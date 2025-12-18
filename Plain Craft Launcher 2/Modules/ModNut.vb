Imports System.Net
Imports System.Reflection
Imports System.Security.Cryptography
Imports Microsoft.VisualBasic.CompilerServices
Imports Newtonsoft.Json
Imports System.IO
Imports System.Text
Imports System.Threading.Tasks
Imports System.IO.Compression
Imports System.Security.Policy
Imports System.Diagnostics
Imports Newtonsoft.Json.Linq
Public Module ModNut
    '版本号
    Public VersionNumber As Single = 213.0F

    Private IsCheckUpdate As Boolean = False

    Public Sub CheckPCLUpdate()
        If IsCheckUpdate = False And IsUpdateStarted = False Then
            IsCheckUpdate = True
            Try
                Dim jsonText As String = ModNet.NetRequestByClientRetry(
                        "https://flyplayteam.fun/software/server/pclinfo.json",
                        Accept:="application/json, text/javascript, */*; q=0.01",
                        RequireJson:=True,
                        Encoding:=Encoding.UTF8
                    )
                Dim jobject As JObject = DirectCast(JsonConvert.DeserializeObject(jsonText), JObject)

                Dim version As String = jobject("version").ToString()
                Dim title As String = jobject("title").ToString()
                Dim newtxt As String = jobject("newtxt").ToString()
                newtxt = "当前启动器版本:" & VersionNumber & "---->最新版本:" & jobject("version").ToString() & vbCrLf & vbCrLf & newtxt

                If jobject Is Nothing Then
                    ModMain.Hint("连接服务器失败，或服务器缓存已被清理，请先确认系统时间与北京时间一致，然后重启 PCL2 再试！", ModMain.HintType.Red, True)
                End If
                If Single.Parse(version) <= VersionNumber Then
                    ModBase.Log("当前启动器版本已为最新，无需更新！", ModBase.LogLevel.Debug, "Debug")
                ElseIf ModMain.MyMsgBox(newtxt, title, "确定", "取消", "", False, True, False) <> 2 Then
                    ModSecret.UpdateStart(jobject("downloadurl").ToString(), False)
                End If
            Catch ex As Exception
                ModMain.Hint("CheckPCLUpdate 尝试检查更新失败", ModMain.HintType.Red, True)
                IsCheckUpdate = False
            End Try
            IsCheckUpdate = False
        Else
            ModMain.Hint("请等待上一次更新执行完毕", ModMain.HintType.Red, True)
        End If
    End Sub
    Public Sub GetNotice()
        Try
            Dim jsonResponse As String = ModNet.NetRequestByClientRetry(
                        "https://flyplayteam.fun/software/server/notice17.json",
                        Accept:="application/json, text/javascript, */*; q=0.01",
                        RequireJson:=True,
                        Encoding:=Encoding.UTF8
                    )
            Dim notices As JArray = DirectCast(JsonConvert.DeserializeObject(jsonResponse), JArray)

            For Each jToken As JToken In notices
                Dim title As String = jToken("title").ToString()
                Dim notice As String = jToken("notice").ToString()
                Dim version As String = jToken("version").ToString()
                Dim button1Text As String = jToken("btn1").ToString()
                Dim button2Text As String = jToken("btn2").ToString()
                Dim optionAction As String = jToken("opt").ToString()
                Dim currentNoticeVersion As String = ""

                Try
                    currentNoticeVersion = Convert.ToString(ModBase.Setup.Get("NoticeVersion", Nothing))
                Catch
                    ModBase.Setup.Set("NoticeVersion", "0", False, Nothing)
                End Try

                ModBase.Setup.Set("NoticeVersion", version, False, Nothing)
                Dim showTime As String = jToken("showtime").ToString()
                Dim responseCode As Integer = 0

                If showTime = "always" Then
                    responseCode = ModMain.MyMsgBox(notice, title, button1Text, button2Text, "", False, True, True)
                End If

                If version <> currentNoticeVersion AndAlso showTime = "once" Then
                    responseCode = ModMain.MyMsgBox(notice, title, button1Text, button2Text, "", False, True, True)
                End If

                If optionAction <> "" AndAlso responseCode = 1 Then
                    Dim optionText As String = jToken("opttxt").ToString()

                    Select Case optionAction
                        Case "openurl"
                            ModBase.OpenWebsite(optionText)
                        Case "clipbord"
                            ModBase.ClipboardSet(optionText, True)
                        Case "exit"
                            Environment.Exit(0)
                    End Select
                End If
            Next
        Catch ex As Exception
            ModBase.Log(ex, "GetNotice 获取公告失败", ModBase.LogLevel.Feedback, "出现错误")
        End Try
    End Sub

    Public Sub CheckFileUpdate()
        Dim sw As Stopwatch = Stopwatch.StartNew()

        '（可选）统计本次检查做了哪些事
        Dim ruleTotal As Integer = 0
        Dim unpackCount As Integer = 0
        Dim downloadCount As Integer = 0
        Dim deleteCount As Integer = 0
        Dim skipAsNewestCount As Integer = 0

        ' MD5 缓存：同一次检查中同一文件只算一次
        Dim md5Cache As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Try
            ModBase.Log("[ModNut]CheckFileUpdate - 开始检查更新", ModBase.LogLevel.Normal, "Normal")

            Dim jsonResponse As String = ModNet.NetRequestByClientRetry(
            "https://flyplayteam.fun/software/server/ClientUpdateInfo2.json",
            Accept:="application/json, text/javascript, */*; q=0.01",
            RequireJson:=True,
            Encoding:=Encoding.UTF8
        )

            Dim root As JObject = DirectCast(JsonConvert.DeserializeObject(jsonResponse), JObject)
            Dim updateFiles As JArray = DirectCast(root("updatefile"), JArray)

            For Each jToken As JToken In updateFiles
                ruleTotal += 1

                Dim folderPath As String = jToken("floderpath").ToString()
                Dim oldName As String = jToken("oldname").ToString()
                Dim oldMD5 As String = jToken("oldmd5").ToString().ToLower().Trim()
                Dim newName As String = jToken("newname").ToString()
                Dim newMD5 As String = jToken("newmd5").ToString().ToLower().Trim()
                Dim downloadPath As String = jToken("downloadpath").ToString()
                Dim fullPath As String = IO.Path.Combine(AppContext.BaseDirectory, folderPath)

                Dim showNoticeStr As String = jToken("shownotice").ToString()
                Dim noticeText As String = jToken("notice").ToString()
                Dim operation As String = jToken("opt").ToString().ToLower().Trim()
                Dim enableStr As String = jToken("enable").ToString().ToLower().Trim()

                ' 禁用则跳过
                If enableStr = "false" Then Continue For

                ' delete 模式兼容：如果 oldmd5 为空而 newmd5 有值，就当 oldmd5 用
                If operation = "delete" AndAlso oldMD5 = "" AndAlso newMD5 <> "" Then oldMD5 = newMD5

                ' unpack 单独处理
                If operation = "unpack" Then
                    unpackCount += 1
                    PackUpdate(jToken)
                    Continue For
                End If

                Dim showNotice As Boolean = (Not String.IsNullOrEmpty(showNoticeStr) AndAlso showNoticeStr.Trim().ToLower() = "true")

                ' 路径检查 / 创建
                If Not Directory.Exists(fullPath) Then
                    If operation = "add" OrElse operation = "replace" Then
                        Directory.CreateDirectory(fullPath)
                    Else
                        Continue For
                    End If
                End If

                ' =========================
                ' delete：oldname 命中 OR oldmd5 命中 → 删除（满足其一即可）
                ' =========================
                If operation = "delete" Then
                    Dim deleteSet As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                    ' 1) oldname：包含命中（通配符枚举，避免全量 FileInfo 分配）
                    If Not String.IsNullOrEmpty(oldName) Then
                        Try
                            For Each p In Directory.EnumerateFiles(fullPath, "*" & oldName & "*", SearchOption.TopDirectoryOnly)
                                deleteSet.Add(p)
                            Next
                        Catch ex As Exception
                            ModBase.Log(ex, "CheckFileUpdate(delete) 枚举文件失败：" & fullPath, ModBase.LogLevel.Hint, "出现错误")
                        End Try
                    End If

                    ' 2) oldmd5：为了严格满足 OR 语义，只要配置了 oldMD5 就需要全目录比对 MD5
                    If Not String.IsNullOrEmpty(oldMD5) Then
                        Try
                            For Each p In Directory.EnumerateFiles(fullPath, "*", SearchOption.TopDirectoryOnly)
                                Dim md5 As String = GetMd5Cached(p, md5Cache)
                                If md5 <> "" AndAlso md5 = oldMD5 Then deleteSet.Add(p)
                            Next
                        Catch ex As Exception
                            ModBase.Log(ex, "CheckFileUpdate(delete) MD5 扫描失败：" & fullPath, ModBase.LogLevel.Hint, "出现错误")
                        End Try
                    End If

                    If deleteSet.Count = 0 Then Continue For

                    If showNotice Then
                        Dim userResponseDel As Integer = ModMain.MyMsgBox(
                        noticeText,
                        "检测到需删除的文件",
                        "删除", "取消", "",
                        False, True, True
                    )
                        If userResponseDel <> 1 Then Continue For
                    End If

                    For Each filePath In deleteSet
                        Try
                            If showNotice Then ModMain.Hint("删除文件:" & filePath, ModMain.HintType.Green, True)
                            ModBase.Log("[ModNut]CheckFileUpdate - 删除文件：" & filePath, ModBase.LogLevel.Normal, "Normal")
                            File.Delete(filePath)
                            deleteCount += 1
                            md5Cache.Remove(filePath)
                        Catch ex As Exception
                            ModBase.Log(ex, "[ModNut]CheckFileUpdate - 删除文件失败：" & filePath, ModBase.LogLevel.Feedback, "出现错误")
                        End Try
                    Next

                    Continue For
                End If

                ' =========================
                ' add / replace
                ' =========================
                Dim newFilePath As String = If(String.IsNullOrEmpty(newName), "", IO.Path.Combine(fullPath, newName))

                ' 先用“直接路径”判断是否已是新版本（避免扫描目录）
                If Not String.IsNullOrEmpty(newFilePath) AndAlso File.Exists(newFilePath) Then
                    If String.IsNullOrEmpty(newMD5) Then
                        skipAsNewestCount += 1
                        Continue For
                    Else
                        Dim md5 As String = GetMd5Cached(newFilePath, md5Cache)
                        If md5 <> "" AndAlso md5 = newMD5 Then
                            skipAsNewestCount += 1
                            Continue For
                        End If
                    End If
                End If

                ' replace：定位旧文件（优先精确路径，其次通配符枚举；必要时才全目录 MD5）
                Dim targetOldPath As String = Nothing

                If operation = "replace" Then
                    ' 1) oldname 看起来是完整文件名时：直接命中
                    If Not String.IsNullOrEmpty(oldName) Then
                        Dim oldExactPath As String = IO.Path.Combine(fullPath, oldName)
                        If File.Exists(oldExactPath) Then
                            targetOldPath = oldExactPath
                        Else
                            ' 2) oldname 不是完整文件名：通配符枚举缩小范围
                            Try
                                For Each p In Directory.EnumerateFiles(fullPath, "*" & oldName & "*", SearchOption.TopDirectoryOnly)
                                    ' 排除 newName 本身被当成 old
                                    If Not String.IsNullOrEmpty(newName) AndAlso
                                   String.Equals(IO.Path.GetFileName(p), newName, StringComparison.OrdinalIgnoreCase) Then
                                        Continue For
                                    End If

                                    If String.IsNullOrEmpty(oldMD5) Then
                                        targetOldPath = p
                                        Exit For
                                    Else
                                        Dim md5 As String = GetMd5Cached(p, md5Cache)
                                        If md5 <> "" AndAlso md5 = oldMD5 Then
                                            targetOldPath = p
                                            Exit For
                                        End If
                                    End If
                                Next
                            Catch ex As Exception
                                ModBase.Log(ex, "CheckFileUpdate(replace) 枚举旧文件失败：" & fullPath, ModBase.LogLevel.Hint, "出现错误")
                            End Try
                        End If
                    ElseIf Not String.IsNullOrEmpty(oldMD5) Then
                        ' 3) 只有 oldMD5：只能全目录找
                        Try
                            For Each p In Directory.EnumerateFiles(fullPath, "*", SearchOption.TopDirectoryOnly)
                                Dim md5 As String = GetMd5Cached(p, md5Cache)
                                If md5 <> "" AndAlso md5 = oldMD5 Then
                                    targetOldPath = p
                                    Exit For
                                End If
                            Next
                        Catch ex As Exception
                            ModBase.Log(ex, "CheckFileUpdate(replace) oldMD5 扫描失败：" & fullPath, ModBase.LogLevel.Hint, "出现错误")
                        End Try
                    End If

                    ' replace 模式但没找到旧文件 → 当成 add（保持你原逻辑）
                    If targetOldPath Is Nothing Then operation = "add"
                End If

                ' 提示（add / replace 共用）
                If showNotice Then
                    Dim userResponse As Integer = ModMain.MyMsgBox(
                    noticeText & vbCrLf & vbCrLf &
                    "需更新文件:" & vbCrLf & IO.Path.Combine(fullPath, newName) & vbCrLf & vbCrLf &
                    "自动下载更新过程中请勿关闭启动器！" & vbCrLf &
                    "也可以点击【下载链接】自行下载更新",
                    "检测到文件更新",
                    "自动更新", "打开下载链接", "取消更新",
                    False, True, True
                )

                    If userResponse <> 1 Then
                        If userResponse = 2 Then ModBase.OpenWebsite(downloadPath)
                        Continue For
                    End If
                End If

                ' 执行下载（add/replace 都是下载到 newName）
                Try
                    DownloadAndSaveFile(downloadPath, fullPath, newName)
                    downloadCount += 1

                    ' 新文件下载后，让其 MD5 缓存失效（防止下一条规则读取到旧值）
                    If Not String.IsNullOrEmpty(newFilePath) Then md5Cache.Remove(newFilePath)
                Catch ex As Exception
                    ModBase.Log(ex, "[ModNut]CheckFileUpdate - 下载新文件失败：" & IO.Path.Combine(fullPath, newName), ModBase.LogLevel.Feedback, "出现错误")
                    Continue For
                End Try

                ' replace：下载成功后再删除旧文件（保持你原“只删一个旧文件”的行为）
                If operation = "replace" AndAlso Not String.IsNullOrEmpty(targetOldPath) Then
                    Try
                        If File.Exists(targetOldPath) AndAlso
                       Not String.Equals(IO.Path.GetFileName(targetOldPath), newName, StringComparison.OrdinalIgnoreCase) Then

                            If showNotice Then ModMain.Hint("删除文件:" & targetOldPath, ModMain.HintType.Green, True)
                            ModBase.Log("[ModNut]CheckFileUpdate - 删除文件：" & targetOldPath, ModBase.LogLevel.Normal, "Normal")
                            File.Delete(targetOldPath)
                            deleteCount += 1
                            md5Cache.Remove(targetOldPath)
                        End If
                    Catch ex As Exception
                        ModBase.Log(ex, "[ModNut]CheckFileUpdate - 删除旧文件失败：" & targetOldPath, ModBase.LogLevel.Feedback, "出现错误")
                    End Try
                End If
            Next

        Catch ex As Exception
            ModBase.Log(ex, "CheckFileUpdate 检查MOD更新失败", ModBase.LogLevel.Feedback, "出现错误")
        Finally
            sw.Stop()
            ModBase.Log(
            "[ModNut]CheckFileUpdate - 检查完成，耗时 " & sw.ElapsedMilliseconds & " ms" &
            " | 规则=" & ruleTotal &
            " | 解包=" & unpackCount &
            " | 下载=" & downloadCount &
            " | 删除=" & deleteCount &
            " | 跳过(已最新)=" & skipAsNewestCount,
            ModBase.LogLevel.Normal,
            "Normal"
        )
        End Try
    End Sub

    Private Function GetMd5Cached(path As String, cache As Dictionary(Of String, String)) As String
        Try
            Dim v As String = Nothing
            If cache.TryGetValue(path, v) Then Return v
            Dim md5 = GetMD5HashFromFile(path).ToLower().Trim()
            cache(path) = md5
            Return md5
        Catch ex As Exception
            ModBase.Log(ex, "CheckFileUpdate 计算 MD5 失败（" & path & "）", ModBase.LogLevel.Hint, "出现错误")
            Return ""
        End Try
    End Function





    Private Sub PackUpdate(jToken As JToken)
        Try
            Dim folderPath As String = jToken("floderpath").ToString()
            Dim oldName As String = jToken("oldname").ToString()
            Dim oldMD5 As String = jToken("oldmd5").ToString().ToLower()
            Dim newName As String = jToken("newname").ToString()
            Dim newMD5 As String = jToken("newmd5").ToString().ToLower()
            Dim downloadPath As String = jToken("downloadpath").ToString()
            Dim fullPath As String = IO.Path.Combine(AppContext.BaseDirectory, folderPath)
            Dim showNotice As String = jToken("shownotice").ToString()
            Dim noticeText As String = jToken("notice").ToString()
            Dim operation As String = jToken("opt").ToString().ToLower()
            Dim enable As String = jToken("enable").ToString().ToLower()

            Dim configFilePath As String = IO.Path.Combine(AppContext.BaseDirectory, "PCL\PackUpdateInfo.ini")

            ' 确保目标目录存在
            If Not Directory.Exists(fullPath) Then
                Directory.CreateDirectory(fullPath)
                ModBase.Log("[PackUpdate] 创建目标目录：" & fullPath, ModBase.LogLevel.Normal, "信息")
            End If

            ' 检查配置文件是否存在，如果没有则创建
            If Not File.Exists(configFilePath) Then
                Try
                    Dim iniContent As String = "[PackUpdate]" & vbCrLf
                    File.WriteAllText(configFilePath, iniContent, Encoding.UTF8)
                    ModBase.Log("[PackUpdate] 配置文件不存在，已创建：" & configFilePath, ModBase.LogLevel.Normal, "信息")
                Catch ex As Exception
                    ModBase.Log(ex, "[PackUpdate] 创建配置文件失败", ModBase.LogLevel.Feedback, "错误")
                    Exit Sub
                End Try
            End If

            ' 读取配置文件并检查 MD5 是否匹配
            Dim config As Dictionary(Of String, String) = ReadConfigFile(configFilePath)
            If config.ContainsKey(newName) Then
                If config(newName) = newMD5 Then
                    ' 如果 MD5 相同，则不需要更新
                    ModBase.Log("[PackUpdate] 文件 " & newName & " 已经是最新版本，不需要更新", ModBase.LogLevel.Normal, "信息")
                    Exit Sub
                End If
            End If

            ' 显示更新提示
            If showNotice = "true" Then
                Dim userResponse As Integer = ModMain.MyMsgBox(noticeText & vbCrLf & vbCrLf & "需更新文件:" & vbCrLf & fullPath & newName & vbCrLf & vbCrLf & "自动下载更新过程中请勿关闭启动器！" & vbCrLf & "也可以点击【下载链接】自行下载更新", "检测到文件更新", "自动更新", "打开下载链接", "取消更新", False, True, True)
                If userResponse <> 1 Then
                    If userResponse = 2 Then
                        ModBase.OpenWebsite(downloadPath)
                    End If
                    ModBase.Log("[PackUpdate] 用户取消更新", ModBase.LogLevel.Normal, "信息")
                    Exit Sub
                End If
            End If

            ' 下载并解压更新包
            Try
                DownloadAndExtractPackage(downloadPath, fullPath， oldName)
                ModBase.Log("[PackUpdate] 更新包已成功下载并解压到 " & fullPath, ModBase.LogLevel.Normal, "信息")
            Catch ex As Exception
                ModBase.Log(ex, "[PackUpdate] 更新包下载或解压失败", ModBase.LogLevel.Feedback, "错误")
                Exit Sub
            End Try

            ' 更新配置文件
            Try
                ModMain.Hint("更新包解压完成!", ModMain.HintType.Green, True)
                UpdateConfigFile(configFilePath, newName, newMD5)
                ModBase.Log("[PackUpdate] 配置文件已更新，添加/修改了 " & newName & " 的 MD5 值", ModBase.LogLevel.Normal, "信息")
            Catch ex As Exception
                ModBase.Log(ex, "[PackUpdate] 更新配置文件失败", ModBase.LogLevel.Feedback, "错误")
            End Try

        Catch ex As Exception
            ModBase.Log(ex, "[PackUpdate] 执行过程中发生未知错误", ModBase.LogLevel.Feedback, "错误")
        End Try
    End Sub



    ' 读取配置文件并返回键值对
    Private Function ReadConfigFile(configFilePath As String) As Dictionary(Of String, String)
        Dim config As New Dictionary(Of String, String)()
        If File.Exists(configFilePath) Then
            Dim lines As String() = File.ReadAllLines(configFilePath, Encoding.UTF8)
            For Each line In lines
                If line.StartsWith("[") OrElse String.IsNullOrWhiteSpace(line) Then Continue For
                Dim parts As String() = line.Split("="c)
                If parts.Length = 2 Then
                    config(parts(0).Trim()) = parts(1).Trim()
                End If
            Next
        End If
        Return config
    End Function

    ' 下载并解压更新包
    Private Sub DownloadAndExtractPackage(downloadPath As String, fullPath As String, fileName As String)
        Dim tempZipPath As String = IO.Path.Combine(AppContext.BaseDirectory, fileName)

        DownloadAndSaveFile(downloadPath, AppContext.BaseDirectory, fileName)

        ' 等待文件下载（建议改用异步等待或事件通知）
        While Not File.Exists(tempZipPath)
            Thread.Sleep(100)
        End While

        If File.Exists(tempZipPath) Then
            Using archive As ZipArchive = ZipFile.OpenRead(tempZipPath)
                For Each entry As ZipArchiveEntry In archive.Entries
                    Dim targetPath As String = IO.Path.GetFullPath(IO.Path.Combine(fullPath, entry.FullName))

                    If Not targetPath.StartsWith(fullPath, StringComparison.Ordinal) Then
                        Throw New IOException("安全警告：尝试解压到非法路径！")
                    End If

                    If entry.FullName.EndsWith("/") Then
                        Directory.CreateDirectory(targetPath)
                        Continue For
                    End If

                    Directory.CreateDirectory(IO.Path.GetDirectoryName(targetPath))
                    entry.ExtractToFile(targetPath, True)
                Next
            End Using

            File.Delete(tempZipPath)
        Else
            Throw New Exception("更新包下载失败: " & tempZipPath)
        End If
    End Sub

    ' 更新配置文件，写入新的键值对
    Private Sub UpdateConfigFile(configFilePath As String, newName As String, newMD5 As String)
        Dim config As Dictionary(Of String, String) = ReadConfigFile(configFilePath)

        ' 添加或更新键值对
        If config.ContainsKey(newName) Then
            config(newName) = newMD5
        Else
            config.Add(newName, newMD5)
        End If

        ' 将更新后的配置写入文件
        Dim sb As New StringBuilder()
        sb.AppendLine("[PackUpdate]")
        For Each kvp In config
            sb.AppendLine($"{kvp.Key}={kvp.Value}")
        Next

        File.WriteAllText(configFilePath, sb.ToString(), Encoding.UTF8)
    End Sub



    Private Sub DownloadAndSaveFile(downloadUrl As String, savePath As String, fileName As String)
        ' 处理 seafile 跳转
        If downloadUrl.Contains("seafile") Then
            downloadUrl = GetResponseUri(downloadUrl)
        End If

        ' savePath 在你上面的调用里传的是文件夹路径，这里统一整理成标准目录格式
        Dim folder As String = savePath
        If String.IsNullOrWhiteSpace(folder) Then
            folder = AppContext.BaseDirectory
        End If
        folder = folder.Replace("/", "\").TrimEnd("\"c) & "\"

        ' 触发自定义事件：下载文件
        ' Arg 格式：Url|FileName|Folder
        CustomEvent.Raise(
        CustomEvent.EventType.下载文件,
        $"{downloadUrl}|{fileName}|{folder}"
    )
    End Sub

    Public Function GetResponseUri(ByVal Url As String) As String
        Dim resultUri As String
        Try
            Dim httpRequest As HttpWebRequest = DirectCast(WebRequest.Create(Url), HttpWebRequest)
            Using httpResponse As HttpWebResponse = DirectCast(httpRequest.GetResponse(), HttpWebResponse)
                resultUri = httpResponse.ResponseUri.ToString()
            End Using
        Catch ex As Exception
            ModBase.Log(ex, "GetResponseUri 获取响应链接失败", ModBase.LogLevel.Feedback, "出现错误")
            resultUri = Nothing
        End Try
        Return resultUri
    End Function


    Public Function GetMD5HashFromFile(ByVal fileName As String) As String
        Try
            Using fileStream As FileStream = File.OpenRead(fileName) ' 更安全的打开文件方式，只读模式
                Dim md5Provider As New MD5CryptoServiceProvider()
                Dim hash As Byte() = md5Provider.ComputeHash(fileStream)
                Dim stringBuilder As New StringBuilder()
                For Each b As Byte In hash
                    stringBuilder.Append(b.ToString("x2"))
                Next
                Return stringBuilder.ToString()
            End Using
        Catch ex As Exception
            Throw New Exception("GetMD5HashFromFile() failed, error: " & ex.Message)
        End Try
    End Function




End Module
