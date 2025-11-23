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

Public Module ModNut
    '版本号
    Public VersionNumber As Single = 211.0F

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
        Try
            Dim jsonResponse As String = ModNet.NetRequestByClientRetry(
            "https://flyplayteam.fun/software/server/ClientUpdateInfo2.json",
            Accept:="application/json, text/javascript, */*; q=0.01",
            RequireJson:=True,
            Encoding:=Encoding.UTF8
        )
            Dim updateFiles As JArray = DirectCast(JsonConvert.DeserializeObject(jsonResponse), JObject)("updatefile")

            For Each jToken As JToken In updateFiles
                Dim folderPath As String = jToken("floderpath").ToString()
                Dim oldName As String = jToken("oldname").ToString()
                Dim oldMD5 As String = jToken("oldmd5").ToString().ToLower().Trim()
                Dim newName As String = jToken("newname").ToString()
                Dim newMD5 As String = jToken("newmd5").ToString().ToLower().Trim()
                Dim downloadPath As String = jToken("downloadpath").ToString()
                Dim fullPath As String = IO.Path.Combine(AppContext.BaseDirectory, folderPath)
                Dim showNotice As String = jToken("shownotice").ToString()
                Dim noticeText As String = jToken("notice").ToString()
                Dim operation As String = jToken("opt").ToString().ToLower()
                Dim enable As String = jToken("enable").ToString().ToLower()

                ' 禁用则跳过
                If enable = "false" Then Continue For

                ' 对 delete 模式做个兼容：如果 oldmd5 为空而 newmd5 有值，就当 oldmd5 用
                If operation = "delete" AndAlso oldMD5 = "" AndAlso newMD5 <> "" Then
                    oldMD5 = newMD5
                End If

                ' 压缩包更新单独处理
                If operation = "unpack" Then
                    PackUpdate(jToken)
                    Continue For
                End If

                ' 路径检查 / 创建
                Dim files As FileInfo() = {}
                If Directory.Exists(fullPath) Then
                    files = New DirectoryInfo(fullPath).GetFiles()
                Else
                    If operation = "add" OrElse operation = "replace" Then
                        Directory.CreateDirectory(fullPath)
                    Else
                        Continue For
                    End If
                End If

                Dim targetOld As FileInfo = Nothing      ' 命中的旧文件
                Dim targetOldMD5 As String = Nothing
                Dim hasCorrectNew As Boolean = False     ' 是否已存在正确的新文件
                Dim deleteList As New List(Of FileInfo)  ' 所有要删掉的文件（delete 模式）

                ' 扫描目录内所有文件
                For Each file As FileInfo In files
                    Dim nameMatchOld As Boolean =
                    (Not String.IsNullOrEmpty(oldName) AndAlso
                     file.Name.IndexOf(oldName, StringComparison.OrdinalIgnoreCase) >= 0)
                    Dim nameMatchNew As Boolean =
                    (Not String.IsNullOrEmpty(newName) AndAlso
                     file.Name.Equals(newName, StringComparison.OrdinalIgnoreCase))

                    ' 是否需要算 MD5：只有当配置里写了 MD5 时才算，避免浪费
                    Dim needMd5 As Boolean = False
                    If oldMD5 <> "" Then needMd5 = True
                    If (operation = "add" OrElse operation = "replace") AndAlso newMD5 <> "" Then needMd5 = True

                    Dim fileMd5 As String = Nothing
                    If needMd5 Then
                        Try
                            fileMd5 = GetMD5HashFromFile(file.FullName).ToLower()
                        Catch ex As Exception
                            ModBase.Log(ex, "CheckFileUpdate 计算 MD5 失败（" & file.FullName & "）", ModBase.LogLevel.Hint, "出现错误")
                        End Try
                    End If

                    Dim md5MatchOld As Boolean = (oldMD5 <> "" AndAlso fileMd5 IsNot Nothing AndAlso fileMd5 = oldMD5)
                    Dim md5MatchNew As Boolean = (newMD5 <> "" AndAlso fileMd5 IsNot Nothing AndAlso fileMd5 = newMD5)

                    If operation = "delete" Then
                        ' 删除模式：只要名字命中 或 MD5 命中，就进删除列表
                        If nameMatchOld OrElse md5MatchOld Then
                            deleteList.Add(file)
                        End If
                    Else
                        ' add / replace 模式

                        ' 1) 记录旧文件（老版本）
                        If (nameMatchOld OrElse md5MatchOld) AndAlso targetOld Is Nothing Then
                            targetOld = file
                            targetOldMD5 = fileMd5
                        End If

                        ' 2) 检查是否已经存在“正确的新文件”
                        If nameMatchNew OrElse md5MatchNew Then
                            If newMD5 <> "" Then
                                If md5MatchNew Then
                                    ' 名字 / MD5 都符合，认为已经是新版本
                                    hasCorrectNew = True
                                Else
                                    ' 名字一样但 MD5 不同，当成旧文件替换掉
                                    If targetOld Is Nothing Then
                                        targetOld = file
                                        targetOldMD5 = fileMd5
                                    End If
                                End If
                            Else
                                ' 未提供 newMD5，只要同名就当成已有新版本
                                hasCorrectNew = True
                            End If
                        End If
                    End If
                Next

                '============ delete 模式：删文件 ============
                If operation = "delete" Then
                    If deleteList.Count = 0 Then
                        Continue For
                    End If

                    Dim userResponseDel As Integer = 1
                    If showNotice = "true" Then
                        userResponseDel = ModMain.MyMsgBox(
                        noticeText,
                        "检测到需删除的文件",
                        "删除", "取消", "",
                        False, True, True
                    )
                        If userResponseDel <> 1 Then Continue For
                    End If

                    For Each fileToDelete In deleteList
                        Try
                            If showNotice = "true" Then
                                ModMain.Hint("删除文件:" & fileToDelete.FullName, ModMain.HintType.Green, True)
                            End If

                            ModBase.Log("[ModNut]CheckFileUpdate - 删除文件：" & fileToDelete.FullName, ModBase.LogLevel.Normal, "Normal")
                            fileToDelete.Delete()

                        Catch ex As Exception
                            ModBase.Log(ex, "[ModNut]CheckFileUpdate - 删除文件失败：" & fileToDelete.FullName, ModBase.LogLevel.Feedback, "出现错误")
                        End Try
                    Next

                    Continue For ' 这一条配置处理完了，看下一条
                End If

                '============ add / replace 模式：更新文件 ============

                ' 如果目录里已经有“正确的新文件”，直接跳过（不会再弹更新提示）
                If hasCorrectNew Then
                    Continue For
                End If

                ' replace 模式但压根没找到旧文件 → 当成 add 处理
                If targetOld Is Nothing AndAlso operation = "replace" Then
                    operation = "add"
                End If

                ' 是否需要弹提示（add / replace 共用原来的提示）
                Dim userResponse As Integer = 1
                If showNotice = "true" Then
                    userResponse = ModMain.MyMsgBox(
                    noticeText & vbCrLf & vbCrLf &
                    "需更新文件:" & vbCrLf & fullPath & newName & vbCrLf & vbCrLf &
                    "自动下载更新过程中请勿关闭启动器！" & vbCrLf &
                    "也可以点击【下载链接】自行下载更新",
                    "检测到文件更新",
                    "自动更新", "打开下载链接", "取消更新",
                    False, True, True
                )
                    If userResponse <> 1 Then
                        If userResponse = 2 Then
                            ModBase.OpenWebsite(downloadPath)
                        End If
                        Continue For
                    End If
                End If

                Select Case operation
                    Case "replace"
                        ' 先下载新的，再删除旧文件，防止中途关闭导致旧文件丢失
                        Try
                            DownloadAndSaveFile(downloadPath, fullPath, newName)
                        Catch ex As Exception
                            ModBase.Log(ex, "[ModNut]CheckFileUpdate - 下载新文件失败（replace）：" & IO.Path.Combine(fullPath, newName), ModBase.LogLevel.Feedback, "出现错误")
                            Continue For
                        End Try

                        If targetOld IsNot Nothing AndAlso targetOld.Exists AndAlso
                       Not String.Equals(targetOld.Name, newName, StringComparison.OrdinalIgnoreCase) Then
                            Try

                                ModMain.Hint("删除文件:" & targetOld.FullName, ModMain.HintType.Green, True)
                                ModBase.Log("[ModNut]CheckFileUpdate - 删除文件：" & targetOld.FullName, ModBase.LogLevel.Normal, "Normal")
                                targetOld.Delete()
                            Catch ex As Exception
                                ModBase.Log(ex, "[ModNut]CheckFileUpdate - 删除旧文件失败：" & targetOld.FullName, ModBase.LogLevel.Feedback, "出现错误")
                            End Try
                        End If

                    Case "add"
                        ' 新增：只负责把新文件下载过来
                        Try
                            DownloadAndSaveFile(downloadPath, fullPath, newName)
                        Catch ex As Exception
                            ModBase.Log(ex, "[ModNut]CheckFileUpdate - 下载新文件失败（add）：" & IO.Path.Combine(fullPath, newName), ModBase.LogLevel.Feedback, "出现错误")
                            Continue For
                        End Try
                End Select
            Next
        Catch ex As Exception
            ModBase.Log(ex, "CheckFileUpdate 检查MOD更新失败", ModBase.LogLevel.Feedback, "出现错误")
        End Try
    End Sub





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
