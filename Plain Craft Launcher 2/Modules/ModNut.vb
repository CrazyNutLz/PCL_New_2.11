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
                Dim oldMD5 As String = jToken("oldmd5").ToString().ToLower()
                Dim newName As String = jToken("newname").ToString()
                Dim newMD5 As String = jToken("newmd5").ToString().ToLower()
                Dim downloadPath As String = jToken("downloadpath").ToString()
                Dim fullPath As String = IO.Path.Combine(AppContext.BaseDirectory, folderPath)
                Dim showNotice As String = jToken("shownotice").ToString()
                Dim noticeText As String = jToken("notice").ToString()
                Dim operation As String = jToken("opt").ToString().ToLower()
                Dim enable As String = jToken("enable").ToString().ToLower()

                Dim files As FileInfo() = {}
                Dim targetFile As FileInfo = Nothing
                Dim targetFileMD5 As String = Nothing

                ' 判断是否启用
                If enable = "false" Then
                    Continue For '跳过操作
                End If

                ' 判断是否压缩包更新
                If operation = "unpack" Then
                    PackUpdate(jToken)
                    Continue For '跳过操作
                End If

                ' 判断路径是否存在
                If Directory.Exists(fullPath) Then
                    files = New DirectoryInfo(fullPath).GetFiles() ' 存在则获取路径中的文件
                Else
                    If operation = "add" Then
                        Directory.CreateDirectory(fullPath) ' 路径不存在但为添加模式，创建文件夹
                    Else
                        Continue For ' 路径不存在且为替换或删除模式，跳过操作
                    End If
                End If

                ' 查找老目标文件并计算其 MD5
                For Each file As FileInfo In files
                    If file.Name.Contains(oldName) OrElse GetMD5HashFromFile(file.FullName).ToLower() = oldMD5 Then
                        targetFile = file
                        targetFileMD5 = GetMD5HashFromFile(file.FullName).ToLower()
                        Exit For
                    End If
                Next

                ' 判断文件与操作之间的关系
                If targetFile IsNot Nothing Then
                    If operation = "add" Or operation = "replace" Then
                        If targetFileMD5 = newMD5 Then
                            ' 如果文件存在 且操作模式为添加 或者 替换 且 MD5 与新的相同
                            Continue For
                        End If
                    End If
                ElseIf operation = "delete" Or operation = "replace" Then
                    ' 如果文件不存在 且为删除或替换
                    Continue For
                End If

                ' 判断是否需要显示通知
                Dim userResponse As Integer = 1
                If showNotice = "true" Then
                    userResponse = ModMain.MyMsgBox(noticeText & vbCrLf & vbCrLf & "需更新文件:" & vbCrLf & fullPath & newName & vbCrLf & vbCrLf & "自动下载更新过程中请勿关闭启动器！" & vbCrLf & "也可以点击【下载链接】自行下载更新", "检测到文件更新", "自动更新", "打开下载链接", "取消更新", False, True, True)
                    If userResponse <> 1 Then
                        If userResponse = 2 Then
                            ModBase.OpenWebsite(downloadPath)
                        End If
                        Continue For
                    End If
                End If

                ' 执行指定操作
                Select Case operation
                    Case "replace"
                        If targetFile IsNot Nothing Then
                            ModMain.Hint("删除文件:" & targetFile.FullName, ModMain.HintType.Green, True)
                            ModBase.Log("[ModNut]CheckFileUpdate - 删除文件：" & targetFile.FullName, ModBase.LogLevel.Normal, "Normal")
                            targetFile.Delete()
                            DownloadAndSaveFile(downloadPath, fullPath, newName)
                        End If
                    Case "add"
                        If targetFile IsNot Nothing Then
                            ModMain.Hint("删除文件:" & targetFile.FullName, ModMain.HintType.Green, True)
                            ModBase.Log("[ModNut]CheckFileUpdate - 删除文件：" & targetFile.FullName, ModBase.LogLevel.Normal, "Normal")
                            targetFile.Delete()
                        End If
                        DownloadAndSaveFile(downloadPath, fullPath, newName)
                    Case "delete"
                        If targetFile IsNot Nothing Then
                            ModMain.Hint("删除文件:" & targetFile.FullName, ModMain.HintType.Green, True)
                            ModBase.Log("[ModNut]CheckFileUpdate - 删除文件：" & targetFile.FullName, ModBase.LogLevel.Normal, "Normal")
                            targetFile.Delete()
                        End If
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
