Public Class PageOtherTest

    Public Shared Sub StartCustomDownload(Url As String, FileName As String, Optional Folder As String = Nothing)
        Try
            'Dim fileName As String = "File_" & ModBase.RandomInteger(0, 9999999).ToString()
            Try
                fileName = WebUtility.UrlEncode(ModBase.GetFileNameFromPath(Url))
            Catch ex As Exception
                ModBase.Log(ex, "自定义下载文件没有文件名", ModBase.LogLevel.Debug, "出现错误")
            End Try


            If String.IsNullOrWhiteSpace(Folder) Then
                Folder = ModBase.SelectSaveFile("选择文件保存位置", fileName, Nothing, Nothing)
                If Not Folder.Contains("\") Then
                    Return
                End If
                If Folder.EndsWith(fileName) Then
                    Folder = Folder.Substring(0, Folder.Length - fileName.Length)
                End If
            End If

            Folder = Folder.Replace("/", "\").TrimEnd("\"c) & "\"
            Try
                Directory.CreateDirectory(Folder)
                ModBase.CheckPermissionWithException(Folder)
            Catch ex2 As Exception
                ModBase.Log(ex2, "访问文件夹失败（" & Folder & "）", ModBase.LogLevel.Hint, "出现错误")
                Return
            End Try

            ModBase.Log("[Download] 自定义下载文件名：" & fileName, ModBase.LogLevel.Normal, "出现错误")
            ModBase.Log("[Download] 自定义下载文件目标：" & Folder, ModBase.LogLevel.Normal, "出现错误")

            Dim uuid As Integer = ModBase.GetUuid()
            Dim loaderDownload As New ModNet.LoaderDownload("自定义下载文件：" & fileName & " ", New List(Of ModNet.NetFile) From {
            New ModNet.NetFile(New String() {Url}, Folder & fileName, Nothing)
        })
            Dim loaderCombo As New ModLoader.LoaderCombo(Of Integer)("自定义下载 (" & uuid.ToString() & ") ", New List(Of ModLoader.LoaderBase) From {loaderDownload})

            ' 直接为 OnStateChanged 属性赋值委托
            loaderCombo.OnStateChanged = Sub(sender As Object)
                                             PageOtherTest.DownloadState(DirectCast(sender, ModLoader.LoaderCombo(Of Integer)))
                                         End Sub

            loaderCombo.Start(0, False)
            ModLoader.LoaderTaskbarAdd(loaderCombo)
            ModMain.FrmMain.BtnExtraDownload.ShowRefresh()
            ModMain.FrmMain.BtnExtraDownload.Ribble()
        Catch ex3 As Exception
            ModBase.Log(ex3, "开始自定义下载失败", ModBase.LogLevel.Feedback, "出现错误")
        End Try
    End Sub


    Private Shared Sub DownloadState(ByVal Loader As ModLoader.LoaderCombo(Of Integer))
        Try
            Select Case Loader.State
                Case ModBase.LoadState.Finished
                    ModMain.Hint(Loader.Name & " 完成！")
                    Interaction.Beep()
                Case ModBase.LoadState.Failed
                    ModBase.Log(Loader.Error, Loader.Name & " 失败", ModBase.LogLevel.Msgbox, "出现错误")
                    Interaction.Beep()
                Case ModBase.LoadState.Aborted
                    ModMain.Hint(Loader.Name & " 已取消！")
            End Select
        Catch ex As Exception
            ModBase.Log(ex, "处理下载状态变化时出错", ModBase.LogLevel.Feedback, "出现错误")
        End Try
    End Sub

    Public Shared Sub Jrrp()
        Hint("为便于维护，开源内容中不包含百宝箱功能……")
    End Sub
    Public Shared Sub RubbishClear()
        Hint("为便于维护，开源内容中不包含百宝箱功能……")
    End Sub
    Public Shared Sub MemoryOptimize(ShowHint As Boolean)
        If ShowHint Then Hint("为便于维护，开源内容中不包含百宝箱功能……")
    End Sub
    Public Shared Sub MemoryOptimizeInternal(ShowHint As Boolean)
    End Sub
    Public Shared Function GetRandomCave() As String
        Return "为便于维护，开源内容中不包含百宝箱功能……"
    End Function
    Public Shared Function GetRandomHint() As String
        Return "为便于维护，开源内容中不包含百宝箱功能……"
    End Function
    Public Shared Function GetRandomPresetHint() As String
        Return "为便于维护，开源内容中不包含百宝箱功能……"
    End Function

End Class
