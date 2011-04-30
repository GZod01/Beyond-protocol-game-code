Public Class Form1

    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"

        Dim sCmd As String = Command()
        sCmd = sCmd.Replace("""", "")
        If IsDate(sCmd) = False Then End

        If Exists(sFile & "UpdaterClient.ex_") = True Then
            If Exists(sFile & "UpdaterClient.exe") = True Then
                EnsureNoFileLock(sFile & "UpdaterClient.exe")
                Kill(sFile & "UpdaterClient.exe")
            End If
            EnsureNoFileLock(sFile & "UpdaterClient.ex_")
            FileCopy(sFile & "UpdaterClient.ex_", sFile & "UpdaterClient.exe")
            Threading.Thread.Sleep(100)

            SetFileDateTime(sCmd, sFile & "UpdaterClient.exe")

            Shell("""" & sFile & "UpdaterClient.exe""", AppWinStyle.NormalFocus)
        End If
        End
    End Sub

    Public Sub SetFileDateTime(ByVal sDate As String, ByVal sFilename As String)
        Dim tMyDate As Date

        'now update the datetime stamp
        '2/17/2003 11:56:41 AM
        tMyDate = CDate(sDate)

        Dim oFSInfo As IO.FileInfo = New IO.FileInfo(sFilename)
        oFSInfo.CreationTime = tMyDate
        oFSInfo.LastAccessTime = tMyDate
        oFSInfo.LastWriteTime = tMyDate
        oFSInfo = Nothing

    End Sub

    Public Function Exists(ByVal sFilename As String) As Boolean
        If Len(Trim$(sFilename)) > 0 Then
            On Error Resume Next
            sFilename = Dir$(sFilename, FileAttribute.Archive Or FileAttribute.Directory Or FileAttribute.Hidden Or FileAttribute.Normal Or FileAttribute.ReadOnly Or FileAttribute.System Or FileAttribute.Volume)
            Return Err.Number = 0 And Len(sFilename) > 0
        Else
            Return False
        End If

    End Function

    Private Sub EnsureNoFileLock(ByVal sFileSrc As String)
        'For file lock...
        Dim bGood As Boolean = False
        Dim lAttempts As Int32 = 0
        While bGood = False AndAlso lAttempts < 10
            lAttempts += 1

            Dim oStream As IO.FileStream = Nothing
            Try
                oStream = New IO.FileStream(sFileSrc, IO.FileMode.Append)
                If oStream Is Nothing = False Then oStream.Dispose() 
                bGood = True
            Catch ex As Exception
                bGood = False
            Finally
                If oStream Is Nothing = False Then oStream.Dispose()
                oStream = Nothing
            End Try

            'If we're not valid yet then we wait...
            If bGood = False Then Threading.Thread.Sleep(500)
        End While
    End Sub
End Class
