Public Enum UpdaterMsgCode As Short
    eLoginMsg = 1
    eRequestFile
    eFileUpdate
End Enum

Public Enum AppTypeCode As int32
    BeyondProtocol = 0
    StarAlliances = 1
End Enum

Public Class frmMain
    Private gsFilePath As String

    Private mbWaiting As Boolean = False
    Private mbSilentSaveToDisk As Boolean = True

    Private mlCurrentFileTotalLen As Int32
    Private mlTotalLen As Int32
    Private mlOverallProgress As Int32

    Private mbHadError As Boolean = False
    Private WithEvents moServer As NetSock

    Private msServerIPList() As String = Nothing
    Private mlServerIPPort() As Int32 = Nothing
    Private mlServerIPUB As Int32 = -1

    Private mdtStartedUpdate As Date

    Private msEntries() As String
    Private mlEntryUB As Int32 = -1

    Private mlAppType As Int32

    Private msCurrentFile As String = ""
    Private mlCurrentFileLen As Int32 = -1

	Public Sub AddTextLine(ByVal sLine As String)
		txtStatus.Text &= sLine & vbCrLf
		txtStatus.SelectionStart = txtStatus.Text.Length
		txtStatus.ScrollToCaret()
    End Sub

    Private Sub frmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Application.Exit()
        End
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If moVersions Is Nothing = False Then
            moVersions.Close()
            moVersions.Dispose()
            moVersions = Nothing
        End If
        

    End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Form.CheckForIllegalCrossThreadCalls = False
        Dim oMe As Process = Process.GetCurrentProcess()

        Dim bDone As Boolean = False
        Dim lCnt As Int32 = 0
        While bDone = False
            bDone = True
            lCnt += 1

            If lCnt > 100 Then
                MsgBox("The Updater encountered an error while attempting to initialize. This may be due to a previous updater instance running.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Error Encountered")
                Application.Exit()
                Return
            End If
            For Each oProc As Process In Process.GetProcesses
                If oProc.Id = oMe.Id Then Continue For

                If oProc.ProcessName = oMe.ProcessName Then
                    Try
                        oProc.Kill()
                        bDone = False
                    Catch
                    End Try
                    Exit For
                End If
            Next
        End While
        
    End Sub

    Private Sub frmMain_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        mlAppType = AppTypeCode.BeyondProtocol

        If Command() Is Nothing = False AndAlso Command.Length > 0 Then
            Dim sAppCode As String = Command()
            If sAppCode.Trim.ToUpper.StartsWith("SA") Then
                mlAppType = AppTypeCode.StarAlliances
            End If
        End If

        Dim oINI As New InitFile()
        Dim lTempVal As Int32 = CInt(Val(oINI.GetString("RESTART", "FORCETYPE", "-1")))
        If lTempVal <> -1 Then
            mlAppType = lTempVal
        End If
        oINI.WriteString("RESTART", "FORCETYPE", "-1")

        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile &= "\"

        If Exists(sFile & "SAClient.exe") = True AndAlso mlAppType <> AppTypeCode.StarAlliances Then
            If Exists(sFile & "BPClient.exe") = False Then
                mlAppType = AppTypeCode.StarAlliances
            End If
        End If

        If mlAppType = AppTypeCode.StarAlliances Then
            Dim sFinal As String = sFile & "SAEULA.htm"
            If Exists(sFinal) = False Then
                sFile &= "EULA.htm"
            Else
                sFile = sFinal
            End If
        Else
            sFile &= "EULA.htm"
        End If

        If Exists(sFile) = False Then Me.Close()
        wbEULA.Navigate(sFile)
    End Sub

    Private moVersions As IO.FileStream = Nothing
	Private Sub DoUpdate()
		'get news file
        'mbWaiting = True
        'moServer.SendData(MakeGetFileRequest("News.txt"))
        'While mbWaiting
        '	Application.DoEvents()
        'End While

        'If Exists(gsFilePath & "News.txt") = False Then
        '	txtNews.Text = "Unable to display news."
        'Else
        '	'Now display the news/updates
        'End If
        mbWaiting = True
        moServer.SendData(MakeGetFileRequest("Versions.txt"))
		While mbWaiting
			Application.DoEvents()
		End While

        If Exists(gsFilePath & "UpdaterUpdater.exe") = True Then
            Try
                Kill(gsFilePath & "UpdaterUpdater.exe")
            Catch
            End Try
        End If
        If Exists(gsFilePath & "UpdaterUpdater.exe") = False Then
            mbWaiting = True
            Dim yReq() As Byte = MakeGetFileRequest("UpdaterUpdater.exe")
            mlCurrentFileLen = 36864
            moServer.SendData(yReq)
            While mbWaiting
                Application.DoEvents()
            End While
        End If

        If Exists(gsFilePath & "UpdaterClient.in_") Then Kill(gsFilePath & "UpdaterClient.in_")
        mbWaiting = True
        moServer.SendData(MakeGetFileRequest("UpdaterClient.in_"))
        While mbWaiting
            Application.DoEvents()
        End While

        If INIUpdated() = True Then
            Me.Close()
            MsgBox("Updater Client has been updated and will be restarted.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Information, "Updated")
            Shell(gsFilePath & "UpdaterClient.exe", AppWinStyle.NormalFocus)
            Return
        End If

        If Exists(gsFilePath & "FinalizeInstall.exe") = True Then
            Try
                Kill(gsFilePath & "FinalizeInstall.exe")
            Catch
            End Try
        End If

        Try
            'if exists(gsfilepath & "
            If Exists(gsFilePath & "dxdiag.txt") = False Then
                Shell("dxdiag /t " & gsFilePath & "dxdiag.txt", AppWinStyle.Hide, False, -1)
            End If
        Catch           'do nothing
        End Try

        moVersions = New IO.FileStream(gsFilePath & "Versions.txt", IO.FileMode.Open)
        Dim oReader As IO.StreamReader = New IO.StreamReader(moVersions)
        Dim sInput As String
        Dim sEntries() As String
        Dim sFileName As String
        Dim bDownload As Boolean
        Dim sDestName As String
        Dim sFilePath As String

        Dim sININame As String = "BPCLIENT.INI"
        If mlAppType = AppTypeCode.StarAlliances Then sININame = "SACLIENT.INI"

        mlOverallProgress = 0
        mlTotalLen = 0
        While oReader.EndOfStream = False
            sInput = oReader.ReadLine()
            If Trim$(sInput) <> "" AndAlso sInput <> vbCrLf Then
                sEntries = Split(sInput, "|")
                sFileName = sEntries(0)

                bDownload = False
                If UBound(sEntries) > 1 Then
                    'Check if the file exists


                    If sEntries.GetUpperBound(0) >= 3 Then
                        If sEntries(3).ToUpper = "SA" Then
                            If mlAppType <> AppTypeCode.StarAlliances Then Continue While
                        ElseIf sEntries(3).ToUpper = "BP" Then
                            If mlAppType <> AppTypeCode.BeyondProtocol Then Continue While
                        End If
                    End If

                    If Exists(gsFilePath & sFileName) Then

                        'MSC - 1/26/2007 - do not update their INI settings...
                        If sFileName.ToUpper.Contains(sININame) = False Then
                            'Ok, it exists, check file length
                            If FileLen(gsFilePath & sFileName) <> CInt(Val(sEntries(1))) Then
                                bDownload = True
                                'Else
                                'filesystem.FileDateTime(gsfilepath & sfilename)
                            End If

                            Dim oFSInfo As IO.FileInfo = New IO.FileInfo(gsFilePath & sFileName)
                            If DateDiff("s", oFSInfo.CreationTime, GetLocaleSpecificDT(sEntries(2))) <> 0 Then
                                bDownload = True
                            End If
                            oFSInfo = Nothing

                            'Dim ofVer As FileVersionInfo = FileVersionInfo.GetVersionInfo(gsFilePath & sFileName)
                            'Dim sVerVal As String = ofVer.FileMajorPart & "_" & ofVer.FileMinorPart & "_" & ofVer.FileVersion
                            'ofVer = Nothing
                            'If sEntries(1) <> sVerVal Then
                            '    bDownload = True
                            'End If

                            'End If
                        End If
                    Else
                        bDownload = True
                    End If
                Else
                    bDownload = True
                End If

                If bDownload Then

                    If sFileName.ToUpper.Contains("UPDATERCLIENT.EX_") = True Then
                        If Exists(gsFilePath & "UpdaterClient.ex_") Then Kill(gsFilePath & "UpdaterClient.ex_")
                        mbWaiting = True
                        Dim yReq() As Byte = MakeGetFileRequest("UpdaterClient.ex_")
                        mlCurrentFileLen = CInt(Val(sEntries(1)))
                        moServer.SendData(yReq)
                        While mbWaiting
                            Application.DoEvents()
                        End While

                        Call SetFileDateTime(sEntries(2), gsFilePath & sFileName)

                        Dim oTmpINI As New InitFile()
                        oTmpINI.WriteString("RESTART", "FORCETYPE", mlAppType.ToString)
                        oTmpINI = Nothing

                        MsgBox("The Updater Client has been updated and will restart.", MsgBoxStyle.OkOnly, "Updated")
                        Shell("""" & gsFilePath & "UpdaterUpdater.exe "" """ & sEntries(2) & """")
                        End
                    End If

                    mlTotalLen += CInt(Val(sEntries(1)))
                End If
            End If
        End While
        oReader.BaseStream.Seek(0, IO.SeekOrigin.Begin)

        mbSilentSaveToDisk = False

        While oReader.EndOfStream = False
            sInput = oReader.ReadLine()

            If Trim$(sInput) <> "" AndAlso sInput <> vbCrLf Then
                sEntries = Split(sInput, "|")
                sFileName = sEntries(0)

                bDownload = False
                If UBound(sEntries) > 1 Then
                    'Check if the file exists

                    If sEntries.GetUpperBound(0) >= 3 Then
                        If sEntries(3).ToUpper = "SA" Then
                            If mlAppType <> AppTypeCode.StarAlliances Then Continue While
                        ElseIf sEntries(3).ToUpper = "BP" Then
                            If mlAppType <> AppTypeCode.BeyondProtocol Then Continue While
                        End If
                    End If

                    If Exists(gsFilePath & sFileName) Then

                        'MSC - 1/26/2007 - do not update their INI settings...
                        If sFileName.ToUpper.Contains(sININame) = False Then
                            'Ok, it exists, check file length
                            If FileLen(gsFilePath & sFileName) <> CInt(Val(sEntries(1))) Then
                                bDownload = True
                                'Else
                                'filesystem.FileDateTime(gsfilepath & sfilename)
                            End If

                            Dim oFSInfo As IO.FileInfo = New IO.FileInfo(gsFilePath & sFileName)
                            'If DateDiff("s", oFSInfo.CreationTime, CDate(sEntries(2))) <> 0 Then
                            If DateDiff("s", oFSInfo.CreationTime, GetLocaleSpecificDT(sEntries(2))) <> 0 Then
                                bDownload = True
                            End If
                            oFSInfo = Nothing

                            'Dim ofVer As FileVersionInfo = FileVersionInfo.GetVersionInfo(gsFilePath & sFileName)
                            'Dim sVerVal As String = ofVer.FileMajorPart & "_" & ofVer.FileMinorPart & "_" & ofVer.FileVersion
                            'ofVer = Nothing
                            'If sEntries(1) <> sVerVal Then
                            '    bDownload = True
                            'End If

                            'End If
                        End If
                    Else
                        bDownload = True
                    End If
                Else
                    bDownload = True
                End If

                If bDownload Then
                    prgCurrent.Value = 0
                    sDestName = sFileName

                    Call EnsureDirectoryStructure(sDestName)
                    sFilePath = StripFilePath(sDestName)

                    Mid$(sDestName, Len(sDestName), 1) = "_"
                    'Kill the dest file
                    'If UBound(sEntries) >= 3 Then
                    '    If sEntries(3) = "1" Then
                    '        'Have to unregister/register
                    '        If Exists(gsFilePath & sFileName) Then
                    '            'need to unregister
                    '            Shell("regsvr32 /u /s """ & gsFilePath & sFileName & """")
                    '        End If
                    '    End If
                    'End If


                    AddTextLine(Space$(3) & "Retrieving " & sFileName & "...")
                    mlCurrentFileTotalLen = CInt(Val(sEntries(1)))
                    mbWaiting = True
                    msCurrentFile = sInput
                    Dim yRequest() As Byte = MakeGetFileRequest(sFileName)
                    mlCurrentFileLen = CInt(Val(sEntries(1)))
                    moServer.SendData(yRequest)
                    Me.Refresh()

                    While mbWaiting
                        Application.DoEvents()
                        Threading.Thread.Sleep(1)
                    End While

                    If mbHadError = True Then Continue While
                    Try
                        Call SetFileDateTime(sEntries(2), gsFilePath & sFileName)

                        'If UBound(sEntries) >= 3 Then
                        '    If sEntries(3) = "1" Then
                        '        'register it
                        '        Shell("regsvr32 /s """ & gsFilePath & sFileName & """")
                        '    End If
                        'End If

                        AddTextLine(Space$(3) & "File Saved Successfully!")

                        mlOverallProgress += mlCurrentFileTotalLen
                        If mlTotalLen <> 0 Then
                            prgOverall.Value = CInt(prgOverall.Maximum * (Math.Min(mlTotalLen, mlOverallProgress) / mlTotalLen))
                        End If
                    Catch
                    End Try

                End If

                'prgOverall.Value = CInt(prgOverall.Maximum * (lCurrent / lTotal))
            End If
        End While
        prgOverall.Value = prgOverall.Maximum
        prgCurrent.Value = prgCurrent.Maximum

        ''Now, go back through retry files
        'If msRetryFiles Is Nothing = False Then
        '    AddTextLine("Retrying delayed requests...")
        '    For X As Int32 = 0 To msRetryFiles.GetUpperBound(0)
        '        sEntries = Split(msRetryFiles(X), "|")
        '        sFileName = sEntries(0)

        '        AddTextLine(Space$(3) & "Retrieving " & sFileName & "...")
        '        mbWaiting = True
        '        msCurrentFile = sFileName
        '        moServer.SendData(MakeGetFileRequest(sFileName))
        '        Me.Refresh()

        '        While mbWaiting
        '            Application.DoEvents()
        '            Threading.Thread.Sleep(1)
        '        End While

        '        If mbHadError = True Then Continue For
        '        Try
        '            Call SetFileDateTime(sEntries(2), gsFilePath & sFileName)

        '            If UBound(sEntries) >= 3 Then
        '                If sEntries(3) = "1" Then
        '                    'register it
        '                    Shell("regsvr32 /s """ & gsFilePath & sFileName & """")
        '                End If
        '            End If

        '            AddTextLine(Space$(3) & "File Saved Successfully!")

        '            mlOverallProgress += mlCurrentFileTotalLen
        '            If mlTotalLen <> 0 Then
        '                prgOverall.Value = CInt(prgOverall.Maximum * (Math.Min(mlTotalLen, mlOverallProgress) / mlTotalLen))
        '            End If
        '        Catch
        '        End Try
        '    Next X
        'End If

        'Ok, do a verify pass now
        If mbHadError = False Then
            oReader.BaseStream.Seek(0, IO.SeekOrigin.Begin)
            While oReader.EndOfStream = False
                sInput = oReader.ReadLine()

                If Trim$(sInput) <> "" AndAlso sInput <> vbCrLf Then
                    mlEntryUB += 1
                    ReDim Preserve msEntries(mlEntryUB)
                    msEntries(mlEntryUB) = sInput

                    sEntries = Split(sInput, "|")
                    sFileName = sEntries(0)

                    bDownload = False
                    If UBound(sEntries) > 1 Then
                        'Check if the file exists
                        If sEntries.GetUpperBound(0) >= 3 Then
                            If sEntries(3).ToUpper = "SA" Then
                                If mlAppType <> AppTypeCode.StarAlliances Then Continue While
                            ElseIf sEntries(3).ToUpper = "BP" Then
                                If mlAppType <> AppTypeCode.BeyondProtocol Then Continue While
                            End If
                        End If

                        If Exists(gsFilePath & sFileName) Then

                            'MSC - 1/26/2007 - do not update their INI settings...
                            If sFileName.ToUpper.Contains(sININame) = False Then

                                Dim oFSInfo As IO.FileInfo = New IO.FileInfo(gsFilePath & sFileName)
                                'If DateDiff("s", oFSInfo.CreationTime, CDate(sEntries(2))) <> 0 Then
                                If DateDiff("s", oFSInfo.CreationTime, GetLocaleSpecificDT(sEntries(2))) <> 0 Then
                                    bDownload = True
                                End If
                                oFSInfo = Nothing
                            End If
                        Else
                            bDownload = True
                        End If
                    Else
                        bDownload = True
                    End If

                    If bDownload = True Then
                        mbHadError = True
                        AddTextLine("  Out of Date: " & sFileName)
                    End If
                End If
            End While
            If mbHadError = True Then
                MsgBox("Not all files were updated. If you are running Vista, you MUST allow the" & vbCrLf & _
                   "updater to update all of the files before being able to log on. Please" & vbCrLf & _
                   "close the updater and try again. If you continue to get this message," & vbCrLf & _
                   "send an email to support@darkskyentertainment.com.", MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Update Incomplete")
            End If
        End If

        oReader.Close()
        oReader.Dispose()
        oReader = Nothing
        moVersions.Close()
        moVersions.Dispose()
        moVersions = Nothing

        'mbWaiting = True
        'moServer.SendData(MakeGetFileRequest("Verify.txt"))
        'While mbWaiting
        '    Application.DoEvents()
        'End While
        'Dim oVerify As New IO.FileStream(gsFilePath & "Verify.txt", IO.FileMode.Open)
        'oReader = New IO.StreamReader(oVerify)
        'While oReader.EndOfStream = False
        '    Dim sVLine As String = oReader.ReadLine
        '    If sVLine Is Nothing Then Continue While

        '    Dim lIdx As Int32 = sVLine.IndexOf("|")
        '    If lIdx > -1 Then
        '        Dim sFile As String = sVLine.Substring(0, lIdx)
        '        Dim lVal As Int32 = CInt(sVLine.Substring(lIdx + 1))
        '    End If
        'End While


        If mbHadError = False Then
            AddTextLine("All files updated!")
        Else
            AddTextLine("Some errors were encountered... the process will begin again in 30 seconds.")
            Dim oT As New Threading.Thread(AddressOf BeginAgain)
            oT.Start()
        End If

        'If mlAppType = AppTypeCode.StarAlliances Then
        '    Kill(gsFilePath & "SAVersions.txt")
        'Else
        Try
            Kill(gsFilePath & "Versions.txt")
        Catch
        End Try
        Try
            Kill(gsFilePath & "Verify.txt")
        Catch
        End Try
        'End If

        If mbHadError = False Then
            mdtStartedUpdate = Now
            btnLaunch.Enabled = True
            btnConfig.Enabled = True
        End If

    End Sub

    Private Sub BeginAgain()
        Threading.Thread.Sleep(30000)
        DoUpdate()
    End Sub

    Private Sub EnsureDirectoryStructure(ByVal sFile As String)
        'ok, get our directories...
        Dim sDirs() As String
        Dim sPath As String
        Dim X As Int32

        sPath = gsFilePath

        sDirs = Split(sFile, "\")

        For X = 0 To UBound(sDirs) - 1
            If Exists(sPath & sDirs(X)) = False Then
                MkDir(sPath & sDirs(X))
            End If
            sPath = sPath & sDirs(X)
            If sPath.EndsWith("\") = False Then sPath &= "\"
        Next X

    End Sub

    Private Function StripFilePath(ByVal sFile As String) As String
        'Meshes\MeshTwo\file4.txt

        Dim lTemp As Int32
        Dim sResult As String

        sResult = ""
        lTemp = InStrRev(sFile, "\")

        If lTemp <> 0 Then
            sResult = Mid$(sFile, 1, lTemp)
            sFile = Mid$(sFile, lTemp, sFile.Length - lTemp)
        End If

        Return sResult
    End Function

    Private Sub DisplayNews()
        ''Open the file and display it
        'Dim oFS As IO.FileStream
        'Dim oReader As IO.StreamReader
        'Dim sInput As String

        'Dim sSB As System.Text.StringBuilder = New System.Text.StringBuilder()

        'If Exists(gsFilePath & "News.txt") Then
        '    oFS = New IO.FileStream(gsFilePath & "News.txt", IO.FileMode.Open)
        '    oReader = New IO.StreamReader(oFS)

        '    'txtNews.Text = ""
        '    While oReader.EndOfStream = False
        '        sInput = oReader.ReadLine()
        '        If sInput.StartsWith("======") = True AndAlso sInput.ToUpper.Contains("SERVER") = True Then
        '            sInput = Trim$(Replace$(sInput, "=", ""))
        '            lblServerStatus.Text = sInput

        '            If sInput.ToUpper.Contains("UP") = True Then
        '                lblServerStatus.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
        '                lblServerStatus.BackColor = System.Drawing.Color.FromArgb(255, 0, 92, 0)
        '            Else
        '                lblServerStatus.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
        '                lblServerStatus.BackColor = System.Drawing.Color.FromArgb(255, 92, 0, 0)
        '            End If
        '        Else
        '            sSB.AppendLine(sInput)
        '        End If
        '    End While

        '    oReader.Close()
        '    oReader.Dispose()
        '    oReader = Nothing
        '    oFS.Close()
        '    oFS.Dispose()
        '    oFS = Nothing
        'End If
        'txtNews.Text = sSB.ToString
        'Me.Refresh()

        If mlAppType = AppTypeCode.StarAlliances Then
            wbStatus.Navigate("http://www.star-alliances.com/phpnuke/html/SANotes.php")
        Else
            wbStatus.Navigate("http://www.beyondprotocol.com/phpnuke/html/BPNotes.php")
        End If

    End Sub

    Private Sub PickServer()

        Dim lTimes(mlServerIPUB) As Int32
        Dim lCurrMin As Int32 = 10000 '60 seconds should be plenty to download a news file

        If mlServerIPUB = 0 Then
            lTimes(0) = 0
            lCurrMin = 0
        Else
            For X As Int32 = 0 To mlServerIPUB
                Try
                    Dim oSock As New NetSock()
                    AddHandler oSock.onConnect, AddressOf TestSocketOnConnect
                    AddHandler oSock.onDataArrival, AddressOf TestSocketOnDataArrival

                    Dim lWaitCnt As Int32 = 0

                    mbTestSocketConnecting = True
                    oSock.Connect(msServerIPList(X), mlServerIPPort(X))
                    While mbTestSocketConnecting = True AndAlso lWaitCnt < 50
                        lWaitCnt += 1
                        Application.DoEvents()
                        Threading.Thread.Sleep(100)
                    End While

                    If mbTestSocketConnecting = True Then
                        lTimes(X) = Int32.MaxValue
                    Else
                        Dim yMsg(41) As Byte

                        System.BitConverter.GetBytes(UpdaterMsgCode.eLoginMsg).CopyTo(yMsg, 0)
                        System.Text.ASCIIEncoding.ASCII.GetBytes("Ep1caTe5ter").CopyTo(yMsg, 2)
                        System.Text.ASCIIEncoding.ASCII.GetBytes("ZaQ1XsW2#eDcVfR4").CopyTo(yMsg, 22)

                        oSock.SendData(yMsg)

                        Dim oSW As Stopwatch = Stopwatch.StartNew()
                        mbWaiting = True
                        oSock.SendData(MakeGetFileRequest("News.txt"))
                        While mbWaiting = True AndAlso lCurrMin > oSW.ElapsedMilliseconds
                            Application.DoEvents()
                        End While
                        If mbWaiting = True Then
                            lTimes(X) = Int32.MaxValue
                        Else
                            lTimes(X) = CInt(oSW.ElapsedMilliseconds)
                            lCurrMin = lTimes(X)
                        End If
                        mbWaiting = False
                        oSW.Stop()
                    End If

                    RemoveHandler oSock.onConnect, AddressOf TestSocketOnConnect
                    RemoveHandler oSock.onDataArrival, AddressOf TestSocketOnDataArrival
                    oSock.Disconnect()
                    oSock = Nothing
                Catch
                End Try
            Next X
        End If

        If lCurrMin = Int32.MaxValue Then
            AddTextLine("Unable to find a suitable server at this time. Please try again later.")
            Return
        End If
        For X As Int32 = 0 To mlServerIPUB
            If lTimes(X) = lCurrMin Then
                moServer = New NetSock()
                moServer.Connect(msServerIPList(X), mlServerIPPort(X))
                Exit For
            End If
        Next X
    End Sub

#Region " Test Socket Events "
    Private mbTestSocketConnecting As Boolean = False
    Private Sub TestSocketOnConnect(ByVal Index As Int32)
        mbTestSocketConnecting = False
    End Sub
    Private Sub TestSocketOnDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer)
        Dim iMsgID As Int16

        iMsgID = System.BitConverter.ToInt16(Data, 0)
        Select Case iMsgID
            Case UpdaterMsgCode.eFileUpdate
                HandleFileUpdate(Data)
                mbWaiting = False
        End Select
    End Sub
#End Region

#Region " Socket Events "
    Private Sub moServer_onConnect(ByVal Index As Integer) Handles moServer.onConnect
        Dim yMsg(41) As Byte

        System.BitConverter.GetBytes(UpdaterMsgCode.eLoginMsg).CopyTo(yMsg, 0)
        System.Text.ASCIIEncoding.ASCII.GetBytes("Ep1caTe5ter").CopyTo(yMsg, 2)
        System.Text.ASCIIEncoding.ASCII.GetBytes("ZaQ1XsW2#eDcVfR4").CopyTo(yMsg, 22)

        moServer.SendData(yMsg)

        Dim oThread As Threading.Thread = New Threading.Thread(AddressOf DoUpdate)
        oThread.Start()
    End Sub

    Private Sub moServer_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moServer.onDataArrival
        Dim iMsgID As Int16

        iMsgID = System.BitConverter.ToInt16(Data, 0)
        Select Case iMsgID
            Case UpdaterMsgCode.eFileUpdate
                HandleFileUpdate(Data)
                mbWaiting = False
        End Select
    End Sub

    Private Sub moServer_onDataContinued(ByVal lRemainingBytes As Integer) Handles moServer.onDataContinued
        Try
            If mlCurrentFileTotalLen <> 0 Then
                prgCurrent.Value = prgCurrent.Maximum - CInt((lRemainingBytes / mlCurrentFileTotalLen) * prgCurrent.Maximum)

                If mlTotalLen <> 0 Then
                    Dim lTempVal As Int32 = mlOverallProgress + (mlCurrentFileTotalLen - lRemainingBytes)
                    prgOverall.Value = CInt(prgOverall.Maximum * (lTempVal / mlTotalLen))
                End If
            End If
            prgCurrent.Refresh()
        Catch
            'no worries..
        End Try
    End Sub

    Private Sub moServer_onDisconnect(ByVal Index As Integer) Handles moServer.onDisconnect
        AddTextLine("Connection Closed.")
    End Sub

    Private Sub moServer_onError(ByVal Index As Integer, ByVal Description As String) Handles moServer.onError
        If Description.ToUpper.Contains("THE REQUESTED NAME IS VALID AND WAS FOUND IN THE DATABASE, BUT") = True Then
            Description &= vbCrLf & "This may be due to being behind a firewall or other security settings."
        End If
        AddTextLine("Error: " & Description)
    End Sub
#End Region
 
    Private Function MakeGetFileRequest(ByVal sFileName As String) As Byte()
        Dim yMsg(256) As Byte       '255 for the Name, 2 for the msg code
        Dim yName() As Byte = System.Text.ASCIIEncoding.ASCII.GetBytes(sFileName)

        'msLastFileRequested = sFileName
        mlCurrentFileLen = -1

        System.BitConverter.GetBytes(UpdaterMsgCode.eRequestFile).CopyTo(yMsg, 0)
        yName.CopyTo(yMsg, 2)

        Return yMsg
    End Function

    Private Sub HandleFileUpdate(ByVal yData() As Byte)
        Dim yTemp(254) As Byte
        Array.Copy(yData, 2, yTemp, 0, 255)
        Dim sFileName As String
        Dim lLen As Int32
        Dim X As Int32

		If mbSilentSaveToDisk = False Then AddTextLine(Space$(3) & "File Retrieved Saving to Disk...")

        Try
            lLen = yTemp.Length
            For X = 0 To yTemp.Length - 1
                If yTemp(X) = 0 Then
                    lLen = X
                    Exit For
                End If
            Next X
            sFileName = Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yTemp), 1, lLen)

            Dim oFS As IO.FileStream = New IO.FileStream(gsFilePath & sFileName, IO.FileMode.Create)
            Dim oBW As IO.BinaryWriter = New IO.BinaryWriter(oFS)

            Dim lDataLen As Int32
            If mlCurrentFileLen = -1 Then
                lDataLen = yData.Length - 258
            Else
                lDataLen = Math.Min(mlCurrentFileLen, yData.Length - 257)
            End If

            Dim yFile(lDataLen - 1) As Byte       '259
            Array.Copy(yData, 257, yFile, 0, lDataLen)    '258
            oBW.Write(yFile)
            oBW.Flush()
            oBW.Close()
            oBW = Nothing
            oFS.Close()
            oFS.Dispose()
            oFS = Nothing
        Catch ex As Exception
            AddTextLine(Space$(3) & "Could not save file: " & ex.Message)
            mbHadError = True
        End Try

    End Sub

    Private Function DoubleCheck() As Boolean
        Dim sININame As String = "BPCLIENT.INI"
        If mlAppType = AppTypeCode.StarAlliances Then sININame = "SACLIENT.INI"

        For X As Int32 = 0 To mlEntryUB
            Dim sInput As String = msEntries(X)

            If Trim$(sInput) <> "" AndAlso sInput <> vbCrLf Then
                mlEntryUB += 1
                ReDim Preserve msEntries(mlEntryUB)
                msEntries(mlEntryUB) = sInput

                Dim sEntries() As String = Split(sInput, "|")
                Dim sFileName As String = sEntries(0)

                Dim bDownload As Boolean = False
                If UBound(sEntries) > 1 Then
                    'Check if the file exists
                    If sEntries.GetUpperBound(0) >= 3 Then
                        If sEntries(3).ToUpper = "SA" Then
                            If mlAppType <> AppTypeCode.StarAlliances Then Continue For
                        ElseIf sEntries(3).ToUpper = "BP" Then
                            If mlAppType <> AppTypeCode.BeyondProtocol Then Continue For
                        End If
                    End If

                    If Exists(gsFilePath & sFileName) Then

                        'MSC - 1/26/2007 - do not update their INI settings...
                        If sFileName.ToUpper.Contains(sININame) = False Then

                            Dim oFSInfo As IO.FileInfo = New IO.FileInfo(gsFilePath & sFileName)
                            'If DateDiff("s", oFSInfo.CreationTime, CDate(sEntries(2))) <> 0 Then
                            If DateDiff("s", oFSInfo.CreationTime, GetLocaleSpecificDT(sEntries(2))) <> 0 Then
                                bDownload = True
                            End If
                            oFSInfo = Nothing
                        End If
                    Else
                        bDownload = True
                    End If
                Else
                    bDownload = True
                End If

                If bDownload = True Then
                    'mbHadError = True
                    'AddTextLine("  Out of Date: " & sFileName)
                    Return False
                End If
            End If
        Next X

        Return True
    End Function

    Private Sub btnLaunch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLaunch.Click

        If DoubleCheck() = False Then
            btnLaunch.Enabled = False
            btnConfig.Enabled = False
            BeginAgain()
            Return
        End If

        Dim sAppName As String = "BPClient.exe"
        If mlAppType = AppTypeCode.StarAlliances Then sAppName = "SAClient.exe"

        If Exists(gsFilePath & sAppName) Then
            Shell("""" & gsFilePath & sAppName & " "" EPICA_NORMAL_LAUNCH", AppWinStyle.NormalFocus, False, -1)
        End If
        Me.Close()
    End Sub

    Private Sub btnAccept_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAccept.Click
        grpEULA.Visible = False

        DisplayNews()

        gsFilePath = AppDomain.CurrentDomain.BaseDirectory
        If gsFilePath.EndsWith("\") = False Then gsFilePath &= "\"
        ChDir(gsFilePath)

        Dim sAppPrefix As String = "BP"
        If mlAppType = AppTypeCode.StarAlliances Then
            sAppPrefix = "SA"
        End If
        btnManual.Visible = (mlAppType = AppTypeCode.StarAlliances)

        Dim oINI As InitFile = New InitFile()
        Dim sImg As String = oINI.GetString("SETTINGS", "DisplayImage", "")

        If Exists(gsFilePath & "epica.bmp") = True Then
            Try
                Rename(gsFilePath & "epica.bmp", gsFilePath & sAppPrefix.ToLower & ".bmp")
            Catch ex As Exception
            End Try
        End If
        If Exists(gsFilePath & "epica.bm_") = True Then
            Try
                Rename(gsFilePath & "epica.bm_", gsFilePath & sAppPrefix.ToLower & ".bm_")
            Catch ex As Exception
            End Try
        End If
        If Exists(gsFilePath & "epicaclient.exe") = True Then
            Try
                Rename(gsFilePath & "epicaclient.exe", gsFilePath & sAppPrefix & "Client.exe")
            Catch ex As Exception
            End Try
        End If
        If Exists(gsFilePath & "epicaconfig.exe") = True Then
            Try
                Rename(gsFilePath & "epicaconfig.exe", gsFilePath & "ClientConfig.exe")
            Catch ex As Exception
            End Try
        End If
        If Exists(gsFilePath & "epicaclient.ini") = True Then
            Try
                Rename(gsFilePath & "epicaclient.ini", gsFilePath & sAppPrefix & "Client.ini")
            Catch ex As Exception
            End Try
        End If

        Dim sAddress As String = oINI.GetString("SETTINGS", "ServerAddress", "")
        Dim iPort As Integer = CInt(Val(oINI.GetString("SETTINGS", "ServerPort", "")))

        If sAddress <> "" Then
            mlServerIPUB += 1
            ReDim Preserve msServerIPList(mlServerIPUB)
            ReDim Preserve mlServerIPPort(mlServerIPUB)
            msServerIPList(mlServerIPUB) = sAddress
            mlServerIPPort(mlServerIPUB) = iPort
        End If

        Dim sEXEName As String = oINI.GetString("SETTINGS", "EXEName", "")
        oINI.WriteString("SETTINGS", "DisplayImage", sAppPrefix.ToLower & ".bmp")


        Dim bDone As Boolean = False
        Dim lIdx As Int32 = 0
        While bDone = False
            bDone = True
            lIdx += 1

            sAddress = oINI.GetString("SETTINGS", "ServerAddress_" & lIdx, "")
            iPort = CInt(Val(oINI.GetString("SETTINGS", "ServerPort_" & lIdx, "")))
            If sAddress <> "" AndAlso iPort <> 0 Then

                For X As Int32 = 0 To mlServerIPUB
                    If msServerIPList(X).ToUpper = sAddress.ToUpper Then
                        bDone = False
                        Continue While
                    End If
                Next

                mlServerIPUB += 1
                ReDim Preserve msServerIPList(mlServerIPUB)
                ReDim Preserve mlServerIPPort(mlServerIPUB)
                msServerIPList(mlServerIPUB) = sAddress
                mlServerIPPort(mlServerIPUB) = iPort
                bDone = False
            End If
        End While


        oINI = Nothing

        If sEXEName.ToUpper = "EPICACLIENT.EXE" Then
            sEXEName = sAppPrefix & "Client.exe"
        End If

        Me.Icon = Nothing
        AddTextLine("Initializing Application...")
        'If sImg <> "" AndAlso Exists(gsFilePath & sImg) = True Then
        '    'Me.PictureBox1.Image = Image.FromFile(gsFilePath & sImg)
        'End If

        prgCurrent.Value = 0
        prgOverall.Value = 0

        AddTextLine("Establishing Connection with Server...")

        'moServer = New NetSock()
        'moServer.Connect(sAddress, iPort)
        PickServer()
    End Sub

    Private Sub btnDecline_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDecline.Click
        Me.Close()
    End Sub

    Private Function INIUpdated() As Boolean
        If Exists(gsFilePath & "UpdaterClient.in_") = False Then Return False

        Dim oCurrentINI As InitFile = New InitFile()
        Dim oUpdatedINI As InitFile = New InitFile(gsFilePath & "UpdaterClient.in_")

        Dim sOriginalAddress As String = oCurrentINI.GetString("SETTINGS", "ServerAddress", "")
        Dim iOriginalPort As Int32 = CInt(Val(oCurrentINI.GetString("SETTINGS", "ServerPort", "")))
        Dim sOriginalEXE As String = oCurrentINI.GetString("SETTINGS", "EXEName", "")
        Dim sOriginalBMP As String = oCurrentINI.GetString("SETTINGS", "DisplayImage", "")

        Dim sNewAddress As String = oUpdatedINI.GetString("SETTINGS", "ServerAddress", "")
        Dim iNewPort As Int32 = CInt(Val(oUpdatedINI.GetString("SETTINGS", "ServerPort", "")))
        Dim sNewEXE As String = oUpdatedINI.GetString("SETTINGS", "EXEName", "")
        Dim sNewBMP As String = oUpdatedINI.GetString("SETTINGS", "DisplayImage", "")

        Dim bResult As Boolean = False

        If sNewAddress <> "" AndAlso sOriginalAddress.ToUpper <> sNewAddress.ToUpper Then
            bResult = True
            oCurrentINI.WriteString("SETTINGS", "ServerAddress", sNewAddress)
        End If
        If iNewPort <> 0 AndAlso iOriginalPort <> iNewPort Then
            bResult = True
            oCurrentINI.WriteString("SETTINGS", "ServerPort", iNewPort.ToString)
        End If

        Dim bDone As Boolean = False
        Dim lIdx As Int32 = 0
        While bDone = False
            bDone = True
            lIdx += 1

            Dim sAddress As String = oUpdatedINI.GetString("SETTINGS", "ServerAddress_" & lIdx, "")
            Dim iPort As Int32 = CInt(Val(oUpdatedINI.GetString("SETTINGS", "ServerPort_" & lIdx, "")))

            Dim sOrigAdd As String = oCurrentINI.GetString("SETTINGS", "ServerAddress_" & lIdx, "")
            Dim lOrigPort As Int32 = CInt(Val(oCurrentINI.GetString("SETTINGS", "ServerPort_" & lIdx, "")))

            If sAddress <> "" AndAlso iPort <> 0 Then
                'ok, overwrite it
                bDone = False
                If sAddress <> sOrigAdd Then
                    oCurrentINI.WriteString("SETTINGS", "ServerAddress_" & lIdx, sAddress)
                    bResult = True
                End If
                If lOrigPort <> iPort Then
                    oCurrentINI.WriteString("SETTINGS", "ServerPort_" & lIdx, iPort.ToString)
                    bResult = True
                End If
            Else
                'ok, new address/port is blank, clear the old one
                If sOrigAdd <> "" Then oCurrentINI.WriteString("SETTINGS", "ServerAddress_" & lIdx, "")
                If lOrigPort <> 0 Then oCurrentINI.WriteString("SETTINGS", "ServerPort_" & lIdx, "")
            End If
        End While

        oCurrentINI = Nothing
        oUpdatedINI = Nothing

        Return bResult
	End Function

	Private Sub btnConfig_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnConfig.Click
		If Exists(gsFilePath & "clientconfig.exe") = True Then
			Shell(gsFilePath & "clientconfig.exe", AppWinStyle.NormalFocus, False, -1)
		End If 
	End Sub 

    Private Sub DoRetryCurrent()
        Threading.Thread.Sleep(10000)
        moServer.SendData(MakeGetFileRequest(msCurrentFile))
    End Sub

    Private Sub moServer_ZeroReturnNeedRerequest() Handles moServer.ZeroReturnNeedRerequest

        AddTextLine("   Request delayed due to congestion.")
        AddTextLine("   Retrying in 10 seconds")
        Dim oT As New Threading.Thread(AddressOf DoRetryCurrent)
        oT.Start()

        'If msCurrentFile <> "" Then
        '    AddTextLine("  Request delayed... retrying later")
        '    If msRetryFiles Is Nothing Then ReDim msRetryFiles(-1)

        '    Dim bFound As Boolean = False
        '    For X As Int32 = 0 To msRetryFiles.GetUpperBound(0)
        '        If msRetryFiles(X) = msCurrentFile Then
        '            bFound = True
        '            Exit For
        '        End If
        '    Next X
        '    If bFound = False Then
        '        ReDim Preserve msRetryFiles(msRetryFiles.GetUpperBound(0) + 1)
        '        msRetryFiles(msRetryFiles.GetUpperBound(0)) = msCurrentFile
        '    End If
        'Else
        '    AddTextLine("  Aborting request...")
        'End If

        'mbWaiting = False
    End Sub

    Private Sub btnManual_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnManual.Click
        Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        If sPath.EndsWith("\") = False Then sPath &= "\"
        sPath &= "StarAlliancesUserGuide.pdf"
        If Exists(sPath) = False Then
            MsgBox("Unable to find Manual. It may not have downloaded yet.", MsgBoxStyle.OkOnly, "Missing Manual")
        Else
            Process.Start(sPath)
        End If
    End Sub
End Class
