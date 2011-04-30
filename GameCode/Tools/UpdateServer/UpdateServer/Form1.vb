Public Enum UpdaterMsgCode As Short
    eLoginMsg = 1
    eRequestFile
    eFileUpdate
End Enum

Public Class Form1
    Private Declare Function timeGetTime Lib "winmm.dll" Alias "timeGetTime" () As Int32

    Private Const ml_AUTO_DISCONNECT_TIME As Int32 = 50000
    Private Const ml_MAX_NOTE_LEN As Int32 = 5000

    Private gsFilePath As String

    Private msNotes As String = "Server Started..."
    Private mlNextRefresh As Int32

    Private Class UpdaterFileEntry
        Public sFileName As String
        Public yFileMsg() As Byte
    End Class
    Private moFiles() As UpdaterFileEntry
    Private mlFileUB As Int32 = -1

    'Sockets
    Private mlClientUB As Int32 = -1
    Private mlClientConnectTime() As Int32
    Private myClientUsed() As Byte
    Private moClients() As NetSock
    Private WithEvents moClientListener As NetSock

    Private msVersionsFileLines() As String
    Private mlVersionsFileLineUB As Int32 = -1

#Region "Client Listener Handling"
    Private Sub moClientListener_onConnect(ByVal Index As Integer) Handles moClientListener.onConnect
        '
    End Sub

    Private Sub moClientListener_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moClientListener.onConnectionRequest
        Dim X As Int32
        Dim lIdx As Int32 = -1
        On Error Resume Next
        msNotes = "Connection requested..." & vbCrLf & msNotes
        'Try
        For X = 0 To mlClientUB
            If moClients(X) Is Nothing Then
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            mlClientUB += 1
            ReDim Preserve moClients(mlClientUB)
            ReDim Preserve mlClientConnectTime(mlClientUB)
            ReDim Preserve myClientUsed(mlClientUB)

            lIdx = mlClientUB
        End If

        mlClientConnectTime(lIdx) = timeGetTime
        myClientUsed(lIdx) = 255
        moClients(lIdx) = New NetSock(oClient)
        moClients(lIdx).SocketIndex = lIdx

        msNotes = "Connection accepted (" & oClient.RemoteEndPoint.ToString & ") as Index " & lIdx & vbCrLf & msNotes


        'add event handlers
        AddHandler moClients(lIdx).onDataArrival, AddressOf moClients_onDataArrival
        AddHandler moClients(lIdx).onDisconnect, AddressOf moClients_onDisconnect
        AddHandler moClients(lIdx).onError, AddressOf moClients_onError

        'and then tell the socket to expect data
        moClients(lIdx).MakeReadyToReceive()
        'Catch ex As Exception
        '	msNotes = "ClientListener_ConnectionRequest:" & ex.Message & vbCrLf & msNotes
        '	Try
        '		If oClient Is Nothing = False Then oClient.Disconnect(True)
        '	Catch
        '	End Try
        '	oClient = Nothing
        'End Try
    End Sub

    Private Sub moClientListener_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moClientListener.onDataArrival
        '
    End Sub

    Private Sub moClientListener_onDisconnect(ByVal Index As Integer) Handles moClientListener.onDisconnect
        '
    End Sub

    Private Sub moClientListener_onError(ByVal Index As Integer, ByVal Description As String) Handles moClientListener.onError
        '
    End Sub

    Private Sub moClientListener_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moClientListener.onSendComplete
        '
    End Sub
#End Region

#Region "Client Connections Handling"
    Private Sub moClients_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer)
        Dim iMsgID As Int16

        iMsgID = System.BitConverter.ToInt16(Data, 0)
        Select Case iMsgID
            Case UpdaterMsgCode.eFileUpdate
            Case UpdaterMsgCode.eLoginMsg
                HandleLogin(Data, Index)
            Case UpdaterMsgCode.eRequestFile
                HandleGetFileRequest(Data, Index)
        End Select
    End Sub

    Private Sub moClients_onDisconnect(ByVal Index As Integer)
        On Error Resume Next
        myClientUsed(Index) = 0
        moClients(Index) = Nothing
        msNotes = "Disconnect (" & Index & ")" & vbCrLf & msNotes
    End Sub

    Private Sub moClients_onError(ByVal Index As Integer, ByVal Description As String)
        On Error Resume Next
        msNotes = "Error (" & Index & "): " & Description & vbCrLf & msNotes

        If Description.ToUpper.Contains("AN EXISTING CONNECTION") = True Then
            moClients(Index).Disconnect()
            moClients(Index) = Nothing
            myClientUsed(Index) = 0
        End If

        'if description.ToUpper.Contains("OUT
        
    End Sub
#End Region

    Private Sub HandleGetFileRequest(ByVal yData() As Byte, ByVal lIndex As Int32)
        Try
            Dim yTemp(254) As Byte
            Array.Copy(yData, 2, yTemp, 0, 255)
            Dim sFileName As String
            Dim lLen As Int32
            Dim X As Int32

            lLen = yTemp.Length
            For X = 0 To yTemp.Length - 1
                If yTemp(X) = 0 Then
                    lLen = X
                    Exit For
                End If
            Next X
            sFileName = Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yTemp), 1, lLen)

            Dim sTest As String = sFileName.ToUpper
            Dim bFound As Boolean = False
            For X = 0 To mlVersionsFileLineUB
                If msVersionsFileLines(X) = sTest Then
                    bFound = True
                    Exit For
                End If
            Next X
            If bFound = False Then Return

            If sFileName <> "" Then 'AndAlso Exists(gsFilePath & sFileName) Then
                'msNotes = "File Request (" & lIndex & "): " & sFileName & vbCrLf & msNotes
                Dim yResp() As Byte = GetFileAsMsg(sFileName)
                If yResp Is Nothing = False Then moClients(lIndex).SendData(yResp)
            End If
        Catch ex As Exception
            msNotes &= "HandleGetFileRequest: " & ex.Message & vbCrLf & msNotes
        End Try
    End Sub

    Private Function GetFileIndex(ByVal sFileName As String) As Int32
        Dim sTempFileName As String = sFileName.ToUpper
        For X As Int32 = 0 To mlFileUB
            If moFiles(X).sFileName = sTempFileName Then
                Return X
            End If
        Next X
        If mlFileUB = -1 Then ReDim moFiles(-1)

        Dim lIdx As Int32 = -1
        SyncLock moFiles
            lIdx = mlFileUB + 1
            ReDim Preserve moFiles(lIdx)
            moFiles(lIdx).sFileName = sFileName.ToUpper '.Trim
            moFiles(lIdx).yFileMsg = Nothing
            mlFileUB = lIdx
        End SyncLock
        Return lIdx
    End Function

    Private Function GetFileAsMsg(ByVal sFileName As String) As Byte()
        Try

            Dim lFileIndex As Int32 = GetFileIndex(sFileName)
            If lFileIndex > -1 Then
                If moFiles(lFileIndex).yFileMsg Is Nothing = False Then
                    'Dim yResult(muFiles(lFileIndex).yFileMsg.GetUpperBound(0)) As Byte
                    'muFiles(lFileIndex).yFileMsg.CopyTo(yResult, 0)
                    Return moFiles(lFileIndex).yFileMsg
                End If
            End If

            Dim yMsg() As Byte
            Dim yName() As Byte = System.Text.ASCIIEncoding.ASCII.GetBytes(sFileName)
            Dim yFile() As Byte

            'If Exists(gsFilePath & sFileName) = True Then
            Dim oFS As IO.FileStream = New IO.FileStream(gsFilePath & sFileName, IO.FileMode.Open)
            Dim oBRead As IO.BinaryReader = New IO.BinaryReader(oFS)

            yFile = oBRead.ReadBytes(CInt(oFS.Length))

            oBRead.Close()
            oBRead = Nothing
            oFS.Close()
            oFS.Dispose()
            oFS = Nothing

            ReDim yMsg(yFile.Length + 256)
            System.BitConverter.GetBytes(UpdaterMsgCode.eFileUpdate).CopyTo(yMsg, 0)
            yName.CopyTo(yMsg, 2)
            yFile.CopyTo(yMsg, 257)

            If lFileIndex > -1 AndAlso yMsg.Length > 1000000 Then
                ReDim moFiles(lFileIndex).yFileMsg(yMsg.GetUpperBound(0))
                yMsg.CopyTo(moFiles(lFileIndex).yFileMsg, 0)
                'muFiles(lFileIndex).yFileMsg = yMsg
            End If

            Return yMsg
            'Else
            'Return Nothing
            'End If
        Catch ex As Exception
            msNotes &= "GetFileAsMsg (" & sFileName & "): " & ex.Message & vbCrLf & msNotes
            Return Nothing
        End Try
    End Function

    Private Sub HandleLogin(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim yTemp(19) As Byte
        Dim sUserName As String
        Dim sPassword As String

        Dim lLen As Int32
        Dim X As Int32

        Try

            'First, the username
            Array.Copy(yData, 2, yTemp, 0, 20)
            lLen = yTemp.Length
            For X = 0 To yTemp.Length - 1
                If yTemp(X) = 0 Then
                    lLen = X
                    Exit For
                End If
            Next X
            sUserName = Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yTemp), 1, lLen)

            'Now the password
            Array.Copy(yData, 22, yTemp, 0, 20)
            lLen = yTemp.Length
            For X = 0 To yTemp.Length - 1
                If yTemp(X) = 0 Then
                    lLen = X
                    Exit For
                End If
            Next X
            sPassword = Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yTemp), 1, lLen)

            If sUserName = "Ep1caTe5ter" AndAlso sPassword = "ZaQ1XsW2#eDcVfR4" Then
                myClientUsed(lIndex) = 254
                'msNotes &= vbCrLf & "Login Accepted (" & lIndex & ")"
                'msNotes = "Login Accepted (" & lIndex & ")" & vbCrLf & msNotes
            Else
                'msNotes &= vbCrLf & "Login Rejected (" & lIndex & ")"
                msNotes = "Login Rejected (" & lIndex & ")" & vbCrLf & msNotes
                moClients(lIndex).Disconnect()
                myClientUsed(lIndex) = 0
            End If
        Catch ex As Exception
            msNotes &= "HandleLogin: " & ex.Message & vbCrLf & msNotes
        End Try
    End Sub

    Private Sub tmrCleanup_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrCleanup.Tick
        Dim lNow As Int32 = timeGetTime
        tmrCleanup.Enabled = False
        Try
            For X As Int32 = 0 To mlClientUB
                If myClientUsed(X) = 255 Then
                    If lNow - mlClientConnectTime(X) > ml_AUTO_DISCONNECT_TIME Then
                        'msNotes &= vbCrLf & "Auto-Disconnect (" & X & ")"
                        msNotes = "Auto-Disconnect (" & X & ")" & vbCrLf & msNotes
                        moClients(X).Disconnect()
                        myClientUsed(X) = 0
                    End If
                End If
            Next X

            If msNotes.Length > ml_MAX_NOTE_LEN Then
                msNotes = msNotes.Substring(0, ml_MAX_NOTE_LEN)
            End If

            mlNextRefresh -= 1
            If mlNextRefresh < 0 Then Me.Refresh()
        Catch
        End Try

        tmrCleanup.Enabled = True
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        gsFilePath = AppDomain.CurrentDomain.BaseDirectory
        If gsFilePath.EndsWith("\") = False Then gsFilePath &= "\"

        Try
            Dim oFS As New IO.FileStream(gsFilePath & "versions.txt", IO.FileMode.Open)
            Dim oRead As New IO.StreamReader(oFS)
            While oRead.EndOfStream = False
                Dim sLine As String = oRead.ReadLine
                'UpdaterClient.ex_|57344|11/16/2007 11:40:24 AM
                Dim lIdx As Int32 = sLine.IndexOf("|"c)
                If lIdx > -1 Then
                    sLine = sLine.Substring(0, lIdx)
                End If
                mlVersionsFileLineUB += 1
                ReDim Preserve msVersionsFileLines(mlVersionsFileLineUB)
                msVersionsFileLines(mlVersionsFileLineUB) = sLine.Trim.ToUpper
            End While
            oRead.Close()
            oRead.Dispose()
            oFS.Close()
            oFS.Dispose()
            oRead = Nothing
            oFS = Nothing
        Catch ex As Exception
            MsgBox("Initialization.LoadVersionsFile: " & ex.Message)
        End Try

        mlVersionsFileLineUB += 1
        ReDim Preserve msVersionsFileLines(mlVersionsFileLineUB)
        msVersionsFileLines(mlVersionsFileLineUB) = "NEWS.TXT"
        mlVersionsFileLineUB += 1
        ReDim Preserve msVersionsFileLines(mlVersionsFileLineUB)
        msVersionsFileLines(mlVersionsFileLineUB) = "VERSIONS.TXT"
        mlVersionsFileLineUB += 1
        ReDim Preserve msVersionsFileLines(mlVersionsFileLineUB)
        msVersionsFileLines(mlVersionsFileLineUB) = "UPDATERUPDATER.EXE"
        mlVersionsFileLineUB += 1
        ReDim Preserve msVersionsFileLines(mlVersionsFileLineUB)
        msVersionsFileLines(mlVersionsFileLineUB) = "UPDATERCLIENT.IN_"
        mlVersionsFileLineUB += 1
        ReDim Preserve msVersionsFileLines(mlVersionsFileLineUB)
        msVersionsFileLines(mlVersionsFileLineUB) = "UPDATERCLIENT.EX_"


        mlFileUB = mlVersionsFileLineUB
        ReDim moFiles(mlVersionsFileLineUB)
        For X As Int32 = 0 To mlFileUB
            moFiles(X) = New UpdaterFileEntry()
            With moFiles(X)
                .sFileName = msVersionsFileLines(X).ToUpper
                '.yFileMsg = GetFileAsMsg(.sFileName)
                'GetFileAsMsg(.sFileName)
                .yFileMsg = Nothing
            End With
        Next X


        Dim oINI As New InitFile
        Dim lListenPort As Int32 = CInt(oINI.GetString("SOCKET", "ListenPort", "7400"))

        moClientListener = New NetSock()
        moClientListener.PortNumber = lListenPort
        moClientListener.Listen()

        tmrCleanup.Enabled = True
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        moClientListener.StopListening()
        moClientListener = Nothing

        For X As Int32 = 0 To mlClientUB
            If myClientUsed(X) <> 0 Then
                moClients(X).Disconnect()
                moClients(X) = Nothing
                myClientUsed(X) = 0
            End If
        Next X

        Application.Exit()
    End Sub

    Private Sub Form1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Me.Paint
        If msNotes <> txtNotes.Text Then
            txtNotes.Text = msNotes
            txtNotes.SelectionStart = msNotes.Length - 1
        End If
        mlNextRefresh = 10
    End Sub

    Private Sub btnDisconnectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisconnectAll.Click
        On Error Resume Next
        For X As Int32 = 0 To mlClientUB
            If myClientUsed(X) <> 0 Then moClients(X).Disconnect()
        Next X
    End Sub

    Private Sub btnResetMsgs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResetMsgs.Click
        On Error Resume Next
        For X As Int32 = 0 To mlFileUB
            moFiles(X).yFileMsg = Nothing
        Next X
    End Sub

    Private Sub btnDownloaders_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDownloaders.Click
        On Error Resume Next
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To mlClientUB
            If myClientUsed(X) <> 0 Then
                If moClients(X) Is Nothing = False Then lCnt += 1
            End If
        Next X
        msNotes = "Concurrent Logins: " & lCnt & vbCrLf & msNotes

    End Sub
End Class
