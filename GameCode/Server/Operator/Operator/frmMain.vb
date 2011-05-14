Public Class frmMain

    Private mlBoxOperatorCnt As Int32 = 0

    'Public Sub AddEventLine(ByVal sEvent As String)
    '    If txtEvents.Text Is Nothing = False AndAlso txtEvents.Text.Length > 500000 Then
    '        txtEvents.Text = txtEvents.Text.Substring(0, 500000)
    '    End If
    '    txtEvents.Text = sEvent & vbCrLf & txtEvents.Text
    '    Me.Refresh()
    'End Sub

    Private Delegate Sub doUpdate(ByVal sEvent As String)

    Private Sub delegateUpdateEvents(ByVal sEvent As String)
        If txtEvents.Text Is Nothing = False AndAlso txtEvents.Text.Length > 500000 Then txtEvents.Text = txtEvents.Text.Substring(0, 500000)
        txtEvents.Text = sEvent & vbCrLf & txtEvents.Text
        Me.Refresh()
    End Sub
    Public Sub AddEventLine(ByVal sEvent As String)
        txtEvents.Invoke(New doUpdate(AddressOf delegateUpdateEvents), sEvent)
    End Sub

    Private Sub btnShutdown_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShutdown.Click
        If btnShutdown.Text.ToUpper = "CONFIRM" Then
            btnShutdown.Enabled = False

            gbRunning = False

            Dim yMsg(1) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eServerShutdown).CopyTo(yMsg, 0)
            For X As Int32 = 0 To goMsgSys.mlServerUB
                If goMsgSys.oServerObject(X) Is Nothing = False Then
                    If goMsgSys.oServerObject(X).lConnectionType = ConnectionType.ePrimaryServerApp Then
                        goMsgSys.oServerObject(X).oSocket.SendData(yMsg)
                    End If
                End If
            Next X
            For X As Int32 = 0 To goMsgSys.mlServerUB
                If goMsgSys.oServerObject(X) Is Nothing = False Then
                    If goMsgSys.oServerObject(X).lConnectionType = ConnectionType.eEmailServerApp Then
                        goMsgSys.oServerObject(X).oSocket.SendData(yMsg)
                    End If
                End If
            Next X

            'TODO: send to all box operators the shutdown sequence



            'save our senate
            Try
                Dim sText As String = Senate.GetSaveObjectText()
                If sText Is Nothing = False AndAlso sText <> "" Then
                    Dim oComm As New OleDb.OleDbCommand(sText, goCN)
                    oComm.ExecuteNonQuery()
                End If
            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "Saving Senate Error: " & ex.Message)
            End Try

            'CloseConn()
        Else
            btnShutdown.Text = "Confirm"
        End If
	End Sub

	Private Sub frmMain_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		Application.Exit()
		End
	End Sub

	Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
		goMsgSys.AcceptingClients = False
		goMsgSys.AcceptingServers = False
		goMsgSys.ForceDisconnectAll()
		goMsgSys = Nothing
	End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Form.CheckForIllegalCrossThreadCalls = False
        gfrmDisplayForm = Me
    End Sub

    Private Function ParseCommandLine() As Boolean
        Try
            If My.Application.CommandLineArgs.Count > 0 Then
                Dim sValue As String = My.Application.CommandLineArgs(0)
                If sValue Is Nothing = False AndAlso IsNumeric(sValue) = True Then
                    If CInt(sValue) <> 0 Then
                        gyBackupOperator = eyOperatorState.EmergencyOperator
                    End If
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox("ParseCommandLine: " & ex.Message)
            Return False
        End Try
    End Function

    Private Sub frmMain_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        ParseCommandLine()

        AddEventLine("Initializing Logging...")
        If InitializeEventLogging() = False Then
            AddEventLine("Could not initialize logging...")
            Return
        End If
        LogEvent(LogEventType.Informational, "Initializing Data...")
        If LoadDataFromDB() = False Then
            LogEvent(LogEventType.CriticalError, "Unable to load database data!")
            Return
        End If
        LogEvent(LogEventType.Informational, "Initializing Message System...")
        goMsgSys = New MsgSystem()
        goMsgSys.AcceptingServers = True
        goMsgSys.AcceptingClients = True
        goMsgSys.ListenForAdmins()
        LogEvent(LogEventType.Informational, "Server ready... waiting for at least " & txtExpectedBoxes.Text & " box operators to connect...")

        cboStatus.Items.Clear()
        cboStatus.Items.Add("Active")
        cboStatus.Items.Add("Suspended")
        cboStatus.Items.Add("Inactive")
        cboStatus.Items.Add("Banned")
        cboStatus.Items.Add("Open Beta")
        cboStatus.Items.Add("Trial")

        Dim oMainThread As New Threading.Thread(AddressOf MainLoop)
        oMainThread.Start()
    End Sub

    Private Sub btnSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSet.Click
        glExpectedBoxCnt = CInt(Val(txtExpectedBoxes.Text))
        LogEvent(LogEventType.Informational, "New Expected Box Count: " & glExpectedBoxCnt.ToString)
    End Sub

    Private Delegate Sub doCheckState()
    Private Sub DoSetCheckState()
        chkAcceptClients.Checked = goMsgSys.bAcceptNewClients
        If chkAcceptClients.Checked = True Then chkAcceptClients.CheckState = CheckState.Checked Else chkAcceptClients.CheckState = CheckState.Unchecked
    End Sub
    Public Sub SetAcceptClientsCheck()
        If chkAcceptClients.InvokeRequired = True Then
            chkAcceptClients.Invoke(New doCheckState(AddressOf DoSetCheckState))
        Else : DoSetCheckState()
        End If
    End Sub

    Public Sub IncrementAndCheckBoxes()
        mlBoxOperatorCnt += 1
        If mlBoxOperatorCnt >= CInt(Val(txtExpectedBoxes.Text)) AndAlso (gyBackupOperator = eyOperatorState.LonelyOperator OrElse gyBackupOperator = eyOperatorState.MainOperatorWithBackup) Then
            'Ok, we're ready
            SpawnNeighborhoods()
        End If
    End Sub

    Private Sub btnSetExtAddy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSetExtAddy.Click
        gsExternalAddress = txtExtAddy.Text
    End Sub

	Private Sub chkAcceptClients_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAcceptClients.CheckedChanged
		goMsgSys.bAcceptNewClients = chkAcceptClients.Checked
	End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If IsNumeric(txtUserID.Text) = False Then
            Try
                Dim sUserName As String = txtUserID.Text.ToUpper
                For X As Int32 = 0 To glPlayerUB
                    If glPlayerIdx(X) > -1 Then
                        If goPlayer(X).sPlayerName = sUserName Then
                            txtUserID.Text = goPlayer(X).ObjectID.ToString
                            Exit For
                        End If
                    End If
                Next X
            Catch
                LogEvent(LogEventType.CriticalError, "Don't be an idiot.")
                Return
            End Try
        End If
        Dim lPlayerID As Int32 = CInt(Val(txtUserID.Text))
        If lPlayerID < 1 Then
            LogEvent(LogEventType.CriticalError, "Could not find user")
            Return
        End If
        Dim lStatus As Int32 = 0

        Select Case cboStatus.Text.ToUpper
            Case "ACTIVE"
                lStatus = AccountStatusType.eActiveAccount
            Case "SUSPENDED"
                lStatus = AccountStatusType.eSuspendedAccount
            Case "INACTIVE"
                lStatus = AccountStatusType.eInactiveAccount
            Case "BANNED"
                lStatus = AccountStatusType.eBannedAccount
            Case "OPEN BETA"
                lStatus = AccountStatusType.eOpenBetaAccount
            Case "TRIAL"
                lStatus = AccountStatusType.eTrialAccount
            Case Else
                Return
        End Select

        For X As Int32 = 0 To glPlayerUB
            If glPlayerIdx(X) = lPlayerID Then
                goPlayer(X).AccountStatus = lStatus
                goPlayer(X).SaveStatusOnly()
                LogEvent(LogEventType.Informational, "Status updated for player ID: " & lPlayerID)

                'Now, let's send our update to all servers
                Dim yMsg(69) As Byte
                Dim lPos As Int32 = 0
                With goPlayer(X)
                    System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerDetails).CopyTo(yMsg, lPos) : lPos += 2
                    System.BitConverter.GetBytes(.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                    .PlayerUserName.CopyTo(yMsg, lPos) : lPos += 20
                    .PlayerPassword.CopyTo(yMsg, lPos) : lPos += 20
                    System.BitConverter.GetBytes(.AccountStatus).CopyTo(yMsg, lPos) : lPos += 4
                    .sPlayerNameProper = BytesToString(.PlayerName)
                    .sPlayerName = .sPlayerNameProper.ToUpper
                    StringToBytes(.sPlayerNameProper).CopyTo(yMsg, lPos) : lPos += 20
                    goMsgSys.SendToPrimaryServers(yMsg)
                End With

                Exit For
            End If
        Next X
    End Sub

    Private Sub btnLegalMins_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLegalMins.Click
        GeoSpawner.LegalizeMineralLocs()
        LogEvent(LogEventType.Informational, "Mineral Locs Legalized.")
    End Sub

    Private Sub btnSpawnNeighborhood_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSpawnNeighborhood.Click
        If btnSpawnNeighborhood.Text = "Confirm" Then
            btnSpawnNeighborhood.Text = "Spawn Neighborhood"
            GeoSpawner.IncrementGenerateCounter()
        Else
            btnSpawnNeighborhood.Text = "Confirm"
        End If
    End Sub

    Private Sub btnRegionDist_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRegionDist.Click
        If uSpawnRequests Is Nothing Then Return

        LogEvent(LogEventType.Informational, "Distribution List")
        For X As Int32 = 0 To uSpawnRequests.GetUpperBound(0)
            If uSpawnRequests(X).lConnectionType = ConnectionType.eRegionServerApp Then
                LogEvent(LogEventType.Informational, "  Server At " & uSpawnRequests(X).oBoxOperator.sIPAddress & " (" & X & ")")
                For Y As Int32 = 0 To uSpawnRequests(X).lSpawnUB
                    Dim oSystem As SolarSystem = GetEpicaSystem(uSpawnRequests(X).lSpawnID(Y))
                    If oSystem Is Nothing = False Then
                        LogEvent(LogEventType.Informational, "    " & BytesToString(oSystem.SystemName))
                    End If
                Next Y
            End If
        Next X
    End Sub

    Private Sub RespawnTheFlux()
        Try
            Dim sSQL As String = "select * from tblplayercommarchive where msgtitle = 'Strange Event'"
            Dim oComm As New OleDb.OleDbCommand(sSQL, goCN)
            Dim oData As OleDb.OleDbDataReader = oComm.ExecuteReader()

            While oData.Read
                Dim sBody As String = CStr(oData("MsgBody"))
                'Another report has come in from a unit encountering massive levels of quantum flux causing a catastrophic event.    
                'The unit in question, -GCA Manus Deus PA, seemed to be destroyed almost instantly.    
                'Witnesses say that an object was seen falling from the explosion towards the surface of Domjba IV.

                Dim lIdx As Int32 = sBody.IndexOf("towards the surface of")
                If lIdx > -1 Then
                    Dim sName As String = sBody.Substring(lIdx).Trim()
                    sName = sName.Replace("towards the surface of", "").Trim()
                    sName = sName.ToUpper()

                    For X As Int32 = 0 To glPlanetUB
                        If glPlanetIdx(X) > -1 Then
                            Dim oPlanet As Planet = goPlanet(X)
                            If oPlanet Is Nothing = False Then
                                Dim sPName As String = BytesToString(oPlanet.PlanetName)
                                If sPName.ToUpper = sName Then

                                    Dim yMsg(21) As Byte
                                    Dim lPos As Int32 = 0
                                    System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, lPos) : lPos += 2
                                    System.BitConverter.GetBytes(41991I).CopyTo(yMsg, lPos) : lPos += 4
                                    System.BitConverter.GetBytes(ObjectType.eMineral).CopyTo(yMsg, lPos) : lPos += 2
                                    oPlanet.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                                    System.BitConverter.GetBytes(1I).CopyTo(yMsg, lPos) : lPos += 4
                                    System.BitConverter.GetBytes(1I).CopyTo(yMsg, lPos) : lPos += 4

                                    If oPlanet.ParentSystem Is Nothing = False Then
                                        If oPlanet.ParentSystem.oPrimaryServer Is Nothing = False Then
                                            oPlanet.ParentSystem.oPrimaryServer.oSocket.SendData(yMsg)
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next X
                End If
            End While

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
End Class

