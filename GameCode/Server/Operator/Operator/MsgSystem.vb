Option Strict On
Public Enum ConnectionType As Int32
    eUnknown = -1
    eNoConnection = 0
    eBackupOperator = 1
    eClientConnection = 2
    ePrimaryServerApp = 3
    eRegionServerApp = 4
    ePathfindingServerApp = 5
    eEmailServerApp = 6
    eBoxOperator = 7
    eOperator = 8
    eArenaServer = 9
End Enum
Public Class MsgSystem

    Private moServers() As NetSock
    Public mlServerUB As Int32 = -1
    Public oServerObject() As ServerObject

    Private moBOPServers() As ServerObject
    Private mlBOPServerUB As Int32 = -1

    Private moClients() As NetSock
    Private mlClientPlayer() As Int32
    Private mlClientUB As Int32 = -1

    Private moAdmins() As NetSock
    Private mlAdminPlayer() As Int32
    Private mlAdminUB As Int32 = -1

    Private WithEvents moAdminListener As NetSock
    Private WithEvents moServerListener As NetSock
    Private WithEvents moClientListener As NetSock

    Private mbAcceptingClients As Boolean = False
    Private mbAcceptingServers As Boolean = False

    Public bAcceptNewClients As Boolean = False

    'Pointer to the server object for the backup operator if I am the primary operator
    Private oBackupOperator As ServerObject = Nothing
    'Socket connection for operator
    Private WithEvents moOperator As NetSock = Nothing

    Public Property AcceptingClients() As Boolean
        Get
            Return mbAcceptingClients
        End Get
        Set(ByVal Value As Boolean)
            Dim lPort As Int32
            If mbAcceptingClients <> Value Then
                mbAcceptingClients = Value
                If mbAcceptingClients Then
                    Dim oINI As New InitFile()
                    Try
                        lPort = CInt(Val(oINI.GetString("SETTINGS", "ClientListenerPort", "0")))
                        If lPort = 0 Then Err.Raise(-1, "AcceptingClients", "Could not get Client Listen Port Number from INI")
                        moClientListener = Nothing
                        moClientListener = New NetSock()
                        moClientListener.PortNumber = lPort
                        moClientListener.Listen()
                    Catch
                        MsgBox("An error was caught in " & Err.Source & ":" & vbCrLf & Err.Description, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, Err.Source)
                        mbAcceptingClients = False
                    Finally
                        oINI = Nothing
                    End Try
                Else
                    'ok, stop listening
                    moClientListener.StopListening()
                End If
            End If
        End Set
    End Property

    Public Property AcceptingServers() As Boolean
        Get
            Return mbAcceptingServers
        End Get
        Set(ByVal Value As Boolean)
            Dim lPort As Int32
            If mbAcceptingServers <> Value Then
                mbAcceptingServers = Value
                If mbAcceptingServers Then
                    Dim oINI As New InitFile()
                    Try
                        lPort = CInt(Val(oINI.GetString("SETTINGS", "ServerListenerPort", "0")))
                        If lPort = 0 Then Err.Raise(-1, "AcceptingServers", "Could not get Server Listen Port Number from INI")
                        moServerListener = Nothing
                        moServerListener = New NetSock()
                        moServerListener.PortNumber = lPort
                        moServerListener.Listen()
                    Catch
                        MsgBox("An error was caught in " & Err.Source & ":" & vbCrLf & Err.Description, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, Err.Source)
                        mbAcceptingServers = False
                    Finally
                        oINI = Nothing
                    End Try
                Else
                    'ok, stop listening
                    moServerListener.StopListening()
                End If
            End If
        End Set
    End Property

    Public Sub ListenForAdmins()
        Dim oINI As New InitFile()
        Try
            Dim lPort As Int32 = CInt(Val(oINI.GetString("SETTINGS", "AdminListenerPort", "7396")))
            If lPort = 0 Then Err.Raise(-1, "ListenForAdmins", "Could not get Admin Listen Port Number from INI")
            moAdminListener = Nothing
            moAdminListener = New NetSock()
            moAdminListener.PortNumber = lPort
            moAdminListener.Listen()
        Catch
            MsgBox("An error was caught in " & Err.Source & ":" & vbCrLf & Err.Description, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, Err.Source)
        Finally
            oINI = Nothing
        End Try
    End Sub

    Public Sub ForceDisconnectAll()
        Dim X As Int32

        On Error Resume Next

        For X = 0 To mlClientUB
            If moClients(X) Is Nothing = False Then
                moClients(X).Disconnect()
                Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(X))
                If oPlayer Is Nothing = False Then oPlayer.TotalPlayTime += CInt(DateDiff(DateInterval.Second, oPlayer.LastLogin, Now))
                moClients(X) = Nothing
            End If
        Next X

        For X = 0 To mlServerUB
            If moServers(X) Is Nothing = False Then
                moServers(X).Disconnect()
                moServers(X) = Nothing
            End If
        Next X

    End Sub

    Private mbConnectingToOperator As Boolean = False
    Private mbConnectedToOperator As Boolean = False
    Public Function ConnectToOperator() As Boolean
        Dim mlTimeout As Int32 = 10000

        Try
            mbConnectingToOperator = True
            mbConnectedToOperator = False
            moOperator = New NetSock()
            moOperator.Connect(gsMainOperatorIP, glMainOperatorPort)
            Dim sw As Stopwatch = Stopwatch.StartNew
            While mbConnectedToOperator = False AndAlso (sw.ElapsedMilliseconds < mlTimeout) AndAlso mbConnectingToOperator = True
                Application.DoEvents()
            End While
            sw.Stop()
            sw = Nothing
            mbConnectingToOperator = False
            'If mbConnectedToOperator = False Then
            'Err.Raise(-1, "EstablishConnection", "Unable to connect to OPERATOR: Connection timed out!")
            'End If
        Catch
            mbConnectedToOperator = False
        End Try

        Return mbConnectedToOperator
    End Function

#Region "Server Listener Handling"
    Private Sub moServerListener_onConnect(ByVal Index As Integer) Handles moServerListener.onConnect
        '
    End Sub

    Private Sub moServerListener_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moServerListener.onConnectionRequest
        Dim X As Int32
        Dim lIdx As Int32 = -1

        If AcceptingServers Then
            Try
                For X = 0 To mlServerUB
                    If moServers(X) Is Nothing AndAlso oServerObject(X) Is Nothing Then
                        lIdx = X
                        Exit For
                    End If
                Next X

                If lIdx = -1 Then
                    mlServerUB += 1
                    ReDim Preserve moServers(mlServerUB)
                    ReDim Preserve oServerObject(mlServerUB)
                    lIdx = mlServerUB
                End If

                moServers(lIdx) = New NetSock(oClient)
                oServerObject(lIdx) = New ServerObject()
                oServerObject(lIdx).oSocket = moServers(lIdx)
                oServerObject(lIdx).bSocketConnected = True
                moServers(lIdx).SocketIndex = lIdx
                moServers(lIdx).lSpecificID = lIdx

                LogEvent(LogEventType.Informational, "Server Connected")

                'add event handlers
                AddHandler moServers(lIdx).onConnect, AddressOf moServers_onConnect
                AddHandler moServers(lIdx).onDataArrival, AddressOf moServers_onDataArrival
                AddHandler moServers(lIdx).onDisconnect, AddressOf moServers_onDisconnect
                AddHandler moServers(lIdx).onError, AddressOf moServers_onError

                'and then tell the socket to expect data
                moServers(lIdx).MakeReadyToReceive()

                'If gyBackupOperator <> eyOperatorState.BackupOperator Then
                '    Dim yMsg(7) As Byte
                '    System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, 0)
                '    System.BitConverter.GetBytes(lIdx).CopyTo(yMsg, 2)
                '    System.BitConverter.GetBytes(Int16.MinValue).CopyTo(yMsg, 6)
                '    ForwardToBackupOperator(yMsg, -1)
                'End If
            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "ServerListener.ConnectionRequest: " & ex.Message)
            End Try
        Else
            LogEvent(LogEventType.Informational, "Server Connection Denied: Not Accepting!")
        End If
    End Sub

    Private Sub moServerListener_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moServerListener.onDataArrival
        '
    End Sub

    Private Sub moServerListener_onDisconnect(ByVal Index As Integer) Handles moServerListener.onDisconnect
        '
    End Sub

    Private Sub moServerListener_onError(ByVal Index As Integer, ByVal Description As String) Handles moServerListener.onError
        '
    End Sub

    Private Sub moServerListener_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moServerListener.onSendComplete
        '
    End Sub
#End Region
#Region "Client Listener Handling"
    Private Sub moClientListener_onConnect(ByVal Index As Integer) Handles moClientListener.onConnect
        '
    End Sub

    Private Sub moClientListener_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moClientListener.onConnectionRequest
        Dim X As Int32
        Dim lIdx As Int32 = -1

        If AcceptingClients = True AndAlso bAcceptNewClients = True Then
            Try
                For X = 0 To mlClientUB
                    If moClients(X) Is Nothing Then
                        lIdx = X
                        Exit For
                    End If
                Next X

                If lIdx = -1 Then
                    mlClientUB += 1
                    ReDim Preserve moClients(mlClientUB)
                    ReDim Preserve mlClientPlayer(mlClientUB)
                    lIdx = mlClientUB
                End If

                moClients(lIdx) = New NetSock(oClient)
                moClients(lIdx).SocketIndex = lIdx
                moClients(lIdx).lSpecificID = -1
                mlClientPlayer(lIdx) = -1

                LogEvent(LogEventType.Informational, "Client Connected")

                'add event handlers
                AddHandler moClients(lIdx).onConnect, AddressOf moClients_onConnect
                AddHandler moClients(lIdx).onDataArrival, AddressOf moClients_onDataArrival
                AddHandler moClients(lIdx).onDisconnect, AddressOf moClients_onDisconnect
                AddHandler moClients(lIdx).onError, AddressOf moClients_onError

                'and then tell the socket to expect data
                moClients(lIdx).MakeReadyToReceive()
            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "ClientConnectionRequest: " & ex.Message)
                If moClients(lIdx) Is Nothing = False Then
                    moClients(lIdx).Disconnect()
                    moClients(lIdx) = Nothing
                End If
            End Try
        Else
            LogEvent(LogEventType.Informational, "Client Connection Denied: Not Accepting")
            oClient.Disconnect(True)
        End If
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
#Region "Admin Listener Handling"
    Private Sub moAdminListener_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moAdminListener.onConnectionRequest
        Dim X As Int32
        Dim lIdx As Int32 = -1

        Try
            For X = 0 To mlAdminUB
                If moAdmins(X) Is Nothing Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                mlAdminUB += 1
                ReDim Preserve moAdmins(mlAdminUB)
                ReDim Preserve mlAdminPlayer(mlAdminUB)
                lIdx = mlAdminUB
            End If

            moAdmins(lIdx) = New NetSock(oClient)
            moAdmins(lIdx).SocketIndex = lIdx
            moAdmins(lIdx).lSpecificID = -1
            mlAdminPlayer(lIdx) = -1

            LogEvent(LogEventType.Informational, "Admin Connected: " & moAdmins(lIdx).GetRemoteIP)

            'add event handlers
            AddHandler moAdmins(lIdx).onDataArrival, AddressOf moAdmins_onDataArrival
            AddHandler moAdmins(lIdx).onDisconnect, AddressOf moAdmins_onDisconnect
            AddHandler moAdmins(lIdx).onError, AddressOf moAdmins_onError

            'and then tell the socket to expect data
            moAdmins(lIdx).MakeReadyToReceive()
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "AdminConnectionRequest: " & ex.Message)
            If moAdmins(lIdx) Is Nothing = False Then
                moAdmins(lIdx).Disconnect()
                moAdmins(lIdx) = Nothing
            End If
        End Try
    End Sub
#End Region
#Region "Server Connections Handling"
    Private Sub moServers_onConnect(ByVal Index As Integer)
        Try
            If oServerObject(Index) Is Nothing = False Then
                If oServerObject(Index).lConnectionType = ConnectionType.eBoxOperator Then
                    LogEvent(LogEventType.Informational, "Box Operator Connected")
                    oServerObject(Index).SendPendingSpawnRequests()
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub moServers_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket)
        '
    End Sub

    Private Sub moServers_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer)
        'ok, gotta get the data...
        'Data() contains our data... let's dissect it...

        If gyBackupOperator = eyOperatorState.BackupOperator Then
            'I am the backup operator, so the server index precedes the msg...
            Index = System.BitConverter.ToInt32(Data, 0)
            'Now, make the final msg
            Dim yTemp(Data.GetUpperBound(0) - 4) As Byte
            Array.Copy(Data, 4, yTemp, 0, Data.Length - 4)
            Data = yTemp
        End If

        Dim iMsgID As Int16
        iMsgID = System.BitConverter.ToInt16(Data, 0)

        Select Case iMsgID
            Case GlobalMessageCode.eCrossPrimaryBudgetAdjust
                HandleCrossPrimaryBudgetAdjust(Data, Index)
            Case GlobalMessageCode.eRequestPlayerBudget
                Dim lPlayer As Int32 = System.BitConverter.ToInt32(Data, 2)
                Dim oPlayer As Player = GetEpicaPlayer(lPlayer)
                If oPlayer Is Nothing = False Then
                    If oPlayer.oBudget Is Nothing = True Then
                        oPlayer.oBudget = New PlayerBudgetData()
                        oPlayer.oBudget.oPlayer = oPlayer
                    End If
                    oPlayer.oBudget.RequestDetails(Index, -1, iMsgID)
                End If
            Case GlobalMessageCode.eOperatorRequestPassThru
                HandleOperatorRequestPassThru(Data, Index)
            Case GlobalMessageCode.eOperatorResponsePassThru
                HandleOperatorResponsePassThru(Data, Index)
            Case GlobalMessageCode.eRegisterDomain
                HandleServerIdentification(Data, Index)
            Case GlobalMessageCode.eDomainServerReady
                HandleServerReady(Data, Index)
            Case GlobalMessageCode.eRequestObject
                HandleKeepAliveMsg(Data, Index)
            Case GlobalMessageCode.eAddObjectCommand
                HandleAddObjectCommand(Data, Index)
            Case GlobalMessageCode.ePlayerAliasConfig
                HandlePlayerAliasConfig(Data, Index)
            Case GlobalMessageCode.eRemovePlayerRel
                HandleRemovePlayerRel(Data)
            Case GlobalMessageCode.ePlayerConnectedPrimary
                HandlePlayerConnectedPrimary(Data, Index)
            Case GlobalMessageCode.eEnvironmentDomain
                HandleEnvironmentDomainMsg(Data, Index)
            Case GlobalMessageCode.eBackupOperatorSyncMsg
                HandleSyncMessageFromBackup(Data)
            Case GlobalMessageCode.eSetPlayerSpecialAttribute
                HandleSetPlayerSpecialAttribute(Data, Index)
            Case GlobalMessageCode.eForcePrimarySync
                SendToPrimaryServers(Data)
            Case GlobalMessageCode.eAddSenateProposal
                Dim yResp() As Byte = Senate.HandleCreateProposalMsg(Data)
                If yResp Is Nothing = False Then moServers(Index).SendData(yResp)
            Case GlobalMessageCode.eAddSenateProposalMessage
                Senate.HandleAddProposalMessage(Data)
            Case GlobalMessageCode.eGetSenateObjectDetails
                Dim bSendLenAppended As Boolean = False
                Dim yResp() As Byte = Senate.HandleGetSenateObjectDetails(Data, bSendLenAppended)
                If yResp Is Nothing = False Then
                    If bSendLenAppended = True Then
                        moServers(Index).SendLenAppendedData(yResp)
                    Else
                        moServers(Index).SendData(yResp)
                    End If
                End If
            Case GlobalMessageCode.eUpdatePlayerVote
                Senate.HandlePlayerSetVote(Data)
            Case GlobalMessageCode.ePlayerIsDead
                HandlePlayerIsDead(Data, Index)
            Case GlobalMessageCode.eUpdatePlanetOwnership
                HandleUpdatePlanetOwnership(Data)
            Case GlobalMessageCode.ePlayerPrimaryOwner
                HandlePlayerPrimaryOwner(Data, Index)
            Case GlobalMessageCode.eForwardToPlayerAtPrimary
                HandleForwardToPlayerAtPrimary(Data, Index)
            Case GlobalMessageCode.eEntityChangingPrimary
                HandleEntityChangingPrimary(Data, Index)
            Case GlobalMessageCode.eAddFormation
                HandleAddFormation(Data, Index)
            Case GlobalMessageCode.eBudgetResponse
                Dim lPlayerID As Int32 = System.BitConverter.ToInt32(Data, 2)
                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                If oPlayer Is Nothing = False AndAlso oPlayer.oBudget Is Nothing = False Then
                    oPlayer.oBudget.HandlePrimaryResponse(Data, Index)
                End If
            Case GlobalMessageCode.eRequestServerDetails
                HandleServerRequestServerDetails(Data, Index)
            Case GlobalMessageCode.eRequestGlobalPlayerScores
                HandleRequestGlobalPlayerScores(Data, Index)
            Case GlobalMessageCode.ePlayerCurrentEnvironment
                HandleServerGetPlayerCurrentEnvironment(Data, Index)
            Case GlobalMessageCode.eChangingEnvironment
                SendToPrimaryServers(Data)
            Case GlobalMessageCode.eRequestPlayerStartLoc
                HandlePrimarySpawnLocUpdate(Data, Index)
            Case GlobalMessageCode.eCheckPrimaryReady
                GeoSpawner.HandleCheckPrimaryReady(Data)
            Case GlobalMessageCode.ePlayerTitleChange
                HandlePlayerTitleChange(Data)
            Case GlobalMessageCode.eChatMessage
                HandleChatMessageCmd(Data, Index)
            Case GlobalMessageCode.eSenateStatusReport
                Dim yResp() As Byte = Senate.HandleSenateStatusReport(Data)
                If yResp Is Nothing = False Then moServers(Index).SendData(yResp)
            Case GlobalMessageCode.eSetEntityName
                Senate.PlayerRenamed(Data)
        End Select
    End Sub

    Private Sub moServers_onDisconnect(ByVal Index As Integer)
        Try
            LogEvent(LogEventType.Informational, "Server " & Index & " disconnected. Type: " & oServerObject(Index).lConnectionType.ToString)
            oServerObject(Index).bSocketConnected = False
        Catch
        End Try
    End Sub

    Private Sub moServers_onError(ByVal Index As Integer, ByVal Description As String)
        LogEvent(LogEventType.Informational, "Server Connection Error (" & Index & "): " & Description)
    End Sub

    Private Sub moServers_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer)
        '
    End Sub

#End Region
#Region "Client Connections Handling"
    Private Sub moClients_onConnect(ByVal Index As Integer)
        '
    End Sub

    Private Sub moClients_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket)
        '
    End Sub

    Private Sub moClients_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer)
        Dim iMsgID As Int16

        iMsgID = System.BitConverter.ToInt16(Data, 0)

        Select Case iMsgID
            Case GlobalMessageCode.eClientVersion
                HandleClientVersion(Data, Index)
            Case GlobalMessageCode.ePlayerLoginRequest
                HandleLoginRequest(Data, Index)
            Case GlobalMessageCode.ePlayerInitialSetup
                If gyBackupOperator = eyOperatorState.BackupOperator Then
                    HandleSetPlayerInitialSetup(Data, Index)
                Else : HandlePlayerInitialSetup(Data, Index)
                End If
            Case GlobalMessageCode.ePathfindingConnectionInfo
                HandleClientGetPrimaryConnData(Data, Index)
        End Select
    End Sub

    Private Sub moClients_onDisconnect(ByVal Index As Integer)
        moClients(Index) = Nothing
        LogEvent(LogEventType.Informational, "Client " & Index & " Disconnected")
    End Sub

    Private Sub moClients_onError(ByVal Index As Integer, ByVal Description As String)
        On Error Resume Next
        LogEvent(LogEventType.Informational, "Client Connection Error (" & Index & "): " & Description)
        If InStr(Description.ToLower, "an existing connection was forcibly closed by the remote host", CompareMethod.Binary) <> 0 Then
            moClients(Index).Disconnect()
            moClients(Index) = Nothing
        ElseIf Err.Number = 5 Then
            If moClients(Index) Is Nothing = False Then moClients(Index).Disconnect()
            moClients(Index) = Nothing
        End If
    End Sub

    Private Sub moClients_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer)
        '
    End Sub

    Private Sub HandleClientGetPrimaryConnData(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        If mlClientPlayer(lIndex) <> lPlayerID Then
            LogEvent(LogEventType.PossibleCheat, "Invalid PlayerID for requesting PrimaryConnData: " & mlClientPlayer(lIndex) & " is mine, passed in: " & lPlayerID)
            Return
        End If

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = True Then Return

        Dim oPrimary As ServerObject = Nothing
        If iEnvirTypeID = ObjectType.ePlanet Then
            Dim oPlanet As Planet = GetEpicaPlanet(lEnvirID)
            If oPlanet Is Nothing = False Then
                oPrimary = oPlanet.ParentSystem.oPrimaryServer
            End If
        ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
            Dim oSystem As SolarSystem = GetEpicaSystem(lEnvirID)
            If oSystem Is Nothing = False Then
                oPrimary = oSystem.oPrimaryServer
            End If
        Else
            LogEvent(LogEventType.PossibleCheat, "Invalid TypeID for PrimaryConnData (" & iEnvirTypeID & "): " & mlClientPlayer(lIndex))
            Return
        End If
        If oPrimary Is Nothing = False Then
            lPos = 0
            Dim yResp(25) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.ePathfindingConnectionInfo).CopyTo(yResp, lPos) : lPos += 2
            Dim bFound As Boolean = False
            With oPrimary
                If .uListeners Is Nothing = False Then
                    For Y As Int32 = 0 To .uListeners.GetUpperBound(0)
                        If .uListeners(Y).lConnectionType = ConnectionType.eClientConnection Then
                            bFound = True
                            If .uListeners(Y).sIPAddress.ToUpper = "EXTERNALIPADDY" Then
                                StringToBytes(gsExternalAddress).CopyTo(yResp, lPos) : lPos += 20
                            Else
                                StringToBytes(.uListeners(Y).sIPAddress).CopyTo(yResp, lPos) : lPos += 20
                            End If
                            System.BitConverter.GetBytes(.uListeners(Y).lPortNumber).CopyTo(yResp, lPos) : lPos += 4
                            Exit For
                        End If
                    Next Y
                End If
            End With
            If bFound = False Then
                LogEvent(LogEventType.CriticalError, "HandleLoginRequest (client): Could not find Primary Listener")
            Else
                moClients(lIndex).SendData(yResp)
            End If
        Else
            LogEvent(LogEventType.PossibleCheat, "Invalid GUID for PrimaryConnData (" & lEnvirID & ", " & iEnvirTypeID & "): " & mlClientPlayer(lIndex))
        End If
    End Sub

#End Region
#Region "  Admin Connections Handling  "
    'IMPORTANT: THESE MUST DIRECTLY MATCH THE GlobalMessageCode ENUM IN BPAdminConsole!!!
    Public Enum AdminMessageCode As Int16

        RequestPlayerDetails
        SearchForPlayer
        RequestDxDiag
        RequestHistories
        RequestSupportHistoryDetails
        RequestForumPostDetails
        RequestBugDetails
        RequestGNSDetails
        RequestMissionDetails
        RequestEmailDetails

        'Colony Requests
        RequestColoniesForPlayer
        ClearColonyProductionQueues
        AbandonColony
        ChangeColonyName
        ChangeColonyTaxRate
        SetOfflineInvulnerability

        'Agent Requests
        RequestAgentsForPlayer
        RequestAgentLocation
        DismissAgent
        RemoveAgentFromMission
        ChangeAgentStatus

        'Station Requests
        RequestStationsForPlayer
        DestroyStation
        ForceLaunchAll
        ClearStationProductionQueues

        'Special Techs Available
        RequestSpecialTechsForPlayer
        CompleteResearchOnSpecialTech
        NumberOfTechsProduced
        StopAllSpecialTechResearchStarted
        WhenSpecialTechResearchStarted
        ForceSpecialTechLinkTest

        'Components
        RequestComponentsForPlayer
        StopResearchOnComponents
        CompleteResearchOnComponet

        'Diplomacy Contacts
        RequestDiplomacyContacts
        SetRelationToContact
        SetContactRelationTo

        'Missions
        RequestMissionsForPlayer
        SetCurrentPhaseOfMission
        CancelMission

        'Units
        RequestEnvironmentsForPlayer
        RequestUnitsForPlayer
        RequestUnitsByEnvironment
        RepairUnit
        FinishProductionOfUnit
        CancelProductionOfUnit
        MoveUnitToNewLocation
        SetUnitName
        ToggleUnitGuildAssetFlag
        RequestUnitCargo
        AddItemToUnitCargo
        RemoveItemFromUnitCargo
        LaunchAllUnits
        'Routes
        RequestRouteForUnit
        AddDestinationToRoute
        RemoveDestinationFromRoute
        ClearRoute
        BeginRoute
        ForceNextDestinationInRoute
        LastDockItemAtThatDestination

        'Minerals
        RequestMineralsForColony
        RequestPlayerTechKnowledge
        AdjustMineralQuantity
        AddMineralQuantity
        GivePlayerTechKnowledge
        CompletePlayerTechResearch
        StopPlayerTechResearch
        RenewPlayerTechKnowledge
        RequestPlayerBudgetSummary

        'User Login
        ValidateUserCredentials
        RequestPlayerAliases
        RequestAliasPermissions
        RequestPlayerSubscriptions
        RequestComponentDetails

        SetPlayerCredits = 92   'TODO: Not really accurate
    End Enum
    Private Sub moAdmins_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer)
        Dim iMsgID As Int16

        iMsgID = System.BitConverter.ToInt16(Data, 0)

        'TODO: verify the user is a player... verify their username and password... associate the playerid to mlAdminPlayer
        'TODO: Do other messages here...

        Dim yResp() As Byte = Nothing

        Select Case iMsgID
            Case AdminMessageCode.RequestDxDiag
                yResp = AdminConsoleMsg.HandleRequestDxDiag(Data, mlAdminPlayer(Index))
            Case AdminMessageCode.ValidateUserCredentials
                mlAdminPlayer(Index) = AdminConsoleMsg.HandleAdminLogin(Data)
                ReDim yResp(5)
                System.BitConverter.GetBytes(AdminMessageCode.ValidateUserCredentials).CopyTo(yResp, 0)
                System.BitConverter.GetBytes(mlAdminPlayer(Index)).CopyTo(yResp, 2)
            Case AdminMessageCode.SearchForPlayer
                yResp = AdminConsoleMsg.HandleSearchForPlayer(Data, mlAdminPlayer(Index))
            Case AdminMessageCode.RequestPlayerDetails
                yResp = AdminConsoleMsg.HandleRequestPlayerDetails(Data, mlAdminPlayer(Index))
            Case AdminMessageCode.RequestColoniesForPlayer
                yResp = AdminConsoleMsg.HandleGetPlayerColonies(Data, mlAdminPlayer(Index))
            Case AdminMessageCode.RequestPlayerBudgetSummary
                Dim lPlayerID As Int32 = System.BitConverter.ToInt32(Data, 2)
                LogEvent(LogEventType.Informational, "Admin " & mlAdminPlayer(Index) & " requesting player budget for player " & lPlayerID)
                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                If oPlayer Is Nothing = False Then
                    If oPlayer.oBudget Is Nothing = True Then
                        oPlayer.oBudget = New PlayerBudgetData()
                        oPlayer.oBudget.oPlayer = oPlayer
                    End If
                    oPlayer.oBudget.RequestDetails(-1, Index, iMsgID)
                End If
            Case AdminMessageCode.RequestStationsForPlayer
                yResp = AdminConsoleMsg.HandleRequestStationsForPlayer(Data, mlAdminPlayer(Index))
            Case AdminMessageCode.RequestAgentsForPlayer
                yResp = AdminConsoleMsg.HandleRequestAgentsForPlayer(Data, mlAdminPlayer(Index))
            Case AdminMessageCode.RequestSpecialTechsForPlayer
                yResp = AdminConsoleMsg.HandleGetPlayerSpecialTechs(Data, mlAdminPlayer(Index), moAdmins(Index))
            Case AdminMessageCode.RequestComponentsForPlayer
                yResp = AdminConsoleMsg.HandleRequestComponentsForPlayer(Data, mlAdminPlayer(Index))
            Case AdminMessageCode.RequestComponentDetails
                yResp = AdminConsoleMsg.HandleRequestComponentDetails(Data, mlAdminPlayer(Index))
            Case AdminMessageCode.RequestDiplomacyContacts
                yResp = AdminConsoleMsg.HandleRequestDiplomacyContacts(Data, mlAdminPlayer(Index))
            Case AdminMessageCode.RequestUnitsByEnvironment
                yResp = AdminConsoleMsg.HandleRequestEntitiesByEnvironment(Data, mlAdminPlayer(Index), moAdmins(Index))
            Case AdminMessageCode.RequestUnitCargo
                yResp = AdminConsoleMsg.HandleRequestEntityContents(Data, mlAdminPlayer(Index), moAdmins(Index))
            Case AdminMessageCode.RequestRouteForUnit
                yResp = AdminConsoleMsg.HandleRequestRouteForUnit(Data, mlAdminPlayer(Index))
            Case AdminMessageCode.SetPlayerCredits
                yResp = AdminConsoleMsg.HandleSetPlayerCredits(Data, mlAdminPlayer(Index))
        End Select

        If yResp Is Nothing = False Then moAdmins(Index).SendData(yResp)
    End Sub

    Private Sub moAdmins_onDisconnect(ByVal Index As Integer)
        moAdmins(Index) = Nothing
        LogEvent(LogEventType.Informational, "Admin " & Index & " Disconnected")
    End Sub

    Private Sub moAdmins_onError(ByVal Index As Integer, ByVal Description As String)
        On Error Resume Next
        LogEvent(LogEventType.Informational, "Admin Connection Error (" & Index & "): " & Description)
        If InStr(Description.ToLower, "an existing connection was forcibly closed by the remote host", CompareMethod.Binary) <> 0 Then
            moAdmins(Index).Disconnect()
            moAdmins(Index) = Nothing
        ElseIf Err.Number = 5 Then
            If moAdmins(Index) Is Nothing = False Then moAdmins(Index).Disconnect()
            moAdmins(Index) = Nothing
        End If
    End Sub
#End Region
#Region "Operator Connection Handling"

    Private myConnectedToOperatorBefore As Byte = 0
    Private Sub moOperator_onConnect(ByVal Index As Integer) Handles moOperator.onConnect
        Dim yMsg(18) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eRegisterDomain).CopyTo(yMsg, lPos) : lPos += 2
        'put box operator id here
        lPos += 4
        System.BitConverter.GetBytes(ConnectionType.eBackupOperator).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(Process.GetCurrentProcess.Id).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(0I).CopyTo(yMsg, lPos) : lPos += 4
        yMsg(lPos) = myConnectedToOperatorBefore : lPos += 1
        myConnectedToOperatorBefore = 1
        moOperator.SendData(yMsg)
        mbConnectedToOperator = True
    End Sub

    Private Sub moOperator_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moOperator.onConnectionRequest
        '
    End Sub

    Private Sub moOperator_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moOperator.onDataArrival
        'ok, gotta get the data...
        'Data() contains our data... let's dissect it...

        'If gyBackupOperator = eyOperatorState.BackupOperator Then
        '    'I am the backup operator, so the server index precedes the msg...
        '    If Data.Length > 3 Then
        '        Index = System.BitConverter.ToInt32(Data, 0)
        '        'Now, make the final msg
        '        Dim yTemp(Data.GetUpperBound(0) - 4) As Byte
        '        Array.Copy(Data, 4, yTemp, 0, Data.Length - 4)
        '        Data = yTemp
        '    End If
        'End If

        Dim iMsgID As Int16
        iMsgID = System.BitConverter.ToInt16(Data, 0)

        Select Case iMsgID
            Case GlobalMessageCode.eRegisterDomain
                HandleServerIdentification(Data, Index)
            Case GlobalMessageCode.eDomainServerReady
                'HandleServerReady(Data, Index)
                HandleBackupOpServerReady(Data, Index)
            Case GlobalMessageCode.eRequestObject
                HandleKeepAliveMsg(Data, Index)
            Case GlobalMessageCode.eAddObjectCommand
                HandleAddObjectCommand(Data, Index)
            Case GlobalMessageCode.ePlayerAliasConfig
                HandlePlayerAliasConfig(Data, Index)
            Case GlobalMessageCode.eRemovePlayerRel
                HandleRemovePlayerRel(Data)
            Case GlobalMessageCode.eEnvironmentDomain
                HandleEnvironmentDomainMsg(Data, Index)
            Case GlobalMessageCode.eBackupOperatorSyncMsg
                HandleSyncMessageFromOperator(Data)
            Case GlobalMessageCode.eUpdatePlayerDetails
                HandleOperatorUpdatePlayer(Data)
        End Select
    End Sub

    Private Sub moOperator_onDisconnect(ByVal Index As Integer) Handles moOperator.onDisconnect
        Try
            LogEvent(LogEventType.Informational, "Server " & Index & " disconnected. Type: " & oServerObject(Index).lConnectionType.ToString)
            oServerObject(Index).bSocketConnected = False
        Catch
        End Try
    End Sub

    Private Sub moOperator_onError(ByVal Index As Integer, ByVal Description As String) Handles moOperator.onError
        LogEvent(LogEventType.Informational, "Server Connection Error (" & Index & "): " & Description)
    End Sub

    Private Sub moOperator_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moOperator.onSendComplete
        '
    End Sub
#End Region

    Private Sub HandleChatMessageCmd(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lFromPlayer As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yType As Byte = yData(lPos) : lPos += 1
        Dim lLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim sMsg As String = GetStringFromBytes(yData, lPos, lLen) : lPos += lLen

        sMsg = sMsg.Trim

        Dim oPlayer As Player = GetEpicaPlayer(lFromPlayer)
        If oPlayer Is Nothing = True Then Return

        Dim sSQL As String = ""
        Dim lTime As Int32 = GetDateAsNumber(Now)
        If sMsg.StartsWith("/csrloggedin") = True Then
            sSQL = "INSERT INTO CSRTrans (CSRName, LogTime, LogType) VALUES ('" & BytesToString(oPlayer.PlayerName) & "', " & lTime & ", 1)"
        ElseIf sMsg.StartsWith("/csrloggedout") = True Then
            sSQL = "INSERT INTO CSRTrans (CSRName, LogTime, LogType) VALUES ('" & BytesToString(oPlayer.PlayerName) & "', " & lTime & ", 0)"
        End If

        Dim oComm As Odbc.OdbcCommand = Nothing
        Try
            oComm = New Odbc.OdbcCommand(sSQL, goSuiteCN)
            oComm.ExecuteNonQuery()
            moServers(lIndex).SendData(yData)
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Could set csrlogged status: " & sMsg)
        End Try
        If oComm Is Nothing = False Then oComm.Dispose()
        oComm = Nothing
    End Sub

    Private Sub HandleOperatorUpdatePlayer(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        LoadSinglePlayer(lPlayerID) 
    End Sub

    Public Sub ForwardToBackupOperator(ByRef yMsg() As Byte, ByVal lServerIndex As Int32)
        Try
            If gyBackupOperator = eyOperatorState.BackupOperator Then Return
            If oBackupOperator Is Nothing Then
                'ok, attempt to find the backup operator in our server list
                Dim lUB As Int32 = -1
                If oServerObject Is Nothing = False Then lUB = Math.Min(mlServerUB, oServerObject.GetUpperBound(0))
                For X As Int32 = 0 To lUB
                    If oServerObject(X) Is Nothing = False Then
                        If oServerObject(X).lConnectionType = ConnectionType.eBackupOperator Then
                            If oServerObject(X).oSocket Is Nothing = False AndAlso oServerObject(X).oSocket.IsConnected = True Then
                                oBackupOperator = oServerObject(X)
                                Exit For
                            End If
                        End If
                    End If
                Next X

                If oBackupOperator Is Nothing = False Then
                    'Resync the backup operator with everything!!!
                    SendResyncMessages(oBackupOperator.oSocket)

                    gyBackupOperator = eyOperatorState.MainOperatorWithBackup
                End If
            Else
                'Dim yForward(yMsg.Length + 3) As Byte
                'System.BitConverter.GetBytes(lServerIndex).CopyTo(yForward, 0)
                'yMsg.CopyTo(yForward, 4)
                oBackupOperator.oSocket.SendData(yMsg)
            End If
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "ForwardToBackupOperator: " & ex.Message)
            If oBackupOperator Is Nothing = False Then
                If oBackupOperator.oSocket Is Nothing = False Then
                    If oBackupOperator.oSocket.IsConnected = False Then
                        'We have lost the backup operator
                        gyBackupOperator = eyOperatorState.LonelyOperator   'I am now a lonely operator
                        oBackupOperator = Nothing
                        SendSupportEmail("The backup operator has become unresponsive. Switching to Lonely Operator Mode. Waiting for Backup Operator to return.", "DSE Backup Operator Fail")
                    End If
                End If
            End If
        End Try

    End Sub

    Private Sub HandlePlayerTitleChange(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yTitle As Byte = yData(lPos) : lPos += 1
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then oPlayer.yPlayerTitle = yTitle
    End Sub

    Private Sub HandleServerIdentification(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lBoxOperatorID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lType As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lProcessID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        oServerObject(lIndex).lConnectionType = CType(lType, ConnectionType)
        oServerObject(lIndex).lProcessID = lProcessID

        With oServerObject(lIndex)

            .sIPAddress = moServers(lIndex).GetRemoteIP

            .lBoxOperatorID = lBoxOperatorID
            If .lBoxOperatorID <> -1 Then .oBoxOperator = GetBoxOperator(.lBoxOperatorID)

            Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            ReDim .uListeners(lCnt - 1)
            For X As Int32 = 0 To lCnt - 1
                .uListeners(X).sIPAddress = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                .uListeners(X).lPortNumber = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .uListeners(X).lConnectionType = CType(System.BitConverter.ToInt32(yData, lPos), ConnectionType) : lPos += 4
            Next X
        End With

        'Now, before proceeding, check if we already have this connection in backup op mode
        If gyBackupOperator = eyOperatorState.BackupOperator OrElse gyBackupOperator = eyOperatorState.EmergencyOperator Then
            'ok, check my bops
            For X As Int32 = 0 To mlBOPServerUB
                If moBOPServers(X) Is Nothing = False AndAlso moBOPServers(X).lBoxOperatorID = lBoxOperatorID AndAlso moBOPServers(X).lProcessID = lProcessID AndAlso moBOPServers(X).lConnectionType = lType Then
                    If moServers(lIndex).GetRemoteIP = moBOPServers(X).sIPAddress Then
                        'ok, found our moBopServer... so copy our data over
                        'moBOPServers(x).oRelatedServer
                        'oServerObject(lIndex).oRelatedServer = moBOPServers(X).oRelatedServer
                        If moBOPServers(X).oRelatedServer Is Nothing = False Then
                            For Y As Int32 = 0 To mlServerUB
                                If oServerObject(Y) Is Nothing = False Then
                                    If oServerObject(Y).lProcessID = moBOPServers(X).oRelatedServer.lProcessID AndAlso oServerObject(Y).lBoxOperatorID = moBOPServers(X).oRelatedServer.lBoxOperatorID AndAlso oServerObject(Y).sIPAddress = moBOPServers(X).oRelatedServer.sIPAddress Then
                                        oServerObject(lIndex).oRelatedServer = oServerObject(Y)
                                        Exit For
                                    End If
                                End If
                            Next Y
                        End If

                        'moBOPServers(x).lSpawnUB
                        oServerObject(lIndex).lSpawnUB = moBOPServers(X).lSpawnUB
                        ReDim oServerObject(lIndex).lSpawnID(oServerObject(lIndex).lSpawnUB)
                        ReDim oServerObject(lIndex).iSpawnTypeID(oServerObject(lIndex).lSpawnUB)
                        'moBOPServers(x).lSpawnID
                        'moBOPServers(x).iSpawnTypeID
                        For Y As Int32 = 0 To oServerObject(lIndex).lSpawnUB
                            oServerObject(lIndex).lSpawnID(Y) = moBOPServers(X).lSpawnID(Y)
                            oServerObject(lIndex).iSpawnTypeID(Y) = moBOPServers(X).iSpawnTypeID(Y)

                            If oServerObject(lIndex).iSpawnTypeID(Y) = ObjectType.eSolarSystem Then
                                Dim oSys As SolarSystem = GetEpicaSystem(oServerObject(lIndex).lSpawnID(Y))
                                If oSys Is Nothing = False Then oSys.oPrimaryServer = oServerObject(lIndex)
                            End If
                        Next Y

                        'and then return
                        Return
                    End If
                End If
            Next X
        End If

        'Ok, now pull our related data from the spawn request
        Dim oRelServer As ServerObject = Nothing
        For X As Int32 = 0 To lSpawnRequestUB
            If uSpawnRequests(X).lBoxOperatorID = oServerObject(lIndex).lBoxOperatorID AndAlso uSpawnRequests(X).lConnectionType = oServerObject(lIndex).lConnectionType Then
                If uSpawnRequests(X).ySpawnRequestState = 1 Then
                    'set the state to fulfilled
                    uSpawnRequests(X).ySpawnRequestState = 2

                    'Ok, now copy our stuff over from the spawn request
                    oServerObject(lIndex).oRelatedServer = uSpawnRequests(X).oRelatedServer

                    oServerObject(lIndex).lSpawnUB = uSpawnRequests(X).lSpawnUB
                    ReDim oServerObject(lIndex).lSpawnID(uSpawnRequests(X).lSpawnUB)
                    ReDim oServerObject(lIndex).iSpawnTypeID(uSpawnRequests(X).lSpawnUB)
                    For Y As Int32 = 0 To uSpawnRequests(X).lSpawnUB
                        oServerObject(lIndex).lSpawnID(Y) = uSpawnRequests(X).lSpawnID(Y)
                        oServerObject(lIndex).iSpawnTypeID(Y) = uSpawnRequests(X).iSpawnTypeID(Y)
                    Next Y

                    Exit For
                End If
            End If
        Next X

        Select Case oServerObject(lIndex).lConnectionType
            Case ConnectionType.ePrimaryServerApp
                FinalizePrimaryIdentification(oServerObject(lIndex), lIndex)
            Case ConnectionType.eRegionServerApp
                FinalizeRegionIdentification(oServerObject(lIndex))
            Case ConnectionType.ePathfindingServerApp
                FinalizePathfindingIdentification(oServerObject(lIndex))
            Case ConnectionType.eBoxOperator
                'reinstantiate our object as box operator class
                oServerObject(lIndex) = New BoxOperator(oServerObject(lIndex))
                If oServerObject(lIndex).lBoxOperatorID = -1 Then
                    'Ok, assign it a new box operator ID
                    oServerObject(lIndex).lBoxOperatorID = GetNextBoxOperatorID()
                    oServerObject(lIndex).oBoxOperator = CType(oServerObject(lIndex), BoxOperator)
                End If
                If gyBackupOperator <> eyOperatorState.BackupOperator Then
                    Dim yResp(5) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eRegisterDomain).CopyTo(yResp, 0)
                    System.BitConverter.GetBytes(oServerObject(lIndex).lBoxOperatorID).CopyTo(yResp, 2)
                    oServerObject(lIndex).oSocket.SendData(yResp)
                    'Increment our connected box operator count and check to make sure we have enoguh
                    gfrmDisplayForm.IncrementAndCheckBoxes()

                    'Store our resulting box operator ID so that the backup operator gets the good result
                    System.BitConverter.GetBytes(oServerObject(lIndex).lBoxOperatorID).CopyTo(yData, 2)
                End If
            Case ConnectionType.eBackupOperator
                oBackupOperator = oServerObject(lIndex)
                gyBackupOperator = eyOperatorState.MainOperatorWithBackup
                Dim yRebound As Byte = yData(lPos) : lPos += 1
                If yRebound = 0 Then
                    'Im in control, send the status to the backup
                    SendResyncMessages(oBackupOperator.oSocket)
                Else
                    'I was lost but now im found, 
                End If
        End Select

        If oServerObject(lIndex).oBoxOperator Is Nothing = False Then
            oServerObject(lIndex).oBoxOperator.AddServerProc(oServerObject(lIndex))
        End If

        If oServerObject(lIndex).lConnectionType <> ConnectionType.eBackupOperator AndAlso gyBackupOperator <> eyOperatorState.BackupOperator Then
            'ForwardToBackupOperator(yData, lIndex)
            ForwardToBackupOperator(SendBackupOpServerReady(oServerObject(lIndex)), -1)
        End If
    End Sub

    Private Sub HandleBackupOpServerReady(ByVal yData() As Byte, ByVal lIndex As Int32)

        Dim lPos As Int32 = 2
        Dim lBoxOpID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lProcID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lConnType As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim sIP As String = yData(lPos).ToString & "." & yData(lPos + 1).ToString & "." & yData(lPos + 2).ToString & "." & yData(lPos + 3).ToString
        lPos += 4

        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlBOPServerUB
            If moBOPServers(X) Is Nothing = False AndAlso moBOPServers(X).lBoxOperatorID = lBoxOpID AndAlso moBOPServers(X).lProcessID = lProcID AndAlso moBOPServers(X).lConnectionType = lConnType Then
                If moBOPServers(X).sIPAddress = sIP Then
                    lIdx = X
                    Exit For
                End If
            End If
        Next X
        If lIdx = -1 Then
            mlBOPServerUB += 1
            ReDim Preserve moBOPServers(mlBOPServerUB)
            lIdx = mlBOPServerUB
            moBOPServers(lIdx) = New ServerObject()
        End If

        With moBOPServers(lIdx)

            .lBoxOperatorID = lBoxOpID
            .lProcessID = lProcID
            .lConnectionType = CType(lConnType, ConnectionType)
            .sIPAddress = sIP

            .oRelatedServer = New ServerObject()
            With .oRelatedServer
                .lProcessID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lBoxOperatorID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .sIPAddress = yData(lPos).ToString & "." & yData(lPos + 1).ToString & "." & yData(lPos + 2).ToString & "." & yData(lPos + 3).ToString
                lPos += 4
            End With
            If .oRelatedServer.lProcessID = 0 AndAlso .oRelatedServer.lBoxOperatorID = -1 Then .oRelatedServer = Nothing

            Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            ReDim .uListeners(lCnt - 1)
            For X As Int32 = 0 To lCnt - 1
                sIP = yData(lPos).ToString & "." & yData(lPos + 1).ToString & "." & yData(lPos + 2).ToString & "." & yData(lPos + 3).ToString
                lPos += 4
                Dim lPort As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lListenConnType As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                .uListeners(X).sIPAddress = sIP
                .uListeners(X).lPortNumber = lPort
                .uListeners(X).lConnectionType = CType(lListenConnType, ConnectionType)
            Next X

            .lSpawnUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
            ReDim .lSpawnID(.lSpawnUB)
            ReDim .iSpawnTypeID(.lSpawnUB)
            For X As Int32 = 0 To .lSpawnUB
                .lSpawnID(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .iSpawnTypeID(X) = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Next X
        End With

    End Sub

    Private mlRegionReadyCnt As Int32 = 0
    Private Sub HandleServerReady(ByRef yData() As Byte, ByVal lIndex As Int32)
        With oServerObject(lIndex)
            Select Case .lConnectionType
                Case ConnectionType.ePrimaryServerApp
                    'Ok, need to spin up a pathfinding server
                    SpawnServer(oServerObject(lIndex), ConnectionType.ePathfindingServerApp, Nothing, Nothing, -1)
                Case ConnectionType.ePathfindingServerApp
                    'Ok, spin up region servers
                    Dim oPrimary As ServerObject = Nothing
                    'If oServerObject(lIndex).lSpawnRequestIdx <> -1 Then
                    '    oPrimary = uSpawnRequests(oServerObject(lIndex).lSpawnRequestIdx).oRelatedServer
                    'End If
                    oPrimary = oServerObject(lIndex).oRelatedServer


                    If oPrimary Is Nothing = False Then SpawnRegions(oPrimary) Else LogEvent(LogEventType.CriticalError, "Unable to spawn regions: could not determine pathfinding server's Primary!")
                    If gyBackupOperator <> eyOperatorState.BackupOperator Then moServers(lIndex).Disconnect()
                Case ConnectionType.eRegionServerApp
                    'are we done? if so, indicate server start to Primary and Regions
                    mlRegionReadyCnt += 1
                    If mlRegionReadyCnt = 9 OrElse gb_IS_TEST_SERVER = True Then
                        Dim yMsg(1) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eDomainServerReady).CopyTo(yMsg, 0)
                        'uSpawnRequests(.lSpawnRequestIdx).oRelatedServer.oSocket.SendData(yMsg)
                        oServerObject(lIndex).oRelatedServer.oSocket.SendData(yMsg)
                    End If
                    moServers(lIndex).Disconnect()
                    Return
                    'If .lSpawnRequestIdx <> -1 Then
                    '    Dim lPrimaryBoxOp As Int32 = -1
                    '    lPrimaryBoxOp = uSpawnRequests(uSpawnRequests(.lSpawnRequestIdx).oRelatedServer.lSpawnRequestIdx).lBoxOperatorID
                    '    For X As Int32 = 0 To lSpawnRequestUB
                    '        If uSpawnRequests(X).ySpawnRequestState = 1 AndAlso uSpawnRequests(X).lConnectionType = ConnectionType.eRegionServerApp Then
                    '            If uSpawnRequests(X).oRelatedServer Is Nothing = False AndAlso uSpawnRequests(X).oRelatedServer.lConnectionType = ConnectionType.ePrimaryServerApp Then
                    '                If uSpawnRequests(X).oRelatedServer.lBoxOperatorID = lPrimaryBoxOp Then
                    '                    'return now, another region is still required to be ready
                    '                    Return
                    '                End If
                    '            End If
                    '        End If
                    '    Next X

                    '    If gyBackupOperator <> eyOperatorState.BackupOperator Then
                    '        'if we are here, we are all ready... send to the Primary server to start
                    '        Dim yMsg(1) As Byte
                    '        System.BitConverter.GetBytes(GlobalMessageCode.eDomainServerReady).CopyTo(yMsg, 0)
                    '        uSpawnRequests(.lSpawnRequestIdx).oRelatedServer.oSocket.SendData(yMsg)

                    '        moServers(lIndex).Disconnect()
                    '    End If

                    'End If
                Case ConnectionType.eEmailServerApp
                    SpawnNeighborhoods()

                    'We do not disconnect the mail server because we are responsible for shutting it down

                Case ConnectionType.eArenaServer
                    'Simply disconnect the Arena server                    
                    moServers(lIndex).Disconnect()
                    Return
            End Select

            If .lConnectionType <> ConnectionType.eBackupOperator Then
                ForwardToBackupOperator(SendBackupOpServerReady(oServerObject(lIndex)), -1)
            End If
        End With
    End Sub
    Private Function SendBackupOpServerReady(ByVal oServer As ServerObject) As Byte()
        With oServer
            'ForwardToBackupOperator(yData, lIndex)
            Dim lListenCnt As Int32 = 0
            Dim lSpawnCnt As Int32 = 0
            If .uListeners Is Nothing = False Then lListenCnt = .uListeners.Length
            lSpawnCnt = .lSpawnUB + 1

            Dim yToBOP(37 + (lListenCnt * 12) + (lSpawnCnt * 6)) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eDomainServerReady).CopyTo(yToBOP, lPos) : lPos += 2
            System.BitConverter.GetBytes(.lBoxOperatorID).CopyTo(yToBOP, lPos) : lPos += 4
            System.BitConverter.GetBytes(.lProcessID).CopyTo(yToBOP, lPos) : lPos += 4
            System.BitConverter.GetBytes(.lConnectionType).CopyTo(yToBOP, lPos) : lPos += 4
            Dim sDots() As String = Split(.sIPAddress, ".")
            yToBOP(lPos) = CByte(sDots(0)) : lPos += 1
            yToBOP(lPos) = CByte(sDots(1)) : lPos += 1
            yToBOP(lPos) = CByte(sDots(2)) : lPos += 1
            yToBOP(lPos) = CByte(sDots(3)) : lPos += 1

            If .oRelatedServer Is Nothing = False Then
                System.BitConverter.GetBytes(.oRelatedServer.lProcessID).CopyTo(yToBOP, lPos) : lPos += 4
                System.BitConverter.GetBytes(.oRelatedServer.lBoxOperatorID).CopyTo(yToBOP, lPos) : lPos += 4
                sDots = Split(.oRelatedServer.sIPAddress, ".")

                yToBOP(lPos) = CByte(sDots(0)) : lPos += 1
                yToBOP(lPos) = CByte(sDots(1)) : lPos += 1
                yToBOP(lPos) = CByte(sDots(2)) : lPos += 1
                yToBOP(lPos) = CByte(sDots(3)) : lPos += 1
            Else
                System.BitConverter.GetBytes(0I).CopyTo(yToBOP, lPos) : lPos += 4
                System.BitConverter.GetBytes(-1I).CopyTo(yToBOP, lPos) : lPos += 4
                System.BitConverter.GetBytes(0I).CopyTo(yToBOP, lPos) : lPos += 4
            End If

            System.BitConverter.GetBytes(lListenCnt).CopyTo(yToBOP, lPos) : lPos += 4

            For X As Int32 = 0 To .uListeners.GetUpperBound(0)
                sDots = Split(.uListeners(X).sIPAddress, ".")
                yToBOP(lPos) = CByte(sDots(0)) : lPos += 1
                yToBOP(lPos) = CByte(sDots(1)) : lPos += 1
                yToBOP(lPos) = CByte(sDots(2)) : lPos += 1
                yToBOP(lPos) = CByte(sDots(3)) : lPos += 1

                System.BitConverter.GetBytes(.uListeners(X).lPortNumber).CopyTo(yToBOP, lPos) : lPos += 4
                System.BitConverter.GetBytes(.uListeners(X).lConnectionType).CopyTo(yToBOP, lPos) : lPos += 4
            Next X

            System.BitConverter.GetBytes(lSpawnCnt).CopyTo(yToBOP, lPos) : lPos += 4

            For X As Int32 = 0 To .lSpawnUB
                System.BitConverter.GetBytes(.lSpawnID(X)).CopyTo(yToBOP, lPos) : lPos += 4
                System.BitConverter.GetBytes(.iSpawnTypeID(X)).CopyTo(yToBOP, lPos) : lPos += 2
            Next X

            'ForwardToBackupOperator(yToBOP, -1)

            Return yToBOP
        End With

    End Function
    'Private Sub HandleServerReady(ByRef yData() As Byte, ByVal lIndex As Int32)
    '    With oServerObject(lIndex)
    '        Select Case .lConnectionType
    '            Case ConnectionType.ePrimaryServerApp
    '                'Ok, need to spin up a pathfinding server
    '                SpawnServer(oServerObject(lIndex), ConnectionType.ePathfindingServerApp, Nothing, Nothing, -1)
    '            Case ConnectionType.ePathfindingServerApp
    '                'Ok, spin up region servers
    '                Dim oPrimary As ServerObject = Nothing
    '                If oServerObject(lIndex).lSpawnRequestIdx <> -1 Then
    '                    oPrimary = uSpawnRequests(oServerObject(lIndex).lSpawnRequestIdx).oRelatedServer
    '                End If

    '                If oPrimary Is Nothing = False Then SpawnRegions(oPrimary) Else LogEvent(LogEventType.CriticalError, "Unable to spawn regions: could not determine pathfinding server's Primary!")
    '                If gyBackupOperator <> eyOperatorState.BackupOperator Then moServers(lIndex).Disconnect()
    '            Case ConnectionType.eRegionServerApp
    '                'are we done? if so, indicate server start to Primary and Regions
    '                If .lSpawnRequestIdx <> -1 Then
    '                    Dim lPrimaryBoxOp As Int32 = -1
    '                    lPrimaryBoxOp = uSpawnRequests(uSpawnRequests(.lSpawnRequestIdx).oRelatedServer.lSpawnRequestIdx).lBoxOperatorID
    '                    For X As Int32 = 0 To lSpawnRequestUB
    '                        If uSpawnRequests(X).ySpawnRequestState = 1 AndAlso uSpawnRequests(X).lConnectionType = ConnectionType.eRegionServerApp Then
    '                            If uSpawnRequests(X).oRelatedServer Is Nothing = False AndAlso uSpawnRequests(X).oRelatedServer.lConnectionType = ConnectionType.ePrimaryServerApp Then
    '                                If uSpawnRequests(X).oRelatedServer.lBoxOperatorID = lPrimaryBoxOp Then
    '                                    'return now, another region is still required to be ready
    '                                    Return
    '                                End If
    '                            End If
    '                        End If
    '                    Next X

    '                    If gyBackupOperator <> eyOperatorState.BackupOperator Then
    '                        'if we are here, we are all ready... send to the Primary server to start
    '                        Dim yMsg(1) As Byte
    '                        System.BitConverter.GetBytes(GlobalMessageCode.eDomainServerReady).CopyTo(yMsg, 0)
    '                        uSpawnRequests(.lSpawnRequestIdx).oRelatedServer.oSocket.SendData(yMsg)

    '                        moServers(lIndex).Disconnect()
    '                    End If

    '                End If
    '            Case ConnectionType.eEmailServerApp
    '                SpawnNeighborhoods()

    '                'We do not disconnect the mail server because we are responsible for shutting it down
    '        End Select

    '        If .lConnectionType <> ConnectionType.eBackupOperator Then
    '            ForwardToBackupOperator(yData, lIndex)
    '        End If
    '    End With
    'End Sub
    Private Sub FinalizePrimaryIdentification(ByRef oServer As ServerObject, ByVal lServerIndex As Int32)
        'For X As Int32 = 0 To lSpawnRequestUB
        '    If uSpawnRequests(X).lBoxOperatorID = oServer.lBoxOperatorID AndAlso uSpawnRequests(X).lConnectionType = oServer.lConnectionType Then
        '        'Ok, same box, same process
        '        If uSpawnRequests(X).ySpawnRequestState = 1 Then
        '            'Ok, must be it
        '            uSpawnRequests(X).ySpawnRequestState = 2
        '            oServer.lSpawnRequestIdx = X

        '            'Now, process anything remaining to do
        '            If uSpawnRequests(X).lSpawnID Is Nothing = False Then
        '                Dim lCnt As Int32 = uSpawnRequests(X).lSpawnID.GetUpperBound(0) + 1
        '                Dim yMsg(9 + (lCnt * 6)) As Byte
        '                Dim lPos As Int32 = 0
        '                System.BitConverter.GetBytes(GlobalMessageCode.eRegisterDomain).CopyTo(yMsg, lPos) : lPos += 2
        '                System.BitConverter.GetBytes(lServerIndex).CopyTo(yMsg, lPos) : lPos += 4
        '                System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4
        '                For Y As Int32 = 0 To uSpawnRequests(X).lSpawnID.GetUpperBound(0)
        '                    System.BitConverter.GetBytes(uSpawnRequests(X).lSpawnID(Y)).CopyTo(yMsg, lPos) : lPos += 4
        '                    System.BitConverter.GetBytes(uSpawnRequests(X).iSpawnTypeID(Y)).CopyTo(yMsg, lPos) : lPos += 2

        '                    If uSpawnRequests(X).iSpawnTypeID(Y) = ObjectType.eSolarSystem Then
        '                        Dim oSystem As SolarSystem = GetEpicaSystem(uSpawnRequests(X).lSpawnID(Y))
        '                        If oSystem Is Nothing = False Then oSystem.oPrimaryServer = oServer
        '                    End If
        '                Next Y
        '                If gyBackupOperator <> eyOperatorState.BackupOperator Then oServer.oSocket.SendData(yMsg)
        '            End If

        '            'Next, send the email server related to it
        '            If gyBackupOperator <> eyOperatorState.BackupOperator AndAlso uSpawnRequests(X).oRelatedServer Is Nothing = False Then
        '                'this is my email server
        '                If uSpawnRequests(X).oRelatedServer.lConnectionType = ConnectionType.eEmailServerApp Then
        '                    With uSpawnRequests(X).oRelatedServer
        '                        If .uListeners Is Nothing = False Then
        '                            Dim bFound As Boolean = False
        '                            For Y As Int32 = 0 To .uListeners.GetUpperBound(0)
        '                                If .uListeners(Y).lConnectionType = ConnectionType.ePrimaryServerApp Then
        '                                    'Ok, found it... create our message
        '                                    Dim yMsg(25) As Byte
        '                                    Dim lPos As Int32 = 0
        '                                    System.BitConverter.GetBytes(GlobalMessageCode.eEmailSettings).CopyTo(yMsg, lPos) : lPos += 2
        '                                    StringToBytes(.uListeners(Y).sIPAddress).CopyTo(yMsg, lPos) : lPos += 20
        '                                    System.BitConverter.GetBytes(.uListeners(Y).lPortNumber).CopyTo(yMsg, lPos) : lPos += 4
        '                                    oServer.oSocket.SendData(yMsg)
        '                                    bFound = True
        '                                    Exit For
        '                                End If
        '                            Next Y
        '                            If bFound = False Then
        '                                LogEvent(LogEventType.CriticalError, "Unable to determine EmailServer's Listener for Primary Server!")
        '                            End If
        '                        End If
        '                    End With
        '                Else
        '                    LogEvent(LogEventType.CriticalError, "Unexpected Server Type for related Server: " & uSpawnRequests(X).oRelatedServer.lConnectionType.ToString & ", expected EmailServer")
        '                End If
        '            End If

        '            'finally exit
        '            Return
        '        End If
        '    End If
        'Next X

        'LogEvent(LogEventType.CriticalError, "FinalizePrimaryIdentification on a server object could not determine SpawnRequest")

        'Now, process anything remaining to do
        If oServer.lSpawnID Is Nothing = False Then
            Dim lCnt As Int32 = oServer.lSpawnID.GetUpperBound(0) + 1
            Dim yMsg(9 + (lCnt * 6)) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eRegisterDomain).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lServerIndex).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4
            For Y As Int32 = 0 To oServer.lSpawnID.GetUpperBound(0)
                System.BitConverter.GetBytes(oServer.lSpawnID(Y)).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(oServer.iSpawnTypeID(Y)).CopyTo(yMsg, lPos) : lPos += 2

                If oServer.iSpawnTypeID(Y) = ObjectType.eSolarSystem Then
                    Dim oSystem As SolarSystem = GetEpicaSystem(oServer.lSpawnID(Y))
                    If oSystem Is Nothing = False Then oSystem.oPrimaryServer = oServer
                End If
            Next Y
            If gyBackupOperator <> eyOperatorState.BackupOperator Then oServer.oSocket.SendData(yMsg)
        End If

        'Next, send the email server related to it
        If gyBackupOperator <> eyOperatorState.BackupOperator AndAlso oServer.oRelatedServer Is Nothing = False Then
            'this is my email server
            If oServer.oRelatedServer.lConnectionType = ConnectionType.eEmailServerApp Then
                With oServer.oRelatedServer
                    If .uListeners Is Nothing = False Then
                        Dim bFound As Boolean = False
                        For Y As Int32 = 0 To .uListeners.GetUpperBound(0)
                            If .uListeners(Y).lConnectionType = ConnectionType.ePrimaryServerApp Then
                                'Ok, found it... create our message
                                Dim yMsg(25) As Byte
                                Dim lPos As Int32 = 0
                                System.BitConverter.GetBytes(GlobalMessageCode.eEmailSettings).CopyTo(yMsg, lPos) : lPos += 2
                                StringToBytes(.uListeners(Y).sIPAddress).CopyTo(yMsg, lPos) : lPos += 20
                                System.BitConverter.GetBytes(.uListeners(Y).lPortNumber).CopyTo(yMsg, lPos) : lPos += 4
                                oServer.oSocket.SendData(yMsg)
                                bFound = True
                                Exit For
                            End If
                        Next Y
                        If bFound = False Then
                            LogEvent(LogEventType.CriticalError, "Unable to determine EmailServer's Listener for Primary Server!")
                        End If
                    End If
                End With
            Else
                LogEvent(LogEventType.CriticalError, "Unexpected Server Type for related Server: " & oServer.oRelatedServer.lConnectionType.ToString & ", expected EmailServer")
            End If
        End If
    End Sub
    Private Sub FinalizeRegionIdentification(ByRef oServer As ServerObject)
        'For X As Int32 = 0 To lSpawnRequestUB
        '    If uSpawnRequests(X).lBoxOperatorID = oServer.lBoxOperatorID AndAlso uSpawnRequests(X).lConnectionType = oServer.lConnectionType Then
        '        'Ok, same box, same process
        '        If uSpawnRequests(X).ySpawnRequestState = 1 Then
        '            'Ok, must be it
        '            uSpawnRequests(X).ySpawnRequestState = 2
        '            oServer.lSpawnRequestIdx = X

        '            'Now, process anything remaining to do
        '            If gyBackupOperator <> eyOperatorState.BackupOperator Then
        '                If uSpawnRequests(X).lSpawnID Is Nothing = False Then
        '                    Dim lCnt As Int32 = uSpawnRequests(X).lSpawnID.GetUpperBound(0) + 1
        '                    Dim yMsg(29 + (lCnt * 6)) As Byte
        '                    Dim lPos As Int32 = 0
        '                    System.BitConverter.GetBytes(GlobalMessageCode.eRegisterDomain).CopyTo(yMsg, lPos) : lPos += 2

        '                    If uSpawnRequests(X).oRelatedServer Is Nothing = False Then
        '                        If uSpawnRequests(X).oRelatedServer.lConnectionType = ConnectionType.ePrimaryServerApp Then
        '                            With uSpawnRequests(X).oRelatedServer
        '                                If .uListeners Is Nothing = False Then
        '                                    Dim bFound As Boolean = False
        '                                    For Y As Int32 = 0 To .uListeners.GetUpperBound(0)
        '                                        If .uListeners(Y).lConnectionType = ConnectionType.eRegionServerApp Then
        '                                            bFound = True
        '                                            StringToBytes(.uListeners(Y).sIPAddress).CopyTo(yMsg, lPos) : lPos += 20
        '                                            System.BitConverter.GetBytes(.uListeners(Y).lPortNumber).CopyTo(yMsg, lPos) : lPos += 4
        '                                            Exit For
        '                                        End If
        '                                    Next Y
        '                                    If bFound = False Then
        '                                        LogEvent(LogEventType.CriticalError, "Unable to determine listener credentials for primary server to connect a region.")
        '                                        Return
        '                                    End If
        '                                End If
        '                            End With
        '                        Else
        '                            LogEvent(LogEventType.CriticalError, "Unexpected RelatedServer for Region Server: " & uSpawnRequests(X).oRelatedServer.lConnectionType.ToString)
        '                            Return
        '                        End If
        '                    Else
        '                        LogEvent(LogEventType.CriticalError, "Unexpected RelatedServer for Region Server: Nothing")
        '                        Return
        '                    End If

        '                    System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4
        '                    For Y As Int32 = 0 To uSpawnRequests(X).lSpawnID.GetUpperBound(0)
        '                        System.BitConverter.GetBytes(uSpawnRequests(X).lSpawnID(Y)).CopyTo(yMsg, lPos) : lPos += 4
        '                        System.BitConverter.GetBytes(uSpawnRequests(X).iSpawnTypeID(Y)).CopyTo(yMsg, lPos) : lPos += 2
        '                    Next Y
        '                    oServer.oSocket.SendData(yMsg)
        '                Else : LogEvent(LogEventType.CriticalError, "Unexpected SpawnList for Region Server: Empty.")
        '                End If
        '            End If

        '            'finally exit
        '            Return
        '        End If
        '    End If
        'Next X

        'LogEvent(LogEventType.CriticalError, "FinalizeRegionIdentification on a server object could not determine SpawnRequest")

        'Now, process anything remaining to do
        If gyBackupOperator <> eyOperatorState.BackupOperator Then
            If oServer.lSpawnID Is Nothing = False Then
                Dim lCnt As Int32 = oServer.lSpawnID.GetUpperBound(0) + 1
                Dim yMsg(29 + (lCnt * 6)) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eRegisterDomain).CopyTo(yMsg, lPos) : lPos += 2

                If oServer.oRelatedServer Is Nothing = False Then
                    If oServer.oRelatedServer.lConnectionType = ConnectionType.ePrimaryServerApp Then
                        With oServer.oRelatedServer
                            If .uListeners Is Nothing = False Then
                                Dim bFound As Boolean = False
                                For Y As Int32 = 0 To .uListeners.GetUpperBound(0)
                                    If .uListeners(Y).lConnectionType = ConnectionType.eRegionServerApp Then
                                        bFound = True
                                        StringToBytes(.uListeners(Y).sIPAddress).CopyTo(yMsg, lPos) : lPos += 20
                                        System.BitConverter.GetBytes(.uListeners(Y).lPortNumber).CopyTo(yMsg, lPos) : lPos += 4
                                        Exit For
                                    End If
                                Next Y
                                If bFound = False Then
                                    LogEvent(LogEventType.CriticalError, "Unable to determine listener credentials for primary server to connect a region.")
                                    Return
                                End If
                            End If
                        End With
                    Else
                        LogEvent(LogEventType.CriticalError, "Unexpected RelatedServer for Region Server: " & oServer.oRelatedServer.lConnectionType.ToString)
                        Return
                    End If
                Else
                    LogEvent(LogEventType.CriticalError, "Unexpected RelatedServer for Region Server: Nothing")
                    Return
                End If

                System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4
                For Y As Int32 = 0 To oServer.lSpawnID.GetUpperBound(0)
                    System.BitConverter.GetBytes(oServer.lSpawnID(Y)).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(oServer.iSpawnTypeID(Y)).CopyTo(yMsg, lPos) : lPos += 2
                Next Y
                oServer.oSocket.SendData(yMsg)
            Else : LogEvent(LogEventType.CriticalError, "Unexpected SpawnList for Region Server: Empty.")
            End If
        End If
    End Sub
    Private Sub FinalizePathfindingIdentification(ByRef oServer As ServerObject)
        'For X As Int32 = 0 To lSpawnRequestUB
        '    If uSpawnRequests(X).lBoxOperatorID = oServer.lBoxOperatorID AndAlso uSpawnRequests(X).lConnectionType = oServer.lConnectionType Then
        '        'Ok, same box, same process
        '        If uSpawnRequests(X).ySpawnRequestState = 1 Then
        '            'Ok, must be it
        '            uSpawnRequests(X).ySpawnRequestState = 2
        '            oServer.lSpawnRequestIdx = X

        '            'Now, process anything remaining to do
        '            If gyBackupOperator <> eyOperatorState.BackupOperator Then
        '                If uSpawnRequests(X).oRelatedServer Is Nothing = False Then
        '                    If uSpawnRequests(X).oRelatedServer.lConnectionType = ConnectionType.ePrimaryServerApp Then
        '                        With uSpawnRequests(X).oRelatedServer
        '                            If .uListeners Is Nothing = False Then
        '                                Dim bFound As Boolean = False
        '                                For Y As Int32 = 0 To .uListeners.GetUpperBound(0)
        '                                    If .uListeners(Y).lConnectionType = ConnectionType.ePathfindingServerApp Then
        '                                        bFound = True
        '                                        Dim yMsg(25) As Byte
        '                                        Dim lPos As Int32 = 0
        '                                        System.BitConverter.GetBytes(GlobalMessageCode.ePathfindingConnectionInfo).CopyTo(yMsg, lPos) : lPos += 2
        '                                        StringToBytes(.uListeners(Y).sIPAddress).CopyTo(yMsg, lPos) : lPos += 20
        '                                        System.BitConverter.GetBytes(.uListeners(Y).lPortNumber).CopyTo(yMsg, lPos) : lPos += 4
        '                                        oServer.oSocket.SendData(yMsg)
        '                                        Exit For
        '                                    End If
        '                                Next Y
        '                                If bFound = False Then
        '                                    LogEvent(LogEventType.CriticalError, "Unable to determine listener credentials for primary server to connect a pathfinding server.")
        '                                    Return
        '                                End If
        '                            End If
        '                        End With
        '                    Else
        '                        LogEvent(LogEventType.CriticalError, "Unexpected RelatedServer for pathfinding Server: " & uSpawnRequests(X).oRelatedServer.lConnectionType.ToString)
        '                        Return
        '                    End If
        '                Else
        '                    LogEvent(LogEventType.CriticalError, "Unexpected RelatedServer for pathfinding Server: Nothing")
        '                    Return
        '                End If
        '            End If

        '            'finally exit
        '            Return
        '        End If
        '    End If
        'Next X

        'LogEvent(LogEventType.CriticalError, "FinalizePathfindingIdentification on a server object could not determine SpawnRequest")

        'Now, process anything remaining to do
        If gyBackupOperator <> eyOperatorState.BackupOperator Then
            If oServer.oRelatedServer Is Nothing = False Then
                If oServer.oRelatedServer.lConnectionType = ConnectionType.ePrimaryServerApp Then
                    With oServer.oRelatedServer
                        If .uListeners Is Nothing = False Then
                            Dim bFound As Boolean = False
                            For Y As Int32 = 0 To .uListeners.GetUpperBound(0)
                                If .uListeners(Y).lConnectionType = ConnectionType.ePathfindingServerApp Then
                                    bFound = True
                                    Dim yMsg(25) As Byte
                                    Dim lPos As Int32 = 0
                                    System.BitConverter.GetBytes(GlobalMessageCode.ePathfindingConnectionInfo).CopyTo(yMsg, lPos) : lPos += 2
                                    StringToBytes(.uListeners(Y).sIPAddress).CopyTo(yMsg, lPos) : lPos += 20
                                    System.BitConverter.GetBytes(.uListeners(Y).lPortNumber).CopyTo(yMsg, lPos) : lPos += 4
                                    oServer.oSocket.SendData(yMsg)
                                    Exit For
                                End If
                            Next Y
                            If bFound = False Then
                                LogEvent(LogEventType.CriticalError, "Unable to determine listener credentials for primary server to connect a pathfinding server.")
                                Return
                            End If
                        End If
                    End With
                Else
                    LogEvent(LogEventType.CriticalError, "Unexpected RelatedServer for pathfinding Server: " & oServer.oRelatedServer.lConnectionType.ToString)
                    Return
                End If
            Else
                LogEvent(LogEventType.CriticalError, "Unexpected RelatedServer for pathfinding Server: Nothing")
                Return
            End If
        End If
    End Sub
    Private Sub HandleRemovePlayerRel(ByRef yData() As Byte)
        For X As Int32 = 0 To mlServerUB
            If oServerObject(X) Is Nothing = False Then
                If oServerObject(X).lConnectionType = ConnectionType.ePrimaryServerApp Then
                    Try
                        oServerObject(X).oSocket.SendData(yData)
                    Catch ex As Exception
                        LogEvent(LogEventType.CriticalError, "HandleRemovePlayerRel: " & ex.Message)
                    End Try
                End If
            End If
        Next X
    End Sub
    Private Sub HandleEnvironmentDomainMsg(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lType As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lProcCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim bulTotalPhysMem As UInt64 = System.BitConverter.ToUInt64(yData, lPos) : lPos += 8
        Dim bulAvailPhysMem As UInt64 = System.BitConverter.ToUInt64(yData, lPos) : lPos += 8
        Dim bulTotalVirtMem As UInt64 = System.BitConverter.ToUInt64(yData, lPos) : lPos += 8
        Dim bulAvailVirtMem As UInt64 = System.BitConverter.ToUInt64(yData, lPos) : lPos += 8

        If lType <> ConnectionType.eBoxOperator Then Return

        'Now, find our server object
        If lIndex > -1 AndAlso lIndex <= mlServerUB Then
            If oServerObject(lIndex) Is Nothing = False Then
                With CType(oServerObject(lIndex), BoxOperator)
                    If .lConnectionType = lType AndAlso lID = .lBoxOperatorID Then
                        .lProcessorCount = lProcCnt
                        .bulAvailablePhysicalMemory = bulAvailPhysMem
                        .bulAvailableVirtualMemory = bulAvailVirtMem
                        .bulTotalPhysicalMemory = bulTotalPhysMem
                        .bulTotalVirtualMemory = bulTotalVirtMem
                    End If
                End With
            End If
        End If

        'ForwardToBackupOperator(yData, lIndex)
    End Sub

    Private Sub HandlePrimarySpawnLocUpdate(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lSpawnsAvail As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lRespawnsAvail As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        oServerObject(lIndex).lSpawnLocsAvail = lSpawnsAvail
        oServerObject(lIndex).lRespawnLocsAvail = lRespawnsAvail

        Dim lTotalSpawns As Int32 = 0
        Dim lTotalRespawns As Int32 = 0
        For X As Int32 = 0 To mlServerUB
            If oServerObject(lIndex) Is Nothing = False AndAlso oServerObject(lIndex).lConnectionType = ConnectionType.ePrimaryServerApp Then
                lTotalSpawns += oServerObject(lIndex).lSpawnLocsAvail
                lTotalRespawns += oServerObject(lIndex).lRespawnLocsAvail
            End If
        Next X
        If lTotalSpawns < 20 OrElse lTotalRespawns < 5 Then
            GeoSpawner.IncrementGenerateCounter()
        End If
    End Sub

    Private Sub HandleClientVersion(ByVal yData() As Byte, ByVal lSocketIndex As Int32)
        Try
            Dim lTheirVersion As Int32 = System.BitConverter.ToInt32(yData, 2)

            Dim yMsg(5) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eClientVersion).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(gl_CLIENT_VERSION).CopyTo(yMsg, 2)

            moClients(lSocketIndex).SendData(yMsg)

            If lTheirVersion <> gl_CLIENT_VERSION Then
                moClients(lSocketIndex).Disconnect()
                moClients(lSocketIndex) = Nothing
            End If
        Catch
            moClients(lSocketIndex).Disconnect()
            moClients(lSocketIndex) = Nothing
        End Try
    End Sub
    Private Function DecBytes(ByVal yBytes() As Byte) As Byte()
        Const ml_ENCRYPT_SEED As Int32 = 777

        'Now, we do the exact opposite...
        Dim lLen As Int32 = yBytes.GetUpperBound(0)
        Dim lKey As Int32
        'Dim lOffset As Int32
        Dim X As Int32
        Dim yFinal(lLen - 1) As Byte
        Dim lChrCode As Int32
        Dim lMod As Int32

        'Get our key value...
        lKey = yBytes(0)

        'set up our seed
        Rnd(-1)
        Call Randomize(ml_ENCRYPT_SEED + lKey)

        For X = 1 To lLen
            'Now, find out what we got here...
            lChrCode = yBytes(X)
            'now, subtract our value... 1 to 5
            lMod = CInt(Int(Rnd() * 5) + 1)
            lChrCode = lChrCode - lMod
            If lChrCode < 0 Then lChrCode = 256 + lChrCode
            yFinal(X - 1) = CByte(lChrCode)
        Next X
        DecBytes = yFinal
    End Function

    Private Function GetPlayerPlayedTime(ByVal lPlayerID As Int32) As Int32
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim oData As OleDb.OleDbDataReader = Nothing
        Dim sSQL As String = ""
        Dim lResult As Int32 = 0

        Try
            sSQL = "SELECT TotalPlayTime FROM tblPlayer WHERE PlayerID = " & lPlayerID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oData = oComm.ExecuteReader()
            If oData.Read = True Then
                lResult = CInt(oData(0))
            End If
            oData.Close()
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "GetPlayerPlayedTime: " & ex.Message)
        Finally
            If oData Is Nothing = False Then oData.Close()
            oData = Nothing
            If oComm Is Nothing = False Then oComm.Dispose()
            oComm = Nothing
        End Try

        Return lResult
    End Function
    Private Function SaveLoginTransaction(ByVal lPlayerID As Int32, ByVal sPlayerName As String, ByVal sIP As String, ByVal lStatus As Int32, ByVal lPrevStatus As Int32) As Boolean

        If gb_IS_TEST_SERVER = True Then Return True

        Dim oComm As Odbc.OdbcCommand = Nothing
        Dim sSQL As String = ""

        Dim bResult As Boolean = False
        Dim lPlayTime As Int32 = GetPlayerPlayedTime(lPlayerID)

        'Now, is the player's status a 1?
        If lStatus = AccountStatusType.eActiveAccount AndAlso lPrevStatus <> AccountStatusType.eActiveAccount Then
            Try
                'Ok, get the recent status from usertrans
                sSQL = "SELECT CurrentStatus FROM User_Play_HX WHERE PlayerID = " & lPlayerID & " ORDER BY LoginTime DESC LIMIT 1"
                oComm = New Odbc.OdbcCommand(sSQL, goSuiteCN)
                Dim oData As Odbc.OdbcDataReader = oComm.ExecuteReader()

                'If first time logging in since server down, lPrevStatus will be -1
                '  thus, if there are no records to read, or the only record is null then prevstatus will = -1
                '  otherwise, prevstatus will be loaded...

                If oData.Read = True Then
                    If oData("CurrentStatus") Is DBNull.Value = False Then
                        lPrevStatus = CInt(oData("CurrentStatus"))
                    End If
                End If
                oData.Close()
                oComm.Dispose()
                oComm = Nothing

                'So, we ask here, if the status is -1 (never logged in before) or status is 99 (trial) then we need to set their 30D because they have logged in as active
                If lPrevStatus = -1 OrElse lPrevStatus = 99 Then
                    'Ok, need to update the subscription to indicate the end date of the 30 Day part
                    LogEvent(LogEventType.Informational, "Setting 30 day end period on " & sPlayerName & ".")

                    sSQL = "UPDATE fc_subscriptions SET StartDate = '" & Now.ToString("yyyy-MM-dd") & "', EndDate = '" & Now.AddDays(30).ToString("yyyy-MM-dd") & "' WHERE UserID = "
                    sSQL &= "(SELECT ID FROM fc_module_users WHERE UserName = '" & sPlayerName & "') AND ProdID = 64"

                    oComm = New Odbc.OdbcCommand(sSQL, goSuiteCN)
                    If oComm.ExecuteNonQuery() = 0 Then
                        'Throw New Exception("No Records updated!")
                        LogEvent(LogEventType.CriticalError, "SaveLoginTransaction New Player 30Day Period Update: No Records update!" & vbCrLf & sSQL)
                    End If
                    oComm.Dispose()
                    oComm = Nothing

                    sSQL = "UPDATE fc_module_users SET GameStartDate = '" & Now.ToString("yyyy-MM-dd") & "', SubExpirationDate = '" & Now.AddDays(30).ToString("yyyy-MM-dd") & _
                      "' WHERE UserName = '" & sPlayerName & "'"
                    oComm = New Odbc.OdbcCommand(sSQL, goSuiteCN)
                    If oComm.ExecuteNonQuery() = 0 Then
                        Throw New Exception("No Records Updated!")
                    End If
                    oComm.Dispose()
                    oComm = Nothing

                End If

                bResult = True
            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "SaveLoginTransaction New Player 30Day Period Update: " & ex.Message & vbCrLf & sSQL)
            Finally
                If oComm Is Nothing = False Then oComm.Dispose()
                oComm = Nothing
            End Try
        Else
            bResult = True
        End If

        Try
            sSQL = "INSERT INTO User_Play_HX (PlayerID, UserName, LoginTime, IPAddress, CurrentStatus, CurrentPlayed) VALUES (" & lPlayerID
            sSQL &= ", '" & MakeDBStr(sPlayerName) & "', '" & Date.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime.ToString("yyyy-MM-dd HH:mm:ss") & "', '" & _
            sIP & "', " & lStatus & ", " & lPlayTime & ")"

            oComm = New Odbc.OdbcCommand(sSQL, goSuiteCN)
            oComm.ExecuteNonQuery()
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "SaveLoginTransaction: " & ex.Message & vbCrLf & sSQL)
        Finally
            If oComm Is Nothing = False Then oComm.Dispose()
            oComm = Nothing
        End Try

        Return bResult
    End Function

    Private Sub HandleLoginRequest(ByVal yData() As Byte, ByVal lSocketIndex As Int32)
        Dim X As Int32
        Dim Y As Int32
        Dim yUserName(19) As Byte
        Dim yPassword(19) As Byte
        Dim lID As Int32
        Dim ySame As Byte           '0 = no match, 1 = username match but password mismatch, 2 = perfect match
        Dim bAlias As Boolean = False
        Dim yAliasUserName(19) As Byte
        Dim yAliasPassword(19) As Byte
        Dim lAliasID As Int32
        Dim lRights As Int32 = 0
        Dim oPlayer As Player = Nothing

        Try
            Dim oEnc As New StrEncDec

            Dim lPos As Int32 = 2   'for msgcode
            Array.Copy(yData, lPos, yUserName, 0, 20) : lPos += 20
            yUserName = oEnc.Decrypt(yUserName)

            Array.Copy(yData, lPos, yPassword, 0, 20) : lPos += 21
            yPassword = oEnc.Decrypt(yPassword)
            
            If yData(lPos) <> 0 Then bAlias = True
            lPos += 1
            Array.Copy(yData, lPos, yAliasUserName, 0, 20) : lPos += 20
            yAliasUserName = oEnc.Decrypt(yAliasUserName)
            Array.Copy(yData, lPos, yAliasPassword, 0, 20) : lPos += 21
            yAliasPassword = oEnc.Decrypt(yAliasPassword)

            For X = 0 To glPlayerUB
                If glPlayerIdx(X) <> -1 Then
                    With goPlayer(X)
                        ySame = 2

                        For Y = 0 To 19
                            If yUserName(Y) <> .PlayerUserName(Y) Then
                                ySame = 0
                                Exit For
                            End If
                        Next Y

                        If ySame = 2 Then
                            'Ok, check the password
                            For Y = 0 To 19
                                If yPassword(Y) <> .PlayerPassword(Y) Then
                                    ySame = 1
                                    Exit For
                                End If
                            Next Y

                            If ySame = 2 Then

                                LogEvent(LogEventType.Informational, "Login Account: " & BytesToString(.PlayerName) & " at " & moClients(lSocketIndex).GetRemoteIP())

                                If SaveLoginTransaction(.ObjectID, .sPlayerName, moClients(lSocketIndex).GetRemoteIP(), .AccountStatus, .lPreviousStatus) = True Then .lPreviousStatus = .AccountStatus

                                If .lPlayerIcon = 0 Then
                                    lID = LoginResponseCodes.eAccountSetup
                                ElseIf .swInDyingProcess Is Nothing = False AndAlso .swInDyingProcess.ElapsedMilliseconds < 30000 Then
                                    lID = LoginResponseCodes.ePlayerIsDying
                                Else
                                    .swInDyingProcess = Nothing
                                    .lPrimaryReturned = Nothing

                                    'Ok, found the player... check the account status
                                    Select Case .AccountStatus
                                        Case AccountStatusType.eActiveAccount, AccountStatusType.eOpenBetaAccount, AccountStatusType.eTrialAccount, AccountStatusType.eMondelisActive

                                            If .AccountStatus = AccountStatusType.eOpenBetaAccount AndAlso gb_IN_OPEN_BETA = False Then
                                                lID = LoginResponseCodes.eAccountInactive
                                            Else
                                                lID = glPlayerIdx(X)
                                                lAliasID = glPlayerIdx(X)

                                                oPlayer = goPlayer(X)

                                                If bAlias = True Then
                                                    'Ok, now, do a lookup for the aliased player
                                                    Dim bFound As Boolean = False

                                                    'lSenderID = 7 OrElse lSenderID = 2076 OrElse lSenderID = 1780 Then
                                                    If .ObjectID = 1 OrElse .ObjectID = 2 OrElse .ObjectID = 6 OrElse .ObjectID = 221 OrElse .ObjectID = 131 OrElse .ObjectID = 2067 OrElse .ObjectID = 7 OrElse .ObjectID = 2076 OrElse .ObjectID = 1780 Then
                                                        For Y = 0 To glPlayerUB
                                                            If glPlayerIdx(Y) > -1 Then
                                                                Dim oTmp As Player = goPlayer(Y)
                                                                If oTmp Is Nothing = False Then
                                                                    bFound = True
                                                                    For Z As Int32 = 0 To 19
                                                                        If oTmp.PlayerUserName(Z) <> yAliasUserName(Z) Then
                                                                            bFound = False
                                                                            Exit For
                                                                        End If
                                                                    Next Z

                                                                    If bFound = True Then
                                                                        lAliasID = oTmp.ObjectID
                                                                        lRights = AliasingRights.eAddProduction Or AliasingRights.eAddResearch Or AliasingRights.eAlterAgents Or AliasingRights.eAlterAutoLaunchPower Or AliasingRights.eAlterColonyStats Or AliasingRights.eAlterDiplomacy Or AliasingRights.eAlterEmail Or AliasingRights.eAlterTrades Or AliasingRights.eCancelProduction Or AliasingRights.eCancelResearch Or AliasingRights.eChangeBehavior Or AliasingRights.eChangeEnvironment Or AliasingRights.eCreateBattleGroups Or AliasingRights.eCreateDesigns Or AliasingRights.eDockUndockUnits Or AliasingRights.eModifyBattleGroups Or AliasingRights.eMoveUnits Or AliasingRights.eTransferCargo Or AliasingRights.eViewAgents Or AliasingRights.eViewBattleGroups Or AliasingRights.eViewBudget Or AliasingRights.eViewColonyStats Or AliasingRights.eViewDiplomacy Or AliasingRights.eViewEmail Or AliasingRights.eViewMining Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns Or AliasingRights.eViewTrades Or AliasingRights.eViewTreasury Or AliasingRights.eViewUnitsAndFacilities
                                                                        .lAliasingPlayerID = oTmp.ObjectID
                                                                        oPlayer = oTmp
                                                                        Exit For
                                                                    End If
                                                                End If
                                                            End If
                                                        Next Y
                                                    Else
                                                        For Y = 0 To .lAllowanceUB
                                                            If .lAllowanceIdx(Y) <> -1 Then
                                                                bFound = True
                                                                For Z As Int32 = 0 To 19
                                                                    If .uAllowanceLogin(Y).yUserName(Z) <> yAliasUserName(Z) OrElse .uAllowanceLogin(Y).yPassword(Z) <> yAliasPassword(Z) Then
                                                                        bFound = False
                                                                        Exit For
                                                                    End If
                                                                Next Z

                                                                If bFound = True Then
                                                                    'Ok, found it
                                                                    lAliasID = .lAllowanceIdx(Y)
                                                                    lRights = .uAllowanceLogin(Y).lRights
                                                                    .lAliasingPlayerID = .oAllowances(Y).ObjectID
                                                                    oPlayer = .oAllowances(Y)

                                                                    If oPlayer.AccountStatus <> AccountStatusType.eActiveAccount AndAlso oPlayer.AccountStatus <> AccountStatusType.eMondelisActive Then
                                                                        lID = LoginResponseCodes.eAccountInactive
                                                                    End If
                                                                    Exit For
                                                                End If
                                                            End If
                                                        Next Y
                                                        If lID = LoginResponseCodes.eAccountInactive Then Exit For
                                                    End If


                                                    If bFound = False Then
                                                        'return bad result
                                                        ySame = 0
                                                        Exit For
                                                    End If
                                                End If

                                                'Do our associations here...
                                                mlClientPlayer(lSocketIndex) = glPlayerIdx(X)
                                                moClients(lSocketIndex).lSpecificID = glPlayerIdx(X)
                                                .oSocket = moClients(lSocketIndex)

                                                'Set the last login time...
                                                .LastLogin = Now
                                            End If

                                        Case AccountStatusType.eBannedAccount
                                            lID = LoginResponseCodes.eAccountBanned
                                        Case AccountStatusType.eInactiveAccount
                                            lID = LoginResponseCodes.eAccountInactive
                                        Case AccountStatusType.eMondelisInactive
                                            lID = LoginResponseCodes.eMondelisInactive
                                        Case AccountStatusType.eSuspendedAccount
                                            lID = LoginResponseCodes.eAccountSuspended
                                        Case Else
                                            lID = LoginResponseCodes.eUnknownFailure
                                    End Select
                                End If

                                Exit For
                            ElseIf ySame = 1 Then
                                Exit For
                            End If
                        End If
                    End With

                End If
            Next X

            'Ok, check our ySame...
            If ySame = 0 Then
                lID = LoginResponseCodes.eInvalidUserName
            ElseIf ySame = 1 Then
                lID = LoginResponseCodes.eInvalidPassword
            End If

            'Now, send our response...
            If lID < 1 Then
                lAliasID = lID
                lRights = 0
            End If
            Dim yResp(13) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerLoginResponse).CopyTo(yResp, 0)
            System.BitConverter.GetBytes(lAliasID).CopyTo(yResp, 2)
            System.BitConverter.GetBytes(lID).CopyTo(yResp, 6)
            System.BitConverter.GetBytes(lRights).CopyTo(yResp, 10)
            moClients(lSocketIndex).SendData(yResp)

            If oPlayer Is Nothing = False AndAlso lAliasID > 0 Then
                'ok, let's send what we need
                Dim oPrimary As ServerObject = Nothing
                If oPlayer.iStartedEnvirTypeID = ObjectType.ePlanet Then
                    'ok, find the planet
                    Dim oPlanet As Planet = GetEpicaPlanet(oPlayer.lStartedEnvirID)

                    If oPlanet Is Nothing = False Then
                        If oPlanet.ParentSystem Is Nothing = False AndAlso oPlanet.ParentSystem.oPrimaryServer Is Nothing = False Then
                            oPrimary = oPlanet.ParentSystem.oPrimaryServer
                        End If
                    Else
                        'new player... is assumed
                        'TODO: This is a bad assumption
                        'ok, find a place to place it
                        oPrimary = GetPrimaryServerForSpawn()
                        If oPrimary Is Nothing = True Then
                            'TODO: Generate new geography, place it and do it up
                        End If
                    End If

                Else
                    'new player... is assumed
                    'TODO: This is a bad assumption
                    'ok, find a place to place it
                    oPrimary = GetPrimaryServerForSpawn()
                    If oPrimary Is Nothing = True Then
                        'TODO: Generate new geography, place it and do it up
                    End If
                    Dim yAssign(5) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.ePlayerDomainAssignment).CopyTo(yAssign, 0)
                    System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yAssign, 2)
                    oPrimary.oSocket.SendData(yAssign)
                End If
                lPos = 0
                ReDim yResp(25)
                System.BitConverter.GetBytes(GlobalMessageCode.ePathfindingConnectionInfo).CopyTo(yResp, lPos) : lPos += 2
                Dim bFound As Boolean = False
                With oPrimary
                    If .oSocket Is Nothing = False Then oPlayer.lOwnerPrimaryIdx = .oSocket.SocketIndex
                    If .uListeners Is Nothing = False Then
                        For Y = 0 To .uListeners.GetUpperBound(0)
                            If .uListeners(Y).lConnectionType = ConnectionType.eClientConnection Then
                                bFound = True
                                If .uListeners(Y).sIPAddress.ToUpper = "EXTERNALIPADDY" Then
                                    StringToBytes(gsExternalAddress).CopyTo(yResp, lPos) : lPos += 20
                                Else
                                    StringToBytes(.uListeners(Y).sIPAddress).CopyTo(yResp, lPos) : lPos += 20
                                End If
                                System.BitConverter.GetBytes(.uListeners(Y).lPortNumber).CopyTo(yResp, lPos) : lPos += 4
                                Exit For
                            End If
                        Next Y
                    End If
                End With
                If bFound = False Then
                    LogEvent(LogEventType.CriticalError, "HandleLoginRequest (client): Could not find Primary Listener")
                End If
                moClients(lSocketIndex).SendData(yResp)
            End If

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "HandleLoginRequest: " & ex.Message)
            Try
                moClients(lSocketIndex).Disconnect()
            Catch
            End Try
        End Try
    End Sub

    Private Sub HandleOperatorRequestPassThru(ByVal yData() As Byte, ByVal lIndex As Int32)
        'ok, a primary is requesting data that it does not have from another primary... here is what we do...
        Dim lPos As Int32 = 2
        '1) Put the index of the requesting server into the request... we will use this when the requested primary responds
        System.BitConverter.GetBytes(lIndex).CopyTo(yData, lPos) : lPos += 4
        lPos += 4        'for playerid isnt important

        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        'ok, determine what primary to use
        Dim lForwardToIdx As Int32 = -1
        Select Case iObjTypeID
            Case ObjectType.ePlayer
                Dim oPlayer As Player = GetEpicaPlayer(lObjID)
                If oPlayer Is Nothing = False Then
                    lForwardToIdx = oPlayer.lOwnerPrimaryIdx
                End If
            Case ObjectType.eSolarSystem
                Dim oSystem As SolarSystem = GetEpicaSystem(lObjID)
                If oSystem Is Nothing = False Then
                    lForwardToIdx = oSystem.oPrimaryServer.oSocket.SocketIndex
                End If
        End Select

        If lForwardToIdx > -1 Then
            moServers(lForwardToIdx).SendData(yData)
        Else
            LogEvent(LogEventType.CriticalError, "HandleOperatorRequestPassThru ForwardToIdx is invalid!")
        End If


    End Sub
    Private Sub HandleOperatorResponsePassThru(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lPrimIdx As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'ok, lPrimIdx is our index in our array to send the msg to
        If lPrimIdx > -1 AndAlso lPrimIdx <= mlServerUB Then
            moServers(lPrimIdx).SendData(yData)
        Else : LogEvent(LogEventType.Warning, "Primary Pass Thru did not respond with valid Primary Index.")
        End If
    End Sub

    Private Sub HandlePlayerIsDead(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2       'for msgcode
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yRespawnFlag As Byte = yData(lPos) : lPos += 1
        Dim yForcedRestart As Byte = yData(lPos) : lPos += 1

        Dim bTurnIn As Boolean = False
        If lPlayerID < 0 Then
            bTurnIn = True
            lPlayerID = Math.Abs(lPlayerID)
        End If

        'ok, a primary server is alerting us to the fact that a player is telling it that they are dead. The primary server has already
        'verified that the player is in fact dead. The operator's job is to now:
        '   1) Alert all primaries of the death in the family
        '   2) Set an array of primary flag indicators indicating that a primary has or has not replied with a received msg of the event
        '   3) Lock out the player by returning "You are in the process of dying" messages


        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then

            If bTurnIn = True Then
                If oPlayer.lPrimaryReturned Is Nothing = False Then
                    Dim bAllTurnedIn As Boolean = True
                    For X As Int32 = 0 To oPlayer.lPrimaryReturned.GetUpperBound(0)
                        If oPlayer.lPrimaryReturned(X) = lIndex Then
                            oPlayer.lPrimaryReturned(X) = -1
                        ElseIf oPlayer.lPrimaryReturned(X) > -1 Then
                            bAllTurnedIn = False
                        End If
                    Next X

                    If bAllTurnedIn = True Then
                        oPlayer.swInDyingProcess = Nothing
                    End If
                End If
                Return
            Else
                If oPlayer.TotalPlayTime <> 0 Then oPlayer.lPlayerIcon = 0
                oPlayer.lAttachedPrimaryUB = -1
                oPlayer.swInDyingProcess = Stopwatch.StartNew()
                Dim lCnt As Int32 = 0
                For X As Int32 = 0 To mlServerUB
                    If oServerObject(X) Is Nothing = False AndAlso (oServerObject(X).lConnectionType = ConnectionType.ePrimaryServerApp) Then
                        lCnt += 1
                    End If
                Next X
                ReDim oPlayer.lPrimaryReturned(lCnt - 1)
                Dim lIdx As Int32 = -1
                For X As Int32 = 0 To mlServerUB
                    If oServerObject(X) Is Nothing = False AndAlso (oServerObject(X).lConnectionType = ConnectionType.ePrimaryServerApp) Then
                        lIdx += 1
                        oPlayer.lPrimaryReturned(lIdx) = X
                    End If
                Next X
            End If
        End If

        Me.SendToPrimaryServers(yData)


    End Sub

    Private Sub HandleSetPlayerSpecialAttribute(ByVal yData() As Byte, ByVal lIndex As Int32)

        'ok, let's see what we got
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iAttr As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        If iAttr = ePlayerSpecialAttributeSetting.ePlayerTitle Then
            'Ok, we need to do a test because the title that the player use to have may have changed
            Dim yNewTitle As Byte = CByte(System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
            Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
            If oPlayer Is Nothing = False Then
                oPlayer.PrimaryReportingTitleValueChange(lIndex, yNewTitle)
            End If
        Else
            'we are nothing more than a passthru for this
            Me.SendToPrimaryServers(yData)
        End If
    End Sub

    Private Sub HandleUpdatePlanetOwnership(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPlanetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To glPlanetUB
            If glPlanetIdx(X) = lPlanetID Then
                lIdx = X
                Exit For
            End If
        Next X
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)

        If lIdx = -1 OrElse oPlayer Is Nothing Then Return
        Dim oPlanet As Planet = goPlanet(lIdx)
        If oPlanet Is Nothing Then Return

        If oPlanet.OwnerID <> oPlayer.ObjectID AndAlso oPlanet.OwnerID > 0 Then
            Dim oPrevOwner As Player = GetEpicaPlayer(oPlanet.OwnerID)
            If oPrevOwner Is Nothing = False Then
                oPrevOwner.RemovePlanetControl(lIdx)
            End If
        End If
        oPlanet.OwnerID = oPlayer.ObjectID
        oPlayer.AddPlanetControl(lIdx)

    End Sub

    Public Sub SubmitKeepAlives()
        Dim yMsg(3) As Byte
        'We send a message neither server cares about from this server
        System.BitConverter.GetBytes(CShort(2)).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestObject).CopyTo(yMsg, 2)

        For X As Int32 = 0 To mlServerUB
            If oServerObject(X) Is Nothing = False AndAlso oServerObject(X).oSocket Is Nothing = False AndAlso (oServerObject(X).lConnectionType = ConnectionType.eBoxOperator OrElse oServerObject(X).lConnectionType = ConnectionType.ePrimaryServerApp OrElse oServerObject(X).lConnectionType = ConnectionType.eEmailServerApp OrElse oServerObject(X).lConnectionType = ConnectionType.eBackupOperator) Then
                Try
                    oServerObject(X).oSocket.SendLenAppendedData(yMsg)
                Catch
                End Try
            End If
        Next X
    End Sub
    Private Sub HandleKeepAliveMsg(ByRef yData() As Byte, ByVal lIndex As Int32)

        If gyBackupOperator = eyOperatorState.BackupOperator Then
            'ok, operator is telling me it is alive
            If goOperatorSW Is Nothing Then
                goOperatorSW = Stopwatch.StartNew
            Else
                goOperatorSW.Reset()
                goOperatorSW.Start()
            End If
            Return
        End If

        If oServerObject Is Nothing = False AndAlso oServerObject(lIndex) Is Nothing = False Then
            If oServerObject(lIndex).lConnectionType = ConnectionType.eBoxOperator Then
                HandleBoxOperatorUpdate(yData, lIndex)
                'Return
            ElseIf oServerObject(lIndex).lConnectionType = ConnectionType.eBackupOperator Then
                'send the msg back to the backup operator as it is expecting a response
                moServers(lIndex).SendData(yData)
            End If

            oServerObject(lIndex).dtLastMsgRecd = Now
            oServerObject(lIndex).bResentKeepAlive = False
            'If oServerObject(lIndex).lConnectionType <> ConnectionType.eBackupOperator Then
            '    ForwardToBackupOperator(yData, lIndex)
            'End If
        End If
    End Sub
    Public Sub CheckKeepAliveStatus()
        Const l_MAX_KEEP_ALIVE As Int32 = 60000         'one minute
        Const l_FIRST_KEEP_ALIVE As Int32 = 30000       '30 seconds


        For X As Int32 = 0 To mlServerUB
            If oServerObject(X) Is Nothing = False AndAlso _
               (oServerObject(X).lConnectionType = ConnectionType.ePrimaryServerApp OrElse _
               oServerObject(X).lConnectionType = ConnectionType.eBoxOperator OrElse _
               oServerObject(X).lConnectionType = ConnectionType.eEmailServerApp OrElse _
               oServerObject(X).lConnectionType = ConnectionType.eBackupOperator) Then

                Dim lTotalMilliseconds As Int32 = CInt(Now.Subtract(oServerObject(X).dtLastMsgRecd).TotalMilliseconds)

                If lTotalMilliseconds > l_MAX_KEEP_ALIVE Then
                    If oServerObject(X).bResentKeepAlive = True Then
                        If oServerObject(X).lConnectionType = ConnectionType.eOperator Then
                            LogEvent(LogEventType.Informational, "Server Outage: Operator!")
                            gyBackupOperator = eyOperatorState.EmergencyOperator
                            bAcceptNewClients = True
                            SendSupportEmail("Lost contact with Operator. Switching to Emergency Operator. Accepting clients and attempting to rebound.", "DSE Operator Server Down")
                        ElseIf oServerObject(X).lConnectionType = ConnectionType.eBoxOperator Then
                            'box operator, not good
                            LogEvent(LogEventType.CriticalError, "Server Outage: Box " & oServerObject(X).lBoxOperatorID & ", " & oServerObject(X).lConnectionType.ToString)
                        ElseIf oServerObject(X).lConnectionType = ConnectionType.eEmailServerApp Then
                            'email server
                            LogEvent(LogEventType.CriticalError, "Server Outage: Box " & oServerObject(X).lBoxOperatorID & ", " & oServerObject(X).lConnectionType.ToString)
                            'Clear Email Server's Details (so that primary requests do not get answered)
                            'Is the box operator associated to the email server available?
                            '  NO: Mark box operator critical, box operator handles this from here on, break out.
                            '  YES: Query Email process
                            '    Is Process dead?
                            '      NO: Kill Process
                            '    Spawn new Email Process
                            '    Email Start up Normal
                            '  Email sends Operator details
                            '  Operator stores details
                            'Primaries will continue to request the details, when the rebound is done, the details will be available
                        ElseIf oServerObject(X).lConnectionType = ConnectionType.ePrimaryServerApp Then
                            'primary server
                            LogEvent(LogEventType.CriticalError, "Server Outage: Box " & oServerObject(X).lBoxOperatorID & ", " & oServerObject(X).lConnectionType.ToString)
                            'Ok, listen for Stray Region and PF servers...
                            'If Is the Box Operator associated to the Primary available?
                            Dim bBoxAvail As Boolean = False
                            If oServerObject(X).oBoxOperator Is Nothing = False Then
                                If oServerObject(X).oBoxOperator.oSocket Is Nothing = False Then
                                    bBoxAvail = True
                                    '  YES: Query the Primary Process ID
                                    '    Box Operator gets it - is the Primary Process Dead?
                                    '      Yes - Kill It
                                    '    Box Operator responds

                                    'Dim yKillProc() As Byte
                                    'System.BitConverter.GetBytes(GlobalMessageCode.eRemoveObject).CopyTo(yKillProc, 0)
                                End If
                            End If
                            If bBoxAvail = False Then
                                '  NO: Is the Box Operator already Critical? If not, mark it critical. Box Operator rebound handles this. break out.
                            End If

                            '  Spawn new Primary Process
                            '  Primary Starts up Normal
                            '  Primary will send Ready
                            '  Any Stray PF's and Regions are notified
                            '  Operator sends Primary data to the strays
                            '  Strays connect to primary
                            '  Strays pass relevant data immediately
                            '  Primary Updates as necessary
                            '  Strays receive all start, primary starts, all is good
                        ElseIf oServerObject(X).lConnectionType = ConnectionType.eBackupOperator Then
                            'backup operator!!!
                            LogEvent(LogEventType.CriticalError, "Server Outage: Box " & oServerObject(X).lBoxOperatorID & ", " & oServerObject(X).lConnectionType.ToString)
                        Else
                            'No issue
                        End If
                    Else
                        'Likely, its the midnight roll-over
                        If Now.Hour <> 0 Then
                            'ok, maybe not...
                            Dim yMsg(3) As Byte
                            'We send a message neither server cares about from this server
                            System.BitConverter.GetBytes(CShort(2)).CopyTo(yMsg, 0)
                            System.BitConverter.GetBytes(GlobalMessageCode.eRequestObject).CopyTo(yMsg, 2)
                            oServerObject(X).bResentKeepAlive = True
                            oServerObject(X).oSocket.SendLenAppendedData(yMsg)
                        End If
                    End If
                ElseIf lTotalMilliseconds > l_FIRST_KEEP_ALIVE Then
                    If oServerObject(X).bResentKeepAlive = False Then
                        oServerObject(X).bResentKeepAlive = True
                        Dim yMsg(3) As Byte
                        'We send a message neither server cares about from this server
                        System.BitConverter.GetBytes(CShort(2)).CopyTo(yMsg, 0)
                        System.BitConverter.GetBytes(GlobalMessageCode.eRequestObject).CopyTo(yMsg, 2)
                        oServerObject(X).bResentKeepAlive = True
                        oServerObject(X).oSocket.SendLenAppendedData(yMsg)
                        LogEvent(LogEventType.Informational, "Possible Server outage: Box " & oServerObject(X).lBoxOperatorID & ", " & oServerObject(X).lConnectionType.ToString)
                    End If
                End If
            End If
        Next X
    End Sub

    Public Function GetBoxOperator(ByVal lBoxOperatorID As Int32) As BoxOperator
        For X As Int32 = 0 To mlServerUB
            If oServerObject(X) Is Nothing = False AndAlso oServerObject(X).lConnectionType = ConnectionType.eBoxOperator Then
                If oServerObject(X).lBoxOperatorID = lBoxOperatorID Then
                    Return CType(oServerObject(X), BoxOperator)
                End If
            End If
        Next X
        Return Nothing
    End Function
    Public Sub ConnectToBoxOperator(ByVal lSocketIdx As Int32, ByVal sIP As String, ByVal lPort As Int32, ByRef oBoxOperator As ServerObject)
        If moServers(lSocketIdx) Is Nothing = False Then
            Try
                moServers(lSocketIdx).Disconnect()
                moServers(lSocketIdx) = Nothing
            Catch
            End Try
        End If
        oBoxOperator.oSocket = Nothing

        moServers(lSocketIdx) = New NetSock()
        'add event handlers
        AddHandler moServers(lSocketIdx).onConnect, AddressOf moServers_onConnect
        AddHandler moServers(lSocketIdx).onDataArrival, AddressOf moServers_onDataArrival
        AddHandler moServers(lSocketIdx).onDisconnect, AddressOf moServers_onDisconnect
        AddHandler moServers(lSocketIdx).onError, AddressOf moServers_onError
        moServers(lSocketIdx).SocketIndex = lSocketIdx
        oBoxOperator.oSocket = moServers(lSocketIdx)
        oBoxOperator.bSocketConnected = True

        moServers(lSocketIdx).Connect(sIP, lPort)
    End Sub

    Private Sub HandlePlayerInitialSetup(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim yUserName(19) As Byte
        Dim yPassword(19) As Byte

        Try

            Array.Copy(yData, lPos, yUserName, 0, 20) : lPos += 20
            Array.Copy(yData, lPos, yPassword, 0, 20) : lPos += 20

            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) <> -1 Then
                    With goPlayer(X)
                        Dim ySame As Byte = 2

                        For Y As Int32 = 0 To 19
                            If yUserName(Y) <> .PlayerUserName(Y) Then
                                ySame = 0
                                Exit For
                            End If
                        Next Y

                        If ySame = 2 Then
                            'Ok, check the password
                            For Y As Int32 = 0 To 19
                                If yPassword(Y) <> .PlayerPassword(Y) Then
                                    ySame = 1
                                    Exit For
                                End If
                            Next Y

                            If ySame = 2 Then
                                'Ok, found it... 
                                Dim yPlayerName(19) As Byte
                                Dim yEmpireName(19) As Byte
                                Dim yGender As Byte
                                Dim lIcon As Int32

                                Dim yNewUserName(19) As Byte
                                Dim yNewPassword(19) As Byte

                                Array.Copy(yData, lPos, yPlayerName, 0, 20) : lPos += 20
                                Array.Copy(yData, lPos, yEmpireName, 0, 20) : lPos += 20
                                yGender = yData(lPos) : lPos += 1
                                lIcon = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                Array.Copy(yData, lPos, yNewUserName, 0, 20) : lPos += 20
                                Array.Copy(yData, lPos, yNewPassword, 0, 20) : lPos += 20

                                Dim yResp(2) As Byte
                                System.BitConverter.GetBytes(GlobalMessageCode.ePlayerInitialSetup).CopyTo(yResp, 0)
                                yResp(2) = 0

                                If lIcon = 0 Then
                                    yResp(2) = 3
                                Else
                                    Try
                                        Dim sPlayerName As String = BytesToString(yPlayerName)
                                        Dim sTest As String = FilterBadWords(sPlayerName)
                                        If sTest.ToUpper <> sPlayerName.ToUpper Then
                                            yResp(2) = 5
                                            Exit Try
                                        End If
                                        Dim sEmpireName As String = BytesToString(yEmpireName)
                                        sTest = FilterBadWords(sEmpireName)
                                        If sTest.ToUpper <> sEmpireName.ToUpper Then
                                            yResp(2) = 6
                                            Exit Try
                                        End If
                                        Dim sSQL As String = "SELECT COUNT(*) FROM tblPlayer WHERE PlayerName = '" & MakeDBStr(sPlayerName) & "' and PlayerID <> " & .ObjectID
                                        Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand(sSQL, goCN)
                                        Dim oData As OleDb.OleDbDataReader = oComm.ExecuteReader()
                                        Dim lCnt As Int32 = 0
                                        If oData.Read = True Then
                                            lCnt = CInt(oData(0))
                                        End If
                                        oData.Close()
                                        oData = Nothing
                                        oComm.Dispose()
                                        oComm = Nothing

                                        If lCnt > 0 Then
                                            yResp(2) = 1
                                        Else
                                            sSQL = "SELECT COUNT(*) FROM tblPlayer WHERE EmpireName = '" & MakeDBStr(sEmpireName) & "' and PlayerID <> " & .ObjectID
                                            oComm = New OleDb.OleDbCommand(sSQL, goCN)
                                            oData = oComm.ExecuteReader()
                                            If oData.Read = True Then
                                                lCnt = CInt(oData(0))
                                            End If
                                            oData.Close()
                                            oData = Nothing
                                            oComm.Dispose()
                                            oComm = Nothing

                                            If lCnt > 0 Then yResp(2) = 2
                                        End If

                                        If lCnt = 0 Then
                                            Dim sNewUserName As String = BytesToString(yNewUserName)
                                            sSQL = "SELECT COUNT(*) FROM tblPlayer WHERE PlayerUserName = '" & MakeDBStr(sNewUserName) & "' AND PlayerID <> " & .ObjectID
                                            oComm = New OleDb.OleDbCommand(sSQL, goCN)
                                            oData = oComm.ExecuteReader()
                                            If oData.Read = True Then lCnt = CInt(oData(0))
                                            oData.Close()
                                            oData = Nothing
                                            oComm.Dispose()
                                            oComm = Nothing

                                            If lCnt > 0 Then yResp(2) = 4
                                        End If
                                    Catch
                                        yResp(2) = 4
                                    End Try
                                End If
                                ForwardToBackupOperator(yData, lIndex)
                                If gyBackupOperator = eyOperatorState.BackupOperator Then Return

                                If yResp(2) = 0 Then
                                    .yGender = yGender
                                    .lPlayerIcon = lIcon
                                    .PlayerName = yPlayerName
                                    .sPlayerNameProper = BytesToString(yPlayerName)
                                    .sPlayerName = .sPlayerNameProper.ToUpper
                                    .EmpireName = yEmpireName

                                    ReDim .PlayerUserName(19)
                                    StringToBytes(BytesToString(yNewUserName)).CopyTo(.PlayerUserName, 0)   'use to be toupper
                                    ReDim .PlayerPassword(19)
                                    StringToBytes(BytesToString(yNewPassword)).CopyTo(.PlayerPassword, 0)   'use to be toupper

                                    .DataChanged()
                                End If

                                moClients(lIndex).SendData(yResp)

                                If yResp(2) = 0 Then
                                    'ReDim yResp(45)
                                    'lPos = 0
                                    'System.BitConverter.GetBytes(EpicaMessageCode.ePlayerInitialSetup).CopyTo(yResp, lPos) : lPos += 2
                                    'System.BitConverter.GetBytes(.ObjectID).CopyTo(yResp, lPos) : lPos += 4
                                    '.PlayerUserName.CopyTo(yResp, lPos) : lPos += 20
                                    '.PlayerPassword.CopyTo(yResp, lPos) : lPos += 20
                                    For lSrvr As Int32 = 0 To mlServerUB
                                        If moServers(lSrvr) Is Nothing = False AndAlso moServers(lSrvr).IsConnected = True Then
                                            Try
                                                'MSC - 4/11/08 - not sure why this was Is Nothing... OrElse... changed it.
                                                '  if it causes issues, change it back and document WHY it is this way.
                                                'If oServerObject(lSrvr) Is Nothing OrElse oServerObject(lSrvr).lConnectionType = ConnectionType.ePrimaryServerApp Then
                                                If oServerObject(lSrvr) Is Nothing = False AndAlso oServerObject(lSrvr).lConnectionType = ConnectionType.ePrimaryServerApp Then
                                                    moServers(lSrvr).SendData(yData)
                                                End If
                                            Catch ex As Exception
                                                '
                                            End Try
                                        End If
                                    Next lSrvr
                                End If

                                Exit For

                            End If
                        End If
                    End With
                End If
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "HandlePlayerInitialSetup: " & ex.Message)
        End Try
    End Sub

    Private Sub HandleSetPlayerInitialSetup(ByRef yData() As Byte, ByVal lIndex As Int32)

        'ONLY to be called if I am in backup mode
        If gyBackupOperator <> eyOperatorState.BackupOperator Then
            HandlePlayerInitialSetup(yData, lIndex)
            Return
        End If

        Dim lPos As Int32 = 2
        Dim yUserName(19) As Byte
        Dim yPassword(19) As Byte

        Try

            Array.Copy(yData, lPos, yUserName, 0, 20) : lPos += 20
            Array.Copy(yData, lPos, yPassword, 0, 20) : lPos += 20

            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) <> -1 Then
                    With goPlayer(X)
                        Dim ySame As Byte = 2

                        For Y As Int32 = 0 To 19
                            If yUserName(Y) <> .PlayerUserName(Y) Then
                                ySame = 0
                                Exit For
                            End If
                        Next Y

                        If ySame = 2 Then
                            'Ok, check the password
                            For Y As Int32 = 0 To 19
                                If yPassword(Y) <> .PlayerPassword(Y) Then
                                    ySame = 1
                                    Exit For
                                End If
                            Next Y

                            If ySame = 2 Then
                                'Ok, found it... 
                                Dim yPlayerName(19) As Byte
                                Dim yEmpireName(19) As Byte
                                Dim yGender As Byte
                                Dim lIcon As Int32

                                Dim yNewUserName(19) As Byte
                                Dim yNewPassword(19) As Byte

                                Array.Copy(yData, lPos, yPlayerName, 0, 20) : lPos += 20
                                Array.Copy(yData, lPos, yEmpireName, 0, 20) : lPos += 20
                                yGender = yData(lPos) : lPos += 1
                                lIcon = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                Array.Copy(yData, lPos, yNewUserName, 0, 20) : lPos += 20
                                Array.Copy(yData, lPos, yNewPassword, 0, 20) : lPos += 20

                                Exit For

                            End If
                        End If
                    End With
                End If
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "HandleSetPlayerInitialSetup: " & ex.Message)
        End Try
    End Sub

    Private Sub HandleCrossPrimaryBudgetAdjust(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            If oPlayer.lOwnerPrimaryIdx > -1 AndAlso oPlayer.lOwnerPrimaryIdx <= mlServerUB Then
                moServers(oPlayer.lOwnerPrimaryIdx).SendData(yData)
            End If
        End If
    End Sub

    Private Sub HandleAddObjectCommand(ByRef yData() As Byte, ByVal lIndex As Int32)
        Try

            If lIndex > -1 AndAlso oServerObject(lIndex) Is Nothing = False AndAlso oServerObject(lIndex).lConnectionType = ConnectionType.eBoxOperator Then
                'ok, a box operator is telling us that a spawn request was made, the process was spawned, and we now have the process id
                Dim lPos As Int32 = 2
                Dim lProcType As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lRequestIdx As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lResult As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                'ok, is the lREsult = int32.minvalue
                If lResult = Int32.MinValue Then
                    'TODO: Unabel to spawn the process, what do we do now?
                End If

                'ForwardToBackupOperator(yData, lIndex)
            Else
                Dim lPos As Int32 = 2       'for msgcode
                Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                Select Case iTypeID
                    Case ObjectType.ePlayer
                        'ok, primary server is telling us to load a new player
                        If LoadSinglePlayer(lID) = True AndAlso gyBackupOperator <> eyOperatorState.BackupOperator Then
                            For X As Int32 = 0 To Me.mlServerUB
                                If oServerObject(X) Is Nothing = False AndAlso oServerObject(X).lConnectionType = ConnectionType.ePrimaryServerApp Then
                                    Try
                                        oServerObject(X).oSocket.SendData(yData)
                                    Catch
                                    End Try
                                End If
                            Next X
                        End If
                    Case ObjectType.eMineral
                        'Ok, the server is telling me that a mineral cache depleted somewhere and we need to respawn it...
                        Dim lEID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        Dim iETypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        MineralGeographyRel.SpawnNewMineralCache(lID, lEID, iETypeID)
                    Case Int16.MinValue
                        'Ok, likely, i am a backup operator and the operator is telling me a new server connection has appeared
                        If gyBackupOperator = eyOperatorState.BackupOperator Then

                            Try

                                'the lID value is the index of the new server connection
                                If lID > mlServerUB Then
                                    mlServerUB = lID
                                    ReDim Preserve moServers(mlServerUB)
                                    ReDim Preserve oServerObject(mlServerUB)
                                End If

                                moServers(lID) = Nothing
                                oServerObject(lID) = New ServerObject()
                                oServerObject(lID).oSocket = moServers(lID)
                                oServerObject(lID).bSocketConnected = True
                                'moServers(lID).SocketIndex = lID
                                'moServers(lID).lSpecificID = lID

                                LogEvent(LogEventType.Informational, "Operator Informing Connected Server Socket " & lID)

                            Catch ex As Exception
                                LogEvent(LogEventType.CriticalError, "HandleAddObjectCommand.Server: " & ex.Message)
                            End Try

                        End If
                    Case Else ' ObjectType.eUnitDef, ObjectType.eFacilityDef, Techs...
                        Dim lOwnerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        'ok, we have the ownerid, now, get the player
                        Dim oPlayer As Player = GetEpicaPlayer(lOwnerID)
                        If oPlayer Is Nothing = False Then
                            'ok, now, determine if the msg was received from the owner primary
                            If oPlayer.lOwnerPrimaryIdx <> lIndex Then
                                'no, it was received from someone else... send to the owner primary
                                moServers(oPlayer.lOwnerPrimaryIdx).SendData(yData)
                            End If
                            'now, go thru interested parties
                            For X As Int32 = 0 To oPlayer.lAttachedPrimaryUB
                                Dim lAttached As Int32 = oPlayer.lAttachedPrimaryIdx(X)
                                If lAttached <> lIndex AndAlso lAttached > -1 AndAlso lAttached <= mlServerUB Then
                                    moServers(lAttached).SendData(yData)
                                End If
                            Next X
                        End If

                End Select
            End If

        Catch
        End Try

    End Sub

    Private Sub HandlePlayerAliasConfig(ByRef yData() As Byte, ByVal lIndex As Int32)
        ForwardToBackupOperator(yData, lIndex)
        Try
            'MsgCode 2, Type(1), AliasPlayer(20) but is as PlayerID (4) and OtherPlayer(4)... other 12 is nothing, AliasUN(20), AliasPW(20), Rights(4)
            Dim lPos As Int32 = 2   'for msgcode
            Dim yType As Byte = yData(lPos) : lPos += 1

            Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lOtherPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            lPos += 12      'remainder of the aliasplayer msg

            Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
            If oPlayer Is Nothing Then Return

            If yType = 0 Then       'remove
                For X As Int32 = 0 To oPlayer.lAliasUB
                    If oPlayer.lAliasIdx(X) = lOtherPlayerID Then
                        Dim oOther As Player = oPlayer.oAliases(X)
                        oPlayer.lAliasIdx(X) = -1
                        oPlayer.uAliasLogin(X).lRights = 0
                        ReDim oPlayer.uAliasLogin(X).yPassword(19)
                        ReDim oPlayer.uAliasLogin(X).yUserName(19)
                        oPlayer.oAliases(X) = Nothing

                        For Y As Int32 = 0 To oOther.lAllowanceUB
                            If oOther.lAllowanceIdx(Y) = oPlayer.ObjectID Then
                                oOther.lAllowanceIdx(Y) = -1
                                oOther.oAllowances(Y) = Nothing
                                oOther.uAllowanceLogin(Y).lRights = 0
                                ReDim oOther.uAllowanceLogin(Y).yPassword(19)
                                ReDim oOther.uAllowanceLogin(Y).yUserName(19)
                            End If
                        Next Y
                        Exit For
                    End If
                Next X
            Else
                'add/update - get our remaining items
                Dim yUserName(19) As Byte
                Dim yPassword(19) As Byte
                Dim lRights As Int32 = 0
                Array.Copy(yData, lPos, yUserName, 0, 20) : lPos += 20
                Array.Copy(yData, lPos, yPassword, 0, 20) : lPos += 20
                lRights = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                'see if it is an update
                Dim lIdx As Int32 = -1
                For X As Int32 = 0 To oPlayer.lAliasUB
                    If oPlayer.lAliasIdx(X) = lOtherPlayerID Then
                        lIdx = X
                        Exit For
                    End If
                Next X

                If lIdx = -1 Then
                    'Ok, need to add an alias... find the player
                    Dim oOther As Player = GetEpicaPlayer(lOtherPlayerID)
                    If oOther Is Nothing Then Return

                    oPlayer.AddPlayerAlias(oOther, yUserName, yPassword, lRights)
                    oOther.AddPlayerAliasAllowance(oPlayer, yUserName, yPassword, lRights)
                Else
                    'An update...
                    Dim oOther As Player = oPlayer.oAliases(lIdx)
                    If oOther Is Nothing = False Then
                        oPlayer.AddPlayerAlias(oOther, yUserName, yPassword, lRights)
                        oOther.AddPlayerAliasAllowance(oPlayer, yUserName, yPassword, lRights)
                    End If
                End If
            End If

            'Now, send our update
            If gyBackupOperator <> eyOperatorState.BackupOperator Then
                For X As Int32 = 0 To mlServerUB
                    If oServerObject(X) Is Nothing = False AndAlso oServerObject(X).lConnectionType = ConnectionType.ePrimaryServerApp Then
                        Try
                            oServerObject(X).oSocket.SendData(yData)
                        Catch
                        End Try
                    End If
                Next X
            End If

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "HandlePlayerAliasConfig: " & ex.Message)
        End Try
    End Sub

    Private Sub HandleBoxOperatorUpdate(ByRef yData() As Byte, ByVal lIndex As Int32)
        'ForwardToBackupOperator(yData, lIndex)
        If oServerObject(lIndex) Is Nothing = False AndAlso oServerObject(lIndex).lConnectionType = ConnectionType.eBoxOperator Then
            CType(oServerObject(lIndex), BoxOperator).HandleBoxOperatorUpdate(yData)
        End If
    End Sub

    Public Sub SendToPrimaryServers(ByVal yMsg() As Byte)
        For X As Int32 = 0 To mlServerUB
            If oServerObject(X) Is Nothing = False AndAlso (oServerObject(X).lConnectionType = ConnectionType.ePrimaryServerApp) Then
                moServers(X).SendData(yMsg)
            End If
        Next X
    End Sub

    Public Sub SendReboundMsgToServers()
        'Should send the operator a series of messages informing it of the situation
        SendResyncMessages(moOperator)

        Dim yMsg(3) As Byte
        System.BitConverter.GetBytes(CShort(2)).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(GlobalMessageCode.eBackupOperatorSyncMsg).CopyTo(yMsg, 2)

        For X As Int32 = 0 To mlServerUB
            If oServerObject(X) Is Nothing = False AndAlso (oServerObject(X).lConnectionType = ConnectionType.eBoxOperator OrElse _
             oServerObject(X).lConnectionType = ConnectionType.ePrimaryServerApp OrElse oServerObject(X).lConnectionType = ConnectionType.eEmailServerApp) Then
                If moServers(X) Is Nothing = False Then moServers(X).SendLenAppendedData(yMsg)
            End If
        Next X

    End Sub

    Private Sub SendResyncMessages(ByRef oSocket As NetSock)
        Dim yMsg() As Byte
        For X As Int32 = 0 To mlServerUB
            If oServerObject(X) Is Nothing = False Then
                yMsg = SendBackupOpServerReady(oServerObject(X))
                oSocket.SendData(yMsg)
                'With oServerObject(X)
                '    Dim lPos As Int32 = 0
                '    Dim lCnt As Int32 = 0
                '    If .uListeners Is Nothing = False Then
                '        lCnt = .uListeners.Length
                '    End If

                '    ReDim yMsg(25 + (lCnt * 28))

                '    System.BitConverter.GetBytes(X).CopyTo(yMsg, lPos) : lPos += 4
                '    System.BitConverter.GetBytes(GlobalMessageCode.eBackupOperatorSyncMsg).CopyTo(yMsg, lPos) : lPos += 2
                '    System.BitConverter.GetBytes(X).CopyTo(yMsg, lPos) : lPos += 4
                '    System.BitConverter.GetBytes(.lConnectionType).CopyTo(yMsg, lPos) : lPos += 4
                '    System.BitConverter.GetBytes(.lBoxOperatorID).CopyTo(yMsg, lPos) : lPos += 4
                '    System.BitConverter.GetBytes(.lProcessID).CopyTo(yMsg, lPos) : lPos += 4

                '    System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4
                '    If .uListeners Is Nothing = False Then
                '        For Y As Int32 = 0 To .uListeners.GetUpperBound(0)
                '            StringToBytes(.uListeners(Y).sIPAddress).CopyTo(yMsg, lPos) : lPos += 20
                '            System.BitConverter.GetBytes(.uListeners(Y).lPortNumber).CopyTo(yMsg, lPos) : lPos += 4
                '            System.BitConverter.GetBytes(.uListeners(Y).lConnectionType).CopyTo(yMsg, lPos) : lPos += 4
                '        Next Y
                '    End If

                '    oSocket.SendData(yMsg)
                'End With
            End If
        Next X
    End Sub

    Private Sub HandleSyncMessageFromBackup(ByRef yData() As Byte)
        Dim lPos As Int32 = 2

        Dim lIdx As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lType As ConnectionType = CType(System.BitConverter.ToInt32(yData, lPos), ConnectionType) : lPos += 4
        Dim lBoxID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If lIdx > mlServerUB Then
            mlServerUB = lIdx
            ReDim Preserve moServers(mlServerUB)
            ReDim Preserve oServerObject(mlServerUB)
        End If
        If oServerObject(lIdx) Is Nothing Then
            oServerObject(lIdx) = New ServerObject()
            oServerObject(lIdx).lConnectionType = lType
        End If

        If oServerObject(lIdx).lConnectionType <> lType Then
            LogEvent(LogEventType.CriticalError, "HandleSyncMessageFromBackup informed me to make Idx a " & lType.ToString & ". It is currently registered as " & oServerObject(lIdx).lConnectionType)
            Return
        End If

        With oServerObject(lIdx)
            .lConnectionType = lType
            .lBoxOperatorID = lBoxID
            ReDim .uListeners(lCnt - 1)
            For X As Int32 = 0 To lCnt - 1
                .uListeners(X).sIPAddress = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                .uListeners(X).lPortNumber = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .uListeners(X).lConnectionType = CType(System.BitConverter.ToInt32(yData, lPos), ConnectionType) : lPos += 4
            Next X
        End With

    End Sub

    Private Sub HandleSyncMessageFromOperator(ByRef yData() As Byte)
        Dim lPos As Int32 = 2

        Dim lIdx As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lType As ConnectionType = CType(System.BitConverter.ToInt32(yData, lPos), ConnectionType) : lPos += 4
        Dim lBoxID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If lIdx > mlServerUB Then
            mlServerUB = lIdx
            ReDim Preserve moServers(mlServerUB)
            ReDim Preserve oServerObject(mlServerUB)
        End If
        If oServerObject(lIdx) Is Nothing Then
            If lType = ConnectionType.eBoxOperator Then
                oServerObject(lIdx) = New BoxOperator(New ServerObject())
            Else
                oServerObject(lIdx) = New ServerObject()
            End If

            oServerObject(lIdx).lConnectionType = lType
        End If

        With oServerObject(lIdx)
            .lConnectionType = lType
            .lBoxOperatorID = lBoxID
            ReDim .uListeners(lCnt - 1)
            For X As Int32 = 0 To lCnt - 1
                .uListeners(X).sIPAddress = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                .uListeners(X).lPortNumber = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .uListeners(X).lConnectionType = CType(System.BitConverter.ToInt32(yData, lPos), ConnectionType) : lPos += 4
            Next X
        End With
    End Sub

    Public Function FindNeighborhoodSuitor(ByVal lConnType As ConnectionType, ByRef oPreferredPrimary As ServerObject) As ServerObject
        'TODO: Finish this
        Select Case lConnType
            Case ConnectionType.ePrimaryServerApp
                'returns a primary server on a box with enough resources for a new system
                'NOT A GOOD IDEA HERE
                For X As Int32 = 0 To mlServerUB
                    If oServerObject(X) Is Nothing = False AndAlso oServerObject(X).lConnectionType = ConnectionType.ePrimaryServerApp Then
                        Return oServerObject(X)
                    End If
                Next X
            Case ConnectionType.eRegionServerApp
                'returns a region server on a box with enough resources for a new system
                'NOT A GOOD IDEA HERE
                Dim oMinSrvr As ServerObject = Nothing
                Dim lMinIDCnt As Int32 = Int32.MaxValue
                For X As Int32 = 0 To mlServerUB
                    If oServerObject(X) Is Nothing = False AndAlso oServerObject(X).lConnectionType = ConnectionType.eRegionServerApp Then

                        'If oServerObject(X).lSpawnRequestIdx > -1 AndAlso uSpawnRequests(oServerObject(X).lSpawnRequestIdx).lSpawnID Is Nothing = False Then
                        If oServerObject(X).lSpawnID Is Nothing = False Then
                            'Dim lCnt As Int32 = uSpawnRequests(oServerObject(X).lSpawnRequestIdx).GetSpawnIDScore()
                            Dim lCnt As Int32 = oServerObject(X).GetSpawnIDScore()
                            If lCnt < lMinIDCnt Then
                                lMinIDCnt = lCnt
                                oMinSrvr = oServerObject(X)
                            End If
                        Else
                            Return oServerObject(X)
                        End If
                        'Return oServerObject(X)
                    End If
                Next X
                If oMinSrvr Is Nothing = False Then Return oMinSrvr
        End Select

        Return Nothing
    End Function

    Private mbFirstRegionHandled As Boolean = False
    Private mlRgn As Int32 = 3
    Public Function GetAvailableBoxOperator(ByVal lConnType As ConnectionType) As Int32
        If gb_IS_TEST_SERVER = True Then Return 1

        If lConnType = ConnectionType.eEmailServerApp Then Return 1
        If lConnType = ConnectionType.ePrimaryServerApp Then Return 2
        If lConnType = ConnectionType.ePathfindingServerApp Then Return 3
        If lConnType = ConnectionType.eRegionServerApp Then
            mlRgn += 1
            If mlRgn > 6 Then mlRgn = 4
            Return mlRgn
        End If
        'If lConnType = ConnectionType.eEmailServerApp Then Return 2
        'If lConnType = ConnectionType.ePrimaryServerApp Then Return 1
        'If lConnType = ConnectionType.ePathfindingServerApp Then Return 1
        'If lConnType = ConnectionType.eRegionServerApp Then 'Return 1
        '    If mbFirstRegionHandled = True Then Return 2
        '    mbFirstRegionHandled = True
        '    Return 1
        'End If

        'For X As Int32 = 0 To mlServerUB
        '    If oServerObject(X) Is Nothing = False AndAlso oServerObject(X).lConnectionType = ConnectionType.eBoxOperator Then
        '        'ok, check this box operator's usage
        '        With CType(oServerObject(X), BoxOperator)
        '            If .CanHandleNewServer(lConnType) = True Then Return .lBoxOperatorID
        '        End With
        '    End If
        'Next X
        Return -1
    End Function
    'Public Function GetAvailableBoxOperator(ByVal lConnType As ConnectionType) As Int32

    '    If lConnType = ConnectionType.eEmailServerApp Then Return 2
    '    If lConnType = ConnectionType.ePrimaryServerApp Then Return 1
    '    If lConnType = ConnectionType.ePathfindingServerApp Then Return 1
    '    If lConnType = ConnectionType.eRegionServerApp Then 'Return 1
    '        If mbFirstRegionHandled = True Then Return 2
    '        mbFirstRegionHandled = True
    '        Return 1
    '    End If

    '    For X As Int32 = 0 To mlServerUB
    '        If oServerObject(X) Is Nothing = False AndAlso oServerObject(X).lConnectionType = ConnectionType.eBoxOperator Then
    '            'ok, check this box operator's usage
    '            With CType(oServerObject(X), BoxOperator)
    '                If .CanHandleNewServer(lConnType) = True Then Return .lBoxOperatorID
    '            End With
    '        End If
    '    Next X
    '    Return -1
    'End Function

    Private Sub HandlePlayerConnectedPrimary(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        Dim lOriginalIndex As Int32 = System.BitConverter.ToInt32(yData, lPos)
        If lOriginalIndex <> Int32.MinValue Then
            lOriginalIndex = lIndex
        Else
            'Ok, determine if the index of this primary is the index of the current...
            If oPlayer.lConnectedPrimaryIdx <> lIndex Then Return 'ok, i am being told that a primary i was connected to is now disconnected... however, i dont think that is the primary the player is connected to
            lOriginalIndex = -1
        End If
        System.BitConverter.GetBytes(lOriginalIndex).CopyTo(yData, lPos) : lPos += 4

        oPlayer.lConnectedPrimaryIdx = lOriginalIndex

        SendToPrimaryServers(yData)
    End Sub
    Private Sub HandlePlayerPrimaryOwner(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            oPlayer.lOwnerPrimaryIdx = lIndex
        End If
    End Sub

    Private Sub HandleForwardToPlayerAtPrimary(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2 'for msgcode
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        If oPlayer.lConnectedPrimaryIdx < 0 OrElse oPlayer.lConnectedPrimaryIdx > mlServerUB Then
            'TODO: What do we do? Tell the primary the player is disconnected?
        Else
            Dim oServer As NetSock = moServers(oPlayer.lConnectedPrimaryIdx)
            If oServer Is Nothing = False Then
                oServer.SendData(yData)
            End If
        End If
    End Sub

    Private Sub HandleEntityChangingPrimary(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lOwnerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lParentDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iParentDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim oDestServer As ServerObject = Nothing
        Dim oPlayer As Player = GetEpicaPlayer(lOwnerID)
        If oPlayer Is Nothing Then Return

        Select Case iDestTypeID
            Case ObjectType.ePlanet
                Dim oPlanet As Planet = GetEpicaPlanet(lDestID)
                If oPlanet Is Nothing = False AndAlso oPlanet.ParentSystem Is Nothing = False Then
                    oDestServer = oPlanet.ParentSystem.oPrimaryServer
                End If
            Case ObjectType.eSolarSystem
                Dim oSystem As SolarSystem = GetEpicaSystem(lDestID)
                If oSystem Is Nothing = False Then oDestServer = oSystem.oPrimaryServer
            Case ObjectType.eUnit, ObjectType.eFacility
                If iParentDestTypeID = ObjectType.ePlanet Then
                    Dim oPlanet As Planet = GetEpicaPlanet(lDestID)
                    If oPlanet Is Nothing = False AndAlso oPlanet.ParentSystem Is Nothing = False Then oDestServer = oPlanet.ParentSystem.oPrimaryServer
                ElseIf iParentDestTypeID = ObjectType.eSolarSystem Then
                    Dim oSystem As SolarSystem = GetEpicaSystem(lDestID)
                    If oSystem Is Nothing = False Then oDestServer = oSystem.oPrimaryServer
                End If
        End Select

        If oDestServer Is Nothing = False Then
            Dim bFound As Boolean = False
            Dim lFirstIdx As Int32 = -1
            Dim lNewDestIdx As Int32 = oDestServer.oSocket.SocketIndex
            For X As Int32 = 0 To oPlayer.lAttachedPrimaryUB
                If oPlayer.lAttachedPrimaryIdx(X) > -1 Then
                    If oPlayer.lAttachedPrimaryIdx(X) = lNewDestIdx Then
                        bFound = True
                        Exit For
                    End If
                ElseIf lFirstIdx = -1 Then
                    lFirstIdx = X
                End If
            Next X

            If bFound = False Then
                If lFirstIdx = -1 Then
                    lFirstIdx = oPlayer.lAttachedPrimaryUB + 1
                    ReDim Preserve oPlayer.lAttachedPrimaryIdx(lFirstIdx)
                    oPlayer.lAttachedPrimaryUB += 1
                End If
                oPlayer.lAttachedPrimaryIdx(lFirstIdx) = lNewDestIdx
                Dim yMsg(5) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.ePrimaryLoadSharedPlayerData).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, 2)
                oDestServer.oSocket.SendData(yMsg)
            End If

            'Now, send the msg to the primary
            oDestServer.oSocket.SendData(yData)

        End If
    End Sub

    Private Sub HandleAddFormation(ByVal yData() As Byte, ByVal lIndex As Int32)
        'Ok, here's what we cAre about... the owner id...
        Dim lPos As Int32 = 2    'for msgcode
        lPos += 4       'formationid
        lPos += 1       'default
        lPos += 1       'criteria
        lPos += 20      'name

        Dim lOwnerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oPlayer As Player = GetEpicaPlayer(lOwnerID)
        If oPlayer Is Nothing = False Then

            'Now, forward this message to the player's primary servers....
            If oPlayer.lOwnerPrimaryIdx <> lIndex Then
                oServerObject(oPlayer.lOwnerPrimaryIdx).oSocket.SendData(yData)
            End If
            For X As Int32 = 0 To oPlayer.lAttachedPrimaryUB
                If oPlayer.lAttachedPrimaryIdx(X) > -1 Then
                    oServerObject(oPlayer.lAttachedPrimaryIdx(X)).oSocket.SendData(yData)
                End If
            Next X

        End If
    End Sub

    Public Sub SendToServerIndex(ByVal lIdx As Int32, ByVal yMsg() As Byte)
        If lIdx > -1 AndAlso lIdx <= mlServerUB Then
            If oServerObject(lIdx) Is Nothing = False AndAlso oServerObject(lIdx).oSocket Is Nothing = False Then
                oServerObject(lIdx).oSocket.SendData(yMsg)
            End If
        End If
    End Sub
    Public Sub SendToAdminIndex(ByVal lIdx As Int32, ByVal yMsg() As Byte)
        If lIdx > -1 AndAlso lIdx <= mlAdminUB Then
            If moAdmins(lIdx) Is Nothing = False Then
                moAdmins(lIdx).SendData(yMsg)
            End If
        End If
    End Sub

    Private Sub HandleServerRequestServerDetails(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lConnType As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Select Case lConnType
            Case ConnectionType.eEmailServerApp
                For X As Int32 = 0 To mlServerUB
                    If oServerObject(X) Is Nothing = False AndAlso oServerObject(X).lConnectionType = ConnectionType.eEmailServerApp Then
                        'ok, is it connected
                        If moServers(X).IsConnected = True Then
                            With oServerObject(X)
                                For Y As Int32 = 0 To .uListeners.GetUpperBound(0)
                                    If .uListeners(Y).lConnectionType = ConnectionType.ePrimaryServerApp Then
                                        Dim yMsg(25) As Byte
                                        lPos = 0
                                        System.BitConverter.GetBytes(GlobalMessageCode.eEmailSettings).CopyTo(yMsg, lPos) : lPos += 2
                                        StringToBytes(.uListeners(Y).sIPAddress).CopyTo(yMsg, lPos) : lPos += 20
                                        System.BitConverter.GetBytes(.uListeners(Y).lPortNumber).CopyTo(yMsg, lPos) : lPos += 4
                                        moServers(lIndex).SendData(yMsg)
                                        Exit For
                                    End If
                                Next Y
                            End With

                            Exit For
                        End If
                    End If
                Next X
        End Select

    End Sub


    Private Sub HandleRequestGlobalPlayerScores(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yTypeID As Byte = yData(lPos) : lPos += 1
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        Dim lForPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If oPlayer Is Nothing Then Return

        If (yTypeID And 128) <> 0 Then
            'ok, a server is responding...
            Dim lRequestIdx As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            With oPlayer
                Dim blTemp As Int64 = .lTechnologyScore
                blTemp += System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                If blTemp > Int32.MaxValue Then blTemp = Int32.MaxValue
                .lTechnologyScore = CInt(blTemp)

                blTemp = .lWealthScore
                blTemp += System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                If blTemp > Int32.MaxValue Then blTemp = Int32.MaxValue
                .lWealthScore = CInt(blTemp)

                blTemp = .lPopulationScore
                blTemp += System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                If blTemp > Int32.MaxValue Then blTemp = Int32.MaxValue
                .lPopulationScore = CInt(blTemp)

                blTemp = .lProductionScore
                blTemp += System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                If blTemp > Int32.MaxValue Then blTemp = Int32.MaxValue
                .lProductionScore = CInt(blTemp)

                blTemp = .lMilitaryScore
                blTemp += System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                If blTemp > Int32.MaxValue Then blTemp = Int32.MaxValue
                .lMilitaryScore = CInt(blTemp)

                blTemp = .lDiplomacyScore
                blTemp += System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                If blTemp > Int32.MaxValue Then blTemp = Int32.MaxValue
                .lDiplomacyScore = CInt(blTemp)

                blTemp = .lTotalScore
                blTemp += System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                If blTemp > Int32.MaxValue Then blTemp = Int32.MaxValue
                .lTotalScore = CInt(blTemp)


                'Now, check if we are done
                .lPlayerScoreRequestCnt += 1
                If .lPlayerScoreRequestCnt > -1 Then
                    'ok, send it
                    Dim yOutMsg(42) As Byte
                    lPos = 0

                    System.BitConverter.GetBytes(GlobalMessageCode.eRequestGlobalPlayerScores).CopyTo(yOutMsg, lPos) : lPos += 2
                    System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yOutMsg, lPos) : lPos += 4
                    If (yTypeID And 128) <> 0 Then yTypeID = CByte(yTypeID Xor 128)
                    yOutMsg(lPos) = yTypeID : lPos += 1
                    System.BitConverter.GetBytes(lForPlayerID).CopyTo(yOutMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(-1I).CopyTo(yOutMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lTechnologyScore).CopyTo(yOutMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lWealthScore).CopyTo(yOutMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lPopulationScore).CopyTo(yOutMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lProductionScore).CopyTo(yOutMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lMilitaryScore).CopyTo(yOutMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lDiplomacyScore).CopyTo(yOutMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.lTotalScore).CopyTo(yOutMsg, lPos) : lPos += 4

                    moServers(lRequestIdx).SendData(yOutMsg)
                End If
            End With
        Else
            System.BitConverter.GetBytes(lIndex).CopyTo(yData, lPos) : lPos += 4

            'ok, a server is requesting it...
            Dim lCnt As Int32 = 1
            For X As Int32 = 0 To oPlayer.lAttachedPrimaryUB
                If oPlayer.lAttachedPrimaryIdx(X) > -1 AndAlso oPlayer.lOwnerPrimaryIdx <> oPlayer.lAttachedPrimaryIdx(X) Then
                    lCnt += 1
                End If
            Next X

            With oPlayer
                .lPlayerScoreRequestCnt = -lCnt
                .lTechnologyScore = 0
                .lWealthScore = 0
                .lPopulationScore = 0
                .lProductionScore = 0
                .lTotalScore = 0
                .lMilitaryScore = 0
                .lDiplomacyScore = 0
            End With

            For X As Int32 = 0 To oPlayer.lAttachedPrimaryUB
                If oPlayer.lAttachedPrimaryIdx(X) > -1 AndAlso oPlayer.lOwnerPrimaryIdx <> oPlayer.lAttachedPrimaryIdx(X) Then
                    moServers(oPlayer.lAttachedPrimaryIdx(X)).SendData(yData)
                End If
            Next X


        End If
    End Sub

    Private Sub HandleServerGetPlayerCurrentEnvironment(ByVal yData() As Byte, ByVal lIndex As Int32)
        'ok, a server is requesting the owning primary of the environment implied

        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lGalaxyID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lSystemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPlanetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yInDomain As Byte = yData(lPos) : lPos += 1

        Dim oPrimary As ServerObject = Nothing
        If iEnvirTypeID = ObjectType.ePlanet Then
            Dim oPlanet As Planet = GetEpicaPlanet(lEnvirID)
            If oPlanet Is Nothing = False Then
                oPrimary = oPlanet.ParentSystem.oPrimaryServer
            End If
        ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
            Dim oSystem As SolarSystem = GetEpicaSystem(lEnvirID)
            If oSystem Is Nothing = False Then
                oPrimary = oSystem.oPrimaryServer
            End If
        Else
            LogEvent(LogEventType.PossibleCheat, "Invalid TypeID for HandleServerGetPlayerCurrentEnvironment (" & iEnvirTypeID & "): " & mlClientPlayer(lIndex))
            Return
        End If
        If oPrimary Is Nothing = False Then
            Dim bFound As Boolean = False
            With oPrimary
                If .uListeners Is Nothing = False Then
                    For Y As Int32 = 0 To .uListeners.GetUpperBound(0)
                        If .uListeners(Y).lConnectionType = ConnectionType.eClientConnection Then
                            bFound = True
                            lPos = 0
                            Dim yResp(48) As Byte
                            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerCurrentEnvironment).CopyTo(yResp, lPos) : lPos += 2
                            System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4
                            System.BitConverter.GetBytes(lEnvirID).CopyTo(yResp, lPos) : lPos += 4
                            System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yResp, lPos) : lPos += 2
                            System.BitConverter.GetBytes(lGalaxyID).CopyTo(yResp, lPos) : lPos += 4
                            System.BitConverter.GetBytes(lSystemID).CopyTo(yResp, lPos) : lPos += 4
                            System.BitConverter.GetBytes(lPlanetID).CopyTo(yResp, lPos) : lPos += 4
                            yResp(lPos) = yInDomain : lPos += 1


                            If .uListeners(Y).sIPAddress.ToUpper = "EXTERNALIPADDY" Then
                                StringToBytes(gsExternalAddress).CopyTo(yResp, lPos) : lPos += 20
                            Else
                                StringToBytes(.uListeners(Y).sIPAddress).CopyTo(yResp, lPos) : lPos += 20
                            End If
                            System.BitConverter.GetBytes(.uListeners(Y).lPortNumber).CopyTo(yResp, lPos) : lPos += 4

                            moServers(lIndex).SendData(yResp)
                            Exit For
                        End If
                    Next Y
                End If
            End With
            If bFound = False Then
                LogEvent(LogEventType.CriticalError, "HandleServerGetPlayerCurrentEnvironment (client): Could not find Primary Listener")
            End If
        Else
            LogEvent(LogEventType.PossibleCheat, "Invalid GUID for PrimaryConnData (" & lEnvirID & ", " & iEnvirTypeID & "): " & mlClientPlayer(lIndex))
        End If

        ''k, now append that data
        'System.BitConverter.GetBytes(lGalaxyID).CopyTo(yResp, 12)
        'System.BitConverter.GetBytes(lSystemID).CopyTo(yResp, 16)
        'System.BitConverter.GetBytes(lPlanetID).CopyTo(yResp, 20)

        'If bInMyDomain = False Then
        '    yResp(24) = 0
        '    moOperator.SendData(yResp)
        '    Return Nothing
        'Else : yResp(24) = 1
        'End If
    End Sub
End Class
