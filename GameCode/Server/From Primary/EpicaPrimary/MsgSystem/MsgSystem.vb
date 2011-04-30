Option Strict On
Public Enum elPlayerDetailsType As Int32
	ePlayerObject = 0		'first now
	eMineralProperties = 1
	ePlayerMinerals = 2
	eModelDefs = 3
	eMySpecials = 4
	eMyAlloys = 5
	eMyArmor = 6
	eMyEngines = 7
	eMyHull = 8
	eMyRadar = 9
	eMyShield = 10
	eMyWeapon = 11
	eMyPrototype = 12
	eGlobalAlloys = 13
	eGlobalArmor = 14
	eGlobalEngines = 15
	eGlobalHull = 16
	eGlobalRadar = 17
	eGlobalShield = 18
	eGlobalWeapon = 19
	eGlobalPrototype = 20
	eUnitGroup = 21
	ePlayerRel = 22
	eUnitDefs = 23
	eFacilityDefs = 24
	eEmails = 25
	eTrades = 26
	ePlayerIntel = 27
	eSpecialAttrs = 28
	ePlayerWormholes = 29
	eAgents = 30
	ePlayerMissions = 31
	eAliasConfigs = 32		'TODO: not needed at login, may be removed. If so, add a request for them in frmAliasing.vb

	eLastPlayerDetailsType	'always the last in this enum!!!
End Enum

Public Class MsgSystem
    Private Const ml_AVAILABLE_RESOURCES_REQUEST_TIME As Int32 = 300      'in cycles...

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
	End Enum
	Public Enum ClientStatusFlags As Int32
		eRequestedGalaxyAndSystems = 1
        eRequestedPlayerDetails = 2
        ePlayerLoggedIn = 4
	End Enum

	Private mbAcceptingClients As Boolean = False
    Private mbAcceptingDomains As Boolean = False
    Private mbAcceptingPathfinding As Boolean = False

	Private WithEvents moOperator As NetSock				'the operator server I am connected to
	Private WithEvents moEmailSrvr As NetSock

	Public bDebug As Boolean = False

	Public moMonitor As MsgMonitor = New MsgMonitor()
	Public Const mb_MONITOR_MSG_ACTIVITY As Boolean = True

	Public Sub ForceDisconnectAll()
		Dim X As Int32

		On Error Resume Next

		For X = 0 To mlClientUB
			If moClients(X) Is Nothing = False Then
				moClients(X).Disconnect()
				Dim oPlayer As Player = GetEpicaPlayer(mlClientPlayer(X))
                If oPlayer Is Nothing = False AndAlso (mlClientStatusFlags(X) And ClientStatusFlags.ePlayerLoggedIn) <> 0 Then oPlayer.TotalPlayTime += CInt(DateDiff(DateInterval.Second, oPlayer.LastLogin, Now))
                If oPlayer Is Nothing = False Then oPlayer.oSocket = Nothing
				moClients(X) = Nothing
			End If
		Next X

		For X = 0 To mlDomainUB
			If moDomains(X) Is Nothing = False Then
				moDomains(X).Disconnect()
				moDomains(X) = Nothing
			End If
		Next X

		'Notify the PF server to disconnect
		Dim yMsg(1) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eServerShutdown).CopyTo(yMsg, 0)
		moPathfinding.SendData(yMsg)
		moPathfinding.Disconnect()
		moPathfinding = Nothing
	End Sub

	Public Sub BroadcastServerShutdownsAndDisconnectClients()
		Dim yMsg(1) As Byte
		Dim X As Int32

		System.BitConverter.GetBytes(GlobalMessageCode.eServerShutdown).CopyTo(yMsg, 0)
		Dim yTemp(yMsg.Length + 1) As Byte
		System.BitConverter.GetBytes(CShort(yMsg.Length)).CopyTo(yTemp, 0)
		yMsg.CopyTo(yTemp, 2)

		For X = 0 To mlDomainUB
			moDomains(X).SendLenAppendedData(yTemp)
		Next X

		For X = 0 To mlClientUB
			If moClients(X) Is Nothing = False Then
				moClients(X).SendLenAppendedData(yTemp)
			End If
		Next X
	End Sub

#Region "  Email Server Connection Handling  "
	Private mbConnectingToEmailSrvr As Boolean = False
    Private mbConnectedToEmailSrvr As Boolean = False
    Private mbConectionToEmailSrvrLost As Boolean = False
    Private mcolEmailFailQueue As Collection = Nothing
    Public swEmailSrvrReconnect As Stopwatch

	Private Sub moEmailSrvr_onConnect(ByVal Index As Integer) Handles moEmailSrvr.onConnect
		mbConnectedToEmailSrvr = True
		mbConnectingToEmailSrvr = False
		'Send only the players that I have domain over (when it is implemented)
		For X As Int32 = 0 To glPlayerUB
            If glPlayerIdx(X) <> -1 Then 'AndAlso goPlayer(X).InMyDomain = True Then
                'Email Srvr expects a different Add Object than the other applications
                CreateAndSendEmailSrvrPlayerMsg(goPlayer(X))
                'Dim yMsg(329) As Byte       '325
                'Dim lPos As Int32 = 0

                'If MsgSystem.mb_MONITOR_MSG_ACTIVITY = True Then
                '                goMsgSys.moMonitor.AddOutMsg(GlobalMessageCode.eAddObjectCommand, MsgMonitor.eMM_AppType.EmailServer, 328, 0)   '324
                'End If

                '            System.BitConverter.GetBytes(328S).CopyTo(yMsg, lPos) : lPos += 2       '324
                'System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, lPos) : lPos += 2
                'With goPlayer(X)
                '	.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                '	.PlayerName.CopyTo(yMsg, lPos) : lPos += 20
                '	.EmpireName.CopyTo(yMsg, lPos) : lPos += 20
                '	.RaceName.CopyTo(yMsg, lPos) : lPos += 20
                '	.ExternalEmailAddress.CopyTo(yMsg, lPos) : lPos += 255
                '                yMsg(lPos) = .yGender : lPos += 1
                '                System.BitConverter.GetBytes(.lGuildID).CopyTo(yMsg, lPos) : lPos += 4
                'End With

                'moEmailSrvr.SendLenAppendedData(yMsg)
            End If
        Next X

        If mcolEmailFailQueue Is Nothing = False Then
            For Each yAry() As Byte In mcolEmailFailQueue
                moEmailSrvr.SendData(yAry)
            Next
            mcolEmailFailQueue.Clear()
        End If
        mcolEmailFailQueue = Nothing
	End Sub

	Private Sub moEmailSrvr_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moEmailSrvr.onConnectionRequest
		'do nothing, should never happen
	End Sub

	Private Sub moEmailSrvr_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moEmailSrvr.onDataArrival
		Dim iMsgID As Int16
		iMsgID = System.BitConverter.ToInt16(Data, 0)
		If mb_MONITOR_MSG_ACTIVITY = True Then moMonitor.AddInMsg(iMsgID, MsgMonitor.eMM_AppType.EmailServer, Data.Length, Index)
        Select Case iMsgID
            Case GlobalMessageCode.eChatMessage
                HandleChatMessageResponse(Data)
            Case GlobalMessageCode.eSetPlayerRel
                HandleSetPlayerRel(Data, -1)
            Case GlobalMessageCode.eColonyLowResources
                HandleMailSrvr_ColonyLowResources(Data)
            Case GlobalMessageCode.ePlayerAlert
                HandleMailSrvr_PlayerAlert(Data)
            Case GlobalMessageCode.eRebuildAISetting
                HandleMailSrvr_RebuildAISetting(Data)
            Case GlobalMessageCode.eRequestChannelDetails
                HandleChannelDetailsResponse(Data)
            Case GlobalMessageCode.eRequestChannelList
                HandleChannelListResponse(Data)
            Case GlobalMessageCode.eSetMineralBid
                HandleEmailSetMineralBid(Data)
            Case GlobalMessageCode.eSendOutMailMsg
                HandleEmailSendMail(Data)
        End Select
	End Sub

	Private Sub moEmailSrvr_onDisconnect(ByVal Index As Integer) Handles moEmailSrvr.onDisconnect
		LogEvent(LogEventType.Informational, "Email Server disconnected.")
	End Sub

	Private Sub moEmailSrvr_onError(ByVal Index As Integer, ByVal Description As String) Handles moEmailSrvr.onError
        LogEvent(LogEventType.Informational, "Email Server Error: " & Description)
        If Description.ToUpper.Contains("AN EXISTING CONNECTION WAS FORCIBLY CLOSED") = True Then
            'Lost connection with the email server
            mbConectionToEmailSrvrLost = True
            moEmailSrvr.Disconnect()
        End If
	End Sub

	Private Sub moEmailSrvr_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moEmailSrvr.onSendComplete
		'
	End Sub

	Public Function ConnectToEmailSrvr() As Boolean
		Dim oINI As InitFile = New InitFile()
		Dim bRes As Boolean = False
		Dim mlTimeout As Int32

		Try
			mlTimeout = CInt(Val(oINI.GetString("CONNECTION", "ConnectTimeout", "10000")))
			If glEmailSrvrPort = 0 OrElse gsEmailSrvrIP = "" Then
				LogEvent(LogEventType.Warning, "ConnectToEmailSrvr without valid Port and IP address attempting to load from INI.")
				glEmailSrvrPort = CInt(Val(oINI.GetString("CONNECTION", "EmailSrvrPort", "")))
				gsEmailSrvrIP = oINI.GetString("CONNECTION", "EmailSrvrIP", "")
				If glEmailSrvrPort = 0 OrElse gsEmailSrvrIP = "" Then Err.Raise(-1, "EstablishConnection", "Unable to connect to server: could not find connection credentials.")
			End If
			mbConnectingToEmailSrvr = True
			moEmailSrvr = New NetSock()
			moEmailSrvr.Connect(gsEmailSrvrIP, glEmailSrvrPort)
			Dim sw As Stopwatch = Stopwatch.StartNew
			While mbConnectedToEmailSrvr = False AndAlso (sw.ElapsedMilliseconds < mlTimeout) AndAlso mbConnectingToEmailSrvr = True
				'NOTE: This causes very nasty context errors if left in
				Threading.Thread.Sleep(10)
			End While
			sw.Stop()
			sw = Nothing
			mbConnectingToEmailSrvr = False
			If mbConnectedToEmailSrvr = False Then
				Err.Raise(-1, "EstablishConnection", "Unable to connect to Email server: Connection timed out!")
			Else
				oINI.WriteString("CONNECTION", "ConnectTimeout", mlTimeout.ToString)
				oINI.WriteString("CONNECTION", "EmailSrvrPort", glEmailSrvrPort.ToString)
				oINI.WriteString("CONNECTION", "EmailSrvrIP", gsEmailSrvrIP)
			End If
		Catch
			mbConnectedToEmailSrvr = False
		Finally
			bRes = mbConnectedToEmailSrvr
			oINI = Nothing
		End Try
		Return bRes
	End Function

    Public Sub SendToEmailSrvr(ByRef yData() As Byte)
        If mbConnectedToEmailSrvr = True AndAlso moEmailSrvr Is Nothing = False Then
            moEmailSrvr.SendData(yData)
        Else
            If mcolEmailFailQueue Is Nothing Then
                mbConectionToEmailSrvrLost = True
                mcolEmailFailQueue = New Collection
                swEmailSrvrReconnect = Stopwatch.StartNew
            End If
            mcolEmailFailQueue.Add(yData)
        End If 
    End Sub
#Region "  Email Srvr Inbound Msgs  "
    Private Sub HandleEmailSendMail(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPC_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lTo As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim sValue As String = GetStringFromBytes(yData, lPos, lLen) : lPos += lLen

        Dim oTo As Player = GetEpicaPlayer(lTo)
        If oTo Is Nothing = False Then
            Dim oPC As PlayerComm = oTo.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sValue, "RE:", lPlayerID, GetDateAsNumber(Now), False, oTo.sPlayerNameProper, Nothing)
            If oPC Is Nothing = False Then oTo.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
        End If
    End Sub

    Private Sub HandleEmailSetMineralBid(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lFacID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lBid As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return
        Dim oFac As Facility = GetEpicaFacility(lFacID)
        If oFac Is Nothing Then Return

        If oFac.oMiningBid Is Nothing = False Then
            oFac.oMiningBid.AddBid(lBid, oFac.oMiningBid.oMineralCache.Quantity, oPlayer, True)
        End If
    End Sub
    Private Sub HandleChannelDetailsResponse(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            If oPlayer.oSocket Is Nothing = False Then oPlayer.oSocket.SendData(yData)
        End If
    End Sub
    Private Sub HandleChannelListResponse(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            If oPlayer.oSocket Is Nothing = False Then oPlayer.oSocket.SendData(yData)
        End If
    End Sub
    Private Sub HandleChatMessageResponse(ByVal yData() As Byte)
        'should contain...
        'MsgCode (2)
        'FromPlayer (4)
        'Type (1)
        'MsgLen (4)
        'MsgBody ()
        'RecipientCnt (4)
        'Recipients ()
        'we need to take this message and send it to the recipients as...
        'MsgCode (2)
        'FromPlayer (4)
        'Type (1)
        'MsgLen (4)
        'MsgBody ()

        Dim lPos As Int32 = 7       'for code, player, type
        Dim lMsgLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lPos += lMsgLen

        Dim yForward(lPos - 1) As Byte
        Array.Copy(yData, 0, yForward, 0, lPos)

        'Now, get our recipient count
        Dim lRecipients As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'Now, loop through our recipients and send them the message...
        For X As Int32 = 0 To lRecipients - 1
            Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            If lPlayerID = -1 Then
                For Y As Int32 = 0 To glPlayerUB
                    If glPlayerIdx(Y) > -1 Then
                        Dim oTmpPlayer As Player = goPlayer(Y)
                        If oTmpPlayer Is Nothing = False AndAlso oTmpPlayer.lConnectedPrimaryID > -1 Then
                            oTmpPlayer.CrossPrimarySafeSendMsg(yForward)
                        End If
                    End If
                Next Y
                Return
            End If

            Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
            If oPlayer Is Nothing = False AndAlso oPlayer.lConnectedPrimaryID > -1 Then oPlayer.CrossPrimarySafeSendMsg(yForward)
        Next X
    End Sub
    Private Sub HandleMailSrvr_ColonyLowResources(ByRef yData() As Byte)
        'MsgCode (2), PlayerID (4), ColonyID (4), ItemID (4), ItemTypeID (2), lQty (4)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lColonyID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lItemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lQty As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        Dim lIdx As Int32 = oPlayer.GetColonyFromColonyID(lColonyID)
        If lIdx = -1 OrElse glColonyIdx(lIdx) <> lColonyID Then Return

        Dim oColony As Colony = goColony(lIdx)
        If oColony Is Nothing Then Return

        'typeid can be an alloytech (build mineral from refinery)
        ' can be a component (build component from production capable)
        ' mineral (gather, find a mining facility that mines that cache, launch trucks if available and turn auto-launch on)
        If iTypeID = ObjectType.eAlloyTech Then
            Dim lFacIdx As Int32 = -1
            For X As Int32 = 0 To oColony.ChildrenUB
                If oColony.lChildrenIdx(X) <> -1 AndAlso oColony.oChildren(X).Active = True AndAlso oColony.oChildren(X).yProductionType = ProductionType.eRefining Then
                    'Ok, is this facility building?
                    If oColony.oChildren(X).bProducing = True Then
                        lFacIdx = X
                    Else
                        'No, use it
                        lFacIdx = X
                        Exit For
                    End If
                End If
            Next X
            If lFacIdx <> -1 Then
                If oColony.oChildren(lFacIdx).AddProduction(lItemID, ObjectType.eMineral, 254, lQty, 0) = True Then
                    For Y As Int32 = 0 To glFacilityUB
                        If glFacilityIdx(Y) = oColony.oChildren(lFacIdx).ObjectID Then
                            AddEntityProducing(Y, ObjectType.eFacility, oColony.oChildren(lFacIdx).ObjectID)
                            Exit For
                        End If
                    Next Y
                End If
            End If
        ElseIf iTypeID = ObjectType.eMineral OrElse lQty = -1 Then
            'Gather
            For X As Int32 = 0 To oColony.ChildrenUB
                If oColony.lChildrenIdx(X) <> -1 AndAlso oColony.oChildren(X).Active = True AndAlso oColony.oChildren(X).yProductionType = ProductionType.eMining Then
                    'ok, is this facility mining our mineral?
                    If oColony.oChildren(X).bMining = True AndAlso oColony.oChildren(X).lCacheIndex <> -1 AndAlso glMineralCacheIdx(oColony.oChildren(X).lCacheIndex) <> -1 AndAlso (oColony.oChildren(X).CurrentStatus And elUnitStatus.eHangarOperational) <> 0 Then
                        Dim oCache As MineralCache = goMineralCache(oColony.oChildren(X).lCacheIndex)
                        If oCache Is Nothing = False AndAlso oCache.ObjectID = oColony.oChildren(X).lCacheID Then
                            If oCache.oMineral.ObjectID = lItemID Then
                                If oColony.oChildren(X).AutoLaunch = False Then
                                    oColony.oChildren(X).AutoLaunch = True
                                    For Y As Int32 = 0 To oColony.oChildren(X).lHangarUB
                                        If oColony.oChildren(X).lHangarIdx(Y) <> -1 AndAlso oColony.oChildren(X).oHangarContents(Y).ObjTypeID = ObjectType.eUnit Then
                                            With CType(oColony.oChildren(X).oHangarContents(Y), Unit)
                                                AddToQueue(glCurrentCycle, EngineCode.QueueItemType.eUndockAndReturnToRefinery_QIT, .ObjectID, .ObjTypeID, oColony.oChildren(X).ObjectID, oColony.oChildren(X).ObjTypeID, 0, 0, 0, 0)
                                            End With
                                        End If
                                    Next Y
                                End If
                            End If
                        Else
                            oColony.oChildren(X).lCacheIndex = -1
                            oColony.oChildren(X).lCacheID = -1
                        End If
                    End If
                End If
            Next X
        Else
            'Component
            Dim lFacIdx As Int32 = -1
            For X As Int32 = 0 To oColony.ChildrenUB
                If oColony.lChildrenIdx(X) <> -1 AndAlso oColony.oChildren(X).Active = True AndAlso (oColony.oChildren(X).yProductionType And ProductionType.eProduction) <> 0 Then
                    If oColony.oChildren(X).bProducing = True Then
                        lFacIdx = X
                    Else
                        'No, use it
                        lFacIdx = X
                        Exit For
                    End If
                End If
            Next X
            If lFacIdx <> -1 Then
                If oColony.oChildren(lFacIdx).AddProduction(lItemID, ObjectType.eMineral, 254, lQty, 0) = True Then
                    For Y As Int32 = 0 To glFacilityUB
                        If glFacilityIdx(Y) = oColony.oChildren(lFacIdx).ObjectID Then
                            AddEntityProducing(Y, ObjectType.eFacility, oColony.oChildren(lFacIdx).ObjectID)
                            Exit For
                        End If
                    Next Y
                End If
            End If
        End If


    End Sub
    Private Sub HandleMailSrvr_PlayerAlert(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Try
            Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim yType As Byte = yData(lPos) : lPos += 1
            Dim lEnemyID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim sSpecific As String = GetStringFromBytes(yData, lPos, 20).ToUpper : lPos += 20

            'ok, find our parent environment's domain
            Dim oSocket As NetSock = Nothing
            If iEnvirTypeID = ObjectType.ePlanet Then
                Dim oPlanet As Planet = GetEpicaPlanet(lEnvirID)
                If oPlanet Is Nothing = False AndAlso oPlanet.oDomain Is Nothing = False Then oSocket = oPlanet.oDomain.DomainSocket
            ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
                Dim oSystem As SolarSystem = GetEpicaSystem(lEnvirID)
                If oSystem Is Nothing = False AndAlso oSystem.oDomain Is Nothing = False Then oSocket = oSystem.oDomain.DomainSocket
            End If
            If oSocket Is Nothing AndAlso yType <> 255 Then
                LogEvent(LogEventType.Warning, "HandleMailSrvr_PlayerAlert: Unable to get domain socket: " & lEnvirID & ", " & iEnvirTypeID & ".")
                Return
            End If

            If yType = 0 Then   'Launch to Attack - everything is launched as possible
                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                If oPlayer Is Nothing = False Then
                    Dim lColonyIdx As Int32 = oPlayer.GetColonyFromParent(lEnvirID, iEnvirTypeID)
                    If lColonyIdx > -1 Then
                        Dim oColony As Colony = goColony(lColonyIdx)
                        If oColony Is Nothing = False Then

                            For X As Int32 = 0 To oColony.ChildrenUB
                                Dim oFac As Facility = oColony.oChildren(X)
                                If oFac Is Nothing = False Then
                                    If oFac.Hangar_Cap <> oFac.EntityDef.Hangar_Cap Then
                                        AddToQueue(glCurrentCycle + X, QueueItemType.eLaunchToAttack, oFac.ObjectID, oFac.ObjTypeID, lEnvirID, iEnvirTypeID, lLocX, lLocZ, -1, -1)
                                    End If
                                End If
                            Next X

                        End If
                    End If
                End If
            ElseIf yType = 255 Then
                'initiate the full invulnerability field
                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                If oPlayer Is Nothing = False Then
                    Select Case oPlayer.yIronCurtainState
                        Case eIronCurtainState.IronCurtainIsDown
                            oPlayer.yIronCurtainState = eIronCurtainState.RaisingIronCurtainOnEverything
                            AddToQueue(glCurrentCycle + 108000, QueueItemType.eIronCurtainRaise, oPlayer.ObjectID, -1, -1, -1, 0, 0, 0, 0)
                        Case eIronCurtainState.IronCurtainIsUpOnSelectedPlanet
                            oPlayer.bInFullLockDown = False
                            oPlayer.InitiateFullLockdown()
                        Case eIronCurtainState.RaisingIronCurtainOnSelectedPlanet
                            oPlayer.yIronCurtainState = eIronCurtainState.RaisingIronCurtainOnEverything
                    End Select

                    oPlayer.lStatusFlags = oPlayer.lStatusFlags Or elPlayerStatusFlag.FullInvulnerabilityRaised
                End If
                Return
            Else
                'Now, get our data...
                Dim yForward(24) As Byte
                Dim lDestPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eSetCounterAttack).CopyTo(yForward, lDestPos) : lDestPos += 2
                Array.Copy(yData, 2, yForward, lDestPos, 23) : lDestPos += 23

                'All but specific string is sent... now, get our unit list, if type is battlegroup (1)
                If yType = 1 Then       'battlegroup
                    Dim lCnt As Int32 = 0
                    Dim oUnitGroup As UnitGroup = Nothing
                    For X As Int32 = 0 To glUnitGroupUB
                        If glUnitGroupIdx(X) <> -1 AndAlso goUnitGroup(X).oOwner.ObjectID = lPlayerID Then
                            If sSpecific.Contains(BytesToString(goUnitGroup(X).UnitGroupName).ToUpper) = True Then
                                'ok, found it
                                oUnitGroup = goUnitGroup(X)
                                Exit For
                            End If
                        End If
                    Next X

                    If oUnitGroup Is Nothing Then
                        LogEvent(LogEventType.Warning, "HandleMailSrvr_PlayerAlert: Unable to find Unit group: " & sSpecific)
                        Return
                    End If

                    For X As Int32 = 0 To oUnitGroup.UnitUB
                        Dim lIdx As Int32 = oUnitGroup.GetUnitIdx(X)
                        If lIdx <> -1 AndAlso glUnitIdx(lIdx) <> -1 Then
                            lCnt += 1
                        End If
                    Next X

                    If lCnt <> 0 Then
                        ReDim Preserve yForward(28 + (lCnt * 4))
                        System.BitConverter.GetBytes(lCnt).CopyTo(yForward, lDestPos) : lDestPos += 4
                        For X As Int32 = 0 To oUnitGroup.UnitUB
                            Dim lIdx As Int32 = oUnitGroup.GetUnitIdx(X)
                            If lIdx <> -1 AndAlso glUnitIdx(lIdx) <> -1 Then
                                System.BitConverter.GetBytes(glUnitIdx(lIdx)).CopyTo(yForward, lDestPos) : lDestPos += 4
                            End If
                        Next X
                    Else : Return
                    End If
                Else
                    Dim lTmpUB As Int32 = -1
                    Dim lTempIdx(99) As Int32
                    For X As Int32 = 0 To glUnitUB
                        If glUnitIdx(X) <> -1 Then
                            Dim oUnit As Unit = goUnit(X)
                            If oUnit Is Nothing = False AndAlso oUnit.Owner.ObjectID = lPlayerID AndAlso oUnit.EntityDef.CombatRating > 0 Then
                                If oUnit.lRouteUB = -1 Then
                                    'Ok, part of the order
                                    lTmpUB += 1
                                    If lTmpUB > lTempIdx.GetUpperBound(0) Then ReDim Preserve lTempIdx(lTmpUB + 100)
                                    lTempIdx(lTmpUB) = X
                                End If
                            End If
                        End If
                    Next X

                    If yType = 3 Then lTmpUB \= 2 'half units

                    If lTmpUB <> -1 Then
                        ReDim Preserve yForward(28 + ((lTmpUB + 1) * 4))
                        System.BitConverter.GetBytes(lTmpUB + 1).CopyTo(yForward, lDestPos) : lDestPos += 4
                        For X As Int32 = 0 To lTmpUB
                            Dim lIdx As Int32 = lTempIdx(X)
                            If lIdx <> -1 AndAlso glUnitIdx(lIdx) <> -1 Then
                                System.BitConverter.GetBytes(glUnitIdx(lIdx)).CopyTo(yForward, lDestPos) : lDestPos += 4
                            End If
                        Next X
                    Else : Return
                    End If
                End If

                oSocket.SendData(yData)
            End If

        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "HandleMailSrvr_PlayerAlert: " & ex.Message)
        End Try
    End Sub
    Private Sub HandleMailSrvr_RebuildAISetting(ByRef yData() As Byte)
        Dim lColonyID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim oColony As Colony = GetEpicaColony(lColonyID)
        If oColony Is Nothing = False Then
            oColony.lRebuilderQueueIdx = -1
        End If
    End Sub
#End Region
#End Region

#Region "  Operator Events  "
    Private mbConnectingToOperator As Boolean = False
	Private mbConnectedToOperator As Boolean = False
	Private moOperatorSW As Stopwatch = Nothing
	Private mcolOperatorFailQueue As Collection = Nothing
	Private msReconnectOperatorIP As String
	Private mlReconnectOperatorPort As Int32
	Public Sub OperatorFailure()
		msReconnectOperatorIP = gsBackupOperatorIP
		mlReconnectOperatorPort = glBackupOperatorPort
		mbConnectedToOperator = False
		moOperatorSW = Nothing
	End Sub

	Public Function ConnectToOperator(ByVal sIP As String, ByVal lPort As Int32) As Boolean
		Dim mlTimeout As Int32 = 10000

		Try
			mbConnectingToOperator = True
			moOperator = New NetSock()
			moOperator.Connect(sIP, lPort)
			Dim sw As Stopwatch = Stopwatch.StartNew
			While mbConnectedToOperator = False AndAlso (sw.ElapsedMilliseconds < mlTimeout) AndAlso mbConnectingToOperator = True
				Application.DoEvents()
			End While
			sw.Stop()
			sw = Nothing
			mbConnectingToOperator = False
			If mbConnectedToOperator = False Then
				Err.Raise(-1, "EstablishConnection", "Unable to connect to OPERATOR: Connection timed out!")
			End If

			moOperatorSW = Stopwatch.StartNew()
		Catch
			mbConnectedToOperator = False
		End Try
		Return mbConnectedToOperator
	End Function

    Private Sub moOperator_onConnect(ByVal Index As Integer) Handles moOperator.onConnect
        mbConnectedToOperator = True
        mbConnectingToOperator = False

        'Indicate to the Operator who I am... needs to indicate my server type
        Dim oEP As Net.EndPoint = moOperator.GetLocalDetails()
        If oEP Is Nothing = False Then
            With CType(oEP, Net.IPEndPoint)
                Dim yMsg(101) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eRegisterDomain).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(glBoxOperatorID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(ConnectionType.ePrimaryServerApp).CopyTo(yMsg, lPos) : lPos += 4

                Dim oMyProcess As System.Diagnostics.Process = System.Diagnostics.Process.GetCurrentProcess()
                Dim lProcessID As Int32 = oMyProcess.Id
                System.BitConverter.GetBytes(lProcessID).CopyTo(yMsg, lPos) : lPos += 4
                oMyProcess = Nothing

                System.BitConverter.GetBytes(3I).CopyTo(yMsg, lPos) : lPos += 4     'indicates 1 connection specifics

                'Get our Port number data
                Dim oIni As New InitFile()
                Dim lRegionPort As Int32 = CInt(Val(oIni.GetString("SETTINGS", "DomainListenerPort", "0")))
                Dim lClientPort As Int32 = CInt(Val(oIni.GetString("SETTINGS", "ClientListenerPort", "0")))
                Dim lPathfindingPort As Int32 = CInt(Val(oIni.GetString("SETTINGS", "PathfindingListenerPort", "0")))
                If lRegionPort = 0 Then Err.Raise(-1, "moOperator.onConnect", "moOperator.onConnect: Unable to determine Region Listen Port from INI")
                If lClientPort = 0 Then Err.Raise(-1, "moOperator.onConnect", "moOperator.onConnect: Unable to determine Client Listen Port from INI")
                If lPathfindingPort = 0 Then Err.Raise(-1, "moOperator.onConnect", "moOperator.onConnect: Unable to determine Pathfinding Listen Port from INI")
                oIni = Nothing

                'Now, we'll indicate our client listener... so use "EXTERNALIPADDY" key word
                StringToBytes(gsExternalIP).CopyTo(yMsg, lPos) : lPos += 20
                System.BitConverter.GetBytes(lClientPort).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(ConnectionType.eClientConnection).CopyTo(yMsg, lPos) : lPos += 4

                'Next, we'll indicate our region listener... so use our local ip address
                StringToBytes(Mid$(.Address.ToString(), 1, 20)).CopyTo(yMsg, lPos) : lPos += 20
                System.BitConverter.GetBytes(lRegionPort).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(ConnectionType.eRegionServerApp).CopyTo(yMsg, lPos) : lPos += 4

                'Then, the Pathfinding Server Listener... use our local IP address again
                StringToBytes(Mid$(.Address.ToString(), 1, 20)).CopyTo(yMsg, lPos) : lPos += 20
                System.BitConverter.GetBytes(lPathfindingPort).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(ConnectionType.ePathfindingServerApp).CopyTo(yMsg, lPos) : lPos += 4

                moOperator.SendData(yMsg)
            End With
		End If

		If mcolOperatorFailQueue Is Nothing = False Then
			For Each yAry() As Byte In mcolOperatorFailQueue
				moOperator.SendData(yAry)
			Next
			mcolOperatorFailQueue.Clear()
		End If
		mcolOperatorFailQueue = Nothing
    End Sub

    Private Sub moOperator_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moOperator.onConnectionRequest
        'should never happen
    End Sub

    Private Sub moOperator_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moOperator.onDataArrival
        Dim iMsgID As Int16
        iMsgID = System.BitConverter.ToInt16(Data, 0)
        If mb_MONITOR_MSG_ACTIVITY = True Then moMonitor.AddInMsg(iMsgID, MsgMonitor.eMM_AppType.OperatorServer, Data.Length, Index)
        Select Case iMsgID
            Case GlobalMessageCode.eCrossPrimaryBudgetAdjust
                HandleCrossPrimaryBudgetAdjust(Data)
            Case GlobalMessageCode.eOperatorRequestPassThru
                HandleOperatorRequestPassThru(Data)
            Case GlobalMessageCode.eRegisterDomain
                HandleOperatorDomainRegister(Data)
            Case GlobalMessageCode.eEmailSettings
                HandleEmailServerData(Data)
            Case GlobalMessageCode.ePlayerConnectedPrimary
                HandlePlayerConnectedPrimary(Data)
            Case GlobalMessageCode.eRequestObject
                If moOperatorSW Is Nothing Then
                    moOperatorSW = Stopwatch.StartNew
                Else
                    moOperatorSW.Reset()
                    moOperatorSW.Start()
                End If
                moOperator.SendData(Data)
            Case GlobalMessageCode.eDomainServerReady
                HandleServersReady()
            Case GlobalMessageCode.eServerShutdown
                gfrmDisplayForm.btnShutdown_Click(Nothing, Nothing)
            Case GlobalMessageCode.ePlayerInitialSetup
                HandlePlayerInitialSetup(Data, -1)
            Case GlobalMessageCode.ePlayerDomainAssignment
                HandlePlayerDomainAssignment(Data)
            Case GlobalMessageCode.eAddObjectCommand
                HandleOperatorAddObject(Data)
            Case GlobalMessageCode.ePlayerAliasConfig
                HandleOperatorAliasConfig(Data)
            Case GlobalMessageCode.eRemovePlayerRel
                HandleRemovePlayerRel(Data)
            Case GlobalMessageCode.eBackupOperatorSyncMsg
                HandleOperatorRebound(Data)
            Case GlobalMessageCode.eEnvironmentDomain
                HandleSpawnSystemRegion(Data)
            Case GlobalMessageCode.eUpdatePlayerDetails
                HandleOperatorUpdatePlayerDetails(Data)
            Case GlobalMessageCode.ePlayerIsDead
                HandlePlayerDied(Data)
            Case GlobalMessageCode.eSetPlayerSpecialAttribute
                HandleOperatorSetPlayerSpecialAttribute(Data)
            Case GlobalMessageCode.eForcePrimarySync
                HandleForcePrimarySync(Data)
            Case GlobalMessageCode.eAddSenateProposal
                Dim lPlayerID As Int32 = System.BitConverter.ToInt32(Data, 2)
                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                If oPlayer Is Nothing = False Then
                    If oPlayer.lConnectedPrimaryID > -1 Then oPlayer.CrossPrimarySafeSendMsg(Data)
                End If
            Case GlobalMessageCode.eGetSenateObjectDetails
                Dim lPlayerID As Int32 = System.BitConverter.ToInt32(Data, 2)
                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                If oPlayer Is Nothing = False Then
                    If oPlayer.lConnectedPrimaryID > -1 Then oPlayer.CrossPrimarySafeSendMsg(Data)
                End If
            Case GlobalMessageCode.eForwardToPlayerAtPrimary
                HandleForwardToPlayerAtPrimary(Data)
            Case GlobalMessageCode.eEntityChangingPrimary
                HandleEntityChangingPrimary(Data)
            Case GlobalMessageCode.ePrimaryLoadSharedPlayerData
                HandlePrimaryLoadSharedPlayerData(Data)
            Case GlobalMessageCode.eAddFormation
                HandleOperatorAddFormation(Data)
            Case GlobalMessageCode.eRequestPlayerBudget
                HandleOperatorRequestPlayerBudget(Data)
            Case GlobalMessageCode.eRequestGlobalPlayerScores
                HandleRequestGlobalPlayerScores(Data)
            Case GlobalMessageCode.ePlayerCurrentEnvironment
                HandleOperatorPlayerCurrEnvir(Data)
            Case GlobalMessageCode.eChangingEnvironment
                Dim lPlayerID As Int32 = System.BitConverter.ToInt32(Data, 2)
                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                If oPlayer Is Nothing = False Then
                    oPlayer.lLastViewedEnvir = System.BitConverter.ToInt32(Data, 6)
                    oPlayer.iLastViewedEnvirType = System.BitConverter.ToInt16(Data, 10)
                End If
            Case GlobalMessageCode.eUpdatePlayerCredits
                Dim lPlayerID As Int32 = System.BitConverter.ToInt32(Data, 2)
                Dim blCredits As Int64 = System.BitConverter.ToInt64(Data, 6)
                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                If oPlayer Is Nothing = False AndAlso oPlayer.InMyDomain = True Then oPlayer.blCredits += blCredits
            Case GlobalMessageCode.eCheckPrimaryReady
                HandleCheckPrimaryReady(Data)
            Case GlobalMessageCode.eChatMessage
                HandleOperatorChatMessage(Data)
            Case GlobalMessageCode.eSenateStatusReport
                Dim lPlayerID As Int32 = System.BitConverter.ToInt32(Data, 2)
                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                If oPlayer Is Nothing = False Then oPlayer.SendPlayerMessage(Data, False, 0)
        End Select
    End Sub

    Private Sub moOperator_onDisconnect(ByVal Index As Integer) Handles moOperator.onDisconnect
        LogEvent(LogEventType.Informational, "Operator Disconnected")
    End Sub

    Private Sub moOperator_onError(ByVal Index As Integer, ByVal Description As String) Handles moOperator.onError
        LogEvent(LogEventType.Informational, "moOperator Error: " & Description)
    End Sub

    Private Sub moOperator_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moOperator.onSendComplete
        'do nothing
    End Sub

    Private Sub HandleCheckPrimaryReady(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lUB As Int32 = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
        Dim lIDs(lUB) As Int32
        For X As Int32 = 0 To lUB
            lIDs(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Next X

        Dim bDone As Boolean = False
        Dim lCnt As Int32 = 0
        While bDone = False
            bDone = True
            lCnt += 1
            For X As Int32 = 0 To lUB
                Dim oSys As SolarSystem = GetEpicaSystem(lIDs(X))
                If oSys Is Nothing = True Then
                    bDone = False
                    Exit For
                End If
            Next X
            If bDone = False Then
                Threading.Thread.Sleep(1000)
            End If
            If lCnt > 120 Then
                LogEvent(LogEventType.CriticalError, "CheckPrimaryReady is taking a very long time! Returning true!")
                bDone = True
            End If
        End While

        moOperator.SendData(yData)
    End Sub

    Public Sub SendRequestGlobalPlayerScores(ByVal lPlayerID As Int32, ByVal yTypeID As Byte, ByVal lForPlayerID As Int32)
        Dim yMsg(14) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestGlobalPlayerScores).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lPlayerID).CopyTo(yMsg, lPos) : lPos += 4
        yMsg(lPos) = yTypeID : lPos += 1
        System.BitConverter.GetBytes(lForPlayerID).CopyTo(yMsg, lPos) : lPos += 4
        lPos += 4
        SendMsgToOperator(yMsg)
    End Sub
    Private Sub HandleRequestGlobalPlayerScores(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yTypeID As Byte = yData(lPos) : lPos += 1
        Dim lForPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lRequestIdx As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        If lRequestIdx = -1 Then
            'responding to me, ok, determine who we are referring to...
            With oPlayer
                .lLastGlobalRequestTurnIn = glCurrentCycle
                .lLGTechScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lLGWealthScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lLGPopulationScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lLGProductionScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lLGMilitaryScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lLGDiplomacyScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lLGTotalScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            End With

            If lForPlayerID <> lPlayerID Then
                Dim oForPlayer As Player = GetEpicaPlayer(lForPlayerID)
                Dim lUpdate As Int32 = GetDateAsNumber(Now)
                If oForPlayer Is Nothing = False Then
                    Dim oPI As PlayerIntel = oForPlayer.GetOrAddPlayerIntel(lPlayerID, False)
                    If oPI Is Nothing = False Then
                        Select Case yTypeID
                            Case 1
                                oPI.TechnologyScore = oPlayer.lLGTechScore
                                oPI.TechnologyUpdate = lUpdate
                            Case 2
                                oPI.WealthScore = oPlayer.lLGWealthScore
                                oPI.WealthUpdate = lUpdate
                            Case 3
                                oPI.PopulationScore = oPlayer.lLGPopulationScore
                                oPI.PopulationUpdate = lUpdate
                            Case 4
                                oPI.ProductionScore = oPlayer.lLGProductionScore
                                oPI.ProductionUpdate = lUpdate
                            Case 5
                                oPI.MilitaryScore = oPlayer.lLGMilitaryScore \ 50
                                If oPI.MilitaryScore * 100L < Int32.MaxValue Then oPI.MilitaryScore *= 100
                                oPI.MilitaryUpdate = lUpdate
                            Case 6
                                oPI.DiplomacyScore = oPlayer.lLGDiplomacyScore
                                oPI.DiplomacyUpdate = lUpdate
                            Case 64
                                oPI.TechnologyScore = oPlayer.lLGTechScore
                                oPI.TechnologyUpdate = lUpdate
                                oPI.WealthScore = oPlayer.lLGWealthScore
                                oPI.WealthUpdate = lUpdate
                                oPI.PopulationScore = oPlayer.lLGPopulationScore
                                oPI.PopulationUpdate = lUpdate
                                oPI.ProductionScore = oPlayer.lLGProductionScore
                                oPI.ProductionUpdate = lUpdate
                                oPI.MilitaryScore = oPlayer.lLGMilitaryScore \ 50
                                If oPI.MilitaryScore * 100L < Int32.MaxValue Then oPI.MilitaryScore *= 100
                                oPI.MilitaryUpdate = lUpdate
                                oPI.DiplomacyScore = oPlayer.lLGDiplomacyScore
                                oPI.DiplomacyUpdate = lUpdate
                        End Select

                        oForPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPI, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                    End If
                End If
            End If

        Else
            'requesting from me
            Dim yResp(42) As Byte
            lPos = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestGlobalPlayerScores).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yResp, lPos) : lPos += 4
            yResp(lPos) = CByte(yTypeID Or 128) : lPos += 1
            System.BitConverter.GetBytes(lForPlayerID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(lRequestIdx).CopyTo(yResp, lPos) : lPos += 4

            With oPlayer
                System.BitConverter.GetBytes(.TechnologyScore).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(.WealthScore).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(.PopulationScore).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(.ProductionScore).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lMilitaryScore \ 100).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(.DiplomacyScore).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(.TotalScore).CopyTo(yResp, lPos) : lPos += 4
            End With

            moOperator.SendData(yResp)

        End If
    End Sub

    Private Sub HandleOperatorChatMessage(ByVal yData() As Byte)
        Dim lPos As Int32 = 2

        Dim lFromPlayer As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yType As Byte = yData(lPos) : lPos += 1
        Dim lLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim sMsg As String = GetStringFromBytes(yData, lPos, lLen) : lPos += lLen

        Dim oPlayer As Player = GetEpicaPlayer(lFromPlayer)
        If oPlayer Is Nothing = True Then Return

        If sMsg.StartsWith("/csrloggedin") = True Then
            oPlayer.SendPlayerMessage(CreateChatMsg(-1, "You are logged in as the CSR", ChatMessageType.eSysAdminMessage), False, 0)
        ElseIf sMsg.StartsWith("/csrloggedout") = True Then
            oPlayer.SendPlayerMessage(CreateChatMsg(-1, "You are logged out as the CSR", ChatMessageType.eSysAdminMessage), False, 0)
        End If

    End Sub

    Public Sub SendOperatorRequestServerDetails(ByVal lConnType As ConnectionType)
        Dim yMsg(5) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestServerDetails).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lConnType).CopyTo(yMsg, lPos) : lPos += 4
        SendMsgToOperator(yMsg)
    End Sub

    Private Sub HandleOperatorRequestPlayerBudget(ByVal yData() As Byte)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)

        Dim bResponse As Boolean = lPlayerID < 0
        lPlayerID = Math.Abs(lPlayerID)

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)

        Dim yResp() As Byte = Nothing

        If oPlayer Is Nothing OrElse oPlayer.oBudget Is Nothing Then
            ReDim yResp(5)
            System.BitConverter.GetBytes(GlobalMessageCode.eBudgetResponse).CopyTo(yResp, 0)
            System.BitConverter.GetBytes(-lPlayerID).CopyTo(yResp, 2)
        Else
            If bResponse = True Then
                System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yData, 0)
                oPlayer.SendPlayerMessage(yData, False, AliasingRights.eViewBudget)
                Return
            Else
                Dim yTemp() As Byte = oPlayer.oBudget.GetObjAsString
                ReDim yResp(1 + yTemp.Length)
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eBudgetResponse).CopyTo(yResp, lPos) : lPos += 2
                yTemp.CopyTo(yResp, lPos) : lPos += yTemp.Length
            End If
        End If
        If yResp Is Nothing = False Then moOperator.SendData(yResp)

    End Sub

    Private Sub HandleOperatorAddFormation(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lFormationID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oFormation As FormationDef = GetEpicaFormation(lFormationID)
        Dim bAdd As Boolean = False
        If oFormation Is Nothing = True Then
            bAdd = True
            oFormation = New FormationDef()
        End If
        If oFormation.FillFromMsg(yData, True) = False Then Return
        If bAdd = True Then AddFormationToGlobalArray(oFormation)
    End Sub

    Private Sub HandleCrossPrimaryBudgetAdjust(ByVal yData() As Byte)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim blCashFlow As Int64 = System.BitConverter.ToInt64(yData, 6)
        Dim blStraightCreditAdjust As Int64 = System.BitConverter.ToInt64(yData, 14)
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then oPlayer.blCredits += blCashFlow + blStraightCreditAdjust
    End Sub

    Private Sub HandleEntityChangingPrimary(ByVal yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lOwnerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lParentDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iParentDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        'ok, Primary now "Owns" the entity
        '  Load the entity from the DB
        Dim oEntity As Epica_Entity = LoadEntityChangingEnvironment(lEntityID, iEntityTypeID)
        If oEntity Is Nothing = True Then
            LogEvent(LogEventType.CriticalError, "Entity Not Load Correctly!!! " & lEntityID & ", " & iEntityTypeID)
            Return
        End If

        'Then, finally, add the unit to the target environment
        If iDestTypeID = ObjectType.ePlanet Then
            'ok, get our planet
            Dim oPlanet As Planet = GetEpicaPlanet(lDestID)
            If oPlanet Is Nothing = False Then
                If oPlanet.oDomain Is Nothing = False Then
                    oPlanet.oDomain.DomainSocket.SendData(GetAddObjectMessage(oEntity, GlobalMessageCode.eAddObjectCommand_CE))
                End If
            End If
        ElseIf iDestTypeID = ObjectType.eSolarSystem Then
            Dim oSystem As SolarSystem = GetEpicaSystem(lDestID)
            If oSystem Is Nothing = False Then
                If oSystem.oDomain Is Nothing = False Then
                    oSystem.oDomain.DomainSocket.SendData(GetAddObjectMessage(oEntity, GlobalMessageCode.eAddObjectCommand_CE))
                End If
            End If
        End If

    End Sub
    Private Sub HandlePrimaryLoadSharedPlayerData(ByVal yData() As Byte)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        If LoadSharedPlayerData(lPlayerID) = False Then
            'TODO: What do we do?
        End If
    End Sub

    Private Sub HandleForwardToPlayerAtPrimary(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return

        If oPlayer.lConnectedPrimaryID = glServerID Then
            If oPlayer.oSocket Is Nothing = False Then
                Dim lLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim yMsg(lLen - 1) As Byte
                Array.Copy(yData, lPos, yMsg, 0, lLen)
                oPlayer.oSocket.SendData(yMsg)
            End If
        End If
    End Sub

    Private Sub HandleOperatorDomainRegister(ByRef yData() As Byte)
        'going to contain a GUID list of all objects I am to control
        '  each object's children list is controlled by me too
        Dim lPos As Int32 = 2       'for msgcode
        glServerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        For X As Int32 = 0 To lCnt - 1
            Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Select Case iTypeID
                Case ObjectType.eUniverse
                    'all geography objects... so loop thru the galaxy object array
                    LogEvent(LogEventType.Informational, "Operator Indicating Entire Universe As Domain")
                    For lGalaxy As Int32 = 0 To glGalaxyUB
                        If glGalaxyIdx(lGalaxy) <> -1 Then goGalaxy(lGalaxy).InMyDomain = True
                    Next lGalaxy
                    'then systems
                    For lSystem As Int32 = 0 To glSystemUB
                        If glSystemIdx(lSystem) <> -1 Then goSystem(lSystem).InMyDomain = True
                    Next lSystem
                    'then planets
                    For lPlanet As Int32 = 0 To glPlanetUB
                        If glPlanetIdx(lPlanet) <> -1 Then goPlanet(lPlanet).InMyDomain = True
                    Next lPlanet
                Case ObjectType.eGalaxy
                    'all of a galaxy
                    For lGalaxy As Int32 = 0 To glGalaxyUB
                        If glGalaxyIdx(lGalaxy) = lID Then
                            LogEvent(LogEventType.Informational, "Operator Indicating Entire " & BytesToString(goGalaxy(lGalaxy).GalaxyName) & " Galaxy as Domain")
                            goGalaxy(lGalaxy).InMyDomain = True
                            Exit For
                        End If
                    Next lGalaxy
                    'systems in galaxy
                    For lSystem As Int32 = 0 To glSystemUB
                        If glSystemIdx(lSystem) <> -1 AndAlso goSystem(lSystem).ParentGalaxy.ObjectID = lID Then
                            goSystem(lSystem).InMyDomain = True
                            goSystem(lSystem).MarkChildrenInMyDomain(True)
                        End If
                    Next lSystem
                Case ObjectType.eSolarSystem
                    'solar system
                    Dim oSystem As SolarSystem = GetEpicaSystem(lID)
                    If oSystem Is Nothing = False Then
                        LogEvent(LogEventType.Informational, "Operator Indicating " & BytesToString(oSystem.SystemName) & " (System) As Domain")
                        oSystem.InMyDomain = True
                        oSystem.MarkChildrenInMyDomain(True)
                    End If
                Case ObjectType.ePlanet
                    'planet
                    Dim oPlanet As Planet = GetEpicaPlanet(lID)
                    If oPlanet Is Nothing = False Then
                        LogEvent(LogEventType.Informational, "Operator Indicating " & BytesToString(oPlanet.PlanetName) & " (Planet) As Domain")
                        oPlanet.InMyDomain = True
                    End If
                Case Else
                    LogEvent(LogEventType.Warning, "Unknown TypeID to register the domain: " & CType(iTypeID, ObjectType).ToString)
            End Select
        Next X

        'Just to be sure...
        For X As Int32 = 0 To glSystemUB
            If glSystemIdx(X) > -1 Then
                Dim oSys As SolarSystem = goSystem(X)
                If oSys Is Nothing = False Then
                    oSys.InMyDomain = True
                    oSys.MarkChildrenInMyDomain(True)
                End If
            End If
        Next X

        bReceivedDomains = True
    End Sub

    Private Sub HandleEmailServerData(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        gsEmailSrvrIP = GetStringFromBytes(yData, lPos, 20) : lPos += 20
        glEmailSrvrPort = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        LogEvent(LogEventType.Informational, "Operator indicating Email Server Connection Data: " & gsEmailSrvrIP & ":" & glEmailSrvrPort)

        If mbConectionToEmailSrvrLost = True Then
            'ok, try to reconnect
            swEmailSrvrReconnect.Stop()
            If ConnectToEmailSrvr() = False Then
                LogEvent(LogEventType.Informational, "Still unable to connect to email server... continuing to wait.")
                swEmailSrvrReconnect.Reset()
                swEmailSrvrReconnect.Start()
            End If
        End If
    End Sub

    Public Sub SendReadyStateToOperator()
        Dim yMsg(1) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eDomainServerReady).CopyTo(yMsg, 0)
        moOperator.SendData(yMsg)
    End Sub

    Private Sub HandleServersReady()
        LogEvent(LogEventType.Informational, "Operator indicating servers ready... starting game loop.")
        gfrmDisplayForm.BeginServerEngine()
    End Sub

    Private Sub HandleOperatorShutdown()
        LogEvent(LogEventType.Informational, "Operator Server indicating shutdown...")
        gfrmDisplayForm.btnShutdown_Click(Nothing, Nothing)
    End Sub

    Private Sub CreateAndSendEmailSrvrPlayerMsg(ByRef oPlayer As Player)
        With oPlayer
            Dim yMsg(329) As Byte       '325
            Dim lPos As Int32 = 0

            If MsgSystem.mb_MONITOR_MSG_ACTIVITY = True Then
                goMsgSys.moMonitor.AddOutMsg(GlobalMessageCode.eAddObjectCommand, MsgMonitor.eMM_AppType.EmailServer, 328, 0)   '324
            End If

            System.BitConverter.GetBytes(328S).CopyTo(yMsg, lPos) : lPos += 2       '324
            System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, lPos) : lPos += 2
            .GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
            .PlayerName.CopyTo(yMsg, lPos) : lPos += 20
            .EmpireName.CopyTo(yMsg, lPos) : lPos += 20
            .RaceName.CopyTo(yMsg, lPos) : lPos += 20
            .ExternalEmailAddress.CopyTo(yMsg, lPos) : lPos += 255
            yMsg(lPos) = .yGender : lPos += 1
            System.BitConverter.GetBytes(.lGuildID).CopyTo(yMsg, lPos) : lPos += 4
            moEmailSrvr.SendLenAppendedData(yMsg)
        End With
    End Sub

	Private Sub HandlePlayerDomainAssignment(ByRef yData() As Byte)
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing = False Then
			oPlayer.InMyDomain = True
            With oPlayer
                CreateAndSendEmailSrvrPlayerMsg(oPlayer)

                Dim yEmail(262) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eEmailSettings).CopyTo(yEmail, 0) 'lPos += 2
                System.BitConverter.GetBytes(.ObjectID).CopyTo(yEmail, 2)   'lpos += 4
                System.BitConverter.GetBytes(.iEmailSettings).CopyTo(yEmail, 6) 'lpos += 2
                .ExternalEmailAddress.CopyTo(yEmail, 8) 'lpos += 255
                Try
                    moEmailSrvr.SendData(yEmail)
                Catch
                End Try
            End With
        End If
	End Sub

	Public Sub SendMsgToOperator(ByRef yData() As Byte)
		If mbConnectedToOperator = True Then
			moOperator.SendData(yData)
		Else
			If mcolOperatorFailQueue Is Nothing Then mcolOperatorFailQueue = New Collection
			mcolOperatorFailQueue.Add(yData)
		End If
	End Sub

	Private Sub HandleOperatorAddObject(ByRef yData() As Byte)
		Dim lPos As Int32 = 2
		Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

		Select Case iTypeID
			Case ObjectType.ePlayer
				If LoadSinglePlayer(lID) = True Then
					Dim oPlayer As Player = GetEpicaPlayer(lID)
					If oPlayer Is Nothing = False Then
						Dim yMsg() As Byte = GetAddObjectMessage(oPlayer, GlobalMessageCode.eAddObjectCommand)
						Me.BroadcastToDomains(yMsg)
					End If
				End If
			Case ObjectType.eSolarSystem
				'ok, a single solar system being added
				If LoadSingleSystem(lID) = True Then
					'ok, let's send our add object for the system to all domains...
					Dim oSystem As SolarSystem = GetEpicaSystem(lID)
					If oSystem Is Nothing = False Then
						Dim yAddMsg() As Byte = GetAddObjectMessage(oSystem, GlobalMessageCode.eAddObjectCommand)
						moPathfinding.SendData(yAddMsg)
						BroadcastToDomains(yAddMsg)
					Else
						LogEvent(LogEventType.CriticalError, "LoadSingleSystem returned true but system is not found.")
					End If

				End If
			Case ObjectType.eMineral
				Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

				If iEnvirTypeID = ObjectType.ePlanet Then
					Dim oPlanet As Planet = GetEpicaPlanet(lEnvirID)
					If oPlanet.oDomain Is Nothing = False Then oPlanet.oDomain.DomainSocket.SendData(yData)
				ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
					Dim oSystem As SolarSystem = GetEpicaSystem(lEnvirID)
					If oSystem.oDomain Is Nothing = False Then oSystem.oDomain.DomainSocket.SendData(yData)
                End If
            Case ObjectType.eUnitDef
                Dim oDef As Epica_Entity_Def = GetEpicaUnitDef(lID)
                If oDef Is Nothing = True Then
                    oDef = New Epica_Entity_Def
                    oDef.FillFromForwardedAddMsg(yData)
                End If
                AddUnitDefToGlobalArray(oDef)
            Case ObjectType.eFacilityDef
                'ok, another primary is saying a unitdef/facilitydef has been added and needs to be added to me as well
                Dim oDef As FacilityDef = GetEpicaFacilityDef(lID)
                If oDef Is Nothing = True Then
                    oDef = New FacilityDef()
                    oDef.FillFromForwardedAddMsg(yData)
                End If
                AddFacilityDefToGlobalArray(oDef)
            Case ObjectType.eAlloyTech
                Dim oTech As New AlloyTech
                oTech.FillFromPrimaryAddMsg(yData)
            Case ObjectType.eArmorTech
                Dim oTech As New ArmorTech
                oTech.FillFromPrimaryAddMsg(yData)
            Case ObjectType.eEngineTech
                Dim oTech As New EngineTech
                oTech.FillFromPrimaryAddMsg(yData)
            Case ObjectType.eHullTech
                Dim oTech As New HullTech
                oTech.FillFromPrimaryAddMsg(yData)
            Case ObjectType.ePrototype
                Dim oTech As New Prototype
                oTech.FillFromPrimaryAddMsg(yData)
            Case ObjectType.eRadarTech
                Dim oTech As New RadarTech
                oTech.FillFromPrimaryAddMsg(yData)
            Case ObjectType.eShieldTech
                Dim oTech As New ShieldTech
                oTech.FillFromPrimaryAddMsg(yData)
            Case ObjectType.eWeaponTech
                Dim yClassType As Byte = Epica_Tech.BASE_OBJ_STRING_SIZE + 22       '2 for msgcode, 20 for name
                Dim oTech As BaseWeaponTech = BaseWeaponTech.CreateWeaponClass(yClassType)
                If oTech Is Nothing = False Then oTech.FillFromPrimaryAddMsg(yData)
        End Select
	End Sub

	Private Sub HandleOperatorAliasConfig(ByRef yData() As Byte)
		'MsgCode 2, Type(1), AliasPlayer(20) but is as PlayerID (4) and OtherPlayer(4)... other 12 is nothing, AliasUN(20), AliasPW(20), Rights(4)
		Dim lPos As Int32 = 2		'for msgcode
		Dim yType As Byte = yData(lPos) : lPos += 1

		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lOtherPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lPos += 12		'remainder of the aliasplayer msg

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing Then Return

		If yType = 0 Then		'remove
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

					For X As Int32 = 0 To mlClientUB
						If mlClientPlayer(X) = oOther.ObjectID Then
							If mlAliasedAs(X) = oPlayer.ObjectID Then mlAliasedRights(X) = lRights
							Exit For
						End If
					Next X
				End If
			End If
		End If

		BroadcastToDomains(yData)
	End Sub

	Private Sub HandleRemovePlayerRel(ByRef yData() As Byte)
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim lWithPlayerID As Int32 = System.BitConverter.ToInt32(yData, 6)

		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing = False Then
			oPlayer.RemovePlayerRel(lWithPlayerID, False)
			oPlayer.SendPlayerMessage(yData, False, AliasingRights.eViewDiplomacy)
		End If
		BroadcastToDomains(yData)
	End Sub

	Public Function CheckOperatorConnection() As Boolean
		If moOperatorSW Is Nothing = False Then
            If moOperatorSW.ElapsedMilliseconds > 20000 Then
                Return False
            End If
		Else
			If mbConnectedToOperator = False Then
				'ok, not connected... should be... so let's connect to the backup operator
				If moOperator Is Nothing = False Then
					Try
						moOperator.Disconnect()
					Catch
					End Try
				End If
				moOperator = Nothing

				If ConnectToOperator(msReconnectOperatorIP, mlReconnectOperatorPort) = False Then
					moOperatorSW = Stopwatch.StartNew()		'start a new stopwatch to wait 10 secs before trying again
				End If
			End If
		End If
		Return True
	End Function

	Public Sub HandleOperatorRebound(ByRef yData() As Byte)
		'the backup operator is telling us that the operator has returned and to connect to it
		msReconnectOperatorIP = gsOperatorIP
		mlReconnectOperatorPort = glOperatorPort

		'so, disconnect from the operator
		mbConnectedToOperator = False
		mbConnectingToOperator = False
		moOperator.Disconnect()
		moOperatorSW = Nothing

		'CheckOperatorConnection will connect us again
	End Sub

	Private Sub HandleSpawnSystemRegion(ByRef yData() As Byte)
		Dim lPos As Int32 = 2
		Dim lSystemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iSystemTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim sIP As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20

		Dim oSystem As SolarSystem = GetEpicaSystem(lSystemID)
		If oSystem Is Nothing Then Return

		'Ok, first, let's set this system and all planets within to my domain
		oSystem.InMyDomain = True
		oSystem.MarkChildrenInMyDomain(True)

		'TODO: We need to load the data related to this system (players, units, facilities, minerals, etc...)
		'  all of it before we send this message to the region

		Dim yMsg(46) As Byte
		lPos = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eRegisterDomain).CopyTo(yMsg, lPos) : lPos += 2
		oSystem.GetObjAsString().CopyTo(yMsg, lPos) : lPos += 45

		'Now, find our region server
		For X As Int32 = 0 To mlDomainUB
			If moDomains(X) Is Nothing = False Then
				If CType(moDomains(X).GetRemoteDetails, System.Net.IPEndPoint).Address.ToString = sIP Then
					moDomains(X).SendData(yMsg)
					Exit For
				End If
			End If
		Next X

    End Sub

    Private Sub HandleOperatorUpdatePlayerDetails(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then
            If LoadSinglePlayer(lPlayerID) = False Then Return
            oPlayer = GetEpicaPlayer(lPlayerID)
            If oPlayer Is Nothing Then Return
            Dim yMsg() As Byte = GetAddObjectMessage(oPlayer, GlobalMessageCode.eAddObjectCommand)
            Me.BroadcastToDomains(yMsg)

            lPos = 0
            ReDim yMsg(323)
            With oPlayer
                System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, lPos) : lPos += 2
                .GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                .PlayerName.CopyTo(yMsg, lPos) : lPos += 20
                .EmpireName.CopyTo(yMsg, lPos) : lPos += 20
                .RaceName.CopyTo(yMsg, lPos) : lPos += 20
                .ExternalEmailAddress.CopyTo(yMsg, lPos) : lPos += 255
                yMsg(lPos) = .yGender : lPos += 1
            End With
            SendToEmailSrvr(yMsg)
            
        Else
            oPlayer.dtAccountWentInactive = Date.MinValue
            ReDim oPlayer.PlayerUserName(19)
            Array.Copy(yData, lPos, oPlayer.PlayerUserName, 0, 20) : lPos += 20
            ReDim oPlayer.PlayerPassword(19)
            Array.Copy(yData, lPos, oPlayer.PlayerPassword, 0, 20) : lPos += 20
            oPlayer.AccountStatus = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            ReDim oPlayer.PlayerName(19)
            Array.Copy(yData, lPos, oPlayer.PlayerName, 0, 20) : lPos += 20
            oPlayer.sPlayerNameProper = BytesToString(oPlayer.PlayerName)
            oPlayer.sPlayerName = oPlayer.sPlayerName.ToUpper
            Me.BroadcastToDomains(yData)
            SendToEmailSrvr(yData)

            If oPlayer.oSocket Is Nothing = False Then
                oPlayer.oSocket.SendData(yData)
            End If
        End If
    End Sub

    Private Sub HandleOperatorSetPlayerSpecialAttribute(ByVal yData() As Byte)
        Dim lPos As Int32 = 2       'for msgcode

        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iSpecialAttr As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lValue As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then

            Select Case iSpecialAttr
                Case ePlayerSpecialAttributeSetting.eBadWarDecCPPenalty
                    oPlayer.BadWarDecCPIncrease = lValue
                Case ePlayerSpecialAttributeSetting.eBadWarDecMoralePenalty
                    oPlayer.BadWarDecMoralePenalty = lValue
                Case ePlayerSpecialAttributeSetting.eDoctrineOfLeadership
                    oPlayer.yCurrentDoctrine = CByte(lValue)
                Case ePlayerSpecialAttributeSetting.ePlayerTitle
                    oPlayer.yPlayerTitle = CByte(lValue)
                    Return
            End Select

            If oPlayer.InMyDomain = True Then
                oPlayer.SendPlayerMessage(yData, False, AliasingRights.eViewUnitsAndFacilities)
            End If
        End If

        Me.BroadcastToDomains(yData)

    End Sub

    Private Sub HandlePlayerDied(ByVal yData() As Byte)
        Dim lPos As Int32 = 2       'for msgcode
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yRespawnFlag As Byte = yData(lPos) : lPos += 1
        Dim yForcedRestart As Byte = yData(lPos) : lPos += 1

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then

            If oPlayer.InMyDomain = True Then
                'MarkPlayerDead

                ''Store our seconds played here in case we need it
                'Dim lSeconds As Int32 = oPlayer.TotalPlayTime
                'If oPlayer.lConnectedPrimaryID > -1 Then
                '    lSeconds += CInt(Now.Subtract(oPlayer.LastLogin).TotalSeconds)
                'End If

                'where the player currently resides
                'If oPlayer.lStartedEnvirID < 500000000 Then
                '    Dim oTmpPlanet As Planet = GetEpicaPlanet(oPlayer.lStartedEnvirID)
                '    If oTmpPlanet Is Nothing = False OrElse oTmpPlanet.ParentSystem Is Nothing = False Then
                '        Try
                '            Dim yGNS(81) As Byte
                '            'Dim lPos As Int32 = 0
                '            lPos = 0
                '            System.BitConverter.GetBytes(GlobalMessageCode.eNewsItem).CopyTo(yGNS, lPos) : lPos += 2
                '            yGNS(lPos) = NewsItemType.ePlayerDeath : lPos += 1
                '            oTmpPlanet.ParentSystem.GetGUIDAsString.CopyTo(yGNS, lPos) : lPos += 6
                '            System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yGNS, lPos) : lPos += 4
                '            System.BitConverter.GetBytes(oPlayer.blMaxPopulation).CopyTo(yGNS, lPos) : lPos += 8
                '            oTmpPlanet.PlanetName.CopyTo(yGNS, lPos) : lPos += 20

                '            oPlayer.PlayerName.CopyTo(yGNS, lPos) : lPos += 20
                '            yGNS(lPos) = oPlayer.yGender : lPos += 1
                '            oPlayer.EmpireName.CopyTo(yGNS, lPos) : lPos += 20

                '            SendToEmailSrvr(yGNS)
                '        Catch
                '        End Try
                '    End If
                '    oTmpPlanet = Nothing
                'End If

                'TODO: Handle the deathbudget better
                If oPlayer.DeathBudgetBalance <> -1 Then 'AndAlso oPlayer.bForcedRestart = False Then
                    'Ok, we're good to go, the player wants to reset
                    oPlayer.blCredits = 0 'oPlayer.DeathBudgetBalance
                    If oPlayer.blCredits < 3000000 Then oPlayer.blCredits += 3000000
                    'oPlayer.DeathBudgetBalance = -1 'set to -1 as another state machine

                    'Begin 24 hour timer on the death budget
                    oPlayer.DeathBudgetEndTime = glCurrentCycle + 2592000       '24 hrs
                Else
                    oPlayer.DeathBudgetBalance = 0
                    oPlayer.blCredits = 3000000
                    'Begin 24 hour timer on the death budget
                    oPlayer.DeathBudgetEndTime = glCurrentCycle + 2592000       '24 hrs
                End If

                oPlayer.blMaxPopulation = 0
                oPlayer.blDBPopulation = 0

                oPlayer.lWarSentiment = 0
                oPlayer.BadWarDecCPIncrease = 0
                oPlayer.BadWarDecCPIncreaseEndCycle = 0
                oPlayer.BadWarDecMoralePenalty = 0
                oPlayer.BadWarDecMoralePenaltyEndCycle = 0
            End If

            'Send to the region servers, they will forward to clients, etc...
            BroadcastToDomains(yData)

            Dim lCurUB As Int32 = -1
            If glUnitIdx Is Nothing = False Then lCurUB = Math.Min(glUnitUB, glUnitIdx.GetUpperBound(0))
            For X As Int32 = 0 To lCurUB
                If glUnitIdx(X) <> -1 Then
                    Dim oUnit As Unit = goUnit(X)
                    If oUnit Is Nothing Then Continue For
                    If oUnit.Owner.ObjectID = oPlayer.ObjectID Then
                        Try
                            oUnit.DeleteEntity(X)
                        Catch
                        End Try
                        glUnitIdx(X) = -1
                    End If
                End If
            Next X

            lCurUB = -1
            If glFacilityIdx Is Nothing = False Then lCurUB = Math.Min(glFacilityUB, glFacilityIdx.GetUpperBound(0))
            For X As Int32 = 0 To lCurUB
                If glFacilityIdx(X) <> -1 Then
                    Dim oFac As Facility = goFacility(X)
                    If oFac Is Nothing = False AndAlso oFac.Owner.ObjectID = oPlayer.ObjectID Then
                        Try
                            oFac.DeleteEntity(X)
                        Catch
                        End Try
                        glFacilityIdx(X) = -1
                        RemoveLookupFacility(oFac.ObjectID, oFac.ObjTypeID)
                    End If
                End If
            Next X

            lCurUB = -1
            If glColonyIdx Is Nothing = False Then lCurUB = Math.Min(glColonyIdx.GetUpperBound(0), glColonyUB)
            For X As Int32 = 0 To lCurUB
                If glColonyIdx(X) <> -1 Then
                    Dim oColony As Colony = goColony(X)
                    If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID = oPlayer.ObjectID Then
                        Try
                            oColony.DeleteColony(Colony.ColonyLostReason.Destruction)
                        Catch
                        End Try
                        glColonyIdx(X) = -1
                    End If
                End If
            Next X

            'remove agents belonging to me
            goAgentEngine.CancelPlayerAgency(lPlayerID)
            goGTC.CancelPlayersOrders(lPlayerID)

            'Remove our factions
            For X As Int32 = 0 To 4
                If oPlayer.lSlotID(X) > 0 Then
                    'ok, need to remove me from the faction of the player
                    Dim oSlotPlayer As Player = GetEpicaPlayer(oPlayer.lSlotID(X))
                    If oSlotPlayer Is Nothing = False Then
                        For Y As Int32 = 0 To 2
                            If oSlotPlayer.lFactionID(Y) = oPlayer.ObjectID Then
                                oSlotPlayer.lFactionID(Y) = -1
                                oSlotPlayer.ReverifySlots()
                                Exit For
                            End If
                        Next Y
                    End If
                End If
                oPlayer.lSlotID(X) = -1
                oPlayer.ySlotState(X) = 0
            Next X
            For X As Int32 = 0 To 2
                If oPlayer.lFactionID(X) > 0 Then
                    Dim oFactionPlayer As Player = GetEpicaPlayer(oPlayer.lFactionID(X))
                    If oFactionPlayer Is Nothing = False Then
                        For Y As Int32 = 0 To 4
                            If oFactionPlayer.lSlotID(Y) = oPlayer.ObjectID Then
                                oFactionPlayer.lSlotID(Y) = -1
                                oFactionPlayer.ReverifySlots()
                                Exit For
                            End If
                        Next Y
                    End If
                End If
                oPlayer.lFactionID(X) = -1
            Next X
            oPlayer.ReverifySlots()

            For X As Int32 = 0 To glPlanetUB
                If glPlanetIdx(X) > -1 Then
                    Dim oPlanet As Planet = goPlanet(X)
                    If oPlanet Is Nothing = False Then
                        If oPlanet.OwnerID = oPlayer.ObjectID Then
                            oPlanet.OwnerID = 0
                            oPlanet.SaveObject()
                        End If
                    End If
                End If
            Next X

            If oPlayer.InMyDomain = True Then
                oPlayer.PlayerIsDead = False
                'If oPlayer.yPlayerPhase = eyPlayerPhase.eInitialPhase AndAlso yRespawnFlag <> 255 Then
                '	'Ok, player is in the initial tutorial phase...
                '	' - add Aurelius as a diplomatic contact and set Aurelius to blood war but my rel towards aurelius to neutral
                '	Dim oPlayer2 As Player = GetEpicaPlayer(gl_HARDCODE_PIRATE_PLAYER_ID)
                '	If oPlayer2 Is Nothing = False Then
                '		Dim oTmpRel As PlayerRel = New PlayerRel
                '		oTmpRel.oPlayerRegards = oPlayer
                '		oTmpRel.oThisPlayer = oPlayer2
                '		oTmpRel.WithThisScore = elRelTypes.eNeutral
                '		oPlayer.SetPlayerRel(gl_HARDCODE_PIRATE_PLAYER_ID, oTmpRel)

                '		Dim yMsg2() As Byte = GetAddPlayerRelMessage(oTmpRel)
                '		For X As Int32 = 0 To mlDomainUB
                '			moDomains(X).SendData(yMsg2)
                '		Next X
                '	End If

                '	' - store the time played so far and reset it to 0
                '	oPlayer.PlayedTimeInTutorialOne = lSeconds
                '	oPlayer.TotalPlayTime = 0

                '	' - update phase to phase 2
                '	oPlayer.yPlayerPhase = eyPlayerPhase.eSecondPhase

                '	' - spawn the player in tutorial system 2
                '	' Because of the above, when the player logs back in, they will be in tutorial system 2

                'ElseIf oPlayer.yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then
                If yRespawnFlag = 254 Then oPlayer.yRespawnWithGuild = 1

                If oPlayer.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                    oPlayer.lLastViewedEnvir = 0
                    oPlayer.iLastViewedEnvirType = 0
                    oPlayer.yPlayerPhase = eyPlayerPhase.eSecondPhase
                    oPlayer.lTutorialStep = 301
                    oPlayer.DeathBudgetFundsRemaining = CInt(oPlayer.blCredits)
                    oPlayer.PlayedTimeInTutorialOne = oPlayer.TotalPlayTime

                    oPlayer.lStartedEnvirID = -1
                    'InitializePlayer(oPlayer, oPlayer.lStartedEnvirID)

                    Dim yDisco(2) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.ePlayerIsDead).CopyTo(yDisco, 0)
                    yDisco(2) = 2
                    oPlayer.SendPlayerMessage(yDisco, False, 0)
                ElseIf oPlayer.yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then
                    'Player is in second or higher tutorial phase, from here...
                    ' - remove all technologies of the player and all Mineral Knowledge and any score-based values, gns values, running tallies, etc...
                    oPlayer.WipePlayer()
                    ' - remove all death budget stuff

                    'If oPlayer.AccountStatus = AccountStatusType.eTrialAccount AndAlso gb_IN_OPEN_BETA = False Then
                    '    oPlayer.lLastViewedEnvir = 0
                    '    oPlayer.iLastViewedEnvirType = 0
                    '    oPlayer.lStartedEnvirID = -1
                    '    oPlayer.yPlayerPhase = eyPlayerPhase.eSecondPhase
                    '    oPlayer.DeathBudgetEndTime = -1
                    '    oPlayer.DeathBudgetBalance = 0
                    '    oPlayer.DeathBudgetFundsRemaining = 0
                    '    oPlayer.blCredits = 0
                    'Else
                    oPlayer.lLastViewedEnvir = 0
                    oPlayer.iLastViewedEnvirType = 0
                    oPlayer.lStartedEnvirID = -1
                    oPlayer.yPlayerPhase = eyPlayerPhase.eFullLivePhase
                    oPlayer.DeathBudgetEndTime = -1
                    oPlayer.DeathBudgetBalance = 0
                    oPlayer.DeathBudgetFundsRemaining = 0
                    oPlayer.blCredits = 0

                    'if the player went thru the tutorial normally...
                    If oPlayer.TutorialPhaseWaves <> 0 AndAlso oPlayer.lTutorialStep > 309 Then
                        oPlayer.DeathBudgetEndTime = glCurrentCycle + 2592000
                        oPlayer.DeathBudgetFundsRemaining = 30000000
                        oPlayer.blCredits = 30000000
                    End If

                    If oPlayer.blCredits < 30000000 Then
                        oPlayer.blCredits = 30000000
                        oPlayer.DeathBudgetEndTime = glCurrentCycle + 2592000
                        oPlayer.DeathBudgetFundsRemaining = 30000000
                    End If
                    'End If

                    ''The player is going to get disconnected soon...
                    'Dim yResp(2) As Byte
                    'Dim lPos As Int32 = 0
                    'System.BitConverter.GetBytes(GlobalMessageCode.ePlayerIsDead).CopyTo(yResp, 0) : lPos += 2
                    'yResp(lPos) = 2 : lPos += 1
                    'oPlayer.SendPlayerMessage(yResp, False, 0)

                    'With the above command, the player will disconnect, we do not call initialize player
                    'InitializePlayer(oPlayer, -1)
                Else

                    oPlayer.RemovePlayerRel(-1, False)

                    oPlayer.bSpawnAtSameLoc = False
                    If oPlayer.dtLastRespawn <> Date.MinValue AndAlso Now.Subtract(oPlayer.dtLastRespawn).TotalHours < 24 Then
                        'ok, did the player get declared war on?
                        If oPlayer.bDeclaredWarOn = False Then
                            oPlayer.bSpawnAtSameLoc = True
                        End If
                    End If

                    oPlayer.lLastViewedEnvir = 0
                    oPlayer.iLastViewedEnvirType = 0
                    oPlayer.DeathBudgetFundsRemaining = CInt(oPlayer.blCredits)
                    oPlayer.dtLastRespawn = Now
                End If
            End If
        End If

        'regardless, return the message to the primary with a negative playerid
        System.BitConverter.GetBytes(-lPlayerID).CopyTo(yData, 2)
        moOperator.SendData(yData)

    End Sub

    Private Sub HandleForcePrimarySync(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Select Case iObjTypeID
            Case ObjectType.eMineral
                Dim oMineral As Mineral = GetEpicaMineral(lObjID)
                If oMineral Is Nothing = True Then
                    oMineral = New Mineral
                    oMineral.FillFromPrimarySyncMsg(yData)
                    glMineralUB += 1
                    ReDim Preserve goMineral(glMineralUB)
                    ReDim Preserve glMineralIdx(glMineralUB)
                    goMineral(glMineralUB) = oMineral
                    goMineral(glMineralUB).ServerIndex = glMineralUB
                    glMineralIdx(glMineralUB) = oMineral.ObjectID
                End If
            Case ObjectType.eWormhole
                'ok, player wormhole knowledge
                Dim oWormhole As Wormhole = GetEpicaWormhole(lObjID)
                Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lSystemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                Dim oSystem As SolarSystem = GetEpicaSystem(lSystemID)
                If oPlayer Is Nothing = False AndAlso oSystem Is Nothing = False Then
                    oPlayer.AddWormholeKnowledge(oWormhole, False, oSystem, False)

                    Dim yTemp(13) As Byte
                    Dim lTempPos As Int32 = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.ePlayerDiscoversWormhole).CopyTo(yTemp, lTempPos) : lTempPos += 2
                    System.BitConverter.GetBytes(lPlayerID).CopyTo(yTemp, lTempPos) : lTempPos += 4
                    System.BitConverter.GetBytes(lObjID).CopyTo(yTemp, lTempPos) : lTempPos += 4
                    System.BitConverter.GetBytes(lSystemID).CopyTo(yTemp, lTempPos) : lTempPos += 4


                    If oSystem.InMyDomain = True Then
                        If oSystem.oDomain Is Nothing = False Then
                            oSystem.oDomain.DomainSocket.SendData(yTemp)
                        End If
                    End If
                End If
        End Select
    End Sub

    Private Sub HandleOperatorRequestPassThru(ByVal yData() As Byte)
        Dim lPos As Int32 = 2       'for msgcode
        Dim lPrimIdx As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        'Ok, the operator is telling me that a Primary out there is requesting data that I may know about... so, let's look it up
        '  the request is coming from a client that the primary could not answer... so... we need to put any pass thru msgs in here...
        Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yData, lPos) ': lPos += 2

        Dim yRespMsg() As Byte = Nothing

        Select Case iMsgCode
            Case GlobalMessageCode.eRequestEmail
                Dim oPlayer As Player = GetEpicaPlayer(lObjID)
                If oPlayer Is Nothing = False Then
                    If oPlayer.InMyDomain = False Then
                        LogEvent(LogEventType.CriticalError, "PassThruRequest on a player that is not in my domain!")
                        Return
                    End If
                    yRespMsg = DoRequestEmail(oPlayer, yData, lPos)
                End If
            Case GlobalMessageCode.eMarkEmailReadStatus
                Dim oPlayer As Player = GetEpicaPlayer(lObjID)
                If oPlayer Is Nothing OrElse oPlayer.InMyDomain = False Then
                    LogEvent(LogEventType.CriticalError, "PassThruRequest on a player that is not in my domain!")
                    Return
                End If
                DoMarkEmailReadStatus(oPlayer, yData, lPos)
                Return
            Case GlobalMessageCode.eDeleteEmailItem
                Dim oPlayer As Player = GetEpicaPlayer(lObjID)
                If oPlayer Is Nothing OrElse oPlayer.InMyDomain = False Then
                    LogEvent(LogEventType.CriticalError, "PassThruRequest on a player that is not in my domain!")
                    Return
                End If
                yRespMsg = DoDeleteEmailItem(oPlayer, yData, lPos)
            Case GlobalMessageCode.eMoveEmailToFolder
                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                If oPlayer Is Nothing OrElse oPlayer.InMyDomain = False Then
                    LogEvent(LogEventType.CriticalError, "PassThruRequest on a player that is not in my domain!")
                    Return
                End If
                yRespMsg = DoMoveEmailToFolder(oPlayer, yData, lPos)
            Case GlobalMessageCode.eAddPlayerCommFolder
                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                If oPlayer Is Nothing OrElse oPlayer.InMyDomain = False Then
                    LogEvent(LogEventType.CriticalError, "PassThruRequest on a player that is not in my domain!")
                    Return
                End If
                yRespMsg = DoAddPlayerCommFolder(oPlayer, yData, lPos)
            Case GlobalMessageCode.eEmailSettings
                Dim oPlayer As Player = GetEpicaPlayer(lObjID)
                If oPlayer Is Nothing OrElse oPlayer.InMyDomain = False Then
                    LogEvent(LogEventType.CriticalError, "PassThruRequest on a player that is not in my domain!")
                    Return
                End If
                yRespMsg = DoEmailSettings(oPlayer, yData, lPos)
            Case GlobalMessageCode.eSendEmail, GlobalMessageCode.eSaveEmailDraft
                Dim oPlayer As Player = GetEpicaPlayer(lObjID)
                If oPlayer Is Nothing OrElse oPlayer.InMyDomain = False Then
                    LogEvent(LogEventType.CriticalError, "PassThruRequest on a player that is not in my domain!")
                    Return
                End If
                yRespMsg = DoSendEmail(oPlayer, yData, lPos)
            Case GlobalMessageCode.eAddObjectCommand, GlobalMessageCode.eAddObjectCommand_CE
                yRespMsg = DoPassThruAddCommand(lObjID, yData, lPos)
            Case GlobalMessageCode.eGetTradeHistory
                Dim oPlayer As Player = GetEpicaPlayer(lObjID)
                If oPlayer Is Nothing OrElse oPlayer.InMyDomain = False Then
                    LogEvent(LogEventType.CriticalError, "PassThruRequest on a player that is not in my domain!")
                    Return
                End If
                yRespMsg = oPlayer.GetTradeHistoryMsg()
            Case GlobalMessageCode.eDeleteTradeHistoryItem
                Dim oPlayer As Player = GetEpicaPlayer(lObjID)
                If oPlayer Is Nothing OrElse oPlayer.InMyDomain = False Then
                    LogEvent(LogEventType.CriticalError, "PassThruRequest on a player that is not in my domain!")
                    Return
                End If
                DoDeleteTradeHistoryItem(oPlayer, yData, lPos)
            Case GlobalMessageCode.eGetIntelSellOrderDetail
                Dim oPlayer As Player = GetEpicaPlayer(lObjID)
                If oPlayer Is Nothing OrElse oPlayer.InMyDomain = False Then
                    LogEvent(LogEventType.CriticalError, "PassThruRequest on a player that is not in my domain!")
                    Return
                End If
                yRespMsg = DoGetIntelSellOrderDetail(oPlayer, yData, lPos)
        End Select

        If yRespMsg Is Nothing = False Then
            Dim yResp(15 + yRespMsg.Length) As Byte
            lPos = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eOperatorResponsePassThru).CopyTo(yResp, lPos) : lPos += 2
            System.BitConverter.GetBytes(lPrimIdx).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(lObjID).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, lPos) : lPos += 2
            yRespMsg.CopyTo(yResp, lPos) : lPos += yRespMsg.Length

            moOperator.SendData(yResp)
        End If
    End Sub
    Private Sub HandleOperatorResponsePassThru(ByVal yData() As Byte)
        Dim lPos As Int32 = 6       'msgcode and primidx
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lPos += 6

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            If oPlayer.oSocket Is Nothing = False Then
                Dim lLen As Int32 = yData.Length - lPos
                Dim yForward(lLen) As Byte
                Array.Copy(yData, lPos, yForward, 0, lLen)
                oPlayer.oSocket.SendData(yForward)
            End If
        End If
    End Sub

    Private Sub HandlePlayerConnectedPrimary(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPrimaryID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then oPlayer.lConnectedPrimaryID = lPrimaryID
    End Sub

    Public Sub SendPassThruMsg(ByVal yData() As Byte, ByVal lPlayerID As Int32, ByVal lObjID As Int32, ByVal iObjTypeID As Int16)

        'need to route this request thru to the operator
        Dim yForward(15 + yData.Length) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eOperatorRequestPassThru).CopyTo(yForward, lPos) : lPos += 2
        lPos += 4        'for return to primary Idx...
        System.BitConverter.GetBytes(lPlayerID).CopyTo(yForward, lPos) : lPos += 4
        System.BitConverter.GetBytes(lObjID).CopyTo(yForward, lPos) : lPos += 4
        System.BitConverter.GetBytes(iObjTypeID).CopyTo(yForward, lPos) : lPos += 2
        yData.CopyTo(yForward, lPos) : lPos += yData.Length
        moOperator.SendData(yForward)

    End Sub
#End Region

#Region "  Encryption and Decryption Functions  "
    'Encryption and Decryption routines here...
    Private ml_ENCRYPT_SEED As Int32 = 777
    Public Function Encrypt(ByVal sVal As String) As String
        Dim lLen As Int32
        Dim lKey As Int32
        Dim oSB As New System.Text.StringBuilder()
        Dim X As Int32
        Dim lOffset As Int32
        Dim lChrCode As Int32
        Dim lMod As Int32

        lLen = sVal.Length
        lKey = CInt(Int(Rnd() * 51))

        'set our seed
        Rnd(-1)
        Call Randomize(ml_ENCRYPT_SEED + lKey)

        'Now encrypt it... we do this by taking the loc - the length as an offset...
        oSB.Append(ChrW(lKey))
        For X = 1 To lLen
            lOffset = X - lLen

            'Now, find out what we got here... it is assumed that the string is already encoded regarding numerics
            lChrCode = Asc(Mid$(sVal, X, 1))

            'Now, add our value... 1 to 5
            lMod = CInt(Int(Rnd() * 5) + 1)
            lChrCode += lMod
            If lChrCode > 255 Then lChrCode = 255 - lChrCode
            oSB.Append(ChrW(lChrCode))
        Next X

        Return oSB.ToString
    End Function

    Public Function Decrypt(ByVal sVal As String) As String
        Dim lLen As Int32
        Dim lKey As Int32
        Dim oSB As New System.Text.StringBuilder()
        Dim X As Int32
        Dim lChrCode As Int32
        Dim lMod As Int32

        'Get our key value...
        lKey = Asc(Mid$(sVal, 1, 1))
        lLen = sVal.Length

        'Set up our seed
        Rnd(-1)
        Randomize(ml_ENCRYPT_SEED + lKey)

        For X = 2 To lLen
            'Now, find out what we got
            lChrCode = Asc(Mid$(sVal, X, 1))
            'now, subtract our value... 1 to 5
            lMod = CInt(Int(Rnd() * 5) + 1)
            lChrCode -= lMod
            If lChrCode < 0 Then lChrCode = 255 + lChrCode
            oSB.Append(ChrW(lChrCode))
        Next X
        Return oSB.ToString
    End Function

#End Region

#Region "  Shared Message Definitions and Handling  "
    Private Sub HandleSetEntityProduction(ByVal lClientIndex As Int32, ByVal yData() As Byte)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lProdID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iProdTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
        Dim lQuantity As Int32 = 1

        If yData.GetUpperBound(0) >= 15 Then
            lQuantity = System.BitConverter.ToInt32(yData, 14)
        End If

        Dim lLocX As Int32
        Dim lLocZ As Int32
        Dim iLocA As Int16

        If lClientIndex = -1 Then
            'came from pathfinding server, includes locational data
            lLocX = System.BitConverter.ToInt32(yData, 14)
            lLocZ = System.BitConverter.ToInt32(yData, 18)
            iLocA = System.BitConverter.ToInt16(yData, 22)
        Else : lLocX = 0 : lLocZ = 0 : iLocA = 0
        End If

        Dim X As Int32
        Dim oED As Epica_Entity_Def = Nothing
        Dim lItemOwnerID As Int32

        If iProdTypeID = ObjectType.eUnitDef OrElse iProdTypeID = ObjectType.eFacilityDef Then
            oED = CType(GetEpicaObject(lProdID, iProdTypeID), Epica_Entity_Def)
            lItemOwnerID = oED.OwnerID
        End If

        'TODO: Verify that the player can in fact build the item referred to by lProdID and iProdTypeID
        If lClientIndex <> -1 Then
            If lItemOwnerID <> 0 AndAlso lItemOwnerID <> mlClientPlayer(lClientIndex) AndAlso (lItemOwnerID <> mlAliasedAs(lClientIndex) OrElse (mlAliasedRights(lClientIndex) And AliasingRights.eAddProduction) = 0) Then
                Return
            End If
        End If

        Dim oPlayer As Player = Nothing

        'only facilities and units can produce items
        If iObjTypeID = ObjectType.eFacility Then
            For X = 0 To glFacilityUB
                If glFacilityIdx(X) = lObjID Then
                    Dim oFac As Facility = goFacility(X)
                    If oFac Is Nothing Then Exit For
                    oPlayer = oFac.Owner
                    If oFac.Owner.ObjectID = lItemOwnerID OrElse lItemOwnerID = 0 OrElse oFac.Active = False Then
                        'If goFacility(X).Owner.oSocket Is Nothing = False Then
                        '    lClientIndex = goFacility(X).Owner.oSocket.SocketIndex
                        'End If
                        If oFac.yProductionType = ProductionType.eResearch AndAlso oFac.bProducing = True AndAlso lProdID <> -1 AndAlso iProdTypeID <> -1 AndAlso lQuantity <> -1 Then
                            System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                        Else
                            If oFac.yProductionType = ProductionType.eSpaceStationSpecial AndAlso iProdTypeID = ObjectType.eFacilityDef Then
                                Dim lFacCnt As Int32 = 0
                                For Y As Int32 = 0 To oFac.ParentColony.ChildrenUB
                                    If oFac.ParentColony.lChildrenIdx(Y) <> -1 Then
                                        lFacCnt += 1
                                    End If
                                Next Y
                                If oFac.CurrentProduction Is Nothing = False Then
                                    If oFac.CurrentProduction.ProductionTypeID = ObjectType.eFacilityDef Then
                                        lFacCnt += oFac.CurrentProduction.lProdCount
                                    End If
                                End If
                                For Y As Int32 = 0 To oFac.ProductionUB
                                    If oFac.ProductionIdx(Y) > -1 Then
                                        Dim oProd As EntityProduction = oFac.Production(Y)
                                        If oProd Is Nothing = False Then
                                            If oProd.ProductionTypeID = ObjectType.eFacilityDef Then lFacCnt += oProd.lProdCount
                                        End If
                                    End If
                                Next Y

                                'ok, fac count
                                Dim lMaxModuleCnt As Int32 = oFac.EntityDef.Structure_MaxHP \ 50000
                                lMaxModuleCnt += 1      'for the station itself
                                If lFacCnt + lQuantity > lMaxModuleCnt Then
                                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                                    Exit For
                                End If
                            End If

                            'If the prodtypeid is -MineralTech, then we need to change lQuantity to Int32.MaxValue
                            If iProdTypeID = -ObjectType.eMineralTech Then
                                iProdTypeID = ObjectType.eMineralTech
                                lQuantity = Int32.MaxValue
                            ElseIf iProdTypeID = ObjectType.eMineralTech AndAlso lProdID = 41991 Then
                                If oFac.ParentColony Is Nothing = False AndAlso oFac.ParentColony.GetColonyIntelligence < 160 Then
                                    oPlayer.SendPlayerMessage(CreateChatMsg(-1, "This material must be studied at a colony with at least 160 population intelligence.", ChatMessageType.eSysAdminMessage), False, 0)
                                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                                    Exit For
                                End If
                            End If

                            If iProdTypeID = ObjectType.eSpecialTech Then
                                Dim oST As PlayerSpecialTech = CType(oPlayer.GetTech(lProdID, iProdTypeID), PlayerSpecialTech)
                                If oST Is Nothing Then Return
                                If oST.bLinked = False Then
                                    LogEvent(LogEventType.PossibleCheat, "Player attempting to research unlinked tech: " & mlClientPlayer(lClientIndex))
                                    Return
                                ElseIf oST.bInTheTank = True Then
                                    LogEvent(LogEventType.PossibleCheat, "Player attempting to research bypassed tech: " & mlClientPlayer(lClientIndex))
                                    Return
                                ElseIf oST.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                                    LogEvent(LogEventType.PossibleCheat, "Player attempting to research researched tech: " & mlClientPlayer(lClientIndex))
                                    Return
                                End If
                            End If

                            If oPlayer.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                                If iProdTypeID = ObjectType.eUnitDef Then
                                    'ok, look for the units in the initial phase
                                    Dim lCPUsage As Int32 = oPlayer.oBudget.GetEnvirCPUsage(CType(oFac.ParentObject, Epica_GUID).ObjectID, ObjectType.ePlanet)
                                    If lCPUsage + (10 * lQuantity) > 450 Then
                                        lQuantity = (450 - lCPUsage) \ 10
                                        If lQuantity < 1 Then Return
                                    End If
                                End If
                            End If

                            If lProdID = Int32.MinValue Then
                                oFac.bExcludeFromColonyQueue = Not oFac.bExcludeFromColonyQueue
                                Return
                            End If

                            'pass 254 as our ordernumber, tells it to put it at the end if its not there, or leave it unchanged if it already exists
                            If oFac.AddProduction(lProdID, iProdTypeID, 254, lQuantity, 0) = True Then
                                'Ok, we got it... respond with success
                                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdSucceed).CopyTo(yData, 0)

                                AddEntityProducing(X, ObjectType.eFacility, oFac.ObjectID)

                                'Ok, now, forcefully update the production status
                                If oFac.Owner.lConnectedPrimaryID > -1 OrElse (oFac.Owner.HasOnlineAliases(AliasingRights.eAddProduction Or AliasingRights.eCancelProduction) = True) Then
                                    Dim yTmpMsg(7) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.eEntityProductionStatus).CopyTo(yTmpMsg, 0)
                                    oFac.GetGUIDAsString.CopyTo(yTmpMsg, 2)
                                    Dim yResp() As Byte = HandleRequestProductionStatus(yTmpMsg)
                                    If yResp Is Nothing = False Then oFac.Owner.CrossPrimarySafeSendMsg(yResp) ', oFac.Owner.oSocket)
                                End If
                            Else
                                Return
                                'System.BitConverter.GetBytes(EpicaMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                            End If
                        End If
                    Else
                        System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                    End If
                    Exit For
                End If
            Next X
        ElseIf iObjTypeID = ObjectType.eUnit Then
            For X = 0 To glUnitUB
                If glUnitIdx(X) = lObjID Then

                    Dim oUnit As Unit = goUnit(X)
                    If oUnit Is Nothing Then
                        glUnitIdx(X) = -1
                        Return
                    End If

                    oPlayer = oUnit.Owner
                    If (oUnit.Owner.ObjectID <> lItemOwnerID AndAlso lItemOwnerID <> 0) OrElse oUnit.bProducing = True Then
                        System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                    ElseIf CType(oUnit.ParentObject, Epica_GUID).ObjTypeID <> ObjectType.ePlanet AndAlso CType(oUnit.ParentObject, Epica_GUID).ObjTypeID <> ObjectType.eSolarSystem Then
                        System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                    Else
                        If lClientIndex = -1 Then
                            'Ok, pathfinding server is stating a unit is building a facility, the unit MUST be at the location specified for this to occur
                            oUnit.LocX = lLocX
                            oUnit.LocZ = lLocZ
                            oUnit.LocAngle = iLocA
                        End If

                        'Set our Client Index
                        'If oUnit.Owner.oSocket Is Nothing = False Then
                        '    lClientIndex = oUnit.Owner.oSocket.SocketIndex
                        'End If

                        'Confirm that this object is not a command center special
                        Dim bGood As Boolean = True
                        Dim bAlertWarDecPot As Boolean = False
                        Dim lMaxWpnRng As Int32 = 0
                        Dim oColony As Colony = Nothing
                        Dim bCheckWpns As Boolean = False

                        If oED Is Nothing = False Then

                            If oUnit.EntityDef.ObjectID = 28344 Then            'replace with the super engineer id
                                If oED.ProductionTypeID <> ProductionType.eCommandCenterSpecial Then
                                    bGood = False
                                End If
                            End If

                            If (oED.ModelID And 255) = 148 Then
                                'ensure we are in a solarsystem 
                                If CType(oUnit.ParentObject, Epica_GUID).ObjTypeID <> ObjectType.eSolarSystem Then
                                    bGood = False
                                Else
                                    Dim lRingPlanetDist As Int32 = 0
                                    Dim oRingPlanet As Planet = CType(oUnit.ParentObject, SolarSystem).GetNearestRingedPlanet(lLocX, lLocZ, lRingPlanetDist)
                                    If oRingPlanet Is Nothing = False Then
                                        If oRingPlanet.RingMiner Is Nothing = False Then
                                            bGood = False
                                        ElseIf oRingPlanet.RingMineralID < 1 Then
                                            bGood = False
                                        ElseIf oRingPlanet.OuterRingRadius < lRingPlanetDist Then
                                            bGood = False
                                        End If
                                    Else
                                        bGood = False
                                    End If
                                End If
                            End If

                            'get our max wpn range
                            For Y As Int32 = 0 To oED.WeaponDefUB
                                If oED.WeaponDefs(Y).oWeaponDef.Range > lMaxWpnRng Then lMaxWpnRng = oED.WeaponDefs(Y).oWeaponDef.Range
                                bCheckWpns = True
                            Next Y
                            If oED.OptRadarRange > lMaxWpnRng Then lMaxWpnRng = oED.OptRadarRange
                            lMaxWpnRng += 10

                            'which it should always be SOMETHING...
                            If oED.ProductionTypeID = ProductionType.eCommandCenterSpecial OrElse oED.ProductionTypeID = ProductionType.eTradePost Then
                                'Ok, now, check if there is a colony here belonging to the owner

                                Dim lTmpIdx As Int32 = -1
                                'Get the colony Idx....
                                With CType(oUnit.ParentObject, Epica_GUID)
                                    lTmpIdx = oUnit.Owner.GetColonyFromParent(.ObjectID, .ObjTypeID)
                                End With
                                'Did we get a result?
                                If lTmpIdx <> -1 Then
                                    'Yup, ok that index is our goColony index...
                                    oColony = goColony(lTmpIdx)
                                    If oColony Is Nothing = False Then
                                        With oColony
                                            If .bCCInProduction = True AndAlso oED.ProductionTypeID = ProductionType.eCommandCenterSpecial Then
                                                bGood = False
                                            ElseIf .bTradepostInProduction = True AndAlso oED.ProductionTypeID = ProductionType.eTradePost Then
                                                bGood = False
                                            Else
                                                For lTmpFacIdx As Int32 = 0 To .ChildrenUB
                                                    If .lChildrenIdx(lTmpFacIdx) <> -1 Then
                                                        If (oED.ProductionTypeID = ProductionType.eCommandCenterSpecial AndAlso .oChildren(lTmpFacIdx).yProductionType = ProductionType.eCommandCenterSpecial) OrElse _
                                                           (oED.ProductionTypeID = ProductionType.eTradePost AndAlso .oChildren(lTmpFacIdx).yProductionType = ProductionType.eTradePost) Then
                                                            'return invalid...
                                                            bGood = False
                                                            Exit For
                                                        End If
                                                    End If
                                                Next
                                            End If

                                            If bGood = True Then
                                                If oED.ProductionTypeID = ProductionType.eTradePost Then
                                                    .bTradepostInProduction = True
                                                ElseIf oED.ProductionTypeID = ProductionType.eCommandCenterSpecial Then
                                                    .bCCInProduction = True
                                                End If
                                            End If
                                        End With
                                    Else
                                        'No, ok, we have to check all units within this environment that belong to me that are building something, if they are building the cc/tp
                                        With CType(oUnit.ParentObject, Epica_GUID)
                                            If BuildingCommandCenterOrTradepost(oUnit.Owner.ObjectID, .ObjectID, .ObjTypeID, oED.ProductionTypeID) = True Then
                                                bGood = False
                                            End If
                                        End With
                                    End If
                                Else
                                    With CType(oUnit.ParentObject, Epica_GUID)
                                        If BuildingCommandCenterOrTradepost(oUnit.Owner.ObjectID, .ObjectID, .ObjTypeID, oED.ProductionTypeID) = True Then
                                            bGood = False
                                        End If
                                    End With
                                End If
                            End If
                        End If

                        If bGood = True Then
                            'Region Server handled the terrain grade for us... we need to handle Planet Proximity
                            If CType(oUnit.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
                                Dim lMinDist As Int32 = 0
                                If oED Is Nothing OrElse oED.ProductionTypeID = ProductionType.eSpaceStationSpecial Then
                                    lMinDist = 15000
                                    Dim oTech As Epica_Tech = oUnit.Owner.GetTech(308, ObjectType.eSpecialTech)
                                    If oTech Is Nothing = False AndAlso oTech.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                                        lMinDist = 6000
                                    End If
                                    oTech = Nothing
                                End If

                                'Now, find any planets nearby
                                With CType(oUnit.ParentObject, SolarSystem)
                                    bGood = .StationProximityTest(lMinDist, lLocX, lLocZ)
                                End With
                            End If

                            'Ok, get our rcTemp which will be the rectangle that the facility will be placed
                            Dim rcTemp As Rectangle = goModelDefs.GetModelDef(oED.ModelID).rcSpacing
                            rcTemp.Location = New Point(lLocX - (rcTemp.Width \ 2), lLocZ - (rcTemp.Height \ 2))
                            Dim lTmpPOID As Int32
                            Dim iTmpPOID As Int16
                            With CType(oUnit.ParentObject, Epica_GUID)
                                lTmpPOID = .ObjectID
                                iTmpPOID = .ObjTypeID
                            End With
                            'First, verify that no facilities will be a problem and we'll check entity def's wpn ranges too
                            lMaxWpnRng *= gl_FINAL_GRID_SQUARE_SIZE
                            Dim rcWpnRng As Rectangle = Rectangle.FromLTRB(lLocX - lMaxWpnRng, lLocZ - lMaxWpnRng, lLocX + lMaxWpnRng, lLocZ + lMaxWpnRng)
                            For Y As Int32 = 0 To glFacilityUB
                                If glFacilityIdx(Y) <> -1 Then
                                    Dim oFac As Facility = goFacility(Y)
                                    If oFac Is Nothing Then
                                        glFacilityIdx(Y) = -1
                                        Continue For
                                    End If
                                    If oFac.ParentObject Is Nothing Then
                                        glFacilityIdx(Y) = -1
                                        Continue For
                                    End If
                                    With CType(oFac.ParentObject, Epica_GUID)
                                        If .ObjectID <> lTmpPOID OrElse .ObjTypeID <> iTmpPOID Then Continue For
                                    End With
                                    Dim rcTest As Rectangle = goModelDefs.GetModelDef(oFac.EntityDef.ModelID).rcSpacing

                                    'MSC - 12/09/08 - put this in for reverse lookup
                                    Dim bReverseCheckWpns As Boolean = False
                                    Dim oTmpDef As FacilityDef = oFac.EntityDef
                                    Dim lTmpRng As Int32 = 0
                                    For Z As Int32 = 0 To oTmpDef.WeaponDefUB
                                        If oTmpDef.WeaponDefs(Z).oWeaponDef.Range > lTmpRng Then lTmpRng = oTmpDef.WeaponDefs(Z).oWeaponDef.Range
                                        bReverseCheckWpns = True
                                    Next Z
                                    If oTmpDef.OptRadarRange > lTmpRng Then lTmpRng = oTmpDef.OptRadarRange
                                    lTmpRng += 10

                                    rcTest.Location = New Point(oFac.LocX - (rcTest.Width \ 2), oFac.LocZ - (rcTest.Height \ 2))
                                    If rcTest.IntersectsWith(rcTemp) = True Then
                                        bGood = False
                                        Exit For
                                    End If

                                    If bReverseCheckWpns = True Then
                                        rcTest = goModelDefs.GetModelDef(oFac.EntityDef.ModelID).rcSpacing
                                        rcTest.Width += lTmpRng * gl_FINAL_GRID_SQUARE_SIZE
                                        rcTest.Height += lTmpRng * gl_FINAL_GRID_SQUARE_SIZE
                                        rcTest.Location = New Point(oFac.LocX - (rcTest.Width \ 2), oFac.LocZ - (rcTest.Height \ 2))
                                    End If

                                    If (bCheckWpns = True OrElse bReverseCheckWpns = True) AndAlso bAlertWarDecPot = False AndAlso oFac.Owner.ObjectID <> oUnit.Owner.ObjectID AndAlso rcWpnRng.IntersectsWith(rcTest) = True Then
                                        'MSC - 12/31/07 - Made this change to stop facilities with weapons being built next to other facilities
                                        'Dim yRelScore As Byte = oFac.Owner.GetPlayerRelScore(oUnit.Owner.ObjectID)
                                        'If yRelScore < elRelTypes.ePeace AndAlso yRelScore > elRelTypes.eWar Then
                                        bAlertWarDecPot = True
                                        'End If
                                        bGood = False
                                        Exit For
                                    End If
                                End If
                            Next Y
                            'Now, verify that no units will be a problem
                            If bGood = True Then
                                For Y As Int32 = 0 To glUnitUB
                                    If glUnitIdx(Y) <> -1 AndAlso glUnitIdx(Y) <> glUnitIdx(X) Then
                                        Dim oTmpUnit As Unit = goUnit(Y)
                                        If oTmpUnit Is Nothing Then
                                            'glUnitIdx(Y) = -1
                                            Continue For
                                        End If
                                        If oTmpUnit.ParentObject Is Nothing Then
                                            'glUnitIdx(Y) = -1
                                            Continue For
                                        End If
                                        With CType(oTmpUnit.ParentObject, Epica_GUID)
                                            If .ObjectID <> lTmpPOID OrElse .ObjTypeID <> iTmpPOID Then Continue For
                                        End With

                                        'TODO: Not sure if removing this is a good idea
                                        'Dim rcTest As Rectangle = goModelDefs.GetModelDef(goUnit(Y).EntityDef.ModelID).rcSpacing
                                        'rcTest.Location = New Point(goUnit(Y).LocX - (rcTest.Width \ 2), goUnit(Y).LocZ - (rcTest.Height \ 2))
                                        'If rcTest.IntersectsWith(rcTemp) = True Then
                                        '    bGood = False
                                        '    Exit For
                                        'End If

                                        If oTmpUnit.bProducing = True AndAlso oTmpUnit.CurrentProduction Is Nothing = False Then
                                            If oTmpUnit.CurrentProduction.ProductionTypeID = ObjectType.eFacilityDef Then
                                                'Ok, get that facility's model
                                                Dim oTmpFacDef As FacilityDef = GetEpicaFacilityDef(oTmpUnit.CurrentProduction.ProductionID)
                                                If oTmpFacDef Is Nothing = False Then
                                                    Dim rcTest As Rectangle = goModelDefs.GetModelDef(oTmpFacDef.ModelID).rcSpacing
                                                    rcTest.Location = New Point(oTmpUnit.LocX - (rcTest.Width \ 2), oTmpUnit.LocZ - (rcTest.Height \ 2))
                                                    If rcTest.IntersectsWith(rcTemp) = True Then
                                                        bGood = False
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next Y
                            End If
                        End If

                        If bGood = True Then

                            'Ok, first, check our alert
                            If bAlertWarDecPot = True Then
                                Dim yBuilderAlert(13) As Byte
                                System.BitConverter.GetBytes(GlobalMessageCode.eActOfWarNotice).CopyTo(yBuilderAlert, 0)
                                CType(oUnit.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yBuilderAlert, 2)
                                oUnit.GetGUIDAsString.CopyTo(yBuilderAlert, 8)
                                oUnit.Owner.SendPlayerMessage(yBuilderAlert, False, AliasingRights.eCancelProduction Or AliasingRights.eMoveUnits)
                            End If

                            oUnit.bProducing = False    'tell the unit to stopp producing
                            oUnit.CurrentProduction = Nothing
                            oUnit.CurrentProduction = New EntityProduction()

                            With oUnit.CurrentProduction
                                .oParent = oUnit
                                .ProductionID = lProdID
                                .ProductionTypeID = iProdTypeID
                                .OrderNumber = 1
                                .PointsProduced = 0
                                .lProdCount = 1
                                .lProdX = lLocX
                                .lProdZ = lLocZ
                                .iProdA = iLocA
                                .lLastUpdateCycle = glCurrentCycle
                                .lFinishCycle = 0
                            End With

                            If oUnit.EntityDef.ObjectID = 28344 OrElse oUnit.PopulateRequirements() = True Then         'replace with super engineerid
                                oUnit.bProducing = True
                                AddEntityProducing(X, ObjectType.eUnit, oUnit.ObjectID)
                                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdSucceed).CopyTo(yData, 0)

                                'Need to inform the Region Server
                                Dim iTemp As Int16 = CType(oUnit.ParentObject, Epica_GUID).ObjTypeID
                                If iTemp = ObjectType.ePlanet Then
                                    CType(oUnit.ParentObject, Planet).oDomain.DomainSocket.SendData(yData)
                                ElseIf iTemp = ObjectType.eSolarSystem Then
                                    CType(oUnit.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yData)
                                End If
                            Else
                                If oED Is Nothing = False AndAlso oColony Is Nothing = False Then
                                    If oED.ProductionTypeID = ProductionType.eTradePost Then
                                        oColony.bTradepostInProduction = False
                                    ElseIf oED.ProductionTypeID = ProductionType.eCommandCenterSpecial Then
                                        oColony.bCCInProduction = False
                                    End If
                                End If
                                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)

                                'Now, if this is a rebuilder
                                If oUnit.oRebuilderFor Is Nothing = False Then
                                    oUnit.oRebuilderFor.DoNextRebuildOrder(oUnit)
                                Else
                                    oUnit.bProdQueueMoveSent = False
                                    oUnit.CurrentProduction = Nothing
                                    oUnit.CheckProdQueue()
                                End If

                            End If
                        Else

                            If bAlertWarDecPot = True Then
                                Dim yBuilderAlert(13) As Byte
                                System.BitConverter.GetBytes(GlobalMessageCode.eActOfWarNotice).CopyTo(yBuilderAlert, 0)
                                CType(oUnit.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yBuilderAlert, 2)
                                oUnit.GetGUIDAsString.CopyTo(yBuilderAlert, 8)
                                oUnit.Owner.SendPlayerMessage(yBuilderAlert, False, AliasingRights.eCancelProduction Or AliasingRights.eMoveUnits)
                            End If

                            If oED Is Nothing = False AndAlso oColony Is Nothing = False Then
                                If oED.ProductionTypeID = ProductionType.eTradePost Then
                                    oColony.bTradepostInProduction = False
                                ElseIf oED.ProductionTypeID = ProductionType.eCommandCenterSpecial Then
                                    oColony.bCCInProduction = False
                                End If
                            End If
                            System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)

                            'Now, if this is a rebuilder
                            If oUnit.oRebuilderFor Is Nothing = False Then
                                oUnit.oRebuilderFor.DoNextRebuildOrder(oUnit)
                            Else
                                oUnit.bProdQueueMoveSent = False
                                oUnit.CurrentProduction = Nothing
                                oUnit.CheckProdQueue()
                            End If
                        End If
                    End If
                    Exit For
                End If
            Next X
        Else
            'respond with fail
            System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
        End If

        'send the response 
        If oPlayer Is Nothing = False Then oPlayer.SendPlayerMessage(yData, False, AliasingRights.eAddProduction Or AliasingRights.eCancelProduction)
    End Sub

    Public Function HandleGetEntityProduction(ByVal yData() As Byte) As Byte()
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        Dim X As Int32
        Dim lCnt As Int32 = 0
        Dim lIdx As Int32 = -1

        Dim yResp() As Byte
        Dim lPos As Int32

        'Only units and facilities can produce items
        If iObjTypeID = ObjectType.eFacility Then
            For X = 0 To glFacilityUB
                If glFacilityIdx(X) = lObjID Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then Return Nothing
            Dim oFac As Facility = goFacility(lIdx)
            If oFac Is Nothing Then Return Nothing

            If oFac.CurrentProduction Is Nothing = False Then
                lCnt = 1
            Else : lCnt = 0
            End If

            'Now, go thru and get the data... first, how much data are we talking?
            For X = 0 To oFac.ProductionUB
                If oFac.ProductionIdx(X) <> -1 Then lCnt += 1
            Next X

            'Now, the message is...
            ReDim yResp(21 + (lCnt * 15))

            'Messagecode
            System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityProduction).CopyTo(yResp, 0)
            'EntityID (4), 'TypeID (2)
            oFac.GetGUIDAsString.CopyTo(yResp, 2)
            'CurrentCycle (4)
            System.BitConverter.GetBytes(glCurrentCycle).CopyTo(yResp, 8)
            'QueueCnt (2)
            System.BitConverter.GetBytes(CShort(lCnt)).CopyTo(yResp, 12)

            lPos = 14

            With oFac
                System.BitConverter.GetBytes(.RallyPointX).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(.RallyPointZ).CopyTo(yResp, lPos) : lPos += 4

                'First, do the current production if it is available
                If .CurrentProduction Is Nothing = False Then ' AndAlso .bProducing = True Then
                    System.BitConverter.GetBytes(.CurrentProduction.ProductionID).CopyTo(yResp, lPos)
                    lPos += 4
                    System.BitConverter.GetBytes(.CurrentProduction.ProductionTypeID).CopyTo(yResp, lPos)
                    lPos += 2
                    System.BitConverter.GetBytes(.CurrentProduction.lProdCount).CopyTo(yResp, lPos)
                    lPos += 4
                    If .CurrentProduction.PointsRequired = 0 OrElse .bProducing = False Then yResp(lPos) = CByte(100) Else yResp(lPos) = CByte((.CurrentProduction.PointsProduced / .CurrentProduction.PointsRequired) * 100)
                    lPos += 1
                    System.BitConverter.GetBytes(.CurrentProduction.lFinishCycle).CopyTo(yResp, lPos)
                    lPos += 4
                End If

                For X = 0 To .ProductionUB
                    If .ProductionIdx(X) <> -1 Then
                        '  ProdID (4)
                        System.BitConverter.GetBytes(.Production(X).ProductionID).CopyTo(yResp, lPos)
                        lPos += 4
                        '  ProdTypeID (2)
                        System.BitConverter.GetBytes(.Production(X).ProductionTypeID).CopyTo(yResp, lPos)
                        lPos += 2
                        '  Quantity (2)
                        System.BitConverter.GetBytes(.Production(X).lProdCount).CopyTo(yResp, lPos)
                        lPos += 4
                        '  PercentComplete (1)
                        If .Production(X).PointsRequired = 0 Then
                            yResp(lPos) = 0
                        ElseIf .Production(X).PointsProduced > .Production(X).PointsRequired Then
                            yResp(lPos) = 100
                        Else
                            yResp(lPos) = CByte((.Production(X).PointsProduced / .Production(X).PointsRequired) * 100)
                        End If
                        lPos += 1
                        '  FinishCycle (4)
                        System.BitConverter.GetBytes(.Production(X).lFinishCycle).CopyTo(yResp, lPos)
                        lPos += 4
                    End If
                Next X
            End With

            'send our response
            'If oOwner Is Nothing = False Then oOwner.SendPlayerMessage(yResp, False, AliasingRights.eViewUnitsAndFacilities)
            'If oSocket Is Nothing = False Then oSocket.SendData(yResp)
            Return yResp

        ElseIf iObjTypeID = ObjectType.eUnit Then
            For X = 0 To glUnitUB
                If glUnitIdx(X) = lObjID Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then Return Nothing
            Dim oUnit As Unit = goUnit(lIdx)
            If oUnit Is Nothing Then Return Nothing

            ReDim yResp(34)

            System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityProduction).CopyTo(yResp, 0)
            'EntityID (4), 'TypeID (2)
            oUnit.GetGUIDAsString.CopyTo(yResp, 2)
            'CurrentCycle (4)
            System.BitConverter.GetBytes(glCurrentCycle).CopyTo(yResp, 8)
            If oUnit.bProducing = False OrElse oUnit.CurrentProduction Is Nothing Then
                System.BitConverter.GetBytes(CShort(0)).CopyTo(yResp, 12)
                System.BitConverter.GetBytes(oUnit.RallyPointX).CopyTo(yResp, 14)
                System.BitConverter.GetBytes(oUnit.RallyPointZ).CopyTo(yResp, 18)
            Else
                With oUnit.CurrentProduction
                    'QueueCnt (2)
                    System.BitConverter.GetBytes(CShort(1)).CopyTo(yResp, 12)
                    lPos = 14
                    System.BitConverter.GetBytes(oUnit.RallyPointX).CopyTo(yResp, lPos) : lPos += 4
                    System.BitConverter.GetBytes(oUnit.RallyPointZ).CopyTo(yResp, lPos) : lPos += 4
                    '  ProdID (4)
                    System.BitConverter.GetBytes(.ProductionID).CopyTo(yResp, lPos)
                    lPos += 4
                    '  ProdTypeID (2)
                    System.BitConverter.GetBytes(.ProductionTypeID).CopyTo(yResp, lPos)
                    lPos += 2
                    '  Quantity (2)
                    System.BitConverter.GetBytes(CShort(.lProdCount)).CopyTo(yResp, lPos)
                    lPos += 2
                    '  PercentComplete (1)
                    yResp(lPos) = CByte((.PointsProduced / .PointsRequired) * 100)
                    lPos += 1
                    '  FinishCycle (4)
                    System.BitConverter.GetBytes(.lFinishCycle).CopyTo(yResp, lPos)
                    lPos += 4
                End With
            End If

            'If oOwner Is Nothing = False Then oOwner.SendPlayerMessage(yResp, False, 0)
            'If oSocket Is Nothing = False Then oSocket.SendData(yResp)
            Return yResp
        End If
        Return Nothing
    End Function

    Private Sub HandleOperatorPlayerCurrEnvir(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        For X As Int32 = 0 To mlClientUB
            If mlClientPlayer(X) = lPlayerID Then
                moClients(X).SendData(yData)
                Exit For
            End If
        Next X
    End Sub
    Private Function GetPlayerEnvirMsg(ByVal lPlayerID As Int32, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16) As Byte()
        Dim yResp(24) As Byte

        Dim X As Int32
        Dim lPlanetID As Int32 = -1
        Dim lSystemID As Int32 = -1
        Dim lGalaxyID As Int32 = -1

        System.BitConverter.GetBytes(GlobalMessageCode.ePlayerCurrentEnvironment).CopyTo(yResp, 0)
        System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, 2)
        System.BitConverter.GetBytes(lEnvirID).CopyTo(yResp, 6)
        System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yResp, 10)

        'Now, append the geography tree to it... geography tree is the following...
        '  Galaxy, System, Planet... we always send all three and the client will use what it needs to
        Dim bInMyDomain As Boolean = False
        Select Case iEnvirTypeID
            Case ObjectType.ePlanet
                lPlanetID = lEnvirID
                For X = 0 To glPlanetUB
                    If glPlanetIdx(X) = lEnvirID Then
                        'ok, now get the system id
                        bInMyDomain = goPlanet(X).InMyDomain
                        If bInMyDomain = False Then goPlanet(X).InMyDomain = True
                        If goPlanet(X).ParentSystem Is Nothing = False Then
                            lSystemID = goPlanet(X).ParentSystem.ObjectID
                            'now, get the galaxy id
                            If goPlanet(X).ParentSystem.ParentGalaxy Is Nothing = False Then
                                lGalaxyID = goPlanet(X).ParentSystem.ParentGalaxy.ObjectID
                            End If
                        End If
                        Exit For
                    End If
                Next X
            Case ObjectType.eSolarSystem
                lPlanetID = -1
                lSystemID = lEnvirID
                For X = 0 To glSystemUB
                    If glSystemIdx(X) = lEnvirID Then
                        'ok, get the galaxy
                        bInMyDomain = goSystem(X).InMyDomain
                        If bInMyDomain = False Then goSystem(X).InMyDomain = True
                        If goSystem(X).ParentGalaxy Is Nothing = False Then
                            lGalaxyID = goSystem(X).ParentGalaxy.ObjectID
                        End If
                        Exit For
                    End If
                Next X
        End Select
        bInMyDomain = True

        'k, now append that data
        System.BitConverter.GetBytes(lGalaxyID).CopyTo(yResp, 12)
        System.BitConverter.GetBytes(lSystemID).CopyTo(yResp, 16)
        System.BitConverter.GetBytes(lPlanetID).CopyTo(yResp, 20)

        If bInMyDomain = False Then
            yResp(24) = 0
            moOperator.SendData(yResp)
            Return Nothing
        Else : yResp(24) = 1
        End If

        Return yResp
    End Function

    Public Function GetAddObjectMessage(ByVal oObj As Object, ByVal iMsgCode As Int16) As Byte()
        Dim yObj() As Byte
        Dim yFinal() As Byte
        Dim iLen As Int32

        Try

            yObj = CType(CallByName(oObj, "GetObjAsString", CallType.Get), Byte()) 'oObj.GetObjAsString()
            iLen = yObj.Length + 1
            ReDim yFinal(iLen)
            System.BitConverter.GetBytes(iMsgCode).CopyTo(yFinal, 0)
            yObj.CopyTo(yFinal, 2)

            Return yFinal
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "GetAddObjectMessage: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function GetUndockCommandMsg(ByRef oEntity As Unit, ByVal lFinalParentID As Int32, ByVal iFinalParentTypeID As Int16, ByVal lBackupLocX As Int32, ByVal lBackupLocZ As Int32) As Byte()
        Dim yObj() As Byte
        Dim yFinal() As Byte
        Dim iLen As Int16

        Try
            Dim lPos As Int32
            yObj = oEntity.GetObjAsString()
            iLen = CShort(yObj.Length + 15)
            ReDim yFinal(iLen)
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestUndock).CopyTo(yFinal, 0)
            lPos = 2
            yObj.CopyTo(yFinal, 2) : lPos += yObj.Length
            System.BitConverter.GetBytes(lFinalParentID).CopyTo(yFinal, lPos) : lPos += 4
            System.BitConverter.GetBytes(iFinalParentTypeID).CopyTo(yFinal, lPos) : lPos += 2
            System.BitConverter.GetBytes(lBackupLocX).CopyTo(yFinal, lPos) : lPos += 4
            System.BitConverter.GetBytes(lBackupLocZ).CopyTo(yFinal, lPos) : lPos += 4

            Return yFinal
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "GetUndockCommandMsg: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function GetAddPlayerRelMessage(ByVal oPRel As PlayerRel) As Byte()
        Dim yObj() As Byte
        Dim yFinal() As Byte
        Dim iLen As Int16

        yObj = oPRel.GetObjAsString()
        iLen = CShort(yObj.Length + 1)
        ReDim yFinal(iLen)
        System.BitConverter.GetBytes(GlobalMessageCode.eAddPlayerRelObject).CopyTo(yFinal, 0)
        yObj.CopyTo(yFinal, 2)
        Return yFinal
    End Function

    Private Sub HandleRequestStarTypes(ByVal Index As Int32, ByVal lConnType As ConnectionType)
        'ok, socket @ Index is requestin the star type list
        Dim X As Int32
        Dim lPos As Int32 = 0
        Dim lOneLen As Int32

        If gbStarTypeMsgReady = False Then
            lOneLen = goStarType(0).GetObjAsString.Length
            ReDim gyStarTypeMsg((lOneLen * (glStarTypeUB + 1)) + 1)     'should be -1 but we need 2 for the msg ID
            System.BitConverter.GetBytes(GlobalMessageCode.eResponseStarTypes).CopyTo(gyStarTypeMsg, 0)
            lPos += 2
            For X = 0 To glStarTypeUB
                goStarType(X).GetObjAsString.CopyTo(gyStarTypeMsg, lPos)
                lPos += lOneLen
            Next X
            gbStarTypeMsgReady = True
        End If

        Select Case lConnType
            Case ConnectionType.eClientConnection
                moClients(Index).SendData(gyStarTypeMsg)
            Case ConnectionType.eRegionServerApp
                moDomains(Index).SendData(gyStarTypeMsg)
            Case ConnectionType.ePathfindingServerApp
                moPathfinding.SendData(gyStarTypeMsg)
        End Select

    End Sub

    Private Sub HandleRequestGalaxy(ByVal Index As Int32, ByVal lConnType As ConnectionType)
        'ok, socket @ index is requesting the galaxy and system list
        Dim lTotalLen As Int32
        Dim lGalLen As Int32
        Dim lSysLen As Int32
        Dim lPos As Int32
        Dim X As Int32

        If lConnType = ConnectionType.eClientConnection Then
            If (mlClientStatusFlags(Index) And ClientStatusFlags.eRequestedGalaxyAndSystems) <> 0 Then
                LogEvent(LogEventType.PossibleCheat, "Multiple RequestGalaxy requests from client. Player: " & mlClientPlayer(Index))
                Return
            Else
                mlClientStatusFlags(Index) = mlClientStatusFlags(Index) Or ClientStatusFlags.eRequestedGalaxyAndSystems
            End If
        End If

        If gbGalaxyMsgReady = False Then
            'should be -1 but we need 2 for the message ID
            lGalLen = goGalaxy(0).GetObjAsString.Length
            lSysLen = goSystem(0).GetObjAsString.Length
            lTotalLen = lGalLen + (lSysLen * (glSystemUB + 1)) + 1
            ReDim gyGalaxySystemMsg(lTotalLen)
            System.BitConverter.GetBytes(GlobalMessageCode.eResponseGalaxyAndSystems).CopyTo(gyGalaxySystemMsg, 0)
            lPos += 2

            'now the galaxy, NOTE: this only handles one galaxy right now... I will want to expand on this eventually
            goGalaxy(0).GetObjAsString.CopyTo(gyGalaxySystemMsg, lPos)
            lPos += lGalLen

            'Now the systems
            For X = 0 To glSystemUB
                goSystem(X).GetObjAsString.CopyTo(gyGalaxySystemMsg, lPos)
                lPos += lSysLen
            Next X
        End If

        'Now send it
        Select Case lConnType
            Case ConnectionType.eClientConnection
                moClients(Index).SendData(gyGalaxySystemMsg)
            Case ConnectionType.eRegionServerApp
                moDomains(Index).SendData(gyGalaxySystemMsg)

                For X = 0 To glWormholeUB
                    If glWormholeIdx(X) > -1 Then
                        moDomains(Index).SendData(GetAddObjectMessage(goWormhole(X), GlobalMessageCode.eAddObjectCommand))
                    End If
                Next X
            Case ConnectionType.ePathfindingServerApp
                moPathfinding.SendData(gyGalaxySystemMsg)
        End Select

    End Sub

    Private Sub HandleRequestSystemDetails(ByVal lSocketIndex As Int32, ByVal lConnType As ConnectionType, ByVal yData() As Byte)
        Dim lSysID As Int32
        Dim X As Int32
        Dim lIdx As Int32
        Dim yResp() As Byte
        'Dim lLen As Int32

        'Ok, the message coming to us is:
        ' MsgID (2), SystemID (4)
        lIdx = -1
        lSysID = System.BitConverter.ToInt32(yData, 2)
        For X = 0 To glSystemUB
            If glSystemIdx(X) = lSysID Then
                'right now, just send back planets...
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            'send back a negative response
            System.BitConverter.GetBytes(GlobalMessageCode.eResponseSystemDetailsFail).CopyTo(yData, 0)
            Select Case lConnType
                Case ConnectionType.eClientConnection
                    moClients(lSocketIndex).SendData(yData)
                Case ConnectionType.eRegionServerApp
                    moDomains(lSocketIndex).SendData(yData)
                Case ConnectionType.ePathfindingServerApp
                    moPathfinding.SendData(yData)
            End Select
        Else
            'send back our data
            'lLen = goSystem(lIdx).GetDetailsStringLen() + 2     '+2 for our msg ID
            'ReDim yResp(lLen)
            'System.BitConverter.GetBytes(GlobalMessageCode.eResponseSystemDetails).CopyTo(yResp, 0)
            'goSystem(lIdx).GetDetailsString.CopyTo(yResp, 2)
            yResp = goSystem(lIdx).GetSystemDetailsResponse()

            'and send it
            Select Case lConnType
                Case ConnectionType.eClientConnection
                    moClients(lSocketIndex).SendData(yResp)
                Case ConnectionType.eRegionServerApp
                    System.GC.Collect()
                    moDomains(lSocketIndex).SendData(yResp)

                Case ConnectionType.ePathfindingServerApp
                    moPathfinding.SendData(yResp)
            End Select
        End If
    End Sub

    Public Sub HandleSetPlayerRel(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, 6)
        Dim yRel As Byte = yData(10)

        If lPlayerID = lTargetID Then Return

        If lIndex <> -1 AndAlso lPlayerID <> mlClientPlayer(lIndex) AndAlso (mlAliasedAs(lIndex) <> lPlayerID OrElse (mlAliasedRights(lIndex) And AliasingRights.eAlterDiplomacy) = 0) Then
            LogEvent(LogEventType.PossibleCheat, "PlayerID does not match socket's player ID for SetPlayerRel: PlayerID = " & lPlayerID & ", SocketPlayerID = " & mlClientPlayer(lIndex))
            Return
        End If

        Dim oInitiator As Player = GetEpicaPlayer(lPlayerID)
        Dim oTarget As Player = GetEpicaPlayer(lTargetID)

        If oInitiator Is Nothing OrElse oTarget Is Nothing Then Return

        'Ok, check guilds
        If oInitiator.oGuild Is Nothing = False AndAlso oTarget.oGuild Is Nothing = False Then
            'ok... check if initiator is in a guild with unified rels
            'If (oInitiator.oGuild.lBaseGuildFlags And elGuildFlags.UnifiedForeignPolicy) <> 0 Then Return
            'ok, are members in same guild?
            If oInitiator.lGuildID = oTarget.lGuildID Then
                'yes, is relationship war?
                If yRel <= elRelTypes.eWar Then
                    'yes, does guild permit war?
                    If (oInitiator.oGuild.lBaseGuildFlags And elGuildFlags.RequirePeaceBetweenMembers) <> 0 Then Return
                End If
            End If
        End If

        'Ok, now, let's do this...
        Dim yCurrent As Byte = oInitiator.GetPlayerRelScore(lTargetID)
        Dim yInitiatorOriginal As Byte = yCurrent

        Dim lCurrentNextVal As Int32 = -1
        With oInitiator
            Dim oRel As PlayerRel = .GetPlayerRel(lTargetID)
            If oRel Is Nothing = False Then lCurrentNextVal = oRel.lCycleOfNextScoreUpdate
        End With

        'Ok, new rules... if the new rel score < current relationship score then we need to queue it up
        If yRel < yCurrent Then
            Dim oRel As PlayerRel
            oRel = oInitiator.GetPlayerRel(lTargetID)
            If oRel Is Nothing Then Return


            Dim lNextVal As Int32 = yCurrent
            If oInitiator.oGuild Is Nothing = False Then
                For X As Int32 = 0 To oInitiator.oGuild.lMemberUB
                    Dim oMem As GuildMember = oInitiator.oGuild.oMembers(X)
                    If oMem Is Nothing = False AndAlso (oMem.yMemberState And GuildMemberState.Approved) <> 0 Then
                        If oMem.oMember Is Nothing = False Then
                            Dim yGldRel As Byte = oMem.oMember.GetPlayerRelScore(lTargetID)
                            If yGldRel <= yCurrent Then
                                lNextVal = Math.Min(lNextVal, yGldRel)  'Math.Max(oRel.TargetScore, yGldRel)
                            End If
                        End If
                    End If
                Next X
            End If
            If oTarget.oGuild Is Nothing = False Then
                For X As Int32 = 0 To oTarget.oGuild.lMemberUB
                    Dim oMem As GuildMember = oTarget.oGuild.oMembers(X)
                    If oMem Is Nothing = False Then
                        Dim yGldRel As Byte = oMem.oMember.GetPlayerRelScore(lPlayerID)
                        If yGldRel <= yCurrent Then
                            lNextVal = Math.Min(lNextVal, yGldRel)
                        End If
                    End If
                Next X
            End If
            For X As Int32 = 0 To 4
                Dim lTmp As Int32 = oInitiator.lSlotID(X)
                If lTmp > -1 Then
                    Dim oFaction As Player = GetEpicaPlayer(lTmp)
                    If oFaction Is Nothing = False Then
                        Dim yFacRel As Byte = oFaction.GetPlayerRelScore(lTargetID)
                        If yFacRel <= yCurrent Then
                            lNextVal = Math.Min(lNextVal, yFacRel)
                        End If
                    End If
                End If
                lTmp = oTarget.lSlotID(X)
                If lTmp > -1 Then
                    Dim oFaction As Player = GetEpicaPlayer(lTmp)
                    If oFaction Is Nothing = False Then
                        Dim yFacRel As Byte = oFaction.GetPlayerRelScore(lPlayerID)
                        If yFacRel <= yCurrent Then
                            lNextVal = Math.Min(lNextVal, yFacRel)
                        End If
                    End If
                End If
            Next X
            For X As Int32 = 0 To 2
                Dim lTmp As Int32 = oInitiator.lFactionID(X)
                If lTmp > -1 Then
                    Dim oFaction As Player = GetEpicaPlayer(lTmp)
                    If oFaction Is Nothing = False Then
                        Dim yFacRel As Byte = oFaction.GetPlayerRelScore(lTargetID)
                        If yFacRel <= yCurrent Then
                            lNextVal = Math.Min(lNextVal, yFacRel)
                        End If
                    End If
                End If
                lTmp = oTarget.lFactionID(X)
                If lTmp > -1 Then
                    Dim oFaction As Player = GetEpicaPlayer(lTmp)
                    If oFaction Is Nothing = False Then
                        Dim yFacRel As Byte = oFaction.GetPlayerRelScore(lPlayerID)
                        If yFacRel <= yCurrent Then
                            lNextVal = Math.Min(lNextVal, yFacRel)
                        End If
                    End If
                End If
            Next X

            Dim bSendMsg As Boolean = False
            If lNextVal < yRel Then lNextVal = yRel
            If lNextVal < yCurrent Then
                oRel.WithThisScore = CByte(lNextVal)
                If yCurrent > elRelTypes.eWar AndAlso oTarget.GetPlayerRelScore(oInitiator.ObjectID) > elRelTypes.eWar AndAlso lNextVal <= elRelTypes.eWar Then
                    ProcessWarDeclared(oInitiator, oTarget, lNextVal)
                End If
                bSendMsg = True
            End If

            oRel.TargetScore = yRel
            If oRel.lCycleOfNextScoreUpdate <= glCurrentCycle Then
                If bSendMsg = True Then
                    oRel.lCycleOfNextScoreUpdate = glCurrentCycle + gl_DELAY_FOUR_HOURS
                Else
                    oRel.lCycleOfNextScoreUpdate = glCurrentCycle + 1
                End If
                AddToQueue(oRel.lCycleOfNextScoreUpdate, QueueItemType.eReprocessPlayerRel, lPlayerID, -1, lTargetID, -1, -1, -1, -1, -1)
            End If
            oRel.SaveObject()

            If bSendMsg = True Then
                Dim yNameForward(27) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityName).CopyTo(yNameForward, 0)
                oInitiator.GetGUIDAsString.CopyTo(yNameForward, 2)
                oInitiator.PlayerName.CopyTo(yNameForward, 8)
                oTarget.SendPlayerMessage(yNameForward, False, 0)

                Dim yTempRel() As Byte = GetAddPlayerRelMessage(oRel)
                System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerRel).CopyTo(yTempRel, 0)
                oInitiator.SendPlayerMessage(yTempRel, False, 0)
                oTarget.SendPlayerMessage(yTempRel, True, 0)
            End If

            Return
        End If

        'yCurrent = Math.Min(yCurrent, oTarget.GetPlayerRelScore(lPlayerID))



        If yCurrent > elRelTypes.eWar AndAlso yRel <= elRelTypes.eWar Then
            ''handle War Sentiment... war is being declared, so if either player doesn't want war...
            ''If oInitiator.lWarSentiment < 0 Then oInitiator.lWarSentiment -= 10
            ''If oTarget.lWarSentiment < 0 Then oTarget.lWarSentiment -= 10

            'oTarget.bDeclaredWarOn = True

            ''Ok, we are declaring war... let's do it cascadingly
            'Dim oCascadeWarDec As New CascadeWarDec
            'oCascadeWarDec.HandleCascadeWarDec(oInitiator, oTarget, yRel)
            'oCascadeWarDec = Nothing
        Else
            'Ok, normal set rel
            Dim oTmpRel As PlayerRel = New PlayerRel
            oTmpRel.oPlayerRegards = oInitiator
            oTmpRel.oThisPlayer = oTarget
            oTmpRel.WithThisScore = yRel
            oTmpRel.TargetScore = yRel
            oTmpRel.lCycleOfNextScoreUpdate = lCurrentNextVal
            oInitiator.SetPlayerRel(lTargetID, oTmpRel, True)

            oInitiator.SendPlayerMessage(yData, False, AliasingRights.eViewDiplomacy Or AliasingRights.eViewAgents)

            Dim bSendEmail As Boolean = False

            Dim yCurRelTypeVal As Byte
            Dim yNewRelTypeVal As Byte

            If yCurrent <= elRelTypes.eWar Then
                yCurRelTypeVal = 0
            ElseIf yCurrent < elRelTypes.eNeutral Then
                yCurRelTypeVal = 1
            ElseIf yCurrent < elRelTypes.ePeace Then
                yCurRelTypeVal = 2
            Else : yCurRelTypeVal = 3
            End If
            If yRel <= elRelTypes.eWar Then
                yNewRelTypeVal = 0
            ElseIf yRel < elRelTypes.eNeutral Then
                yNewRelTypeVal = 1
            ElseIf yRel < elRelTypes.ePeace Then
                yNewRelTypeVal = 2
            Else : yNewRelTypeVal = 3
            End If

            bSendEmail = yCurRelTypeVal <> yNewRelTypeVal
            oTarget.SendPlayerMessage(yData, bSendEmail, AliasingRights.eViewDiplomacy Or AliasingRights.eViewAgents)

            'Ok, if my enemy is at war with me, but I am setting my rel to non-war with them, we do not send.
            Dim bDoNotSend As Boolean = False
            Dim bSendTargetsRel As Boolean = True
            'If yRel > elRelTypes.eWar Then
            '    If oTarget.GetPlayerRelScore(oInitiator.ObjectID) <= elRelTypes.eWar Then
            '        'bDoNotSend = True          'MSC - 12/22/08 - always send
            '    Else
            '        'Now, if my rel WAS war, and the target has me at non-war, and my new rel is non-war... send the other player's rel msg too
            '        If yInitiatorOriginal <= elRelTypes.eWar Then
            '            bSendTargetsRel = True
            '        End If
            '    End If

            'End If

            'Now, broadcast to domains...
            If bDoNotSend = False Then
                For Y As Int32 = 0 To mlDomainUB
                    moDomains(Y).SendData(yData)
                Next Y
            End If
            If bSendTargetsRel = True Then
                Dim yTemp(10) As Byte

                System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerRel).CopyTo(yTemp, 0)
                System.BitConverter.GetBytes(oTarget.ObjectID).CopyTo(yTemp, 2)
                System.BitConverter.GetBytes(oInitiator.ObjectID).CopyTo(yTemp, 6)
                yTemp(10) = oTarget.GetPlayerRelScore(oInitiator.ObjectID)

                For Y As Int32 = 0 To mlDomainUB
                    moDomains(Y).SendData(yTemp)
                Next Y
            End If

            If (yCurrent < elRelTypes.eAlly AndAlso yRel >= elRelTypes.eAlly) OrElse (yCurrent < elRelTypes.ePeace AndAlso yRel >= elRelTypes.ePeace) Then
                oInitiator.TestCustomTitlePermissions_Allies()
                oTarget.TestCustomTitlePermissions_Allies()
            End If

            'If yCurrent < elRelTypes.eWar AndAlso yRel > elRelTypes.eWar Then
            'If oInitiator.lWarSentiment > 0 Then oInitiator.lWarSentiment += 10
            'End If
        End If

        oInitiator.ReverifySlots()
        oTarget.ReverifySlots()


        'For X As Int32 = 0 To glPlayerUB
        '    If glPlayerIdx(X) = lPlayerID Then

        '        'Check the target player...
        '        For Y As Int32 = 0 To glPlayerUB
        '            If glPlayerIdx(Y) = lTargetID Then
        '                Dim oTmpRel As PlayerRel = New PlayerRel()
        '                oTmpRel.oPlayerRegards = goPlayer(X)
        '                oTmpRel.oThisPlayer = goPlayer(Y)
        '                oTmpRel.WithThisScore = yRel
        '                goPlayer(X).SetPlayerRel(lTargetID, oTmpRel)
        '                If yRel < elRelTypes.eWar Then
        '                    'Because one player declared war, both do...
        '                    oTmpRel = Nothing
        '                    oTmpRel = New PlayerRel()
        '                    oTmpRel.oPlayerRegards = goPlayer(Y)
        '                    oTmpRel.oThisPlayer = goPlayer(X)
        '                    oTmpRel.WithThisScore = yRel
        '                    goPlayer(Y).SetPlayerRel(lPlayerID, oTmpRel)

        '                    'And subsequently the allies of each declare war

        '                End If

        '                'Now, send the message to the client if it is connected
        '                If goPlayer(Y).oSocket Is Nothing = False Then
        '                    goPlayer(Y).oSocket.SendData(yData)
        '                End If

        '                Exit For

        '            End If
        '        Next Y

        '        'Now, broadcast to domains...
        '        For Y As Int32 = 0 To mlDomainUB
        '            moDomains(Y).SendData(yData)
        '        Next Y

        '        Exit For
        '    End If
        'Next X
    End Sub

    Private Function HandleRequestProductionStatus(ByVal yData() As Byte) As Byte() ' ByRef oOwner As Player)
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim oEntity As Epica_Entity = CType(GetEpicaObject(lID, iTypeID), Epica_Entity)

        'Dim yResp(22) As Byte
        Dim yResp(29) As Byte
        Dim yPerc As Byte

        'TODO: Ensure OwnerID = ID of the player for the socket

        On Error Resume Next

        If oEntity Is Nothing = False Then
            System.BitConverter.GetBytes(GlobalMessageCode.eEntityProductionStatus).CopyTo(yResp, 0)
            oEntity.GetGUIDAsString.CopyTo(yResp, 2)
            If oEntity.bProducing = False OrElse oEntity.CurrentProduction Is Nothing Then
                'MSC - added to make the production status not jump around so much...
                'yResp(8) = 255     'client will know how to display this
                'System.BitConverter.GetBytes(CInt(-1)).CopyTo(yResp, 13)
                'System.BitConverter.GetBytes(CInt(-1)).CopyTo(yResp, 17)
                'System.BitConverter.GetBytes(CShort(-1)).CopyTo(yResp, 21)

                System.BitConverter.GetBytes(CSng(255)).CopyTo(yResp, 8)
                System.BitConverter.GetBytes(CInt(-1)).CopyTo(yResp, 16)
                System.BitConverter.GetBytes(CInt(-1)).CopyTo(yResp, 20)
                System.BitConverter.GetBytes(CShort(-1)).CopyTo(yResp, 24)
                System.BitConverter.GetBytes(CShort(-1)).CopyTo(yResp, 26)
                System.BitConverter.GetBytes(CShort(-1)).CopyTo(yResp, 28)
            Else
                Dim blTemp As Int64
                With oEntity.CurrentProduction

                    If oEntity.ObjTypeID = ObjectType.eFacility Then CType(oEntity, Facility).RecalcProduction()

                    Dim blPointsProd As Int64 = .PointsProduced
                    Dim lLastUpdate As Int32 = .lLastUpdateCycle
                    Dim lProd As Int32 = oEntity.mlProdPoints
                    Dim lFinish As Int32 = .lFinishCycle



                    'Is this a research facility?
                    If oEntity.yProductionType = ProductionType.eResearch Then
                        'Ok, get the tech
                        Dim oTech As Epica_Tech = oEntity.Owner.GetTech(.ProductionID, .ProductionTypeID)
                        If oTech Is Nothing = False Then
                            oTech.FillPrimarysProductionStatus(blPointsProd, lLastUpdate, lProd, lFinish)
                        End If
                    End If

                    If blPointsProd = 0 AndAlso lLastUpdate = 0 Then
                        oEntity.CurrentProduction = Nothing
                        oEntity.bProducing = False
                        Return Nothing
                    End If

                    'lTemp = .PointsProduced + ((glCurrentCycle - .lLastUpdateCycle) * oEntity.mlProdPoints)
                    blTemp = blPointsProd + ((glCurrentCycle - lLastUpdate) * lProd)

                    Dim blRequired As Int64 = .PointsRequired
                    If oEntity.yProductionType = ProductionType.eResearch Then
                        blRequired = CLng(.PointsRequired * oEntity.Owner.fFactionResearchTimeMultiplier)
                    End If

                    Dim fTmp As Single = CSng((blTemp / blRequired) * 100)
                    System.BitConverter.GetBytes(fTmp).CopyTo(yResp, 8)
                    'System.BitConverter.GetBytes(.lFinishCycle).CopyTo(yResp, 16)
                    System.BitConverter.GetBytes(lFinish).CopyTo(yResp, 16)
                    System.BitConverter.GetBytes(.ProductionID).CopyTo(yResp, 20)
                    System.BitConverter.GetBytes(.ProductionTypeID).CopyTo(yResp, 24)
                    System.BitConverter.GetBytes(.ProductionItemModelID).CopyTo(yResp, 26)
                    System.BitConverter.GetBytes(.iProdA).CopyTo(yResp, 28)
                End With
            End If
            'System.BitConverter.GetBytes(glCurrentCycle).CopyTo(yResp, 9)
            System.BitConverter.GetBytes(glCurrentCycle).CopyTo(yResp, 12)

            'If oSocket Is Nothing = False Then oSocket.SendData(yResp)
            Return yResp
        End If
        Return Nothing
    End Function

    Public Shared Function CreateChatMsg(ByVal lFromPlayer As Int32, ByVal sMsg As String, ByVal yType As ChatMessageType) As Byte()
        Dim yMsg() As Byte

        ReDim yMsg(10 + sMsg.Length)
        System.BitConverter.GetBytes(GlobalMessageCode.eChatMessage).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(lFromPlayer).CopyTo(yMsg, 2)
        yMsg(6) = yType
        System.BitConverter.GetBytes(sMsg.Length).CopyTo(yMsg, 7)
        StringToBytes(sMsg).CopyTo(yMsg, 11)
        Return yMsg
    End Function

    Public Shared Function GetPathfindingNewEntityMsg(ByRef oObjGUID As Epica_GUID, ByRef oParentGUID As Epica_GUID, ByVal lX As Int32, ByVal lZ As Int32, ByVal iModelID As Int16) As Byte()
        Dim yMsg(23) As Byte

        If iModelID = 0 Then Return Nothing
        If oParentGUID Is Nothing Then Return Nothing

        System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, 0)
        oObjGUID.GetGUIDAsString.CopyTo(yMsg, 2)
        oParentGUID.GetGUIDAsString.CopyTo(yMsg, 8)
        System.BitConverter.GetBytes(lX).CopyTo(yMsg, 14)
        System.BitConverter.GetBytes(lZ).CopyTo(yMsg, 18)
        System.BitConverter.GetBytes(iModelID).CopyTo(yMsg, 22)
        Return yMsg
    End Function
    Private mlClearPlayerCommSentBys As Int32 = -1
    Private Sub BeginClearPlayerCommSentBys()
        Dim lID As Int32 = mlClearPlayerCommSentBys
        mlClearPlayerCommSentBys = -1
        ClearPlayerCommSentBys(lID)
    End Sub
    Private Sub ClearPlayerCommSentBys(ByVal lID As Int32)
        Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand("DELETE FROM tblTradeHistory WHERE OtherPlayerID = " & lID, goCN)
        oComm.ExecuteNonQuery()
        oComm.Dispose()
        oComm = Nothing

        goGTC.ClearOrdersAboutIntelOfPlayer(lID)

        For lTemp As Int32 = 0 To glPlayerUB
            If glPlayerIdx(lTemp) > -1 Then
                Dim oTmpP As Player = goPlayer(lTemp)
                If oTmpP Is Nothing = False Then
                    For lEF As Int32 = 0 To oTmpP.EmailFolderUB
                        If oTmpP.EmailFolders(lEF) Is Nothing = False Then
                            Dim oEF As PlayerCommFolder = oTmpP.EmailFolders(lEF)
                            If oEF Is Nothing = False Then
                                For lPC As Int32 = 0 To oEF.PlayerMsgsUB
                                    If oEF.PlayerMsgsIdx(lPC) > -1 Then
                                        Dim oPC As PlayerComm = oEF.PlayerMsgs(lPC)
                                        If oPC Is Nothing = False Then
                                            If oPC.SentByID = lID Then
                                                oPC.SentByID = oPC.PlayerID
                                                oPC.DataChanged()
                                                oPC.SaveObject(oTmpP.ObjectID)
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Next

                    For lPII As Int32 = 0 To oTmpP.mlItemIntelUB
                        Dim oPII As PlayerItemIntel = oTmpP.moItemIntel(lPII)
                        If oPII Is Nothing = False Then
                            If oPII.lOtherPlayerID = lID Then
                                oPII.yArchived = 255
                                oPII.SaveObject()
                            End If
                        End If
                    Next lPII

                    For lPTK As Int32 = 0 To oTmpP.mlPlayerTechKnowledgeUB
                        Dim oPTK As PlayerTechKnowledge = oTmpP.moPlayerTechKnowledge(lPTK)
                        If oPTK Is Nothing = False AndAlso oPTK.oTech Is Nothing = False AndAlso oPTK.oTech.Owner Is Nothing = False AndAlso oPTK.oTech.Owner.ObjectID = lID Then
                            oPTK.yArchived = 255
                            oPTK.SaveObject()
                        End If
                    Next lPTK

                    For lPI As Int32 = 0 To oTmpP.mlPlayerIntelUB
                        Dim oPI As PlayerIntel = oTmpP.moPlayerIntel(lPI)
                        If oPI Is Nothing = False AndAlso oPI.oTarget.ObjectID = lID Then
                            oPI.DiplomacyUpdate = 0
                            oPI.MilitaryUpdate = 0
                            oPI.PopulationUpdate = 0
                            oPI.ProductionUpdate = 0
                            oPI.TechnologyUpdate = 0
                            oPI.WealthUpdate = 0
                            oPI.SaveObject()
                        End If
                    Next lPI


                    oTmpP.RemoveTradeHistoryItemsOfPlayer(lID)
                End If
            End If
        Next lTemp

        For lTemp As Int32 = 0 To glPlayerMissionUB
            If glPlayerMissionIdx(lTemp) > -1 Then
                Dim oPM As PlayerMission = goPlayerMission(lTemp)
                If oPM Is Nothing = False Then
                    If oPM.oTarget Is Nothing = False Then
                        If oPM.oTarget.ObjectID = lID Then
                            oPM.oTarget = Nothing
                            oPM.SaveObject()
                        End If
                    End If
                End If
            End If
        Next lTemp
    End Sub
    Private Sub HandlePlayerInitialSetup(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim yUserName(19) As Byte
        Dim yPassword(19) As Byte

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
                            Dim yEmailAddress(254) As Byte

                            Array.Copy(yData, lPos, yPlayerName, 0, 20) : lPos += 20
                            Array.Copy(yData, lPos, yEmpireName, 0, 20) : lPos += 20
                            yGender = yData(lPos) : lPos += 1
                            lIcon = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                            Array.Copy(yData, lPos, yNewUserName, 0, 20) : lPos += 20
                            Array.Copy(yData, lPos, yNewPassword, 0, 20) : lPos += 20
                            Array.Copy(yData, lPos, yEmailAddress, 0, 255) : lPos += 255

                            Dim yResp(2) As Byte
                            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerInitialSetup).CopyTo(yResp, 0)
                            yResp(2) = 0

                            If lIcon = 0 Then
                                yResp(2) = 3
                            Else
                                Try
                                    Dim sPlayerName As String = BytesToString(yPlayerName)
                                    Dim sEmpireName As String = BytesToString(yEmpireName)
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

                            If yResp(2) = 0 Then
                                .yGender = yGender
                                .lPlayerIcon = lIcon
                                .PlayerName = yPlayerName
                                .sPlayerNameProper = BytesToString(yPlayerName)
                                .sPlayerName = .sPlayerNameProper.ToUpper
                                .EmpireName = yEmpireName

                                If .yPlayerPhase = eyPlayerPhase.eFullLivePhase Then
                                    Dim yMsg(25) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityName).CopyTo(yMsg, 0)
                                    System.BitConverter.GetBytes(.ObjectID).CopyTo(yMsg, 2)
                                    .PlayerName.CopyTo(yMsg, 6)
                                    SendMsgToOperator(yMsg)

                                    Dim oThread As New Threading.Thread(AddressOf BeginClearPlayerCommSentBys)
                                    mlClearPlayerCommSentBys = .ObjectID
                                    oThread.Start()
                                End If



                                ReDim .PlayerUserName(19)
                                StringToBytes(BytesToString(yNewUserName).ToUpper).CopyTo(.PlayerUserName, 0)
                                ReDim .PlayerPassword(19)
                                StringToBytes(BytesToString(yNewPassword).ToUpper).CopyTo(.PlayerPassword, 0)
                                ReDim .ExternalEmailAddress(254)
                                yEmailAddress.CopyTo(.ExternalEmailAddress, 0)
                                .iEmailSettings = eEmailSettings.eEngagedEnemy Or eEmailSettings.eLowResources Or eEmailSettings.ePlayerRels Or eEmailSettings.eResearchComplete Or eEmailSettings.eTradeRequested Or eEmailSettings.eUnderAttack

                                'let's update the email server
                                If .InMyDomain = True Then
                                    Dim yEmail(282) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.eEmailSettings).CopyTo(yEmail, 0) 'lPos += 2
                                    System.BitConverter.GetBytes(.ObjectID).CopyTo(yEmail, 2)   'lpos += 4
                                    System.BitConverter.GetBytes(.iEmailSettings).CopyTo(yEmail, 6) 'lpos += 2
                                    .ExternalEmailAddress.CopyTo(yEmail, 8) 'lpos += 255
                                    .PlayerName.CopyTo(yEmail, 263)
                                    Try
                                        moEmailSrvr.SendData(yEmail)
                                    Catch
                                    End Try
                                End If

                                .DataChanged()
                            End If

                            If lIndex <> -1 Then moClients(lIndex).SendData(yResp)

                            If yResp(2) = 0 Then

                                ReDim yResp(45)
                                lPos = 0
                                System.BitConverter.GetBytes(GlobalMessageCode.ePlayerInitialSetup).CopyTo(yResp, lPos) : lPos += 2
                                System.BitConverter.GetBytes(.ObjectID).CopyTo(yResp, lPos) : lPos += 4
                                .PlayerUserName.CopyTo(yResp, lPos) : lPos += 20
                                .PlayerPassword.CopyTo(yResp, lPos) : lPos += 20
                                BroadcastToDomains(yResp)
                            End If

                            Exit For

                        End If
                    End If
                End With
            End If
        Next X
    End Sub

    Private Sub HandleSetMaintenanceTarget(ByRef yData() As Byte, ByVal iMsgCode As Int16, ByVal bPathfinding As Boolean)
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeId As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim lX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oEntity As Epica_Entity = Nothing
        If iObjTypeId = ObjectType.eFacility Then
            oEntity = GetEpicaFacility(lObjID)
        ElseIf iObjTypeId = ObjectType.eUnit Then
            oEntity = GetEpicaUnit(lObjID)
        End If
        If oEntity Is Nothing Then Return

        If bPathfinding = True Then
            'Ok, we have reached our dest, send the msg to the Region Server responsible
            With CType(oEntity.ParentObject, Epica_GUID)
                If .ObjTypeID = ObjectType.ePlanet Then
                    CType(oEntity.ParentObject, Planet).oDomain.DomainSocket.SendData(yData)
                ElseIf .ObjTypeID = ObjectType.eSolarSystem Then
                    CType(oEntity.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yData)
                End If
            End With
        Else
            'Ok, region server is telling us that repairs or dismantle may begin
            If iMsgCode = GlobalMessageCode.eSetDismantleTarget Then
                'Ok, destroy the target...
                Dim lStatus As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lArmorHP(3) As Int32
                For X As Int32 = 0 To 3
                    lArmorHP(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Next X

                If iTargetTypeID = ObjectType.eFacility Then
                    For X As Int32 = 0 To glFacilityUB
                        If glFacilityIdx(X) = lTargetID Then
                            Dim oFac As Facility = goFacility(X)
                            If oFac Is Nothing = False Then
                                glFacilityIdx(X) = -1
                                With oFac
                                    RemoveLookupFacility(.ObjectID, .ObjTypeID)
                                    If oEntity.Owner.ObjectID <> .Owner.ObjectID Then Return

                                    Dim iTemp As Int16 = CType(.ParentObject, Epica_GUID).ObjTypeID

                                    'If .EntityDef.oPrototype Is Nothing = False AndAlso .EntityDef.oPrototype.Owner Is Nothing = False AndAlso .EntityDef.oPrototype.Owner.ObjectID <> 0 Then
                                    '    Dim oSocket As NetSock = Nothing
                                    '    If iTemp = ObjectType.ePlanet Then
                                    '        oSocket = CType(.ParentObject, Planet).oDomain.DomainSocket
                                    '    ElseIf iTemp = ObjectType.eSolarSystem Then
                                    '        oSocket = CType(.ParentObject, SolarSystem).oDomain.DomainSocket
                                    '    End If
                                    '    If oSocket Is Nothing = False Then
                                    '        Dim lParentID As Int32 = CType(.ParentObject, Epica_GUID).ObjectID

                                    '        Dim yCacheType As Byte = 0
                                    '        If (.EntityDef.yChassisType And ChassisType.eGroundBased) <> 0 Then
                                    '            yCacheType = yCacheType Or MineralCacheType.eGround
                                    '        ElseIf (.EntityDef.yChassisType And ChassisType.eAtmospheric) <> 0 OrElse (.EntityDef.yChassisType And ChassisType.eSpaceBased) <> 0 Then
                                    '            yCacheType = yCacheType Or MineralCacheType.eFlying
                                    '        ElseIf (.EntityDef.yChassisType And ChassisType.eNavalBased) <> 0 Then
                                    '            yCacheType = yCacheType Or MineralCacheType.eNaval
                                    '        End If

                                    '        .EntityDef.oPrototype.CreateDismantledCaches(oSocket, lStatus, lX, lZ, lParentID, iTemp, lArmorHP(0), lArmorHP(1), lArmorHP(2), lArmorHP(3), yCacheType)
                                    '    End If
                                    'End If
                                    If iObjTypeId = ObjectType.eUnit Then
                                        CType(oEntity, Unit).DismantleFacility(oFac)
                                    End If

                                    Dim yMsg(8) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.eRemoveObject).CopyTo(yMsg, 0)
                                    .GetGUIDAsString.CopyTo(yMsg, 2)
                                    yMsg(8) = 0
                                    If iTemp = ObjectType.ePlanet Then
                                        CType(.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
                                    ElseIf iTemp = ObjectType.eSolarSystem Then
                                        CType(.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
                                    End If

                                    If .oMiningBid Is Nothing = False Then .oMiningBid.SendBiddersDeathAlert(.Owner.ObjectID, True)

                                    .RemoveMe()
                                    .DeleteEntity(X)
                                End With
                            End If
                            Exit For
                        End If
                    Next X
                ElseIf iTargetTypeID = ObjectType.eUnit Then
                    For X As Int32 = 0 To glUnitUB
                        If glUnitIdx(X) = lTargetID Then
                            Dim oUnit As Unit = goUnit(X)
                            If oUnit Is Nothing = False Then
                                With oUnit
                                    If oEntity.Owner.ObjectID <> .Owner.ObjectID Then Return
                                    Dim iTemp As Int16 = CType(.ParentObject, Epica_GUID).ObjTypeID

                                    'If .EntityDef.oPrototype Is Nothing = False AndAlso .EntityDef.oPrototype.Owner Is Nothing = False AndAlso .EntityDef.oPrototype.Owner.ObjectID <> 0 Then
                                    '    Dim oSocket As NetSock = Nothing
                                    '    If iTemp = ObjectType.ePlanet Then
                                    '        oSocket = CType(.ParentObject, Planet).oDomain.DomainSocket
                                    '    ElseIf iTemp = ObjectType.eSolarSystem Then
                                    '        oSocket = CType(.ParentObject, SolarSystem).oDomain.DomainSocket
                                    '    End If
                                    '    If oSocket Is Nothing = False Then
                                    '        Dim lParentID As Int32 = CType(.ParentObject, Epica_GUID).ObjectID

                                    '        Dim yCacheType As Byte = 0
                                    '        If (.EntityDef.yChassisType And ChassisType.eGroundBased) <> 0 Then
                                    '            yCacheType = yCacheType Or MineralCacheType.eGround
                                    '        ElseIf (.EntityDef.yChassisType And ChassisType.eAtmospheric) <> 0 OrElse (.EntityDef.yChassisType And ChassisType.eSpaceBased) <> 0 Then
                                    '            yCacheType = yCacheType Or MineralCacheType.eFlying
                                    '        ElseIf (.EntityDef.yChassisType And ChassisType.eNavalBased) <> 0 Then
                                    '            yCacheType = yCacheType Or MineralCacheType.eNaval
                                    '        End If

                                    '        .EntityDef.oPrototype.CreateDismantledCaches(oSocket, lStatus, lX, lZ, lParentID, iTemp, lArmorHP(0), lArmorHP(1), lArmorHP(2), lArmorHP(3), yCacheType)
                                    '    End If
                                    'End If
                                    If iObjTypeId = ObjectType.eUnit Then
                                        CType(oEntity, Unit).DismantleChildUnit(oUnit)
                                    End If

                                    Dim yMsg(8) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.eRemoveObject).CopyTo(yMsg, 0)
                                    .GetGUIDAsString.CopyTo(yMsg, 2)
                                    yMsg(8) = 0
                                    If iTemp = ObjectType.ePlanet Then
                                        CType(.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
                                    ElseIf iTemp = ObjectType.eSolarSystem Then
                                        CType(.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
                                    End If

                                    .DeleteEntity(X)
                                End With
                            End If

                            glUnitIdx(X) = -1
                            Exit For
                        End If
                    Next X
                End If

            Else
                'Repair Target Order
                If iTargetTypeID = ObjectType.eUnit Then
                    'OK, special repair order, we repair the unit's engines only
                    Dim oUnit As Unit = GetEpicaUnit(lTargetID)
                    If oUnit Is Nothing = False Then
                        Dim lTemp As Int32 = 0
                        For X As Int32 = 0 To oUnit.EntityDef.lSideCrits.GetUpperBound(0)
                            lTemp = lTemp Or oUnit.EntityDef.lSideCrits(X)
                        Next X
                        If (lTemp And elUnitStatus.eEngineOperational) <> 0 Then
                            oUnit.CurrentStatus = oUnit.CurrentStatus Or elUnitStatus.eEngineOperational
                            Dim yMsg(11) As Byte
                            System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityStatus).CopyTo(yMsg, 0)
                            oUnit.GetGUIDAsString.CopyTo(yMsg, 2)
                            System.BitConverter.GetBytes(elUnitStatus.eEngineOperational).CopyTo(yMsg, 8)

                            Dim iTemp As Int16 = CType(oUnit.ParentObject, Epica_GUID).ObjTypeID
                            If iTemp = ObjectType.ePlanet Then
                                CType(oUnit.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
                            ElseIf iTemp = ObjectType.eSolarSystem Then
                                CType(oUnit.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
                            End If
                        End If
                    End If
                End If
            End If
        End If


    End Sub

#End Region

#Region "  Outbound Message Handlers  "
    Public Sub SendPathfindingNewEntity(ByRef oObjGuid As Epica_GUID, ByRef oParentGuid As Epica_GUID, ByVal lX As Int32, ByVal lZ As Int32, ByVal iModelID As Int16)
        Dim yMsg() As Byte = GetPathfindingNewEntityMsg(oObjGuid, oParentGuid, lX, lZ, iModelID)
        If yMsg Is Nothing Then Return
        moPathfinding.SendData(yMsg)
    End Sub

    Public Sub SendDomainEntityStatus(ByRef oEntity As Epica_Entity, ByVal lStatusChg As Int32)
        Dim yMsg(11) As Byte

        If gbServerInitializing = True Then Return

        System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityStatus).CopyTo(yMsg, 0)
        oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(lStatusChg).CopyTo(yMsg, 8)

        Dim iTemp As Int16 = CType(oEntity.ParentObject, Epica_GUID).ObjTypeID
        If iTemp = ObjectType.ePlanet Then
            CType(oEntity.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
        ElseIf iTemp = ObjectType.eSolarSystem Then
            CType(oEntity.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
        End If
    End Sub

    Public Sub SendPlayerCPLimitUpdate(ByVal lPlayerID As Int32, ByVal iCPLimit As Int16)
        Dim yMsg(13) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(12S).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerTechValue).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lPlayerID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(PlayerSpecialAttributeID.eCPLimit).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(CInt(iCPLimit)).CopyTo(yMsg, lPos) : lPos += 4

        For X As Int32 = 0 To mlDomainUB
            moDomains(X).SendLenAppendedData(yMsg)
        Next X
    End Sub

    Public Sub SendOutboundEmail(ByRef oComm As PlayerComm, ByRef oPlayer As Player, ByVal iBaseAlertType As Int16, ByVal lExt1 As Int32, ByVal lExt2 As Int32, ByVal lExt3 As Int32, ByVal lExt4 As Int32, ByVal lExt5 As Int32, ByVal lExt6 As Int32, ByVal sExtraBody As String)
        If moEmailSrvr Is Nothing Then Return
        If oPlayer.AccountStatus <> AccountStatusType.eActiveAccount AndAlso oPlayer.AccountStatus <> AccountStatusType.eTrialAccount AndAlso oPlayer.ObjectID <> 19326 AndAlso oPlayer.AccountStatus <> AccountStatusType.eMondelisActive Then Return

        Dim yMsg() As Byte
        Dim lPos As Int32 = 0
        Dim lSubLen As Int32 = 0
        Dim lBodyLen As Int32 = 0

        Dim sFinalBody As String = ""
        If oComm.MsgBody Is Nothing = False Then sFinalBody = BytesToString(oComm.MsgBody)
        If sExtraBody Is Nothing = False Then sFinalBody &= sExtraBody

        If oComm.MsgTitle Is Nothing = False Then lSubLen = oComm.MsgTitle.Length
        lBodyLen = sFinalBody.Length

        ReDim yMsg(43 + lSubLen + lBodyLen)

        System.BitConverter.GetBytes(GlobalMessageCode.eSendOutMailMsg).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(iBaseAlertType).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(oComm.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSubLen).CopyTo(yMsg, lPos) : lPos += 4
        If oComm.MsgTitle Is Nothing = False Then
            oComm.MsgTitle.CopyTo(yMsg, lPos) : lPos += lSubLen
        End If
        System.BitConverter.GetBytes(lBodyLen).CopyTo(yMsg, lPos) : lPos += 4
        If lBodyLen <> 0 Then StringToBytes(sFinalBody).CopyTo(yMsg, lPos) : lPos += lBodyLen

        System.BitConverter.GetBytes(lExt1).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lExt2).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lExt3).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lExt4).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lExt5).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lExt6).CopyTo(yMsg, lPos) : lPos += 4

        SendToEmailSrvr(yMsg) 'moEmailSrvr.SendData(yMsg)
    End Sub

    Public Sub SendPathfindingRemoveObject(ByRef oGUID As Epica_GUID)
        Dim yPFMsg(7) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eRemoveObject).CopyTo(yPFMsg, 0)
        oGUID.GetGUIDAsString.CopyTo(yPFMsg, 2)
        moPathfinding.SendData(yPFMsg)
        Erase yPFMsg
    End Sub

    Public Sub SendToPathfinding(ByRef yMsg() As Byte)
        moPathfinding.SendData(yMsg)
    End Sub

#End Region
End Class

