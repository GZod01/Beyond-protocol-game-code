Option Strict On

Public Enum elBusyStatusType As Int32
	PrimaryServerLogin = 1
	OperatorServerLogin = 2
	GalaxyAndSystems = 4
	StarTypes = 8
	SystemDetails = 16
	PlayerDetails = 32
	PrimaryValidateVersion = 64
	OperatorValidateVersion = 128
	SkillsList = 256
	GoalsList = 512

	TradesNotification = 1024
End Enum
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
Public Enum eyConnType As Byte
	OperatorServer = 0
	PrimaryServer = 1
	RegionServer = 2
	ChatServer = 3			'???
End Enum

Public Class MsgSystem
    'Ok, the good news is... the client doesn't receive connections... it connects to the Primary and to the Domain
    '  It use to be that it would also connect to the Pathfinding Server but I removed that idea to allow the Domain
    '  server as a passthru for the Move Object Messages. So... the client connects to the Primary and the Domain
    Private WithEvents moPrimary As NetSock     'the primary server
    Private WithEvents moRegion As NetSock      'the region server i am currently connected to
    Private WithEvents moOperator As NetSock

	'Private mbConnectedToPrimary As Boolean = False
	'Private mbConnectedToRegion As Boolean = False
	'Private mbConnectedToOperator As Boolean = False
	'Public mbConnectingToPrimary As Boolean = False
	'Public mbConnectingToRegion As Boolean = False
	'Public mbConnectingToOperator As Boolean = False

	Private msPrimaryIP As String = ""
    Private mlPrimaryPort As Int32 = 0

    Private mlTimeout As Int32

    Private myReactionData() As Byte
    Private moReactionThread As Threading.Thread

	'Private mbBusy As Boolean = False
	'Private mbPlayerDetailsArrived As Boolean = False
    Public mlBusyStatus As Int32 = 0

	'Private mlLoginResp As Int32
    Private mbClosing As Boolean = False

    Private msCurrentLoc As String
	Private mbVersionValidated As Boolean = False

	Private mlLastAlert As Int32		'generally speaking

	Private mlLastAddObjectID As Int32 = -1
	Private miLastAddObjectTypeID As Int16 = -1

    Public bDoNotClearBusyStatus As Boolean = False 
	'Public Shared l_REGION_MSGS As Int32 = 0
	'Public Shared l_PRIMARY_MSGS As Int32 = 0
    Private mbDebugMessages As Boolean = False
    Private miLastAgentMsgUpdate As Int32
#Region " Message System Events "
	'Public Event ConnectionWaitStatus(ByVal fPercentage As Single)
    Public Event ReceivedPlayerCurrentEnvironment(ByVal lPlayerID As Int32, ByVal lID As Int32, ByVal iTypeID As Int16, ByVal lGalaxyID As Int32, ByVal lSystemID As Int32, ByVal lPlanetID As Int32, ByVal bInCurrDomain As Boolean)
    Public Event ReceivedEnvironmentsDomain(ByVal sIPAddress As String, ByVal iPort As Int16)
	Public Event ProcessedBurstEnvironment()
	Public Event ReceivedStarTypes(ByVal yData() As Byte)
    Public Event ReceivedGalaxyAndSystems(ByVal yData() As Byte)
    Public Event ReceivedSystemDetails(ByVal yData() As Byte)
	Public Event ServerShutdown()

	Public Event ConnectedToServer(ByVal yConnType As eyConnType)
	Public Event ValidateVersionResponse(ByVal yConnType As eyConnType, ByVal bValue As Boolean)
	Public Event LoginResponse(ByVal yConnType As eyConnType, ByVal lVal As Int32)
	Public Event ReceivedPrimaryServerData()
	Public Event ServerDisconnected(ByVal yConnType As eyConnType)
	Public Event ReceivedAllPlayerDetails()
#End Region

    Public Sub DisconnectAll()  ' disconnects all connections
        'in almost every case I can think of, if an error occurs here, we don't care
        On Error Resume Next
        msCurrentLoc = "DisconnectAll"
        mbClosing = True
        If moRegion Is Nothing = False Then moRegion.Disconnect()
        If moPrimary Is Nothing = False Then moPrimary.Disconnect()
        mlBusyStatus = 0
        'If moOperator Is Nothing = False Then moOperator.Disconnect()
    End Sub

    Public Sub ResetClosing()
        mbClosing = False
    End Sub

	Public Sub SendToPrimary(ByVal yData() As Byte)
        msCurrentLoc = "SendToPrimary"
        Dim iMsgID As Short = System.BitConverter.ToInt16(yData, 0)
        If mbDebugMessages = True Then Debug.Print("SendToPrimary: " & CType(iMsgID, GlobalMessageCode).ToString)
        If moPrimary Is Nothing = False Then moPrimary.SendData(yData)
    End Sub

    Public Sub SendLenAppendMsgToPrimary(ByRef yData() As Byte)
        msCurrentLoc = "SendLenAppendMsgToPrimary"
        If mbDebugMessages = True Then
            'no try-catch here if it pops an error; our packet was crafted wrong and we should go fix it
            Debug.Print("SendLenAppendMsgToPrimary: " & yData.Length.ToString)
            Dim lPos As Int32 = 0
            Dim iLen As Int32 = 0
            Dim iMsgID As Short
            While lPos < yData.Length
                iLen = yData(lPos)
                'iMsgID = yData(lPos + 2)
                iMsgID = System.BitConverter.ToInt16(yData, lPos + 2)
                Debug.Print("SendLenAppendMsgToPrimary: " & CType(iMsgID, GlobalMessageCode).ToString)
                lPos += iLen + 2
            End While
        End If
        moPrimary.SendLenAppendedData(yData)
    End Sub
	Public Sub SendToOperator(ByVal yData() As Byte)
        msCurrentLoc = "SendToOperator"
        Dim iMsgID As Short = System.BitConverter.ToInt16(yData, 0)
        If mbDebugMessages = True Then Debug.Print("SendToOperator: " & CType(iMsgID, GlobalMessageCode).ToString)
        If moOperator Is Nothing = False Then moOperator.SendData(yData)
    End Sub

	Public Sub SendToRegion(ByVal yData() As Byte)
        msCurrentLoc = "SendToRegion"
        Dim iMsgID As Short = System.BitConverter.ToInt16(yData, 0)
        If mbDebugMessages = True Then Debug.Print("SendToRegion: " & CType(iMsgID, GlobalMessageCode).ToString)
        moRegion.SendData(yData)
	End Sub

	Public Sub SendLenAppendMsgToRegion(ByRef yData() As Byte)
        msCurrentLoc = "SendLenAppendMsgToRegion"
        If mbDebugMessages = True Then
            'no try-catch here if it pops an error; our packet was crafted wrong and we should go fix it
            Debug.Print("SendLenAppendMsgToRegion: " & yData.Length.ToString)
            Dim lPos As Int32 = 0
            Dim iLen As Int32 = 0
            Dim iMsgID As Short
            While lPos < yData.Length
                iLen = yData(lPos)
                'iMsgID = yData(lPos + 2)
                iMsgID = System.BitConverter.ToInt16(yData, lPos + 2)
                Debug.Print("SendLenAppendMsgToRegion: " & CType(iMsgID, GlobalMessageCode).ToString)
                lPos += iLen + 2
            End While
        End If
        moRegion.SendLenAppendedData(yData)
    End Sub

	Public Sub DisconnectRegion()
		msCurrentLoc = "DisconnectRegion"
		moRegion.Disconnect()
    End Sub
    Public Sub DisconnectPrimary()
        msCurrentLoc = "DisconnectPrimary"
        moPrimary.Disconnect()
    End Sub

#Region "  Connect to Server Methods  "

	Public Function ConnectToDomainServer(ByVal sIP As String, ByVal iPort As Int16) As Boolean
		Dim bRes As Boolean = False
		Dim lLocalPort As Int32
		Dim oINI As InitFile = New InitFile()

		msCurrentLoc = "ConnectToDomainServer"

		Try
			If iPort = 0 Or sIP = "" Then Err.Raise(-1, "ConnectToDomainServer", "Bad connection data for the requested server.")
			lLocalPort = CInt(Val(oINI.GetString("CONNECTION", "LocalDomainPort", "0")))
			moRegion = New NetSock()
			moRegion.SetLocalBinding(CShort(lLocalPort))
			'mbConnectingToRegion = True
			moRegion.Connect(sIP, iPort)

			'System.Threading.Thread.Sleep(500)

			'Dim sw As Stopwatch = Stopwatch.StartNew
			'While mbConnectedToRegion = False AndAlso (sw.ElapsedMilliseconds < mlTimeout) AndAlso mbConnectingToRegion = True
			'	'Application.DoEvents()
			'	Threading.Thread.Sleep(0)
			'	Threading.Thread.Sleep(1)
			'	RaiseEvent ConnectionWaitStatus(CSng(sw.ElapsedMilliseconds / mlTimeout))
			'End While
			'sw.Stop()
			'sw = Nothing
			'If mbConnectedToRegion = False Then
			'	Err.Raise(-1, "ConnectToDomainServer", "Unable to connect to domain server: Connection timed out!")
			'End If
			'RaiseEvent ConnectedToRegion()

			bRes = True
		Catch
			MsgBox(Err.Description, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Connection Failure")
			'mbConnectedToRegion = False
		Finally
			'bRes = mbConnectedToRegion
			oINI = Nothing
		End Try
		Return bRes
	End Function

	Public Function ConnectToPrimary() As Boolean
		Dim oINI As InitFile = New InitFile()
		'Dim lPort As Int32
		'Dim lLocalPort As Int32
		'Dim sIP As String
		Dim bRes As Boolean = False

		msCurrentLoc = "ConnectToPrimary"

		Try
			mlTimeout = CInt(Val(oINI.GetString("CONNECTION", "ConnectTimeout", "10000")))
			'lPort = CInt(Val(oINI.GetString("CONNECTION", "PrimaryPort", "0")))
			'lLocalPort = CInt(Val(oINI.GetString("CONNECTION", "LocalPrimaryPort", "0")))
			'sIP = oINI.GetString("CONNECTION", "PrimaryIP", "") 
			'If oINI.GetString("CONNECTION", "OverrideIP", "") <> "" Then msPrimaryIP = oINI.GetString("CONNECTION", "OverrideIP", "")

			'TODO: HARDCODED FOR PERFORMANCE TEST
			'msPrimaryIP = "216.231.221.181"

			If mlPrimaryPort = 0 Or msPrimaryIP = "" Then
				Err.Raise(-1, "EstablishConnection", "Unable to connect to server: could not find connection credentials.")
            End If

            moPrimary = New NetSock()
            moPrimary.SetLocalBinding(0)
            'mbConnectingToRegion = True
            'moPrimary.SetLocalBinding(CShort(lLocalPort))
			moPrimary.Connect(msPrimaryIP, mlPrimaryPort)
			'Dim sw As Stopwatch = Stopwatch.StartNew
			'While mbConnectedToPrimary = False AndAlso (sw.ElapsedMilliseconds < mlTimeout) AndAlso mbConnectingToPrimary = True
			'	'Application.DoEvents()
			'	Threading.Thread.Sleep(0)
			'	Threading.Thread.Sleep(1)
			'	RaiseEvent ConnectionWaitStatus(CSng(sw.ElapsedMilliseconds / mlTimeout))
			'End While
			'sw.Stop()
			'sw = Nothing
			'mbConnectingToPrimary = False
			'If mbConnectedToPrimary = False Then
			'	Err.Raise(-1, "EstablishConnection", "Unable to connect to server: Connection timed out!")
			'End If
			'RaiseEvent ConnectedToPrimary()

			bRes = True
		Catch
			'MsgBox(Err.Description, MsgBoxStyle.Critical Or MsgBoxStyle.OKOnly, "Connection Failure")
			'mbConnectedToPrimary = False
		Finally
			'bRes = mbConnectedToPrimary
			oINI = Nothing
		End Try
		Return bRes
	End Function

	Public Function ConnectToOperator() As Boolean
		Dim oINI As InitFile = New InitFile()
		Dim lPort As Int32
		Dim lLocalPort As Int32
		Dim sIP As String
		Dim bRes As Boolean = False

		msCurrentLoc = "ConnectToOperator"

		Try
			mlTimeout = CInt(Val(oINI.GetString("CONNECTION", "ConnectTimeout", "10000")))
			lPort = CInt(Val(oINI.GetString("CONNECTION", "OperatorPort", "7779")))
			lLocalPort = CInt(Val(oINI.GetString("CONNECTION", "LocalPrimaryPort", "0")))
            sIP = oINI.GetString("CONNECTION", "OperatorIP", "74.113.102.133") '"www.beyondprotocol.net")

            If lPort = 0 Or sIP = "" Then
                Err.Raise(-1, "EstablishConnection", "Unable to connect to server: could not find connection credentials.")
            End If
            'mbConnectingToOperator = True
            If moOperator Is Nothing = False Then
                moOperator.Disconnect()
                moOperator = Nothing
                GC.Collect()
            End If
            moOperator = New NetSock()
			moOperator.Connect(sIP, lPort)
			'Dim sw As Stopwatch = Stopwatch.StartNew
			'While mbConnectedToOperator = False AndAlso (sw.ElapsedMilliseconds < mlTimeout) AndAlso mbConnectingToOperator = True
			'	'Application.DoEvents()
			'	Threading.Thread.Sleep(0)
			'	Threading.Thread.Sleep(1)
			'	RaiseEvent ConnectionWaitStatus(CSng(sw.ElapsedMilliseconds / mlTimeout))
			'End While
			'sw.Stop()
			'sw = Nothing
			'mbConnectingToOperator = False
			'If mbConnectedToOperator = False Then
			'	Err.Raise(-1, "EstablishConnection", "Unable to connect to server: Connection timed out!")
			'End If
			bRes = True
		Catch
			'mbConnectedToOperator = False
		Finally
			'bRes = mbConnectedToOperator
			oINI = Nothing
		End Try
		Return bRes
	End Function

    Public Function GiveUpAndConnectToBackupOperator() As Boolean
        Dim oINI As InitFile = New InitFile()
        Dim lPort As Int32
        Dim sIP As String
        Dim bRes As Boolean = False

        msCurrentLoc = "GiveUpAndConnectToBackupOperator"

        Try
            lPort = CInt(Val(oINI.GetString("CONNECTION", "BackupOperatorPort", "7779")))
            sIP = oINI.GetString("CONNECTION", "BackupOperatorIP", "12.228.7.143") '

            If lPort = 0 Or sIP = "" Then
                Err.Raise(-1, "EstablishConnection", "Unable to connect to server: could not find connection credentials.")
            End If
            'mbConnectingToOperator = True
            If moOperator Is Nothing = False Then
                moOperator.Disconnect()
                moOperator = Nothing
                GC.Collect()
            End If
            moOperator = New NetSock()
            moOperator.Connect(sIP, lPort)
            'Dim sw As Stopwatch = Stopwatch.StartNew
            'While mbConnectedToOperator = False AndAlso (sw.ElapsedMilliseconds < mlTimeout) AndAlso mbConnectingToOperator = True
            '	'Application.DoEvents()
            '	Threading.Thread.Sleep(0)
            '	Threading.Thread.Sleep(1)
            '	RaiseEvent ConnectionWaitStatus(CSng(sw.ElapsedMilliseconds / mlTimeout))
            'End While
            'sw.Stop()
            'sw = Nothing
            'mbConnectingToOperator = False
            'If mbConnectedToOperator = False Then
            '	Err.Raise(-1, "EstablishConnection", "Unable to connect to server: Connection timed out!")
            'End If
            bRes = True
        Catch
            'mbConnectedToOperator = False
        Finally
            'bRes = mbConnectedToOperator
            oINI = Nothing
        End Try
        Return bRes
    End Function
#End Region

#Region " Operator Server Events "
	Private Sub moOperator_onConnect(ByVal Index As Integer) Handles moOperator.onConnect
		'mbConnectedToOperator = True
		'mbConnectingToOperator = False
		RaiseEvent ConnectedToServer(eyConnType.OperatorServer)
	End Sub

	Private Sub moOperator_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moOperator.onDataArrival
		Dim iMsgID As Short

        msCurrentLoc = "oda"

        iMsgID = System.BitConverter.ToInt16(Data, 0)
        If mbDebugMessages = True Then Debug.Print("moOperator_onDataArrival: " & CType(iMsgID, GlobalMessageCode).ToString)
		Select Case iMsgID
			Case GlobalMessageCode.ePathfindingConnectionInfo
				HandleOperatorConnInfo(Data)
			Case GlobalMessageCode.ePlayerLoginResponse
				'Ok, the response... the first four bytes of the message *should* be our reason or the player ID
				Dim lResp As Int32 = LoginResponseCodes.eUnknownFailure
				If Data.Length > 13 Then
					lResp = System.BitConverter.ToInt32(Data, 2)
					glActualPlayerID = System.BitConverter.ToInt32(Data, 6)
					glAliasRights = System.BitConverter.ToInt32(Data, 10)
				ElseIf Data.Length > 5 Then
					lResp = System.BitConverter.ToInt32(Data, 2)
					glActualPlayerID = lResp
				Else
					lResp = LoginResponseCodes.eUnknownFailure
				End If
				If (mlBusyStatus And elBusyStatusType.OperatorServerLogin) <> 0 Then
					mlBusyStatus = mlBusyStatus Xor elBusyStatusType.OperatorServerLogin
				End If
				RaiseEvent LoginResponse(eyConnType.OperatorServer, lResp)
			Case GlobalMessageCode.eClientVersion
				HandleValidateVersion(Data, eyConnType.OperatorServer)
			Case GlobalMessageCode.ePlayerInitialSetup
				If goUILib Is Nothing = False Then
					Dim ofrm As frmPlayerSetup = CType(goUILib.GetWindow("frmPlayerSetup"), frmPlayerSetup)
					If ofrm Is Nothing = False Then ofrm.HandleSetupMsg(Data)
					ofrm = Nothing
				End If
		End Select
	End Sub

	Private Sub HandleOperatorConnInfo(ByRef yData() As Byte)
		Dim lPos As Int32 = 2		'msgcode
		msPrimaryIP = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		mlPrimaryPort = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		RaiseEvent ReceivedPrimaryServerData()
		RaiseEvent ServerDisconnected(eyConnType.OperatorServer)
	End Sub

	Public Sub DisconnectOperator()
		moOperator.Disconnect()
    End Sub

#End Region

#Region " Primary Server Events "

    Private Sub moPrimary_onConnect(ByVal Index As Integer) Handles moPrimary.onConnect
        RaiseEvent ConnectedToServer(eyConnType.PrimaryServer)
    End Sub

    Private Sub moPrimary_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moPrimary.onConnectionRequest
        'do nothing, should never happen
    End Sub

    Private Sub moPrimary_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moPrimary.onDataArrival
		Dim iMsgID As Short

        'l_PRIMARY_MSGS += 1
        If Data Is Nothing = True Then Return

        msCurrentLoc = "pda"
		If Data.Length < 2 Then
			'If (mlBusyStatus And elBusyStatusType.PlayerDetails) <> 0 Then
			'	Dim yMsg(5) As Byte
			'	msCurrentLoc = "RequestPlayerDetails"
			'	System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerDetails).CopyTo(yMsg, 0)
			'	System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 2)
			'	moPrimary.SendData(yMsg)
			'End If

			If (mlBusyStatus And elBusyStatusType.PlayerDetails) <> 0 Then
				If mbReRequestInProgress = False Then
					'DoLogEvent("moPrimary Bug. Retrying " & mlLastRequestedType)
					mbReRequestInProgress = True
					WriteToRPDFile("  Retrying " & mlLastRequestedType)
					Dim oThread As New Threading.Thread(AddressOf __ThreadedReRequestPD)
					oThread.Start()
				End If
			End If

            'goUILib.AddNotification("moPrimary bug", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		iMsgID = System.BitConverter.ToInt16(Data, 0)
		'If (mlBusyStatus And elBusyStatusType.PlayerDetails) = 0 Then
		'	goUILib.AddNotification("Primary: " & CType(iMsgID, GlobalMessageCode).ToString, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		'End If
        If mbDebugMessages = True Then Debug.Print("moPrimary_onDataArrival: " & CType(iMsgID, GlobalMessageCode).ToString)
        msCurrentLoc = "pda. (" & iMsgID & ")"

		'goUILib.AddNotification("Primary: " & CType(iMsgID, GlobalMessageCode).ToString, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        Select Case iMsgID
            'Case GlobalMessageCode.eRewardWarpoints
            '    HandleRewardOrReportWarpoints(Data)
            'Case GlobalMessageCode.eReportWarpoints
            '    HandleRewardOrReportWarpoints(Data)
            Case GlobalMessageCode.ePlayerCurrentEnvironment
                'HandlePlayerCurrentEnvirMsg(Data)
                myReactionData = Data
                moReactionThread = New Threading.Thread(AddressOf __ReactionThreadStart)
                moReactionThread.Start()
            Case GlobalMessageCode.eEnvironmentDomain
                HandleEnvirDomainMsg(Data)
            Case GlobalMessageCode.eResponseGalaxyAndSystems
                RaiseEvent ReceivedGalaxyAndSystems(Data)
                If (mlBusyStatus And elBusyStatusType.GalaxyAndSystems) <> 0 Then
                    mlBusyStatus = mlBusyStatus Xor elBusyStatusType.GalaxyAndSystems
                End If
            Case GlobalMessageCode.eResponseStarTypes
                RaiseEvent ReceivedStarTypes(Data)
                If (mlBusyStatus And elBusyStatusType.StarTypes) <> 0 Then
                    mlBusyStatus = mlBusyStatus Xor elBusyStatusType.StarTypes
                End If
            Case GlobalMessageCode.ePlayerLoginResponse
                'Ok, the response... the first four bytes of the message *should* be our reason or the player ID
                Dim lResp As Int32 = LoginResponseCodes.eUnknownFailure
                If Data.Length > 13 Then
                    lResp = System.BitConverter.ToInt32(Data, 2)
                    glActualPlayerID = System.BitConverter.ToInt32(Data, 6)
                    glAliasRights = System.BitConverter.ToInt32(Data, 10)
                    If glActualPlayerID <> glPlayerID Then gbAliased = True
                ElseIf Data.Length > 5 Then
                    lResp = System.BitConverter.ToInt32(Data, 2)
                    glActualPlayerID = lResp
                    If glActualPlayerID <> glPlayerID Then gbAliased = True
                Else
                    lResp = LoginResponseCodes.eUnknownFailure
                End If
                If (mlBusyStatus And elBusyStatusType.PrimaryServerLogin) <> 0 Then
                    mlBusyStatus = mlBusyStatus Xor elBusyStatusType.PrimaryServerLogin
                End If
                RaiseEvent LoginResponse(eyConnType.PrimaryServer, lResp)
            Case GlobalMessageCode.eResponseSystemDetails
                bDoNotClearBusyStatus = False
                RaiseEvent ReceivedSystemDetails(Data)
                If bDoNotClearBusyStatus = False AndAlso (mlBusyStatus And elBusyStatusType.SystemDetails) <> 0 Then
                    mlBusyStatus = mlBusyStatus Xor elBusyStatusType.SystemDetails
                End If
            Case GlobalMessageCode.eResponseSystemDetailsFail
                'TODO: Define this, this is bad
            Case GlobalMessageCode.eAddObjectCommand, GlobalMessageCode.eGetArchivedItems
                HandleAddObjectMsg(Data, True)
            Case GlobalMessageCode.eUpdatePlayerCredits
                HandleUpdatePlayerCredits(Data)
            Case GlobalMessageCode.eSetPlayerRel
                HandleSetPlayerRel(Data)
            Case GlobalMessageCode.eRequestEntityContents
                HandleRequestEntityContents(Data)
            Case GlobalMessageCode.eEntityProductionStatus
                HandleProductionStatusMsg(Data)
            Case GlobalMessageCode.eProductionCompleteNotice
                HandleProductionCompleteNotice(Data)
            Case GlobalMessageCode.eGetEntityProduction
                If goUILib Is Nothing = False Then goUILib.HandleProductionQueueMsg(Data)
            Case GlobalMessageCode.eSetEntityProdSucceed
                HandleSetEntityProdSucceed(Data)
            Case GlobalMessageCode.eSetEntityProdFailed
                HandleSetEntityProdFailed(Data)
            Case GlobalMessageCode.eAddPlayerRelObject
                HandleAddPlayerRelMsg(Data)
            Case GlobalMessageCode.eGetAvailableResources
                If goAvailableResources Is Nothing Then goAvailableResources = New AvailResources()
                goAvailableResources.FillFromMsg(Data)
            Case GlobalMessageCode.eChatMessage
                HandleChatMsg(Data)
            Case GlobalMessageCode.eRequestSkillList
                HandleSkillListMsg(Data)
            Case GlobalMessageCode.eRequestGoalList
                HandleGoalListMsg(Data)
            Case GlobalMessageCode.eGetEntityName
                HandleGetEntityName(Data)
            Case GlobalMessageCode.eGetColonyDetails
                If goUILib Is Nothing = False Then goUILib.HandleColonyStatsMsg(Data)
            Case GlobalMessageCode.eSetEntityPersonnel
                HandleSetEntityPersonnel(Data)
            Case GlobalMessageCode.eServerShutdown
                RaiseEvent ServerShutdown()
            Case GlobalMessageCode.eColonyLowResources
                HandleColonyLowResources(Data)
            Case GlobalMessageCode.eBugList
                If goUILib Is Nothing = False Then
                    Dim oTmpWin As frmBugMain = CType(goUILib.GetWindow("frmBugMain"), frmBugMain)
                    If oTmpWin Is Nothing = False Then oTmpWin.HandleAddBugToList(Data)
                End If
            Case GlobalMessageCode.eSetEntityStatus
                HandleSetEntityStatus(Data)
            Case GlobalMessageCode.eClientVersion
                HandleValidateVersion(Data, eyConnType.PrimaryServer)
            Case GlobalMessageCode.eGetColonyChildList
                HandleGetColonyChildList(Data)
            Case GlobalMessageCode.eRequestPlayerDetails
                HandleRequestPlayerDetailsResponse(Data)
            Case GlobalMessageCode.eRequestDockFail
                HandleRequestDockFail(Data)
            Case GlobalMessageCode.eRemoveFromFleet
                HandleRemoveFromFleet(Data)
            Case GlobalMessageCode.eSetFleetDest
                HandleSetFleetDest(Data)
            Case GlobalMessageCode.eUpdateEntityParent
                HandleUpdateEntityParent(Data)
            Case GlobalMessageCode.eAddPlayerCommFolder
                HandleAddPlayerCommFolder(Data)
            Case GlobalMessageCode.eDeleteEmailItem
                HandleDeleteEmailItem(Data)
            Case GlobalMessageCode.eMoveEmailToFolder
                HandleMoveEmailToFolder(Data)
            Case GlobalMessageCode.ePlayerAlert
                HandlePlayerAlert(Data)
            Case GlobalMessageCode.eGetColonyList
                HandleGetColonyList(Data)
            Case GlobalMessageCode.eGetNonOwnerDetails
                HandleGetNonOwnerDetails(Data)
            Case GlobalMessageCode.eRequestEntityDefenses
                If goUILib Is Nothing = False Then
                    Dim ofrm As frmRepair = CType(goUILib.GetWindow("frmRepair"), frmRepair)
                    If ofrm Is Nothing = False Then
                        ofrm.HandleDefensesMsg(Data)
                    End If
                    ofrm = Nothing
                End If
            Case GlobalMessageCode.eGetPlayerScores
                HandleGetPlayerScores(Data)
            Case GlobalMessageCode.ePlayerTitleChange
                HandlePlayerTitleChange(Data)
            Case GlobalMessageCode.ePlayerInitialSetup
                If goUILib Is Nothing = False Then
                    Dim ofrm As frmPlayerSetup = CType(goUILib.GetWindow("frmPlayerSetup"), frmPlayerSetup)
                    If ofrm Is Nothing = False Then ofrm.HandleSetupMsg(Data)
                    ofrm = Nothing
                End If
            Case GlobalMessageCode.eDeathBudgetDeposit
                HandleDeathBudgetDeposit(Data)
            Case GlobalMessageCode.eGetEntityDetails
                HandleGetEntityDetailsMsg(Data)
            Case GlobalMessageCode.ePlayerIsDead
                HandlePlayerIsDead(Data)
            Case GlobalMessageCode.eDeleteDesign
                HandleDeleteDesign(Data)
            Case GlobalMessageCode.eRequestEntityAmmo
                If goUILib Is Nothing = False Then
                    Dim ofrm As frmAmmo = CType(goUILib.GetWindow("frmAmmo"), frmAmmo)
                    If ofrm Is Nothing = False Then ofrm.HandleRequestAmmo(Data)
                    ofrm = Nothing
                End If
            Case GlobalMessageCode.eEmailSettings
                If goUILib Is Nothing = False Then
                    Dim ofrm As frmEmailSetup = CType(goUILib.GetWindow("frmEmailSetup"), frmEmailSetup)
                    If ofrm Is Nothing = False Then ofrm.HandleEmailSettingsMsg(Data)
                    ofrm = Nothing
                End If

            Case GlobalMessageCode.eGetTradeHistory
                HandleGetTradeHistory(Data)
            Case GlobalMessageCode.eGetGTCList, GlobalMessageCode.eGetTradePostList, GlobalMessageCode.eGetOrderSpecifics, GlobalMessageCode.eGetTradePostTradeables, GlobalMessageCode.eSubmitBuyOrder, GlobalMessageCode.eDeliverBuyOrder, GlobalMessageCode.ePurchaseSellOrder, GlobalMessageCode.eAcceptBuyOrder
                HandleTradeWindowMsg(iMsgID, Data)
            Case GlobalMessageCode.eGetTradeDeliveries
                HandleGetTradeDeliveries(Data)
            Case GlobalMessageCode.eUpdatePlayerTechValue
                HandleUpdatePlayerTechValue(Data)
            Case GlobalMessageCode.eAddPlayerTechKnow
                HandleAddPlayerTechKnow(Data)
            Case GlobalMessageCode.eSetPlayerSpecialAttribute
                HandleSetPlayerSpecialAttribute(Data)
            Case GlobalMessageCode.eTransferContents
                HandleTransferContents(Data)
            Case GlobalMessageCode.eSetFleetReinforcer
                HandleSetFleetReinforcer(Data)
            Case GlobalMessageCode.ePlayerAliasConfig
                HandlePlayerAliasConfig(Data)
            Case GlobalMessageCode.eActOfWarNotice
                HandleActOfWarNotice(Data)
            Case GlobalMessageCode.eDeinfiltrateAgent
                HandleDeinfiltrateAgent(Data)
            Case GlobalMessageCode.eSetInfiltrateSettings
                HandleSetInfiltrateSettings(Data)
            Case GlobalMessageCode.eSetAgentStatus
                HandleSetAgentStatus(Data)
            Case GlobalMessageCode.eSubmitMission
                HandleSubmitMission(Data)
            Case GlobalMessageCode.eGetAgentStatus
                HandleGetAgentStatus(Data)
            Case GlobalMessageCode.eGetPMUpdate
                HandleGetPMUpdate(Data)
            Case GlobalMessageCode.eSetSkipStatus
                HandleSetSkipStatus(Data)
            Case GlobalMessageCode.eCaptureKillAgent
                HandleCapturedAgent(Data)
            Case GlobalMessageCode.eRemovePlayerRel
                HandleRemovePlayerRel(Data)
            Case GlobalMessageCode.eRequestPlayerDetailsPlayerMin
                HandleRequestPlayerDetailsPlayerMin(Data)
            Case GlobalMessageCode.eRequestEmailSummarys
                HandleRequestEmailSummarys(Data)
            Case GlobalMessageCode.eRequestDXDiag
                HandleRequestDXDiag(Data)
            Case GlobalMessageCode.eGetRouteList
                HandleGetRouteList(Data)
            Case GlobalMessageCode.eSetRouteMineral
                HandleSetRouteMineral(Data)
            Case GlobalMessageCode.eUpdateRouteStatus
                HandleUpdateRouteStatus(Data)
            Case GlobalMessageCode.eRemoveRouteItem
                HandleRemoveRouteItem(Data)
            Case GlobalMessageCode.eGetSkillList
                HandleGetSkillList(Data)
            Case GlobalMessageCode.eRequestPlayerDetailsAlert
                HandleRequestPlayerDetailsAlert(Data)
            Case GlobalMessageCode.eAddFormation
                HandleAddFormation(Data)
            Case GlobalMessageCode.eRemoveFormation
                HandleRemoveFormation(Data)
            Case GlobalMessageCode.eUpdateSlotStates
                HandleUpdateSlotStates(Data)
            Case GlobalMessageCode.eUpdateResearcherCnt
                HandleUpdateResearchCnt(Data)
            Case GlobalMessageCode.eGetItemIntelDetail
                HandleGetItemIntelDetail(Data)
            Case GlobalMessageCode.eAlertDestinationReached
                HandleAlertDestArrived(Data)
            Case GlobalMessageCode.eGetIntelSellOrderDetail
                HandleGetIntelSellOrderDetail(Data)
            Case GlobalMessageCode.eGetColonyResearchList
                Dim oWin As frmColonyResearch = CType(goUILib.GetWindow("frmColonyResearch"), frmColonyResearch)
                If oWin Is Nothing = False Then oWin.HandleGetColonyResearchList(Data)
            Case GlobalMessageCode.eAgentMissionCompleted
                HandleAgentMissionCompleted(Data)
            Case GlobalMessageCode.eUpdatePlayerTimer
                HandleUpdatePlayerTimer(Data)
            Case GlobalMessageCode.ePirateWaveSpawn
                HandlePirateWaveSpawn(Data)
            Case GlobalMessageCode.eUpdateMOTD
                HandleUpdateMOTD(Data)
            Case GlobalMessageCode.eUpdateGuildRank
                HandleUpdateGuildRank(Data)
            Case GlobalMessageCode.eUpdateRankPermission
                HandleUpdateRankPermission(Data)
            Case GlobalMessageCode.eGuildMemberStatus
                HandleGuildMemberStatus(Data)
            Case GlobalMessageCode.eProposeGuildVote
                HandleProposeGuildVote(Data)
            Case GlobalMessageCode.eSearchGuilds
                If goUILib Is Nothing = False Then
                    Dim oWin As frmGuildSearch = CType(goUILib.GetWindow("frmGuildSearch"), frmGuildSearch)
                    If oWin Is Nothing = False Then oWin.HandleSearchResults(Data)
                End If
            Case GlobalMessageCode.eAddEventAttachment
                HandleAddEventAttachment(Data)
            Case GlobalMessageCode.eUpdateGuildRecruitment
                HandleUpdateGuildRecruitment(Data)
            Case GlobalMessageCode.eRemoveEventAttachment
                HandleRemoveEventAttachment(Data)
            Case GlobalMessageCode.eUpdateGuildEvent
                HandleUpdateGuildEvent(Data)
            Case GlobalMessageCode.eSetGuildRel
                HandleSetGuildRel(Data)
            Case GlobalMessageCode.eAdvancedEventConfig
                HandleAdvancedEventConfig(Data)
            Case GlobalMessageCode.eCheckGuildName
                HandleCheckGuildName(Data)
            Case GlobalMessageCode.eInviteFormGuild
                HandleInviteFormGuild(Data)
            Case GlobalMessageCode.eGuildRequestDetails
                If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then goCurrentPlayer.oGuild.HandleGuildRequestDetails(Data)
            Case GlobalMessageCode.eGetGuildInvites
                If goUILib Is Nothing = False Then
                    Dim ofrm As frmGuildSearch = CType(goUILib.GetWindow("frmGuildSearch"), frmGuildSearch)
                    If ofrm Is Nothing = False Then ofrm.HandleInvitedGuildList(Data)
                End If
            Case GlobalMessageCode.eInvitePlayerToGuild
                HandleInvitePlayerToGuild(Data)
            Case GlobalMessageCode.eGetMyVoteValue
                HandleGetMyVoteValue(Data)
            Case GlobalMessageCode.eRequestGuildEvents
                HandleRequestGuildEvents(Data)
            Case GlobalMessageCode.eUpdateGuildTreasury
                HandleUpdateGuildTreasury(Data)
            Case GlobalMessageCode.eGetGuildAssets
                TradePostContents.HandleGetGuildAssets(Data)
            Case GlobalMessageCode.eGetSenateObjectDetails
                If goUILib Is Nothing = False Then
                    Dim ofrm As frmSenate = CType(goUILib.GetWindow("frmSenate"), frmSenate)
                    If ofrm Is Nothing = False Then ofrm.HandleGetSenateObjectDetails(Data)
                End If
            Case GlobalMessageCode.eAddSenateProposal
                If goUILib Is Nothing = False Then
                    Dim ofrm As frmSenate = CType(goUILib.GetWindow("frmSenate"), frmSenate)
                    If ofrm Is Nothing = False Then ofrm.HandleAddSenateProposal(Data)
                End If
            Case GlobalMessageCode.eGuildCreationUpdate
                HandleGuildCreationUpdate(Data)
            Case GlobalMessageCode.eGuildCreationAcceptance
                If goUILib Is Nothing = False Then
                    Dim ofrm As frmGuildSetup = CType(goUILib.GetWindow("frmGuildSetup"), frmGuildSetup)
                    If ofrm Is Nothing = False Then ofrm.HandleGuildCreationAcceptance(Data)
                End If
            Case GlobalMessageCode.eRequestChannelDetails
                If goUILib Is Nothing = False Then
                    Dim ofrm As frmChannelConfig = CType(goUILib.GetWindow("frmChannelConfig"), frmChannelConfig)
                    If ofrm Is Nothing = False Then ofrm.HandleChannelDetails(Data)
                End If
            Case GlobalMessageCode.eRequestChannelList
                If goUILib Is Nothing = False Then
                    Dim ofrm As frmChannels = CType(goUILib.GetWindow("frmChannels"), frmChannels)
                    If ofrm Is Nothing = False Then ofrm.HandleRequestChannelList(Data)
                End If
            Case GlobalMessageCode.eGetImposedAgentEffects
                If goUILib Is Nothing = False Then
                    Dim oFrm As frmAgentMain = CType(goUILib.GetWindow("frmAgentMain"), frmAgentMain)
                    If oFrm Is Nothing = False Then oFrm.HandleGetImposedAgentEffects(Data)
                End If
            Case GlobalMessageCode.eUpdatePlayerCustomTitle
                HandleUpdatePlayerCustomTitle(Data)
            Case GlobalMessageCode.eMineralBid
                If goUILib Is Nothing = False Then
                    Dim oFrm As frmMineralBid = CType(goUILib.GetWindow("frmMineralBid"), frmMineralBid)
                    If oFrm Is Nothing = False Then oFrm.HandleMineralBidData(Data)
                End If
            Case GlobalMessageCode.eRequestBidStatus
                If goUILib Is Nothing = False Then
                    Dim oFrm As frmMining = CType(goUILib.GetWindow("frmMining"), frmMining)
                    If oFrm Is Nothing = False Then oFrm.HandleMiningBidListMsg(Data)
                End If
            Case GlobalMessageCode.eOutBidAlert
                HandleOutbidAlert(Data)
            Case GlobalMessageCode.eShiftClickAddProduction
                HandleShiftClickAddProduction(Data)
            Case GlobalMessageCode.eUpdatePlanetOwnership
                HandleUpdatePlanetOwnership(Data)
            Case GlobalMessageCode.eGetGuildBillboards
                HandleGetGuildBillboards(Data)
            Case GlobalMessageCode.eGetSpecTechGuaranteeList
                If goUILib Is Nothing = False Then
                    Dim oFrm As frmSpecTechSelect = CType(goUILib.GetWindow("frmSpecTechSelect"), frmSpecTechSelect)
                    If oFrm Is Nothing = False Then oFrm.HandleSpecialTechListMsg(Data)
                End If
            Case GlobalMessageCode.eUpdatePlayerDetails
                HandleUpdatePlayerDetails(Data)
            Case GlobalMessageCode.eClaimItem
                HandleClaimables(Data)
            Case GlobalMessageCode.eSetSpecTechGuarantee
                Dim oFrm As New frmSpecTechSelect(goUILib)
                oFrm.Visible = True
            Case GlobalMessageCode.eRequestObject
                HandleRequestObjectResponse(Data)
            Case GlobalMessageCode.eGetCPPenaltyList
                If goUILib Is Nothing = False Then
                    Dim oWin As frmEnvirDisplay = CType(goUILib.GetWindow("frmEnvirDisplay"), frmEnvirDisplay)
                    If oWin Is Nothing = False Then oWin.HandleGetCPPenaltyMsg(Data)
                End If
            Case GlobalMessageCode.eDeleteFleet
                HandleDeleteFleet(Data)
            Case GlobalMessageCode.eGetRouteTemplates
                frmRouteTemplate.HandleRequestTemplates(Data)
            Case GlobalMessageCode.eSenateStatusReport
                HandleSenateStatusReport(Data)
            Case GlobalMessageCode.eRequestMaxWpnRng
                HandleRequestMaxWpnRng(Data)
            Case GlobalMessageCode.eRequestColonyTransportPotentials
                frmAddTransport.HandleRequestColonyTransportPotentials(Data)
            Case GlobalMessageCode.eRequestTransportName
                frmTransportManagement.SetTransportName(Data)
            Case GlobalMessageCode.eRequestTransportRouteDestList
                frmTransportOrders.HandleTransportRouteDestList(Data)
            Case GlobalMessageCode.eSetTransportStatus
                frmTransportManagement.HandleSetTransportStatus(Data)
            Case GlobalMessageCode.eAddTransport
                frmTransportManagement.HandleAddTransport(Data)
            Case GlobalMessageCode.eRequestTransports
                frmTransportManagement.HandleRequestTransports(Data)
            Case GlobalMessageCode.eRequestTransportDetails
                Dim lID As Int32 = System.BitConverter.ToInt32(Data, 2)
                Dim oTrans As Transport = frmTransportManagement.GetTransport(lID)
                If oTrans Is Nothing = False Then oTrans.HandleTransportDetails(Data)
            Case GlobalMessageCode.eDecommissionTransport
                frmTransportManagement.HandleDecommissionTransport(Data)
            Case GlobalMessageCode.eGetDAValues
                HandleGetDAValues(Data)
            Case GlobalMessageCode.eRespawnSelection
                HandleRespawnSelection(Data)
            Case GlobalMessageCode.eWPTimers
                HandleWPTimers(Data)
            Case GlobalMessageCode.eGetGuildShareAssets
                If goUILib Is Nothing = False Then
                    Dim oWin As frmGuildMain = CType(goUILib.GetWindow("frmGuildMain"), frmGuildMain)
                    If oWin Is Nothing = False Then oWin.HandleGetGuildShareAssets(Data)
                End If
            Case GlobalMessageCode.eRequestUniversalAssets
                If goUILib Is Nothing = False Then
                    Dim oWin As frmCommand = CType(goUILib.GetWindow("frmCommand"), frmCommand)
                    If oWin Is Nothing = False Then oWin.HandleUniverseInventory(Data)
                End If
        End Select

    End Sub

    Private Sub moPrimary_onDisconnect(ByVal Index As Integer) Handles moPrimary.onDisconnect
        'TODO: if it was intentional, then we do nothing... if it was not... then we notify the user
    End Sub

    Private Sub moPrimary_onError(ByVal Index As Integer, ByVal Description As String) Handles moPrimary.onError
        If Description.ToUpper = "AN EXISTING CONNECTION WAS FORCIBLY CLOSED BY THE REMOTE HOST" OrElse Description.ToUpper.StartsWith("CANNOT ACCESS A DISPOSED OBJECT.") = True Then
            'ok, bad... means the connection has been disconnected
            Dim bResult As Boolean = False
            If (mlBusyStatus And elBusyStatusType.PrimaryServerLogin) <> 0 Then
                MsgBox("Connection to the server has been lost and could not be re-established." & vbCrLf & _
                 "This may either mean the server is down or something is wrong with your" & vbCrLf & _
                 "connection to the internet. Please try again later.", MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Connection Lost")
                If mlBusyStatus <> 0 Then
                    mlBusyStatus = 0
                Else
                    frmMain.Close()
                    Application.Exit()
                End If
            End If
        End If

		If mbClosing = False Then
			If (mlBusyStatus And elBusyStatusType.PlayerDetails) <> 0 Then
				If mbReRequestInProgress = False Then
					mbReRequestInProgress = True
					WriteToRPDFile("  Retrying " & mlLastRequestedType)
					Dim oThread As New Threading.Thread(AddressOf __ThreadedReRequestPD)
					oThread.Start()
				End If
				Return
			End If

			If goUILib Is Nothing = False Then goUILib.AddNotification("GAME ERROR (" & msCurrentLoc & "): " & Description, System.Drawing.Color.FromArgb(255, 255, 0, 0), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		End If
    End Sub

#End Region

#Region " Regional Server Events "

	Private Sub moRegion_onConnect(ByVal Index As Integer) Handles moRegion.onConnect
		'mbConnectedToRegion = True
		'mbConnectingToRegion = False
		RaiseEvent ConnectedToServer(eyConnType.RegionServer)
	End Sub

    Private Sub moRegion_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moRegion.onConnectionRequest
        '
    End Sub

    Private Sub moRegion_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moRegion.onDataArrival
        Dim iMsgID As Int16

        msCurrentLoc = "rda"

		'l_REGION_MSGS += 1

        iMsgID = System.BitConverter.ToInt16(Data, 0)
        'Debug.WriteLine("Region.MsgID: " & iMsgID)
		If Data.Length < 2 Then Return
        'goUILib.AddNotification("Region: " & CType(iMsgID, GlobalMessageCode).ToString, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        If mbDebugMessages = True Then Debug.Print("moRegion_onDataArrival: " & CType(iMsgID, GlobalMessageCode).ToString)
		Select Case iMsgID
			Case GlobalMessageCode.eFireWeapon
				HandleFireWeaponMsg(Data, True)
			Case GlobalMessageCode.eFireWeaponMiss
                HandleFireWeaponMsg(Data, False)
            Case GlobalMessageCode.ePDS_AtMissileHit, GlobalMessageCode.ePDS_AtMissileMiss, GlobalMessageCode.ePDS_AtUnitHit, GlobalMessageCode.ePDS_AtUnitMiss
                HandlePDSShot(Data)
            Case GlobalMessageCode.eBurstEnvironmentResponse
                If goCurrentEnvir Is Nothing = False Then
                    'goCurrentEnvir.ProcessRequestEnvironmentResponse(Data)
                    RaiseEvent ProcessedBurstEnvironment()
                End If
			Case GlobalMessageCode.eMoveObjectCommand, GlobalMessageCode.eEntityChangeEnvironment, GlobalMessageCode.eFinalMoveCommand
				HandleMoveObjectCommand(Data)
			Case GlobalMessageCode.eStopObjectCommand
				HandleStopObjectCommand(Data)
			Case GlobalMessageCode.eSetEntityTarget
				HandleSetEntityTargetMessage(Data)
			Case GlobalMessageCode.eMissileFired
				HandleMissileFired(Data)
			Case GlobalMessageCode.eMissileImpact
                'HandleMissileImpact(Data)
			Case GlobalMessageCode.eMissileDetonated
				HandleMissileDetonated(Data)
			Case GlobalMessageCode.eAddObjectCommand, GlobalMessageCode.eAddObjectCommand_CE
				HandleAddObjectMsg(Data, False)
			Case GlobalMessageCode.eUnitHPUpdate
				HandleUnitHPUpdateMsg(Data)
			Case GlobalMessageCode.eRemoveObject
				HandleRemoveObject(Data)
			Case GlobalMessageCode.eGetEntityDetails
				HandleGetEntityDetailsMsg(Data)
			Case GlobalMessageCode.eRequestDockFail
				HandleRequestDockFail(Data)
			Case GlobalMessageCode.eRequestUndock
                HandleAddObjectMsg(Data, False)
			Case GlobalMessageCode.eBombardFireMsg
				HandleBombFireMsg(Data)
			Case GlobalMessageCode.eChatMessage
				HandleChatMsg(Data)
			Case GlobalMessageCode.eUpdateEntityAttrs
				HandleUpdateEntityAttrs(Data)
			Case GlobalMessageCode.eUpdateCommandPoints
				HandleCommandPointUpdate(Data)
			Case GlobalMessageCode.eFleetInterSystemMoving
				HandleFleetInterSystemMoving(Data)
			Case GlobalMessageCode.eSetRepairTarget, GlobalMessageCode.eSetDismantleTarget
				HandleSetMaintenanceTarget(Data, iMsgID)
			Case GlobalMessageCode.eMoveObjectRequestDeny
				HandleMoveObjectRequestDeny(Data)
			Case GlobalMessageCode.eMoveFormation
				HandleMoveFormation(Data)
			Case GlobalMessageCode.eSetIronCurtain
                HandleSetIronCurtain(Data)
            Case GlobalMessageCode.eSetEntityStatus
                HandleRegionSetEntityStatus(Data)
            Case GlobalMessageCode.ePlayerIsDead
                HandleOtherPlayerDeath(Data)
        End Select
    End Sub

    Private Sub moRegion_onDisconnect(ByVal Index As Integer) Handles moRegion.onDisconnect
        '
    End Sub

    Private Sub moRegion_onError(ByVal Index As Integer, ByVal Description As String) Handles moRegion.onError
        'On Error Resume Next
        'moPrimary.Disconnect()
        'moRegion.Disconnect()
        'mbConnectingToPrimary = False
        'mbConnectingToRegion = False
        If mbClosing = False Then
            If goUILib Is Nothing = False Then goUILib.AddNotification("GAME ERROR (" & msCurrentLoc & "): " & Description, System.Drawing.Color.FromArgb(255, 255, 0, 0), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

#End Region

#Region " Outgoing Message Handlers "
    Public Sub SendSetMaintenanceTarget(ByVal iMsgCode As Int16, ByRef oObject As BaseEntity, ByRef oTarget As BaseEntity)
        SendClearEntityProdQueue(oObject)

        msCurrentLoc = "SendSetMaintenanceTarget"
        If HasAliasedRights(AliasingRights.eMoveUnits) = False Then
            goUILib.AddNotification("You lack rights to move units.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If

        'Also verify Dismantle rights
        If iMsgCode = GlobalMessageCode.eSetDismantleTarget Then
            If HasAliasedRights(AliasingRights.eDismantle) = False Then
                goUILib.AddNotification("You lack rights to dismantle units.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
                Return
            End If
        End If

        Dim yMsg(13) As Byte

        System.BitConverter.GetBytes(iMsgCode).CopyTo(yMsg, 0)
        oObject.GetGUIDAsString.CopyTo(yMsg, 2)
        oTarget.GetGUIDAsString.CopyTo(yMsg, 8)

        If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, SoundMgr.SpeechType.eMoveRequest)

        SendToRegion(yMsg)
    End Sub

    Public Sub RequestEntityHPUpdate(ByVal lObjectID As Int32, ByVal iObjTypeID As Int16)
        Dim yData(7) As Byte

        msCurrentLoc = "RequestEntityHPUpdate"

        'TODO: A fix to the bombard exploit for bombarding the server with requests
        'Dim oEnvir As EpicaEnvironment = goCurrentEnvir
        'Dim lCnt As Int32 = 0
        'Dim lPos As Int32 = 0
        'If oEnvir Is Nothing = False Then
        '    For X As Int32 = 0 To oEnvir.lEntityUB
        '        If oEnvir.lEntityIdx(X) <> -1 AndAlso oEnvir.oEntity(X).bSelected = True Then
        '            lCnt += 1
        '        End If
        '    Next X
        '    ReDim yData((lCnt * 6) + 3)
        '    System.BitConverter.GetBytes(EpicaMessageCode.eUnitHPUpdate).CopyTo(yData, lPos) : lPos += 2
        '    System.BitConverter.GetBytes(CShort(lCnt)).CopyTo(yData, lPos) : lPos += 2
        '    For X As Int32 = 0 To oEnvir.lEntityUB
        '        If oEnvir.lEntityIdx(X) <> -1 AndAlso oEnvir.oEntity(X).bSelected = True Then
        '            oEnvir.oEntity(X).GetGUIDAsString.CopyTo(yData, lPos) : lPos += 6
        '        End If
        '    Next X
        'End If

        System.BitConverter.GetBytes(GlobalMessageCode.eUnitHPUpdate).CopyTo(yData, 0)
        System.BitConverter.GetBytes(lObjectID).CopyTo(yData, 2)
        System.BitConverter.GetBytes(iObjTypeID).CopyTo(yData, 6)

        moRegion.SendData(yData)
    End Sub

	Public Sub SendSetPrimaryTarget(ByRef oTarget As Base_GUID, ByVal bFromConfirm As Boolean)
		Dim yData(7) As Byte
		Dim X As Int32
		Dim lLen As Int32

		If goCurrentEnvir Is Nothing Then Exit Sub
		If HasAliasedRights(AliasingRights.eMoveUnits) = False Then Return

		msCurrentLoc = "SendSetPrimaryTarget"

		If NewTutorialManager.TutorialOn = True Then
			If goUILib Is Nothing = False AndAlso goUILib.CommandAllowed(True, "AssignTarget") = False Then Return
		End If

		System.BitConverter.GetBytes(GlobalMessageCode.eSetPrimaryTarget).CopyTo(yData, 0)
		'now the target
		oTarget.GetGUIDAsString.CopyTo(yData, 2)

		lLen = 8

		Dim lCnt As Int32 = 0
		'Now, any items selected
		For X = 0 To goCurrentEnvir.lEntityUB
			If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
				If goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    If goCurrentEnvir.oEntity(X).bProducing = True AndAlso goCurrentPlayer.yPlayerPhase <> 0 Then
                        If bFromConfirm = False AndAlso goUILib Is Nothing = False AndAlso muSettings.bDoNotShowEngineerCancelAlert = False Then
                            Dim ofrm As New frmCancelBuild(goUILib)
                            ofrm.SetOrder(frmCancelBuild.eyOrderType.SetPrimaryTarget, -1, -1, oTarget.ObjectID, oTarget.ObjTypeID, -1, -1, -1)
                            ofrm.Visible = True
                            Return
                        End If
                    End If

                    SendClearEntityProdQueue(goCurrentEnvir.oEntity(X))


					ReDim Preserve yData(lLen + 6)
					goCurrentEnvir.oEntity(X).GetGUIDAsString.CopyTo(yData, lLen)
					lLen += 6
					lCnt += 1
				End If
			End If
		Next X

		If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, goSound.GetUnitSpeechTypeCnt(SoundMgr.SpeechType.eSetPrimaryTarget, lCnt))

        moRegion.SendData(yData)

        BPMetrics.MetricMgr.AddActivity(BPMetrics.eOrders.eSetPrimaryTarget, lCnt, 0)
	End Sub

    Public Sub SendDockRequest(ByRef oDockee As Base_GUID)
        Dim yData(7) As Byte
        Dim lLen As Int32
        Dim X As Int32
        msCurrentLoc = "SendDockRequest"

        If HasAliasedRights(AliasingRights.eDockUndockUnits) = False Then
            goUILib.AddNotification("You lack rights to dock/undock units.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If

		If NewTutorialManager.TutorialOn = True Then
			If goUILib Is Nothing = False AndAlso goUILib.CommandAllowed(True, "OrderDock") = False Then Return
		End If

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestDock).CopyTo(yData, 0)
        'Now, the target GUID
        oDockee.GetGUIDAsString.CopyTo(yData, 2)

        'Ok, before moving on... we do our own validation here to reduce server load
        If oDockee.ObjTypeID = ObjectType.eUnit OrElse oDockee.ObjTypeID = ObjectType.eFacility Then
            With CType(oDockee, BaseEntity)
                If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                    goUILib.AddNotification("Docking Request Rejected: Unable to dock with moving target", Color.FromArgb(255, 255, 255, 0), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                    Return
                End If

                'TODO: Add other possible checks here... like door size
            End With
        End If

        'Note: if its not a unit or facility, then we just leave it to the server to decide...

        lLen = 8

        'Now, any items selected
        Dim lCnt As Int32 = 0
        For X = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
                If oDockee.ObjectID <> goCurrentEnvir.oEntity(X).ObjectID OrElse oDockee.ObjTypeID <> goCurrentEnvir.oEntity(X).ObjTypeID Then
					If goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
						ReDim Preserve yData(lLen + 6)
                        goCurrentEnvir.oEntity(X).GetGUIDAsString.CopyTo(yData, lLen)

                        SendClearEntityProdQueue(goCurrentEnvir.oEntity(X))

						lCnt += 1
						lLen += 6
					End If
                End If
            End If
        Next X
        If lCnt = 0 Then Return
        If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, goSound.GetUnitSpeechTypeCnt(SoundMgr.SpeechType.eDockRequest, lCnt))

        moRegion.SendData(yData)

        BPMetrics.MetricMgr.AddActivity(BPMetrics.eOrders.eDockRequest, lCnt, 0)
    End Sub

#Region " Single Object Messages (Obsolete) "
    'Public Sub SendMoveObjectRequest(ByVal oUnit As EpicaEntity, ByVal lDestX As Int32, ByVal lDestZ As Int32, ByVal iDestAngle As Int16, ByVal lDestID As Int32, ByVal iDestTypeID As Int16)
    '    Dim yData(23) As Byte
    '    msCurrentLoc = "SendMORequest"
    '    System.BitConverter.GetBytes(EpicaMessageCode.eMoveObjectRequest).CopyTo(yData, 0)
    '    oUnit.GetGUIDAsString.CopyTo(yData, 2)
    '    System.BitConverter.GetBytes(lDestX).CopyTo(yData, 8)
    '    System.BitConverter.GetBytes(lDestZ).CopyTo(yData, 12)
    '    System.BitConverter.GetBytes(iDestAngle).CopyTo(yData, 16)
    '    System.BitConverter.GetBytes(lDestID).CopyTo(yData, 18)
    '    System.BitConverter.GetBytes(iDestTypeID).CopyTo(yData, 22)
    '    moRegion.SendData(yData)
    '    Erase yData
    'End Sub

    'Public Sub SendAddWaypointRequest(ByVal oUnit As EpicaEntity, ByVal lDestX As Int32, ByVal lDestZ As Int32, ByVal iDestAngle As Int16, ByVal lDestID As Int32, ByVal iDestTypeID As Int16)
    '    Dim yData(23) As Byte

    '    msCurrentLoc = "SendAddWaypointRequest"

    '    System.BitConverter.GetBytes(EpicaMessageCode.eAddWaypointMsg).CopyTo(yData, 0)
    '    oUnit.GetGUIDAsString.CopyTo(yData, 2)
    '    System.BitConverter.GetBytes(lDestX).CopyTo(yData, 8)
    '    System.BitConverter.GetBytes(lDestZ).CopyTo(yData, 12)
    '    System.BitConverter.GetBytes(iDestAngle).CopyTo(yData, 16)
    '    System.BitConverter.GetBytes(lDestID).CopyTo(yData, 18)
    '    System.BitConverter.GetBytes(iDestTypeID).CopyTo(yData, 22)
    '    moRegion.SendData(yData)
    '    Erase yData
    'End Sub
#End Region

    Public Sub RequestPlayerDetails()
		Dim yData(15) As Byte

		msCurrentLoc = "RequestPlayerDetails"

		Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
		If sPath.EndsWith("\") = False Then sPath &= "\"
		If Exists(sPath & "RPD.txt") = True Then
			FileCopy(sPath & "RPD.txt", sPath & "BadRPD_" & Now.ToString("MMddyyyyhhmmss") & ".txt")
			KillRPDFile()
		End If

		WriteToRPDFile("RPD Begin: " & Now.ToString("MM/dd/yyyy hh:mm:ss"))

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerDetails).CopyTo(yData, 0)
		System.BitConverter.GetBytes(glPlayerID).CopyTo(yData, 2)
		System.BitConverter.GetBytes(elPlayerDetailsType.ePlayerObject).CopyTo(yData, 6)
		System.BitConverter.GetBytes(-1I).CopyTo(yData, 10)
		System.BitConverter.GetBytes(-1S).CopyTo(yData, 14)

		mlLastAddObjectID = -1
		miLastAddObjectTypeID = -1
		Dim lLastAddID As Int32 = -1
		Dim iLastAddTypeID As Int16 = -1
		Dim lCntr As Int32 = 0
		mlBusyStatus = mlBusyStatus Or elBusyStatusType.PlayerDetails
		SendToPrimary(yData)
		'While (mlBusyStatus And elBusyStatusType.PlayerDetails) <> 0
		'	'Threading.Thread.Sleep(10)
		'	Application.DoEvents()
		'	'Threading.Thread.Sleep(0)
		'	Threading.Thread.Sleep(1)

		'	If lLastAddID = mlLastAddObjectID AndAlso miLastAddObjectTypeID = iLastAddTypeID Then
		'		lCntr += 1
		'		If lCntr > 5000 Then
		'			lCntr = 0
		'			mbReRequestInProgress = True
		'			Dim oThread As New Threading.Thread(AddressOf __ThreadedReRequestPD)
		'			oThread.Start()
		'		End If
		'	Else
		'		lCntr = 0
		'		lLastAddID = mlLastAddObjectID
		'		iLastAddTypeID = miLastAddObjectTypeID
		'	End If
		'End While

		''ok, we're here, so kill our rdpfile
		'KillRPDFile()

    End Sub

    Public Sub RequestGalaxyAndSystems()
        'in this routine, we are going to send to the primary that we want the Galaxy and System layout...
        Dim yData(1) As Byte    'for right now, we send just the message itself...
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestGalaxyAndSystems).CopyTo(yData, 0)

        msCurrentLoc = "RequestGalaxyAndSystems"

		mlBusyStatus = mlBusyStatus Or elBusyStatusType.GalaxyAndSystems
        SendToPrimary(yData)
		'While (mlBusyStatus And elBusyStatusType.GalaxyAndSystems) <> 0
		'	'Threading.Thread.Sleep(10)
		'	Application.DoEvents()
		'	'Threading.Thread.Sleep(0)
		'	Threading.Thread.Sleep(1)
		'End While
    End Sub

    Public Sub RequestStarTypes()
        'this needs to be called PRETTY early in application life
        Dim yData(1) As Byte    'for right now, we send just the message itself...
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestStarTypes).CopyTo(yData, 0)

        msCurrentLoc = "RequestStarTypes"

		mlBusyStatus = mlBusyStatus Or elBusyStatusType.StarTypes
        SendToPrimary(yData)
		While (mlBusyStatus And elBusyStatusType.StarTypes) <> 0
			'Threading.Thread.Sleep(10)
			Application.DoEvents()
			'Threading.Thread.Sleep(0)
			Threading.Thread.Sleep(1)
		End While
    End Sub

    Public Sub RequestSystemDetails(ByVal lSystemID As Int32)
        Dim yData(5) As Byte

        msCurrentLoc = "RequestSystemDetails"

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestSystemDetails).CopyTo(yData, 0)
        System.BitConverter.GetBytes(lSystemID).CopyTo(yData, 2)
		mlBusyStatus = mlBusyStatus Or elBusyStatusType.SystemDetails
        SendToPrimary(yData)
		While (mlBusyStatus And elBusyStatusType.SystemDetails) <> 0
			'Threading.Thread.Sleep(10)
			Application.DoEvents()
			'Threading.Thread.Sleep(0)
			Threading.Thread.Sleep(1)
		End While
	End Sub

	Private Function EncBytes(ByVal yBytes() As Byte) As Byte()
		Dim lLen As Int32 = UBound(yBytes)
		Dim lKey As Int32
		Dim lOffset As Int32
		Dim X As Int32
		Dim lChrCode As Int32
		Dim lMod As Int32

		Dim yFinal(lLen + 1) As Byte

        lKey = CInt(Math.Floor((Rnd() * 51)))

        VBMath.Rnd(-1)
        Call Randomize(777 + lKey)

        yFinal(0) = CByte(lKey)
        For X = 0 To lLen
            lOffset = X - lLen
            'now, found out what we got here..
            lChrCode = yBytes(X)
            lMod = CInt(Math.Floor((Rnd() * 5) + 1))
            lChrCode = lChrCode + lMod
            If lChrCode > 255 Then lChrCode = lChrCode - 256
            yFinal(X + 1) = CByte(lChrCode)
        Next X

        Return yFinal

    End Function

    Public Function RequestLogin(ByVal sUserName As String, ByVal sPassword As String, ByVal bAlias As Boolean, ByVal sAliasName As String, ByVal sAliasPassword As String, ByVal yType As eyConnType) As Boolean 'As Int32
        Dim yData() As Byte

        msCurrentLoc = "RequestLogin"

        'Ok, so here, we encrypt our password's
        Dim oEnc As New StrEncDec
        Dim yEncPassword() As Byte = oEnc.Encrypt(System.Text.ASCIIEncoding.ASCII.GetBytes(sPassword))
        Dim yEncAliasPW() As Byte = oEnc.Encrypt(System.Text.ASCIIEncoding.ASCII.GetBytes(sAliasPassword))

        ReDim yData(84) 'all entries should be 0
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.ePlayerLoginRequest).CopyTo(yData, lPos) : lPos += 2
        oEnc.Encrypt(System.Text.ASCIIEncoding.ASCII.GetBytes(sUserName)).CopyTo(yData, lPos) : lPos += 20
 
        yEncPassword.CopyTo(yData, lPos) : lPos += 21

        If bAlias = True Then
            yData(lPos) = 1
        Else : yData(lPos) = 0
        End If
        lPos += 1
        oEnc.Encrypt(System.Text.ASCIIEncoding.ASCII.GetBytes(sAliasName)).CopyTo(yData, lPos) : lPos += 20
        yEncAliasPW.CopyTo(yData, lPos) : lPos += 21
 
        Dim lStatus As Int32
        If yType = eyConnType.OperatorServer Then
            lStatus = elBusyStatusType.OperatorServerLogin
            mlBusyStatus = mlBusyStatus Or lStatus
            moOperator.SendData(yData)
        ElseIf yType = eyConnType.PrimaryServer Then
            lStatus = elBusyStatusType.PrimaryServerLogin
            mlBusyStatus = mlBusyStatus Or lStatus
            moPrimary.SendData(yData)
        End If
 
        Return True
    End Function

    Public Sub SendChangeEnvironment(ByVal lID As Int32, ByVal iTypeID As Int16)
        Dim yData(11) As Byte

        If goCurrentEnvir Is Nothing = False Then
            goCurrentEnvir.SaveEnvironmentTacticalData()

        End If

        Dim ofrmED As frmEnvirDisplay = CType(goUILib.GetWindow("frmEnvirDisplay"), frmEnvirDisplay)
        If Not ofrmED Is Nothing Then
            ofrmED.ChangeEnviroments()
        End If

        msCurrentLoc = "SendChangeEnvironment"
        If goUILib Is Nothing = False Then goUILib.AddNotification("Establishing link with our commanders in that environment...", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        'Player ID, the new Envir ID, the new Envir Type ID
        System.BitConverter.GetBytes(GlobalMessageCode.eChangingEnvironment).CopyTo(yData, 0)
        System.BitConverter.GetBytes(glPlayerID).CopyTo(yData, 2)
        System.BitConverter.GetBytes(lID).CopyTo(yData, 6)
        System.BitConverter.GetBytes(iTypeID).CopyTo(yData, 10)

        'Send to the domain server..
        If moRegion Is Nothing = False Then moRegion.SendData(yData)

        'Now, send to the Primary...
        SendToPrimary(yData)

    End Sub

    Public Sub RequestEntityDetails(ByVal oObjGuid As Base_GUID, ByVal oEnvirGuid As Base_GUID)
        'TODO: the type ID may tell us a better server to request from, for now, request from region
        Dim yData(13) As Byte

        msCurrentLoc = "RequestEntityDetails"

        System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityDetails).CopyTo(yData, 0)
        oObjGuid.GetGUIDAsString.CopyTo(yData, 2)
        oEnvirGuid.GetGUIDAsString.CopyTo(yData, 8)

        If oObjGuid.ObjTypeID = ObjectType.eMineralCache Then
            moPrimary.SendData(yData)
        Else : moRegion.SendData(yData)
        End If


    End Sub

    Public Sub SendSetMiningLoc(ByVal lCacheID As Int32, ByVal iCacheTypeID As Int16)
        Dim yData(7) As Byte
        Dim X As Int32
        Dim lLen As Int32

        msCurrentLoc = "SendSetMiningLoc"

        If HasAliasedRights(AliasingRights.eMoveUnits) = False Then Return

        System.BitConverter.GetBytes(GlobalMessageCode.eSetMiningLoc).CopyTo(yData, 0)
        'Target Cache Object ID
        System.BitConverter.GetBytes(lCacheID).CopyTo(yData, 2)
        System.BitConverter.GetBytes(iCacheTypeID).CopyTo(yData, 6)
        lLen = 8

        'Now, any items selected
        Dim lCnt As Int32 = 0
        For X = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
                If goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    ReDim Preserve yData(lLen + 6)
                    goCurrentEnvir.oEntity(X).GetGUIDAsString.CopyTo(yData, lLen)

                    SendClearEntityProdQueue(goCurrentEnvir.oEntity(X))

                    lLen += 6
                    lCnt += 1
                End If
            End If
        Next X
        If lCnt = 0 Then Return
        If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, goSound.GetUnitSpeechTypeCnt(SoundMgr.SpeechType.eMoveRequest, lCnt))

        moRegion.SendData(yData)
    End Sub

    Public Sub SendSetGuildFacility(ByVal lBuildObjID As Int32, ByVal iBuildObjTypeID As Int16, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal iLocA As Int16)
        Dim yData(23) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eUpdateGuildTreasury).CopyTo(yData, lPos) : lPos += 2
        System.BitConverter.GetBytes(lBuildObjID).CopyTo(yData, lPos) : lPos += 4
        System.BitConverter.GetBytes(iBuildObjTypeID).CopyTo(yData, lPos) : lPos += 2
        goCurrentEnvir.GetGUIDAsString.CopyTo(yData, lPos) : lPos += 6
        System.BitConverter.GetBytes(lLocX).CopyTo(yData, lPos) : lPos += 4
        System.BitConverter.GetBytes(lLocZ).CopyTo(yData, lPos) : lPos += 4
        System.BitConverter.GetBytes(iLocA).CopyTo(yData, lPos) : lPos += 2

        moPrimary.SendData(yData)
    End Sub

    Public Sub SendQueueSetProductionMsg(ByRef oBuilder As Base_GUID, ByVal lBuildObjID As Int32, ByVal iBuildObjTypeID As Int16, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal iLocA As Int16)
        Dim yData(23) As Byte

        If NewTutorialManager.TutorialOn = True AndAlso goCurrentPlayer.yPlayerPhase = 0 Then
            SendSetProductionMsg(oBuilder, lBuildObjID, iBuildObjTypeID, lLocX, lLocZ, iLocA)
            Return
        End If

        msCurrentLoc = "SendQueueSetProductionMsg"

        If iLocA > 3600 Then iLocA -= 3600S
        If iLocA < 0 Then iLocA += 3600S

        If oBuilder Is Nothing Then Return
        With CType(oBuilder, BaseEntity)
            .bProducing = True
        End With

        System.BitConverter.GetBytes(GlobalMessageCode.eShiftClickAddProduction).CopyTo(yData, 0)
        oBuilder.GetGUIDAsString.CopyTo(yData, 2)
        System.BitConverter.GetBytes(lBuildObjID).CopyTo(yData, 8)
        System.BitConverter.GetBytes(iBuildObjTypeID).CopyTo(yData, 12)
        System.BitConverter.GetBytes(lLocX).CopyTo(yData, 14)
        System.BitConverter.GetBytes(lLocZ).CopyTo(yData, 18)
        System.BitConverter.GetBytes(iLocA).CopyTo(yData, 22)

        If oBuilder.ObjTypeID = ObjectType.eUnit Then
            If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, goSound.GetUnitSpeechTypeCnt(SoundMgr.SpeechType.eBuildingConstruction, 1))
        End If

        moRegion.SendData(yData)
    End Sub

    Public Sub SendSetProductionMsg(ByRef oBuilder As Base_GUID, ByVal lBuildObjID As Int32, ByVal iBuildObjTypeID As Int16, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal iLocA As Int16)
        Dim yData(23) As Byte

        msCurrentLoc = "SendSetProductionMsg"

        SendClearEntityProdQueue(CType(oBuilder, BaseEntity))

        If iLocA > 3600 Then iLocA -= 3600S
        If iLocA < 0 Then iLocA += 3600S

        If oBuilder Is Nothing Then Return
        With CType(oBuilder, BaseEntity)
            If .bProducing = True AndAlso goCurrentPlayer.yPlayerPhase <> 0 Then
                If goUILib Is Nothing = False Then
                    Dim ofrm As New frmCancelBuild(goUILib)
                    ofrm.SetOrder(frmCancelBuild.eyOrderType.SetProduction, .ObjectID, .ObjTypeID, lBuildObjID, iBuildObjTypeID, lLocX, lLocZ, iLocA)
                    ofrm.Visible = True
                    Return
                End If
            End If
            .bProducing = True
        End With

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eBuildOrderSent, lBuildObjID, iBuildObjTypeID, 1, "")
        End If

        System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yData, 0)
        oBuilder.GetGUIDAsString.CopyTo(yData, 2)
        System.BitConverter.GetBytes(lBuildObjID).CopyTo(yData, 8)
        System.BitConverter.GetBytes(iBuildObjTypeID).CopyTo(yData, 12)
        System.BitConverter.GetBytes(lLocX).CopyTo(yData, 14)
        System.BitConverter.GetBytes(lLocZ).CopyTo(yData, 18)
        System.BitConverter.GetBytes(iLocA).CopyTo(yData, 22)

        If oBuilder.ObjTypeID = ObjectType.eUnit Then
            If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, goSound.GetUnitSpeechTypeCnt(SoundMgr.SpeechType.eBuildingConstruction, 1))
        End If

        moRegion.SendData(yData)
    End Sub

    Public Sub SendBombardRequest(ByVal lPlanetID As Int32, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal yBombardType As Byte)
        Dim yMsg(18) As Byte

        msCurrentLoc = "SendBombardRequest"

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestOrbitalBombard).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(lPlanetID).CopyTo(yMsg, 6)
        System.BitConverter.GetBytes(lLocX).CopyTo(yMsg, 10)
        System.BitConverter.GetBytes(lLocZ).CopyTo(yMsg, 14)
        yMsg(18) = yBombardType

        If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, goSound.GetUnitSpeechTypeCnt(SoundMgr.SpeechType.eBombardRequest, 1))

        moRegion.SendData(yMsg)
    End Sub

    Public Sub SendBombardStop(ByVal lPlanetID As Int32)
        Dim yMsg(9) As Byte

        msCurrentLoc = "SendBombardStop"
        System.BitConverter.GetBytes(GlobalMessageCode.eStopOrbitalBombard).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(lPlanetID).CopyTo(yMsg, 6)

        If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, goSound.GetUnitSpeechTypeCnt(SoundMgr.SpeechType.eStopBombard, 1))

        moRegion.SendData(yMsg)
    End Sub

    Public Sub AddOwnerPlayerRel(ByVal lOwnerID As Int32)
        Dim yMsg(9) As Byte
        Dim oTmpRel As PlayerRel
        Dim sMsg As String
        msCurrentLoc = "AddOwnerPlayerRel"

        oTmpRel = goCurrentPlayer.GetPlayerRel(lOwnerID)

        If oTmpRel Is Nothing Then
            msCurrentLoc = "AddOwnerPlayerRel"
            System.BitConverter.GetBytes(GlobalMessageCode.eAddPlayerRelObject).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 2)
            System.BitConverter.GetBytes(lOwnerID).CopyTo(yMsg, 6)

            moPrimary.SendData(yMsg)

            sMsg = "Player relationship created (use F5 to view)."
        Else : sMsg = "That player relationship already exists (use F5 to view)."
        End If

        If goUILib Is Nothing = False Then
            goUILib.AddNotification(sMsg, Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Public Sub SubmitAlloyDesign(ByRef oResearcher As Base_GUID, ByVal sName As String, ByVal lMineral1ID As Int32, ByVal lMineral2ID As Int32, ByVal lMineral3ID As Int32, ByVal lMineral4ID As Int32, ByVal yProperty1ID As Byte, ByVal yProperty2ID As Byte, ByVal yProperty3ID As Byte, ByVal yVal1 As Byte, ByVal yVal2 As Byte, ByVal yVal3 As Byte, ByVal yResearchLevel As Byte, ByVal lTechID As Int32)
        Dim yMsg(56) As Byte
        Dim lPos As Int32 = 0
        msCurrentLoc = "SubmitAlloyDesign"
        System.BitConverter.GetBytes(GlobalMessageCode.eSubmitComponentPrototype).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(ObjectType.eAlloyTech).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lTechID).CopyTo(yMsg, lPos) : lPos += 4
        oResearcher.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        System.BitConverter.GetBytes(lMineral1ID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMineral2ID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMineral3ID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMineral4ID).CopyTo(yMsg, lPos) : lPos += 4
        yMsg(lPos) = yProperty1ID : lPos += 1
        yMsg(lPos) = yProperty2ID : lPos += 1
        yMsg(lPos) = yProperty3ID : lPos += 1
        'yMsg(lPos) = yHL_Flags : lPos += 1
        yMsg(lPos) = yVal1 : lPos += 1
        yMsg(lPos) = yVal2 : lPos += 1
        yMsg(lPos) = yVal3 : lPos += 1
        yMsg(lPos) = yResearchLevel : lPos += 1

        If sName.Length > 20 Then sName = Mid$(sName, 1, 20)
        System.Text.ASCIIEncoding.ASCII.GetBytes(sName).CopyTo(yMsg, lPos) : lPos += 20

        moPrimary.SendData(yMsg)
    End Sub

    'Public Sub SubmitEngineDesign(ByRef oResearcher As Epica_GUID, ByVal sName As String, ByVal lPower As Int32, ByVal lThrust As Int32, ByVal ySpeed As Byte, ByVal yManeuver As Byte, ByVal lStructBodyMineralID As Int32, ByVal lStructFrameMineralID As Int32, ByVal lStructMeldMineralID As Int32, ByVal lDriveBodyMineralID As Int32, ByVal lDriveFrameMineralID As Int32, ByVal lDriveMeldMineralID As Int32, ByVal lFuelCompMineralID As Int32, ByVal lFuelCatalystMineralID As Int32)
    Public Sub SubmitEngineDesign(ByRef oResearcher As Base_GUID, ByVal sName As String, ByVal lPower As Int32, ByVal lThrust As Int32, ByVal ySpeed As Byte, ByVal yManeuver As Byte, ByVal lStructBodyMineralID As Int32, ByVal lStructFrameMineralID As Int32, ByVal lStructMeldMineralID As Int32, ByVal lDriveBodyMineralID As Int32, ByVal lDriveFrameMineralID As Int32, ByVal lDriveMeldMineralID As Int32, ByVal yColorValue As Byte, ByVal lTechID As Int32, ByVal lHullReq As Int32, ByVal lPowerReq As Int32, ByVal blResCost As Int64, ByVal blResTime As Int64, ByVal blProdCost As Int64, ByVal blProdTime As Int64, ByVal lColonists As Int32, ByVal lEnlisted As Int32, ByVal lOfficers As Int32, ByVal lSpecMin1 As Int32, ByVal lSpecMin2 As Int32, ByVal lSpecMin3 As Int32, ByVal lSpecMin4 As Int32, ByVal lSpecMin5 As Int32, ByVal lSpecMin6 As Int32, ByVal yHullTypeID As Byte)
        Dim yMsg(145) As Byte '67) As Byte
        Dim bResult As Boolean = False
        Dim lAttempts As Int32 = 0
        msCurrentLoc = "SubmitEngineDesign"
        While bResult = False AndAlso lAttempts < 2
            bResult = True
            lAttempts += 1
            Try
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eSubmitComponentPrototype).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(ObjectType.eEngineTech).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(lTechID).CopyTo(yMsg, lPos) : lPos += 4
                oResearcher.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                System.BitConverter.GetBytes(lPower).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lThrust).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = ySpeed : lPos += 1
                yMsg(lPos) = yManeuver : lPos += 1
                System.BitConverter.GetBytes(lStructBodyMineralID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lStructFrameMineralID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lStructMeldMineralID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lDriveBodyMineralID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lDriveFrameMineralID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lDriveMeldMineralID).CopyTo(yMsg, lPos) : lPos += 4
                'System.BitConverter.GetBytes(lFuelCompMineralID).CopyTo(yMsg, 44)
                'System.BitConverter.GetBytes(lFuelCatalystMineralID).CopyTo(yMsg, 48)
                yMsg(lPos) = yColorValue : lPos += 1

                If sName.Length > 20 Then sName = Mid$(sName, 1, 20)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sName).CopyTo(yMsg, lPos) : lPos += 20

                System.BitConverter.GetBytes(lHullReq).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lPowerReq).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(blResCost).CopyTo(yMsg, lPos) : lPos += 8
                System.BitConverter.GetBytes(blResTime).CopyTo(yMsg, lPos) : lPos += 8
                System.BitConverter.GetBytes(blProdCost).CopyTo(yMsg, lPos) : lPos += 8
                System.BitConverter.GetBytes(blProdTime).CopyTo(yMsg, lPos) : lPos += 8
                System.BitConverter.GetBytes(lColonists).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lEnlisted).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lOfficers).CopyTo(yMsg, lPos) : lPos += 4

                System.BitConverter.GetBytes(lSpecMin1).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lSpecMin2).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lSpecMin3).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lSpecMin4).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lSpecMin5).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lSpecMin6).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = yHullTypeID : lPos += 1

                moPrimary.SendData(yMsg)
            Catch ex As Exception
                bResult = False
            End Try
        End While

    End Sub

    Public Sub SubmitShieldDesign(ByRef oResearcher As Base_GUID, ByVal sName As String, ByVal lMaxHP As Int32, ByVal lRechargeRate As Int32, ByVal lRechargeFreq As Int32, ByVal lProjectionHullSize As Int32, ByVal lCoilMineralID As Int32, ByVal lAcceleratorMineralID As Int32, ByVal lCasingMineralID As Int32, ByVal yColorValue As Byte, ByVal lTechID As Int32, ByVal lHullReq As Int32, ByVal lPowerReq As Int32, ByVal blResCost As Int64, ByVal blResTime As Int64, ByVal blProdCost As Int64, ByVal blProdTime As Int64, ByVal lColonists As Int32, ByVal lEnlisted As Int32, ByVal lOfficers As Int32, ByVal lSpecMin1 As Int32, ByVal lSpecMin2 As Int32, ByVal lSpecMin3 As Int32, ByVal lSpecMin4 As Int32, ByVal lSpecMin5 As Int32, ByVal lSpecMin6 As Int32, ByVal yHullTypeID As Byte)
        Dim yMsg(139) As Byte '62) As Byte
        Dim lPos As Int32 = 0
        msCurrentLoc = "SubmitShieldDesign"
        System.BitConverter.GetBytes(GlobalMessageCode.eSubmitComponentPrototype).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(ObjectType.eShieldTech).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lTechID).CopyTo(yMsg, lPos) : lPos += 4
        oResearcher.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        System.BitConverter.GetBytes(lMaxHP).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lRechargeRate).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lRechargeFreq).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lProjectionHullSize).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lCoilMineralID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lAcceleratorMineralID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lCasingMineralID).CopyTo(yMsg, lPos) : lPos += 4
        yMsg(lPos) = yColorValue : lPos += 1

        If sName.Length > 20 Then sName = Mid$(sName, 1, 20)
        System.Text.ASCIIEncoding.ASCII.GetBytes(sName).CopyTo(yMsg, lPos) : lPos += 20

        System.BitConverter.GetBytes(lHullReq).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lPowerReq).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(blResCost).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(blResTime).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(blProdCost).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(blProdTime).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(lColonists).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lEnlisted).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lOfficers).CopyTo(yMsg, lPos) : lPos += 4

        System.BitConverter.GetBytes(lSpecMin1).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSpecMin2).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSpecMin3).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSpecMin4).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSpecMin5).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSpecMin6).CopyTo(yMsg, lPos) : lPos += 4
        yMsg(lPos) = yHullTypeID : lPos += 1

        moPrimary.SendData(yMsg)
    End Sub

    Public Sub SubmitArmorDesign(ByRef oResearcher As Base_GUID, ByVal sName As String, ByVal yBeam As Byte, ByVal yImpact As Byte, ByVal yPiercing As Byte, ByVal yMagnetic As Byte, ByVal yToxic As Byte, ByVal yBurn As Byte, ByVal yRadar As Byte, ByVal lHullUsagePerPlate As Int32, ByVal lHPPerPlate As Int32, ByVal lOuterLayerMineralID As Int32, ByVal lMiddleLayerMineralID As Int32, ByVal lInnerLayerMineralID As Int32, ByVal lTechID As Int32, ByVal lHullReq As Int32, ByVal lPowerReq As Int32, ByVal blResCost As Int64, ByVal blResTime As Int64, ByVal blProdCost As Int64, ByVal blProdTime As Int64, ByVal lColonists As Int32, ByVal lEnlisted As Int32, ByVal lOfficers As Int32, ByVal lSpecMin1 As Int32, ByVal lSpecMin2 As Int32, ByVal lSpecMin3 As Int32, ByVal lSpecMin4 As Int32, ByVal lSpecMin5 As Int32, ByVal lSpecMin6 As Int32)
        Dim yMsg(136) As Byte
        Dim lPos As Int32 = 0
        msCurrentLoc = "SubmitArmorDesign"
        System.BitConverter.GetBytes(GlobalMessageCode.eSubmitComponentPrototype).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(ObjectType.eArmorTech).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lTechID).CopyTo(yMsg, lPos) : lPos += 4
        oResearcher.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        yMsg(lPos) = yBeam : lPos += 1
        yMsg(lPos) = yImpact : lPos += 1
        yMsg(lPos) = yPiercing : lPos += 1
        yMsg(lPos) = yMagnetic : lPos += 1
        yMsg(lPos) = yToxic : lPos += 1
        yMsg(lPos) = yBurn : lPos += 1
        yMsg(lPos) = yRadar : lPos += 1
        System.BitConverter.GetBytes(lHullUsagePerPlate).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lHPPerPlate).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lOuterLayerMineralID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMiddleLayerMineralID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lInnerLayerMineralID).CopyTo(yMsg, lPos) : lPos += 4

        If sName.Length > 20 Then sName = Mid$(sName, 1, 20)
        System.Text.ASCIIEncoding.ASCII.GetBytes(sName).CopyTo(yMsg, lPos) : lPos += 20

        System.BitConverter.GetBytes(lHullReq).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lPowerReq).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(blResCost).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(blResTime).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(blProdCost).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(blProdTime).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(lColonists).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lEnlisted).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lOfficers).CopyTo(yMsg, lPos) : lPos += 4

        System.BitConverter.GetBytes(lSpecMin1).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSpecMin2).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSpecMin3).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSpecMin4).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSpecMin5).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSpecMin6).CopyTo(yMsg, lPos) : lPos += 4

        moPrimary.SendData(yMsg)
    End Sub

    Public Sub SubmitRadarDesign(ByRef oResearcher As Base_GUID, ByVal sName As String, ByVal yWpnAcc As Byte, ByVal yScanRes As Byte, ByVal yOptRange As Byte, ByVal yMaxRange As Byte, ByVal yDisRes As Byte, ByVal lEmitterMineralID As Int32, ByVal lDetectionMineralID As Int32, ByVal lCollectionMineralID As Int32, ByVal lCasingMineralID As Int32, ByVal yType As Byte, ByVal yJamImmune As Byte, ByVal yJamStr As Byte, ByVal yJamTargets As Byte, ByVal yJamEffect As Byte, ByVal lTechID As Int32, ByVal lHullReq As Int32, ByVal lPowerReq As Int32, ByVal blResCost As Int64, ByVal blResTime As Int64, ByVal blProdCost As Int64, ByVal blProdTime As Int64, ByVal lColonists As Int32, ByVal lEnlisted As Int32, ByVal lOfficers As Int32, ByVal lSpecMin1 As Int32, ByVal lSpecMin2 As Int32, ByVal lSpecMin3 As Int32, ByVal lSpecMin4 As Int32, ByVal lSpecMin5 As Int32, ByVal lSpecMin6 As Int32)
        Dim yMsg(135) As Byte '59) As Byte
        Dim lPos As Int32 = 0
        msCurrentLoc = "SubmitRadarDesign"
        System.BitConverter.GetBytes(GlobalMessageCode.eSubmitComponentPrototype).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(ObjectType.eRadarTech).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lTechID).CopyTo(yMsg, lPos) : lPos += 4
        oResearcher.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        yMsg(lPos) = yWpnAcc : lPos += 1
        yMsg(lPos) = yScanRes : lPos += 1
        yMsg(lPos) = yOptRange : lPos += 1
        yMsg(lPos) = yMaxRange : lPos += 1
        yMsg(lPos) = yDisRes : lPos += 1
        System.BitConverter.GetBytes(lEmitterMineralID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lDetectionMineralID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lCollectionMineralID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lCasingMineralID).CopyTo(yMsg, lPos) : lPos += 4

        yMsg(lPos) = yType : lPos += 1
        yMsg(lPos) = yJamImmune : lPos += 1
        yMsg(lPos) = yJamStr : lPos += 1
        yMsg(lPos) = yJamTargets : lPos += 1
        yMsg(lPos) = yJamEffect : lPos += 1


        If sName.Length > 20 Then sName = Mid$(sName, 1, 20)
        System.Text.ASCIIEncoding.ASCII.GetBytes(sName).CopyTo(yMsg, lPos) : lPos += 20

        System.BitConverter.GetBytes(lHullReq).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lPowerReq).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(blResCost).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(blResTime).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(blProdCost).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(blProdTime).CopyTo(yMsg, lPos) : lPos += 8
        System.BitConverter.GetBytes(lColonists).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lEnlisted).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lOfficers).CopyTo(yMsg, lPos) : lPos += 4

        System.BitConverter.GetBytes(lSpecMin1).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSpecMin2).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSpecMin3).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSpecMin4).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSpecMin5).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lSpecMin6).CopyTo(yMsg, lPos) : lPos += 4

        moPrimary.SendData(yMsg)
    End Sub

    Public Sub SendTetherPoint(ByVal yAction As Byte)
        
        msCurrentLoc = "SendTetherPoint"
        Try
            If HasAliasedRights(AliasingRights.eMoveUnits) = False Then Return

            Dim lOrderCnt As Int32 = 0
            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                    If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                        lOrderCnt += 1
                    End If
                End If
            Next X

            If lOrderCnt = 0 Then Return
            Dim yTether(21 + (lOrderCnt * 6)) As Byte
            Dim lPos As Int32 = 0

            'set rally point
            System.BitConverter.GetBytes(GlobalMessageCode.eSetTetherPoint).CopyTo(yTether, 0) : lPos += 2
            System.BitConverter.GetBytes(lOrderCnt).CopyTo(yTether, lPos) : lPos += 4
            yTether(lPos) = yAction : lPos += 1

            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                    If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                        goCurrentEnvir.oEntity(X).GetGUIDAsString.CopyTo(yTether, lPos) : lPos += 6
                    End If
                End If
            Next X
            If yAction = 1 Then
                goUILib.AddNotification("Tether point set.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Else
                goUILib.AddNotification("Tether point cleared.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If

            'Now, send the command
            moRegion.SendData(yTether)
        Catch
        End Try
    End Sub

    Public Sub SendRallyPointMsg(ByVal lDestX As Int32, ByVal lDestZ As Int32, ByVal iDestA As Int16)
        Dim lDestID As Int32 = -1
        Dim iDestTypeID As Int16 = -1
        msCurrentLoc = "SendRallyPointMsg"
        If HasAliasedRights(AliasingRights.eMoveUnits) = False Then Return
        'Ok, determine the currently viewed environment...
        If goGalaxy.CurrentSystemIdx <> -1 Then
            If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx <> -1 AndAlso (glCurrentEnvirView = CurrentView.ePlanetMapView OrElse glCurrentEnvirView = CurrentView.ePlanetView) Then
                lDestID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx).ObjectID
                iDestTypeID = ObjectType.ePlanet
            Else
                lDestID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID
                iDestTypeID = ObjectType.eSolarSystem
            End If
        Else
            Return
        End If

        Dim lOrderCnt As Int32 = 0
        For X As Int32 = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    lOrderCnt += 1
                End If
            End If
        Next X

        If lOrderCnt = 0 Then Return
        If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, goSound.GetUnitSpeechTypeCnt(SoundMgr.SpeechType.eSetRallyPoint, lOrderCnt))
        Dim yMultiMove(21 + (lOrderCnt * 6)) As Byte
        Dim lPos As Int32 = 0

        'set rally point
        System.BitConverter.GetBytes(GlobalMessageCode.eSetRallyPoint).CopyTo(yMultiMove, 0) : lPos += 2

        'Destination location
        System.BitConverter.GetBytes(lDestX).CopyTo(yMultiMove, lPos) : lPos += 4
        System.BitConverter.GetBytes(lDestZ).CopyTo(yMultiMove, lPos) : lPos += 4
        System.BitConverter.GetBytes(iDestA).CopyTo(yMultiMove, lPos) : lPos += 2
        System.BitConverter.GetBytes(lDestID).CopyTo(yMultiMove, lPos) : lPos += 4
        System.BitConverter.GetBytes(iDestTypeID).CopyTo(yMultiMove, lPos) : lPos += 2
        System.BitConverter.GetBytes(lOrderCnt).CopyTo(yMultiMove, lPos) : lPos += 4

        'Now, go thru each entity and add them to the list
        For X As Int32 = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    goCurrentEnvir.oEntity(X).GetGUIDAsString.CopyTo(yMultiMove, lPos) : lPos += 6
                End If
            End If
        Next X

        'Now, send the command
        moPrimary.SendData(yMultiMove)
    End Sub

    Public Function ValidateVersion(ByVal yType As eyConnType) As Boolean
        Dim yData(5) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eClientVersion).CopyTo(yData, 0)
        System.BitConverter.GetBytes(gl_CLIENT_VERSION).CopyTo(yData, 2)

        msCurrentLoc = "ValidateVersion"

        Dim lStatus As Int32
        If yType = eyConnType.OperatorServer Then
            lStatus = elBusyStatusType.OperatorValidateVersion
            mlBusyStatus = mlBusyStatus Or lStatus
            moOperator.SendData(yData)
        Else
            lStatus = elBusyStatusType.PrimaryValidateVersion
            mlBusyStatus = mlBusyStatus Or lStatus
            moPrimary.SendData(yData)
        End If
        'While (mlBusyStatus And lStatus) <> 0
        '	'Threading.Thread.Sleep(10)
        '	Application.DoEvents()
        '	'Threading.Thread.Sleep(0)
        '	Threading.Thread.Sleep(1)
        'End While

        Return True 'mbVersionValidated
    End Function

    Public Sub SendClearEntityProdQueue(ByRef oEntity As BaseEntity)
        Dim yMsg(7) As Byte
        Dim lPos As Int32 = 0

        If goCurrentEnvir Is Nothing = False Then
            If goCurrentEnvir.ClearUnitsProdQueue(oEntity.ObjectID, oEntity.ObjTypeID) = False Then Return
        End If

        System.BitConverter.GetBytes(GlobalMessageCode.eClearEntityProdQueue).CopyTo(yMsg, lPos) : lPos += 2
        oEntity.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        moPrimary.SendData(yMsg)
    End Sub

    Public Sub SendMoveRequestMsg(ByVal lDestX As Int32, ByVal lDestZ As Int32, ByVal iDestA As Int16, ByVal bShiftDown As Boolean, ByVal bFromConfirm As Boolean)
        Dim lDestID As Int32 = -1
        Dim iDestTypeID As Int16 = -1

        msCurrentLoc = "SendMoveRequestMsg"

        If HasAliasedRights(AliasingRights.eMoveUnits) = False Then Return

        If NewTutorialManager.TutorialOn = True Then
            If goUILib Is Nothing = False AndAlso goUILib.CommandAllowed(True, "OrderMovement") = False Then Return
        End If

        'Ok, determine the currently viewed environment...
        lDestID = -1 : iDestTypeID = -1

        If goGalaxy.CurrentSystemIdx <> -1 Then
            If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx <> -1 AndAlso (glCurrentEnvirView = CurrentView.ePlanetMapView OrElse glCurrentEnvirView = CurrentView.ePlanetView) Then
                lDestID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx).ObjectID
                iDestTypeID = ObjectType.ePlanet
            Else
                lDestID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID
                iDestTypeID = ObjectType.eSolarSystem
            End If
        ElseIf goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 0 Then
            lDestID = goCurrentEnvir.ObjectID
            iDestTypeID = goCurrentEnvir.ObjTypeID
        Else
            Exit Sub
        End If

        If goUILib Is Nothing = False Then
            Dim ofrm As frmMultiDisplay = CType(goUILib.GetWindow("frmMultiDisplay"), frmMultiDisplay)
            If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
                ofrm.StoreMultiSelectLookup()
            End If
        End If

        SendMoveRequestMsgEx(lDestX, lDestZ, iDestA, bShiftDown, bFromConfirm, lDestID, iDestTypeID)

    End Sub
    Public Sub SendMoveRequestMsgEx(ByVal lDestX As Int32, ByVal lDestZ As Int32, ByVal iDestA As Int16, ByVal bShiftDown As Boolean, ByVal bFromConfirm As Boolean, ByVal lDestID As Int32, ByVal iDestTypeID As Int16)
        Dim yMultiMove() As Byte        'used for group move, attack, mining, orders etc...
        Dim lPos As Int32 = 0
        Dim X As Int32

        Dim bCancelledAlertPosted As Boolean = False

        msCurrentLoc = "SendMoveRequestMsgEx"

        If HasAliasedRights(AliasingRights.eMoveUnits) = False Then Return

        If NewTutorialManager.TutorialOn = True Then
            If goUILib Is Nothing = False AndAlso goUILib.CommandAllowed(True, "OrderMovement") = False Then Return
        End If

        If lDestID <> goCurrentEnvir.ObjectID OrElse iDestTypeID <> goCurrentEnvir.ObjTypeID Then
            If HasAliasedRights(AliasingRights.eChangeEnvironment) = False Then
                goUILib.AddNotification("You lack rights to change unit environments.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
                Return
            End If
        End If

        'now, generate our movement requests... determine if we have a multiple selection (of say > 3)
        Dim lOrderCnt As Int32 = 0
        Dim lUnitCnt As Int32 = 0
        Dim bCheckGuild As Boolean = goCurrentPlayer.oGuild Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 255
        For X = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso (goCurrentEnvir.oEntity(X).OwnerID = glPlayerID OrElse ((goCurrentEnvir.oEntity(X).CurrentStatus And elUnitStatus.eGuildAsset) <> 0) AndAlso bCheckGuild = True AndAlso goCurrentPlayer.oGuild.MemberInGuild(goCurrentEnvir.oEntity(X).OwnerID) = True) Then
                    If goCurrentEnvir.oEntity(X).bProducing = True AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit AndAlso goCurrentPlayer.yPlayerPhase <> 0 Then
                        If muSettings.bDoNotShowEngineerCancelAlert = False AndAlso bFromConfirm = False AndAlso goUILib Is Nothing = False Then
                            Dim ofrm As New frmCancelBuild(goUILib)
                            If bShiftDown = True Then
                                ofrm.SetOrder(frmCancelBuild.eyOrderType.MoveRequest, Int32.MinValue, -1, lDestID, iDestTypeID, lDestX, lDestZ, iDestA)
                            Else
                                ofrm.SetOrder(frmCancelBuild.eyOrderType.MoveRequest, -1, -1, lDestID, iDestTypeID, lDestX, lDestZ, iDestA)
                            End If

                            ofrm.Visible = True
                            Return
                        End If
                    End If
                    If bCancelledAlertPosted = False AndAlso goCurrentEnvir.oEntity(X).bProducing = True AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit Then
                        bCancelledAlertPosted = True
                        If goUILib Is Nothing = False Then
                            goUILib.AddNotification("Facility Construction Cancelled.", Color.Red, goCurrentEnvir.ObjectID, goCurrentEnvir.ObjTypeID, goCurrentEnvir.oEntity(X).ObjectID, goCurrentEnvir.oEntity(X).ObjTypeID, CInt(goCurrentEnvir.oEntity(X).LocX), CInt(goCurrentEnvir.oEntity(X).LocZ))
                            If goSound Is Nothing = False Then
                                goSound.StartSound("Game Narrative\ProductionCancelled.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
                            End If
                        End If
                    End If
                    lOrderCnt += 1
                    If goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit Then lUnitCnt += 1
                End If
            End If
        Next X

        If lOrderCnt = 0 Then Return

        ReDim yMultiMove(17 + (lOrderCnt * 6))
        lPos = 0
        If bShiftDown = True Then
            System.BitConverter.GetBytes(GlobalMessageCode.eAddWaypointMsg).CopyTo(yMultiMove, 0) : lPos += 2
        Else
            System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequest).CopyTo(yMultiMove, 0) : lPos += 2

            If goSound Is Nothing = False AndAlso lUnitCnt <> 0 Then goSound.StartUnitSpeechSound(True, goSound.GetUnitSpeechTypeCnt(SoundMgr.SpeechType.eMoveRequest, lUnitCnt))
        End If

        System.BitConverter.GetBytes(lDestX).CopyTo(yMultiMove, lPos) : lPos += 4
        System.BitConverter.GetBytes(lDestZ).CopyTo(yMultiMove, lPos) : lPos += 4
        'TODO: eventually, we will want to determine the final facing if the player wants to specify one
        System.BitConverter.GetBytes(CShort(-1)).CopyTo(yMultiMove, lPos) : lPos += 2
        System.BitConverter.GetBytes(lDestID).CopyTo(yMultiMove, lPos) : lPos += 4
        System.BitConverter.GetBytes(iDestTypeID).CopyTo(yMultiMove, lPos) : lPos += 2

        'Now, go thru each entity and add them to the list
        For X = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso (goCurrentEnvir.oEntity(X).OwnerID = glPlayerID OrElse ((goCurrentEnvir.oEntity(X).CurrentStatus And elUnitStatus.eGuildAsset) <> 0) AndAlso bCheckGuild = True AndAlso goCurrentPlayer.oGuild.MemberInGuild(goCurrentEnvir.oEntity(X).OwnerID) = True) Then
                    goCurrentEnvir.oEntity(X).GetGUIDAsString.CopyTo(yMultiMove, lPos) : lPos += 6
                    SendClearEntityProdQueue(goCurrentEnvir.oEntity(X))
                End If
            End If
        Next X

        'Now, send it...
        SendToRegion(yMultiMove)

        If goCurrentEnvir.ObjectID <> lDestID OrElse goCurrentEnvir.ObjTypeID <> iDestTypeID Then
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eOrders.eChangeEnvironmentMovement, lOrderCnt, 0)
        Else
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eOrders.eGroupMovement, lOrderCnt, 0)
        End If

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eMoveCommandGiven, -1, -1, -1, "")
        End If
    End Sub

    Private mlPrimaryFailures As Int32 = 0
    Private mlRegionFailures As Int32 = 0
    Public Function SendKeepAliveMsgs() As Boolean
        'we just send a message that neither server cares about from the client
        Dim yMsg(1) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.ePathfindingConnectionInfo).CopyTo(yMsg, 0)
        msCurrentLoc = "SendKeepAliveMsgs"
        Try
            If moPrimary Is Nothing = False AndAlso moPrimary.IsConnected = True Then moPrimary.SendData(yMsg)
            mlPrimaryFailures = 0
        Catch
            mlPrimaryFailures += 1
        End Try
        Try
            If moRegion Is Nothing = False AndAlso moRegion.IsConnected = True Then moRegion.SendData(yMsg)
            mlRegionFailures = 0
        Catch
            mlRegionFailures += 1
        End Try

        If mlPrimaryFailures > 3 OrElse mlRegionFailures > 3 Then
            If goUILib Is Nothing = False Then
                goUILib.AddNotification("Connection lost to the server!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                'RaiseEvent ServerShutdown()
                'MsgBox("Connection lost to the server.", MsgBoxStyle.OkOnly, "Connection Lost")
                Return False
            End If
        End If

        Return True
    End Function

    Public Sub SendJumpTargetMsg(ByVal lJumpTargetID As Int32, ByVal iJumpTargetTypeID As Int16)
        Dim lDestID As Int32
        Dim iDestTypeID As Int16
        Dim yMultiMove() As Byte        'used for group move, attack, mining, orders etc...
        Dim lPos As Int32 = 0
        Dim X As Int32

        msCurrentLoc = "SendJumpTargetMsg"

        If HasAliasedRights(AliasingRights.eChangeEnvironment) = False Then Return

        'Ok, determine the currently viewed environment...
        lDestID = -1 : iDestTypeID = -1
        If goGalaxy.CurrentSystemIdx <> -1 Then
            If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx <> -1 AndAlso (glCurrentEnvirView = CurrentView.ePlanetMapView OrElse glCurrentEnvirView = CurrentView.ePlanetView) Then
                lDestID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx).ObjectID
                iDestTypeID = ObjectType.ePlanet
            Else
                lDestID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID
                iDestTypeID = ObjectType.eSolarSystem
            End If
        Else
            Exit Sub
        End If
        If iDestTypeID <> ObjectType.eSolarSystem Then Return

        'now, generate our movement requests... determine if we have a multiple selection (of say > 3)
        Dim lOrderCnt As Int32 = 0
        For X = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    lOrderCnt += 1
                End If
            End If
        Next X

        If lOrderCnt = 0 Then Return

        If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(True, goSound.GetUnitSpeechTypeCnt(SoundMgr.SpeechType.eMoveRequest, lOrderCnt))

        ReDim yMultiMove(13 + (lOrderCnt * 6))
        lPos = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eJumpTarget).CopyTo(yMultiMove, 0) : lPos += 2
        System.BitConverter.GetBytes(lDestID).CopyTo(yMultiMove, lPos) : lPos += 4
        System.BitConverter.GetBytes(iDestTypeID).CopyTo(yMultiMove, lPos) : lPos += 2
        System.BitConverter.GetBytes(lJumpTargetID).CopyTo(yMultiMove, lPos) : lPos += 4
        System.BitConverter.GetBytes(iJumpTargetTypeID).CopyTo(yMultiMove, lPos) : lPos += 2

        'Now, go thru each entity and add them to the list
        For X = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    goCurrentEnvir.oEntity(X).GetGUIDAsString.CopyTo(yMultiMove, lPos) : lPos += 6
                    SendClearEntityProdQueue(goCurrentEnvir.oEntity(X))
                End If
            End If
        Next X

        'Now, send it...
        SendToRegion(yMultiMove)
    End Sub

    Public Sub SendFormationMoveRequestMsg(ByVal lDestX As Int32, ByVal lDestZ As Int32, ByVal iDestA As Int16, ByVal bShiftDown As Boolean)
        Dim lDestID As Int32
        Dim iDestTypeID As Int16
        Dim yMultiMove() As Byte        'used for group move, attack, mining, orders etc...
        Dim lPos As Int32 = 0
        Dim X As Int32

        Dim bCancelledAlertPosted As Boolean = False

        msCurrentLoc = "SendMoveRequestMsg"

        If HasAliasedRights(AliasingRights.eMoveUnits) = False Then Return

        Dim oWin As frmMultiDisplay = Nothing
        If goUILib Is Nothing = False Then
            oWin = CType(goUILib.GetWindow("frmMultiDisplay"), frmMultiDisplay)
        End If
        If oWin Is Nothing OrElse oWin.Visible = False Then
            SendMoveRequestMsg(lDestX, lDestZ, iDestA, bShiftDown, False)
            Return
        End If
        Dim lFormationID As Int32 = oWin.GetFormationID
        If lFormationID < 1 Then
            SendMoveRequestMsg(lDestX, lDestZ, iDestA, bShiftDown, False)
            Return
        End If
        oWin.StoreMultiSelectLookup()

        'Ok, determine the currently viewed environment...
        lDestID = -1 : iDestTypeID = -1

        If goGalaxy.CurrentSystemIdx <> -1 Then
            If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx <> -1 AndAlso (glCurrentEnvirView = CurrentView.ePlanetMapView OrElse glCurrentEnvirView = CurrentView.ePlanetView) Then
                lDestID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx).ObjectID
                iDestTypeID = ObjectType.ePlanet
            Else
                lDestID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID
                iDestTypeID = ObjectType.eSolarSystem
            End If
        Else
            Exit Sub
        End If
        If lDestID <> goCurrentEnvir.ObjectID OrElse iDestTypeID <> goCurrentEnvir.ObjTypeID Then
            SendMoveRequestMsg(lDestX, lDestZ, iDestA, bShiftDown, False)
            Return
        End If

        'now, generate our movement requests... determine if we have a multiple selection (of say > 3)
        Dim lOrderCnt As Int32 = 0
        Dim lUnitCnt As Int32 = 0
        For X = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    If bCancelledAlertPosted = False AndAlso goCurrentEnvir.oEntity(X).bProducing = True AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit Then
                        bCancelledAlertPosted = True
                        If goUILib Is Nothing = False Then
                            goUILib.AddNotification("Facility Construction Cancelled.", Color.Red, goCurrentEnvir.ObjectID, goCurrentEnvir.ObjTypeID, goCurrentEnvir.oEntity(X).ObjectID, goCurrentEnvir.oEntity(X).ObjTypeID, CInt(goCurrentEnvir.oEntity(X).LocX), CInt(goCurrentEnvir.oEntity(X).LocZ))
                            If goSound Is Nothing = False Then
                                goSound.StartSound("Game Narrative\ProductionCancelled.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
                            End If
                        End If
                    End If
                    lOrderCnt += 1
                    If goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit Then lUnitCnt += 1
                End If
            End If
        Next X

        If lOrderCnt = 0 Then Return

        ReDim yMultiMove(26 + (lOrderCnt * 6))
        lPos = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eMoveFormation).CopyTo(yMultiMove, 0) : lPos += 2
        System.BitConverter.GetBytes(lFormationID).CopyTo(yMultiMove, lPos) : lPos += 4

        If muSettings.bFormationMoveThenForm = True Then yMultiMove(lPos) = 255 Else yMultiMove(lPos) = 0
        lPos += 1

        System.BitConverter.GetBytes(lDestX).CopyTo(yMultiMove, lPos) : lPos += 4
        System.BitConverter.GetBytes(lDestZ).CopyTo(yMultiMove, lPos) : lPos += 4
        System.BitConverter.GetBytes(iDestA).CopyTo(yMultiMove, lPos) : lPos += 2
        System.BitConverter.GetBytes(lDestID).CopyTo(yMultiMove, lPos) : lPos += 4
        System.BitConverter.GetBytes(iDestTypeID).CopyTo(yMultiMove, lPos) : lPos += 2
        System.BitConverter.GetBytes(lOrderCnt).CopyTo(yMultiMove, lPos) : lPos += 4


        'Now, go thru each entity and add them to the list
        For X = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                If goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                    goCurrentEnvir.oEntity(X).GetGUIDAsString.CopyTo(yMultiMove, lPos) : lPos += 6
                    SendClearEntityProdQueue(goCurrentEnvir.oEntity(X))
                End If
            End If
        Next X

        'Now, send it...
        If goSound Is Nothing = False AndAlso lUnitCnt <> 0 Then goSound.StartUnitSpeechSound(True, goSound.GetUnitSpeechTypeCnt(SoundMgr.SpeechType.eMoveRequest, lUnitCnt))
        SendToRegion(yMultiMove)

        BPMetrics.MetricMgr.AddActivity(BPMetrics.eOrders.eMoveFormation, lOrderCnt, lFormationID)
    End Sub

#End Region

#Region " Incoming Message Handlers "
    Private Sub HandleWPTimers(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If glPlayerID = lPlayerID AndAlso goUILib Is Nothing = False Then
            Dim lLastWPUpkeepTime As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lLastGuildShareUpkeep As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lSecondsSinceLastUpkeep As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            Dim lSecondsToUpkeep As Int32 = 86400 - lSecondsSinceLastUpkeep

            Dim ofrm As frmEnvirDisplay = CType(goUILib.GetWindow("frmEnvirDisplay"), frmEnvirDisplay)
            If ofrm Is Nothing = False Then
                ofrm.sNextWPUpkeepTime = DateTime.SpecifyKind(GetDateFromNumber(lLastWPUpkeepTime), DateTimeKind.Utc).ToLocalTime.AddDays(30).ToString
                ofrm.sNextGuildShareUpkeep = DateTime.SpecifyKind(GetDateFromNumber(lLastGuildShareUpkeep), DateTimeKind.Utc).ToLocalTime.AddDays(1).ToString
                ofrm.dtSessionNextUpkeep = DateTime.SpecifyKind(Now, DateTimeKind.Local).AddSeconds(lSecondsToUpkeep)
            End If
        End If
    End Sub
    Private Sub HandleRespawnSelection(ByVal yData() As Byte)
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, 2)
        If lID = -1 Then
            'Ok, primary is telling us we need to select a respawn system, so open our respawn select window
            If goUILib Is Nothing = False Then
                Dim oFrm As New frmSelectRespawn(goUILib)
                oFrm.Visible = True
            End If
        ElseIf lID = Int32.MinValue Then
            'server giving us the spawn in spawn system option
            If goUILib Is Nothing = False Then
                Dim oFrm As New frmSpawnChoice(goUILib)
                oFrm.Visible = True
            End If
        Else
            'primary is adding a system to the list
            If goUILib Is Nothing = False Then
                Dim oFrm As frmSelectRespawn = CType(goUILib.GetWindow("frmSelectRespawn"), frmSelectRespawn)
                If oFrm Is Nothing = False Then oFrm.HandleAddSystemItem(yData)
            End If
        End If
    End Sub
    'Private Sub HandleRewardOrReportWarpoints(ByVal yData() As Byte)
    '    msCurrentLoc = "HandleRewardOrReportWarpoints"

    '    Dim lPos As Int32 = 0
    '    Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

    '    Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '    Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
    '    Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '    Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
    '    Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '    Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '    Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '    Dim lInfluence As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

    '    If lPlayerID = glPlayerID Then
    '        If goRewards Is Nothing = False Then
    '            Dim lY As Int32 = 0
    '            If goCurrentEnvir Is Nothing = False Then
    '                'if different envir, we dont care
    '                If goCurrentEnvir.ObjectID <> lEnvirID OrElse goCurrentEnvir.ObjTypeID <> iEnvirTypeID Then Return

    '                'Snap it to the unit/facility
    '                Try
    '                    Dim lUB As Int32 = -1
    '                    If goCurrentEnvir.lEntityIdx Is Nothing = False Then lUB = Math.Min(goCurrentEnvir.lEntityIdx.GetUpperBound(0), goCurrentEnvir.lEntityUB)
    '                    For X As Int32 = 0 To lUB
    '                        If goCurrentEnvir.lEntityIdx(X) = lObjID Then
    '                            Dim oEnt As BaseEntity = goCurrentEnvir.oEntity(X)
    '                            If oEnt Is Nothing = False AndAlso oEnt.ObjTypeID = iObjTypeID Then
    '                                lLocX = CInt(oEnt.LocX)
    '                                lLocZ = CInt(oEnt.LocZ)
    '                                Exit For
    '                            End If
    '                        End If
    '                    Next X
    '                Catch
    '                End Try

    '                If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
    '                    Dim oObj As Object = goCurrentEnvir.oGeoObject
    '                    If oObj Is Nothing = False AndAlso CType(oObj, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
    '                        lY = CInt(CType(oObj, Planet).GetHeightAtPoint(lLocX, lLocZ, True))
    '                    End If
    '                End If
    '            End If
    '            lY += 300

    '            goRewards.AddReward(lLocX, lY, lLocZ, lInfluence, iMsgCode = GlobalMessageCode.eRewardWarpoints)
    '        End If
    '    End If

    'End Sub
    Private Sub HandleGetDAValues(ByVal yData() As Byte)
        'MSC - 04/17/10 - remarked out for the time being
        'Just need the type id
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 2)

        If goUILib Is Nothing Then Return

        Select Case iTypeID
            Case ObjectType.eEngineTech
                Dim oFrm As frmEngineBuilder = CType(goUILib.GetWindow("frmEngineBuilder"), frmEngineBuilder)
                If oFrm Is Nothing = False Then
                    oFrm.DADataReceived(yData)
                End If
            Case ObjectType.eRadarTech
                Dim oFrm As frmRadarBuilder = CType(goUILib.GetWindow("frmRadarBuilder"), frmRadarBuilder)
                If oFrm Is Nothing = False Then
                    oFrm.DADataReceived(yData)
                End If
            Case ObjectType.eShieldTech
                Dim oFrm As frmShieldBuilder = CType(goUILib.GetWindow("frmShieldBuilder"), frmShieldBuilder)
                If oFrm Is Nothing = False Then
                    oFrm.DADataReceived(yData)
                End If
            Case ObjectType.eWeaponTech
                Dim oFrm As frmWeaponBuilder = CType(goUILib.GetWindow("frmWeaponBuilder"), frmWeaponBuilder)
                If oFrm Is Nothing = False Then
                    oFrm.DADataReceived(yData)
                End If
        End Select
    End Sub
    Private Sub HandleRequestObjectResponse(ByVal yData() As Byte)

        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        If iTypeID = ObjectType.eFacility Then
            If goUILib Is Nothing = False Then

                Dim yType As Byte = yData(lPos) : lPos += 1

                If yType = 0 Then
                    Dim oWin As frmProdCost = CType(goUILib.GetWindow("frmResCost"), frmProdCost)
                    If oWin Is Nothing = False Then
                        oWin.HandleRequestTimeToDoHere(yData)
                    Else
                        Dim oTmp As frmSpecTech = CType(goUILib.GetWindow("frmSpecTech"), frmSpecTech)
                        If oTmp Is Nothing = False Then
                            If oTmp.mfrmResCost Is Nothing = False Then
                                oTmp.mfrmResCost.HandleRequestTimeToDoHere(yData)
                            End If
                        End If
                    End If
                Else
                    Dim oWin As frmProdCost = CType(goUILib.GetWindow("frmProdCost"), frmProdCost)
                    If oWin Is Nothing = False Then
                        oWin.HandleRequestTimeToDoHere(yData)
                    End If
                End If
            End If
        ElseIf iTypeID = ObjectType.eSolarSystem Then
            If goGalaxy Is Nothing = False Then
                For X As Int32 = 0 To goGalaxy.mlSystemUB
                    If goGalaxy.moSystems(X) Is Nothing = False Then
                        If goGalaxy.moSystems(X).ObjectID = lObjID Then
                            goGalaxy.moSystems(X).HandleRequestGalMapViewDetails(yData)
                            Exit For
                        End If
                    End If
                Next
            End If
        End If

    End Sub
    Private Sub HandleSenateStatusReport(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        'lPos += 4       'for playerid
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If lPlayerID <> glPlayerID Then Return

        Dim lFloorOpen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lFloorVoted As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lEmpChmbrOpen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lEmpChmbrVoted As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If NewTutorialManager.TutorialOn = True Then Return
        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerTitle >= Player.PlayerRank.Duke AndAlso (lFloorOpen > lFloorVoted OrElse lEmpChmbrOpen > lEmpChmbrVoted) Then
            If lFloorOpen > lFloorVoted Then
                Dim sMsg As String = "You have " & (lFloorOpen - lFloorVoted).ToString & " Senate Floor Proposals awaiting your attention."
                If goUILib Is Nothing = False Then goUILib.AddNotification(sMsg, System.Drawing.Color.FromArgb(255, 255, 255, 0), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
            If goCurrentPlayer.yPlayerTitle = Player.PlayerRank.Emperor AndAlso lEmpChmbrOpen > lEmpChmbrVoted Then
                Dim sMsg As String = "You have " & (lEmpChmbrOpen - lEmpChmbrVoted).ToString & " Emperor's Chamber Proposals awaiting your attention."
                If goUILib Is Nothing = False Then goUILib.AddNotification(sMsg, System.Drawing.Color.FromArgb(255, 255, 255, 0), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If

            If goUILib Is Nothing = False Then
                Dim ofrm As frmQuickBar = CType(goUILib.GetWindow("frmQuickBar"), frmQuickBar)
                If ofrm Is Nothing Then ofrm = New frmQuickBar(goUILib)
                If ofrm Is Nothing = False Then ofrm.SenateUpdate()
                ofrm = Nothing
            End If

        End If
    End Sub

    Private Sub HandleDeleteFleet(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lFleetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
            If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
                goCurrentPlayer.moUnitGroups(X) = Nothing
                goCurrentPlayer.mlUnitGroupIdx(X) = -1
            End If
        Next X
    End Sub
    Private Sub HandleClaimables(ByVal yData() As Byte)
        Dim lPos As Int32 = 2

        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        ReDim guClaimables(lCnt - 1)
        For X As Int32 = 0 To lCnt - 1
            With guClaimables(X)
                .lID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .iTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .lOfferCode = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .sName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                .yClaimFlag = yData(lPos) : lPos += 1
            End With
        Next X

    End Sub

    Private Sub HandleUpdatePlayerDetails(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If lPlayerID <> glPlayerID Then Return
        gsUserName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
        gsPassword = GetStringFromBytes(yData, lPos, 20) : lPos += 20
        goCurrentPlayer.lAccountStatus = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        gsUserNameProper = GetStringFromBytes(yData, lPos, 20) : lPos += 20
        If goCurrentPlayer Is Nothing = False Then goCurrentPlayer.PlayerName = gsUserNameProper
    End Sub
    Private Sub HandleGetGuildBillboards(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2


        If iObjTypeID = ObjectType.ePlanet Then
            Dim lSystemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim oPlanet As Planet = Nothing

            If goGalaxy Is Nothing = False Then
                For X As Int32 = 0 To goGalaxy.mlSystemUB
                    If goGalaxy.moSystems(X) Is Nothing = False AndAlso goGalaxy.moSystems(X).ObjectID = lSystemID Then
                        For Y As Int32 = 0 To goGalaxy.moSystems(X).PlanetUB
                            If goGalaxy.moSystems(X).moPlanets(Y) Is Nothing = False AndAlso goGalaxy.moSystems(X).moPlanets(Y).ObjectID = lObjID Then
                                oPlanet = goGalaxy.moSystems(X).moPlanets(Y)
                                Exit For
                            End If
                        Next Y
                        If oPlanet Is Nothing = False Then Exit For
                    End If
                Next X
            End If
            If oPlanet Is Nothing Then Return

            For X As Int32 = 0 To oPlanet.uBillboards.GetUpperBound(0)
                With oPlanet.uBillboards(X)
                    .lGuildID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .iRecruitFlags = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .lIcon = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .BidAmount = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .BillboardText = GetStringFromBytes(yData, lPos, 200) : lPos += 200
                End With
            Next X

        End If

    End Sub

    Private Sub HandleRequestMaxWpnRng(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lMaxRng As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If lMaxRng < 1 Then Return
        Try
            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing = False Then
                For X As Int32 = 0 To oEnvir.lEntityUB
                    If oEnvir.lEntityIdx(X) = lID Then
                        If oEnvir.oEntity(X) Is Nothing = False AndAlso oEnvir.oEntity(X).ObjTypeID = iTypeID Then
                            lMaxRng += oEnvir.oEntity(X).oMesh.RangeOffset
                            oEnvir.oEntity(X).lMaxWpnRngValue = lMaxRng * gl_FINAL_GRID_SQUARE_SIZE
                            Exit For
                        End If
                    End If
                Next X
            End If
        Catch
        End Try

    End Sub

    Private Sub HandleUpdatePlanetOwnership(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlanetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lSystemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If goGalaxy Is Nothing = False Then
            For X As Int32 = 0 To goGalaxy.mlSystemUB
                If goGalaxy.moSystems(X) Is Nothing = False AndAlso goGalaxy.moSystems(X).ObjectID = lSystemID Then

                    For Y As Int32 = 0 To goGalaxy.moSystems(X).PlanetUB
                        If goGalaxy.moSystems(X).moPlanets(Y) Is Nothing = False AndAlso goGalaxy.moSystems(X).moPlanets(Y).ObjectID = lPlanetID Then
                            goGalaxy.moSystems(X).moPlanets(Y).HandleUpdatePlanetOwnership(yData)
                            Exit For
                        End If
                    Next Y
                    Exit For
                End If
            Next X
        End If
    End Sub
    Private Sub HandleOutbidAlert(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lFacID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iFacTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        If goUILib Is Nothing = False Then
            goUILib.AddNotification("You have been outbid at a mining facility.", Color.Yellow, lEnvirID, iEnvirTypeID, lFacID, iFacTypeID, Int32.MinValue, Int32.MinValue)
        End If

    End Sub

    Private Sub HandleUpdatePlayerCustomTitle(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If lPlayerID = glPlayerID Then
            Dim lCustomPerms As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim yCustomTitle As Byte = yData(lPos) : lPos += 1

            If goCurrentPlayer Is Nothing = False Then
                goCurrentPlayer.yCustomTitle = yCustomTitle
                goCurrentPlayer.lCustomTitlePermission = lCustomPerms
            End If
        Else
            Dim yCustomTitle As Byte = yData(lPos) : lPos += 1
            For X As Int32 = 0 To glPlayerIntelUB
                If glPlayerIntelIdx(X) = lPlayerID Then
                    Dim oPII As PlayerIntel = goPlayerIntel(X)
                    If oPII Is Nothing = False Then
                        oPII.yCustomTitle = yCustomTitle
                    End If
                    Exit For
                End If
            Next X
        End If

    End Sub

    Private Sub HandleSetFleetReinforcer(ByRef yData() As Byte)
        msCurrentLoc = "HandleSetFleetReinforcer"
        Dim lPos As Int32 = 2
        Dim lFleetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lFacilityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oUnitGroup As UnitGroup = Nothing
        If goCurrentPlayer Is Nothing = False Then
            For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
                If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
                    oUnitGroup = goCurrentPlayer.moUnitGroups(X)
                    Exit For
                End If
            Next X
        End If
        If oUnitGroup Is Nothing = True Then
            If goUILib Is Nothing = False Then goUILib.AddNotification("Unable to add reinforcing facility to battlegroup!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If
        If lFacilityID = -1 Then
            If goUILib Is Nothing = False Then goUILib.AddNotification("Unable to add reinforcing facility to " & oUnitGroup.sName & " battlegroup!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        oUnitGroup.AddReinforcer(lFacilityID, lParentID, iParentTypeID)

        If goUILib Is Nothing = False Then goUILib.AddNotification(GetCacheObjectValue(lFacilityID, ObjectType.eFacility) & " is now reinforcing the " & oUnitGroup.sName & " battlegroup.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    End Sub

    Private Sub HandleAddPlayerTechKnow(ByRef yData() As Byte)
        'Ok, always first...
        Dim lPos As Int32 = 2       'for msg cd
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yKnowType As Byte = yData(lPos) : lPos += 1
        Dim bArchive As Boolean = yData(lPos) <> 0 : lPos += 1
        msCurrentLoc = "HandleAddPlayerTechKnow"
        If lPlayerID = glPlayerID Then
            Dim oPTK As New PlayerTechKnowledge()
            oPTK.oPlayer = goCurrentPlayer
            oPTK.yKnowledgeType = CType(yKnowType, PlayerTechKnowledge.KnowledgeType)
            oPTK.bArchived = bArchive
            If oPTK.FillFromMsg(yData, lPos) = True Then
                'Add to player's knowledge list
                Dim bFound As Boolean = False
                For X As Int32 = 0 To glPlayerTechKnowledgeUB
                    If glPlayerTechKnowledgeIdx(X) = oPTK.oTech.ObjectID Then
                        If goPlayerTechKnowledge(X).oTech.ObjTypeID = oPTK.oTech.ObjTypeID AndAlso goPlayerTechKnowledge(X).oPlayer.ObjectID = oPTK.oPlayer.ObjectID Then
                            goPlayerTechKnowledge(X) = oPTK
                            bFound = True
                            Exit For
                        End If
                    End If
                Next X
                If bFound = False Then
                    ReDim Preserve glPlayerTechKnowledgeIdx(glPlayerTechKnowledgeUB + 1)
                    ReDim Preserve goPlayerTechKnowledge(glPlayerTechKnowledgeUB + 1)
                    goPlayerTechKnowledge(glPlayerTechKnowledgeUB + 1) = oPTK
                    glPlayerTechKnowledgeIdx(glPlayerTechKnowledgeUB + 1) = oPTK.oTech.ObjectID
                    glPlayerTechKnowledgeUB += 1
                End If
            End If
            'Else
            '	'TODO: Espionage on some else's espionage?
        End If

    End Sub

    Private Sub HandleUpdatePlayerTechValue(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        msCurrentLoc = "HandleUpdatePlayerTechValue"
        If lPlayerID = glPlayerID Then
            goCurrentPlayer.HandleUpdatePlayerTechValue(yData)
            'Else
            '	'TODO: Espionage?
        End If
    End Sub

    Private Sub HandlePlayerTitleChange(ByRef yData() As Byte)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim yVal As Byte = yData(6)
        msCurrentLoc = "HandlePlayerTitleChange"
        If lPlayerID = glPlayerID Then
            If goCurrentPlayer Is Nothing = False Then
                Dim sMsg As String

                'TODO: Record new WAV files for this event
                Dim sWAV As String = "Game Narrative\DiplomaticRelChange.wav"

                If goCurrentPlayer.yPlayerTitle > yVal Then
                    'Demotion
                    sMsg = "You have been demoted by title!"
                Else : sMsg = "You have been promoted by title!"
                End If
                If goUILib Is Nothing = False Then goUILib.AddNotification(sMsg, Color.White, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                If goSound Is Nothing = False Then goSound.StartSound(sWAV, False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)

                goCurrentPlayer.yPlayerTitle = yVal
            End If
        Else
            For X As Int32 = 0 To glPlayerIntelUB
                If glPlayerIntelIdx(X) = lPlayerID Then
                    Dim sMsg As String = goPlayerIntel(X).PlayerName & " has a new title: " & Player.GetPlayerNameWithTitle(yVal, goPlayerIntel(X).PlayerName, goPlayerIntel(X).bIsMale)
                    If goUILib Is Nothing = False Then goUILib.AddNotification(sMsg, Color.White, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    'TODO: Record new WAV files for this event
                    If goSound Is Nothing = False Then goSound.StartSound("Game Narrative\DiplomaticRelChange.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
                    Exit For
                End If
            Next X
        End If
    End Sub

    Private Sub HandleGetPlayerScores(ByRef yData() As Byte)
        'should always be my scores
        msCurrentLoc = "HandleGetPlayerScores"
        If goCurrentPlayer Is Nothing = False Then
            With goCurrentPlayer
                Dim lPos As Int32 = 2
                .lTechnologyScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lDiplomacyScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lMilitaryScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lPopulationScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lProductionScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lWealthScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lTotalScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                Dim iCnt As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                Dim lCP(iCnt - 1) As Int32
                Dim yCP(iCnt - 1) As Byte

                For X As Int32 = 0 To iCnt - 1
                    lCP(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    yCP(X) = yData(lPos) : lPos += 1
                Next X
                .lControlPlanets = lCP
                .yControlPlanetAmt = yCP
            End With
        End If
    End Sub

    Private Sub HandlePlayerAlert(ByRef yData() As Byte)
        msCurrentLoc = "HandlePlayerAlert"
        'msgcode, type, entityguid, envirguid, playerid, enemyid, optional name(20)
        Dim lPos As Int32 = 2
        Dim yType As Byte = yData(lPos) : lPos += 1
        Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lEnemyID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim sMsg As String
        Dim sWAV As String = ""
        Dim lClr As System.Drawing.Color = Color.Red

        Dim lLocX As Int32 = Int32.MinValue
        Dim lLocZ As Int32 = Int32.MinValue

        'Now... what to do what to do
        Select Case yType
            Case PlayerAlertType.eColonyLost
                sMsg = "We have lost contact with a colony: " & GetStringFromBytes(yData, lPos, 20) & "!"
                lPos += 20
                lLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                lLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                sWAV = "Game Narrative\ColonyLost.wav"
            Case PlayerAlertType.eEngagedEnemy
                sWAV = "Game Narrative\EngagedEnemy.wav"
                'lPos += 20
                lLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                lLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                If lPos + 20 > yData.Length Then
                    sMsg = "We have engaged the enemy!"
                Else
                    If lPos + 40 < yData.Length Then lPos += 20
                    Dim sAltName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                    If sAltName <> "" Then
                        sMsg = "We have engaged the enemy at " & sAltName & "!"
                    Else
                        sMsg = "We have engaged the enemy!"
                    End If
                End If
                If NewTutorialManager.TutorialOn = True Then
                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eEngagedEnemy, -1, -1, -1, "")
                End If
            Case PlayerAlertType.eFacilityLost
                sMsg = "Facility Lost (" & GetStringFromBytes(yData, lPos, 20) & ")"
                lPos += 20
                lLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                lLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                sWAV = "Game Narrative\FacilityLost.wav"
                Dim sAltName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                If sAltName <> "" Then sMsg &= " at " & sAltName
            Case PlayerAlertType.eUnderAttack
                sWAV = "Game Narrative\UnderAttack.wav"
                'lPos += 20
                lLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                lLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                If lPos + 20 > yData.Length Then
                    sMsg = "Our forces are under attack!"
                Else
                    If lPos + 40 < yData.Length Then lPos += 20
                    Dim sAltName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                    If sAltName <> "" Then
                        sMsg = "Our forces are under attack at " & sAltName & "!"
                    Else
                        sMsg = "Our forces are under attack!"
                    End If
                End If

            Case PlayerAlertType.eUnitLost
                sMsg = "Unit Lost (" & GetStringFromBytes(yData, lPos, 20) & ")"
                lPos += 20
                lLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                lLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                sWAV = "Game Narrative\UnitLost.wav"
                Dim sAltName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                If sAltName <> "" Then sMsg &= " at " & sAltName
            Case PlayerAlertType.eBuyOrderAccepted
                sMsg = "Buy Order Accepted!"
                sWAV = "UserInterface\Alert.wav"
                lClr = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Case Else
                sMsg = "Undefined Message!"
        End Select

        If goUILib Is Nothing = False Then
            goUILib.AddNotification(sMsg, lClr, lEnvirID, iEnvirTypeID, lEntityID, iEntityTypeID, lLocX, lLocZ)
        End If
        If goSound Is Nothing = False Then goSound.StartSound(sWAV, False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)

    End Sub

    Private Sub HandleAddPlayerCommFolder(ByRef yData() As Byte)
        msCurrentLoc = "HandleAddPlayerCommFolder"
        Dim lPos As Int32 = 2   '2 bytem sg code
        Dim lPCF_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        mlLastAddObjectID = lPCF_ID
        miLastAddObjectTypeID = -1

        If lPlayerID = glPlayerID Then
            goCurrentPlayer.AddEmailFolder(lPCF_ID, GetStringFromBytes(yData, lPos, 20))
            'Else
            '	'TODO: intercepting a folder???
        End If
    End Sub

    Private Sub HandleUpdateEntityParent(ByRef yData() As Byte)
        msCurrentLoc = "HandleUpdateEntityParent"

        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
        Try
            'For now, this only pertains to units in UnitGroups... but it can be expanded on more later
            If iObjTypeID = ObjectType.eUnit Then
                For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
                    If goCurrentPlayer.mlUnitGroupIdx(X) <> -1 Then
                        With goCurrentPlayer.moUnitGroups(X)
                            For Y As Int32 = 0 To .lUnitUB
                                If .uUnitIDs(Y).lUnitID = lObjID Then
                                    .uUnitIDs(Y).lParentID = lParentID
                                    .uUnitIDs(Y).iParentTypeID = iParentTypeID
                                    .LastMessageUpdate = glCurrentCycle
                                    Return
                                End If
                            Next Y
                        End With
                    End If
                Next X
            End If
        Catch
            'do nothing
        End Try
    End Sub

    Private Sub HandleFleetInterSystemMoving(ByRef yData() As Byte)
        If goCurrentEnvir Is Nothing Then Return

        msCurrentLoc = "HandleFleetInterSystemMoving"
        Dim lPos As Int32 = 2
        Dim lFleetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'TODO: Some sort of special fx could be cool here...

        For X As Int32 = 0 To lCnt - 1
            Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            For lIdx As Int32 = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(lIdx) = lID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit Then
                    goCurrentEnvir.oEntity(lIdx).ClearParticleFX()
                    goCurrentEnvir.lEntityIdx(lIdx) = -1
                    goCurrentEnvir.oEntity(lIdx) = Nothing
                End If
            Next lIdx
        Next X

    End Sub

    Private Sub HandleRemoveObject(ByVal yData() As Byte)
        'ok, remove it
        msCurrentLoc = "HandleRemoveObject"
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim yRemoveType As RemovalType = CType(yData(8), RemovalType)
        Dim X As Int32

        Dim lIdx As Int32 = -1

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return

        If iObjTypeID = ObjectType.eMineralCache OrElse iObjTypeID = ObjectType.eComponentCache Then
            For X = 0 To oEnvir.lCacheUB
                If oEnvir.lCacheIdx(X) = lObjID AndAlso oEnvir.oCache(X).ObjTypeID = iObjTypeID Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx <> -1 Then
                oEnvir.lCacheIdx(X) = -1
                oEnvir.oCache(X) = Nothing
            End If
        Else
            Dim lCurUB As Int32 = -1
            If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
            For X = 0 To lCurUB
                If oEnvir.lEntityIdx(X) = lObjID AndAlso oEnvir.oEntity(X).ObjTypeID = iObjTypeID Then
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx <> -1 Then
                Dim oEntity As BaseEntity = oEnvir.oEntity(lIdx)
                If oEntity Is Nothing = False Then
                    If oEntity.ObjTypeID = ObjectType.eFacility AndAlso oEntity.OwnerID = glPlayerID Then
                        If oEntity.yVisibility = eVisibilityType.Visible OrElse oEntity.yVisibility = eVisibilityType.FacilityIntel Then
                            If oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oEnvir.oGeoObject Is Nothing = False AndAlso CType(oEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
                                CType(oEnvir.oGeoObject, Planet).SetCityCreepLoc(oEntity.LocX, oEntity.LocZ, False, oEntity.oMesh.ShieldXZRadius)
                            End If
                        End If
                    End If
                End If

                Select Case yRemoveType
                    Case RemovalType.eObjectDestroyed
                        With oEntity
                            .bObjectDestroyed = True

                            If NewTutorialManager.TutorialOn = True Then
                                'ok, what was destroyed?
                                If .OwnerID = gl_HARDCODE_PIRATE_PLAYER_ID Then
                                    'pirate
                                    If .ObjTypeID = ObjectType.eFacility Then
                                        Dim lMesh As Int32 = (.oMesh.lModelID And 255)
                                        If lMesh = 16 Then
                                            'turret
                                            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eIncrementPirateTurretDeath, -1, -1, -1, "")
                                        ElseIf lMesh = 12 Then
                                            'power gen
                                            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.ePirateMediumPowerKilled, -1, -1, -1, "")
                                        ElseIf lMesh = 15 Then
                                            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.ePirateBaseKilled, -1, -1, -1, "")
                                        End If
                                    End If
                                ElseIf .OwnerID = glPlayerID Then
                                    'mine
                                    If .ObjTypeID = ObjectType.eUnit Then
                                        If .yProductionType = 0 Then
                                            'indicates tank kill
                                            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eIncrementTankDeath, -1, -1, -1, "")
                                        End If
                                    End If
                                End If
                            End If

                            .ClearParticleFX()
                            Try
                                Dim fTX As Single
                                Dim fTZ As Single
                                Dim fDist As Single
                                fTX = .LocX - goCamera.mlCameraX
                                fTZ = .LocZ - goCamera.mlCameraZ
                                fTX *= fTX
                                fTZ *= fTZ
                                fDist = CSng(Math.Sqrt(fTX + fTZ))
                                If fDist < 10000.0F Then
                                    fDist /= 10000.0F
                                    If .oUnitDef.HullSize > 0 Then fDist *= .oUnitDef.HullSize
                                    goCamera.ScreenShake(500 * fDist)
                                End If
                            Catch
                            End Try

                            Dim bDoDeathSequence As Boolean = True
                            If .OwnerID <> glPlayerID AndAlso .OwnerID <> gl_HARDCODE_PIRATE_PLAYER_ID Then
                                If .yVisibility = eVisibilityType.FacilityIntel Then
                                    bDoDeathSequence = False
                                End If
                            End If

                            If bDoDeathSequence = True AndAlso goEntityDeath Is Nothing = False Then goEntityDeath.AddNewDeathSequence(oEntity, lIdx)
                        End With
                    Case Else
                        oEnvir.oEntity(lIdx).ClearParticleFX()
                        oEnvir.lEntityIdx(lIdx) = -1
                        oEnvir.oEntity(lIdx) = Nothing

                        If yRemoveType = RemovalType.eJumping Then
                            'TODO: Do a jump special appearance effect
                        End If
                End Select



            End If
        End If



    End Sub

    Private Sub __ReactionThreadStart()
        'ok, here we go... give everyone a chance to do something heroic
        Threading.Thread.Sleep(200)

        'TODO: For now, this is the only message that we need to exit our thread for... 
        'we may want to put this in other places too, it is a good idea to avoid function call locks from killing I/O
        HandlePlayerCurrentEnvirMsg(myReactionData)

        moReactionThread.Abort()

        moReactionThread = Nothing
    End Sub

    Private Sub HandleUnitHPUpdateMsg(ByVal yData() As Byte)
        Dim X As Int32
        'Dim lGrid As Int32

        msCurrentLoc = "HandleUnitHPUpdateMsg"

        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)


        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing = False Then
            Dim lCurUB As Int32 = -1
            If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
            For X = 0 To lCurUB
                If oEnvir.lEntityIdx(X) = lObjID Then
                    Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                    If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iObjTypeID Then
                        With oEntity
                            .yArmorHP(0) = yData(8)
                            .yArmorHP(1) = yData(9)
                            .yArmorHP(2) = yData(10)
                            .yArmorHP(3) = yData(11)
                            .yShieldHP = yData(12)
                            .yStructureHP = yData(13)
                            .bHPChanged = True
                            .lLastHPUpdate = glCurrentCycle

                            .CritList = System.BitConverter.ToInt32(yData, 14)
                            .TestForBurnFX()
                        End With
                        Return
                    End If
                End If
            Next X
        End If

    End Sub

    Private Sub HandleAddObjectMsg(ByVal yData() As Byte, ByVal bFromPrimary As Boolean)
        'ok, the primary server sent us an Add Object Message... yData has the pertinent information...
        ' an Add Object has the following format: ObjID, ObjTypeID, ParentID, ParentTypeID... data after that
        ' is specific to the object being added
        msCurrentLoc = "HandleAddObjectMsg"
        Dim lObjID As Int32
        Dim iObjTypeID As Int16
        Dim X As Int32
        Dim Y As Int32
        Dim lIdx As Int32
        Dim lFirstIndex As Int32
        Dim lPos As Int32

        Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yData, 0)
        Dim lTemp As Int32

        lObjID = System.BitConverter.ToInt32(yData, 2)
        iObjTypeID = System.BitConverter.ToInt16(yData, 6)
        lPos = 8

        If (mlBusyStatus And elBusyStatusType.PlayerDetails) <> 0 Then WriteToRPDFile("  -- " & lObjID)

        If NewTutorialManager.TutorialOn = True Then
            If NewTutorialManager.iStoreNextGUIDType = iObjTypeID Then NewTutorialManager.StoreGUID(lObjID, iObjTypeID)
        End If

        If iObjTypeID <> ObjectType.eProductionCost Then
            'Production Costs typically come across as 0 ID's at least for the global techs, so, only use their parents
            mlLastAddObjectID = lObjID
            miLastAddObjectTypeID = iObjTypeID
        End If

        'NOTE: Make sure that any object type that requires goCurrentEnvir is added to this list
        If goCurrentEnvir Is Nothing AndAlso (iObjTypeID = ObjectType.eUnit OrElse _
          iObjTypeID = ObjectType.eFacility OrElse iObjTypeID = ObjectType.eMineralCache) Then Exit Sub

        Select Case iObjTypeID
            Case ObjectType.eUnit, ObjectType.eFacility
                lIdx = -1
                lFirstIndex = -1
                For X = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) = lObjID Then
                        'AndAlso goCurrentEnvir.ObjTypeID = iObjTypeID Then
                        Dim oTmp As BaseEntity = goCurrentEnvir.oEntity(X)
                        If oTmp Is Nothing = False AndAlso oTmp.ObjTypeID = iObjTypeID Then
                            lIdx = X
                            Exit For
                        End If
                    ElseIf lFirstIndex = -1 AndAlso goCurrentEnvir.lEntityIdx(X) = -1 Then
                        lFirstIndex = X
                    End If
                Next X

                If lIdx = -1 Then
                    If lFirstIndex <> -1 Then
                        lIdx = lFirstIndex
                    Else
                        goCurrentEnvir.lEntityUB += 1
                        lIdx = goCurrentEnvir.lEntityUB
                        ReDim Preserve goCurrentEnvir.lEntityIdx(lIdx)
                        ReDim Preserve goCurrentEnvir.oEntity(lIdx)
                    End If
                    goCurrentEnvir.oEntity(lIdx) = New BaseEntity()
                End If

                'First, clear our index
                goCurrentEnvir.lEntityIdx(lIdx) = -1

                'Now, parse the values...
                With goCurrentEnvir.oEntity(lIdx)
                    .ObjectID = lObjID
                    .ObjTypeID = iObjTypeID
                    .yRelID = 0

                    'If lObjID = 701749 Then Stop
                    'If lObjID = 701761 Then Stop
                    'If lObjID = 701762 Then Stop

                    .lEnvirEntityIdx = lIdx
                    'If lObjID = 921482 Then Stop

                    If .oUnitDef Is Nothing Then .oUnitDef = New EntityDef()

                    Dim yTempFXColor As Byte = 0
                    If bFromPrimary = True Then
                        'now, get the current status
                        .CurrentStatus = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                        'The next 3 are always the Locational data
                        .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .LocAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                        .CellLocX = CInt(Math.Floor(.LocX / muSettings.EntityClipPlane))
                        .CellLocZ = CInt(Math.Floor(.LocZ / muSettings.EntityClipPlane))

                        'now, if we are moving, the next 3 are dests
                        If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                            .DestX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                            .DestZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                            .TrueDestAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        End If

                        'the next two are for both
                        .OwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        Dim iModelID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        .oMesh = goResMgr.GetMesh(iModelID)
                        .oUnitDef.ModelID = iModelID

                        'Now total velocity
                        If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                            .TotalVelocity = yData(lPos) : lPos += 1
                        End If

                        'Now, the remaining  items
                        .oUnitDef.MaxSpeed = yData(lPos) : lPos += 1
                        .oUnitDef.Maneuver = yData(lPos) : lPos += 1
                        .oUnitDef.OptRadarRange = yData(lPos) : lPos += 1
                        .oUnitDef.MaxRadarRange = yData(lPos) : lPos += 1
                        .yProductionType = yData(lPos) : lPos += 1
                        .SetHPsFromBurstMsg(yData(lPos)) : lPos += 1
                        yTempFXColor = yData(lPos) : lPos += 1
                    Else
                        Select Case lObjID Mod 4
                            Case 0
                                .CurrentStatus = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                .LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                .LocAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                                'Set these here just in case
                                .CellLocX = CInt(Math.Floor(.LocX / muSettings.EntityClipPlane))
                                .CellLocZ = CInt(Math.Floor(.LocZ / muSettings.EntityClipPlane))

                                If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                                    .DestX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    .DestZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    .TrueDestAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                                End If

                                .OwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                Dim iModelID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                                .oMesh = goResMgr.GetMesh(iModelID)
                                .oUnitDef.ModelID = iModelID

                                If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                                    .TotalVelocity = yData(lPos) : lPos += 1
                                End If

                                .oUnitDef.MaxSpeed = yData(lPos) : lPos += 1
                                .oUnitDef.Maneuver = yData(lPos) : lPos += 1
                                .oUnitDef.OptRadarRange = yData(lPos) : lPos += 1
                                .oUnitDef.MaxRadarRange = yData(lPos) : lPos += 1
                                .yProductionType = yData(lPos) : lPos += 1
                                .SetHPsFromBurstMsg(yData(lPos)) : lPos += 1
                                yTempFXColor = yData(lPos) : lPos += 1
                                If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                                    .oUnitDef.yFormationManeuver = yData(lPos) : lPos += 1
                                    .oUnitDef.yFormationMaxSpeed = yData(lPos) : lPos += 1
                                    .oUnitDef.fFormationAcceleration = .oUnitDef.yFormationManeuver / 100.0F
                                End If
                            Case 1
                                .oUnitDef.MaxSpeed = yData(lPos) : lPos += 1
                                .oUnitDef.Maneuver = yData(lPos) : lPos += 1
                                Dim yHPValue As Byte = yData(lPos) : lPos += 1
                                yTempFXColor = yData(lPos) : lPos += 1
                                .CurrentStatus = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                .OwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                                If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                                    .TotalVelocity = yData(lPos) : lPos += 1
                                    .oUnitDef.OptRadarRange = yData(lPos) : lPos += 1
                                    .DestX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    .LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    .DestZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    .LocAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                                    .TrueDestAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                                    .oUnitDef.yFormationManeuver = yData(lPos) : lPos += 1
                                    .oUnitDef.yFormationMaxSpeed = yData(lPos) : lPos += 1
                                    .oUnitDef.fFormationAcceleration = .oUnitDef.yFormationManeuver / 100.0F
                                Else
                                    .oUnitDef.OptRadarRange = yData(lPos) : lPos += 1
                                    .LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    .LocAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                                End If

                                'Set these here just in case
                                .CellLocX = CInt(Math.Floor(.LocX / muSettings.EntityClipPlane))
                                .CellLocZ = CInt(Math.Floor(.LocZ / muSettings.EntityClipPlane))

                                Dim iModelID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                                .oMesh = goResMgr.GetMesh(iModelID)
                                .oUnitDef.ModelID = iModelID
                                .oUnitDef.MaxRadarRange = yData(lPos) : lPos += 1
                                .yProductionType = yData(lPos) : lPos += 1

                                .SetHPsFromBurstMsg(yHPValue)
                            Case 2
                                .OwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                .CurrentStatus = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                .LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                                If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                                    .TrueDestAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                                    .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    .DestZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    Dim iModelID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                                    .oMesh = goResMgr.GetMesh(iModelID)
                                    .oUnitDef.ModelID = iModelID
                                    .DestX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    .LocAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                                    .oUnitDef.Maneuver = yData(lPos) : lPos += 1
                                    .TotalVelocity = yData(lPos) : lPos += 1
                                Else
                                    .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    Dim iModelID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                                    .oMesh = goResMgr.GetMesh(iModelID)
                                    .oUnitDef.ModelID = iModelID
                                    .LocAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                                    .oUnitDef.Maneuver = yData(lPos) : lPos += 1
                                End If

                                'Set these here just in case
                                .CellLocX = CInt(Math.Floor(.LocX / muSettings.EntityClipPlane))
                                .CellLocZ = CInt(Math.Floor(.LocZ / muSettings.EntityClipPlane))

                                .oUnitDef.OptRadarRange = yData(lPos) : lPos += 1
                                .oUnitDef.MaxRadarRange = yData(lPos) : lPos += 1
                                .oUnitDef.MaxSpeed = yData(lPos) : lPos += 1
                                If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                                    .oUnitDef.yFormationMaxSpeed = yData(lPos) : lPos += 1
                                    .oUnitDef.yFormationManeuver = yData(lPos) : lPos += 1
                                    .oUnitDef.fFormationAcceleration = .oUnitDef.yFormationManeuver / 100.0F
                                End If

                                .yProductionType = yData(lPos) : lPos += 1
                                yTempFXColor = yData(lPos) : lPos += 1
                                .SetHPsFromBurstMsg(yData(lPos)) : lPos += 1
                            Case Else
                                Dim yHPValue As Byte = yData(lPos) : lPos += 1
                                yTempFXColor = yData(lPos) : lPos += 1

                                .oUnitDef.Maneuver = yData(lPos) : lPos += 1
                                .oUnitDef.MaxSpeed = yData(lPos) : lPos += 1
                                .OwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                .CurrentStatus = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                .oUnitDef.OptRadarRange = yData(lPos) : lPos += 1

                                If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                                    .TotalVelocity = yData(lPos) : lPos += 1
                                    .DestZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    .DestX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                End If

                                .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                .LocAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                                .LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                                'Set these here just in case
                                .CellLocX = CInt(Math.Floor(.LocX / muSettings.EntityClipPlane))
                                .CellLocZ = CInt(Math.Floor(.LocZ / muSettings.EntityClipPlane))

                                If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                                    .TrueDestAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                                End If

                                Dim iModelID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                                .oMesh = goResMgr.GetMesh(iModelID)
                                .oUnitDef.ModelID = iModelID

                                .oUnitDef.MaxRadarRange = yData(lPos) : lPos += 1
                                .yProductionType = yData(lPos) : lPos += 1
                                If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                                    .oUnitDef.yFormationManeuver = yData(lPos) : lPos += 1
                                    .oUnitDef.yFormationMaxSpeed = yData(lPos) : lPos += 1
                                    .oUnitDef.fFormationAcceleration = .oUnitDef.yFormationManeuver / 100.0F
                                End If

                                .SetHPsFromBurstMsg(yHPValue)
                        End Select
                    End If

                    'Now, set all of our values everywhere
                    'If iMsgCode = GlobalMessageCode.eAddObjectCommand_CE Then
                    '	'.fOriginalDist = System.Math.Abs(.DestX - .LocX) + System.Math.Abs(.DestZ - .LocZ)

                    '	If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                    '		.LocY = 5000
                    '	Else : .LocY = -5000
                    '	End If
                    '	.bForceSetY = False
                    'Else : .bForceSetY = True
                    'End If
                    .bForceSetY = True

                    If .oMesh Is Nothing = False Then
                        .oUnitDef.FOW_OptRadarRange = (.oMesh.RangeOffset + .oUnitDef.OptRadarRange) * gl_FINAL_GRID_SQUARE_SIZE
                        .oUnitDef.FOW_MaxRadarRange = (.oMesh.RangeOffset + .oUnitDef.MaxRadarRange) * gl_FINAL_GRID_SQUARE_SIZE * 4
                    End If

                    Dim lHiValue As Int32 = (yTempFXColor And 240) \ 16
                    Dim lLowValue As Int32 = (yTempFXColor And 15)
                    .clrEngineFX = gclrEngines(lHiValue)
                    .clrShieldFX = gclrShields(lLowValue)

                    'Set us moving...
                    If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                        .SetDest(.DestX, .DestZ, CShort(LineAngleDegrees(CInt(.LocX), CInt(.LocZ), .DestX, .DestZ) * 10))
                    End If

                    If iMsgCode = GlobalMessageCode.eAddObjectCommand_CE Then
                        If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                            .yChangeEnvironment = ChangeEnvironmentType.eSystemToPlanet
                        Else : .yChangeEnvironment = ChangeEnvironmentType.ePlanetToSystem
                        End If
                    End If

                    .LastUpdateCycle = glCurrentCycle
                    If (.yProductionType And ProductionType.eProduction) <> 0 AndAlso .OwnerID = glPlayerID Then
                        TutorialManager.bFirstFactoryBuilt = True
                    End If

                    If goCurrentPlayer Is Nothing = False Then
                        If goCurrentPlayer.oGuild Is Nothing = False Then
                            .bGuildMember = goCurrentPlayer.oGuild.MemberInGuild(.OwnerID)
                        End If
                    End If

                    If .ObjTypeID = ObjectType.eFacility Then
                        If .yProductionType = ProductionType.eMining Then
                            TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.FirstMiningFacilityBuilt)
                        ElseIf .yProductionType = ProductionType.eTradePost Then
                            TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.TradepostBuilt)
                        ElseIf (.yProductionType And ProductionType.eProduction) <> 0 Then
                            TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.FirstProducerBuilt)
                        End If
                    Else
                        If .oUnitDef Is Nothing = False AndAlso (.oUnitDef.yChassisType And ChassisType.eSpaceBased) <> 0 Then
                            TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.BuildSpaceCapableUnit)
                        End If
                    End If
                End With

                'NOW, set our index
                goCurrentEnvir.lEntityIdx(lIdx) = lObjID
            Case ObjectType.ePlayer
                'Ok, the player object is being added...

                If lObjID = glPlayerID Then
                    If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player()
                    With goCurrentPlayer
                        .ObjectID = lObjID
                        .ObjTypeID = iObjTypeID

                        'Player name (20)
                        .PlayerName = GetStringFromBytes(yData, 8, 20)
                        'Empire name (20)
                        .EmpireName = GetStringFromBytes(yData, 28, 20)
                        'Race Name (20)
                        .RaceName = GetStringFromBytes(yData, 48, 20)

                        '.lSenateID = System.BitConverter.ToInt32(yData, 68)
                        .bIsMale = yData(68) <> 2
                        .CommEncryptLevel = System.BitConverter.ToInt16(yData, 72)
                        .EmpireTaxRate = yData(74)

                        lPos = 75
                        .blCredits = System.BitConverter.ToInt32(yData, lPos) : lPos += 8

                        .yPlayerTitle = yData(lPos) : lPos += 1
                        .lPlayerIcon = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lTechnologyScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lDiplomacyScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lMilitaryScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lPopulationScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lProductionScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lWealthScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lTotalScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .DeathBudgetBalance = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                        Dim lGuildID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        Dim lJoinedGuildOn As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        If lGuildID <> -1 Then
                            .oGuild = New Guild()
                            .oGuild.ObjectID = lGuildID
                            .oGuild.ObjTypeID = ObjectType.eGuild
                            .oGuild.dtJoined = GetDateFromNumber(lJoinedGuildOn)
                        End If

                        .BadWarDecCPIncrease = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .BadWarDecMoralePenalty = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .yPlayerPhase = yData(lPos) : lPos += 1
                        .lTutorialStep = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lAccountStatus = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                        .lStatusFlags = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                        If glPlayerID = .ObjectID Then
                            If .yPlayerPhase = 0 OrElse (.yPlayerPhase = 1 AndAlso .lTutorialStep < 315) Then
                                NewTutorialManager.TutorialOn = True
                                NewTutorialManager.SetTutorialStep(.lTutorialStep)
                            Else
                                NewTutorialManager.TutorialOn = False
                                NewTutorialManager.CheckReceivedFirstMsg()
                                TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.EndsTutorialInAurelium)
                            End If
                        End If
                        'Control Groups
                        If goControlGroups Is Nothing Then goControlGroups = New ControlGroups()

                    End With
                End If
            Case ObjectType.eMineralCache, ObjectType.eComponentCache       'NOTE: This assumes the object is received from REGION!!!

                'If iObjTypeID = ObjectType.eComponentCache Then Stop

                lIdx = -1
                lFirstIndex = -1
                For X = 0 To goCurrentEnvir.lCacheUB
                    If goCurrentEnvir.lCacheIdx(X) = lObjID Then
                        lIdx = X
                        Exit For
                    ElseIf lFirstIndex = -1 AndAlso goCurrentEnvir.lCacheIdx(X) = -1 Then
                        lFirstIndex = X
                    End If
                Next X

                If lIdx = -1 Then
                    If lFirstIndex <> -1 Then
                        lIdx = lFirstIndex
                    Else
                        goCurrentEnvir.lCacheUB += 1
                        lIdx = goCurrentEnvir.lCacheUB
                        ReDim Preserve goCurrentEnvir.lCacheIdx(lIdx)
                        ReDim Preserve goCurrentEnvir.oCache(lIdx)
                    End If
                    goCurrentEnvir.oCache(lIdx) = New MineralCache()
                End If

                goCurrentEnvir.lCacheIdx(lIdx) = lObjID
                With goCurrentEnvir.oCache(lIdx)
                    .ObjectID = lObjID
                    .ObjTypeID = iObjTypeID
                    .LocX = System.BitConverter.ToInt32(yData, 8)
                    .LocZ = System.BitConverter.ToInt32(yData, 12)
                    .MineralID = System.BitConverter.ToInt32(yData, 16)
                    .CacheTypeID = yData(20)

                    lTotalCache += 1
                    If .CacheTypeID <> 0 Then lTotalCacheType1 += 1
                End With
            Case ObjectType.eUnitDef, ObjectType.eFacilityDef
                lIdx = -1 : lFirstIndex = -1
                For X = 0 To glEntityDefUB
                    If glEntityDefIdx(X) = lObjID AndAlso goEntityDefs(X).ObjTypeID = iObjTypeID Then
                        lIdx = X
                        Exit For
                    ElseIf lFirstIndex = -1 AndAlso glEntityDefIdx(X) = -1 Then
                        lFirstIndex = X
                    End If
                Next X

                If lIdx = -1 Then
                    If lFirstIndex <> -1 Then
                        lIdx = lFirstIndex
                    Else
                        glEntityDefUB += 1
                        lIdx = glEntityDefUB
                        ReDim Preserve goEntityDefs(glEntityDefUB)
                        ReDim Preserve glEntityDefIdx(glEntityDefUB)
                    End If
                    goEntityDefs(lIdx) = New EntityDef()
                End If
                glEntityDefIdx(lIdx) = lObjID
                With goEntityDefs(lIdx)
                    .ProductionCost = New ProductionCost()

                    .ObjectID = lObjID
                    .ObjTypeID = iObjTypeID
                    lPos = 8
                    lPos += 4       'for ownerid
                    .Armor_MaxHP(0) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .Armor_MaxHP(1) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .Armor_MaxHP(2) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .Armor_MaxHP(3) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .Maneuver = yData(lPos) : lPos += 1
                    .MaxSpeed = yData(lPos) : lPos += 1
                    lPos += 1           'old fuel efficiency
                    '.FuelEfficiency = yData(lPos) : lPos += 1
                    .Structure_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .Hangar_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .HullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .Cargo_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .Shield_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .ShieldRecharge = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .ShieldRechargeFreq = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    ''Max Door Size (4)
                    'lPos += 4
                    ''Number of Doors (1)
                    'lPos += 1

                    '.Fuel_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    lPos += 4       'old fuel_cap

                    .Weapon_Acc = yData(lPos) : lPos += 1
                    .ScanResolution = yData(lPos) : lPos += 1
                    .OptRadarRange = yData(lPos) : lPos += 1
                    .MaxRadarRange = yData(lPos) : lPos += 1
                    .DisruptionResistance = yData(lPos) : lPos += 1
                    .PiercingResist = yData(lPos) : lPos += 1
                    .ImpactResist = yData(lPos) : lPos += 1
                    .BeamResist = yData(lPos) : lPos += 1
                    .ECMResist = yData(lPos) : lPos += 1
                    .FlameResist = yData(lPos) : lPos += 1
                    .ChemicalResist = yData(lPos) : lPos += 1
                    'Detection Resist (1)
                    lPos += 1
                    .ModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                    'DefName (20)
                    .DefName = GetStringFromBytes(yData, lPos, 20)
                    lPos += 20

                    If iObjTypeID = ObjectType.eFacilityDef Then
                        .WorkerFactor = System.BitConverter.ToInt32(yData, lPos) : lPos += 4 'WorkerFactor (4)
                        lPos += 1 'MaxFacilitySize (1)
                        .ProdFactor = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .PowerFactor = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    End If
                    .ProductionTypeID = yData(lPos) : lPos += 1
                    .RequiredProductionTypeID = yData(lPos) : lPos += 1
                    .yChassisType = yData(lPos) : lPos += 1
                    .yFXColors = yData(lPos) : lPos += 1
                    .yArmorIntegrityRoll = yData(lPos) : lPos += 1
                    .JamImmunity = yData(lPos) : lPos += 1
                    .JamStrength = yData(lPos) : lPos += 1
                    .JamTargets = yData(lPos) : lPos += 1
                    .JamEffect = yData(lPos) : lPos += 1

                    'Side Critical Hit Locations...
                    lPos += 16

                    lTemp = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .WeaponDefUB = lTemp - 1
                    ReDim .WeaponDefs(.WeaponDefUB)
                End With

                lTemp = 0

                For X = 0 To goEntityDefs(lIdx).WeaponDefUB
                    goEntityDefs(lIdx).WeaponDefs(lTemp) = New WeaponDef()
                    With goEntityDefs(lIdx).WeaponDefs(lTemp)
                        '.ArcID = yData(lPos) : lPos += 1
                        '.ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                        ''Weapon Name (20)
                        '.WeaponName = GetStringFromBytes(yData, lPos, 20)
                        'lPos += 20

                        '.WeaponType = yData(lPos) : lPos += 1
                        '.ROF_Delay = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        '.Range = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        '.Accuracy = yData(lPos) : lPos += 1
                        '.PiercingMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.PiercingMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.ImpactMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.ImpactMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.BeamMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.BeamMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.ECMMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.ECMMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.FlameMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.FlameMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.ChemicalMinDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.ChemicalMaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.WpnGroup = yData(lPos) : lPos += 1
                        '.lFirePowerRating = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        ''wpn tech id
                        'lPos += 4

                        '.AOERange = yData(lPos) : lPos += 1
                        '.WeaponSpeed = yData(lPos) : lPos += 1
                        '.Maneuver = yData(lPos) : lPos += 1

                        '.lEntityStatusGroup = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.lAmmoCap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.WpnDefID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                        '.MaxRangeAccuracy = .Range * (.Accuracy / 100.0F)
                        lPos = .FillFromMsg(yData, lPos)
                    End With
                    lTemp += 1
                Next X
                If gl_CLIENT_VERSION > 308 Then
                    goEntityDefs(lIdx).lExtendedFlags = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                End If
            Case ObjectType.eProductionCost
                lPos = 8

                lIdx = -1
                lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4    'Object ID
                Y = System.BitConverter.ToInt16(yData, lPos) : lPos += 2        'ObjTypeID

                Dim oProdCost As ProductionCost = Nothing

                'NOTE: POSITION SENSITIVE STUFF!!!!

                If Y = ObjectType.eAlloyTech OrElse Y = ObjectType.eArmorTech OrElse Y = ObjectType.eEngineTech OrElse _
                  Y = ObjectType.eHullTech OrElse Y = ObjectType.eRadarTech OrElse _
                  Y = ObjectType.eShieldTech OrElse Y = ObjectType.eWeaponTech OrElse Y = ObjectType.ePrototype OrElse Y = ObjectType.eSpecialTech Then
                    If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player()

                    Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lTemp, CShort(Y))
                    If oTech Is Nothing = False Then
                        'NOTE: If positions should ever change, we need to make sure this still works
                        If yData(lPos + 28) = 0 Then
                            If oTech.oProductionCost Is Nothing Then oTech.oProductionCost = New ProductionCost()
                            oProdCost = oTech.oProductionCost
                        Else
                            If oTech.oResearchCost Is Nothing Then oTech.oResearchCost = New ProductionCost()
                            oProdCost = oTech.oResearchCost
                        End If
                    End If
                Else
                    For X = 0 To glEntityDefUB
                        If glEntityDefIdx(X) = lTemp Then
                            If goEntityDefs(X).ObjTypeID = Y Then
                                If goEntityDefs(X).ProductionCost Is Nothing Then goEntityDefs(X).ProductionCost = New ProductionCost
                                oProdCost = goEntityDefs(X).ProductionCost
                                Exit For
                            End If
                        End If
                    Next X
                End If

                If oProdCost Is Nothing = False Then
                    With oProdCost
                        .PC_ID = lObjID
                        .ObjectID = lTemp
                        .ObjTypeID = CShort(Y)
                        .CreditCost = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                        '.CreditCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .ColonistCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .EnlistedCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .OfficerCost = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .PointsRequired = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                        .ProductionCostType = yData(lPos) : lPos += 1

                        .ItemCostUB = (System.BitConverter.ToInt32(yData, lPos)) - 1 : lPos += 4
                        ReDim .ItemCosts(.ItemCostUB)
                        For X = 0 To .ItemCostUB
                            .ItemCosts(X) = New ProductionCostItem()
                            .ItemCosts(X).ItemID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                            .ItemCosts(X).ItemTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                            .ItemCosts(X).QuantityNeeded = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        Next X
                    End With
                End If

                oProdCost = Nothing

            Case ObjectType.eMineral
                lIdx = -1 : lFirstIndex = -1
                Dim bPreGot As Boolean = False
                For X = 0 To glMineralUB
                    If glMineralIdx(X) = lObjID Then
                        lIdx = X
                        bPreGot = True
                        Exit For
                    ElseIf lFirstIndex = -1 AndAlso glMineralIdx(X) = -1 Then
                        lFirstIndex = X
                    End If
                Next X

                If lIdx = -1 Then
                    If lFirstIndex <> -1 Then
                        lIdx = lFirstIndex
                    Else
                        glMineralUB += 1
                        lIdx = glMineralUB
                        ReDim Preserve goMinerals(glMineralUB)
                        ReDim Preserve glMineralIdx(glMineralUB)
                    End If
                    goMinerals(lIdx) = New Mineral()
                End If

                lPos = 8

                'NOTE: we are actually getting a PlayerMineral object...

                glMineralIdx(lIdx) = lObjID
                With goMinerals(lIdx)

                    .bDiscovered = (yData(lPos) <> 0)
                    lPos += 1

                    'MineralName (20)
                    .MineralName = GetStringFromBytes(yData, lPos, 20)
                    lPos += 20

                    .ObjectID = lObjID
                    .ObjTypeID = ObjectType.eMineral
                    '.bArchived = (iMsgCode = GlobalMessageCode.eGetArchivedItems)

                    If .bDiscovered = True Then
                        If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.MineralDiscovered)
                        'Ok, more to go... next two is the number of properties we are getting back
                        Dim lCnt As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        For Y = 0 To lCnt - 1
                            'four bytes for property id
                            lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                            'Next 20 bytes is the Property Value Range Name
                            Dim sPropVal As String = GetStringFromBytes(yData, lPos, 20)
                            lPos += 20

                            'Next 20 bytes is the Property Value Range Name of the KNOWN rating
                            Dim sKnowVal As String = GetStringFromBytes(yData, lPos, 20)
                            lPos += 20

                            .SetMineralPropertyValue(lTemp, sPropVal, sKnowVal, yData(lPos)) : lPos += 1
                        Next Y
                    ElseIf bPreGot = False Then
                        goUILib.AddNotification("We have discovered a new mineral!", System.Drawing.Color.FromArgb(255, 0, 255, 255), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.DiscoverNewMineral)
                        If goSound Is Nothing = False Then
                            goSound.StartSound("Game Narrative\DiscoveredMineral.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
                        End If
                        If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.DiscoveredNewMaterial)
                    End If

                    .lLastMsgUpdate = glCurrentCycle
                End With

            Case ObjectType.eMineralProperty
                lIdx = -1 : lFirstIndex = -1
                For X = 0 To glMineralPropertyUB
                    If glMineralPropertyIdx(X) = lObjID Then
                        lIdx = X
                        Exit For
                    ElseIf lFirstIndex = -1 AndAlso glMineralPropertyIdx(X) = -1 Then
                        lFirstIndex = X
                    End If
                Next X

                If lIdx = -1 Then
                    If lFirstIndex <> -1 Then
                        lIdx = lFirstIndex
                    Else
                        glMineralPropertyUB += 1
                        lIdx = glMineralPropertyUB
                        ReDim Preserve goMineralProperty(glMineralPropertyUB)
                        ReDim Preserve glMineralPropertyIdx(glMineralPropertyUB)
                    End If
                    goMineralProperty(lIdx) = New MineralProperty
                End If

                With goMineralProperty(lIdx)
                    'MineralPropertyName (50)
                    .MineralPropertyName = GetStringFromBytes(yData, lPos, 50)
                    lPos += 50
                    .yKnowledgeLevel = yData(lPos) : lPos += 1

                    .ObjectID = lObjID
                    .ObjTypeID = iObjTypeID
                End With
                glMineralPropertyIdx(lIdx) = lObjID
                SortMineralProperties()
            Case ObjectType.eUnitGroup
                Dim lOwnerID As Int32 = System.BitConverter.ToInt32(yData, 8)
                If lOwnerID = glPlayerID Then
                    goCurrentPlayer.AddUnitGroupFromMsg(yData, lObjID)
                Else
                    'TODO: Determine what happens with other player's fleet intel
                End If

            Case ObjectType.ePlayerComm
                'the ONLY time a player should receive a player comm is if it belongs in one of their Mailboxes!
                If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player()
                Dim oMsg As PlayerComm = New PlayerComm()
                oMsg.FillFromMsg(yData)
                goCurrentPlayer.AddPlayerComm(oMsg)

                If oMsg.bMsgRead = False Then
                    frmEmailMain.lCurrentUnreadMessages += 1
                    If goUILib Is Nothing = False Then
                        Dim oWin As frmQuickBar = CType(goUILib.GetWindow("frmQuickBar"), frmQuickBar)
                        If oWin Is Nothing Then oWin = New frmQuickBar(goUILib)
                        If oWin Is Nothing = False Then oWin.NewEmailHasArrived()
                        oWin = Nothing
                    End If
                End If

            Case ObjectType.eTrade
                If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player()
                lIdx = goCurrentPlayer.GetOrAddTrade(lObjID)
                If lIdx <> -1 Then
                    Dim lPrevStatus As Int32 = goCurrentPlayer.moTrades(lIdx).yTradeState
                    goCurrentPlayer.moTrades(lIdx).PopulateFromMessage(yData)
                    'If mbBusy = False AndAlso lPrevStatus <> goCurrentPlayer.moTrades(lIdx).yTradeState Then
                    'If (mlBusyStatus = 0 OrElse ((goCurrentPlayer.moTrades(lIdx).yTradeState And (Trade.eTradeStateValues.TradeCompleted Or Trade.eTradeStateValues.TradeRejected)) = 0)) AndAlso (lPrevStatus <> goCurrentPlayer.moTrades(lIdx).yTradeState OrElse lPrevStatus = 0) Then
                    If (mlBusyStatus = 0 OrElse ((goCurrentPlayer.moTrades(lIdx).yTradeState And Trade.eTradeStateValues.TradeCompleted) = 0 AndAlso (goCurrentPlayer.moTrades(lIdx).yTradeState And Trade.eTradeStateValues.InProgress) = 0 AndAlso (goCurrentPlayer.moTrades(lIdx).yTradeState <> Trade.eTradeStateValues.TradeRejected))) AndAlso (lPrevStatus <> goCurrentPlayer.moTrades(lIdx).yTradeState OrElse lPrevStatus = 0) Then
                        If (mlBusyStatus And elBusyStatusType.TradesNotification) = 0 Then
                            mlBusyStatus = mlBusyStatus Or elBusyStatusType.TradesNotification
                            goUILib.AddNotification("Trade Agreements require your attention.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Dim ofrm As frmQuickBar = CType(goUILib.GetWindow("frmQuickBar"), frmQuickBar)
                            If ofrm Is Nothing Then ofrm = New frmQuickBar(goUILib)
                            If ofrm Is Nothing = False Then ofrm.TradeAgreementUpdate()
                            ofrm = Nothing
                        End If


                    End If
                End If

            Case ObjectType.eBudget
                If glPlayerID = Math.Abs(lObjID) Then
                    'Ok, my budget...
                    If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player
                    If goCurrentPlayer.oBudget Is Nothing Then goCurrentPlayer.oBudget = New Budget()
                    goCurrentPlayer.oBudget.FillFromMsg(yData)
                    'Else
                    '	'TODO: Espionage
                End If
            Case ObjectType.eAgent
                Dim oAgent As New Agent()
                oAgent.FillFromMsg(yData)
                If oAgent.lOwnerID = glPlayerID Then
                    lIdx = -1
                    lFirstIndex = -1
                    Dim bFound As Boolean = False
                    For X = 0 To goCurrentPlayer.AgentUB
                        If goCurrentPlayer.AgentIdx(X) = oAgent.ObjectID Then
                            lIdx = X
                            bFound = True
                            Exit For
                        ElseIf goCurrentPlayer.AgentIdx(X) = -1 AndAlso lFirstIndex = -1 Then
                            lFirstIndex = X
                        End If
                    Next X
                    If lIdx = -1 Then
                        If lFirstIndex = -1 Then
                            lFirstIndex = goCurrentPlayer.AgentUB + 1
                            ReDim Preserve goCurrentPlayer.AgentIdx(lFirstIndex)
                            ReDim Preserve goCurrentPlayer.Agents(lFirstIndex)
                            goCurrentPlayer.AgentIdx(lFirstIndex) = -1
                            goCurrentPlayer.AgentUB += 1
                        End If
                        lIdx = lFirstIndex
                    End If
                    goCurrentPlayer.Agents(lIdx) = oAgent
                    goCurrentPlayer.AgentIdx(lIdx) = oAgent.ObjectID
                    If bFound = False AndAlso (oAgent.lAgentStatus And AgentStatus.NewRecruit) <> 0 Then
                        If glCurrentCycle - miLastAgentMsgUpdate > 150 Then
                            miLastAgentMsgUpdate = glCurrentCycle

                            If NewTutorialManager.TutorialOn = True Then
                                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eAgentAvailable, -1, -1, -1, "")
                            End If
                            TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.GetFirstAgent)
                            If goUILib Is Nothing = False Then goUILib.AddNotification("New Agents are available for recruitment.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        End If
                    End If
                End If
            Case ObjectType.ePlayerItemIntel
                Dim oNewPII As New PlayerItemIntel()
                oNewPII.FillFromMsg(yData)
                oNewPII.bArchived = (iMsgCode = GlobalMessageCode.eGetArchivedItems)
                'now, find a spot
                If oNewPII.lPlayerID = glPlayerID Then
                    lIdx = -1
                    For X = 0 To glItemIntelUB
                        If glItemIntelIdx(X) = oNewPII.lItemID AndAlso goItemIntel(X).iItemTypeID = oNewPII.iItemTypeID AndAlso goItemIntel(X).lOtherPlayerID = oNewPII.lOtherPlayerID Then
                            lIdx = X
                            Exit For
                        End If
                    Next X
                    If lIdx = -1 Then
                        ReDim Preserve glItemIntelIdx(glItemIntelUB + 1)
                        ReDim Preserve goItemIntel(glItemIntelUB + 1)
                        glItemIntelIdx(glItemIntelUB + 1) = -2
                        glItemIntelUB += 1
                        lIdx = glItemIntelUB
                    End If
                    goItemIntel(lIdx) = oNewPII
                    glItemIntelIdx(lIdx) = oNewPII.lItemID
                End If

            Case ObjectType.ePlayerIntel
                Dim lTempPlayerID As Int32 = System.BitConverter.ToInt32(yData, 8)

                'TODO: is this what we want?
                If lTempPlayerID <> glPlayerID Then Return

                If lObjID <> glPlayerID Then
                    lIdx = -1
                    For X = 0 To glPlayerIntelUB
                        If glPlayerIntelIdx(X) = lObjID Then
                            lIdx = X
                            Exit For
                        End If
                    Next X
                    If lIdx = -1 Then
                        ReDim Preserve glPlayerIntelIdx(glPlayerIntelUB + 1)
                        ReDim Preserve goPlayerIntel(glPlayerIntelUB + 1)
                        goPlayerIntel(glPlayerIntelUB + 1) = New PlayerIntel
                        glPlayerIntelIdx(glPlayerIntelUB + 1) = lObjID
                        glPlayerIntelUB += 1
                        lIdx = glPlayerIntelUB
                    End If
                    goPlayerIntel(lIdx).FillFromPlayerIntelMsg(yData)
                    'Else
                    '	'TODO: Intel on someone else' intel?
                End If

            Case ObjectType.eGuild
                If goCurrentPlayer Is Nothing = False Then
                    If goCurrentPlayer.oGuild Is Nothing = False AndAlso lObjID = goCurrentPlayer.oGuild.ObjectID Then
                        goCurrentPlayer.oGuild.FillFromMsg(yData)
                    ElseIf goCurrentPlayer.oGuild Is Nothing Then
                        goCurrentPlayer.oGuild = New Guild()
                        goCurrentPlayer.oGuild.FillFromMsg(yData)
                    End If
                End If

                'TECHS ================================
            Case ObjectType.eAlloyTech
                Dim oAlloy As AlloyTech = New AlloyTech
                oAlloy.FillFromMsg(yData)

                If oAlloy.OwnerID = glPlayerID Then
                    oAlloy.bArchived = (iMsgCode = GlobalMessageCode.eGetArchivedItems)
                    If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player()
                    goCurrentPlayer.AddTech(CObj(oAlloy))
                    'Else
                    '	'TODO: What do we do with other player's technologies??? Espionage?
                End If

            Case ObjectType.eArmorTech
                Dim oArmor As ArmorTech = New ArmorTech
                oArmor.FillFromMsg(yData)

                If oArmor.OwnerID = glPlayerID Then
                    oArmor.bArchived = (iMsgCode = GlobalMessageCode.eGetArchivedItems)
                    If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player()
                    goCurrentPlayer.AddTech(CObj(oArmor))
                    'Else
                    '	'TODO: What do we do with other player's technologies??? Espionage?
                End If

            Case ObjectType.eEngineTech
                Dim oEngine As EngineTech = New EngineTech
                oEngine.FillFromMsg(yData)

                If oEngine.OwnerID = glPlayerID Then
                    oEngine.bArchived = (iMsgCode = GlobalMessageCode.eGetArchivedItems)
                    If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player()
                    goCurrentPlayer.AddTech(CObj(oEngine))
                    'Else
                    '	'TODO: What do we do with other player's technologies??? Espionage?
                End If

                'Case ObjectType.eHangarTech
                '    Dim oHangar As HangarTech = New HangarTech
                '    oHangar.FillFromMsg(yData)

                '    If oHangar.OwnerID = glPlayerID Then
                '        If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player()
                '        goCurrentPlayer.AddTech(CObj(oHangar))
                '    Else
                '        'TODO: What do we do with other player's technologies??? Espionage?
                '    End If
            Case ObjectType.eHullTech
                Dim oHull As HullTech = New HullTech()
                oHull.FillFromMsg(yData)
                If oHull.OwnerID = glPlayerID Then
                    oHull.bArchived = (iMsgCode = GlobalMessageCode.eGetArchivedItems)
                    If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player()
                    goCurrentPlayer.AddTech(CObj(oHull))
                    'Else
                    '	'TODO: what do we do with other player's technologies??? espionage?
                End If
            Case ObjectType.ePrototype
                Dim oPrototype As PrototypeTech = New PrototypeTech
                oPrototype.FillFromMsg(yData)
                If oPrototype.OwnerID = glPlayerID Then
                    oPrototype.bArchived = (iMsgCode = GlobalMessageCode.eGetArchivedItems)
                    If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player()
                    goCurrentPlayer.AddTech(CObj(oPrototype))
                    'Else
                    '	'TODO: what do we do with other player's technologies??? espionage?
                End If
            Case ObjectType.eRadarTech
                Dim oRadar As RadarTech = New RadarTech
                oRadar.FillFromMsg(yData)

                If oRadar.OwnerID = glPlayerID Then
                    oRadar.bArchived = (iMsgCode = GlobalMessageCode.eGetArchivedItems)
                    If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player()
                    goCurrentPlayer.AddTech(CObj(oRadar))
                    'Else
                    '	'TODO: What do we do with other player's technologies??? Espionage?
                End If
            Case ObjectType.eShieldTech
                Dim oShield As ShieldTech = New ShieldTech
                oShield.FillFromMsg(yData)

                If oShield.OwnerID = glPlayerID Then
                    oShield.bArchived = (iMsgCode = GlobalMessageCode.eGetArchivedItems)
                    If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player()
                    goCurrentPlayer.AddTech(CObj(oShield))
                    'Else
                    '	'TODO: What do we do with other player's technologies??? Espionage?
                End If
            Case ObjectType.eWeaponTech
                Dim oWeapon As WeaponTech = WeaponTech.CreateWeaponClass(yData)
                oWeapon.FillFromMsg(yData)

                If oWeapon.OwnerID = glPlayerID Then
                    oWeapon.bArchived = (iMsgCode = GlobalMessageCode.eGetArchivedItems)
                    If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player()
                    goCurrentPlayer.AddTech(CObj(oWeapon))
                    'Else
                    '	'TODO: What do we do with other player's technologies??? Espionage?
                End If
            Case ObjectType.eSpecialTech
                Dim oTech As SpecialTech = New SpecialTech
                oTech.FillFromMsg(yData)

                If oTech.OwnerID = glPlayerID Then
                    If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player()
                    oTech.bArchived = (iMsgCode = GlobalMessageCode.eGetArchivedItems)
                    goCurrentPlayer.AddTech(CObj(oTech))


                    If oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
                        If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.SchoolOfSciencesResearched)
                    End If

                    'Indicative of player requesting details
                    If oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
                        If mlBusyStatus = 0 AndAlso goUILib Is Nothing = False Then
                            If glCurrentCycle - mlLastAlert > 30 Then
                                mlLastAlert = glCurrentCycle
                                'TODO: Play an alert sound...
                                goUILib.AddNotification("New Special Technology Concepts are Ready to be Researched!", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            End If
                        End If
                    End If
                    'Else
                    '	'TODO: What do we do with other player's technologies??? Espionage?
                End If
            Case ObjectType.eWormhole
                Dim oWormhole As New Wormhole()
                oWormhole.FillFromAddMsg(yData)
        End Select

    End Sub

    Private Sub HandleAddPlayerRelMsg(ByVal yData() As Byte)
        'ok... this is a special message
        ' 2 byte message code... 4 Player Regards, 4 byte This Player, 1 byte Score
        msCurrentLoc = "HandleAddPlayerRelMsg"
        Dim lPlayerRegards As Int32
        Dim lThisPlayer As Int32
        Dim yScore As Byte

        lPlayerRegards = System.BitConverter.ToInt32(yData, 2)
        lThisPlayer = System.BitConverter.ToInt32(yData, 6)
        yScore = yData(10)
        Dim yTargetScore As Byte = yData(11)
        Dim lCyclesToNextEvent As Int32 = System.BitConverter.ToInt32(yData, 12)

        If lPlayerRegards = glPlayerID Then
            goCurrentPlayer.SetPlayerRel(lThisPlayer, yScore)
            Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRel(lThisPlayer)
            oRel.TargetScore = yTargetScore
            If lCyclesToNextEvent > -1 AndAlso oRel.TargetScore <> oRel.WithThisScore Then
                oRel.lNextUpdateCycle = glCurrentCycle + lCyclesToNextEvent
            Else
                oRel.lNextUpdateCycle = -1
            End If
        ElseIf lThisPlayer = glPlayerID Then
            goCurrentPlayer.SetPlayerRel(lPlayerRegards, yScore)
            Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRel(lPlayerRegards)
            oRel.TargetScore = yTargetScore
            If lCyclesToNextEvent > -1 AndAlso oRel.TargetScore <> oRel.WithThisScore Then
                oRel.lNextUpdateCycle = glCurrentCycle + lCyclesToNextEvent
            Else
                oRel.lNextUpdateCycle = -1
            End If
        End If

    End Sub

    Private Sub HandlePlayerCurrentEnvirMsg(ByVal yData() As Byte)
        Dim lPlayerID As Int32
        Dim lEnvirID As Int32
        Dim iEnvirTypeID As Int16

        Dim lGalaxyID As Int32
        Dim lSystemID As Int32
        Dim lPlanetID As Int32

        msCurrentLoc = "HandlePlayerCurrentEnvirMsg"

        Dim lPos As Int32 = 2

        '2 byte msg code, 4 byte player id, 4 byte envir id, 2 byte envir type
        lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lEnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        iEnvirTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        lGalaxyID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lSystemID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lPlanetID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim bInCurrDomain As Boolean = yData(lPos) <> 0 : lPos += 1
        bInCurrDomain = True
        If bInCurrDomain = False Then
            msPrimaryIP = GetStringFromBytes(yData, lPos, 20) : lPos += 20
            mlPrimaryPort = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        End If

        RaiseEvent ReceivedPlayerCurrentEnvironment(lPlayerID, lEnvirID, iEnvirTypeID, lGalaxyID, lSystemID, lPlanetID, bInCurrDomain)
    End Sub

    Private Sub HandleEnvirDomainMsg(ByVal yData() As Byte)
        Dim iPort As Int16
        Dim sIP As String

        msCurrentLoc = "HandleEnvirDomainMsg"

        '2 byte msg code, 2 byte port number, 4 byte IP
        iPort = System.BitConverter.ToInt16(yData, 2)
        sIP = CStr(yData(4)) & "." & CStr(yData(5)) & "." & CStr(yData(6)) & "." & CStr(yData(7))

        RaiseEvent ReceivedEnvironmentsDomain(sIP, iPort)
    End Sub

    Private Sub HandleMoveObjectCommand(ByVal yData() As Byte)
        Dim X As Int32
        Dim lUnitID As Int32
        Dim iUnitTypeID As Int16
        Dim DestX As Int32
        Dim DestZ As Int32
        Dim DestA As Int16
        Dim yChangeEnvironment As Byte = 0
        Dim yForcedSpeed As Byte = 0

        msCurrentLoc = "HandleMoveObjectCommand"

        Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yData, 0)
        lUnitID = System.BitConverter.ToInt32(yData, 2)
        iUnitTypeID = System.BitConverter.ToInt16(yData, 6)
        DestX = System.BitConverter.ToInt32(yData, 8)
        DestZ = System.BitConverter.ToInt32(yData, 12)
        'If lUnitID = 1026 Then Stop
        If yData.GetUpperBound(0) >= 18 Then
            If iMsgCode = GlobalMessageCode.eEntityChangeEnvironment Then yChangeEnvironment = yData(18) Else yForcedSpeed = yData(18)
        End If

        If goCurrentEnvir Is Nothing = False Then
            For X = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) = lUnitID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iUnitTypeID Then
                    With goCurrentEnvir.oEntity(X)
                        If .LocX = DestX And .LocZ = DestZ Then
                            'Ok, check the angle
                            DestA = System.BitConverter.ToInt16(yData, 16)
                            If DestA = -1 Then
                                DestA = .LocAngle
                            End If
                        Else
                            DestA = CShort(LineAngleDegrees(CInt(.LocX), CInt(.LocZ), DestX, DestZ) * 10)
                        End If


                        If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 AndAlso .OwnerID = glPlayerID AndAlso goSound Is Nothing = False Then
                            Dim lVal As Int32 = CInt((Rnd() * 8) + 1)
                            goSound.StartSound("RC" & lVal & ".wav", False, SoundMgr.SoundUsage.eRadioChatter, New Microsoft.DirectX.Vector3(.LocX, .LocY, .LocZ), Microsoft.DirectX.Vector3.Empty)
                        End If

                        .oUnitDef.yFormationManeuver = 0
                        .oUnitDef.yFormationMaxSpeed = yForcedSpeed
                        .oUnitDef.fFormationAcceleration = 0

                        .SetDest(DestX, DestZ, DestA)
                        .yChangeEnvironment = yChangeEnvironment

                        '.bDoNotRunSetYTarget = (yChangeEnvironment = ChangeEnvironmentType.ePlanetToSystem OrElse yChangeEnvironment = ChangeEnvironmentType.eSystemToPlanet)
                        If yChangeEnvironment = ChangeEnvironmentType.ePlanetToSystem Then
                            .bDoNotRunSetYTarget = True
                            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                                'planet to system, watching from planet means units move into sky
                                .mlTargetY = 10000
                            Else
                                'planet to system, watching from system, means units move up
                                .mlTargetY = 0
                            End If
                        ElseIf yChangeEnvironment = ChangeEnvironmentType.eSystemToPlanet Then
                            .bDoNotRunSetYTarget = True
                            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                                'moving from a system to a planet, my dest should be the loc on the planet to move to
                                If goCurrentEnvir.oGeoObject Is Nothing = False AndAlso CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
                                    .mlTargetY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(.LocX, .LocZ, True) + 1000)
                                Else : .mlTargetY = 2000
                                End If
                            Else
                                'movedown
                                .mlTargetY = -4000      '??? should be the planet's mid point
                            End If
                        Else
                            .bDoNotRunSetYTarget = False
                        End If

                        .LastUpdateCycle = glCurrentCycle


                    End With
                    Exit For
                End If
            Next X
        End If
    End Sub

    Private Sub HandleStopObjectCommand(ByVal yData() As Byte)
        msCurrentLoc = "HandleStopObjectCommand"
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lX As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim lZ As Int32 = System.BitConverter.ToInt32(yData, 12)
        Dim lA As Int16 = System.BitConverter.ToInt16(yData, 16)

        Dim X As Int32

        If goCurrentEnvir Is Nothing Then Return

        For X = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) = lID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iTypeID Then
                With goCurrentEnvir.oEntity(X)
                    .LocX = lX
                    .LocZ = lZ
                    .LocAngle = lA
                    .LocYaw = 0
                    If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then .CurrentStatus -= elUnitStatus.eUnitMoving
                    .ShutoffEngines()
                End With
                Exit For
            End If
        Next X
    End Sub

    'Private Sub HandleFireWeaponMsg(ByVal yData() As Byte, ByVal bHit As Boolean)
    '	'msgcode = 2 bytes, attackerid = 4, targetid = 4, weapontype = 1, shield = 1, armor1 = 1, armor2= 1, armor3 = 1,
    '	'   armor4 = 1, structure = 1
    '	msCurrentLoc = "HandleFireWeaponMsg"

    '	If frmMain.mbChangingEnvirs = True Then Return

    '	Dim X As Int32
    '	Dim lAttackerID As Int32 = System.BitConverter.ToInt32(yData, 2)
    '	Dim iAttackerTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
    '	Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, 8)
    '	Dim iTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
    '	Dim yWpnTypeID As Byte = yData(14)

    '	Dim bHitShields As Boolean = False
    '	If (yWpnTypeID And WeaponType.eShieldHitBitMask) <> 0 Then
    '		bHitShields = True
    '		yWpnTypeID -= CByte(yWpnTypeID - 128)
    '	End If

    '	'Dim lTargetIdx As Int32
    '	'Dim lAttackerIdx As Int32

    '	Dim oEnvir As BaseEnvironment = goCurrentEnvir
    '	If oEnvir Is Nothing = False Then
    '		Dim lCurUB As Int32 = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
    '		Dim oTargetEntity As BaseEntity = Nothing
    '		Dim oAttackerEntity As BaseEntity = Nothing
    '		For X = 0 To lCurUB
    '			If oEnvir.lEntityIdx(X) = lTargetID Then
    '				Dim oEntity As BaseEntity = oEnvir.oEntity(X)
    '				If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iTargetTypeID Then
    '					oTargetEntity = oEntity
    '					If bHitShields = True Then
    '						If oEntity.yShieldHP = 0 Then oEntity.yShieldHP = 1
    '					ElseIf oEntity.yShieldHP <> 0 Then
    '						oEntity.yShieldHP = 0
    '					End If
    '					Exit For
    '				End If
    '			End If
    '		Next X

    '		For X = 0 To lCurUB
    '			If oEnvir.lEntityIdx(X) = lAttackerID Then
    '				Dim oEntity As BaseEntity = oEnvir.oEntity(X)
    '				If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iAttackerTypeID Then
    '					oAttackerEntity = oEntity
    '					Exit For
    '				End If
    '			End If
    '		Next X

    '		'Now, add our effect to the weapon shot manager
    '		If oTargetEntity Is Nothing = False AndAlso oAttackerEntity Is Nothing = False Then
    '			If goWpnMgr Is Nothing = False Then
    '				goWpnMgr.AddNewEffect(oTargetEntity, oAttackerEntity, yWpnTypeID, bHit)
    '			End If
    '			oTargetEntity.lLastWeaponUpdate = glCurrentCycle
    '		End If
    '	End If

    'End Sub

    Private Sub HandleGetEntityDetailsMsg(ByVal yData() As Byte)
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim X As Int32

        msCurrentLoc = "HandleGetEntityDetailsMsg"

        'TODO: We need to flesh these out...
        Select Case iTypeID
            Case ObjectType.eMineralCache, ObjectType.eComponentCache
                If goCurrentEnvir Is Nothing = False Then
                    For X = 0 To goCurrentEnvir.lCacheUB
                        If goCurrentEnvir.lCacheIdx(X) = lID AndAlso goCurrentEnvir.oCache(X).ObjTypeID = iTypeID Then
                            With goCurrentEnvir.oCache(X)
                                .LocX = System.BitConverter.ToInt32(yData, 8)
                                .LocZ = System.BitConverter.ToInt32(yData, 12)
                                'MineralID is @ 16
                                .CacheTypeID = yData(20)
                                .Quantity = System.BitConverter.ToInt32(yData, 21)
                                .Concentration = System.BitConverter.ToInt32(yData, 25)

                                .DetailsRefreshed()
                                Exit For
                            End With
                        End If
                    Next X
                End If
            Case ObjectType.eUnit, ObjectType.eFacility
                If goCurrentEnvir Is Nothing = False Then
                    For X = 0 To goCurrentEnvir.lEntityUB
                        If goCurrentEnvir.lEntityIdx(X) = lID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iTypeID Then
                            With goCurrentEnvir.oEntity(X)
                                'Entity Name
                                .EntityName = GetStringFromBytes(yData, 8, 20)

                                .Ammo_Cap = yData(28)
                                'ExpLvl
                                .Exp_Level = yData(29)
                                'Targeting tactics
                                .iTargetingTactics = System.BitConverter.ToInt16(yData, 30)
                                'combat tactics
                                .iCombatTactics = System.BitConverter.ToInt32(yData, 32)

                                .lTetherPointX = System.BitConverter.ToInt32(yData, 36)
                                .lTetherPointZ = System.BitConverter.ToInt32(yData, 40)
                                'SetUnitWarpointUpkeep(.ObjectID, System.BitConverter.ToInt32(yData, 44))

                            End With
                            Exit For
                        End If
                    Next X
                End If
        End Select
    End Sub

    Private Sub HandleUpdatePlayerCredits(ByVal yData() As Byte)
        msCurrentLoc = "HandleUpdatePlayerCredits"
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim blCredits As Int64 = System.BitConverter.ToInt64(yData, 8)
        Dim blCashFlow As Int64 = System.BitConverter.ToInt64(yData, 16)
        'Dim blWarpoints As Int64 = System.BitConverter.ToInt64(yData, 24)

        'Dim lCurrentWPUpkeepCost As Int32 = System.BitConverter.ToInt32(yData, 32)

        If goCurrentPlayer Is Nothing Then Exit Sub

        If lID = goCurrentPlayer.ObjectID AndAlso iTypeID = ObjectType.ePlayer Then
            goCurrentPlayer.blCredits = blCredits
            goCurrentPlayer.blCashFlow = blCashFlow
            'goCurrentPlayer.blWarpoints = blWarpoints
            'goCurrentPlayer.lCurrentWarpointUpkeepCost = lCurrentWPUpkeepCost
            'If goCurrentPlayer.blWarpointsSession = -1 AndAlso blWarpoints > 0 Then goCurrentPlayer.blWarpointsSession = blWarpoints
        Else
            Return
        End If

        If goUILib Is Nothing = False Then
            If goUILib.GetWindow("frmBudget") Is Nothing = False Then
                Dim yMsg(5) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerBudget).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 2)
                moPrimary.SendData(yMsg)
            End If
        End If
    End Sub

    Private Sub HandleSetPlayerRel(ByVal yData() As Byte)
        msCurrentLoc = "HandleSetPlayerRel"
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, 6)
        Dim yRel As Byte = yData(10)
        Dim sMsg As String
        Dim cColor As System.Drawing.Color
        Dim X As Int32

        Dim lEnvirID As Int32 = -1
        Dim iEnvirTypeID As Int16 = -1
        Dim lEntityID As Int32 = -1
        Dim iEntityTypeID As Int16 = -1
        Dim lLocX As Int32 = Int32.MinValue
        Dim lLocZ As Int32 = Int32.MinValue
        If yData.GetUpperBound(0) >= 30 Then
            lEnvirID = System.BitConverter.ToInt32(yData, 11)
            iEnvirTypeID = System.BitConverter.ToInt16(yData, 15)
            lEntityID = System.BitConverter.ToInt32(yData, 17)
            iEntityTypeID = System.BitConverter.ToInt16(yData, 21)
            lLocX = System.BitConverter.ToInt32(yData, 23)
            lLocZ = System.BitConverter.ToInt32(yData, 27)
        End If

        Dim yCurrent As Int32 = -1
        'Someone is setting player relations with me
        If lTargetID = glPlayerID Then
            If yRel <= elRelTypes.eWar Then
                'ensure we do the same...
                goCurrentPlayer.SetPlayerRel(lPlayerID, yRel)

                If goCurrentEnvir Is Nothing = False Then
                    For X = 0 To goCurrentEnvir.lEntityUB
                        If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                            If goCurrentEnvir.oEntity(X).OwnerID = lPlayerID Then
                                goCurrentEnvir.oEntity(X).yRelID = yRel
                            End If
                        End If
                    Next X
                End If
            Else
                Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRel(lPlayerID)
                If oRel Is Nothing Then
                    If goUILib Is Nothing = False Then
                        goUILib.AddNotification("We have made first contact with another race!", Color.Teal, lEnvirID, iEnvirTypeID, lEntityID, iEntityTypeID, lLocX, lLocZ)

                        Dim ofrm As frmQuickBar = CType(goUILib.GetWindow("frmQuickBar"), frmQuickBar)
                        If ofrm Is Nothing Then ofrm = New frmQuickBar(goUILib)
                        If ofrm Is Nothing = False Then ofrm.DiplomacyUpdate()
                        ofrm = Nothing
                    End If
                    goCurrentPlayer.SetPlayerRel(lPlayerID, yRel)
                Else
                    If oRel.oPlayerIntel Is Nothing = False Then
                        yCurrent = oRel.oPlayerIntel.yRegardsCurrentPlayer
                    End If
                End If
            End If

            sMsg = GetCacheObjectValue(lPlayerID, ObjectType.ePlayer)
            If yCurrent > -1 AndAlso yCurrent <> yRel Then
                sMsg &= " has changed relations with you to from " & yCurrent & " to " & yRel & " "
            Else
                sMsg &= " has set relations with you to " & yRel & " "
            End If

            Dim sWavFile As String = "Game Narrative\DiplomaticRelChange.wav"

            If yRel <= elRelTypes.eWar Then
                sWavFile = "Game Narrative\WarDec.wav"
                sMsg &= "(WAR)"
                cColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            ElseIf yRel <= elRelTypes.eNeutral Then
                sMsg &= "(Neutral)"
                cColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            ElseIf yRel <= elRelTypes.ePeace Then
                sMsg &= "(Peace)"
                cColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            ElseIf yRel <= elRelTypes.eAlly Then
                sMsg &= "(ALLY)"
                cColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else
                sMsg &= "(BLOOD ALLY)"
                cColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            End If

            'Ok, regardless... do this correctly... play a sound file...
            If goSound Is Nothing = False Then
                goSound.StartSound(sWavFile, False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
            End If

            'Now, show a notification
            If goUILib Is Nothing = False Then
                goUILib.AddNotification(sMsg, cColor, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Dim ofrm As frmQuickBar = CType(goUILib.GetWindow("frmQuickBar"), frmQuickBar)
                If ofrm Is Nothing Then ofrm = New frmQuickBar(goUILib)
                If ofrm Is Nothing = False Then ofrm.DiplomacyUpdate()
                ofrm = Nothing
            End If
        ElseIf lPlayerID = glPlayerID Then
            'Ok, possibly a first contact...
            Dim yTargetScore As Byte = yRel
            Try
                yTargetScore = yData(11)
            Catch
            End Try

            Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRel(lTargetID)
            If oRel Is Nothing Then
                If goUILib Is Nothing = False Then
                    goUILib.AddNotification("We have made first contact with another race!", Color.Teal, lEnvirID, iEnvirTypeID, lEntityID, iEntityTypeID, lLocX, lLocZ)
                    Dim ofrm As frmQuickBar = CType(goUILib.GetWindow("frmQuickBar"), frmQuickBar)
                    If ofrm Is Nothing Then ofrm = New frmQuickBar(goUILib)
                    If ofrm Is Nothing = False Then ofrm.DiplomacyUpdate()
                    ofrm = Nothing
                End If
                goCurrentPlayer.SetPlayerRel(lPlayerID, yRel)
                If goSound Is Nothing = False Then
                    Dim sWavFile As String = "Game Narrative\DiplomaticRelChange.wav"
                    goSound.StartSound(sWavFile, False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
                End If
            Else
                If yTargetScore = 0 Then yTargetScore = yRel
                oRel.TargetScore = yTargetScore
                oRel.WithThisScore = yRel

                Dim lCycles As Int32 = -1
                Try
                    lCycles = System.BitConverter.ToInt32(yData, 12)
                Catch
                End Try
                If lCycles > -1 AndAlso oRel.TargetScore <> oRel.WithThisScore Then
                    oRel.lNextUpdateCycle = glCurrentCycle + lCycles
                Else
                    oRel.lNextUpdateCycle = -1
                End If

            End If
        End If
    End Sub

    Private Sub HandleRequestEntityContents(ByRef yData() As Byte)
        If goCurrentEnvir Is Nothing Then Return
        msCurrentLoc = "HandleRequestEntityContents"

        Dim lPos As Int32 = 2       'for msgcode
        Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lEntityDefID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lTotalCargoCap As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lTotalHangarCap As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lEntityParentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEntityParentTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lItemCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim iDefTypeID As Int16 = 0

        If iEntityTypeID = ObjectType.eUnit Then
            iDefTypeID = ObjectType.eUnitDef
        Else : iDefTypeID = ObjectType.eFacilityDef
        End If

        Try
            Dim colHangar As Collection = Nothing
            Dim colCargo As Collection = Nothing

            If goCurrentEnvir.ObjectID = lEntityParentID AndAlso goCurrentEnvir.ObjTypeID = iEntityParentTypeID Then
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) = lEntityID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iEntityTypeID Then
                        If goCurrentEnvir.oEntity(X).oUnitDef Is Nothing OrElse goCurrentEnvir.oEntity(X).oUnitDef.ObjectID <> lEntityDefID Then
                            'ok, got a unit def mismatch...
                            Dim bFound As Boolean = False
                            Dim lIdx As Int32
                            For lIdx = 0 To glEntityDefUB
                                If glEntityDefIdx(lIdx) = lEntityDefID AndAlso goEntityDefs(lIdx).ObjTypeID = iDefTypeID Then
                                    'What we really care about here is hangar and cargo contents... which we only update for non-space stations
                                    If goCurrentEnvir.oEntity(X).yProductionType <> ProductionType.eSpaceStationSpecial Then
                                        goCurrentEnvir.oEntity(X).oUnitDef.Hangar_Cap = goEntityDefs(lIdx).Hangar_Cap
                                        goCurrentEnvir.oEntity(X).oUnitDef.Cargo_Cap = goEntityDefs(lIdx).Cargo_Cap
                                    End If
                                    goCurrentEnvir.oEntity(X).oUnitDef.ObjectID = lEntityDefID
                                    bFound = True
                                    Exit For
                                End If
                            Next lIdx
                            If bFound = False Then
                                goCurrentEnvir.oEntity(X).oUnitDef.Cargo_Cap = lTotalCargoCap
                                goCurrentEnvir.oEntity(X).oUnitDef.Hangar_Cap = lTotalHangarCap
                            End If
                        End If

                        If goCurrentEnvir.oEntity(X).oUnitDef Is Nothing = False Then
                            With goCurrentEnvir.oEntity(X).oUnitDef
                                .Cargo_Cap = lTotalCargoCap
                                .Hangar_Cap = lTotalHangarCap
                            End With
                        End If

                        If goCurrentEnvir.oEntity(X).colHangar Is Nothing Then goCurrentEnvir.oEntity(X).colHangar = New Collection()
                        colHangar = goCurrentEnvir.oEntity(X).colHangar

                        If goCurrentEnvir.oEntity(X).colCargo Is Nothing Then goCurrentEnvir.oEntity(X).colCargo = New Collection()
                        colCargo = goCurrentEnvir.oEntity(X).colCargo

                        'Indicate that the contents changed
                        goCurrentEnvir.oEntity(X).ContentsLastUpdate = glCurrentCycle

                        Exit For

                    End If
                Next X
            Else
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) = lEntityParentID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iEntityParentTypeID Then
                        Dim oTmp As EntityContents
                        Dim bFound As Boolean = False

                        'we indicate that this entity's contents have also changed
                        goCurrentEnvir.oEntity(X).ContentsLastUpdate = glCurrentCycle

                        If goCurrentEnvir.oEntity(X).colHangar Is Nothing Then goCurrentEnvir.oEntity(X).colHangar = New Collection()

                        If iEntityTypeID = ObjectType.eFacility OrElse iEntityTypeID = ObjectType.eUnit Then
                            If goCurrentEnvir.oEntity(X).colHangar Is Nothing = False Then
                                For Each oTmp In goCurrentEnvir.oEntity(X).colHangar
                                    If oTmp.ObjectID = lEntityID AndAlso oTmp.ObjTypeID = iEntityTypeID Then
                                        bFound = True
                                        oTmp.ContentsLastUpdate = glCurrentCycle

                                        If oTmp.colHangar Is Nothing Then oTmp.colHangar = New Collection()
                                        colHangar = oTmp.colHangar

                                        If oTmp.colCargo Is Nothing Then oTmp.colCargo = New Collection()
                                        colCargo = oTmp.colCargo

                                        Exit For
                                    End If
                                Next
                                oTmp = Nothing
                            End If

                            If bFound = False Then
                                oTmp = New EntityContents()
                                With oTmp
                                    .ObjectID = lEntityID
                                    .ObjTypeID = iEntityTypeID

                                    .oDetailItem = Nothing
                                    For Y As Int32 = 0 To glEntityDefUB
                                        If glEntityDefIdx(Y) = lEntityDefID AndAlso goEntityDefs(Y).ObjTypeID = iDefTypeID Then
                                            .oDetailItem = goEntityDefs(Y)
                                            Exit For
                                        End If
                                    Next Y

                                    If oTmp.colHangar Is Nothing Then oTmp.colHangar = New Collection()
                                    colHangar = oTmp.colHangar
                                    If oTmp.colCargo Is Nothing Then oTmp.colCargo = New Collection()
                                    colCargo = oTmp.colCargo

                                    .ContentsLastUpdate = glCurrentCycle

                                    .lQuantity = 1
                                    If goCurrentEnvir.oEntity(X).colHangar Is Nothing Then goCurrentEnvir.oEntity(X).colHangar = New Collection()
                                    goCurrentEnvir.oEntity(X).colHangar.Add(oTmp) ', "ID" & .ObjectID)
                                End With
                            End If
                        Else
                            If goCurrentEnvir.oEntity(X).colCargo Is Nothing = False Then
                                For Each oTmp In goCurrentEnvir.oEntity(X).colCargo
                                    If oTmp.ObjectID = lEntityID AndAlso oTmp.ObjTypeID = iEntityTypeID Then
                                        bFound = True
                                        oTmp.ContentsLastUpdate = glCurrentCycle

                                        If oTmp.colHangar Is Nothing Then oTmp.colHangar = New Collection()
                                        colHangar = oTmp.colHangar

                                        If oTmp.colCargo Is Nothing Then oTmp.colCargo = New Collection()
                                        colCargo = oTmp.colCargo

                                        Exit For
                                    End If
                                Next
                                oTmp = Nothing
                            End If
                        End If

                        Exit For
                    End If
                Next X
            End If

            'Store any contents messages we may have already received for my children...
            Dim lObjIDs() As Int32 = Nothing
            Dim iObjTypeIDs() As Int16 = Nothing
            Dim colHangars() As Collection = Nothing
            Dim colCargos() As Collection = Nothing
            Dim lUB As Int32 = -1
            For Each oTmp As EntityContents In colHangar
                If oTmp.colCargo Is Nothing = False OrElse oTmp.colHangar Is Nothing = False Then
                    lUB += 1
                    ReDim Preserve lObjIDs(lUB)
                    ReDim Preserve iObjTypeIDs(lUB)
                    ReDim Preserve colHangars(lUB)
                    ReDim Preserve colCargos(lUB)
                    lObjIDs(lUB) = oTmp.ObjectID
                    iObjTypeIDs(lUB) = oTmp.ObjTypeID
                    colHangars(lUB) = oTmp.colHangar
                    colCargos(lUB) = oTmp.colCargo
                End If
            Next oTmp

            colCargo.Clear()
            colHangar.Clear()

            For X As Int32 = 0 To lItemCnt - 1
                Dim yType As Byte = yData(lPos) : lPos += 1

                Dim oNewItem As New EntityContents()
                With oNewItem
                    .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                    Dim lDetailID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim iDetailTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .lQuantity = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                    Dim lHullSize As Int32 = 0
                    If .lQuantity < -1 Then
                        lHullSize = Math.Abs(.lQuantity)
                        .lQuantity = -1
                    End If

                    .ContentsLastUpdate = glCurrentCycle

                    .lDetailItemID = lDetailID
                    .iDetailItemTypeID = iDetailTypeID

                    If .ObjTypeID = ObjectType.eMineralCache Then
                        For Y As Int32 = 0 To glMineralUB
                            If glMineralIdx(Y) = lDetailID Then
                                .oDetailItem = goMinerals(Y)
                                Exit For
                            End If
                        Next Y
                    ElseIf .ObjTypeID = ObjectType.eUnit Then
                        Dim bFound As Boolean = False
                        For Y As Int32 = 0 To glEntityDefUB
                            If glEntityDefIdx(Y) = lDetailID Then
                                .oDetailItem = goEntityDefs(Y)
                                bFound = True
                                Exit For
                            End If
                        Next Y

                        If bFound = False Then
                            .oDetailItem = New EntityDef()
                            With CType(.oDetailItem, EntityDef)
                                .HullSize = lHullSize
                            End With
                        End If

                        'If bFound = False Then
                        '    For Y As Int32 = 0 To glItemIntelUB
                        '        If glItemIntelIdx(Y) = lDetailID Then
                        '            If goItemIntel(Y).iItemTypeID = iDetailTypeID Then
                        '                .oDetailItem = goItemIntel(Y)
                        '                bFound = True
                        '                Exit For
                        '            End If
                        '        End If
                        '    Next Y 
                        'End If
                    ElseIf .ObjTypeID = ObjectType.eComponentCache Then
                        .oDetailItem = goCurrentPlayer.GetTech(lDetailID, iDetailTypeID)
                    End If
                End With

                If oNewItem.ObjectID <> -1 AndAlso oNewItem.ObjTypeID <> -1 Then
                    Try
                        If yType = 0 Then   'hangar
                            colHangar.Add(oNewItem) ', "ID" & oNewItem.ObjectID)
                        Else                'cargo
                            colCargo.Add(oNewItem) ', "ID" & oNewItem.ObjectID)
                        End If
                    Catch
                    End Try
                End If
            Next X

            'Now, populate my children's resource lists with what I already knew
            For Each oTmp As EntityContents In colHangar
                For X As Int32 = 0 To lUB
                    If lObjIDs(X) = oTmp.ObjectID AndAlso iObjTypeIDs(X) = oTmp.ObjTypeID Then
                        oTmp.colCargo = colCargos(X)
                        oTmp.colHangar = colHangars(X)
                        Exit For
                    End If
                Next X
            Next
        Catch ex As Exception
            goUILib.AddNotification(msCurrentLoc & ": " & ex.Message, Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End Try
    End Sub

    'Private Sub HandleRequestEntityContents(ByVal yData() As Byte)
    '    If goCurrentEnvir Is Nothing Then Exit Sub
    '    msCurrentLoc = "HandleRequestEntityContents"
    '    Dim lContainer As Int32 = System.BitConverter.ToInt32(yData, 2)
    '    Dim iContainerTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
    '    Dim yType As Byte = yData(8)

    '    Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 9)
    '    Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 13)

    '    Dim lDetailID As Int32 = System.BitConverter.ToInt32(yData, 15)
    '    Dim iDetailTypeID As Int16 = System.BitConverter.ToInt16(yData, 19)

    '    Dim lQuantity As Int32 = System.BitConverter.ToInt32(yData, 21)

    '    Dim lContainersDefID As Int32 = System.BitConverter.ToInt32(yData, 25)
    '    Dim iUnitDefTypeID As Int16

    '    Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, 29)
    '    Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, 33)

    '    Dim X As Int32
    '    Dim oNewItem As EntityContents = New EntityContents()

    '    Dim colTemp As Collection = Nothing

    '    oNewItem.ObjectID = lObjID
    '    oNewItem.ObjTypeID = iObjTypeID
    '    oNewItem.lQuantity = lQuantity


    '    Dim bClearAll As Boolean = False

    '    If lObjID = -1 AndAlso iObjTypeID = -1 AndAlso lDetailID = -1 AndAlso iDetailTypeID = -1 Then
    '        'ok, this list is clear
    '        bClearAll = True
    '    End If

    '    If iObjTypeID = ObjectType.eMineralCache Then
    '        For X = 0 To glMineralUB
    '            If glMineralIdx(X) = lDetailID Then
    '                oNewItem.oDetailItem = goMinerals(X)
    '                Exit For
    '            End If
    '        Next X
    '    ElseIf iObjTypeID = ObjectType.eUnit Then
    '        For X = 0 To glEntityDefUB
    '            If glEntityDefIdx(X) = lDetailID Then
    '                oNewItem.oDetailItem = goEntityDefs(X)
    '                Exit For
    '            End If
    '        Next X
    '    ElseIf iObjTypeID = ObjectType.eComponentCache Then
    '        oNewItem.oDetailItem = goCurrentPlayer.GetTech(lDetailID, iDetailTypeID)
    '    End If

    '    'Parent will either be the environment or a containing object within the environment
    '    If goCurrentEnvir.ObjectID = lParentID AndAlso goCurrentEnvir.ObjTypeID = iParentTypeID Then
    '        For X = 0 To goCurrentEnvir.lEntityUB
    '            If goCurrentEnvir.lEntityIdx(X) = lContainer AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iContainerTypeID Then

    '                If bClearAll = True Then
    '                    goCurrentEnvir.oEntity(X).colCargo = Nothing
    '                    goCurrentEnvir.oEntity(X).colHangar = Nothing
    '                    goCurrentEnvir.oEntity(X).ContentsLastUpdate = glCurrentCycle
    '                    'Return
    '                End If

    '                'let's associate our def ID if needed
    '                If goCurrentEnvir.oEntity(X).oUnitDef Is Nothing OrElse goCurrentEnvir.oEntity(X).oUnitDef.ObjectID <> lContainersDefID Then
    '                    'ok, got a unit def mismatch...
    '                    If iContainerTypeID = ObjectType.eUnit Then
    '                        iUnitDefTypeID = ObjectType.eUnitDef
    '                    Else : iUnitDefTypeID = ObjectType.eFacilityDef
    '                    End If

    '                    Dim lIdx As Int32
    '                    For lIdx = 0 To glEntityDefUB
    '                        If glEntityDefIdx(lIdx) = lContainersDefID AndAlso goEntityDefs(lIdx).ObjTypeID = iUnitDefTypeID Then
    '                            'What we really care about here is hangar and cargo contents... which we only update for non-space stations
    '                            If goCurrentEnvir.oEntity(X).yProductionType <> ProductionType.eSpaceStationSpecial Then
    '                                goCurrentEnvir.oEntity(X).oUnitDef.Hangar_Cap = goEntityDefs(lIdx).Hangar_Cap
    '                                goCurrentEnvir.oEntity(X).oUnitDef.Cargo_Cap = goEntityDefs(lIdx).Cargo_Cap
    '                            End If
    '                            goCurrentEnvir.oEntity(X).oUnitDef.ObjectID = lContainersDefID

    '                            Exit For
    '                        End If
    '                    Next lIdx
    '                End If

    '                If bClearAll = True Then Return

    '                If yType = 0 Then
    '                    If goCurrentEnvir.oEntity(X).colHangar Is Nothing Then goCurrentEnvir.oEntity(X).colHangar = New Collection()
    '                    colTemp = goCurrentEnvir.oEntity(X).colHangar
    '                Else
    '                    If goCurrentEnvir.oEntity(X).colCargo Is Nothing Then goCurrentEnvir.oEntity(X).colCargo = New Collection()
    '                    colTemp = goCurrentEnvir.oEntity(X).colCargo
    '                End If

    '                'Indicate that the contents changed
    '                goCurrentEnvir.oEntity(X).ContentsLastUpdate = glCurrentCycle

    '                Exit For
    '            End If
    '        Next X
    '    Else
    '        For X = 0 To goCurrentEnvir.lEntityUB
    '            If goCurrentEnvir.lEntityIdx(X) = lParentID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iParentTypeID Then
    '                Dim oTmp As EntityContents
    '                Dim bFound As Boolean = False

    '                'we indicate that this entity's contents have also changed
    '                goCurrentEnvir.oEntity(X).ContentsLastUpdate = glCurrentCycle

    '                If goCurrentEnvir.oEntity(X).colHangar Is Nothing Then goCurrentEnvir.oEntity(X).colHangar = New Collection()

    '                If iContainerTypeID = ObjectType.eFacility OrElse iContainerTypeID = ObjectType.eUnit Then
    '                    If goCurrentEnvir.oEntity(X).colHangar Is Nothing = False Then
    '                        For Each oTmp In goCurrentEnvir.oEntity(X).colHangar
    '                            If oTmp.ObjectID = lContainer AndAlso oTmp.ObjTypeID = iContainerTypeID Then
    '                                bFound = True
    '                                oTmp.ContentsLastUpdate = glCurrentCycle
    '                                If yType = 0 Then
    '                                    If oTmp.colHangar Is Nothing Then oTmp.colHangar = New Collection()
    '                                    colTemp = oTmp.colHangar
    '                                Else
    '                                    If oTmp.colCargo Is Nothing Then oTmp.colCargo = New Collection()
    '                                    colTemp = oTmp.colCargo
    '                                End If

    '                                If bClearAll = True Then
    '                                    oTmp.colCargo = Nothing
    '                                    oTmp.colHangar = Nothing
    '                                    Return
    '                                End If

    '                                Exit For
    '                            End If
    '                        Next
    '                        oTmp = Nothing
    '                    End If

    '                    If bFound = False Then
    '                        oTmp = New EntityContents()
    '                        With oTmp
    '                            .ObjectID = lContainer
    '                            .ObjTypeID = iContainerTypeID

    '                            .oDetailItem = Nothing
    '                            For Y As Int32 = 0 To glEntityDefUB
    '                                If glEntityDefIdx(Y) = lContainersDefID AndAlso goEntityDefs(Y).ObjTypeID = iUnitDefTypeID Then
    '                                    .oDetailItem = goEntityDefs(Y)
    '                                    Exit For
    '                                End If
    '                            Next Y

    '                            If yType = 0 Then
    '                                If oTmp.colHangar Is Nothing Then oTmp.colHangar = New Collection()
    '                                colTemp = oTmp.colHangar
    '                            Else
    '                                If oTmp.colCargo Is Nothing Then oTmp.colCargo = New Collection()
    '                                colTemp = oTmp.colCargo
    '                            End If

    '                            .ContentsLastUpdate = glCurrentCycle

    '                            .lQuantity = 1
    '                            If goCurrentEnvir.oEntity(X).colHangar Is Nothing Then goCurrentEnvir.oEntity(X).colHangar = New Collection()
    '                            goCurrentEnvir.oEntity(X).colHangar.Add(oTmp, "ID" & .ObjectID)
    '                        End With
    '                    End If
    '                Else
    '                    If goCurrentEnvir.oEntity(X).colCargo Is Nothing = False Then
    '                        For Each oTmp In goCurrentEnvir.oEntity(X).colCargo
    '                            If oTmp.ObjectID = lContainer AndAlso oTmp.ObjTypeID = iContainerTypeID Then
    '                                bFound = True
    '                                oTmp.ContentsLastUpdate = glCurrentCycle
    '                                If yType = 0 Then
    '                                    If oTmp.colHangar Is Nothing Then oTmp.colHangar = New Collection()
    '                                    colTemp = oTmp.colHangar
    '                                Else
    '                                    If oTmp.colCargo Is Nothing Then oTmp.colCargo = New Collection()
    '                                    colTemp = oTmp.colCargo
    '                                End If

    '                                If bClearAll = True Then
    '                                    oTmp.colCargo = Nothing
    '                                    oTmp.colHangar = Nothing
    '                                    Return
    '                                End If

    '                                Exit For
    '                            End If
    '                        Next
    '                        oTmp = Nothing
    '                    End If
    '                End If

    '                Exit For
    '            End If
    '        Next X
    '    End If

    '    If colTemp Is Nothing = False AndAlso lObjID <> -1 AndAlso iObjTypeID <> -1 Then
    '        On Error Resume Next
    '        If colTemp.Contains("ID" & lObjID) Then
    '            If CType(colTemp.Item("ID" & lObjID), EntityContents).lQuantity <> oNewItem.lQuantity Then
    '                colTemp.Remove("ID" & lObjID)
    '            Else
    '                CType(colTemp.Item("ID" & lObjID), EntityContents).ContentsLastUpdate = glCurrentCycle
    '                Return
    '            End If
    '        End If
    '        oNewItem.ContentsLastUpdate = glCurrentCycle
    '        colTemp.Add(oNewItem, "ID" & lObjID)
    '        colTemp = Nothing
    '    End If


    'End Sub

    Private Sub HandleProductionStatusMsg(ByVal yData() As Byte)
        msCurrentLoc = "HandleProductionStatusMsg"
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        'MSC - made perc a single to increase precision...
        'Dim yPerc As Byte = yData(8)
        'Dim lCurrentCycle As Int32 = System.BitConverter.ToInt32(yData, 9)
        'Dim lFinishCycle As Int32 = System.BitConverter.ToInt32(yData, 13)
        'Dim lProdID As Int32 = System.BitConverter.ToInt32(yData, 17)
        'Dim iProdTypeID As Int16 = System.BitConverter.ToInt16(yData, 21)
        Dim fPerc As Single = System.BitConverter.ToSingle(yData, 8)
        Dim lCurrentCycle As Int32 = System.BitConverter.ToInt32(yData, 12)
        Dim lFinishCycle As Int32 = System.BitConverter.ToInt32(yData, 16)
        Dim lProdID As Int32 = System.BitConverter.ToInt32(yData, 20)
        Dim iProdTypeID As Int16 = System.BitConverter.ToInt16(yData, 24)
        Dim iModelID As Int16 = System.BitConverter.ToInt16(yData, 26)
        Dim iProductionAngle As Int16 = System.BitConverter.ToInt16(yData, 28)

        Dim X As Int32

        If goCurrentEnvir Is Nothing Then Exit Sub

        If iObjTypeID = ObjectType.eUnit AndAlso yData.Length > 33 Then
            Try
                Dim lPos As Int32 = 30
                Dim lQueueUB As Int32 = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4

                goCurrentEnvir.ClearUnitsProdQueue(lObjID, iObjTypeID)
                For X = 0 To lQueueUB
                    Dim oNewItem As New UnitProdQueue()
                    With oNewItem
                        .lProdID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .iLocA = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        .iBuilderTypeID = ObjectType.eUnit
                        .iModelID = 0
                        .iProdTypeID = ObjectType.eFacilityDef
                        .lBuilderID = lObjID
                        .oModel = Nothing
                    End With
                    goCurrentEnvir.AddUnitProdQueueItem(oNewItem)
                Next X
            Catch
            End Try
        End If

        Dim bFound As Boolean = False
        Try
            For X = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) = lObjID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iObjTypeID Then
                    goCurrentEnvir.oEntity(X).ReceiveProductionStatusMsg(lCurrentCycle, lFinishCycle, fPerc, iModelID)
                    goCurrentEnvir.oEntity(X).lProducingID = lProdID
                    goCurrentEnvir.oEntity(X).iProducingTypeID = iProdTypeID
                    goCurrentEnvir.oEntity(X).iProductionAngle = iProductionAngle
                    bFound = True

                    If NewTutorialManager.TutorialOn = True Then
                        If goCurrentEnvir.oEntity(X).bProducing = False AndAlso goCurrentEnvir.oEntity(X).bSelected = True Then
                            If iObjTypeID = ObjectType.eUnit Then
                                If goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eFacility Then
                                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eIdleEngineerSelected, lObjID, iObjTypeID, ProductionType.eFacility, "")
                                End If
                            ElseIf iObjTypeID = ObjectType.eFacility Then
                                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eIdleFacilitySelected, lObjID, iObjTypeID, goCurrentEnvir.oEntity(X).yProductionType, "")
                            End If
                        End If
                    End If

                    Exit For
                End If
            Next X

            If bFound = False Then
                'Ok, try to find it in a child...
                'TODO: This is pisspoor slow... to speed this up, we could take a server msg hit and send down the parent ID/type id
                Dim oTmpChild As StationChild = Nothing
                For X = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                        oTmpChild = goCurrentEnvir.oEntity(X).GetChild(lObjID, iObjTypeID)
                        If oTmpChild Is Nothing = False Then
                            With oTmpChild
                                .ReceiveProductionStatusMsg(lCurrentCycle, lFinishCycle, fPerc)
                                .lProdID = lProdID
                                .iProdTypeID = iProdTypeID
                            End With
                        End If
                    End If
                Next X
            End If
        Catch
        End Try
    End Sub

    Private Sub HandleRequestDockFail(ByVal yData() As Byte)
        msCurrentLoc = "HandleRequestDockFail"
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yReason As Byte = yData(lPos) : lPos += 1

        'Now, generate our string
        Dim sMsg As String
        Dim sWav As String = "UserInterface\Deny.wav"

        sMsg = "Docking Request Rejected: "
        Select Case CType(yReason, DockRejectType)
            Case DockRejectType.eDockeeMoving
                sMsg &= "Unable to dock with moving target"
            Case DockRejectType.eDoorSizeExceeded
                sMsg &= "Target's Doors are not big enough!"
            Case DockRejectType.eHangarInoperable
                sMsg &= "Target's Hangar is Inoperable!"
            Case DockRejectType.eInsufficientHangarCap
                sMsg &= "Insufficient Hangar Capacity!"
                sWav = "Game Narrative\LowHangarCap.wav"
            Case DockRejectType.eDockerCannotEnterEnvir
                sMsg &= "Entity cannot enter environment!"
        End Select

        If goSound Is Nothing = False Then goSound.StartSound(sWav, False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)

        If goUILib Is Nothing = False Then
            goUILib.AddNotification(sMsg, System.Drawing.Color.FromArgb(255, 255, 255, 0), lEnvirID, iEnvirTypeID, lObjID, iObjTypeID, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Private Sub HandleProductionCompleteNotice(ByVal yData() As Byte)
        msCurrentLoc = "HandleProductionCompleteNotice"

        Dim lProdID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iProdTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lCreatorID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iCreatorTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, 14)
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 18)

        'New as of V. 19
        Dim yProdType As Byte = yData(20)

        Dim X As Int32
        'Dim Y As Int32

        Dim sCreatedBy As String = ""
        Dim sCreatedWhat As String = ""
        Dim sCreatedWhere As String = ""
        Dim sFinal As String

        Dim sProdRes As String = "Production"

        Dim sWAVFile As String = "Game Narrative\ProductionComplete.wav"

        If goGalaxy Is Nothing Then Return

        'Do not alert the player of repairs
        If iProdTypeID = ObjectType.eRepairItem Then Return

        If NewTutorialManager.TutorialOn = True Then
            If iProdTypeID = ObjectType.eUnitDef Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eUnitBuilt, lProdID, iProdTypeID, -1, "")
            ElseIf iProdTypeID = ObjectType.eFacilityDef Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eFacilityBuilt, lProdID, iProdTypeID, -1, "")
            End If
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eProductionCompleted, lProdID, iProdTypeID, yProdType, "")
        End If

        Select Case iProdTypeID
            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.eMineralTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.eSpecialTech
                'if the production type is research... set it to Research, otherwise, it is PRoduction
                If yProdType = ProductionType.eResearch Then
                    If goCurrentPlayer Is Nothing = False Then
                        Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lProdID, iProdTypeID)
                        If oTech Is Nothing = False Then oTech.Researchers = 0
                    End If

                    sWAVFile = "Game Narrative\ResearchComplete.wav"
                    sProdRes = "Research"

                    If NewTutorialManager.TutorialOn = True Then
                        NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eResearchCompleted, lProdID, iProdTypeID, -1, "")
                    End If

                End If
            Case ObjectType.eUnit, ObjectType.eUnitDef, ObjectType.eRepairItem, ObjectType.eAmmunition, ObjectType.eEnlisted, ObjectType.eOfficers
                sWAVFile = "Game Narrative\ProductionComplete.wav"
                sProdRes = "Production"
            Case ObjectType.eMineral
                'sWAVFile = ""
                Return
            Case Else
                sWAVFile = "Game Narrative\ConstructionComplete.wav"
                sProdRes = "Production"
        End Select

        ' Changes to allow for limiting notifications
        Dim bDoNotification As Boolean = True

        If muSettings.MsgPersonnelProductionComplete = 0 AndAlso (yProdType = ProductionType.eEnlisted) OrElse (yProdType = ProductionType.eOfficers) Then
            bDoNotification = False
        ElseIf muSettings.MsgLandProductionComplete = 0 AndAlso (yProdType = ProductionType.eLandProduction) Then
            bDoNotification = False
        ElseIf muSettings.MsgNavalProductionComplete = 0 AndAlso (yProdType = ProductionType.eNavalProduction) Then
            bDoNotification = False
        ElseIf muSettings.MsgAerialProductionComplete = 0 AndAlso (yProdType = ProductionType.eAerialProduction) Then
            bDoNotification = False
        ElseIf muSettings.MsgRefineryProductionComplete = 0 AndAlso (yProdType = ProductionType.eRefining) Then
            bDoNotification = False
        ElseIf muSettings.MsgCommandCenterProductionComplete = 0 AndAlso (yProdType = ProductionType.eCommandCenterSpecial) Then
            bDoNotification = False
        ElseIf muSettings.MsgSpaceStationProductionComplete = 0 AndAlso (yProdType = ProductionType.eSpaceStationSpecial) Then
            bDoNotification = False
        Else
            bDoNotification = True
        End If

        If goSound Is Nothing = False AndAlso bDoNotification = True Then
            goSound.StartSound(sWAVFile, False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
        End If

        If goUILib Is Nothing Then Return

        If iProdTypeID = ObjectType.eAmmunition Then
            Dim ofrm As frmAmmo = CType(goUILib.GetWindow("frmAmmo"), frmAmmo)
            If ofrm Is Nothing = False Then ofrm.CheckForAmmoProdComplete(lCreatorID, iCreatorTypeID)
            ofrm = Nothing
        End If

        If goCurrentEnvir Is Nothing = False Then
            If goCurrentEnvir.ObjectID = lEnvirID AndAlso goCurrentEnvir.ObjTypeID = iEnvirTypeID Then
                For X = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) = lCreatorID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iCreatorTypeID Then
                        'sCreatedBy = goCurrentEnvir.oEntity(X).EntityName

                        'Check if it has a prodqueue
                        If goCurrentEnvir.oEntity(X).ProductionQueue Is Nothing = False AndAlso goCurrentEnvir.oEntity(X).ProdQueueUB > -1 Then
                            If goCurrentEnvir.oEntity(X).ProductionQueue(0) Is Nothing = False Then
                                goCurrentEnvir.oEntity(X).ProductionQueue(0).iQuantity -= 1S
                                If goCurrentEnvir.oEntity(X).ProductionQueue(0).iQuantity = 0 Then
                                    For lProdIdx As Int32 = 1 To goCurrentEnvir.oEntity(X).ProdQueueUB
                                        goCurrentEnvir.oEntity(X).ProductionQueue(lProdIdx - 1) = goCurrentEnvir.oEntity(X).ProductionQueue(lProdIdx)
                                    Next lProdIdx
                                    goCurrentEnvir.oEntity(X).ProdQueueUB -= 1
                                End If
                            End If
                        End If
                        Exit For
                    End If
                Next X
            End If

            goCurrentEnvir.RemoveUnitProdQueueItem(lCreatorID, iCreatorTypeID, lProdID, iProdTypeID)
        End If
        sCreatedBy = GetCacheObjectValue(lCreatorID, iCreatorTypeID)

        'Cannot trust goCurrentEnvir...
        If iEnvirTypeID = ObjectType.ePlanet Then
            sCreatedWhere = GetCacheObjectValue(lEnvirID, iEnvirTypeID)
            'For X = 0 To goGalaxy.mlSystemUB
            '    For Y = 0 To goGalaxy.moSystems(X).PlanetUB
            '        If goGalaxy.moSystems(X).moPlanets(Y).ObjectID = lEnvirID Then
            '            sCreatedWhere = goGalaxy.moSystems(X).moPlanets(Y).PlanetName '& " (" & goGalaxy.moSystems(X).SystemName & " system)"
            '            Exit For
            '        End If
            '    Next Y
            '    If sCreatedWhere <> "" Then Exit For
            'Next X
        ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
            sCreatedWhere = goGalaxy.GetSystemName(lEnvirID)
        End If
        If sCreatedBy = "Unknown" Then sCreatedBy = ""
        If sCreatedWhere = "Unknown" Then sCreatedWhere = ""

        If iProdTypeID = ObjectType.eUnitDef OrElse iProdTypeID = ObjectType.eFacilityDef OrElse iProdTypeID = ObjectType.eEnlisted OrElse iProdTypeID = ObjectType.eOfficers Then
            For X = 0 To glEntityDefUB
                If glEntityDefIdx(X) = lProdID AndAlso goEntityDefs(X).ObjTypeID = iProdTypeID Then
                    sCreatedWhat = goEntityDefs(X).DefName

                    If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then
                        If iProdTypeID = ObjectType.eFacilityDef Then
                            If goEntityDefs(X).ProductionTypeID = ProductionType.eCommandCenterSpecial Then
                                goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.CommandCenterBuilt)
                            ElseIf goEntityDefs(X).ProductionTypeID = ProductionType.ePowerCenter Then
                                goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.PowerGeneratorBuilt)
                            ElseIf goEntityDefs(X).ProductionTypeID = ProductionType.eEnlisted Then
                                goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.BarracksBuilt)
                            ElseIf goEntityDefs(X).ProductionTypeID = ProductionType.eColonists Then
                                goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.ResidentialFacilityBuilt)
                            ElseIf goEntityDefs(X).ProductionTypeID = ProductionType.eOfficers Then
                                goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.OfficersFacilityBuilt)
                            ElseIf goEntityDefs(X).ProductionTypeID = ProductionType.eLandProduction Then
                                goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.FactoryBuilt)
                            ElseIf goEntityDefs(X).ProductionTypeID = ProductionType.eMining Then
                                goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.MiningFacilityBuilt)
                            ElseIf goEntityDefs(X).ProductionTypeID = ProductionType.eResearch Then
                                goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.ResearchFacilityBuilt)
                            End If
                        Else
                            If goEntityDefs(X).DefName.ToLower = "cargo truck" Then
                                goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.CargoTruckBuilt)
                            End If
                        End If
                    End If

                    Exit For
                End If
            Next X
        ElseIf iProdTypeID = ObjectType.eMineralTech Then
            For X = 0 To glMineralUB
                If glMineralIdx(X) = lProdID Then
                    sCreatedWhat = goMinerals(X).MineralName
                    Exit For
                End If
            Next X
        Else
            Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lProdID, Math.Abs(iProdTypeID))
            If oTech Is Nothing = False Then
                sCreatedWhat = oTech.GetComponentName()
            End If
            oTech = Nothing
        End If

        sFinal = ""
        If sCreatedBy = "" Then
            sFinal = sProdRes & " Completed"
        Else
            sFinal = sCreatedBy & " Finished " & sProdRes
        End If

        If sCreatedWhat <> "" Then
            sFinal &= " on a " & sCreatedWhat
        End If

        If sCreatedWhere <> "" Then
            sFinal &= " in " & sCreatedWhere
        End If

        If iEnvirTypeID <> ObjectType.ePlanet AndAlso iEnvirTypeID <> ObjectType.eSolarSystem Then iEnvirTypeID = -1

        ' Changes to allow for limiting notifications

        If bDoNotification = True Then
            goUILib.AddNotification(sFinal, System.Drawing.Color.FromArgb(255, 255, 255, 255), lEnvirID, iEnvirTypeID, lCreatorID, iCreatorTypeID, Int32.MinValue, Int32.MinValue)
        End If


    End Sub

    Private Sub HandleSetEntityProdFailed(ByVal yData() As Byte)
        If goUILib Is Nothing Then Exit Sub
        msCurrentLoc = "HandleSetEntityProdFailed"
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lProdID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iProdTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)


        Dim lEnvirID As Int32 = -1
        Dim iEnvirTypeID As Int16 = -1
        Dim lLocX As Int32 = -1
        Dim lLocZ As Int32 = -1

        Dim X As Int32
        Dim sMsg As String = "Production could not begin!"

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing = False Then
            lEnvirID = oEnvir.ObjectID
            iEnvirTypeID = oEnvir.ObjTypeID
            Dim lCurUB As Int32 = -1
            If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
            For X = 0 To lCurUB
                If oEnvir.lEntityIdx(X) = lObjID AndAlso oEnvir.oEntity(X).ObjTypeID = iObjTypeID Then
                    lLocX = CInt(oEnvir.oEntity(X).LocX)
                    lLocZ = CInt(oEnvir.oEntity(X).LocZ)
                    oEnvir.oEntity(X).bProducing = False
                    Exit For
                End If
            Next X
        End If
        goCurrentEnvir.RemoveUnitProdQueueItem(lObjID, iObjTypeID, lProdID, iProdTypeID)
        If lProdID = -50 Then
            'special,
            sMsg = "Could not redesign " & Base_Tech.GetComponentTypeName(iProdTypeID) & " because it is currently being researched!"
        Else
            For X = 0 To glEntityDefUB
                If lProdID = glEntityDefIdx(X) Then
                    If goEntityDefs(X).ObjTypeID = iProdTypeID Then
                        sMsg = "Could not begin production of " & goEntityDefs(X).DefName & "!"
                        Exit For
                    End If
                End If
            Next X
        End If

        goUILib.AddNotification(sMsg, System.Drawing.Color.FromArgb(255, 255, 0, 0), lEnvirID, iEnvirTypeID, lObjID, iObjTypeID, lLocX, lLocZ)
        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
    End Sub

    Private Sub HandleBombFireMsg(ByVal yData() As Byte)
        msCurrentLoc = "HandleBombFireMsg"
        Dim lAttackerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iAttackerTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lPlanetID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iPlanetTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
        Dim yType As Byte = yData(14)
        Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, 15)
        Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, 19)
        Dim yAOE As Byte = yData(23)

        If goCurrentEnvir Is Nothing Then Return
        If goWpnMgr Is Nothing Then Return

        'Ok, determine what environment we are currently in
        If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
            'ok, we are in the planet... this one is easy

            'TODO: check if the attacker is in this environment, if they are, then it is not an orbital bombardment

            Dim lDestY As Int32 = 0
            Try
                If goCurrentEnvir.oGeoObject Is Nothing = False Then
                    If CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.ePlanet Then
                        lDestY = CInt(CType(goCurrentEnvir.oGeoObject, Planet).GetHeightAtPoint(lLocX, lLocZ, True))
                    End If
                End If
            Catch
            End Try

            'goWpnMgr.AddNewEffect(lLocX, 10000, lLocZ, lLocX, 0, lLocZ, yType, True, yAOE)

            If yType >= WeaponType.eFlickerGreenBeam AndAlso yType <= WeaponType.eSolidPurpleBeam Then
                goWpnMgr.AddNewEffect(lLocX, lDestY + 10000, lLocZ, lLocX, lDestY, lLocZ, yType, True, yAOE)
            Else
                BombMgr.goBombMgr.AddNewBomb(lLocX + CInt((Rnd() * 3000) - 1500), lDestY + 7000, lLocZ + CInt((Rnd() * 3000) - 1500), lLocX, lDestY, lLocZ, yType, yAOE)
            End If

        ElseIf goCurrentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
            'ok, we are in the system...
            Dim X As Int32
            Dim lAtkrIdx As Int32 = -1

            Select Case yType
                Case WeaponType.eBomb_Green
                    yType = WeaponType.eShortGreenPulse
                Case WeaponType.eBomb_Purple
                    yType = WeaponType.eShortPurplePulse
                Case WeaponType.eBomb_Red
                    yType = WeaponType.eShortRedPulse
                Case WeaponType.eBomb_Teal
                    yType = WeaponType.eShortTealPulse
                Case WeaponType.eBomb_Yellow
                    yType = WeaponType.eShortYellowPulse
                Case WeaponType.eBomb_Gray
                    yType = WeaponType.eShortYellowPulse
            End Select

            For X = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) = lAttackerID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iAttackerTypeID Then
                    lAtkrIdx = X
                    Exit For
                End If
            Next X

            If lAtkrIdx <> -1 AndAlso goCurrentEnvir.oGeoObject Is Nothing = False AndAlso CType(goCurrentEnvir.oGeoObject, Base_GUID).ObjTypeID = ObjectType.eSolarSystem Then
                With CType(goCurrentEnvir.oGeoObject, SolarSystem)
                    For X = 0 To .PlanetUB
                        If .moPlanets(X).ObjectID = lPlanetID Then
                            'ok, found it
                            goWpnMgr.AddNewEffect(goCurrentEnvir.oEntity(lAtkrIdx), .moPlanets(X).LocX, .moPlanets(X).LocY, .moPlanets(X).LocZ, yType, True, yAOE)
                            Exit For
                        End If
                    Next X
                End With
            End If

        End If
    End Sub

    Private Sub HandleChatMsg(ByVal yData() As Byte)
        msCurrentLoc = "HandleChatMsg"

        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim yType As Byte = yData(6)

        Dim lLen As Int32 = System.BitConverter.ToInt32(yData, 7)
        Dim sMsg As String = GetStringFromBytes(yData, 11, lLen)

        If goUILib Is Nothing = False Then
            Dim oTmpWin As frmChat = CType(goUILib.GetWindow("frmChat"), frmChat)

            If oTmpWin Is Nothing = False Then
                oTmpWin.AddChatMessage(lPlayerID, sMsg, CType(yType, ChatMessageType), Now, True)
            End If
            oTmpWin = Nothing
        End If

    End Sub

    Private Sub HandleSkillListMsg(ByVal yData() As Byte)
        Dim lCnt As Int32 = System.BitConverter.ToInt16(yData, 2)
        Dim lPos As Int32
        Dim lID As Int32
        Dim iTypeID As Int16
        Dim X As Int32
        Dim lIdx As Int32

        msCurrentLoc = "HandleSkillListMsg"

        lPos = 4
        While lPos < yData.Length - 1
            'Ok, first the GUID...
            lID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            lIdx = -1
            For X = 0 To glSkillUB
                If goSkills(X).ObjectID = lID AndAlso goSkills(X).ObjTypeID = iTypeID Then
                    'Ok, found it
                    lIdx = X
                    Exit For
                End If
            Next X

            If lIdx = -1 Then
                'Ok, new
                glSkillUB += 1
                ReDim Preserve goSkills(glSkillUB)
                goSkills(glSkillUB) = New Skill()
                lIdx = glSkillUB
            End If

            With goSkills(lIdx)
                .ObjectID = lID
                .ObjTypeID = iTypeID

                'Now, the skill name (20)
                .SkillName = GetStringFromBytes(yData, lPos, 20)
                lPos += 20

                .MinVal = yData(lPos) : lPos += 1
                .MaxVal = yData(lPos) : lPos += 1
                .SkillType = yData(lPos) : lPos += 1

                'Now, skill desc
                .SkillDesc = GetStringFromBytes(yData, lPos, 255)
                lPos += 255
            End With
        End While

        If (mlBusyStatus And elBusyStatusType.SkillsList) <> 0 Then
            mlBusyStatus = mlBusyStatus Xor elBusyStatusType.SkillsList
        End If
    End Sub

    Private Sub HandleGoalListMsg(ByVal yData() As Byte)
        'Dim lCnt As Int32 = System.BitConverter.ToInt16(yData, 2)
        'Dim lPos As Int32
        'Dim lID As Int32
        'Dim iTypeID As Int16
        'Dim X As Int32
        'Dim lIdx As Int32

        msCurrentLoc = "HandleGoalListMsg"

        'lPos = 4
        'While lPos < yData.Length - 1
        '    'Ok, first the GUID...
        '    lID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        '    iTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        '    lIdx = -1
        '    For X = 0 To glGoalUB
        '        If goGoals(X).ObjectID = lID AndAlso goGoals(X).ObjTypeID = iTypeID Then
        '            'Ok, found it
        '            lIdx = X
        '            Exit For
        '        End If
        '    Next X

        '    If lIdx = -1 Then
        '        'Ok, new
        '        glGoalUB += 1
        '        ReDim Preserve goGoals(glGoalUB)
        '        goGoals(glGoalUB) = New Goal()
        '        lIdx = glGoalUB
        '    End If

        '    With goGoals(lIdx)
        '        .ObjectID = lID
        '        .ObjTypeID = iTypeID

        '        'Now, the skill name
        '        .GoalName = GetStringFromBytes(yData, lPos, 20)
        '        lPos += 20

        '        .BaseTime = yData(lPos) : lPos += 1
        '        .MaxTime = yData(lPos) : lPos += 1
        '        .MinTime = yData(lPos) : lPos += 1
        '        .RiskOfDetection = yData(lPos) : lPos += 1
        '        .PointRequirement = yData(lPos) : lPos += 1

        '        .GoalDesc = GetStringFromBytes(yData, lPos, 255)
        '        lPos += 255
        '    End With
        'End While

        If (mlBusyStatus And elBusyStatusType.GoalsList) <> 0 Then
            mlBusyStatus = mlBusyStatus Xor elBusyStatusType.GoalsList
        End If
    End Sub

    Private Sub HandleGetEntityName(ByVal yData() As Byte)
        msCurrentLoc = "HandleGetEntityName"
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim sName As String

        If iTypeID = ObjectType.eGuild Then
            sName = GetStringFromBytes(yData, 8, 50)
        Else
            sName = GetStringFromBytes(yData, 8, 20)
        End If

        'Now, tell our cache to set it
        SetCacheObjectValue(lObjID, iTypeID, sName)
    End Sub

    Private Sub HandleSetEntityPersonnel(ByVal yData() As Byte)
        msCurrentLoc = "HandleSetEntityPersonnel"
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lMaxCol As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim lMaxWrk As Int32 = System.BitConverter.ToInt32(yData, 12)
        Dim lPower As Int32 = System.BitConverter.ToInt32(yData, 16)

        Dim yAutoLaunch As Byte = yData(20)

        Dim X As Int32

        If goCurrentEnvir Is Nothing = False Then
            For X = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) = lID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iTypeID Then
                    With goCurrentEnvir.oEntity(X)
                        .MaxColonists = lMaxCol
                        .MaxWorkers = lMaxWrk
                        .PowerConsumption = lPower

                        If yAutoLaunch = 0 Then
                            .AutoLaunch = False
                        Else : .AutoLaunch = True
                        End If
                    End With
                    Exit For
                End If
            Next X
        End If
    End Sub

    Private Sub HandleColonyLowResources(ByVal yData() As Byte)
        Dim X As Int32
        Dim sColonyName As String
        Dim lColonyID As Int32
        Dim iColonyTypeID As Int16
        Dim lEnvirID As Int32
        Dim iEnvirTypeID As Int16
        Dim yLowRes As Byte

        Dim sWhere As String = ""
        Dim sFinal As String = ""

        msCurrentLoc = "HandleColonyLowResources"

        Dim sWavFile As String = "Game Narrative\LowResources.wav"

        If goUILib Is Nothing Then Exit Sub

        'ColonyName (20)
        sColonyName = GetStringFromBytes(yData, 2, 20)

        lColonyID = System.BitConverter.ToInt32(yData, 22)
        iColonyTypeID = System.BitConverter.ToInt16(yData, 26)
        lEnvirID = System.BitConverter.ToInt32(yData, 28)
        iEnvirTypeID = System.BitConverter.ToInt16(yData, 32)
        yLowRes = yData(34)

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eInsufficientResources, -1, -1, -1, "")
        End If

        'Cannot trust goCurrentEnvir...
        If iEnvirTypeID = ObjectType.ePlanet Then
            For X = 0 To goGalaxy.mlSystemUB
                For Y As Int32 = 0 To goGalaxy.moSystems(X).PlanetUB
                    If goGalaxy.moSystems(X).moPlanets(Y).ObjectID = lEnvirID Then
                        sWhere = goGalaxy.moSystems(X).moPlanets(Y).PlanetName & " (" & goGalaxy.moSystems(X).SystemName & " system)"
                        Exit For
                    End If
                Next Y
                If sWhere <> "" Then Exit For
            Next X
        ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
            sWhere = goGalaxy.GetSystemName(lEnvirID)
        End If

        Select Case yLowRes
            Case ProductionType.eColonists
                sFinal = "Low Colonists"
                sWavFile = "Game Narrative\LowPersonnel.wav"
                'TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.LowPersonnel)
            Case ProductionType.eEnlisted
                sFinal = "Low Enlisted"
                sWavFile = "Game Narrative\LowPersonnel.wav"
                TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.LowEnlisted)
            Case ProductionType.eFood
                sFinal = "Low Food Supplies"
            Case ProductionType.eOfficers
                sFinal = "Low Officers"
                sWavFile = "Game Narrative\LowPersonnel.wav"
                TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.LowOfficers)
            Case ProductionType.ePowerCenter
                sFinal = "Low Power"
                sWavFile = "Game Narrative\LowPower.wav"
                TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.LowPowerEvent)
            Case ProductionType.eCredits
                sFinal = "Low Credits"
                sWavFile = "Game Narrative\LowCredits.wav"
                TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.LowCreditsEvent)
            Case ProductionType.eWareHouse
                sFinal = "Low Cargo Capacity"
                sWavFile = "Game Narrative\LowCapacity.wav"
                'TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.LowCapacity)
            Case ProductionType.eProduction 'indicates low hangar capacity
                sFinal = "Low Hangar Capacity"
                sWavFile = "Game Narrative\LowHangarCap.wav"

                Dim lItemID As Int32 = 0
                Try
                    lItemID = System.BitConverter.ToInt32(yData, 35)
                Catch
                End Try
                If lItemID = -1 Then
                    sFinal = "Low Hangar Door Capacity"
                End If
            Case Else
                sFinal = "Low Resources"
                TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.LowResourcesEvent)

                Dim lItemID As Int32 = System.BitConverter.ToInt32(yData, 35)
                Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 39)
                If iTypeID = ObjectType.eMineral OrElse iTypeID = ObjectType.eMineralCache Then
                    For X = 0 To glMineralUB
                        If glMineralIdx(X) = lItemID Then
                            sFinal = "Low " & goMinerals(X).MineralName
                            Exit For
                        End If
                    Next X
                Else
                    Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lItemID, iTypeID)
                    If oTech Is Nothing = False Then
                        sFinal = "Low " & oTech.GetComponentName()
                    Else
                        Dim sTempVal As String = GetCacheObjectValue(lItemID, iTypeID)
                        If sTempVal.ToUpper <> "UNKNOWN" Then
                            sFinal = "Low " & sTempVal
                        End If
                    End If
                End If

        End Select

        Dim lProdID As Int32 = System.BitConverter.ToInt32(yData, 41)
        Dim iProdTypeID As Int16 = System.BitConverter.ToInt16(yData, 45)

        Dim sBuildingWhat As String = ""
        If iProdTypeID = ObjectType.eUnit OrElse iProdTypeID = ObjectType.eFacility OrElse iProdTypeID = ObjectType.eEnlisted OrElse iProdTypeID = ObjectType.eOfficers Then
            For X = 0 To glEntityDefUB
                If glEntityDefIdx(X) = lProdID AndAlso goEntityDefs(X).ObjTypeID = iProdTypeID Then
                    sBuildingWhat = goEntityDefs(X).DefName
                    Exit For
                End If
            Next X
        ElseIf iProdTypeID = ObjectType.eMineralTech Then
            For X = 0 To glMineralUB
                If glMineralIdx(X) = lProdID Then
                    sBuildingWhat = goMinerals(X).MineralName
                    Exit For
                End If
            Next X
        ElseIf iProdTypeID = ObjectType.eMineral Then
            For X = 0 To glMineralUB
                If glMineralIdx(X) = lProdID Then
                    sBuildingWhat = "bid for " & goMinerals(X).MineralName
                    Exit For
                End If
            Next X
        ElseIf lProdID <> -1 Then
            Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lProdID, Math.Abs(iProdTypeID))
            If oTech Is Nothing = False Then
                sBuildingWhat = oTech.GetComponentName()
            End If
            oTech = Nothing
        End If
        If sBuildingWhat <> "" Then sFinal &= " for " & sBuildingWhat
        If sColonyName <> "" Then sFinal &= " at " & sColonyName & " colony"
        If sWhere = "" Then sFinal &= "!" Else sFinal &= " on " & sWhere

        If goSound Is Nothing = False Then goSound.StartSound(sWavFile, False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)

        goUILib.AddNotification(sFinal, System.Drawing.Color.FromArgb(255, 255, 255, 0), lEnvirID, iEnvirTypeID, -1, -1, Int32.MinValue, Int32.MinValue)
    End Sub

    Private Sub HandleSetEntityStatus(ByVal yData() As Byte)
        msCurrentLoc = "HandleSetEntityStatus"
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lStatus As Int32 = System.BitConverter.ToInt32(yData, 8)
        Try
            If goCurrentEnvir Is Nothing = False Then
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) = lObjID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iObjTypeID Then

                        With goCurrentEnvir.oEntity(X)
                            If lStatus = -elUnitStatus.eFacilityPowered Then
                                If (.CurrentStatus And elUnitStatus.eFacilityPowered) <> 0 Then
                                    .CurrentStatus -= elUnitStatus.eFacilityPowered
                                End If
                            ElseIf lStatus = elUnitStatus.eFacilityPowered Then
                                .CurrentStatus = .CurrentStatus Or elUnitStatus.eFacilityPowered
                            End If

                            'Check if we need to update our stuff now
                            If goUILib Is Nothing = False Then
                                Dim ofrmAdv As frmAdvanceDisplay = CType(goUILib.GetWindow("frmAdvanceDisplay"), frmAdvanceDisplay)
                                If ofrmAdv Is Nothing = False AndAlso ofrmAdv.Visible = True Then
                                    If ofrmAdv.SetEntityStatusMsgRcvd(lObjID, iObjTypeID, lStatus) = False Then
                                        Dim ofrmMine As frmMining = CType(goUILib.GetWindow("frmMining"), frmMining)
                                        If ofrmMine Is Nothing = False Then ofrmMine.SetEntityStatusMsgRcvd(lObjID, iObjTypeID, lStatus)
                                    End If
                                Else
                                    Dim ofrmMine As frmMining = CType(goUILib.GetWindow("frmMining"), frmMining)
                                    If ofrmMine Is Nothing = False Then ofrmMine.SetEntityStatusMsgRcvd(lObjID, iObjTypeID, lStatus)
                                End If
                                ofrmAdv = Nothing
                            End If
                        End With

                        Return
                    End If
                Next X

                'still haven't found it, so see if it is a child
                If iObjTypeID = ObjectType.eFacility Then
                    For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                        If goCurrentEnvir.lEntityIdx(X) <> -1 Then
                            For Y As Int32 = 0 To goCurrentEnvir.oEntity(X).lChildUB
                                If goCurrentEnvir.oEntity(X).lChildIdx(Y) = lObjID AndAlso goCurrentEnvir.oEntity(X).oChild(Y).iChildTypeID = iObjTypeID Then

                                    With goCurrentEnvir.oEntity(X).oChild(Y)
                                        If lStatus = -elUnitStatus.eFacilityPowered Then
                                            If (.lChildStatus And elUnitStatus.eFacilityPowered) <> 0 Then
                                                .lChildStatus -= elUnitStatus.eFacilityPowered
                                            End If
                                        ElseIf lStatus = elUnitStatus.eFacilityPowered Then
                                            .lChildStatus = .lChildStatus Or elUnitStatus.eFacilityPowered
                                        End If

                                        'Check if we need to update our stuff now
                                        If goUILib Is Nothing = False Then
                                            Dim ofrmAdv As frmAdvanceDisplay = CType(goUILib.GetWindow("frmAdvanceDisplay"), frmAdvanceDisplay)
                                            If ofrmAdv Is Nothing = False Then
                                                ofrmAdv.SetEntityStatusMsgRcvd(lObjID, iObjTypeID, lStatus)
                                            End If
                                            ofrmAdv = Nothing
                                        End If
                                    End With

                                    Return
                                End If
                            Next Y
                        End If
                    Next X
                End If
            End If
        Catch
        End Try
    End Sub

    Private Sub HandleValidateVersion(ByVal yData() As Byte, ByVal yConnType As eyConnType)
        Dim lVersion As Int32 = System.BitConverter.ToInt32(yData, 2)
        msCurrentLoc = "HandleValidateVersion"
        mbVersionValidated = (lVersion = gl_CLIENT_VERSION)

        'If (mlBusyStatus And elBusyStatusType.OperatorValidateVersion) <> 0 Then
        '	mlBusyStatus = mlBusyStatus Xor elBusyStatusType.OperatorValidateVersion
        'End If
        'If (mlBusyStatus And elBusyStatusType.PrimaryValidateVersion) <> 0 Then
        '	mlBusyStatus = mlBusyStatus Xor elBusyStatusType.PrimaryValidateVersion
        'End If

        RaiseEvent ValidateVersionResponse(yConnType, mbVersionValidated)
    End Sub

    Private Sub HandleSetEntityProdSucceed(ByVal yData() As Byte)
        msCurrentLoc = "HandleSetEntityProdSucceed"
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lProdID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iProdTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)

        Dim lQty As Int32 = 1
        If yData.Length > 17 AndAlso iObjTypeID = ObjectType.eFacility Then
            lQty = System.BitConverter.ToInt32(yData, 14)
        End If

        Dim sSound As String = "Game Narrative\ProductionStarted.wav"
        If lQty < 1 Then
            sSound = "Game Narrative\ProductionCancelled.wav"
            If goUILib Is Nothing = False Then goUILib.AddNotification("Production Order Cancelled.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        Else
            If iObjTypeID = ObjectType.eUnit Then
                If goSound Is Nothing = False Then
                    sSound = "Game Narrative\ConstructionStarted.wav"
                End If

                If NewTutorialManager.TutorialOn = True Then
                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eConstructionStarted, lProdID, iProdTypeID, -1, "")
                End If

                'Now, tell the primary that we want hte production status
                Dim yResp(7) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eEntityProductionStatus).CopyTo(yResp, 0)
                System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
                System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
                SendToPrimary(yResp)
            ElseIf iObjTypeID = ObjectType.eFacility Then
                Select Case iProdTypeID
                    Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.eMineralTech, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.ePrototype, ObjectType.eSpecialTech
                        If goCurrentEnvir Is Nothing = False Then
                            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                                If goCurrentEnvir.lEntityIdx(X) = lObjID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iObjTypeID Then
                                    If goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eResearch Then
                                        sSound = "Game Narrative\ResearchStarted.wav"

                                        Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lProdID, iProdTypeID)
                                        ' If oTech Is Nothing = False AndAlso oTech.Researchers <> 0 Then oTech.Researchers += CByte(1)
                                        If oTech Is Nothing Then Exit For
                                        If oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eComponentDesign Then oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eComponentResearching
                                    End If
                                    Exit For
                                End If
                            Next X
                        End If
                End Select
            End If
        End If
        If sSound <> "" AndAlso goSound Is Nothing = False Then goSound.StartSound(sSound, False, SoundMgr.SoundUsage.eNarrative, Microsoft.DirectX.Vector3.Empty, Microsoft.DirectX.Vector3.Empty)
    End Sub

    Private Sub HandleGetColonyChildList(ByRef yData() As Byte)
        msCurrentLoc = "HandleGetColonyChildList"

        'msg is..
        Dim lPos As Int32 = 2    'for msg code
        'EnvirID (4)
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'EnvirTypeID (2)
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        'PlayerID (4)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'Cnt (4)
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lFacID As Int32
        Dim iFacTypeID As Int16
        Dim lEntityDefID As Int32
        Dim iEntityDefTypeID As Int16
        Dim lStatus As Int32

        If iEnvirTypeID = ObjectType.eFacility Then
            Dim oParent As BaseEntity = Nothing
            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) = lEnvirID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iEnvirTypeID Then
                    oParent = goCurrentEnvir.oEntity(X)
                    Exit For
                End If
            Next X

            If oParent Is Nothing = False Then

                oParent.lChildUB = -1
                Erase oParent.lChildIdx
                Erase oParent.oChild

                For X As Int32 = 0 To lCnt - 1
                    '   [
                    '       FacilityGUID (6)
                    lFacID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    iFacTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    '       EntityDefID (6)
                    lEntityDefID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    iEntityDefTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    '       UnitStatus (4)
                    lStatus = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    '   ]
                    If lFacID <> lEnvirID OrElse iFacTypeID <> iEnvirTypeID Then
                        Dim lIdx As Int32 = -1
                        Dim lFirstIndex As Int32 = -1
                        For Y As Int32 = 0 To oParent.lChildUB
                            If oParent.lChildIdx(Y) = lFacID Then
                                lIdx = Y
                                Exit For
                            ElseIf lFirstIndex = -1 AndAlso oParent.lChildIdx(Y) = -1 Then
                                lFirstIndex = Y
                            End If
                        Next Y

                        If lIdx = -1 Then
                            If lFirstIndex = -1 Then
                                oParent.lChildUB += 1
                                lIdx = oParent.lChildUB
                                ReDim Preserve oParent.lChildIdx(oParent.lChildUB)
                                ReDim Preserve oParent.oChild(oParent.lChildUB)
                                oParent.oChild(oParent.lChildUB) = New StationChild
                            Else
                                lIdx = lFirstIndex
                                oParent.oChild(lIdx) = New StationChild
                            End If
                        End If

                        oParent.lChildIdx(lIdx) = lFacID
                        With oParent.oChild(lIdx)
                            .lChildID = lFacID
                            .iChildTypeID = iFacTypeID
                            .lChildStatus = lStatus
                            For Y As Int32 = 0 To glEntityDefUB
                                If glEntityDefIdx(Y) = lEntityDefID AndAlso goEntityDefs(Y).ObjTypeID = iEntityDefTypeID Then
                                    .oChildDef = goEntityDefs(Y)
                                    Exit For
                                End If
                            Next Y
                        End With
                    End If
                Next X

                'Finally, update our HangarCapMax and CargoCapMax which is available because we are a facility :D
                oParent.oUnitDef.Hangar_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                oParent.oUnitDef.Cargo_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                oParent.bChildListUpdated = True
            End If
        Else
            'TODO: Define this
        End If



    End Sub

    Private Sub HandleUpdateEntityAttrs(ByRef yData() As Byte)
        msCurrentLoc = "HandleUpdateEntityAttrs"
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim yMan As Byte = yData(8)
        Dim yMaxSpeed As Byte = yData(9)
        Dim yExp_Level As Byte = yData(10)

        If goCurrentEnvir Is Nothing Then Return

        For X As Int32 = 0 To goCurrentEnvir.lEntityUB
            If goCurrentEnvir.lEntityIdx(X) = lObjID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iObjTypeID Then
                goCurrentEnvir.oEntity(X).oUnitDef.Maneuver = yMan
                goCurrentEnvir.oEntity(X).oUnitDef.MaxSpeed = yMaxSpeed

                'TODO: Since this only occurs on a level up, we could indicate as much with some form of 'level up' sfx

                'TODO: Can't enable this untill the base_entity add object contains exp_level.  Otherwise any time I re-enter an environment and get even 1xp it will think we ranked to that rank.
                'TODO: 2009-09-08: Still bugged.  Half of the units seem one exp_level behind.
                'If goCurrentEnvir.oEntity(X).Exp_Level <> yExp_Level Then
                '    Dim sMsg As String = ""

                '    Dim sName As String = GetCacheObjectValue(lObjID, iObjTypeID)
                '    Dim loldXPRank As Int32 = Math.Abs((CInt(goCurrentEnvir.oEntity(X).Exp_Level) - 1) \ 25)
                '    Dim lnewXPRank As Int32 = Math.Abs((CInt(yExp_Level) - 1) \ 25)
                '    If sName = "Unknown" Then sName = "" Else sName &= " "
                '    sMsg = sName & "Gained XP " & goCurrentEnvir.oEntity(X).Exp_Level.ToString & " -> " & yExp_Level.ToString
                '    If loldXPRank <> lnewXPRank Then
                '        Dim sRank As String = BaseEntity.GetExperienceLevelName(yExp_Level)
                '        sMsg &= " Rank increased to " & sRank
                '    End If
                '    goUILib.AddNotification(sMsg, Color.Teal, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                'End If
                goCurrentEnvir.oEntity(X).Exp_Level = yExp_Level

                Exit For
            End If
        Next X
    End Sub

    Private Sub HandleCommandPointUpdate(ByRef yData() As Byte)
        msCurrentLoc = "HandleCommandPointUpdate"
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lCPUsage As Int32 = System.BitConverter.ToInt32(yData, 8)

        If goCurrentEnvir Is Nothing = False AndAlso goCurrentEnvir.ObjectID = lEnvirID AndAlso goCurrentEnvir.ObjTypeID = iEnvirTypeID Then
            goCurrentEnvir.lCPUsage = lCPUsage
        End If
    End Sub

    Private Sub HandleUpdateFleetSpeed(ByRef yData() As Byte)
        msCurrentLoc = "HandleUpdateFleetSpeed"
        Dim lFleetID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim ySpeed As Byte = yData(6)

        For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
            If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
                goCurrentPlayer.moUnitGroups(X).yInterSystemMovementSpeed = ySpeed
                goCurrentPlayer.moUnitGroups(X).LastMessageUpdate = glCurrentCycle
                Exit For
            End If
        Next X
    End Sub

    Private Sub HandleRemoveFromFleet(ByRef yData() As Byte)
        msCurrentLoc = "HandleRemoveFromFleet"
        Dim lFleetID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lUnitID As Int32 = System.BitConverter.ToInt32(yData, 6)

        For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
            If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
                goCurrentPlayer.moUnitGroups(X).RemoveUnit(lUnitID)
                Exit For
            End If
        Next X
    End Sub

    Private Sub HandleSetFleetDest(ByRef yData() As Byte)
        msCurrentLoc = "HandleSetFleetDest"
        Dim lFleetID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lSystemID As Int32 = System.BitConverter.ToInt32(yData, 6)

        If lSystemID = -1 Then
            Dim sFleetName As String = "battlegroup."
            For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
                If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then
                    sFleetName = goCurrentPlayer.moUnitGroups(X).sName & "."
                    Exit For
                End If
            Next X
            goUILib.AddNotification("Unable to set destination for " & sFleetName, Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        Else
            Dim lPos As Int32 = 10
            'Ok, find the fleet
            For X As Int32 = 0 To goCurrentPlayer.mlUnitGroupUB
                If goCurrentPlayer.mlUnitGroupIdx(X) = lFleetID Then

                    'Ok... fill the remaining values
                    With goCurrentPlayer.moUnitGroups(X)
                        .lOriginX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lOriginY = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lOriginZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lInterSystemOriginID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lInterSystemTargetID = lSystemID
                        .lInterSystemMoveCyclesRemaining = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lParentID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .iParentTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                        If .iParentTypeID <> ObjectType.eGalaxy Then
                            goUILib.AddNotification("Destination Set for " & .sName & ".", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            goUILib.AddNotification("Regrouping to move to " & goGalaxy.GetSystemName(.lInterSystemTargetID) & ".", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Dim ofrm As frmFleet = CType(goUILib.GetWindow("frmFleet"), frmFleet)
                            If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
                                ofrm.RefreshFleetList()
                            End If
                            ofrm = Nothing
                        Else
                            .SetAllChildrenToParent()
                            goUILib.AddNotification(.sName & " battlegroup is en route to " & goGalaxy.GetSystemName(.lInterSystemTargetID) & ".", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        End If

                        .LastMessageUpdate = glCurrentCycle
                    End With

                    Exit For
                End If
            Next X
        End If
    End Sub

    Private Sub HandleDeleteEmailItem(ByRef yData() As Byte)
        If goCurrentPlayer Is Nothing Then Return

        msCurrentLoc = "HandleDeleteEmailItem"
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        If iObjTypeID = -1 Then
            For X As Int32 = 0 To goCurrentPlayer.mlEmailFolderUB
                If goCurrentPlayer.mlEmailFolderIdx(X) = lObjID Then
                    goCurrentPlayer.mlEmailFolderIdx(X) = -1
                    goCurrentPlayer.moEmailFolders(X) = Nothing
                    Exit For
                End If
            Next X
        Else
            Dim lPCF_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lTrashIdx As Int32 = -1

            For X As Int32 = 0 To goCurrentPlayer.mlEmailFolderUB
                If goCurrentPlayer.mlEmailFolderIdx(X) = PlayerCommFolder.ePCF_ID_HARDCODES.eDeleted_PCF Then
                    lTrashIdx = X
                    Exit For
                End If
            Next X

            For X As Int32 = 0 To goCurrentPlayer.mlEmailFolderUB
                If goCurrentPlayer.mlEmailFolderIdx(X) = lPCF_ID Then
                    With goCurrentPlayer.moEmailFolders(X)
                        For Y As Int32 = 0 To .PlayerMsgUB
                            If .PlayerMsgsIdx(Y) <> -1 AndAlso (lObjID = -1 OrElse .PlayerMsgsIdx(Y) = lObjID) Then
                                If lTrashIdx <> -1 AndAlso lTrashIdx <> X Then
                                    Dim oFolder As PlayerCommFolder = goCurrentPlayer.moEmailFolders(lTrashIdx)
                                    Dim lMsgIdx As Int32 = -1
                                    For Z As Int32 = 0 To oFolder.PlayerMsgUB
                                        If oFolder.PlayerMsgsIdx(Z) = -1 Then
                                            lMsgIdx = Z
                                            Exit For
                                        End If
                                    Next Z

                                    If lMsgIdx = -1 Then
                                        oFolder.PlayerMsgUB += 1
                                        ReDim Preserve oFolder.PlayerMsgs(oFolder.PlayerMsgUB)
                                        ReDim Preserve oFolder.PlayerMsgsIdx(oFolder.PlayerMsgUB)
                                        lMsgIdx = oFolder.PlayerMsgUB
                                    End If
                                    oFolder.PlayerMsgsIdx(lMsgIdx) = .PlayerMsgsIdx(Y)
                                    oFolder.PlayerMsgs(lMsgIdx) = .PlayerMsgs(Y)
                                End If

                                .PlayerMsgsIdx(Y) = -1
                                .PlayerMsgs(Y) = Nothing
                                If lObjID <> -1 Then Exit For
                            End If
                        Next Y
                    End With
                    Exit For
                End If
            Next X
        End If

    End Sub

    Private Sub HandleMoveEmailToFolder(ByRef yData() As Byte)
        msCurrentLoc = "HandleMoveEmailtoFolder"
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lPCF_Src As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPCF_Dest As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If goCurrentPlayer Is Nothing Then Return

        Dim lFromIdx As Int32 = -1
        Dim lToIdx As Int32 = -1

        For X As Int32 = 0 To goCurrentPlayer.mlEmailFolderUB
            If goCurrentPlayer.mlEmailFolderIdx(X) = lPCF_Src Then
                lFromIdx = X
            ElseIf goCurrentPlayer.mlEmailFolderIdx(X) = lPCF_Dest Then
                lToIdx = X
            End If
        Next X

        If lFromIdx = -1 OrElse lToIdx = -1 Then Return

        Dim oMsg As PlayerComm = Nothing
        With goCurrentPlayer.moEmailFolders(lFromIdx)
            For X As Int32 = 0 To .PlayerMsgUB
                If .PlayerMsgsIdx(X) = lObjID Then
                    oMsg = .PlayerMsgs(X)
                    .PlayerMsgs(X) = Nothing
                    .PlayerMsgsIdx(X) = -1
                    Exit For
                End If
            Next
        End With

        If oMsg Is Nothing = False Then
            oMsg.PCF_ID = goCurrentPlayer.mlEmailFolderIdx(lToIdx)
            goCurrentPlayer.AddPlayerComm(oMsg)
        End If
    End Sub

    Private Sub HandleGetColonyList(ByRef yData() As Byte)
        'MsgCode (2), PlayerID (4), Cnt (4)
        'ColonyID (4), Name (20), ParentGUID (6) ...
        msCurrentLoc = "HandleGetColonyList"
        Dim lPos As Int32 = 2   'for 2 byte msg code
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If glPlayerID = lPlayerID Then
            If goCurrentPlayer Is Nothing = False Then
                Dim bUsed() As Boolean
                ReDim bUsed(goCurrentPlayer.mlColonyUB)

                For X As Int32 = 0 To lCnt - 1
                    Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                    Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                    Dim lIdx As Int32 = -1
                    Dim lFirstIdx As Int32 = -1
                    With goCurrentPlayer
                        For Y As Int32 = 0 To .mlColonyUB
                            If .mlColonyIdx(Y) = lID Then
                                lIdx = Y
                                Exit For
                            ElseIf lFirstIdx = -1 AndAlso .mlColonyIdx(Y) = -1 Then
                                lFirstIdx = Y
                            End If
                        Next
                        If lIdx = -1 Then
                            If lFirstIdx = -1 Then
                                .mlColonyUB += 1
                                ReDim Preserve .mlColonyIdx(.mlColonyUB)
                                ReDim Preserve .moColonies(.mlColonyUB)
                                lIdx = .mlColonyUB
                            Else : lIdx = lFirstIdx
                            End If
                            .moColonies(lIdx) = New Colony()
                        Else : bUsed(lIdx) = True
                        End If
                    End With
                    With goCurrentPlayer.moColonies(lIdx)
                        .ObjectID = lID
                        .ObjTypeID = ObjectType.eColony
                        .ParentEnvirID = lParentID
                        .ParentEnvirTypeID = iParentTypeID
                        .sName = sName
                        .sParentName = GetCacheObjectValue(.ParentEnvirID, .ParentEnvirTypeID)
                    End With
                    goCurrentPlayer.mlColonyIdx(lIdx) = lID
                Next X

                For X As Int32 = 0 To bUsed.GetUpperBound(0)
                    If bUsed(X) = False Then
                        goCurrentPlayer.mlColonyIdx(X) = -1
                    End If
                Next X
            End If
        End If

    End Sub

    Private Sub HandleGetNonOwnerDetails(ByRef yData() As Byte)
        msCurrentLoc = "HandleGetNonOwnerDetails"
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim oSB As New System.Text.StringBuilder()

        Dim lFinalObjID As Int32 = lObjID
        Dim iFinalTypeID As Int16 = iObjTypeID
        If iObjTypeID = ObjectType.eComponentCache Then
            lPos += 2   'for new msgcode
            lObjID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        End If

        Select Case Math.Abs(iObjTypeID)
            Case ObjectType.eAgent
                oSB.AppendLine("AGENT")
                oSB.AppendLine("Name: " & GetStringFromBytes(yData, lPos, 30)) : lPos += 30
                oSB.AppendLine(vbCrLf & "Infiltration: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Dagger: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Resourcefulness: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Suspicion: " & yData(lPos)) : lPos += 1

                Dim lUpfront As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lMaint As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                oSB.AppendLine(vbCrLf & "SKILLS")
                Dim lSkillCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                For X As Int32 = 0 To lSkillCnt - 1
                    Dim lSkillID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim yProf As Byte = yData(lPos) : lPos += 1

                    For Z As Int32 = 0 To glSkillUB
                        If glSkillIdx(Z) = lSkillID Then
                            oSB.AppendLine(goSkills(Z).SkillName.PadRight(21, " "c) & yProf & "/" & goSkills(Z).MaxVal)
                            Exit For
                        End If
                    Next Z
                Next X

                oSB.AppendLine(vbCrLf & "Original Cost: " & lUpfront.ToString("#,##0"))
                oSB.AppendLine("Maintenance: " & lMaint.ToString("#,##0"))
            Case ObjectType.eUnit, ObjectType.eUnitDef, ObjectType.eFacility, ObjectType.eFacilityDef
                'Def, but the nonOwner array has Unit/facility
                'iObjTypeID = ObjectType.eUnit
                Dim oTmpDef As New EntityDef
                Dim lWorkerFactor As Int32 = 0

                With oTmpDef
                    lPos += 4       'for ownerid
                    For X As Int32 = 0 To 3
                        .Armor_MaxHP(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Next X
                    .Maneuver = yData(lPos) : lPos += 1
                    .MaxSpeed = yData(lPos) : lPos += 1

                    '.FuelEfficiency = yData(lPos) : lPos += 1
                    lPos += 1       'old fuel efficiency

                    .Structure_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .Hangar_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .HullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .Cargo_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .Shield_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .ShieldRecharge = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .ShieldRechargeFreq = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                    '.Fuel_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    lPos += 4       'old fuel_cap

                    .Weapon_Acc = yData(lPos) : lPos += 1
                    .ScanResolution = yData(lPos) : lPos += 1
                    .OptRadarRange = yData(lPos) : lPos += 1
                    .MaxRadarRange = yData(lPos) : lPos += 1
                    .DisruptionResistance = yData(lPos) : lPos += 1
                    .PiercingResist = yData(lPos) : lPos += 1
                    .ImpactResist = yData(lPos) : lPos += 1
                    .BeamResist = yData(lPos) : lPos += 1
                    .ECMResist = yData(lPos) : lPos += 1
                    .FlameResist = yData(lPos) : lPos += 1
                    .ChemicalResist = yData(lPos) : lPos += 1
                    lPos += 1   'detect res
                    .ModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .DefName = GetStringFromBytes(yData, lPos, 20) : lPos += 20

                    If iObjTypeID = ObjectType.eFacility OrElse iObjTypeID = ObjectType.eFacilityDef Then
                        lWorkerFactor = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        lPos += 1       'max fac size
                        .ProdFactor = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .PowerFactor = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    End If

                    .ProductionTypeID = yData(lPos) : lPos += 1
                    .RequiredProductionTypeID = yData(lPos) : lPos += 1
                    .yChassisType = yData(lPos) : lPos += 1
                    .yFXColors = yData(lPos) : lPos += 1
                    .yArmorIntegrityRoll = yData(lPos) : lPos += 1
                    .JamImmunity = yData(lPos) : lPos += 1
                    .JamStrength = yData(lPos) : lPos += 1
                    .JamTargets = yData(lPos) : lPos += 1
                    .JamEffect = yData(lPos) : lPos += 1

                    lPos += 16  'crit locs

                    'the rest of the data is the weapon-specific data, need to find that out
                    .WeaponDefUB = System.BitConverter.ToInt16(yData, lPos) - 1 : lPos += 2
                    ReDim .WeaponDefs(.WeaponDefUB)
                    For X As Int32 = 0 To .WeaponDefUB
                        .WeaponDefs(X) = New WeaponDef
                        With .WeaponDefs(X)
                            lPos = .FillFromMsg(yData, lPos)
                        End With
                    Next X

                    '================================
                    'Now, create our string
                    oSB.AppendLine(.DefName)
                    If iObjTypeID = ObjectType.eFacility OrElse iObjTypeID = ObjectType.eFacilityDef Then
                        oSB.AppendLine("Facility")
                    Else : oSB.AppendLine("Unit")
                    End If
                    If (.yChassisType And ChassisType.eGroundBased) <> 0 Then
                        oSB.AppendLine("Ground-Based")
                    ElseIf (.yChassisType And ChassisType.eAtmospheric) <> 0 AndAlso (.yChassisType And ChassisType.eSpaceBased) <> 0 Then
                        oSB.AppendLine("Aerospace")
                    ElseIf (.yChassisType And ChassisType.eAtmospheric) <> 0 Then
                        oSB.AppendLine("Atmospheric")
                    ElseIf (.yChassisType And ChassisType.eNavalBased) <> 0 Then
                        oSB.AppendLine("Naval")
                    Else
                        oSB.AppendLine("Space-Only")
                    End If

                    oSB.AppendLine()
                    oSB.AppendLine("Basic Attributes")
                    oSB.AppendLine("  Hull Size: " & .HullSize)

                    oSB.AppendLine("Defenses")
                    If .Shield_MaxHP <> 0 Then oSB.AppendLine("  Shield Strength: " & .Shield_MaxHP)
                    If .ShieldRecharge <> 0 Then oSB.AppendLine("  Recharge Rate: " & .ShieldRecharge)
                    If .ShieldRechargeFreq <> 0 Then oSB.AppendLine("    " & (.ShieldRecharge / (.ShieldRechargeFreq / 30)).ToString("0.#") & " per second")

                    oSB.AppendLine("  Armor Resist")
                    oSB.AppendLine("    Beam/Energy Resist: " & (100 - .BeamResist))
                    oSB.AppendLine("    Burn Resist: " & (100 - .FlameResist))
                    oSB.AppendLine("    Chemical Resist: " & (100 - .ChemicalResist))
                    oSB.AppendLine("    Impact Resist: " & (100 - .ImpactResist))
                    oSB.AppendLine("    Magnetic Resist: " & (100 - .ECMResist))
                    oSB.AppendLine("    Piercing Resist: " & (100 - .PiercingResist))
                    oSB.AppendLine("    Integrity: " & .lDisplayedIntegrity)
                    oSB.AppendLine("  Armor Plating")
                    oSB.AppendLine("    Forward: " & .Armor_MaxHP(UnitArcs.eForwardArc).ToString("0.#") & " hps")
                    oSB.AppendLine("    Left: " & .Armor_MaxHP(UnitArcs.eLeftArc).ToString("0.#") & " hps")
                    oSB.AppendLine("    Rear: " & .Armor_MaxHP(UnitArcs.eBackArc).ToString("0.#") & " hps")
                    oSB.AppendLine("    Right: " & .Armor_MaxHP(UnitArcs.eRightArc).ToString("0.#") & " hps")
                    oSB.AppendLine("  Structural Integrity " & .Structure_MaxHP & " hps")

                    oSB.AppendLine("Electronics")
                    If .Weapon_Acc <> 0 Then oSB.AppendLine("  Targeting Accuracy: " & .Weapon_Acc)
                    'Scan Resolution?
                    If .OptRadarRange <> 0 Then oSB.AppendLine("  Optimum Radar Range: " & .OptRadarRange)
                    If .MaxRadarRange <> 0 Then oSB.AppendLine("  Maximum Radar Range: " & .MaxRadarRange)
                    'Disruption Resist?

                    oSB.AppendLine("Offensive Capabilities")
                    If .WeaponDefUB = -1 Then
                        oSB.AppendLine("  None equipped")
                    Else
                        For X As Int32 = 0 To .WeaponDefUB
                            oSB.AppendLine("  " & .WeaponDefs(X).WeaponName)
                            Select Case .WeaponDefs(X).ArcID
                                Case UnitArcs.eForwardArc
                                    oSB.AppendLine("    Front arc")
                                Case UnitArcs.eLeftArc
                                    oSB.AppendLine("    Left arc")
                                Case UnitArcs.eRightArc
                                    oSB.AppendLine("    Right arc")
                                Case UnitArcs.eBackArc
                                    oSB.AppendLine("    Rear arc")
                                Case Else
                                    oSB.AppendLine("    All arcs")
                            End Select
                            oSB.AppendLine("    ROF: " & (.WeaponDefs(X).ROF_Delay / 30.0F).ToString("#,###.#0"))

                            Dim lMinDmg As Int32 = .WeaponDefs(X).BeamMinDmg + .WeaponDefs(X).ChemicalMinDmg + .WeaponDefs(X).ECMMinDmg + .WeaponDefs(X).FlameMinDmg + .WeaponDefs(X).ImpactMinDmg + .WeaponDefs(X).PiercingMinDmg
                            Dim lMaxDmg As Int32 = .WeaponDefs(X).BeamMaxDmg + .WeaponDefs(X).ChemicalMaxDmg + .WeaponDefs(X).ECMMaxDmg + .WeaponDefs(X).FlameMaxDmg + .WeaponDefs(X).ImpactMaxDmg + .WeaponDefs(X).PiercingMaxDmg
                            oSB.AppendLine("    Min Dmg: " & lMinDmg)
                            oSB.AppendLine("    Max Dmg: " & lMaxDmg)
                        Next X
                    End If

                    oSB.AppendLine("Mobility")

                    If .Maneuver <> 0 Then oSB.AppendLine("  Maneuver: " & .Maneuver)
                    If .MaxSpeed <> 0 Then oSB.AppendLine("  Max Speed: " & .MaxSpeed)
                    If .Maneuver = 0 AndAlso .MaxSpeed = 0 Then oSB.AppendLine("  No Movement Ability")

                    'If .Fuel_Cap <> 0 Then oSB.AppendLine("  Fuel Capacity: " & .Fuel_Cap)

                    oSB.AppendLine("Utility")

                    If .Hangar_Cap <> 0 Then
                        oSB.AppendLine("  Hangar Capacity: " & .Hangar_Cap)
                        'TODO: Need Doors...
                    End If
                    If .Cargo_Cap <> 0 Then oSB.AppendLine("  Cargo Capacity: " & .Cargo_Cap)

                    If .ProdFactor <> 0 AndAlso (iObjTypeID = ObjectType.eFacility OrElse iObjTypeID = ObjectType.eFacilityDef) Then
                        oSB.AppendLine("Production")
                        oSB.AppendLine("  Jobs: " & lWorkerFactor)
                        oSB.AppendLine("  Production Rating: " & .ProdFactor)

                        oSB.AppendLine(vbCrLf & "  Production Ability")

                        Select Case CType(.ProductionTypeID, ProductionType)
                            Case ProductionType.eAerialProduction
                                oSB.AppendLine("    Air/Space Production")
                            Case ProductionType.eColonists
                                oSB.AppendLine("    Residence")
                            Case ProductionType.eCommandCenterSpecial
                                oSB.AppendLine("    Command Center")
                            Case ProductionType.eCredits
                                oSB.AppendLine("    Financials")
                            Case ProductionType.eEnlisted
                                oSB.AppendLine("    Enlisted")
                            Case ProductionType.eEnlistedAndOfficers
                                oSB.AppendLine("    Enlisted and Officers")
                            Case ProductionType.eFacility
                                oSB.AppendLine("    Facilities")
                            Case ProductionType.eFood
                                oSB.AppendLine("    Food")
                            Case ProductionType.eLandProduction
                                oSB.AppendLine("    Ground Production")
                            Case ProductionType.eMining
                                oSB.AppendLine("    Mining")
                            Case ProductionType.eMorale
                                oSB.AppendLine("    Morale")
                            Case ProductionType.eNavalProduction
                                oSB.AppendLine("    Naval Production")
                            Case ProductionType.eOfficers
                                oSB.AppendLine("    Officers")
                            Case ProductionType.ePowerCenter
                                oSB.AppendLine("    Power")
                            Case ProductionType.eRefining
                                oSB.AppendLine("    Refining")
                            Case ProductionType.eResearch
                                oSB.AppendLine("    Research")
                            Case ProductionType.eSpaceStationSpecial
                                oSB.AppendLine("    Space Station")
                            Case ProductionType.eWareHouse
                                oSB.AppendLine("    Warehouse")
                        End Select
                    End If
                End With

                'Finally, change iTypeID to the Unit/Facility as that is what is in the NonOwner list
                If iObjTypeID = ObjectType.eUnitDef Then
                    iObjTypeID = ObjectType.eUnit
                ElseIf iObjTypeID = ObjectType.eFacilityDef Then
                    iObjTypeID = ObjectType.eFacility
                End If
            Case ObjectType.eMineralCache, ObjectType.eMineral
                Dim iCnt As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                oSB.AppendLine("Mineral Properties")
                For X As Int32 = 0 To iCnt - 1
                    Dim lProp As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim yValue As Byte = yData(lPos) : lPos += 1

                    Dim sProperty As String = "Unknown Property"
                    For Y As Int32 = 0 To glMineralPropertyUB
                        If lProp = glMineralPropertyIdx(Y) Then
                            sProperty = goMineralProperty(Y).MineralPropertyName
                            Exit For
                        End If
                    Next Y

                    oSB.AppendLine("  " & sProperty & ": " & yValue)
                Next X
            Case ObjectType.eAlloyTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eSpecialTech
                Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                oSB.AppendLine(sName)
                oSB.AppendLine(Base_Tech.GetComponentTypeName(iObjTypeID))
            Case ObjectType.eArmorTech
                oSB.AppendLine(GetStringFromBytes(yData, lPos, 20)) : lPos += 20
                oSB.AppendLine(Base_Tech.GetComponentTypeName(iObjTypeID))
                oSB.AppendLine("Resistances")
                oSB.AppendLine("-----------")
                oSB.AppendLine("Piercing: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Impact: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Beam: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Magnetic: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Flame: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Chemical: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Radar: " & yData(lPos)) : lPos += 1
                oSB.AppendLine()
                oSB.AppendLine("Hit Points Per Plate: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                oSB.AppendLine("Hull Per Plate: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                oSB.AppendLine("Integrity: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
            Case ObjectType.eEngineTech
                oSB.AppendLine(GetStringFromBytes(yData, lPos, 20)) : lPos += 20
                oSB.AppendLine(Base_Tech.GetComponentTypeName(iObjTypeID))
                oSB.AppendLine("Thrust: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                oSB.AppendLine("Maneuver: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Speed: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Power Production: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                oSB.AppendLine("Hull Required: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                oSB.AppendLine("Intended for a " & Base_Tech.GetHullTypeName(yData(lPos))) : lPos += 1
            Case ObjectType.eRadarTech
                oSB.AppendLine(GetStringFromBytes(yData, lPos, 20)) : lPos += 20
                oSB.AppendLine(Base_Tech.GetComponentTypeName(iObjTypeID))
                oSB.AppendLine("Weapon Accuracy: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Scan Resolution: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Optimum Range: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Maximum Range: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Disruption Resist: " & yData(lPos)) : lPos += 1
                oSB.AppendLine("Power Required: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                oSB.AppendLine("Hull Required: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
            Case ObjectType.eShieldTech
                oSB.AppendLine(GetStringFromBytes(yData, lPos, 20)) : lPos += 20
                oSB.AppendLine(Base_Tech.GetComponentTypeName(iObjTypeID))
                oSB.AppendLine("Max Hit Points: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                oSB.AppendLine("Recharge Rate: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                oSB.AppendLine("Recharge Frequency: " & (System.BitConverter.ToInt32(yData, lPos) / 30.0F).ToString("#,##0.##")) : lPos += 4
                oSB.AppendLine("Projection Hull Size: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                oSB.AppendLine("Power Required: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                oSB.AppendLine("Hull Required: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                oSB.AppendLine("Intended for a " & Base_Tech.GetHullTypeName(yData(lPos))) : lPos += 1
            Case ObjectType.eWeaponTech
                oSB.AppendLine(GetStringFromBytes(yData, lPos, 20)) : lPos += 20
                'oSB.AppendLine(Base_Tech.GetComponentTypeName(iObjTypeID))
                Dim yClassTypeID As Byte = yData(lPos) : lPos += 1
                Dim yWpnTypeID As Byte = yData(lPos) : lPos += 1
                Dim lPower As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lHull As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim yHullTypeID As Byte = yData(lPos) : lPos += 1
                oSB.AppendLine("Weapon for a " & Base_Tech.GetHullTypeName(yHullTypeID))
                Select Case yClassTypeID
                    Case WeaponClassType.eEnergyBeam
                        oSB.AppendLine("Energy Beam")
                        Dim lROF As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        oSB.AppendLine("ROF: " & (lROF / 30.0F).ToString("#,###.#0"))
                        oSB.AppendLine("Maximum Damage: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                        oSB.AppendLine("Optimum Range: " & System.BitConverter.ToInt16(yData, lPos)) : lPos += 1
                        oSB.AppendLine("Accuracy: " & yData(lPos)) : lPos += 1
                    Case WeaponClassType.eMissile
                        oSB.AppendLine("Missile System")
                        Dim lROF As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        oSB.AppendLine("ROF: " & (lROF / 30.0F).ToString("#,###.#0"))
                        oSB.AppendLine("Missile Hull Size: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                        oSB.AppendLine("Maximum Damage: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                        oSB.AppendLine("Homing Accuracy: " & yData(lPos)) : lPos += 1
                        oSB.AppendLine("Max Speed: " & yData(lPos)) : lPos += 1
                        oSB.AppendLine("Maneuver: " & yData(lPos)) : lPos += 1
                        oSB.AppendLine("Explosion Radius: " & yData(lPos)) : lPos += 1
                        lROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        oSB.AppendLine("Flight Time: " & (lROF / 30.0F).ToString("#,###.#0"))
                    Case WeaponClassType.eProjectile
                        oSB.AppendLine("Projectile Weapon")
                        Dim lROF As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        oSB.AppendLine("ROF: " & (lROF / 30.0F).ToString("#,###.#0"))
                        oSB.AppendLine("Cartridge Size: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                        oSB.AppendLine("Max Range: " & System.BitConverter.ToInt16(yData, lPos)) : lPos += 2
                        oSB.AppendLine("Explosion Radius: " & yData(lPos)) : lPos += 1
                        oSB.AppendLine("Pierce Ratio: " & yData(lPos)) : lPos += 1
                    Case WeaponClassType.eEnergyPulse
                        oSB.AppendLine("Energy Pulse Weapon")
                        Dim lROF As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        oSB.AppendLine("ROF: " & (lROF / 30.0F).ToString("#,###.#0"))
                        oSB.AppendLine("Scatter Radius: " & yData(lPos)) : lPos += 1
                        oSB.AppendLine("Max Range: " & yData(lPos)) : lPos += 1
                    Case WeaponClassType.eBomb
                        Dim lROF As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        oSB.AppendLine("ROF: " & (lROF / 30.0F).ToString("#,###.#0"))
                        oSB.AppendLine("Payload Size: " & System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
                        oSB.AppendLine("Range: " & yData(lPos).ToString) : lPos += 1
                        oSB.AppendLine("AOE: " & yData(lPos).ToString) : lPos += 1
                        oSB.AppendLine("Guidance: " & yData(lPos).ToString) : lPos += 1
                        Select Case yData(lPos)
                            Case 0
                                oSB.AppendLine("Explosive Payload")
                            Case 1
                                oSB.AppendLine("Toxic Payload")
                            Case 2
                                oSB.AppendLine("Concussion Payload")
                            Case 3
                                oSB.AppendLine("EMP Payload")
                        End Select
                        lPos += 1
                    Case Else
                        oSB.AppendLine("Need to be defined")
                End Select
                oSB.AppendLine("Power Required: " & lPower)
                oSB.AppendLine("Hull Required: " & lHull)
            Case Else
                oSB.AppendLine("Unable to request data.")
        End Select

        SetNonOwnerItemData(lFinalObjID, iFinalTypeID, oSB.ToString)
    End Sub

    Private Sub HandleDeathBudgetDeposit(ByRef yData() As Byte)
        msCurrentLoc = "HandleDeathBudgetDeposit"
        Dim lVal As Int32 = System.BitConverter.ToInt32(yData, 2)
        If lVal < 0 Then
            If goUILib Is Nothing = False Then goUILib.AddNotification("Unable to deposit funds into the Death Budget account. The account may be full.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        Else
            If goCurrentPlayer Is Nothing = False Then goCurrentPlayer.DeathBudgetBalance = lVal
            If goUILib Is Nothing = False Then goUILib.AddNotification("New Death Budget balance: " & lVal.ToString("#,##0"), Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Private Sub HandlePlayerIsDead(ByVal yData() As Byte)
        msCurrentLoc = "HandlePlayerIsDead"
        Dim yValue As Byte = yData(2)
        If goUILib Is Nothing Then Return

        If yValue = 1 Then
            If frmDeath.mbInDeathSequence = False Then
                Dim frmDeath As frmDeath = CType(goUILib.GetWindow("frmDeath"), frmDeath)
                If frmDeath Is Nothing Then frmDeath = New frmDeath(goUILib)
                frmDeath.Visible = True
                frmDeath = Nothing
            End If
        ElseIf yValue = 2 Then
            'Ok, player was dead but is now reborn... relog the player

            If goCurrentPlayer Is Nothing = False Then
                If goCurrentPlayer.yPlayerPhase = 0 Then
                    frmMain.bTutRelog = True
                    goCurrentPlayer.yPlayerPhase = 1
                    goCurrentPlayer.lTutorialStep = 300
                    NewTutorialManager.ClearSteps()
                    Return
                End If
            End If

            frmMain.bRestartWithUpdater = True
            'frmMain.mbfrmConfirmHandled = True
            'Try
            '    frmMain.Close()
            'Catch
            'End Try

            'Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
            'If sPath.EndsWith("\") = False Then sPath &= "\"
            'sPath &= "UpdaterClient.exe"
            'If Exists(sPath) = True Then Shell(sPath, AppWinStyle.NormalFocus, False, -1)
            'Application.Exit()
            'End


            'Dim lPos As Int32 = 3
            'Dim lStartedEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'Dim iStartedEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            'Dim lGalaxyID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'Dim lSystemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'Dim lPlanetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            ''TODO: For now, always same galaxy

            'For X As Int32 = 0 To goGalaxy.mlSystemUB
            '	If goGalaxy.moSystems(X).ObjectID = lSystemID Then
            '		If goGalaxy.moSystems(X).PlanetUB = -1 Then
            '			RequestSystemDetails(lSystemID)
            '		Else
            '			Dim bGood As Boolean = False
            '			For Y As Int32 = 0 To goGalaxy.moSystems(X).PlanetUB
            '				If goGalaxy.moSystems(X).moPlanets(Y).ObjectID = lPlanetID Then
            '					bGood = True
            '					Exit For
            '				End If
            '			Next Y

            '			If bGood = False Then
            '				RequestSystemDetails(lSystemID)
            '			End If
            '		End If
            '		Exit For
            '	End If
            'Next X

            'frmMain.ForceChangeEnvironment(lStartedEnvirID, iStartedEnvirTypeID)

        Else '0
            goUILib.RemoveWindow("frmDeath")
        End If
    End Sub

    Private Sub HandleSetMaintenanceTarget(ByRef yData() As Byte, ByVal iMsgCode As Int16)
        msCurrentLoc = "HandleSetMaintenanceTarget"
        Dim lPos As Int32 = 2
        Dim lObjectID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        'Only time I get this is to notify me that it cannot be done
        If goUILib Is Nothing = False Then
            If iMsgCode = GlobalMessageCode.eSetRepairTarget Then
                If iObjTypeID = ObjectType.eUnit AndAlso iTargetTypeID = ObjectType.eFacility Then
                    goUILib.AddNotification("Units cannot repair facilities.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Else
                    goUILib.AddNotification("Unable to execute Repair Order: Possibly Multiple targets at location.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
            Else
                goUILib.AddNotification("Unable to execute Dismantle Order: Possibly Multiple targets at location.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
        End If

        If goSound Is Nothing = False Then
            goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
        End If

    End Sub

    Private Sub HandleDeleteDesign(ByRef yData() As Byte)
        msCurrentLoc = "HandleDeleteDesign"
        Dim lTechID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        If goCurrentPlayer Is Nothing = False Then
            goCurrentPlayer.RemoveTech(lTechID, iTypeID)

            If goUILib Is Nothing = False Then
                Dim ofrm As frmResearchMain = CType(goUILib.GetWindow("frmResearchMain"), frmResearchMain)
                If ofrm Is Nothing = False Then ofrm.RefreshComponentList()
                ofrm = Nothing
            End If
        End If
    End Sub

    Private Sub HandleMissileFired(ByRef yData() As Byte)
        msCurrentLoc = "HandleMissileFired"

        Dim lPos As Int32 = 2
        Dim lMissileID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim bHit As Boolean = True
        If lMissileID < 0 Then
            bHit = False
            lMissileID = Math.Abs(lMissileID)
        End If
        Dim lAttackerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iAttackerTypeID As Int16 = ObjectType.eUnit
        If lAttackerID < 0 Then
            lAttackerID = Math.Abs(lAttackerID)
            iAttackerTypeID = ObjectType.eFacility
        End If
        Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTargetTypeID As Int16 = ObjectType.eUnit
        If lTargetID < 0 Then
            lTargetID = Math.Abs(lTargetID)
            iTargetTypeID = ObjectType.eFacility
        End If
        Dim yMissileType As Byte = yData(lPos) : lPos += 1
        Dim ySpeed As Byte = yData(lPos) : lPos += 1
        Dim yManeuver As Byte = CByte(ySpeed \ 2)
        If yManeuver = 0 Then yManeuver = 1

        'goUILib.AddNotification("Missile " & lMissileID & " fired.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If goMissileMgr Is Nothing = False AndAlso oEnvir Is Nothing = False Then
            Dim oTarget As BaseEntity = Nothing
            Dim oAttacker As BaseEntity = Nothing

            For X As Int32 = 0 To oEnvir.lEntityUB
                If oEnvir.lEntityIdx(X) = lAttackerID AndAlso oEnvir.oEntity(X).ObjTypeID = iAttackerTypeID Then
                    oAttacker = oEnvir.oEntity(X)
                    Exit For
                End If
            Next X
            For X As Int32 = 0 To oEnvir.lEntityUB
                If oEnvir.lEntityIdx(X) = lTargetID AndAlso oEnvir.oEntity(X).ObjTypeID = iTargetTypeID Then
                    oTarget = oEnvir.oEntity(X)
                    Exit For
                End If
            Next X

            If oAttacker Is Nothing OrElse oTarget Is Nothing Then Return

            Dim lClrIdx As Int32 = CInt(yMissileType) - CInt(WeaponType.eMissile_Color_1)
            goMissileMgr.AddMissile(New Microsoft.DirectX.Vector3(oAttacker.LocX, oAttacker.LocY, oAttacker.LocZ), oTarget, ySpeed, yManeuver, lClrIdx, lMissileID, bHit)
        End If
    End Sub

    'Private Sub HandleMissileImpact(ByRef yData() As Byte)
    '    msCurrentLoc = "HandleMissileImpact"
    '    Dim lMissileID As Int32 = System.BitConverter.ToInt32(yData, 2)
    '    Dim yAOERange As Byte = yData(6)

    '    If goMissileMgr Is Nothing = False Then goMissileMgr.MissileImpact(lMissileID, yAOERange)
    'End Sub

    Private Sub HandleMissileDetonated(ByRef yData() As Byte)
        msCurrentLoc = "HandleMissileDetonated"
        Dim lMissileID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim yAOERange As Byte = yData(6)
        If yAOERange > 0 Then
            Dim lX As Int32 = System.BitConverter.ToInt32(yData, 7)
            Dim lZ As Int32 = System.BitConverter.ToInt32(yData, 11)
            If goMissileMgr Is Nothing = False Then goMissileMgr.MissileDetonated(lMissileID, lX, lZ, yAOERange)
        Else
            If goMissileMgr Is Nothing = False Then goMissileMgr.MissileBreakApart(lMissileID)
        End If

    End Sub

    Private Sub HandleTradeWindowMsg(ByVal iMsgCode As Int16, ByRef yData() As Byte)
        msCurrentLoc = "HandleTradeWindowMsg"
        If goUILib Is Nothing Then Return
        Dim ofrm As frmTradeMain = CType(goUILib.GetWindow("frmTradeMain"), frmTradeMain)
        'If ofrm Is Nothing Then Return

        Select Case iMsgCode
            Case GlobalMessageCode.eGetGTCList
                If ofrm Is Nothing = False AndAlso ofrm.fraContents Is Nothing = False AndAlso ofrm.fraContents.ControlName = "fraTradeContents" Then
                    CType(ofrm.fraContents, fraTradeContents).HandleGTCList(yData)
                End If
            Case GlobalMessageCode.eGetTradePostList
                If ofrm Is Nothing = False Then ofrm.HandleGetTradePostList(yData)
                Dim frmGuild As frmGuildMain = CType(goUILib.GetWindow("frmGuildMain"), frmGuildMain)
                If frmGuild Is Nothing = False Then
                    frmGuild.HandleGetTradepostList(yData)
                End If
            Case GlobalMessageCode.eGetOrderSpecifics
                If ofrm Is Nothing = False AndAlso ofrm.fraContents Is Nothing = False AndAlso ofrm.fraContents.ControlName = "fraTradeContents" Then
                    CType(ofrm.fraContents, fraTradeContents).HandleGetOrderSpecifics(yData)
                End If
            Case GlobalMessageCode.eGetTradePostTradeables
                TradePostContents.HandleTradepostTradeables(yData)
            Case GlobalMessageCode.eSubmitBuyOrder
                Dim yVal As Byte = yData(6)
                Select Case yVal
                    Case 0
                        goUILib.AddNotification("Buy Order Submitted and Approved.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If ofrm Is Nothing = False Then ofrm.ReturnToCurrentView()
                    Case 1
                        goUILib.AddNotification("Unable to submit Buy Order. Maximum Buy Slots have been exceeded.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                    Case 2
                        goUILib.AddNotification("Unable to submit Buy Order. Please try again later.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                    Case 3
                        goUILib.AddNotification("Invalid Property values were entered.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                    Case 4
                        goUILib.AddNotification("Insufficient credits in treasury for Payment Amount entered.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("Game Narrative\LowCredits.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
                End Select
            Case GlobalMessageCode.eDeliverBuyOrder
                Dim yVal As Byte = yData(2)
                Select Case yVal
                    Case 0
                        goUILib.AddNotification("Delivery Order Approved. Shipment has begun.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If ofrm Is Nothing = False Then ofrm.ReturnToCurrentView()
                    Case 1
                        goUILib.AddNotification("Delivery Order unaccepted: Deadline would be exceeded.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                    Case 2
                        goUILib.AddNotification("Delivery Order unaccepted: Cargo does not meet criteria.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                    Case 255
                        goUILib.AddNotification("Delivery Order Unaccepted: Unknown reason, please try again.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                End Select
            Case GlobalMessageCode.ePurchaseSellOrder
                Dim yVal As Byte = yData(2)
                Select Case yVal
                    Case 0
                        Dim sName As String = GetStringFromBytes(yData, 3, 20)
                        goUILib.AddNotification(sName & " Purchased. Shipment has begun.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If ofrm Is Nothing = False Then ofrm.ReturnToCurrentView()
                    Case 1
                        goUILib.AddNotification("Sell Order Purchase Failed: Not enough supply in Sell Order.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                    Case 255
                        goUILib.AddNotification("Sell Order Purchase Failed: Unknown reason, please try again.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                End Select
            Case GlobalMessageCode.eAcceptBuyOrder
                Dim yVal As Byte = yData(6)
                Select Case yVal
                    Case 0
                        goUILib.AddNotification("Buy Order Accepted. Funds have been escrowed. Complete the buy order by the deadline in order to receive the payment amount and escrow funds.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If ofrm Is Nothing = False Then ofrm.ReturnToCurrentView()
                    Case 1
                        goUILib.AddNotification("Insufficient credits for the Escrow fund requirement.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                    Case 2
                        goUILib.AddNotification("Buy Order Tradepost is out of range of the selected tradepost.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                    Case 3
                        goUILib.AddNotification("Buy Order is no longer available as it has already been accepted.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                End Select
        End Select
    End Sub

    Private Sub HandleGetTradeHistory(ByRef yData() As Byte)
        msCurrentLoc = "HandleGetTradeHistory"
        Dim lPos As Int32 = 2       'for msgcode
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lUB As Int32 = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4

        If glPlayerID = lPlayerID Then
            goCurrentPlayer.mlTradeHistoryUB = -1
            For X As Int32 = 0 To lUB
                Dim lTransDate As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lOtherPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                lTransDate = GetDateAsNumber(Now.Subtract(New TimeSpan(0, 0, lTransDate)))

                Dim lIdx As Int32 = -1
                With goCurrentPlayer
                    For Y As Int32 = 0 To .mlTradeHistoryUB
                        If .moTradeHistory(Y).lOtherPlayerID = lOtherPlayerID AndAlso .moTradeHistory(Y).lTransactionDate = lTransDate Then
                            lIdx = Y
                            Exit For
                        End If
                    Next Y
                    If lIdx = -1 Then
                        .mlTradeHistoryUB += 1
                        ReDim Preserve .moTradeHistory(.mlTradeHistoryUB)
                        .moTradeHistory(.mlTradeHistoryUB) = New TradeHistory()
                        lIdx = .mlTradeHistoryUB
                    End If
                End With
                With goCurrentPlayer.moTradeHistory(lIdx)
                    .lTransactionDate = lTransDate
                    .lOtherPlayerID = lOtherPlayerID
                    .lPlayerID = lPlayerID
                    .yTradeResult = CType(yData(lPos), TradeHistory.TradeResult) : lPos += 1
                    .yTradeEventType = CType(yData(lPos), TradeHistory.TradeEventType) : lPos += 1
                    .blTradeAmt = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                    .lDeliveryTime = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                    .ClearItems()

                    Dim lItemUB As Int32 = yData(lPos) - 1 : lPos += 1
                    For Y As Int32 = 0 To lItemUB
                        Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                        Dim blQty As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                        Dim yType As Byte = yData(lPos) : lPos += 1

                        .AddTradeItem(sName, blQty, CType(yType, TradeHistory.TradeHistoryItemType))
                    Next Y
                End With
            Next X
            'Else
            '	'TODO: Espionage?
        End If
    End Sub

    Private Sub HandleGetTradeDeliveries(ByRef yData() As Byte)
        msCurrentLoc = "HandleGetTradeDeliveries"
        Try
            Dim lPos As Int32 = 2           'for msgcode
            Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'Dim lServerDT As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lCnt As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            For X As Int32 = 0 To lCnt - 1
                Dim lTPID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lSourceID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lDelivery As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lStartedOn As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lColonyID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                lStartedOn = GetDateAsNumber(Now.Subtract(New TimeSpan(0, 0, lStartedOn)))

                Dim oTDP As TradeDeliveryPackage = TradeDeliveryPackage.AddOrGetTradeDelivery(lTPID, lSourceID, lDelivery, lStartedOn)
                If oTDP Is Nothing = False Then
                    oTDP.lColonyID = lColonyID
                    Dim lID1 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim iID2 As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    Dim lID3 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim blQty As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                    oTDP.AddItem(lID1, iID2, lID3, blQty)
                Else : lPos += 18
                End If
            Next X
        Catch ex As Exception
            'do nothing
        End Try
    End Sub

    Private Sub HandleSetPlayerSpecialAttribute(ByRef yData() As Byte)
        msCurrentLoc = "HandleSetPlayerSpecialAttribute"
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iAttrID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lValue As Int32 = System.BitConverter.ToInt32(yData, 8)

        If goCurrentPlayer Is Nothing = False AndAlso lPlayerID = glPlayerID Then
            goCurrentPlayer.SetPlayerSpecialAttributeSetting(CType(iAttrID, ePlayerSpecialAttributeSetting), lValue)
            'Else
            '	'TODO: Espionage?
        End If

    End Sub

    Private Sub HandleTransferContents(ByRef yData() As Byte)
        msCurrentLoc = "HandleTransferContents"
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lFromID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iFromTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
        Dim lToID As Int32 = System.BitConverter.ToInt32(yData, 14)
        Dim iToTypeID As Int16 = System.BitConverter.ToInt16(yData, 18)
        Dim lQty As Int32 = System.BitConverter.ToInt32(yData, 20)

        If lQty = -1 Then
            goUILib.AddNotification("Cargo bay is inoperable.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        Else
            goUILib.AddNotification("Quantity Transferred: " & lQty & ".", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

            Dim lRequestEntityContentsID As Int32 = -1
            Dim iRequestEntityContentsTypeID As Int16 = -1

            'Ok, lets get our stuff
            If iFromTypeID = ObjectType.eColony Then
                If goAvailableResources Is Nothing = False AndAlso goAvailableResources.lColonyID = lFromID Then
                    goAvailableResources.AdjustObjectQuantity(lObjID, iObjTypeID, -lQty)
                End If
                lRequestEntityContentsID = lToID
                iRequestEntityContentsTypeID = iToTypeID
            ElseIf iToTypeID = ObjectType.eColony Then
                'get available resources will handle it
                frmTransfer.bReRequestResources = True
                lRequestEntityContentsID = lFromID
                iRequestEntityContentsTypeID = iFromTypeID
            Else
                '????
            End If

            If lRequestEntityContentsID <> -1 AndAlso iRequestEntityContentsTypeID <> -1 Then
                Dim yMsg(7) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eRequestEntityContents).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(lRequestEntityContentsID).CopyTo(yMsg, 2)
                System.BitConverter.GetBytes(iRequestEntityContentsTypeID).CopyTo(yMsg, 6)
                SendToPrimary(yMsg)
            End If
        End If
    End Sub

    Private Sub HandleMoveObjectRequestDeny(ByRef yData() As Byte)
        msCurrentLoc = "HandleMoveObjectRequestDeny"
        If goSound Is Nothing = False Then goSound.StartUnitSpeechSound(False, goSound.GetUnitSpeechTypeCnt(SoundMgr.SpeechType.eMoveRequest, 1))
    End Sub

    Private Sub HandlePlayerAliasConfig(ByRef yData() As Byte)
        msCurrentLoc = "HandlePlayerAliasConfig"
        Dim lPos As Int32 = 2   'for msgcode
        Dim yType As Byte = yData(lPos) : lPos += 1
        'AliasPlayer(20), AliasUN(20), AliasPW(20), Rights(4)
        Dim yPlayer(19) As Byte
        Array.Copy(yData, lPos, yPlayer, 0, 20) : lPos += 20
        Dim yUserName(19) As Byte
        Array.Copy(yData, lPos, yUserName, 0, 20) : lPos += 20
        Dim yPassword(19) As Byte
        Array.Copy(yData, lPos, yPassword, 0, 20) : lPos += 20
        Dim lRights As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If yType = 1 OrElse yType = 2 Then
            'Add or edit
            If goUILib Is Nothing = False Then goUILib.AddNotification("Player Alias Configuration Updated.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If gbAliased = True AndAlso gsAliasUserName = GetStringFromBytes(yUserName, 0, 20).ToUpper Then
                'Ok, set my stuff up
                'gsAliasUserName = GetStringFromBytes(yUserName, 0, 20)
                gsAliasPassword = GetStringFromBytes(yPassword, 0, 20)
                glAliasRights = lRights
            Else
                Dim bFound As Boolean = False
                Dim sPlayerName As String = GetStringFromBytes(yPlayer, 0, 20).ToUpper
                For X As Int32 = 0 To goCurrentPlayer.mlAliasUB
                    If goCurrentPlayer.muAliases(X).sPlayerName.ToUpper = sPlayerName Then
                        bFound = True
                        goCurrentPlayer.muAliases(X).sPlayerName = GetStringFromBytes(yPlayer, 0, 20)
                        goCurrentPlayer.muAliases(X).sPassword = GetStringFromBytes(yPassword, 0, 20)
                        goCurrentPlayer.muAliases(X).sUserName = GetStringFromBytes(yUserName, 0, 20)
                        goCurrentPlayer.muAliases(X).lRights = lRights
                        Exit For
                    End If
                Next X
                If bFound = False Then
                    ReDim Preserve goCurrentPlayer.muAliases(goCurrentPlayer.mlAliasUB + 1)
                    With goCurrentPlayer.muAliases(goCurrentPlayer.mlAliasUB + 1)
                        .sPlayerName = GetStringFromBytes(yPlayer, 0, 20)
                        .sPassword = GetStringFromBytes(yPassword, 0, 20)
                        .sUserName = GetStringFromBytes(yUserName, 0, 20)
                        .lRights = lRights
                    End With
                    goCurrentPlayer.mlAliasUB += 1
                End If
            End If
        ElseIf yType = 255 Then
            'unable to find player
            If goUILib Is Nothing = False Then goUILib.AddNotification("Unable to find player: " & GetStringFromBytes(yPlayer, 0, 20) & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf yType = 0 Then
            If goUILib Is Nothing = False Then goUILib.AddNotification("Player Alias Configuration removed.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If gbAliased = True Then
                'glCurrentEnvirView = CurrentView.eStartupDSELogo
                'goUILib.RemoveAllWindows()
                'DisconnectAll()
                'goCurrentEnvir = Nothing
                RaiseEvent ServerShutdown()
            Else
                Dim sPlayerName As String = GetStringFromBytes(yPlayer, 0, 20).ToUpper
                For X As Int32 = 0 To goCurrentPlayer.mlAliasUB
                    If goCurrentPlayer.muAliases(X).sPlayerName = sPlayerName Then
                        goCurrentPlayer.muAliases(X).sPlayerName = ""
                    End If
                Next X
            End If
            Return
        ElseIf yType = 254 Then
            'Add to player details...
            Dim sPlayerName As String = GetStringFromBytes(yPlayer, 0, 20).ToUpper
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To goCurrentPlayer.mlAliasUB
                If goCurrentPlayer.muAliases(X).sPlayerName = sPlayerName Then
                    lIdx = X
                    Exit For
                End If
            Next X
            If lIdx = -1 Then
                lIdx = goCurrentPlayer.mlAliasUB + 1
                ReDim Preserve goCurrentPlayer.muAliases(lIdx)
                goCurrentPlayer.mlAliasUB += 1
            End If
            With goCurrentPlayer.muAliases(lIdx)
                .sPlayerName = sPlayerName
                .sPassword = GetStringFromBytes(yPassword, 0, 20)
                .sUserName = GetStringFromBytes(yUserName, 0, 20)
                .lRights = lRights
            End With
        ElseIf yType = 253 Then
            'Add to player details...
            Dim sPlayerName As String = GetStringFromBytes(yPlayer, 0, 20).ToUpper
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To goCurrentPlayer.mlAllowanceUB
                If goCurrentPlayer.muAllowances(X).sPlayerName = sPlayerName Then
                    lIdx = X
                    Exit For
                End If
            Next X
            If lIdx = -1 Then
                lIdx = goCurrentPlayer.mlAllowanceUB + 1
                ReDim Preserve goCurrentPlayer.muAllowances(lIdx)
                goCurrentPlayer.mlAllowanceUB += 1
            End If
            With goCurrentPlayer.muAllowances(lIdx)
                .sPlayerName = sPlayerName
                .sPassword = GetStringFromBytes(yPassword, 0, 20)
                .sUserName = GetStringFromBytes(yUserName, 0, 20)
                .lRights = lRights
            End With
        End If
    End Sub

    Private Sub HandleActOfWarNotice(ByRef yData() As Byte)
        msCurrentLoc = "HandleActOfWarNotice"
        Dim lPos As Int32 = 2 'for msgcode
        Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        If goUILib Is Nothing = False Then
            'goUILib.AddNotification("Placing an armed facility within range of a non-allied faction may be viewed as an act of war!", _
            '   Color.Red, lParentID, iParentTypeID, lEntityID, iEntityTypeID, Int32.MinValue, Int32.MinValue)
            goUILib.AddNotification("Unable to build facilities with weapons in range of other player's facilities!", _
             Color.Red, lParentID, iParentTypeID, lEntityID, iEntityTypeID, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Alert.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
        End If
    End Sub

    Private Sub HandleDeinfiltrateAgent(ByRef yData() As Byte)
        msCurrentLoc = "HandleDeinfiltrateAgent"
        Dim lPos As Int32 = 2 'msgcode
        Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lResult As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If lResult = -2 Then
            If goUILib Is Nothing = False Then goUILib.AddNotification("That agent is dead and can no longer receive orders.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lResult = -3 Then
            If goUILib Is Nothing = False Then goUILib.AddNotification("That agent has been captured and can no longer receive orders.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lResult = -4 Then
            If goUILib Is Nothing = False Then goUILib.AddNotification("That agent is already returning home.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lResult = -5 Then
            If goUILib Is Nothing = False Then goUILib.AddNotification("That agent is currently on a mission and cannot deinfiltrate at this time.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lResult = -1 Then
            If goUILib Is Nothing = False Then goUILib.AddNotification("That agent is lying low waiting for another opportunity to deinfiltrate.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lResult = 0 Then
            If goUILib Is Nothing = False Then goUILib.AddNotification("Agent has deinfiltrated and is on their way home.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lResult = Int32.MinValue Then
            Dim sAgentName As String = ""
            If goCurrentPlayer Is Nothing = False Then
                For X As Int32 = 0 To goCurrentPlayer.AgentUB
                    If goCurrentPlayer.AgentIdx(X) = lAgentID Then
                        sAgentName = goCurrentPlayer.Agents(X).sAgentName
                        If (goCurrentPlayer.Agents(X).lAgentStatus And AgentStatus.ReturningHome) <> 0 Then goCurrentPlayer.Agents(X).lAgentStatus = goCurrentPlayer.Agents(X).lAgentStatus Xor AgentStatus.ReturningHome
                        Exit For
                    End If
                Next X
            End If

            If goUILib Is Nothing = False Then goUILib.AddNotification(sAgentName & " (Agent) has returned home and is ready for deployment.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        If lResult < 256 Then
            If goCurrentPlayer Is Nothing = False Then
                For X As Int32 = 0 To goCurrentPlayer.AgentUB
                    If goCurrentPlayer.AgentIdx(X) = lAgentID Then
                        goCurrentPlayer.Agents(X).Suspicion = CByte(lResult)
                        If lResult = 0 Then
                            goCurrentPlayer.Agents(X).lAgentStatus = goCurrentPlayer.Agents(X).lAgentStatus Or AgentStatus.ReturningHome
                            goCurrentPlayer.Agents(X).lTargetID = -1
                        Else
                            If goUILib Is Nothing = False Then goUILib.AddNotification(goCurrentPlayer.Agents(X).sAgentName & " was stopped at a border checkpoint and questioned. The authorities will not let them leave at this time.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        End If
                        Exit For
                    End If
                Next X
            End If
        End If
    End Sub

    Private Sub HandleSetAgentStatus(ByRef yData() As Byte)
        msCurrentLoc = "HandleSetAgentStatus"
        Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lStatus As Int32 = System.BitConverter.ToInt32(yData, 6)

        Dim oAgent As Agent = Nothing
        For X As Int32 = 0 To goCurrentPlayer.AgentUB
            If goCurrentPlayer.AgentIdx(X) = lAgentID Then
                oAgent = goCurrentPlayer.Agents(X)
                Exit For
            End If
        Next X

        If oAgent Is Nothing = False Then
            If lStatus = Int32.MinValue Then
                'indicates failure...
                If (oAgent.lAgentStatus And AgentStatus.NewRecruit) <> 0 Then
                    'ok, this agent is a new recruit, reason should be that we could not afford it
                    If goUILib Is Nothing = False Then goUILib.AddNotification("Unable to recruit " & oAgent.sAgentName & " due to low credits.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Else
                    'ok this agent is recruited, reason should be that the agent cannot be dissmised??
                    If goUILib Is Nothing = False Then goUILib.AddNotification("Unable to dismiss " & oAgent.sAgentName & " at this time.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
            ElseIf lStatus = Int32.MaxValue Then
                'ok, the agent is gone gone...
                For X As Int32 = 0 To goCurrentPlayer.AgentUB
                    If goCurrentPlayer.AgentIdx(X) = lAgentID Then
                        goCurrentPlayer.AgentIdx(X) = -1
                        goCurrentPlayer.Agents(X) = Nothing
                        Exit For
                    End If
                Next X
                oAgent = Nothing
            ElseIf lStatus = AgentStatus.NewRecruit Then
                'ok, recruiting, remove the recruit status and set to normal
                oAgent.lAgentStatus = AgentStatus.NormalStatus
                oAgent.dtRecruited = Now
            ElseIf lStatus = AgentStatus.Dismissed Then
                'ok, dismissing... what to do...
                oAgent.lAgentStatus = AgentStatus.Dismissed
            ElseIf lStatus = AgentStatus.ReturningHome Then
                'ok... check what we were doing before returning home... if I am returning home, I cannot be:
                '  dismissed, dead, infiltrating, infiltrated, on a mission, a counter agent, captured
                'therefore, the agent's status is returnhome, only... however, if the agent was infiltrating, then that is important to know
                If (oAgent.lAgentStatus And AgentStatus.Infiltrating) <> 0 Then
                    Dim sName As String = GetCacheObjectValue(oAgent.lTargetID, oAgent.iTargetTypeID)
                    If goUILib Is Nothing = False Then goUILib.AddNotification(oAgent.sAgentName & " (Agent) could not infiltrate " & sName & " and is returning home.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                ElseIf (oAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then
                    If goUILib Is Nothing = False Then goUILib.AddNotification(oAgent.sAgentName & " (Agent) has escaped and is returning home!", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Else
                    If goUILib Is Nothing = False Then goUILib.AddNotification(oAgent.sAgentName & " (Agent) is returning home.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
                oAgent.lAgentStatus = AgentStatus.ReturningHome
                'Now, because this is returning home, let's get the cycle delay
                Dim lCycles As Int32 = System.BitConverter.ToInt32(yData, 10)
                oAgent.mlArrivalCycles = lCycles
                oAgent.mlReceivedCycle = glCurrentCycle
                oAgent.InfiltrationLevel = 0
                oAgent.InfiltrationType = eInfiltrationType.eGeneralInfiltration
                oAgent.lTargetID = -1
                oAgent.iTargetTypeID = -1
            ElseIf lStatus = AgentStatus.IsInfiltrated Then
                If (oAgent.lAgentStatus And AgentStatus.Infiltrating) <> 0 Then oAgent.lAgentStatus = oAgent.lAgentStatus Xor AgentStatus.Infiltrating
                oAgent.lAgentStatus = oAgent.lAgentStatus Or AgentStatus.IsInfiltrated
                Dim sName As String = GetCacheObjectValue(oAgent.lTargetID, oAgent.iTargetTypeID)
                If NewTutorialManager.TutorialOn = True Then
                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eAgentInfiltrated, -1, -1, -1, "")
                End If
                If goUILib Is Nothing = False Then goUILib.AddNotification(oAgent.sAgentName & " (Agent) has successfully infiltrated " & sName & " and is ready for missions.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
        End If
    End Sub

    Private Sub HandleSetInfiltrateSettings(ByRef yData() As Byte)
        msCurrentLoc = "HandleSetInfiltrateSettings"
        Dim lPos As Int32 = 2
        Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lInfType As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lInfTargetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iInfTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lFreq As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If goCurrentPlayer Is Nothing = False Then
            For X As Int32 = 0 To goCurrentPlayer.AgentUB
                If goCurrentPlayer.AgentIdx(X) = lAgentID Then

                    If lInfType = Int32.MinValue Then
                        'invalid target
                        If goUILib Is Nothing = False Then goUILib.AddNotification("That is an invalid infiltration target.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    Else
                        Dim lArrivalCycles As Int32 = -1
                        If yData.Length >= lPos + 4 Then
                            lArrivalCycles = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        End If

                        With goCurrentPlayer.Agents(X)
                            .InfiltrationType = CType(lInfType, eInfiltrationType)
                            .yReportFreq = CByte(lFreq)
                            .lTargetID = lInfTargetID
                            .iTargetTypeID = iInfTargetTypeID
                            .mlArrivalCycles = lArrivalCycles
                            .mlReceivedCycle = glCurrentCycle
                        End With
                        If goUILib Is Nothing = False Then goUILib.AddNotification("Infiltration Settings Confirmed.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    End If

                    Exit For
                End If
            Next X
        End If
    End Sub

    Private Sub HandleSubmitMission(ByRef yData() As Byte)
        msCurrentLoc = "HandleSubmitMission"

        Dim lPos As Int32 = 2   'for msgcode
        Dim lPM_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        mlLastAddObjectID = lPM_ID
        miLastAddObjectTypeID = -1

        Dim oPM As PlayerMission = Nothing

        If goCurrentPlayer Is Nothing Then Return
        For X As Int32 = 0 To goCurrentPlayer.PlayerMissionUB
            If goCurrentPlayer.PlayerMissionIdx(X) = lPM_ID Then
                oPM = goCurrentPlayer.PlayerMissions(X)
                Exit For
            End If
        Next X
        If oPM Is Nothing Then oPM = New PlayerMission()

        If oPM.FillFromMsg(yData) = True Then
            'ok, find our mission
            Dim lIdx As Int32 = -1
            Dim lFirstIdx As Int32 = -1
            For X As Int32 = 0 To goCurrentPlayer.PlayerMissionUB
                If goCurrentPlayer.PlayerMissionIdx(X) = oPM.PM_ID Then
                    lIdx = X
                    Exit For
                ElseIf goCurrentPlayer.PlayerMissionIdx(X) = -1 AndAlso lFirstIdx = -1 Then
                    lFirstIdx = X
                End If
            Next X

            If lIdx = -1 Then
                If lFirstIdx = -1 Then
                    lFirstIdx = goCurrentPlayer.PlayerMissionUB + 1
                    ReDim Preserve goCurrentPlayer.PlayerMissionIdx(lFirstIdx)
                    ReDim Preserve goCurrentPlayer.PlayerMissions(lFirstIdx)
                    goCurrentPlayer.PlayerMissionIdx(lFirstIdx) = -1
                    goCurrentPlayer.PlayerMissionUB = lFirstIdx
                End If
                lIdx = lFirstIdx
                goCurrentPlayer.PlayerMissions(lIdx) = oPM
                goCurrentPlayer.PlayerMissionIdx(lIdx) = oPM.PM_ID
            End If

            If goUILib Is Nothing = False Then
                Dim oWin As frmMission = CType(goUILib.GetWindow("frmMission"), frmMission)
                If oWin Is Nothing = False Then
                    Dim lVal As Int32 = oWin.GetMissionPMID()
                    If lVal = -1 OrElse lVal = oPM.PM_ID Then
                        'goUILib.RemoveWindow(oWin.ControlName)
                        goUILib.AddNotification("Mission Plan Accepted!", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    End If
                End If
                oWin = Nothing
            End If

        End If
    End Sub

    Private Sub HandleGetAgentStatus(ByRef yData() As Byte)
        msCurrentLoc = "HandleGetAgentStatus"
        Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, 2)

        Dim oAgent As Agent = Nothing
        For X As Int32 = 0 To goCurrentPlayer.AgentUB
            If goCurrentPlayer.AgentIdx(X) = lAgentID Then
                oAgent = goCurrentPlayer.Agents(X)
                Exit For
            End If
        Next X

        If oAgent Is Nothing = False Then
            Dim lStatus As Int32 = System.BitConverter.ToInt32(yData, 6)

            oAgent.lAgentStatus = CType(lStatus, AgentStatus)

            oAgent.Suspicion = yData(10)
            oAgent.InfiltrationLevel = yData(11)
            oAgent.mlArrivalCycles = System.BitConverter.ToInt32(yData, 12)
            oAgent.mlReceivedCycle = glCurrentCycle
        End If
    End Sub

    Private Sub HandleGetPMUpdate(ByRef yData() As Byte)
        msCurrentLoc = "HandleGetPMUpdate"
        Dim lPos As Int32 = 2   'for msgcode
        Dim lPM_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If goCurrentPlayer Is Nothing Then Return
        Dim oPM As PlayerMission = Nothing
        For X As Int32 = 0 To goCurrentPlayer.PlayerMissionUB
            If goCurrentPlayer.PlayerMissionIdx(X) = lPM_ID Then
                oPM = goCurrentPlayer.PlayerMissions(X)
                Exit For
            End If
        Next X
        If oPM Is Nothing = True Then Return

        Dim yValue As Byte = yData(lPos) : lPos += 1
        With oPM
            If (yValue And 128) <> 0 Then
                .bAlarmThrown = True
                yValue = CByte(yValue Xor 128)
            End If
            .ySafeHouseSetting = yValue
            yValue = yData(lPos) : lPos += 1
            .yCurrentPhase = CType(yValue, eMissionPhase)


            If .ySafeHouseSetting <> 0 Then
                If .oSafehouseMissionGoal Is Nothing Then
                    .oSafehouseMissionGoal = New PlayerMissionGoal()
                    .oSafehouseMissionGoal.oGoal = Goal.GetSafehouseGoal()
                    .oSafehouseMissionGoal.oMission = oPM
                End If
                'skillsetid
                Dim lTemp As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .oSafehouseMissionGoal.oSkillSet = .oSafehouseMissionGoal.oGoal.GetSkillset(lTemp)
                'assignment1 agentid
                Dim lAgent1 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                'assignment1 skillid
                Dim lSkill1 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim iPoints1 As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                Dim yStatus1 As Byte = yData(lPos) : lPos += 1
                'assignment2 agentid
                Dim lAgent2 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                'assignment2 skillid
                Dim lSkill2 As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim iPoints2 As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                Dim yStatus2 As Byte = yData(lPos) : lPos += 1

                Dim oPMG As PlayerMissionGoal = .oSafehouseMissionGoal
                If lAgent1 > -1 Then
                    For Z As Int32 = 0 To oPMG.lAssignmentUB
                        If oPMG.oAssignments(Z) Is Nothing = False AndAlso oPMG.oAssignments(Z).oAgent.ObjectID = lAgent1 AndAlso oPMG.oAssignments(Z).oSkill.ObjectID = lSkill1 Then
                            oPMG.oAssignments(Z).PointsProduced = iPoints1
                            oPMG.oAssignments(Z).yStatus = yStatus1
                            Exit For
                        End If
                    Next Z
                End If
                If lAgent2 > -1 Then
                    For Z As Int32 = 0 To oPMG.lAssignmentUB
                        If oPMG.oAssignments(Z) Is Nothing = False AndAlso oPMG.oAssignments(Z).oAgent.ObjectID = lAgent2 AndAlso oPMG.oAssignments(Z).oSkill.ObjectID = lSkill2 Then
                            oPMG.oAssignments(Z).PointsProduced = iPoints2
                            oPMG.oAssignments(Z).yStatus = yStatus2
                            Exit For
                        End If
                    Next Z
                End If
            End If


            Dim lCnt As Int32 = yData(lPos) : lPos += 1

            If .oMission Is Nothing = False Then
                For X As Int32 = 0 To lCnt - 1
                    Dim lGoalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    Dim lAssignCnt As Int32 = yData(lPos) : lPos += 1

                    'now, find our missiongoal
                    Dim oPMG As PlayerMissionGoal = Nothing
                    For Y As Int32 = 0 To .oMission.GoalUB
                        If .oMission.MethodIDs(Y) = .lMethodID AndAlso .oMissionGoals(Y) Is Nothing = False AndAlso .oMission.Goals(Y).ObjectID = lGoalID Then
                            oPMG = .oMissionGoals(Y)
                            Exit For
                        End If
                    Next Y
                    If oPMG Is Nothing = False Then
                        For Y As Int32 = 0 To lAssignCnt - 1
                            Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                            Dim lSkillID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                            Dim iPoints As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                            Dim yStatus As Byte = yData(lPos) : lPos += 1

                            For Z As Int32 = 0 To oPMG.lAssignmentUB
                                If oPMG.oAssignments(Z) Is Nothing = False AndAlso oPMG.oAssignments(Z).oAgent.ObjectID = lAgentID AndAlso oPMG.oAssignments(Z).oSkill.ObjectID = lSkillID Then
                                    oPMG.oAssignments(Z).PointsProduced = iPoints
                                    oPMG.oAssignments(Z).yStatus = yStatus
                                    Exit For
                                End If
                            Next Z
                        Next Y
                    Else : Return
                    End If
                Next X
            End If
        End With
    End Sub

    Private Sub HandleSetSkipStatus(ByRef yData() As Byte)
        msCurrentLoc = "HandleSetSkipStatus"
        Dim lPos As Int32 = 2   'for msgcode
        Dim lPMID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lGoalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oPM As PlayerMission = Nothing

        If goCurrentPlayer Is Nothing = False Then
            For X As Int32 = 0 To goCurrentPlayer.PlayerMissionUB
                If goCurrentPlayer.PlayerMissionIdx(X) = lPMID Then
                    oPM = goCurrentPlayer.PlayerMissions(X)
                    Exit For
                End If
            Next X
        Else : Return
        End If
        If oPM Is Nothing Then Return

        If lGoalID > 0 Then
            If oPM.oMission Is Nothing Then Return

            Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lSkillID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim ySkipVal As Byte = yData(lPos) : lPos += 1

            Dim oPMG As PlayerMissionGoal = Nothing
            For X As Int32 = 0 To oPM.oMission.GoalUB
                If oPM.lMethodID = oPM.oMission.MethodIDs(X) AndAlso oPM.oMission.Goals(X).ObjectID = lGoalID Then
                    oPMG = oPM.oMissionGoals(X)
                    Exit For
                End If
            Next X
            If oPMG Is Nothing Then Return
            For X As Int32 = 0 To oPMG.lAssignmentUB
                If oPMG.oAssignments(X) Is Nothing = False AndAlso oPMG.oAssignments(X).oAgent Is Nothing = False AndAlso oPMG.oAssignments(X).oSkill Is Nothing = False Then
                    If oPMG.oAssignments(X).oAgent.ObjectID = lAgentID AndAlso oPMG.oAssignments(X).oSkill.ObjectID = lSkillID Then
                        If ySkipVal = 0 Then
                            'ok not skipping
                            If (oPMG.oAssignments(X).yStatus And AgentAssignmentStatus.Skipped) <> 0 Then
                                oPMG.oAssignments(X).yStatus = oPMG.oAssignments(X).yStatus Xor AgentAssignmentStatus.Skipped
                            End If
                        Else
                            'skipping
                            oPMG.oAssignments(X).yStatus = oPMG.oAssignments(X).yStatus Or AgentAssignmentStatus.Skipped
                        End If
                        Exit For
                    End If
                End If
            Next X
        Else
            'cancel, resume, pause or Cancel Error
            If lGoalID = -1 Then
                'ok, cancel...
                oPM.yCurrentPhase = eMissionPhase.eCancelled
                If goUILib Is Nothing = False Then goUILib.AddNotification("Mission Cancelled.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            ElseIf lGoalID = Int32.MinValue Then
                'cancel error
                If goUILib Is Nothing = False Then goUILib.AddNotification("Unable to cancel this mission as it has passed the point of no return.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Else
                'pause or resume
                'regardless, we need to get the next value
                Dim yStatus As Byte = yData(lPos) : lPos += 1
                oPM.yCurrentPhase = CType(yStatus, eMissionPhase)
            End If
        End If
    End Sub

    Private Sub HandleCapturedAgent(ByRef yData() As Byte)
        msCurrentLoc = "HandleCapturedAgent"
        'I have captured an agent... or data about an agent I have captured has been updated
        Dim lPos As Int32 = 2   'for msgcode

        Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yAction As Byte = yData(lPos) : lPos += 1

        mlLastAddObjectID = lAgentID
        miLastAddObjectTypeID = ObjectType.eAgent

        If yAction = 0 Then     'add captured agent
            'see if we have this agent in our list already
            Dim oCA As CapturedAgent = Nothing
            If goCurrentPlayer Is Nothing = False Then
                For X As Int32 = 0 To goCurrentPlayer.CapturedAgentUB
                    If goCurrentPlayer.CapturedAgentIdx(X) = lAgentID Then
                        oCA = goCurrentPlayer.CapturedAgents(X)
                        Exit For
                    End If
                Next X
            End If

            Dim bNew As Boolean = False
            If oCA Is Nothing Then
                bNew = True
                oCA = New CapturedAgent
                oCA.ObjectID = lAgentID
                oCA.ObjTypeID = ObjectType.eAgent
            End If

            Dim bNewData As Boolean = False
            With oCA
                Dim sName As String = GetStringFromBytes(yData, lPos, 30) : lPos += 30
                If bNew = False AndAlso .sAgentName <> sName Then bNewData = True
                .sAgentName = sName

                Dim lTemp As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                If bNew = False AndAlso .lOwnerID <> lTemp Then bNewData = True
                .lOwnerID = lTemp

                Dim yTemp As Byte = yData(lPos) : lPos += 1
                'If bNew = False AndAlso .yHealth <> yTemp Then bNewData = True
                .yHealth = yTemp

                yTemp = yData(lPos) : lPos += 1
                If bNew = False AndAlso .yInfLevel <> yTemp Then bNewData = True
                .yInfLevel = yTemp

                yTemp = yData(lPos) : lPos += 1
                If bNew = False AndAlso .yInfType <> yTemp Then bNewData = True
                .yInfType = yTemp

                .yInterrogationProgress = yData(lPos) : lPos += 1
                .lInterrogatorID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                If bNew = False AndAlso .lMissionID <> lTemp Then bNewData = True
                .lMissionID = lTemp

                lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                If bNew = False AndAlso .lMissionTargetID <> lTemp Then bNewData = True
                .lMissionTargetID = lTemp

                lTemp = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                If bNew = False AndAlso .iMissionTargetTypeID <> lTemp Then bNewData = True
                .iMissionTargetTypeID = CShort(lTemp)
            End With

            If bNew = True Then
                Dim lIdx As Int32 = -1
                For X As Int32 = 0 To goCurrentPlayer.CapturedAgentUB
                    If goCurrentPlayer.CapturedAgentIdx(X) = -1 Then
                        lIdx = X
                        Exit For
                    End If
                Next X
                If lIdx = -1 Then
                    lIdx = goCurrentPlayer.CapturedAgentUB + 1
                    ReDim Preserve goCurrentPlayer.CapturedAgentIdx(lIdx)
                    ReDim Preserve goCurrentPlayer.CapturedAgents(lIdx)
                    goCurrentPlayer.CapturedAgentIdx(lIdx) = -1
                    goCurrentPlayer.CapturedAgentUB = lIdx
                End If
                goCurrentPlayer.CapturedAgents(lIdx) = oCA
                goCurrentPlayer.CapturedAgentIdx(lIdx) = oCA.ObjectID
            End If

            If bNewData = True Then
                If goUILib Is Nothing = False Then
                    goUILib.AddNotification("We have acquired new information from a captured agent!", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    If goUILib Is Nothing = False Then
                        Dim ofrm As frmQuickBar = CType(goUILib.GetWindow("frmQuickBar"), frmQuickBar)
                        If ofrm Is Nothing Then ofrm = New frmQuickBar(goUILib)
                        If ofrm Is Nothing = False Then ofrm.AgentEventUpdate()
                        ofrm = Nothing
                    End If
                End If
            End If
        ElseIf yAction = 1 Then 'executed
            If goCurrentPlayer Is Nothing = False Then
                For X As Int32 = 0 To goCurrentPlayer.CapturedAgentUB
                    If goCurrentPlayer.CapturedAgentIdx(X) = lAgentID Then
                        goCurrentPlayer.CapturedAgentIdx(X) = -1
                        Exit For
                    End If
                Next X
            End If
        ElseIf yAction = 2 Then 'interrogation
            If goCurrentPlayer Is Nothing = False Then
                For X As Int32 = 0 To goCurrentPlayer.CapturedAgentUB
                    If goCurrentPlayer.CapturedAgentIdx(X) = lAgentID Then
                        goCurrentPlayer.CapturedAgents(X).yInterrogationProgress = 1
                        Exit For
                    End If
                Next X
            End If
        Else    'released
            Dim oCA As CapturedAgent = Nothing
            If goCurrentPlayer Is Nothing = False Then
                For X As Int32 = 0 To goCurrentPlayer.CapturedAgentUB
                    If goCurrentPlayer.CapturedAgentIdx(X) = lAgentID Then
                        goCurrentPlayer.CapturedAgentIdx(X) = -1
                        Exit For
                    End If
                Next X
            End If

        End If




    End Sub

    Private Sub HandleRemovePlayerRel(ByRef yData() As Byte)
        msCurrentLoc = "HandleRemovePlayerRel"
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lWithPlayerID As Int32 = System.BitConverter.ToInt32(yData, 6)

        Dim lOtherPlayerID As Int32 = lWithPlayerID
        If lWithPlayerID = glPlayerID Then
            lOtherPlayerID = lPlayerID
        End If

        goCurrentPlayer.RemovePlayerRel(lOtherPlayerID)
    End Sub

    Private Sub HandleRequestPlayerDetailsPlayerMin(ByRef yData() As Byte)
        Static xbAlertedPlayer As Boolean = False
        Dim lPos As Int32 = 2
        Dim lMinID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iMinTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim bDiscovered As Boolean = yData(lPos) <> 0 : lPos += 1
        Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
        Dim bArchive As Boolean = yData(lPos) <> 0 : lPos += 1

        If xbAlertedPlayer = False AndAlso bDiscovered = False AndAlso bArchive = False Then
            xbAlertedPlayer = True
            goUILib.AddNotification("We have discovered a new mineral!", System.Drawing.Color.FromArgb(255, 0, 255, 255), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.DiscoverNewMineral)
            If goSound Is Nothing = False Then
                goSound.StartSound("Game Narrative\DiscoveredMineral.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
            End If
        End If

        For X As Int32 = 0 To glMineralUB
            If glMineralIdx(X) = lMinID Then
                If goMinerals(X).MineralName <> sName Then goMinerals(X).MineralName = sName
                If goMinerals(X).bDiscovered <> bDiscovered Then goMinerals(X).bDiscovered = bDiscovered
                goMinerals(X).bArchived = bArchive
                Return
            End If
        Next X

        Dim oMineral As New Mineral()
        oMineral.ObjectID = lMinID
        oMineral.ObjTypeID = iMinTypeID
        oMineral.bDiscovered = bDiscovered
        oMineral.MineralName = sName
        oMineral.bRequestedProps = False
        oMineral.bArchived = bArchive
        oMineral.bReceivedProps = True

        'SyncLock goMinerals
        ReDim Preserve goMinerals(glMineralUB + 1)
        ReDim Preserve glMineralIdx(glMineralUB + 1)
        goMinerals(glMineralUB + 1) = oMineral
        glMineralIdx(glMineralUB + 1) = oMineral.ObjectID
        glMineralUB += 1
        'End SyncLock
    End Sub

    Private Sub HandleRequestPlayerDetailsAlert(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim yType As Byte = yData(lPos) : lPos += 1
        Dim lValue As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If yType = 1 Then
            'email
            If goUILib Is Nothing = False Then
                frmEmailMain.lCurrentUnreadMessages = lValue
                goUILib.AddNotification("You have " & lValue & " unread messages.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                If lValue > 0 Then
                    Dim oWin As frmQuickBar = CType(goUILib.GetWindow("frmQuickBar"), frmQuickBar)
                    If oWin Is Nothing Then oWin = New frmQuickBar(goUILib)
                    If oWin Is Nothing = False Then oWin.NewEmailHasArrived()
                    oWin = Nothing
                End If
            End If
        End If
    End Sub

    Private Sub HandleRequestEmailSummarys(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lEmailID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEmailTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim sSender As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
        Dim yMsgRead As Byte = yData(lPos) : lPos += 1
        Dim lPCF_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lTitleLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim sMsgTitle As String = ""
        If lTitleLen > 0 Then
            sMsgTitle = GetStringFromBytes(yData, lPos, lTitleLen) : lPos += lTitleLen
        End If

        If goCurrentPlayer Is Nothing Then goCurrentPlayer = New Player()
        Dim oMsg As PlayerComm = New PlayerComm()
        With oMsg
            .bMsgRead = yMsgRead <> 0
            .MsgTitle = sMsgTitle
            .ObjectID = lEmailID
            .ObjTypeID = iEmailTypeID
            .PCF_ID = lPCF_ID
            .sSender = sSender
            .bRequestedDetails = False
        End With
        goCurrentPlayer.AddPlayerComm(oMsg)

    End Sub

#Region "  Request Player Details Items  "
    Private moFSRPD As IO.FileStream = Nothing
    Private moFSRPD_WRITE As IO.StreamWriter = Nothing
    Private Sub WriteToRPDFile(ByVal sLine As String)
        Try
            Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
            If sPath.EndsWith("\") = False Then sPath &= "\"
            moFSRPD = New IO.FileStream(sPath & "RPD.txt", IO.FileMode.Append)
            moFSRPD_WRITE = New IO.StreamWriter(moFSRPD)
            moFSRPD_WRITE.WriteLine(sLine)
        Catch
        Finally
            If moFSRPD_WRITE Is Nothing = False Then moFSRPD_WRITE.Close()
            moFSRPD_WRITE = Nothing
            If moFSRPD Is Nothing = False Then moFSRPD.Close()
            moFSRPD = Nothing
        End Try
    End Sub
    Private Sub KillRPDFile()
        Try
            Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
            If sPath.EndsWith("\") = False Then sPath &= "\"
            If Exists(sPath & "RPD.txt") = True Then Kill(sPath & "RPD.txt")
        Catch
        End Try
    End Sub
    Private mlLastRequestedType As Int32 = 0
    Private mbReRequestInProgress As Boolean = False
    Private mlReRequestCount As Int32 = 0
    Private Sub HandleRequestPlayerDetailsResponse(ByRef yData() As Byte)
        Dim lTypeID As Int32 = System.BitConverter.ToInt32(yData, 6)

        'DoLogEvent("HandleRequestPlayerDetailsResponse: " & lTypeID)

        WriteToRPDFile("  RPD: " & lTypeID)
        Threading.Thread.Sleep(10)

        lTypeID += 1

        mlLastRequestedType = lTypeID

        If lTypeID >= elPlayerDetailsType.eLastPlayerDetailsType Then
            If (mlBusyStatus And elBusyStatusType.PlayerDetails) <> 0 Then
                mlBusyStatus = mlBusyStatus Xor elBusyStatusType.PlayerDetails
            End If
            KillRPDFile()
            RaiseEvent ReceivedAllPlayerDetails()
            Return
        End If

        mlReRequestCount = 0

        mlLastAddObjectID = -1
        miLastAddObjectTypeID = -1

        Dim yMsg(15) As Byte
        msCurrentLoc = "RequestPlayerDetails"
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerDetails).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(lTypeID).CopyTo(yMsg, 6)
        System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 10)
        System.BitConverter.GetBytes(-1S).CopyTo(yMsg, 14)

        moPrimary.SendData(yMsg)
    End Sub
    Private Sub __ThreadedReRequestPD()
        mbReRequestInProgress = True

        mlReRequestCount += 1

        If mlReRequestCount > 20 Then
            goUILib.AddNotification("Unable to log in at this time, please try again.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            goUILib.AddNotification("If the problem persists, send an email to support@darkskyentertainment.com.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            goUILib.AddNotification("In that email, include any BadRPD files in the game directory as a zip attachment.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        Threading.Thread.Sleep(1000)
        Dim yMsg(15) As Byte
        msCurrentLoc = "RequestPlayerDetails"
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerDetails).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(mlLastRequestedType).CopyTo(yMsg, 6)
        System.BitConverter.GetBytes(mlLastAddObjectID).CopyTo(yMsg, 10)
        System.BitConverter.GetBytes(miLastAddObjectTypeID).CopyTo(yMsg, 14)
        moPrimary.SendData(yMsg)
        mbReRequestInProgress = False
    End Sub
#End Region

    Public Sub CheckDXDiagProcess()
        If lDXDiagProcessID <> 0 Then
            If lDXDiagProcessID = -5 Then
                lDXDiagProcessID = 0
            Else
                Try
                    Dim oProc As Process = Process.GetProcessById(lDXDiagProcessID)
                    If oProc Is Nothing Then lDXDiagProcessID = 0
                Catch
                    lDXDiagProcessID = 0
                End Try
            End If

            If lDXDiagProcessID = 0 Then
                Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
                If sPath.EndsWith("\") = False Then sPath &= "\"
                If Exists(sPath & "dxdiag.txt") = True Then
                    SendDXDiag()
                Else
                    lDXDiagProcessID = -5
                End If
            End If
        End If
    End Sub

    Private Sub SendDXDiag()
        'ok, can we open the file?
        Dim oFS As IO.FileStream = Nothing
        Dim oRead As IO.StreamReader = Nothing
        Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        Try
            oFS = New IO.FileStream(sPath & "dxdiag.txt", IO.FileMode.Open)
            oRead = New IO.StreamReader(oFS)
            Dim sFullText As String = ""
            While oRead.EndOfStream() = False
                sFullText = oRead.ReadToEnd()
            End While
            If sFullText <> "" Then
                Dim yDXDiag() As Byte = System.Text.ASCIIEncoding.ASCII.GetBytes(sFullText)
                If yDXDiag Is Nothing = False Then
                    'ok, now, send our data... we need to break it down into chunks... we'll use 8192 as our chunksize
                    Dim lTotalPacks As Int32 = CInt(Math.Ceiling(yDXDiag.Length / 8000))

                    Dim yCache(200000) As Byte
                    Dim yFinal() As Byte
                    Dim lPos As Int32
                    Dim lSingleMsgLen As Int32
                    Dim yTemp() As Byte

                    Dim lFromPos As Int32 = 0

                    'Ok, for each pack
                    For X As Int32 = 0 To lTotalPacks - 1
                        Dim lMsgLen As Int32 = Math.Min(8000, yDXDiag.Length - lFromPos)
                        ReDim yTemp(9 + lMsgLen)
                        System.BitConverter.GetBytes(GlobalMessageCode.eRequestDXDiag).CopyTo(yTemp, 0)
                        System.BitConverter.GetBytes(CInt(X + 1)).CopyTo(yTemp, 2)
                        System.BitConverter.GetBytes(lTotalPacks).CopyTo(yTemp, 6)
                        Array.Copy(yDXDiag, lFromPos, yTemp, 10, lMsgLen)
                        lFromPos += lMsgLen

                        lSingleMsgLen = yTemp.Length
                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                            ReDim Preserve yCache(yCache.Length + 200000)
                        End If
                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                        lPos += 2
                        yTemp.CopyTo(yCache, lPos)
                        lPos += lSingleMsgLen
                    Next X

                    'Now, send it all...
                    If lPos <> 0 Then
                        ReDim yFinal(lPos - 1)
                        Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
                        moPrimary.SendLenAppendedData(yFinal)
                    End If
                End If
            End If
        Catch
            'do nothing for now
        Finally
            If oRead Is Nothing = False Then oRead.Close()
            oRead = Nothing
            If oFS Is Nothing = False Then oFS.Close()
            oFS = Nothing
        End Try
    End Sub

    Public lDXDiagProcessID As Int32 = 0
    Private Sub HandleRequestDXDiag(ByRef yData() As Byte)
        'ok, we need to get our dx diag file...
        Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        If sPath.EndsWith("\") = False Then sPath &= "\"
        If Exists(sPath & "dxdiag.txt") = False Then
            'ok, let's do a silent startup of dxdiag
            Try
                lDXDiagProcessID = Shell("dxdiag /t " & sPath & "dxdiag.txt", AppWinStyle.Hide, False, -1)
            Catch           'do nothing
            End Try
        Else
            SendDXDiag()
        End If

    End Sub

    Private Sub HandleGetRouteList(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lUnitID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yValue As Byte = yData(lPos) : lPos += 1
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim uRoute(lCnt - 1) As RouteItem
        For X As Int32 = 0 To lCnt - 1
            With uRoute(X)
                .FillFromMsg(yData, lPos)
            End With
        Next X

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing = False Then
            For X As Int32 = 0 To oEnvir.lEntityUB
                If oEnvir.lEntityIdx(X) = lUnitID AndAlso oEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit Then
                    Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                    If oEntity Is Nothing = False Then
                        oEntity.bRoutePaused = yValue <> 0
                        oEntity.uRoute = uRoute
                        oEntity.lRouteUB = lCnt - 1
                    End If
                    Exit For
                End If
            Next X
        End If
    End Sub

    Private Sub HandleSetRouteMineral(ByRef yData() As Byte)
        Dim lPos As Int32 = 2 'formsgcode
        Dim lUnitID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lRouteIdx As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lMineralID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iMineralTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim yFlag As Byte = yData(lPos) : lPos += 1

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return

        Try
            For X As Int32 = 0 To oEnvir.lEntityUB
                If oEnvir.lEntityIdx(X) = lUnitID AndAlso oEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit Then
                    Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                    If oEntity Is Nothing = False Then
                        If oEntity.lRouteUB >= lRouteIdx Then
                            With oEntity.uRoute(lRouteIdx)
                                If .lDestID = lDestID AndAlso .iDestTypeID = iDestTypeID Then
                                    .lLoadItemID = lMineralID
                                    .iLoadItemTypeID = iMineralTypeID
                                    .yFlag = yFlag
                                End If
                            End With
                        End If
                    End If
                    Exit For
                End If
            Next X
        Catch
        End Try
    End Sub

    Private Sub HandleUpdateRouteStatus(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lUnitID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lValue As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return

        Try
            Dim lCurUB As Int32 = -1
            If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = (Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0)))
            For X As Int32 = 0 To lCurUB
                If oEnvir.lEntityIdx(X) = lUnitID AndAlso oEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit Then
                    Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                    If oEntity Is Nothing = False Then
                        If lValue = Int32.MaxValue OrElse lValue = -4 Then
                            'begin
                            If goUILib Is Nothing = False Then goUILib.AddNotification("Route begun.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        ElseIf lValue = Int32.MinValue Then
                            'force next... do nothing
                        ElseIf lValue = -2 Then
                            'paused
                            oEntity.bRoutePaused = True
                        ElseIf lValue = -3 Then
                            'unpaused
                            oEntity.bRoutePaused = False
                        Else
                            'route location add/update
                            Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                            Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                            Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                            Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                            'Now... update our unit
                            If oEntity.lRouteUB >= lValue Then
                                With oEntity.uRoute(lValue)
                                    .lDestID = lEnvirID
                                    .iDestTypeID = iEnvirTypeID
                                    .lLocX = lLocX
                                    .lLocZ = lLocZ

                                    If iEnvirTypeID = ObjectType.eFacility Then
                                        For Y As Int32 = 0 To lCurUB
                                            If oEnvir.lEntityIdx(Y) = .lDestID AndAlso oEnvir.oEntity(Y).ObjTypeID = .iDestTypeID Then
                                                If oEnvir.oEntity(Y).yProductionType = ProductionType.eMining Then
                                                    .lLoadItemID = Int32.MinValue
                                                End If
                                                Exit For
                                            End If
                                        Next
                                    End If
                                End With
                            Else
                                Dim uNewItem As RouteItem
                                With uNewItem
                                    .lLocX = lLocX
                                    .lLocZ = lLocZ
                                    .lDestID = lEnvirID
                                    .iDestTypeID = iEnvirTypeID

                                    If iEnvirTypeID = ObjectType.eFacility Then
                                        For Y As Int32 = 0 To lCurUB
                                            If oEnvir.lEntityIdx(Y) = .lDestID AndAlso oEnvir.oEntity(Y).ObjTypeID = .iDestTypeID Then
                                                If oEnvir.oEntity(Y).yProductionType = ProductionType.eMining Then
                                                    .lLoadItemID = Int32.MinValue
                                                End If
                                                Exit For
                                            End If
                                        Next
                                    End If
                                End With
                                oEntity.lRouteUB += 1
                                ReDim Preserve oEntity.uRoute(oEntity.lRouteUB)
                                oEntity.uRoute(oEntity.lRouteUB) = uNewItem
                            End If
                        End If
                    End If
                    Exit For
                End If
            Next X
        Catch
        End Try
    End Sub

    Private Sub HandleRemoveRouteItem(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lUnitID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lValue As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return

        Try
            For X As Int32 = 0 To oEnvir.lEntityUB
                If oEnvir.lEntityIdx(X) = lUnitID AndAlso oEnvir.oEntity(X).ObjTypeID = ObjectType.eUnit Then
                    Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                    If oEntity Is Nothing = False Then
                        If lValue = Int32.MinValue Then
                            oEntity.lRouteUB = -1
                        Else
                            For Y As Int32 = lValue To oEntity.lRouteUB - 1
                                oEntity.uRoute(Y) = oEntity.uRoute(Y + 1)
                            Next Y
                            oEntity.lRouteUB -= 1
                        End If
                    End If
                    Exit For
                End If
            Next X
        Catch
        End Try
    End Sub

    Private Sub HandleGetSkillList(ByRef yData() As Byte)
        Dim lPos As Int32 = 2 'msgcode
        Dim lAgentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If goCurrentPlayer Is Nothing = False Then
            For X As Int32 = 0 To goCurrentPlayer.AgentUB
                If goCurrentPlayer.AgentIdx(X) = lAgentID Then
                    goCurrentPlayer.Agents(X).HandleSkillList(yData)
                    Exit For
                End If
            Next X
        End If
    End Sub

    Private Sub HandleAddFormation(ByRef yData() As Byte)
        Dim oFormation As New FormationDef()
        oFormation.FillFromMsg(yData)

        'Now, place it...
        If oFormation.lOwnerID = glPlayerID Then
            If goCurrentPlayer Is Nothing = False Then
                For X As Int32 = 0 To goCurrentPlayer.lFormationUB
                    If goCurrentPlayer.lFormationIdx(X) = oFormation.FormationID Then
                        goCurrentPlayer.oFormations(X) = oFormation
                        Return
                    End If
                Next X
                goCurrentPlayer.lFormationUB += 1
                ReDim Preserve goCurrentPlayer.lFormationIdx(goCurrentPlayer.lFormationUB)
                ReDim Preserve goCurrentPlayer.oFormations(goCurrentPlayer.lFormationUB)
                goCurrentPlayer.oFormations(goCurrentPlayer.lFormationUB) = oFormation
                goCurrentPlayer.lFormationIdx(goCurrentPlayer.lFormationUB) = oFormation.FormationID
            End If
        End If
    End Sub

    Private Sub HandleRemoveFormation(ByRef yData() As Byte)
        Dim lFormationID As Int32 = System.BitConverter.ToInt32(yData, 2)
        If goCurrentPlayer Is Nothing = False Then
            For X As Int32 = 0 To goCurrentPlayer.lFormationUB
                If goCurrentPlayer.lFormationIdx(X) = lFormationID Then
                    goCurrentPlayer.lFormationIdx(X) = -1
                End If
            Next X
        End If
    End Sub

    Private Sub HandleMoveFormation(ByRef yData() As Byte)
        Dim lPos As Int32 = 2 ' for msgcode 
        Dim yMaxSpeed As Byte = yData(lPos) : lPos += 1
        Dim yManeuver As Byte = yData(lPos) : lPos += 1
        Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim iDestA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return

        Dim fAcc As Single = yManeuver / 100.0F

        For X As Int32 = 0 To lCnt - 1
            Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            For Y As Int32 = 0 To oEnvir.lEntityUB
                If oEnvir.lEntityIdx(Y) = lObjID AndAlso oEnvir.oEntity(Y).ObjTypeID = iObjTypeID Then
                    With oEnvir.oEntity(Y)
                        If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 AndAlso .OwnerID = glPlayerID AndAlso goSound Is Nothing = False Then
                            Dim lVal As Int32 = CInt((Rnd() * 8) + 1)
                            goSound.StartSound("RC" & lVal & ".wav", False, SoundMgr.SoundUsage.eRadioChatter, New Microsoft.DirectX.Vector3(.LocX, .LocY, .LocZ), Microsoft.DirectX.Vector3.Empty)
                        End If
                        .oUnitDef.yFormationManeuver = yManeuver
                        .oUnitDef.yFormationMaxSpeed = yMaxSpeed
                        .oUnitDef.fFormationAcceleration = fAcc
                        .SetDest(lDestX, lDestZ, iDestA)
                        .yChangeEnvironment = 0
                        .bDoNotRunSetYTarget = False
                        .LastUpdateCycle = glCurrentCycle
                    End With
                End If
            Next Y
        Next X
    End Sub

    Private Sub HandleSetIronCurtain(ByRef yData() As Byte)
        Dim lPos As Int32 = 2 'for msgcode
        Dim lPlanetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yValue As Byte = yData(lPos) : lPos += 1

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing = False Then
            If oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oEnvir.ObjectID = lPlanetID Then
                Dim lIdx As Int32 = -1
                For X As Int32 = 0 To oEnvir.lIronCurtainPlayerUB
                    If oEnvir.lIronCurtainPlayers(X) = lPlayerID Then
                        If yValue = 0 Then
                            oEnvir.lIronCurtainPlayers(X) = -1
                        Else : Return
                        End If
                    ElseIf oEnvir.lIronCurtainPlayers(X) = -1 AndAlso lIdx = -1 Then
                        lIdx = X
                    End If
                Next X
                If yValue <> 0 Then
                    If lIdx = -1 Then
                        ReDim Preserve oEnvir.lIronCurtainPlayers(oEnvir.lIronCurtainPlayerUB + 1)
                        oEnvir.lIronCurtainPlayers(oEnvir.lIronCurtainPlayerUB + 1) = lPlayerID
                        oEnvir.lIronCurtainPlayerUB += 1
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub HandleUpdateSlotStates(ByRef yData() As Byte)
        Dim lPos As Int32 = 2 'fopr msgcode
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If glPlayerID = lPlayerID Then
            goCurrentPlayer.HandleUpdateSlotMsg(yData)
            'Else
            '	'TODO: Espionage?
        End If
    End Sub

    Private Sub HandleUpdateResearchCnt(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lTechID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTechTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim yCnt As Byte = yData(lPos) : lPos += 1

        If goCurrentPlayer Is Nothing = False Then
            Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lTechID, iTechTypeID)
            If oTech Is Nothing = False Then oTech.Researchers = yCnt
        End If
    End Sub
    Private Sub HandleGetItemIntelDetail(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lItemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iItemTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lOtherPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        For X As Int32 = 0 To glItemIntelUB
            If glItemIntelIdx(X) <> -1 Then
                Dim oPII As PlayerItemIntel = goItemIntel(X)
                If oPII Is Nothing = False Then
                    If oPII.lItemID = lItemID AndAlso oPII.iItemTypeID = iItemTypeID AndAlso oPII.lOtherPlayerID = lOtherPlayerID Then
                        oPII.FillDetails(yData, lPos)
                        Exit For
                    End If
                End If
            End If
        Next X
    End Sub
    Private Sub HandleAlertDestArrived(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20

        frmSingleSelect.UnitAlertReceived(lObjID, iObjTypeID)

        If goUILib Is Nothing = False Then
            goUILib.AddNotification(sName & " has arrived at its destination.", Color.Green, lEnvirID, iEnvirTypeID, lObjID, iObjTypeID, lLocX, lLocZ)
            If goSound Is Nothing = False Then
                goSound.StartSound("UserInterface\Alert.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            End If
        End If
    End Sub

    Private Sub HandleGetIntelSellOrderDetail(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lTP_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4  'unused...
        Dim lItemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim iExtTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim sPlayerName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20

        Dim oSB As New System.Text.StringBuilder
        Select Case iTypeID
            Case ObjectType.ePlayerIntel
                oSB.AppendLine("Stats for " & sPlayerName)
                'ok, make our string
                Dim lDip As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lMil As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lPop As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lProd As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lTech As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lWealth As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                If lDip > 0 Then
                    oSB.AppendLine("Diplomacy Score (> " & GetDurationFromSeconds(lDip, True) & ")")
                End If
                If lMil > 0 Then
                    oSB.AppendLine("Military Score (> " & GetDurationFromSeconds(lMil, True) & ")")
                End If
                If lPop > 0 Then
                    oSB.AppendLine("Population Score (> " & GetDurationFromSeconds(lPop, True) & ")")
                End If
                If lProd > 0 Then
                    oSB.AppendLine("Production Score (> " & GetDurationFromSeconds(lProd, True) & ")")
                End If
                If lTech > 0 Then
                    oSB.AppendLine("Technology Score (> " & GetDurationFromSeconds(lTech, True) & ")")
                End If
                If lWealth > 0 Then
                    oSB.AppendLine("Wealth Score (> " & GetDurationFromSeconds(lWealth, True) & ")")
                End If
            Case ObjectType.ePlayerItemIntel
                Dim yIntelType As Byte = yData(lPos) : lPos += 1
                Dim lAge As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim sItemName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20

                Select Case yIntelType
                    Case PlayerItemIntel.PlayerItemIntelType.eStatus
                        oSB.AppendLine("Location and status of " & sItemName)
                    Case PlayerItemIntel.PlayerItemIntelType.eFullKnowledge
                        oSB.AppendLine("Full knowledge of " & sItemName)
                    Case Else
                        oSB.AppendLine("Location of " & sItemName)
                End Select
                oSB.AppendLine(vbCrLf & "Owner: " & sPlayerName)
                oSB.AppendLine(vbCrLf & "Data is " & GetDurationFromSeconds(lAge, True) & " old")
            Case ObjectType.ePlayerTechKnowledge
                Dim yTechKnowLvl As Byte = yData(lPos) : lPos += 1
                Dim sTechName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20

                Select Case yTechKnowLvl
                    Case PlayerTechKnowledge.KnowledgeType.eSettingsLevel1
                        oSB.AppendLine("Technical readouts for " & sTechName & " belonging to " & sPlayerName & ".")
                    Case PlayerTechKnowledge.KnowledgeType.eSettingsLevel2
                        oSB.AppendLine("Detailed Schematics for " & sTechName & " belonging to " & sPlayerName & ".")
                    Case PlayerTechKnowledge.KnowledgeType.eFullKnowledge
                        oSB.AppendLine("All documents on " & sTechName & " belonging to " & sPlayerName & ".")
                End Select

                oSB.AppendLine(vbCrLf & "Component Type: " & Base_Tech.GetComponentTypeName(iExtTypeID))
        End Select

        SetNonOwnerIntelItemData(lItemID, iTypeID, iExtTypeID, oSB.ToString)
    End Sub

    Private mbArchivedRequested As Boolean = False
    Public Sub LoadArchived()
        If mbArchivedRequested = True Then Return
        mbArchivedRequested = True
        Dim yMsg(1) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eGetArchivedItems).CopyTo(yMsg, 0)
        moPrimary.SendData(yMsg)
    End Sub

    Private Sub HandleAgentMissionCompleted(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPM_ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMissionID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yPhase As eMissionPhase = CType(yData(lPos), eMissionPhase) : lPos += 1

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eMissionCompleted, lMissionID, -1, -1, "")
        End If
        If goUILib Is Nothing = False Then
            If yPhase = eMissionPhase.eMissionOverSuccess Then
                goUILib.AddNotification("Your agency is reporting a successful mission!", System.Drawing.Color.FromArgb(255, 0, 255, 0), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            ElseIf yPhase = eMissionPhase.eMissionOverFailure Then
                goUILib.AddNotification("Your agency is reporting an unsuccessful mission!", System.Drawing.Color.FromArgb(255, 255, 0, 0), -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
        End If
    End Sub

    Private Sub HandleUpdatePlayerTimer(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim yType As Byte = yData(lPos) : lPos += 1
        Dim lSeconds As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim dtEndTime As Date = Now.Add(New TimeSpan(0, 0, lSeconds))

        frmEnvirDisplay.dtEndTimer = dtEndTime
    End Sub

    Public Sub HandlePirateWaveSpawn(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lWave As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim clrVal As System.Drawing.Color
        If lWave < 3 Then
            clrVal = System.Drawing.Color.FromArgb(0, 255, 255, 0)
        ElseIf lWave < 8 Then
            clrVal = System.Drawing.Color.FromArgb(0, 255, 128, 0)
        Else : clrVal = System.Drawing.Color.FromArgb(0, 255, 0, 0)
        End If

        frmEnvirDisplay.clrPirateWave = clrVal
        frmEnvirDisplay.mlPirateWave = lWave

        If goSound Is Nothing = False Then goSound.IncrementExcitementLevel(1000)

        frmEnvirDisplay.dtEndTimer = Now.Add(New TimeSpan(0, (11 - lWave), 0))
    End Sub

    Private Sub HandleUpdateMOTD(ByRef yData() As Byte)
        Dim sMOTD As String = GetStringFromBytes(yData, 2, 200)
        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then goCurrentPlayer.oGuild.sMOTD = sMOTD
        If goUILib Is Nothing = False Then
            Dim oTmpWin As frmChat = CType(goUILib.GetWindow("frmChat"), frmChat)

            If oTmpWin Is Nothing = False Then
                oTmpWin.AddChatMessage(-1, "Guild MOTD:" & sMOTD, ChatMessageType.eNotificationMessage, Date.MinValue, True)
            End If
            oTmpWin = Nothing
        End If
    End Sub
    Private Sub HandleUpdateGuildRank(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lRankID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yChanges As Byte = yData(lPos) : lPos += 1

        If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return

        Dim oRank As GuildRank = goCurrentPlayer.oGuild.GetRankByID(lRankID)
        If oRank Is Nothing = False Then
            If (yChanges And 1) <> 0 Then
                Dim yMoveDir As Byte = yData(lPos) : lPos += 1
                If yMoveDir = 255 Then
                    'remove it
                    goCurrentPlayer.oGuild.RemoveRank(lRankID)
                Else
                    'move the rank
                    goCurrentPlayer.oGuild.MoveRank(lRankID, yMoveDir)
                End If
            End If
            If (yChanges And 2) <> 0 Then
                oRank.sRankName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
            End If
            If (yChanges And 4) <> 0 Then
                oRank.lVoteStrength = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            End If
            If (yChanges And 8) <> 0 Then
                oRank.TaxRateFlat = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                oRank.TaxRatePercentage = yData(lPos) : lPos += 1
                oRank.TaxRatePercType = CType(yData(lPos), eyGuildTaxPercType) : lPos += 1
            End If
            If goUILib Is Nothing = False Then
                goUILib.AddNotification("The Rank of " & oRank.sRankName & " has new settings.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
        End If

    End Sub
    Private Sub HandleUpdateRankPermission(ByRef yData() As Byte)
        Dim lPos As Int32 = 2    'for msgcode
        Dim lRankID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lValueChangeCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return

        Dim oRank As GuildRank = goCurrentPlayer.oGuild.GetRankByID(lRankID)
        If oRank Is Nothing = False Then
            For X As Int32 = 0 To lValueChangeCnt - 1
                Dim lValue As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                If lValue < 0 Then
                    'ok, this means we are removing the value
                    lValue = Math.Abs(lValue)
                    If (oRank.lRankPermissions And lValue) <> 0 Then oRank.lRankPermissions = CType(oRank.lRankPermissions Xor lValue, RankPermissions)
                Else
                    'ok, this means we are adding the value
                    oRank.lRankPermissions = CType(oRank.lRankPermissions Or lValue, RankPermissions)
                End If
            Next X

            If goUILib Is Nothing = False Then
                goUILib.AddNotification("The Rank of " & oRank.sRankName & " has had its permissions updated.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
        End If
    End Sub
    Private Sub HandleGuildMemberStatus(ByRef yData() As Byte)
        Dim lPos As Int32 = 2    'for msgcode
        Dim lGuildID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMemberID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lStatusUpdate As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return

        Dim oGuild As Guild = goCurrentPlayer.oGuild
        If oGuild Is Nothing Then Return

        'this is more of an action, the server is TELLING me what to do... so we know there is a member?

        Select Case lStatusUpdate
            Case -1
                'demote the specified member
                If glPlayerID = lMemberID Then
                    Dim oNextRank As GuildRank = oGuild.GetNextRankPosition(1, oGuild.lCurrentRankID)
                    If oNextRank Is Nothing = False Then
                        'oGuild.lCurrentRankID = oNextRank.lRankID
                        If goUILib Is Nothing = False Then goUILib.AddNotification("You have been demoted in rank.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    End If
                End If

                Dim oMember As GuildMember = oGuild.GetMember(lMemberID)
                If oMember Is Nothing = False Then
                    Dim oNextRank As GuildRank = oGuild.GetNextRankPosition(1, oMember.lRankID)
                    If oNextRank Is Nothing = False Then
                        oMember.lRankID = oNextRank.lRankID
                    End If
                End If
            Case -2
                'promote the specified member
                If glPlayerID = lMemberID Then
                    Dim oNextRank As GuildRank = oGuild.GetNextRankPosition(-1, oGuild.lCurrentRankID)
                    If oNextRank Is Nothing = False Then
                        'oGuild.lCurrentRankID = oNextRank.lRankID
                        If goUILib Is Nothing = False Then goUILib.AddNotification("You have been promoted in rank.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    End If
                End If

                Dim oMember As GuildMember = oGuild.GetMember(lMemberID)
                If oMember Is Nothing = False Then
                    Dim oNextRank As GuildRank = oGuild.GetNextRankPosition(-1, oMember.lRankID)
                    If oNextRank Is Nothing = False Then
                        oMember.lRankID = oNextRank.lRankID
                    End If
                End If
            Case Int32.MinValue
                'remove the specified player or reject the player if they are applying
                'Ok, determine what we are doing... if the target is not a member of the guild, we want the "Reject" member
                ' otherwise, we want to "Remove" the member
                Dim oMember As GuildMember = oGuild.GetMember(lMemberID)
                If oMember Is Nothing = False Then
                    If (oMember.yMemberState And GuildMemberState.Approved) <> 0 Then
                        'ok, member is already a member...
                        oGuild.RemoveMember(lMemberID)
                        If lMemberID = glPlayerID Then
                            If goUILib Is Nothing = False Then goUILib.AddNotification("You have been kicked from the guild!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            goCurrentPlayer.oGuild = Nothing
                        End If
                    Else
                        'ok, member is applied or invited...
                        If (oMember.yMemberState And GuildMemberState.Applied) <> 0 Then
                            oGuild.RemoveMember(lMemberID)
                        ElseIf (oMember.yMemberState And GuildMemberState.Invited) <> 0 Then
                            oGuild.RemoveMember(lMemberID)
                        End If
                    End If
                End If
                oGuild.RecheckGuildMembers()
        End Select

    End Sub
    Private Sub HandleProposeGuildVote(ByRef yData() As Byte)
        If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return

        Dim lPos As Int32 = 2
        Dim lVoteID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lProposedByID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lVoteStarts As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lDuration As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yTypeOfVote As Byte = yData(lPos) : lPos += 1
        Dim lSelectedItem As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lNewValueNumber As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim sNewValueText As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
        Dim sSummary As String = GetStringFromBytes(yData, lPos, 255) : lPos += 255
        Dim yVoteState As Byte = yData(lPos) : lPos += 1

        Dim oVote As GuildVote = goCurrentPlayer.oGuild.GetVote(lVoteID)
        Dim bAdded As Boolean = False
        If oVote Is Nothing Then
            oVote = New GuildVote()
            bAdded = True
        End If
        With oVote
            .dtVoteStarts = Date.SpecifyKind(GetDateFromNumber(lVoteStarts), DateTimeKind.Utc)
            .lNewValue = lNewValueNumber
            .lSelectedItem = lSelectedItem
            .lVoteDuration = lDuration
            .ProposedByID = lProposedByID
            .sSummary = sSummary
            .VoteID = lVoteID
            .yTypeOfVote = yTypeOfVote
            .sNewValueText = sNewValueText
            .yVoteState = yVoteState
        End With
        If bAdded = True Then goCurrentPlayer.oGuild.AddVote(oVote)

        If lProposedByID = glPlayerID Then
            If goUILib Is Nothing = False Then goUILib.AddNotification("Vote Proposal Details Accepted.", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        Else
            'If goUILib Is Nothing = False Then goUILib.AddNotification("New Vote Proposal Details Available.", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub
    Private Sub HandleAddEventAttachment(ByRef yData() As Byte)
        If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return

        Dim lPos As Int32 = 2    'for msgcode
        Dim lEventID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oEvent As GuildEvent = goCurrentPlayer.oGuild.GetEvent(lEventID)
        If oEvent Is Nothing Then Return

        Dim yAttachPos As Byte = yData(lPos) : lPos += 1

        Dim oSketchPad As SketchPad = Nothing
        If oEvent.AttachmentUB >= yAttachPos Then
            oSketchPad = oEvent.Attachments(yAttachPos)
        End If
        Dim bAdded As Boolean = False
        If oSketchPad Is Nothing Then
            oSketchPad = New SketchPad
            bAdded = True
        End If
        With oSketchPad
            .sName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
            .lEnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .iEnvirTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .ViewID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CameraX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CameraY = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CameraZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CameraAtX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CameraAtY = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CameraAtZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            Dim lItemCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            ReDim .uItems(lItemCnt - 1)

            For X As Int32 = 0 To lItemCnt - 1
                With oSketchPad.uItems(X)
                    .yType = CType(yData(lPos), SketchPad.eySketchShapes) : lPos += 1
                    .yClrVal = yData(lPos) : lPos += 1
                    .fPtA_X = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
                    .fPtA_Y = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
                    .fPtB_X = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
                    .fPtB_Y = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
                    If .yType = SketchPad.eySketchShapes.Text Then
                        Dim lLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .sText = GetStringFromBytes(yData, lPos, lLen) : lPos += lLen
                    End If
                End With
            Next X
        End With
        If bAdded = True Then
            If yAttachPos > oEvent.AttachmentUB Then
                oEvent.AttachmentUB = yAttachPos
                ReDim Preserve oEvent.Attachments(oEvent.AttachmentUB)
                oEvent.Attachments(yAttachPos) = oSketchPad
            End If
        End If
    End Sub
    Private Sub HandleUpdateGuildRecruitment(ByRef yData() As Byte)
        If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return
        Dim lPos As Int32 = 2   'for msgcode
        Dim iFlags As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lBillboardLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        goCurrentPlayer.oGuild.iRecruitFlags = CType(iFlags, eiRecruitmentFlags)
        goCurrentPlayer.oGuild.sBillboard = GetStringFromBytes(yData, lPos, lBillboardLen)
    End Sub
    Private Sub HandleRemoveEventAttachment(ByRef yData() As Byte)
        If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return

        Dim lPos As Int32 = 2   'for msgcode
        Dim lEventID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oEvent As GuildEvent = goCurrentPlayer.oGuild.GetEvent(lEventID)
        If oEvent Is Nothing Then Return

        For X As Int32 = 0 To oEvent.AttachmentUB
            If oEvent.Attachments(X) Is Nothing = False AndAlso oEvent.Attachments(X).lID = lID Then
                oEvent.Attachments(X) = Nothing
            End If
        Next X

    End Sub
    Private Sub HandleUpdateGuildEvent(ByRef yData() As Byte)
        'two types of messages here, the update acceptance and full update...
        If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return

        Dim lPos As Int32 = 2    'for msgcode
        Dim lEventID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If yData.Length > 12 Then
            'larger version, are we updating or creating?
            Dim oEvent As GuildEvent = Nothing
            Dim bAdded As Boolean = False
            If lEventID > 0 Then
                'updating
                oEvent = goCurrentPlayer.oGuild.GetEvent(lEventID)
            End If
            If oEvent Is Nothing Then
                oEvent = New GuildEvent()
                bAdded = True
            End If
            If oEvent Is Nothing = False Then
                With oEvent
                    .EventID = lEventID
                    .sTitle = GetStringFromBytes(yData, lPos, 50) : lPos += 50
                    .dtStartsAt = Date.SpecifyKind(GetDateFromNumber(System.BitConverter.ToInt32(yData, lPos)), DateTimeKind.Utc) : lPos += 4
                    .lDuration = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .yEventType = yData(lPos) : lPos += 1
                    .ySendAlerts = yData(lPos) : lPos += 1
                    .yEventIcon = yData(lPos) : lPos += 1
                    .yMembersCanAccept = yData(lPos) : lPos += 1
                    .yRecurrence = yData(lPos) : lPos += 1
                    Dim lLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .sDetails = GetStringFromBytes(yData, lPos, lLen) : lPos += lLen
                    .lPostedBy = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .dtPostedOn = Date.SpecifyKind(GetDateFromNumber(System.BitConverter.ToInt32(yData, lPos)), DateTimeKind.Utc) : lPos += 4
                End With
                If bAdded = True Then goCurrentPlayer.oGuild.AddEvent(oEvent)
            End If
        Else
            'smaller version... event is always present
            Dim oEvent As GuildEvent = goCurrentPlayer.oGuild.GetEvent(lEventID)
            If oEvent Is Nothing Then Return

            Dim yAcceptance As Byte = yData(lPos) : lPos += 1
            'Ok, I am setting my acceptance
            If yAcceptance = 255 Then
                goCurrentPlayer.oGuild.RemoveEvent(lEventID)
            Else
                Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                oEvent.SetPlayerAcceptance(lPlayerID, yAcceptance)
            End If
        End If

    End Sub
    Private Sub HandleSetGuildRel(ByRef yData() As Byte)
        'Dim lPos As Int32 = 2    'for msgcode
        'Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        'Dim yRel As Byte = yData(lPos) : lPos += 1

        'If iEntityTypeID = ObjectType.ePlayer AndAlso lEntityID = glPlayerID Then
        '    'TODO: a guild's rel with me
        '    'yRel is guild towards me
        '    Dim yRelTowardsGuild As Byte = yData(lPos) : lPos += 1
        '    'Dim lGuildID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        '    'Dim iGuildTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        'Else
        '    If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
        '        'is the rel with my guild?
        '        If goCurrentPlayer.oGuild.ObjectID = lEntityID AndAlso iEntityTypeID = ObjectType.eGuild Then
        '            Dim yRelTowardsGuild As Byte = yData(lPos) : lPos += 1
        '            Dim lGuildID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        '            Dim iGuildTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        '        Else
        '            Dim oRel As GuildRel = goCurrentPlayer.oGuild.GetRel(lEntityID, iEntityTypeID)
        '            If oRel Is Nothing = False Then
        '                oRel.yRelTowardsThem = yRel
        '            End If
        '        End If
        '    End If
        'End If
    End Sub
    Private Sub HandleAdvancedEventConfig(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim yType As Byte = yData(lPos)

        '1 is race config
        '2 is tourn config
        If yType = 1 Then
            If goUILib Is Nothing = False Then
                Dim ofrm As frmRaceConfig = CType(goUILib.GetWindow("frmRaceConfig"), frmRaceConfig)
                If ofrm Is Nothing = False Then
                    ofrm.HandleAdvancedEventConfig(yData)
                End If
            End If
        ElseIf yType = 2 Then
            If goUILib Is Nothing = False Then
                Dim ofrm As frmTournament = CType(goUILib.GetWindow("frmTournament"), frmTournament)
                If ofrm Is Nothing = False Then ofrm.HandleAdvancedEventConfig(yData)
            End If
        End If


    End Sub
    Private Sub HandleCheckGuildName(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim sName As String = GetStringFromBytes(yData, lPos, 50) : lPos += 50
        Dim yValue As Byte = yData(lPos) : lPos += 1

        If goUILib Is Nothing = False Then
            If yValue = 0 Then
                goUILib.AddNotification(sName & " is available to use.", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Else
                goUILib.AddNotification(sName & " is already in use.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
        End If

    End Sub
    Private Sub HandleInviteFormGuild(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim yType As Byte = yData(lPos) : lPos += 1

        If yType = 1 Then
            'inviting
            Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
            Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim ofrm As New frmInviteFormGuild(goUILib)
            ofrm.SetFromRequest(sName, lPlayerID)
            ofrm.Visible = True
        ElseIf yType = 128 Then
            goUILib.RemoveWindow("frmInviteFormGuild")
        Else
            Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
            Dim lPlayerID As Int32 = -1
            If yType = 2 Then
                lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            End If
            Dim ofrm As frmGuildSetup = CType(goUILib.GetWindow("frmGuildSetup"), frmGuildSetup)
            If ofrm Is Nothing = False Then
                ofrm.PlayerInviteResponse(yType, sName, lPlayerID)
            End If
        End If
    End Sub
    Private Sub HandleInvitePlayerToGuild(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        'Dim yMsg(18) As Byte			'NOTE: Changing this len means changing the length in guild
        'Dim lPos As Int32 = 0
        Dim lMemberID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yMemberState As Byte = yData(lPos) : lPos += 1
        Dim lRankID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yPlayerTitle As Byte = yData(lPos) : lPos += 1
        Dim yGender As Byte = yData(lPos) : lPos += 1
        Dim lLastLogin As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lJoinedOn As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return

        Dim oMember As GuildMember = goCurrentPlayer.oGuild.GetMember(lMemberID)
        Dim bAdded As Boolean = False
        If oMember Is Nothing Then
            oMember = New GuildMember()
            bAdded = True
        End If
        With oMember
            .dtJoined = GetDateFromNumber(lJoinedOn)
            .dtLastOnline = GetDateFromNumber(lLastLogin)
            .lMemberID = lMemberID
            .lRankID = -1
            .yMemberState = CType(yMemberState, GuildMemberState)
            .yPlayerGender = yGender
            .yPlayerTitle = yPlayerTitle
        End With
        If bAdded = True Then goCurrentPlayer.oGuild.AddMember(oMember)

    End Sub
    Private Sub HandleGetMyVoteValue(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim yType As Byte = yData(lPos) : lPos += 1
        If yType = 0 Then
            If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return

            'guild
            Dim lVoteID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim yVote As eyVoteValue = CType(yData(lPos), eyVoteValue) : lPos += 1
            Dim oVote As GuildVote = goCurrentPlayer.oGuild.GetVote(lVoteID)
            If oVote Is Nothing = False Then oVote.yPlayerVote = yVote
        Else
            'senate
        End If
    End Sub
    Private Sub HandleRequestGuildEvents(ByRef yData() As Byte)
        If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return

        Dim lPos As Int32 = 2    'for msgcode
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        For X As Int32 = 0 To lCnt - 1
            Dim lEventID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim yIcon As Byte = yData(lPos) : lPos += 1
            Dim lStartsAt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            Dim oEvent As GuildEvent = goCurrentPlayer.oGuild.GetEvent(lEventID)
            Dim bAdded As Boolean = False
            If oEvent Is Nothing Then
                bAdded = True
                oEvent = New GuildEvent()
            End If
            With oEvent
                .EventID = lEventID
                .yEventIcon = yIcon
                .dtStartsAt = GetDateFromNumber(lStartsAt)
            End With
            If bAdded = True Then goCurrentPlayer.oGuild.AddEvent(oEvent)
        Next X
    End Sub
    Private Sub HandleUpdateGuildTreasury(ByRef yData() As Byte)
        If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return
        Dim lPos As Int32 = 2
        If yData.Length > 11 Then
            frmGuildMain.fra_Finances.HandleUpdateGuildTreasury(yData)
        Else
            goCurrentPlayer.oGuild.blTreasury = System.BitConverter.ToInt64(yData, lPos)
        End If

    End Sub

    Private Sub HandleShiftClickAddProduction(ByVal yData() As Byte)
        'ok, our production was added properly
        Try

            Dim oUnitProdQueue As New UnitProdQueue()
            With oUnitProdQueue
                Dim lPos As Int32 = 2
                .lBuilderID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .iBuilderTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .lProdID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .iProdTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .lLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .iLocA = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .iModelID = 0
                .oModel = Nothing
            End With

            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing = False Then
                oEnvir.AddUnitProdQueueItem(oUnitProdQueue)
            End If
        Catch
        End Try

    End Sub

    Private Sub HandleSetEntityTargetMessage(ByRef yData() As Byte)
        Dim lPos As Int32 = 2       'for msgcode
        Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEntityTypeID As Int16

        If lEntityID < 0 Then iEntityTypeID = ObjectType.eFacility Else iEntityTypeID = ObjectType.eUnit
        lEntityID = Math.Abs(lEntityID)

        Dim lTarget1ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lTarget2ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lTarget3ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lTarget4ID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim iType1ID As Int16
        Dim iType2ID As Int16
        Dim iType3ID As Int16
        Dim iType4ID As Int16

        If lTarget1ID < 0 Then iType1ID = ObjectType.eFacility Else iType1ID = ObjectType.eUnit
        If lTarget2ID < 0 Then iType2ID = ObjectType.eFacility Else iType2ID = ObjectType.eUnit
        If lTarget3ID < 0 Then iType3ID = ObjectType.eFacility Else iType3ID = ObjectType.eUnit
        If lTarget4ID < 0 Then iType4ID = ObjectType.eFacility Else iType4ID = ObjectType.eUnit
        lTarget1ID = Math.Abs(lTarget1ID)
        lTarget2ID = Math.Abs(lTarget2ID)
        lTarget3ID = Math.Abs(lTarget3ID)
        lTarget4ID = Math.Abs(lTarget4ID)

        Dim lT1Idx As Int32 = -1
        Dim lT2Idx As Int32 = -1
        Dim lT3Idx As Int32 = -1
        Dim lT4Idx As Int32 = -1

        'ok, let's set up our thing...
        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing Then Return
        Dim lCurUB As Int32 = -1
        If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityIdx.GetUpperBound(0), oEnvir.lEntityUB)
        Dim oAttacker As BaseEntity = Nothing
        For X As Int32 = 0 To lCurUB
            If oEnvir.lEntityIdx(X) = lEntityID Then
                Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iEntityTypeID Then
                    oAttacker = oEntity
                End If
            End If
            If oEnvir.lEntityIdx(X) = lTarget1ID AndAlso lTarget1ID > 0 Then
                Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iType1ID Then
                    lT1Idx = X
                End If
            End If
            If oEnvir.lEntityIdx(X) = lTarget2ID AndAlso lTarget2ID > 0 Then
                Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iType2ID Then
                    lT2Idx = X
                End If
            End If
            If oEnvir.lEntityIdx(X) = lTarget3ID AndAlso lTarget3ID > 0 Then
                Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iType3ID Then
                    lT3Idx = X
                End If
            End If
            If oEnvir.lEntityIdx(X) = lTarget4ID AndAlso lTarget4ID > 0 Then
                Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iType4ID Then
                    lT4Idx = X
                End If
            End If
        Next X

        If oAttacker Is Nothing = False Then
            oAttacker.lTargetMsg += 1
            oAttacker.lTargetIdx(0) = lT1Idx
            oAttacker.lTargetIdx(1) = lT2Idx
            oAttacker.lTargetIdx(2) = lT3Idx
            oAttacker.lTargetIdx(3) = lT4Idx
        End If

    End Sub

    Private Sub HandleFireWeaponMsg(ByRef yData() As Byte, ByVal bHit As Boolean)
        If frmMain.mbChangingEnvirs = True Then Return

        Dim lPos As Int32 = 2
        Dim lAttackerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yWpnTypeID As Byte = yData(lPos) : lPos += 1
        Dim ySideShot As Byte = yData(lPos) : lPos += 1

        Dim yAOE As Byte = 0
        'Ok, determine our aoe
        If ySideShot > 3 Then
            yAOE = ySideShot
            ySideShot = CByte(ySideShot And 3)
            yAOE = yAOE Xor ySideShot
            Dim lTemp As Int32 = yAOE
            lTemp = lTemp >> 2
            lTemp *= 4
            If lTemp > 255 Then lTemp = 255
            If lTemp < 0 Then lTemp = 0
            yAOE = CByte(lTemp)
        End If

        Dim iAttackerTypeID As Int16 = ObjectType.eUnit
        If lAttackerID < 0 Then
            iAttackerTypeID = ObjectType.eFacility
            lAttackerID = Math.Abs(lAttackerID)
        End If

        Dim bHitShields As Boolean = False
        If (yWpnTypeID And WeaponType.eShieldHitBitMask) <> 0 Then
            bHitShields = True
            yWpnTypeID -= CByte(yWpnTypeID - 128)
        End If

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing = False Then
            Dim lCurUB As Int32 = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
            Dim oTargetEntity As BaseEntity = Nothing
            Dim oAttackerEntity As BaseEntity = Nothing

            For X As Int32 = 0 To lCurUB
                If oEnvir.lEntityIdx(X) = lAttackerID Then
                    Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                    If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iAttackerTypeID Then
                        oAttackerEntity = oEntity
                        Exit For
                    End If
                End If
            Next X
            If oAttackerEntity Is Nothing = False Then
                Dim lIdx As Int32 = oAttackerEntity.lTargetIdx(ySideShot)
                If lIdx <> -1 Then
                    oTargetEntity = oEnvir.oEntity(lIdx)
                    If oTargetEntity Is Nothing = False Then

                        If bHitShields = True Then
                            If oTargetEntity.yShieldHP = 0 Then oTargetEntity.yShieldHP = 1
                        ElseIf oTargetEntity.yShieldHP <> 0 Then
                            oTargetEntity.yShieldHP = 0
                        End If

                        If goWpnMgr Is Nothing = False Then
                            goWpnMgr.AddNewEffect(oTargetEntity, oAttackerEntity, yWpnTypeID, bHit, False, yAOE)
                        End If
                        oTargetEntity.lLastWeaponUpdate = glCurrentCycle

                    End If
                End If
            End If

        End If

    End Sub

    Private Sub HandleGuildCreationUpdate(ByRef yData() As Byte)
        If goUILib Is Nothing Then Return
        Dim ofrm As frmGuildSetup = CType(goUILib.GetWindow("frmGuildSetup"), frmGuildSetup)
        If ofrm Is Nothing Then Return
        Dim oGuild As Guild = New Guild
        Dim lFormer As Int32 = -1
        With oGuild
            Dim lPos As Int32 = 2   'for msgcode

            lFormer = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            .lBaseGuildRules = CType(System.BitConverter.ToInt32(yData, lPos), elGuildFlags) : lPos += 4
            .lIcon = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .sName = GetStringFromBytes(yData, lPos, 50) : lPos += 50
            .yGuildTaxBaseMonth = yData(lPos) : lPos += 1
            .yGuildTaxBaseDay = yData(lPos) : lPos += 1
            .yGuildTaxRateInterval = CType(yData(lPos), eyGuildInterval) : lPos += 1
            .yState = CType(yData(lPos), eyGuildState) : lPos += 1
            .yVoteWeightType = CType(yData(lPos), eyVoteWeightType) : lPos += 1

            Dim lMemberCnt As Int32 = yData(lPos) : lPos += 1
            ReDim .moMembers(lMemberCnt - 1)
            For X As Int32 = 0 To lMemberCnt - 1
                .moMembers(X) = New GuildMember

                .moMembers(X).lMemberID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .moMembers(X).lRankID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .moMembers(X).yMemberState = CType(yData(lPos), GuildMemberState) : lPos += 1
            Next X

            Dim lRankCnt As Int32 = yData(lPos) : lPos += 1
            ReDim .moRanks(lRankCnt - 1)
            For X As Int32 = 0 To lRankCnt - 1
                .moRanks(X) = New GuildRank()
                .moRanks(X).lRankID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .moRanks(X).lRankPermissions = CType(System.BitConverter.ToInt32(yData, lPos), RankPermissions) : lPos += 4
                .moRanks(X).lVoteStrength = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .moRanks(X).sRankName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                .moRanks(X).TaxRateFlat = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .moRanks(X).TaxRatePercentage = yData(lPos) : lPos += 1
                .moRanks(X).TaxRatePercType = CType(yData(lPos), eyGuildTaxPercType) : lPos += 1
                .moRanks(X).yPosition = yData(lPos) : lPos += 1
            Next X
        End With

        ofrm.SetNewGuildObject(oGuild, lFormer)
    End Sub

    Private Sub HandleRegionSetEntityStatus(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lStatus As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing = False Then
            Dim lCurUB As Int32 = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))


            For X As Int32 = 0 To lCurUB
                If oEnvir.lEntityIdx(X) = lObjID Then
                    If oEnvir.oEntity(X).ObjTypeID = iObjTypeID Then
                        If lStatus < 0 Then
                            If (oEnvir.oEntity(X).CurrentStatus And Math.Abs(lStatus)) <> 0 Then oEnvir.oEntity(X).CurrentStatus = oEnvir.oEntity(X).CurrentStatus Xor Math.Abs(lStatus)
                        Else
                            oEnvir.oEntity(X).CurrentStatus = oEnvir.oEntity(X).CurrentStatus Or lStatus
                        End If
                        Exit For
                    End If
                End If
            Next X
        End If
    End Sub

    Private Sub HandlePDSShot(ByRef yData() As Byte)
        Dim lPos As Int32 = 0
        Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lAttackerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'Dim iAttackerTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim iAttackerTypeID As Int16 = ObjectType.eUnit
        If lAttackerID < 0 Then
            lAttackerID = Math.Abs(lAttackerID)
            iAttackerTypeID = ObjectType.eFacility
        End If
        Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'Dim iTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim iTargetTypeID As Int16 = ObjectType.eUnit
        If lTargetID < 0 Then
            lTargetID = Math.Abs(lTargetID)
            iTargetTypeID = ObjectType.eFacility
        End If
        Dim yWpnTypeID As Byte = yData(lPos) : lPos += 1

        Dim bHitShields As Boolean = False
        If (yWpnTypeID And WeaponType.eShieldHitBitMask) <> 0 Then
            bHitShields = True
            yWpnTypeID = CByte(yWpnTypeID Xor 128)
        End If

        Dim bHit As Boolean = (iMsgCode = GlobalMessageCode.ePDS_AtUnitHit) OrElse (iMsgCode = GlobalMessageCode.ePDS_AtMissileHit)

        'If iMsgCode = GlobalMessageCode.ePDS_AtMissileHit Then goUILib.AddNotification(lAttackerID & " PD Hit Missile " & lTargetID, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        'If iMsgCode = GlobalMessageCode.ePDS_AtMissileMiss Then goUILib.AddNotification(lAttackerID & " PD Missed Missile " & lTargetID, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        'If iMsgCode <> GlobalMessageCode.ePDS_AtMissileMiss AndAlso iMsgCode <> GlobalMessageCode.ePDS_AtMissileHit Then goUILib.AddNotification("Shot at non missile.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing = False AndAlso oEnvir.lEntityIdx Is Nothing = False Then
            Dim lCurUB As Int32 = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))

            If iMsgCode = GlobalMessageCode.ePDS_AtMissileHit OrElse iMsgCode = GlobalMessageCode.ePDS_AtMissileMiss Then
                If goMissileMgr Is Nothing Then Return
                Dim oAttackerEntity As BaseEntity = Nothing
                Dim oMissile As MissileMgr.Missile = goMissileMgr.GetMissile(lTargetID)
                For X As Int32 = 0 To lCurUB
                    If oEnvir.lEntityIdx(X) = lAttackerID Then
                        Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                        If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iAttackerTypeID Then
                            oAttackerEntity = oEntity
                            Exit For
                        End If
                    End If
                Next X
                If oAttackerEntity Is Nothing = False AndAlso oMissile Is Nothing = False Then
                    'ok, fire...
                    If goWpnMgr Is Nothing = False Then
                        goWpnMgr.AddNewEffect(oMissile, oAttackerEntity, yWpnTypeID, bHit, True)
                    End If
                End If

            Else
                Dim oTargetEntity As BaseEntity = Nothing
                Dim oAttackerEntity As BaseEntity = Nothing

                For X As Int32 = 0 To lCurUB
                    If oEnvir.lEntityIdx(X) = lAttackerID Then
                        Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                        If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iAttackerTypeID Then
                            oAttackerEntity = oEntity
                            If oTargetEntity Is Nothing = False Then Exit For
                        End If
                    ElseIf oEnvir.lEntityIdx(X) = lTargetID Then
                        Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                        If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iTargetTypeID Then
                            oTargetEntity = oEntity
                            If oAttackerEntity Is Nothing = False Then Exit For
                        End If
                    End If
                Next X
                If oAttackerEntity Is Nothing = False AndAlso oTargetEntity Is Nothing = False Then
                    If bHitShields = True Then
                        If oTargetEntity.yShieldHP = 0 Then oTargetEntity.yShieldHP = 1
                    ElseIf oTargetEntity.yShieldHP <> 0 Then
                        oTargetEntity.yShieldHP = 0
                    End If

                    If goWpnMgr Is Nothing = False Then
                        goWpnMgr.AddNewEffect(oTargetEntity, oAttackerEntity, yWpnTypeID, bHit, True, 0)
                    End If
                    oTargetEntity.lLastWeaponUpdate = glCurrentCycle
                End If
            End If



        End If

    End Sub

    Private Sub HandleOtherPlayerDeath(ByVal yData() As Byte)
        'a region server is telling us that someone died...
        '  our responsibility is to remove any units and facilities belonging to the player in the environment
        '  remove any player rel I have with them
        '  and if I had a player rel with them give an alert that they are gone
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oEnvir As BaseEnvironment = goCurrentEnvir
        If oEnvir Is Nothing = False Then
            Dim lCurUB As Int32 = -1
            If oEnvir.lEntityIdx Is Nothing = False Then lCurUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
            For X As Int32 = 0 To lCurUB
                If oEnvir.lEntityIdx(X) > -1 Then
                    Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                    If oEntity Is Nothing = False Then
                        If oEntity.OwnerID = lPlayerID Then


                            With oEntity
                                .bObjectDestroyed = True
                                .ClearParticleFX()
                                Try
                                    Dim fTX As Single
                                    Dim fTZ As Single
                                    Dim fDist As Single
                                    fTX = .LocX - goCamera.mlCameraX
                                    fTZ = .LocZ - goCamera.mlCameraZ
                                    fTX *= fTX
                                    fTZ *= fTZ
                                    fDist = CSng(Math.Sqrt(fTX + fTZ))
                                    If fDist < 10000.0F Then
                                        fDist /= 10000.0F
                                        If .oUnitDef.HullSize > 0 Then fDist *= .oUnitDef.HullSize
                                        goCamera.ScreenShake(500 * fDist)
                                    End If
                                Catch
                                End Try

                                Dim bDoDeathSequence As Boolean = True
                                If .OwnerID <> glPlayerID AndAlso .OwnerID <> gl_HARDCODE_PIRATE_PLAYER_ID Then
                                    If .yVisibility = eVisibilityType.FacilityIntel Then
                                        bDoDeathSequence = False
                                    End If
                                End If

                                If bDoDeathSequence = True AndAlso goEntityDeath Is Nothing = False Then goEntityDeath.AddNewDeathSequence(oEntity, X)
                            End With

                        End If
                    End If
                End If
            Next X
        End If

        If goCurrentPlayer Is Nothing = False Then
            Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRel(lPlayerID)
            If oRel Is Nothing = False Then
                If goUILib Is Nothing = False Then goUILib.AddNotification("Intel reports that we have lost contact with " & GetCacheObjectValue(lPlayerID, ObjectType.ePlayer) & ".", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
            goCurrentPlayer.RemovePlayerRel(lPlayerID)
        End If

    End Sub

#End Region


End Class
