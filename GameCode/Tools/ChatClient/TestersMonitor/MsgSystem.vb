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
    
    Public msPrimaryIP As String = ""
    Public mlPrimaryPort As Int32 = 0
    Private mlTimeout As Int32

    Private myReactionData() As Byte
    Private moReactionThread As Threading.Thread

    Private mlBusyStatus As Int32 = 0

	Private mbClosing As Boolean = False

    Private msCurrentLoc As String
	Private mbVersionValidated As Boolean = False

	Private mlLastAlert As Int32		'generally speaking

	Private mlLastAddObjectID As Int32 = -1
	Private miLastAddObjectTypeID As Int16 = -1

    Public bDoNotClearBusyStatus As Boolean = False
 
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
        If moPrimary Is Nothing = False Then moPrimary.Disconnect()
        'If moOperator Is Nothing = False Then moOperator.Disconnect()
    End Sub

    Public Sub ResetClosing()
        mbClosing = False
    End Sub

	Public Sub SendToPrimary(ByVal yData() As Byte)
		msCurrentLoc = "SendToPrimary"
		moPrimary.SendData(yData)
	End Sub
    Public Sub DisconnectPrimary()
        msCurrentLoc = "DisconnectPrimary"
        moPrimary.Disconnect()
    End Sub

#Region "  Connect to Server Methods  "
 
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

        msCurrentLoc = "pda"

        dtLastMsg = Now
		 
		iMsgID = System.BitConverter.ToInt16(Data, 0)
		'If (mlBusyStatus And elBusyStatusType.PlayerDetails) = 0 Then
        '	AddTextLine("Primary: " & CType(iMsgID, GlobalMessageCode).ToString, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        'End If
        'Debug.Print("moPrimary.onDataArrival: " & CType(iMsgID, GlobalMessageCode).ToString)
        msCurrentLoc = "pda. (" & iMsgID & ")"

        'AddTextLine("Primary: " & CType(iMsgID, GlobalMessageCode).ToString, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        Select Case iMsgID
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
            Case GlobalMessageCode.eChatMessage
                HandleChatMsg(Data)
            Case GlobalMessageCode.eRequestSkillList
                HandleSkillListMsg(Data)
            Case GlobalMessageCode.eRequestGoalList
                HandleGoalListMsg(Data)
            Case GlobalMessageCode.eServerShutdown
                RaiseEvent ServerShutdown()
            Case GlobalMessageCode.eClientVersion
                HandleValidateVersion(Data, eyConnType.PrimaryServer)
            Case GlobalMessageCode.eRequestPlayerDetails
                HandleRequestPlayerDetailsResponse(Data)
            Case GlobalMessageCode.eRequestPlayerDetailsPlayerMin
                HandleRequestPlayerDetailsPlayerMin(Data)
            Case GlobalMessageCode.eRequestEmailSummarys
                HandleRequestEmailSummarys(Data)
            Case GlobalMessageCode.eRequestPlayerDetailsAlert
                HandleRequestPlayerDetailsAlert(Data)
            Case GlobalMessageCode.eRequestChannelList
                If gfrmChannels Is Nothing = False Then
                    gfrmChannels.HandleRequestChannelList(Data)
                End If
            Case GlobalMessageCode.eRequestChannelDetails
                If gfrmChannelConfig Is Nothing = False Then
                    gfrmChannelConfig.HandleChannelDetails(Data)
                End If
            Case GlobalMessageCode.eAddObjectCommand
                HandleAddObjectMsg(Data, True)
            Case GlobalMessageCode.eGetEntityName
                HandleGetEntityName(Data)
            Case GlobalMessageCode.eUpdatePlayerCredits
                HandleUpdatePlayerCredits(Data)
            Case GlobalMessageCode.eProductionCompleteNotice
                'TODO: More testing HandleProductionCompleteNotice(Data)
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

            If gfrmChat Is Nothing = False Then
                gfrmChat.AddTextLine(-1, "\cf7 GAME ERROR (" & msCurrentLoc & "): " & Description, ChatMessageType.eSysAdminMessage)
            Else
                MsgBox("Error (" & msCurrentLoc & "): " & Description, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, "Error")
            End If

        End If
    End Sub

#End Region

#Region " Outgoing Message Handlers "

    Public Sub RequestPlayerDetails()
        Dim yData(15) As Byte

        msCurrentLoc = "RequestPlayerDetails"

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerDetails).CopyTo(yData, 0)
        System.BitConverter.GetBytes(glPlayerID).CopyTo(yData, 2)
        System.BitConverter.GetBytes(elPlayerDetailsType.eAliasConfigs).CopyTo(yData, 6)
        System.BitConverter.GetBytes(-1I).CopyTo(yData, 10)
        System.BitConverter.GetBytes(-1S).CopyTo(yData, 14)

        SendToPrimary(yData)

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

        Dim bSA As Boolean = False
        If yType = eyConnType.PrimaryServer Then
            If Command() Is Nothing OrElse Command.ToUpper <> "SA" Then
                ReDim yData(84) 'all entries should be 0
            Else
                ReDim yData(90) 'all entries should be 0
                bSA = True
            End If
        Else
            ReDim yData(84) 'all entries should be 0
        End If

        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.ePlayerLoginRequest).CopyTo(yData, lPos) : lPos += 2
        oEnc.Encrypt(System.Text.ASCIIEncoding.ASCII.GetBytes(sUserName)).CopyTo(yData, lPos) : lPos += 20

        'System.Text.ASCIIEncoding.ASCII.GetBytes(sPassword).CopyTo(yData, lPos) : lPos += 20
        yEncPassword.CopyTo(yData, lPos) : lPos += 21

        If bAlias = True Then
            yData(lPos) = 1
        Else : yData(lPos) = 0
        End If
        lPos += 1
        oEnc.Encrypt(System.Text.ASCIIEncoding.ASCII.GetBytes(sAliasName)).CopyTo(yData, lPos) : lPos += 20
        'System.Text.ASCIIEncoding.ASCII.GetBytes(sAliasPassword).CopyTo(yData, lPos) : lPos += 20
        yEncAliasPW.CopyTo(yData, lPos) : lPos += 21

        If yType = eyConnType.PrimaryServer AndAlso bSA = True Then
            Dim oEP As Net.IPEndPoint = CType(moPrimary.GetLocalDetails(), Net.IPEndPoint)
            Dim sIP As String = oEP.Address().ToString()
            Dim yMAC() As Byte = MACAddress.GetMACAddressFromIP(sIP)
            If yMAC Is Nothing = False Then yMAC.CopyTo(yData, lPos) : lPos += 6
        End If
        'mlLoginResp = LoginResponseCodes.eUnknownFailure

        Dim lStatus As Int32
        If yType = eyConnType.PrimaryServer Then
            lStatus = elBusyStatusType.PrimaryServerLogin
            mlBusyStatus = mlBusyStatus Or lStatus
            moPrimary.SendData(yData)
        End If
 

        Return True
    End Function

    Public Sub SendChangeEnvironment(ByVal lID As Int32, ByVal iTypeID As Int16)
        Dim yData(11) As Byte
 

        msCurrentLoc = "SendChangeEnvironment"

        'Player ID, the new Envir ID, the new Envir Type ID
        System.BitConverter.GetBytes(GlobalMessageCode.eChangingEnvironment).CopyTo(yData, 0)
        System.BitConverter.GetBytes(glPlayerID).CopyTo(yData, 2)
        System.BitConverter.GetBytes(lID).CopyTo(yData, 6)
        System.BitConverter.GetBytes(iTypeID).CopyTo(yData, 10)

        'Now, send to the Primary...
        SendToPrimary(yData)

    End Sub

    Public Function ValidateVersion(ByVal yType As eyConnType) As Boolean
        Dim yData(5) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eClientVersion).CopyTo(yData, 0)

        If frmMain.sCommandArg = "SA" Then
            System.BitConverter.GetBytes(gl_SA_CLIENT_VERSION).CopyTo(yData, 2)
        Else
            System.BitConverter.GetBytes(gl_BP_CLIENT_VERSION).CopyTo(yData, 2)
        End If

        msCurrentLoc = "ValidateVersion"

        Dim lStatus As Int32
        
        lStatus = elBusyStatusType.PrimaryValidateVersion
        mlBusyStatus = mlBusyStatus Or lStatus
        moPrimary.SendData(yData)

        Return True 'mbVersionValidated
    End Function

    Private mlPrimaryFailures As Int32 = 0 
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
 

        If mlPrimaryFailures > 3 Then
            gfrmChat.AddTextLine(-1, "\cf7 {\b Connection lost to the server!}", ChatMessageType.eSysAdminMessage)
            
            Return False
        End If

        Return True
    End Function

#End Region

#Region " Incoming Message Handlers "


    Private Sub HandleAddObjectMsg(ByVal yData() As Byte, ByVal bFromPrimary As Boolean)
        'ok, the primary server sent us an Add Object Message... yData has the pertinent information...
        ' an Add Object has the following format: ObjID, ObjTypeID, ParentID, ParentTypeID... data after that
        ' is specific to the object being added
        msCurrentLoc = "HandleAddObjectMsg"
        Dim lObjID As Int32
        Dim iObjTypeID As Int16
 
        Dim lPos As Int32

        Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yData, 0) 

        lObjID = System.BitConverter.ToInt32(yData, 2)
        iObjTypeID = System.BitConverter.ToInt16(yData, 6)
        lPos = 8
 
        'NOTE: Make sure that any object type that requires goCurrentEnvir is added to this list

        Select Case iObjTypeID 
            Case ObjectType.ePlayer
                'Ok, the player object is being added...

                If lObjID = glPlayerID Then
   
                    '.ObjectID = lObjID
                    '.ObjTypeID = iObjTypeID

                    ''Player name (20)
                    '.PlayerName = GetStringFromBytes(yData, 8, 20)
                    ''Empire name (20)
                    '.EmpireName = GetStringFromBytes(yData, 28, 20)
                    ''Race Name (20)
                    '.RaceName = GetStringFromBytes(yData, 48, 20)

                    ''.lSenateID = System.BitConverter.ToInt32(yData, 68)
                    '.bIsMale = yData(68) <> 2
                    '.CommEncryptLevel = System.BitConverter.ToInt16(yData, 72)
                    '.EmpireTaxRate = yData(74)

                    lPos = 75
                    lPos += 8

                    gyPlayerTitle = yData(lPos) : lPos += 1

                    '.yPlayerTitle = yData(lPos) : lPos += 1
                    '.lPlayerIcon = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    '.lTechnologyScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    '.lDiplomacyScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    '.lMilitaryScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    '.lPopulationScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    '.lProductionScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    '.lWealthScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    '.lTotalScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    '.DeathBudgetBalance = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                    '.BadWarDecCPIncrease = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    '.BadWarDecMoralePenalty = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    '.yPlayerPhase = yData(lPos) : lPos += 1
                    '.lTutorialStep = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    '.lAccountStatus = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                    '.lStatusFlags = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                End If
        End Select

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

    Private mlChatMsg As Int32 = 0
    Private Sub HandleChatMsg(ByVal yData() As Byte)
        msCurrentLoc = "HandleChatMsg"

        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim yType As Byte = yData(6)

        Dim lLen As Int32 = System.BitConverter.ToInt32(yData, 7)
        Dim sMsg As String = GetStringFromBytes(yData, 11, lLen)

        sMsg = sMsg.Replace("\", "\\")

        If glPlayerID = 1 Then
            sMsg = "{\i " & lPlayerID & "} " & sMsg
        End If

        If lPlayerID = glPlayerID Then
            sMsg = "\cf9 " & sMsg
        Else
            sMsg = "\cf" & (CInt(yType) + 1).ToString & " " & sMsg
        End If


        gfrmChat.AddTextLine(lPlayerID, sMsg, CType(yType, ChatMessageType))

    End Sub

    Private Sub HandleSkillListMsg(ByVal yData() As Byte)
        Dim lCnt As Int32 = System.BitConverter.ToInt16(yData, 2)


        msCurrentLoc = "HandleSkillListMsg"



        If (mlBusyStatus And elBusyStatusType.SkillsList) <> 0 Then
            mlBusyStatus = mlBusyStatus Xor elBusyStatusType.SkillsList
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

    Private Sub HandleValidateVersion(ByVal yData() As Byte, ByVal yConnType As eyConnType)
        Dim lVersion As Int32 = System.BitConverter.ToInt32(yData, 2)
        msCurrentLoc = "HandleValidateVersion"

        Dim lTestVer As Int32 = gl_BP_CLIENT_VERSION
        If frmMain.sCommandArg = "SA" Then
            lTestVer = gl_SA_CLIENT_VERSION
        End If
        mbVersionValidated = (lVersion = lTestVer)

        'If (mlBusyStatus And elBusyStatusType.OperatorValidateVersion) <> 0 Then
        '	mlBusyStatus = mlBusyStatus Xor elBusyStatusType.OperatorValidateVersion
        'End If
        'If (mlBusyStatus And elBusyStatusType.PrimaryValidateVersion) <> 0 Then
        '	mlBusyStatus = mlBusyStatus Xor elBusyStatusType.PrimaryValidateVersion
        'End If

        RaiseEvent ValidateVersionResponse(yConnType, mbVersionValidated)
    End Sub

    Private Sub HandleRequestPlayerDetailsPlayerMin(ByRef yData() As Byte)
    End Sub

    Private Sub HandleRequestPlayerDetailsAlert(ByRef yData() As Byte)
    End Sub

    Private Sub HandleRequestEmailSummarys(ByRef yData() As Byte)
    End Sub

    Private Sub HandleUpdatePlayerCredits(ByRef yData() As Byte)
        msCurrentLoc = "HandleUpdatePlayerCredits"
        If frmMain.sCommandArg = "SA" Then Return
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        glCredits = System.BitConverter.ToInt64(yData, 8)
        glCashFlow = System.BitConverter.ToInt64(yData, 16)
        'glWarpoints = System.BitConverter.ToInt64(yData, 24)
        'glCurrentWPUpkeepCost = System.BitConverter.ToInt32(yData, 32)
    End Sub

    Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Integer, ByVal dwFlags As Integer) As Integer
    Private Const SND_ASYNC As Integer = &H1
    Private Const SND_ALIAS As Integer = &H10000

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

        Dim sCreatedBy As String = ""
        Dim sCreatedWhat As String = ""
        Dim sCreatedWhere As String = ""
        Dim sFinal As String

        Dim sProdRes As String = "Production"
        Dim sWAVFile As String = "Game Narrative\ProductionComplete.wav"

        If iProdTypeID = ObjectType.eRepairItem Then Return

        Select Case iProdTypeID
            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.eMineralTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.eSpecialTech
                'if the production type is research... set it to Research, otherwise, it is PRoduction
                If yProdType = ProductionType.eResearch Then
                    sWAVFile = "Game Narrative\ResearchComplete.wav"
                    sProdRes = "Research"
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

        Call PlaySound("SystemNotification", 0, SND_ASYNC Or SND_ALIAS)

        sCreatedBy = GetCacheObjectValue(lCreatorID, iCreatorTypeID)

        'Cannot trust goCurrentEnvir...
        sCreatedWhere = GetCacheObjectValue(lEnvirID, iEnvirTypeID)

        If sCreatedBy = "Unknown" Then sCreatedBy = ""
        If sCreatedWhere = "Unknown" Then sCreatedWhere = ""

        If iProdTypeID = ObjectType.eUnitDef OrElse iProdTypeID = ObjectType.eFacilityDef OrElse iProdTypeID = ObjectType.eEnlisted OrElse iProdTypeID = ObjectType.eOfficers Then
            'For X = 0 To glEntityDefUB
            '    If glEntityDefIdx(X) = lProdID AndAlso goEntityDefs(X).ObjTypeID = iProdTypeID Then
            '        sCreatedWhat = goEntityDefs(X).DefName
            '    End If
            '    Exit For
            '    End If
            'Next X
        ElseIf iProdTypeID = ObjectType.eMineralTech Then
            'For X = 0 To glMineralUB
            '    If glMineralIdx(X) = lProdID Then
            '        sCreatedWhat = goMinerals(X).MineralName
            '        Exit For
            '    End If
            'Next X
        Else
            'Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lProdID, Math.Abs(iProdTypeID))
            'If oTech Is Nothing = False Then
            '    sCreatedWhat = oTech.GetComponentName()
            'End If
            'oTech = Nothing
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

        'If bDoNotification = True Then
        '    goUILib.AddNotification(sFinal, System.Drawing.Color.FromArgb(255, 255, 255, 255), lEnvirID, iEnvirTypeID, lCreatorID, iCreatorTypeID, Int32.MinValue, Int32.MinValue)
        'End If
        gfrmChat.AddTextLine(-1, "\cf3 {\i " & sFinal & "}", ChatMessageType.eNotificationMessage)

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
            gfrmChat.AddTextLine(-1, "\cf7 {\bUnable to log in at this time, please try again.}", ChatMessageType.eSysAdminMessage)
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



#End Region


End Class
