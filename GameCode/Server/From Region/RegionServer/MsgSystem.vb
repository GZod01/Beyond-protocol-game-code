Option Strict On

Public Structure ClientConnData
	Public mlClientLastRequest As Int32
	Public lCameraPosX As Int32
	Public lCameraPosZ As Int32
    Public mlLastGetEntityDetails As Int32
End Structure

Public Class MsgSystem
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

    Private WithEvents moPrimary As NetSock
    Private WithEvents moPathfinding As NetSock
    Private WithEvents moClientListener As NetSock
    Private WithEvents moOperator As NetSock

    Private moClients() As NetSock
	Private mlClientPlayer() As Int32
	Private muClientData() As ClientConnData

    Private mlAliasedAs() As Int32      'the playerID I am aliasing
    Private mlAliasedRights() As Int32  'the rights I am aliasing

    Private mlClientUB As Int32 = -1
    Private mbAcceptingClients As Boolean = False
    Private mbConnectedToPrimary As Boolean = False
    Private mbConnectedToPathfinding As Boolean = False
    Private mbHavePathfindingInfo As Boolean = False

    Private mbSyncWait As Boolean = False
    Private mlTimeout As Int32

    Private Const ml_REQUEST_THROTTLE As Int32 = 15     'cycles
	'Private mlClientLastRequest() As Int32

    Public Function GetConnectedClientCount() As Int32
        Dim X As Int32
        Dim lCnt As Int32 = 0

        For X = 0 To mlClientUB
			If moClients(X) Is Nothing = False Then
				lCnt += 1
			End If
        Next X
        Return lCnt
    End Function

    Public Sub New()
        'Just your standard NEW statement...
        ReDim moClients(0)
        ReDim mlClientPlayer(0)
        ReDim mlAliasedAs(0)
        ReDim mlAliasedRights(0)
		'ReDim mlClientLastRequest(0)
		ReDim muClientData(0)
        mlClientUB = -1
    End Sub

    Public Sub New(ByVal lPreDimClients As Int32)
        mlClientUB = lPreDimClients
        ReDim moClients(mlClientUB)
        ReDim mlAliasedAs(mlClientUB)
        ReDim mlAliasedRights(mlClientUB)
		'ReDim mlClientLastRequest(mlClientUB)
		ReDim muClientData(mlClientUB)
        ReDim mlClientPlayer(mlClientUB)

        For X As Int32 = 0 To mlClientUB
            mlAliasedAs(X) = -1
        Next X
    End Sub

    Public Sub CloseAllConnections()
        Dim X As Int32
        AcceptingClients = False
        For X = 0 To mlClientUB
            If moClients(X) Is Nothing = False Then
                moClients(X).Disconnect()
            End If
            moClients(X) = Nothing
        Next X
        If moPathfinding Is Nothing = False Then moPathfinding.Disconnect()
        moPathfinding = Nothing

        'Ok, before closing our connection to the primary, dump everything to it...
        Dim yAllData() As Byte = GetAll()
        moPrimary.SendLenAppendedData(yAllData)

        If goMsgMonitor Is Nothing = False Then goMsgMonitor.SaveAll()

        'give us time
        Threading.Thread.Sleep(10)

        'Finally ,send the ServerShutdown message indicating that we are finished and shutting down
        Dim yMsg(1) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eServerShutdown).CopyTo(yMsg, 0)
		moPrimary.SendData(yMsg)

		'And disconnect the primary
		If moPrimary Is Nothing = False Then moPrimary.Disconnect()
		moPrimary = Nothing
	End Sub

	Private Function GetAll() As Byte()
		Dim lPos As Int32
		Dim lSingleMsgLen As Int32
		Dim X As Int32
		Dim yTemp() As Byte
		Dim yCache(200000) As Byte
		Dim yFinal() As Byte

		lPos = 0
		lSingleMsgLen = -1
		Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
		For X = 0 To lCurUB
            If glEntityIdx(X) > 0 Then

                'Be sure the entity is disengaged...
                If (goEntity(X).CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then goEntity(X).CurrentStatus = goEntity(X).CurrentStatus Xor elUnitStatus.eUnitEngaged
                If (goEntity(X).CurrentStatus And elUnitStatus.eSide1HasTarget) <> 0 Then goEntity(X).CurrentStatus = goEntity(X).CurrentStatus Xor elUnitStatus.eSide1HasTarget
                If (goEntity(X).CurrentStatus And elUnitStatus.eSide2HasTarget) <> 0 Then goEntity(X).CurrentStatus = goEntity(X).CurrentStatus Xor elUnitStatus.eSide2HasTarget
                If (goEntity(X).CurrentStatus And elUnitStatus.eSide3HasTarget) <> 0 Then goEntity(X).CurrentStatus = goEntity(X).CurrentStatus Xor elUnitStatus.eSide3HasTarget
                If (goEntity(X).CurrentStatus And elUnitStatus.eSide4HasTarget) <> 0 Then goEntity(X).CurrentStatus = goEntity(X).CurrentStatus Xor elUnitStatus.eSide4HasTarget

                yTemp = goEntity(X).GetObjectUpdateMsg(GlobalMessageCode.eUpdateEntityAndSave, 0)
                If yTemp Is Nothing Then Continue For
                lSingleMsgLen = yTemp.Length

                'Ok, before we continue, check if we need to increase our cache
                If lPos + lSingleMsgLen + 2 > yCache.Length Then
                    'increase it
                    ReDim Preserve yCache(yCache.Length + 200000)
                End If
                System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                lPos += 2

                yTemp.CopyTo(yCache, lPos)
                lPos += lSingleMsgLen
            End If
        Next X
		If lPos <> 0 Then
			ReDim yFinal(lPos - 1)
			Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
			Return yFinal
		Else : Return Nothing
		End If
	End Function

    Public Property AcceptingClients() As Boolean
        Get
            Return mbAcceptingClients
        End Get
        Set(ByVal Value As Boolean)
            'Dim lPort As Int32

            If mbAcceptingClients <> Value Then
                mbAcceptingClients = Value
                If mbAcceptingClients = True Then
                    'Ok... start listening
                    'Dim oINI As New InitFile()
                    Try
                        'lPort = CInt(Val(oINI.GetString("SETTINGS", "ClientListenPort", "0")))
                        'If lPort = 0 Then Err.Raise(-1, "AcceptingClients", "Could not get Client Listen Port Number from INI.")
                        moClientListener = Nothing
                        moClientListener = New NetSock()
                        moClientListener.PortNumber = glExternalPort
                        moClientListener.Listen()
                    Catch
                        MsgBox("An error was caught in " & Err.Source & ":" & vbCrLf & Err.Description, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, Err.Source)
                        mbAcceptingClients = False
                        'Finally
                        'oINI = Nothing
                    End Try
                Else
                    'ok, stop listening
                    moClientListener.StopListening()
                End If
            End If
        End Set
    End Property

	Public Function ConnectToPathfinding() As Boolean
		Dim oINI As New InitFile()
		Dim lPort As Int32
		Dim sIP As String

		Try
			lPort = CInt(Val(oINI.GetString("PATHFINDING", "PortNumber", "0")))
			If lPort = 0 Then
				Err.Raise(-1, "ConnectToServers", "Could not find Pathfinding Server Port Number in INI.")
			End If
			sIP = oINI.GetString("PATHFINDING", "IPAddress", "")
			If sIP = "" Then
				Err.Raise(-1, "ConnectToServers", "Could not find Pathfinding Server IP Address in INI.")
			End If
			moPathfinding = New NetSock()
			moPathfinding.Connect(sIP, lPort)
			Dim sw As Stopwatch = Stopwatch.StartNew
			While mbConnectedToPathfinding = False AndAlso (sw.ElapsedMilliseconds < mlTimeout)
				Application.DoEvents()
			End While
			If mbConnectedToPathfinding = False Then
				moPrimary.Disconnect()
				Err.Raise(-1, "ConnectToServers", "Connection to Pathfinding Server timed out.")
			End If
		Catch
			MsgBox("An error was caught in " & Err.Source & ":" & vbCrLf & Err.Description, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, Err.Source)
			mbConnectedToPathfinding = False
		Finally
			oINI = Nothing
		End Try
		Return mbConnectedToPathfinding
	End Function

	Public Function ConnectToPrimary() As Boolean
		Dim oINI As New InitFile()
		Dim lPort As Int32
		Dim sIP As String
		Dim bRes As Boolean

		Try
			mlTimeout = CInt(Val(oINI.GetString("SETTINGS", "ConnectTimeout", "15000")))

			'Connect to the Primary
			lPort = CInt(Val(oINI.GetString("PRIMARY", "PortNumber", "0")))
			If lPort = 0 Then
				Err.Raise(-1, "ConnectToServers", "Could not find Primary Server Port Number in INI.")
			End If
			sIP = oINI.GetString("PRIMARY", "IPAddress", "")
			If sIP = "" Then
				Err.Raise(-1, "ConnectToServers", "Could not find Primary Server IP Address in INI.")
			End If
			moPrimary = New NetSock()
			moPrimary.Connect(sIP, lPort)
			Dim sw As Stopwatch = Stopwatch.StartNew
			While mbConnectedToPrimary = False AndAlso (sw.ElapsedMilliseconds < mlTimeout)
				Application.DoEvents()
			End While
			If mbConnectedToPrimary = False Then
				Err.Raise(-1, "ConnectToServers", "Connection to Primary Server timed out.")
			End If

			bRes = mbConnectedToPrimary
		Catch
			MsgBox("An error was caught in " & Err.Source & ":" & vbCrLf & Err.Description, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, Err.Source)
			bRes = False
		Finally
			oINI = Nothing
		End Try
		Return bRes
	End Function

	Public Sub BroadcastToClients(ByVal yMsg() As Byte)
		'Sends sMsg to ALL clients regardless of environment
		Dim X As Int32

		Dim iMsgCode As Int16 = 0
		If gb_MONITOR_MSGS = True Then iMsgCode = System.BitConverter.ToInt16(yMsg, 0)

		For X = 0 To mlClientUB
			If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(iMsgCode, MsgMonitor.eMM_AppType.ClientConnection, yMsg.Length, mlClientPlayer(X))
			moClients(X).SendData(yMsg)
		Next X
	End Sub

	Public Sub BroadcastToEnvironmentClients(ByVal yMsg() As Byte, ByRef oEnvir As Envir)
		Dim X As Int32
        'oEnvir.lPlayersInEnvirCnt = 0
        'On Error Resume Next

		Dim iMsgCode As Int16 = 0
		If gb_MONITOR_MSGS = True Then iMsgCode = System.BitConverter.ToInt16(yMsg, 0)

		Dim yTemp(yMsg.Length + 1) As Byte
		System.BitConverter.GetBytes(CShort(yMsg.Length)).CopyTo(yTemp, 0)
		yMsg.CopyTo(yTemp, 2)

        Try
            For X = 0 To oEnvir.lPlayersInEnvirUB
                If oEnvir.lPlayersInEnvirIdx(X) <> -1 Then

                    If oEnvir.oPlayersInEnvir(X) Is Nothing = False Then
                        If oEnvir.oPlayersInEnvir(X).oSocket Is Nothing = False Then
                            Try
                                If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(iMsgCode, MsgMonitor.eMM_AppType.ClientConnection, yMsg.Length, oEnvir.lPlayersInEnvirIdx(X))
                                oEnvir.oPlayersInEnvir(X).oSocket.SendLenAppendedData(yTemp)
                            Catch
                            End Try
                        End If
                    End If

                End If
            Next X
        Catch
        End Try

    End Sub

	Public Sub BroadcastToEnvironmentClients_Ex(ByVal iMsgCode As Int16, ByVal yMsg() As Byte, ByRef oEnvir As Envir)
		Dim yFinal() As Byte
        Dim X As Int32

        'On Error Resume Next

		ReDim yFinal(yMsg.GetUpperBound(0) + 2)
		System.BitConverter.GetBytes(iMsgCode).CopyTo(yFinal, 0)
		yMsg.CopyTo(yFinal, 2)

        Try
            For X = 0 To oEnvir.lPlayersInEnvirUB
                If oEnvir.lPlayersInEnvirIdx(X) <> -1 AndAlso oEnvir.oPlayersInEnvir(X) Is Nothing = False AndAlso oEnvir.oPlayersInEnvir(X).oSocket Is Nothing = False Then
                    Try
                        If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(iMsgCode, MsgMonitor.eMM_AppType.ClientConnection, yMsg.Length, oEnvir.lPlayersInEnvirIdx(X))
                        oEnvir.oPlayersInEnvir(X).oSocket.SendData(yFinal)
                    Catch
                    End Try
                End If
            Next X
        Catch
        End Try

        Erase yFinal
    End Sub

	Public Sub BroadcastToEnvironmentClients_Filter(ByVal yMsg() As Byte, ByRef oEnvir As Envir, ByVal lCameraPosX As Int32, ByVal lCameraPosZ As Int32)
		Dim X As Int32

        'On Error Resume Next

		Dim iMsgCode As Int16 = 0
		If gb_MONITOR_MSGS = True Then iMsgCode = System.BitConverter.ToInt16(yMsg, 0)

		Dim yTemp(yMsg.Length + 1) As Byte
		System.BitConverter.GetBytes(CShort(yMsg.Length)).CopyTo(yTemp, 0)
		yMsg.CopyTo(yTemp, 2)

        Try
            For X = 0 To oEnvir.lPlayersInEnvirUB
                If oEnvir.lPlayersInEnvirIdx(X) <> -1 AndAlso oEnvir.oPlayersInEnvir(X) Is Nothing = False AndAlso oEnvir.oPlayersInEnvir(X).oSocket Is Nothing = False Then
                    Try
                        Dim lClientIdx As Int32 = oEnvir.oPlayersInEnvir(X).oSocket.SocketIndex
                        If Math.Abs(muClientData(lClientIdx).lCameraPosX - lCameraPosX) < 2 AndAlso Math.Abs(muClientData(lClientIdx).lCameraPosZ - lCameraPosZ) < 2 Then
                            If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(iMsgCode, MsgMonitor.eMM_AppType.ClientConnection, yMsg.Length, oEnvir.lPlayersInEnvirIdx(X))
                            oEnvir.oPlayersInEnvir(X).oSocket.SendLenAppendedData(yTemp)
                        End If
                    Catch
                    End Try
                End If
            Next X
        Catch
        End Try
    End Sub

	Public Sub SendToPathfinding(ByVal yMsg() As Byte)
		moPathfinding.SendData(yMsg)
	End Sub

	Public Sub SendToPrimary(ByVal yMsg() As Byte)
		moPrimary.SendData(yMsg)
	End Sub

	Public Function RequestPathfindingInfo() As Boolean
		Dim yMsg() As Byte

		gfrmDisplayForm.AddEventLine("Requesting Pathfinding Info")

		ReDim yMsg(9)	'0 to 9 = 10 bytes
		System.BitConverter.GetBytes(GlobalMessageCode.ePathfindingConnectionInfo).CopyTo(yMsg, 0)
		yMsg(2) = 0 : yMsg(3) = 0 : yMsg(4) = 0 : yMsg(5) = 0 : yMsg(6) = 0 : yMsg(7) = 0

		moPrimary.SendData(yMsg)

		Dim sw As Stopwatch = Stopwatch.StartNew
		While mbHavePathfindingInfo = False AndAlso (sw.ElapsedMilliseconds < mlTimeout)
			Application.DoEvents()
		End While
		sw.Stop()
		sw = Nothing

		Return mbHavePathfindingInfo
	End Function

	Private Sub ParsePathfindingConnectionInfo(ByVal yData() As Byte)
		'Pathfinding connection info is:
		'MsgCode(2), IP(4), Port(4)
		Dim sIP As String
		Dim lPort As Int32
		Dim oINI As New InitFile()

		sIP = CStr(yData(2)) & "." & CStr(yData(3)) & "." & CStr(yData(4)) & "." & CStr(yData(5))
		lPort = System.BitConverter.ToInt32(yData, 6)
		oINI.WriteString("PATHFINDING", "IPAddress", sIP)
		oINI.WriteString("PATHFINDING", "PortNumber", CStr(lPort))
		mbHavePathfindingInfo = True
	End Sub

	Public Shared Function StringToBytes(ByVal sData As String) As Byte()
		Return System.Text.ASCIIEncoding.ASCII.GetBytes(sData)
	End Function

	Public Shared Function BytesToString(ByVal yData As Byte()) As String
		Return System.Text.ASCIIEncoding.ASCII.GetString(yData)
	End Function


#Region "  Operator Events  "
	Private mbConnectingToOperator As Boolean = False
	Private mbConnectedToOperator As Boolean = False
	Public bReceivedAssignment As Boolean = False
	Public Function ConnectToOperator() As Boolean

		Try
			mlTimeout = 10000
			mbConnectingToOperator = True
			moOperator = New NetSock()
			moOperator.Connect(gsOperatorIP, glOperatorPort)
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
                Dim yMsg(45) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eRegisterDomain).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(glBoxOperatorID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(ConnectionType.eRegionServerApp).CopyTo(yMsg, lPos) : lPos += 4

                Dim oMyProcess As System.Diagnostics.Process = System.Diagnostics.Process.GetCurrentProcess()
                Dim lProcessID As Int32 = oMyProcess.Id
                System.BitConverter.GetBytes(lProcessID).CopyTo(yMsg, lPos) : lPos += 4
                oMyProcess = Nothing

                System.BitConverter.GetBytes(1I).CopyTo(yMsg, lPos) : lPos += 4     'indicates 1 connection specifics

                'Get our Port number data
                'Dim oIni As New InitFile()
                ''TODO: We recently killed our INI file, so this port number will always be 7710... probably wanna change that later
                'Dim lClientPort As Int32 = CInt(Val(oIni.GetString("SETTINGS", "ClientListenPort", "7710")))
                'If lClientPort = 0 Then Err.Raise(-1, "moOperator.onConnect", "moOperator.onConnect: Unable to determine Client Listen Port from INI")
                'oIni = Nothing

                'Now, we'll indicate our client listener... so use "EXTERNALIPADDY" key word
                StringToBytes(gsExternalIP).CopyTo(yMsg, lPos) : lPos += 20
                System.BitConverter.GetBytes(glExternalPort).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(ConnectionType.eClientConnection).CopyTo(yMsg, lPos) : lPos += 4

                moOperator.SendData(yMsg)
            End With
        End If
    End Sub

	Private Sub moOperator_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moOperator.onConnectionRequest
		'should never happen
	End Sub

	Private Sub moOperator_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moOperator.onDataArrival
		Dim iMsgID As Int16
		iMsgID = System.BitConverter.ToInt16(Data, 0)
		If gb_MONITOR_MSGS = True Then goMsgMonitor.AddInMsg(iMsgID, MsgMonitor.eMM_AppType.OperatorServer, Data.Length, Index)
		Select Case iMsgID
			Case GlobalMessageCode.eRegisterDomain
				HandleOperatorDomainRegister(Data)
			Case GlobalMessageCode.eRequestObject
				moOperator.SendData(Data)
		End Select
	End Sub

	Private Sub moOperator_onDisconnect(ByVal Index As Integer) Handles moOperator.onDisconnect
		gfrmDisplayForm.AddEventLine("Operator Disconnected")
	End Sub

	Private Sub moOperator_onError(ByVal Index As Integer, ByVal Description As String) Handles moOperator.onError
		gfrmDisplayForm.AddEventLine("moOperator Error: " & Description)
	End Sub

	Private Sub moOperator_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moOperator.onSendComplete
		'do nothing
	End Sub

	Private Sub HandleOperatorDomainRegister(ByRef yData() As Byte)
		'going to contain a GUID list of all objects I am to control
		'  each object's children list is controlled by me too
		Dim lPos As Int32 = 2		'for msgcode
		Dim sPrimaryIP As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		Dim lPrimaryPort As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim oINI As New InitFile()

		'write our client listen port
		oINI.WriteString("SETTINGS", "ClientListenPort", "7710")
		'Write our primary data
		oINI.WriteString("PRIMARY", "PortNumber", lPrimaryPort.ToString)
		oINI.WriteString("PRIMARY", "IPAddress", sPrimaryIP)

		For X As Int32 = 0 To lCnt - 1
			Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			'Write the values to file
			oINI.WriteString("DOMAIN", "ID" & (X + 1).ToString, lID.ToString)
			oINI.WriteString("DOMAIN", "TypeID" & (X + 1).ToString, iTypeID.ToString)
		Next X

		gfrmDisplayForm.AddEventLine("Received Assignment from Operator")
		oINI = Nothing

		bReceivedAssignment = True
	End Sub

	Public Sub SendReadyStateToOperator()
		Dim yMsg(1) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eDomainServerReady).CopyTo(yMsg, 0)
		moOperator.SendData(yMsg)
	End Sub
#End Region

#Region "Primary Server Events - Domain Connects to Primary"
	Private Sub moPrimary_onConnect(ByVal Index As Integer) Handles moPrimary.onConnect
		'Index is required for backwards compatibility, on the Server, we do not use it
		gfrmDisplayForm.AddEventLine("Connected to Primary")
		mbConnectedToPrimary = True
	End Sub

	Private Sub moPrimary_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moPrimary.onConnectionRequest
		'Index is required for backwards compatibility, on the Server, we do not use it
	End Sub

	Private Sub moPrimary_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moPrimary.onDataArrival
		'Index is required for backwards compatibility, on the Server, we do not use it
		Dim iMsgID As Int16

		iMsgID = System.BitConverter.ToInt16(Data, 0)
		Select Case iMsgID
			Case GlobalMessageCode.eRouteMoveCommand
				HandleRouteMove(Data)
			Case GlobalMessageCode.ePathfindingConnectionInfo
				'Ok, the primary is telling us what our pathfinding info is
				gfrmDisplayForm.AddEventLine("Received Pathfinding Connection Info")
				ParsePathfindingConnectionInfo(Data)
			Case GlobalMessageCode.eRequestPlayerList
				gfrmDisplayForm.AddEventLine("End of Player List")
				mbSyncWait = False
			Case GlobalMessageCode.eRequestPlayerRelList
				gfrmDisplayForm.AddEventLine("End of Player Rel List")
				mbSyncWait = False
			Case GlobalMessageCode.eDomainRequestEnvirObjects
				gfrmDisplayForm.AddEventLine("End of Envir Object List")
				GeoSpawner.ReceivedEnvironmentObjects(Data)
				'mbSyncWait = False
			Case GlobalMessageCode.eAddObjectCommand, GlobalMessageCode.eAddObjectCommand_CE
				ReceiveAddObject(Data)
			Case GlobalMessageCode.eAddPlayerRelObject
				gfrmDisplayForm.AddEventLine("Add Player Rel Received")
				ReceiveAddPlayerRel(Data)
			Case GlobalMessageCode.eResponseGalaxyAndSystems
				gfrmDisplayForm.AddEventLine("Received Galaxy Map")
				HandleGalaxyAndSystemsMsg(Data)
				mbSyncWait = False
			Case GlobalMessageCode.eResponseStarTypes
				gfrmDisplayForm.AddEventLine("Received Star Types")
				HandleStarTypesMsg(Data)
				mbSyncWait = False
			Case GlobalMessageCode.eResponseSystemDetails
				gfrmDisplayForm.AddEventLine("Received System Details")
				HandleSystemDetailsMsg(Data)
                'mbSyncWait = False
            Case GlobalMessageCode.eUpdateCommandPoints
                HandlePrimaryUpdateCP(Data)
            Case GlobalMessageCode.eSetPlayerRel
                HandleSetPlayerRel(Data)
			Case GlobalMessageCode.eRequestDock
				HandleRequestDock(Data, moPrimary, -1)
			Case GlobalMessageCode.eSetMiningLoc
				HandleSetMiningLoc(Data, -1)
			Case GlobalMessageCode.eSetEntityProdSucceed
				HandleSetEntityProdSucceed(Data)
			Case GlobalMessageCode.eSetEntityStatus
				HandleSetEntityStatus(Data)
			Case GlobalMessageCode.eRemoveObject
				HandleRemoveObject(Data)
			Case GlobalMessageCode.eMoveObjectRequest
				HandleGroupMoveMessage(Data, -1)
			Case GlobalMessageCode.eRequestUndock
				HandleUndockCommand(Data)
			Case GlobalMessageCode.eDockCommand
				HandleDockCommand(Data)
			Case GlobalMessageCode.eReloadWpnMsg
				HandleReloadWpnMsg(Data)
			Case GlobalMessageCode.eRequestPlayerStartLoc
				HandleRequestPlayerStartLoc(Data)
			Case GlobalMessageCode.eServerShutdown
				gbRunning = False		'cause us to break out of running mode
			Case GlobalMessageCode.eSetFleetDest
				HandleSetFleetDest(Data)
			Case GlobalMessageCode.eFleetInterSystemMoving
				HandleFleetInterSystemMoving(Data)
			Case GlobalMessageCode.eGetPirateStartLoc
				HandleGetPirateStartLoc(Data)
			Case GlobalMessageCode.ePlacePirateAssets
				HandlePlacePirateAssets(Data)
			Case GlobalMessageCode.eMoveEngineer
				HandleMoveEngineer(Data)
			Case GlobalMessageCode.eSetRepairTarget, GlobalMessageCode.eSetDismantleTarget
				HandleSetMaintenanceTargetFromPrimary(Data, iMsgID)
			Case GlobalMessageCode.eRepairCompleted
				HandleRepairCompleted(Data)
			Case GlobalMessageCode.eUpdatePlayerTechValue
				HandleUpdatePlayerTechValue(Data)
			Case GlobalMessageCode.eSetEntityName
				HandleSetEntityName(Data)
			Case GlobalMessageCode.eEntityDefCriticalHitChances
				HandleEntityDefCriticalHitChances(Data)
			Case GlobalMessageCode.eSetCounterAttack
				HandleSetCounterAttack(Data)
			Case GlobalMessageCode.ePlayerDiscoversWormhole
				HandlePlayerDiscoversWormhole(Data)
			Case GlobalMessageCode.eJumpTarget
				HandlePrimaryJumpTarget(Data)
			Case GlobalMessageCode.ePlayerInitialSetup
				HandlePlayerInitialSetup(Data)
			Case GlobalMessageCode.ePlayerAliasConfig
				HandlePlayerAliasConfig(Data)
			Case GlobalMessageCode.eRemovePlayerRel
				HandleRemovePlayerRel(Data)
			Case GlobalMessageCode.eSetPlayerSpecialAttribute
				HandleSetPlayerSpecialAttribute(Data)
			Case GlobalMessageCode.eSetEntityProduction
				HandlePrimarySetEntityProduction(Data)
			Case GlobalMessageCode.eAddFormation
				HandleAddFormation(Data)
			Case GlobalMessageCode.eSetIronCurtain
				HandleSetIronCurtain(Data)
			Case GlobalMessageCode.eCreatePlanetInstance
				HandleCreatePlanetInstance(Data)
			Case GlobalMessageCode.eSaveAndUnloadInstance
				HandleSaveAndUnloadInstance(Data)
			Case GlobalMessageCode.eAILaunchAll
				HandleAILaunchAll(Data)
			Case GlobalMessageCode.eSetPrimaryTarget
				HandleSetPrimaryTarget(Data, -1)
			Case GlobalMessageCode.eForcedMoveSpeedMove
				HandleForcedMoveSpeedMove(Data)
			Case GlobalMessageCode.eRegisterDomain
                HandleRegisterDomain(Data)
            Case GlobalMessageCode.eUpdatePlayerDetails
                HandleUpdatePlayerDetails(Data)
            Case GlobalMessageCode.ePlayerIsDead
                HandlePlayerIsDead(Data)
            Case GlobalMessageCode.eRequestEntityDefenses
                HandleRequestEntityDefenses(Data)
            Case GlobalMessageCode.ePlayerConnectedPrimary
                HandlePlayerConnectedPrimary(Data) 
            Case GlobalMessageCode.eAdvancedEventConfig
                HandlePlanetColonyLimitMsg(Data)
            Case GlobalMessageCode.eUpdateUnitExpLevel
                HandleUpdateUnitExpLevel(Data)
        End Select
	End Sub

	Private Sub moPrimary_onDisconnect(ByVal Index As Integer) Handles moPrimary.onDisconnect
		'Index is required for backwards compatibility, on the Server, we do not use it
	End Sub

	Private Sub moPrimary_onError(ByVal Index As Integer, ByVal Description As String) Handles moPrimary.onError
		'Index is required for backwards compatibility, on the Server, we do not use it
		'MsgBox("Primary Socket Error: " & Description)
		If gfrmDisplayForm Is Nothing Then
			Debug.Write("Primary Socket Error: " & Description)
		Else
			gfrmDisplayForm.AddEventLine("Primary Socket Error: " & Description)
		End If
	End Sub

	Private Sub moPrimary_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moPrimary.onSendComplete
		'Index is required for backwards compatibility, on the Server, we do not use it
	End Sub
#End Region

#Region "Pathfinding Server Events - Domain Connects to Pathfinding"
	Private Sub moPathfinding_onConnect(ByVal Index As Integer) Handles moPathfinding.onConnect
		'Index is required for backwards compatibility, on the Server, we do not use it
		gfrmDisplayForm.AddEventLine("Pathfinding Connected")
		mbConnectedToPathfinding = True
	End Sub

	Private Sub moPathfinding_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moPathfinding.onConnectionRequest
		'Index is required for backwards compatibility, on the Server, we do not use it
	End Sub

	Private Sub moPathfinding_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moPathfinding.onDataArrival
		'Index is required for backwards compatibility, on the Server, we do not use it
		'we care about this one...
		Dim iMsgID As Int16

		If TotalBytes = 0 Then Return

		'sData will contain our message... here we will parse messages and handle them, etc...
		iMsgID = System.BitConverter.ToInt16(Data, 0)
		Select Case iMsgID
			Case GlobalMessageCode.eMoveObjectCommand, GlobalMessageCode.eFinalMoveCommand
				'gfrmDisplayForm.AddEventLine("Received Move Object Command From Pathfinding")
				HandleMoveCommand(Data)
			Case GlobalMessageCode.eEntityChangeEnvironment
				'gfrmDisplayForm.AddEventLine("Received Entity Change Environment From Pathfinding")
				HandleMoveCommand(Data)
			Case GlobalMessageCode.eFinalizeStopEvent
				HandleFinalizeStopEvent(Data)
			Case GlobalMessageCode.eMoveFormation
				HandlePFMoveFormation(Data)
		End Select
	End Sub

	Private Sub moPathfinding_onDisconnect(ByVal Index As Integer) Handles moPathfinding.onDisconnect
		'Index is required for backwards compatibility, on the Server, we do not use it
	End Sub

	Private Sub moPathfinding_onError(ByVal Index As Integer, ByVal Description As String) Handles moPathfinding.onError
		'Index is required for backwards compatibility, on the Server, we do not use it
	End Sub

	Private Sub moPathfinding_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moPathfinding.onSendComplete
		'Index is required for backwards compatibility, on the Server, we do not use it
	End Sub
#End Region

#Region "Client Listener Events - Listens for new Client Connections"
	Private Sub moClientListener_onConnect(ByVal Index As Integer) Handles moClientListener.onConnect
		'Not really used I don't think. Does an OnConnect get fired if I am the one being connected to? I dont thinkso
	End Sub

	Private Sub moClientListener_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moClientListener.onConnectionRequest
		Dim X As Int32
		Dim lIdx As Int32

		Try
			lIdx = -1
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
				'ReDim Preserve mlClientLastRequest(mlClientUB)
				ReDim Preserve muclientdata(mlclientub)
				ReDim Preserve mlAliasedAs(mlClientUB)
				ReDim Preserve mlAliasedRights(mlClientUB)
				lIdx = mlClientUB
			End If

			moClients(lIdx) = New NetSock(oClient)
			moClients(lIdx).SocketIndex = lIdx
			mlAliasedAs(lIdx) = -1
            mlAliasedRights(lIdx) = 0
            mlClientPlayer(lIdx) = -1

            If moClients(lIdx) Is Nothing Then
                If oClient Is Nothing = False Then oClient.Shutdown(Net.Sockets.SocketShutdown.Both)
                Return
            End If

			gfrmDisplayForm.AddEventLine("Client Connected " & lIdx)

			'Now, set up our event handlers...
			AddHandler moClients(lIdx).onConnect, AddressOf moClients_onConnect
			AddHandler moClients(lIdx).onDataArrival, AddressOf moClients_onDataArrival
			AddHandler moClients(lIdx).onDisconnect, AddressOf moClients_onDisconnect
            AddHandler moClients(lIdx).onError, AddressOf moClients_onError
            AddHandler moClients(lIdx).onSocketClosed, AddressOf moClients_onSocketClosed

			'And then tell the socket to begin listening for data to arrive
			moClients(lIdx).MakeReadyToReceive()
		Catch
			Try
				moClients(lIdx).Disconnect()
			Catch
			End Try
			moClients(lIdx) = Nothing
		End Try
	End Sub

	Private Sub moClientListener_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moClientListener.onDataArrival
		'Should never happen, the listener doesn't expect data...
	End Sub

	Private Sub moClientListener_onDisconnect(ByVal Index As Integer) Handles moClientListener.onDisconnect
		'shouldn't really happen either
	End Sub

	Private Sub moClientListener_onError(ByVal Index As Integer, ByVal Description As String) Handles moClientListener.onError
		'trap the error? Also, if a connection keeps getting established and broken over a threshold, perhaps
		' we should ban the IP until the user calls support???
	End Sub

	Private Sub moClientListener_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moClientListener.onSendComplete
		'
	End Sub
#End Region

#Region "Client Connections (moClients)"
	Private Sub moClients_onConnect(ByVal Index As Integer)
		'Don't really care
	End Sub

	Private Sub moClients_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer)
		'we care about this one...
		Dim iMsgID As Int16
		Dim yResponse() As Byte = Nothing
		Dim X As Int32

		'sData will contain our message... here we will parse messages and handle them, etc...
		iMsgID = System.BitConverter.ToInt16(Data, 0)

		If gb_MONITOR_MSGS = True Then
			goMsgMonitor.AddInMsg(iMsgID, MsgMonitor.eMM_AppType.ClientConnection, Data.Length, mlClientPlayer(Index))
		End If

		Select Case iMsgID
			Case GlobalMessageCode.eCameraPosUpdate
				HandleCameraPosUpdate(Data, Index)
			Case GlobalMessageCode.eMoveObjectRequest, GlobalMessageCode.eAddWaypointMsg
				HandleGroupMoveMessage(Data, Index) 
			Case GlobalMessageCode.eMoveFormation
				HandleMoveFormation(Data, Index)
			Case GlobalMessageCode.eBurstEnvironmentRequest
				'gfrmDisplayForm.AddEventLine("Received Burst Environment Request")
				'ok, send out the burst response
				For X = 0 To glPlayerUB
					If glPlayerIdx(X) = mlClientPlayer(Index) Then
						yResponse = CreateBurstEnvironmentResponse(Data, goPlayers(X))
						Exit For
					End If
				Next X
				If yResponse Is Nothing = False AndAlso yResponse.Length > 0 Then
					'moClients(Index).SendData(yResponse)
					If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eBurstEnvironmentRequest, MsgMonitor.eMM_AppType.ClientConnection, yResponse.Length, mlClientPlayer(Index))
					moClients(Index).SendLenAppendedData(yResponse)
				End If
				Erase yResponse
			Case GlobalMessageCode.eChangingEnvironment
				'gfrmDisplayForm.AddEventLine("Client Changing Environment")
				HandleChangeEnvironment(Data, Index)
			Case GlobalMessageCode.eUnitHPUpdate
				HandleClientRequestHPMsg(Data, Index)
			Case GlobalMessageCode.ePlayerLoginRequest
				'Ok, this server has a copy of the player, so let's find it
				'ok... association time
				'gfrmDisplayForm.AddEventLine("Received Player Login Message")
				HandleLoginRequest(Data, Index)
			Case GlobalMessageCode.eGetEntityDetails
				HandleGetEntityDetails(Data, Index)
			Case GlobalMessageCode.eSetPrimaryTarget
				HandleSetPrimaryTarget(Data, Index)
			Case GlobalMessageCode.eRequestDock
				HandleRequestDock(Data, moClients(Index), Index)
			Case GlobalMessageCode.eRequestUndock
				HandleRequestUndock(Data, Index)
			Case GlobalMessageCode.eSetMiningLoc
				HandleSetMiningLoc(Data, Index)
			Case GlobalMessageCode.eSetEntityProduction
				HandleSetEntityProduction(Data, Index)
			Case GlobalMessageCode.eRequestOrbitalBombard
				HandleRequestBombard(Data, Index)
			Case GlobalMessageCode.eStopOrbitalBombard
				HandleRequestStopBombard(Data, Index)
			Case GlobalMessageCode.eSetEntityAI
				HandleSetEntityAI(Data, Index)
			Case GlobalMessageCode.eChatMessage
				HandleChatMessage(Data, Index)
			Case GlobalMessageCode.eSetDismantleTarget, GlobalMessageCode.eSetRepairTarget
				HandleSetMaintenanceTarget(Data, Index, iMsgID)
			Case GlobalMessageCode.eJumpTarget
				HandleJumpTarget(Data, Index)
			Case GlobalMessageCode.eAlertDestinationReached
                HandleAlertDestReached(Data, Index)
            Case GlobalMessageCode.eSetEntityStatus
                HandleClientSetEntityStatus(Data, Index)
            Case GlobalMessageCode.eShiftClickAddProduction
                HandleShiftClickAddProduction(Data, Index)
            Case GlobalMessageCode.eSetTetherPoint
                HandleSetTetherPoint(Data, Index)
        End Select
	End Sub

	Private Sub moClients_onDisconnect(ByVal Index As Integer)
		On Error Resume Next
		'Dim X As Int32
		'Dim lIdx As Int32

		gfrmDisplayForm.AddEventLine("Client Disconnected " & Index)

		' Go thru, get our player object, indicate that they are offline on this server and remove them
		'  from the environment... furthermore, decrement the server's players in environment count for that envir

		If mlClientPlayer(Index) <> -1 Then
			'lIdx = -1
			'Dim bInstance As Boolean = False

            'RemovePlayerFromAllEnvirs(mlClientPlayer(Index), -1, -1)
			'For X = 0 To glPlayerUB
			'	If glPlayerIdx(X) = mlClientPlayer(Index) Then
			'		lIdx = goPlayers(X).lEnvirIdx
			'		bInstance = goPlayers(X).yPlayerPhase = 0
			'		goPlayers(X).lEnvirIdx = -1
			'		Exit For
			'	End If
			'Next X

			'If bInstance = True Then
			'	If lIdx > -1 AndAlso lIdx <= glInstanceUB Then
			'		For X = 0 To goInstances(lIdx).lPlayersInEnvirUB
			'			If goInstances(lIdx).lPlayersInEnvirIdx(X) = mlClientPlayer(Index) Then
			'				goInstances(lIdx).lPlayersInEnvirIdx(X) = -1
			'				goInstances(lIdx).oPlayersInEnvir(X) = Nothing
			'				Exit For
			'			End If
			'		Next X
			'	End If
			'Else
			'	If lIdx > -1 AndAlso lIdx <= glEnvirUB Then
			'		For X = 0 To goEnvirs(lIdx).lPlayersInEnvirUB
			'			If goEnvirs(lIdx).lPlayersInEnvirIdx(X) = mlClientPlayer(Index) Then
			'				goEnvirs(lIdx).lPlayersInEnvirIdx(X) = -1
			'				goEnvirs(lIdx).oPlayersInEnvir(X) = Nothing
			'				'Exit For
			'			End If
			'		Next X
			'	End If
			'End If
		End If

		'moClients(Index).Disconnect()	'to ensure it is disconnected
		moClients(Index) = Nothing
		mlClientPlayer(Index) = -1
		mlAliasedAs(Index) = -1
		mlAliasedRights(Index) = 0
	End Sub

	Private Sub moClients_onError(ByVal Index As Integer, ByVal Description As String)
		'Now, because we got an error, I don't care... drop em
		On Error Resume Next
		'Trap the error... we should do more here... perhaps stop people from connecting who are causing errors?
		gfrmDisplayForm.AddEventLine("Client Socket Error (" & Index & "): " & Description)

		If Description.ToLower.Contains("an existing connection was forcibly closed by the remote host") = True OrElse Err.Number = 5 Then
			moClients(Index).Disconnect()

			'Dim lIdx As Int32
			'Dim X As Int32

			If mlClientPlayer(Index) <> -1 Then
				'lIdx = -1

                'RemovePlayerFromAllEnvirs(mlClientPlayer(Index), -1, -1)

				'Dim bInstance As Boolean = False

				'For X = 0 To glPlayerUB
				'	If glPlayerIdx(X) = mlClientPlayer(Index) Then
				'		lIdx = goPlayers(X).lEnvirIdx
				'		bInstance = goPlayers(X).yPlayerPhase = 0
				'		goPlayers(X).lEnvirIdx = -1
				'		Exit For
				'	End If
				'Next X

				'If bInstance = True Then
				'	If lIdx > -1 AndAlso lIdx <= glInstanceUB Then
				'		For X = 0 To goInstances(lIdx).lPlayersInEnvirUB
				'			If goInstances(lIdx).lPlayersInEnvirIdx(X) = mlClientPlayer(Index) Then
				'				goInstances(lIdx).lPlayersInEnvirIdx(X) = -1
				'				goInstances(lIdx).oPlayersInEnvir(X) = Nothing
				'				'Exit For
				'			End If
				'		Next X
				'	End If
				'Else
				'	If lIdx > -1 AndAlso lIdx <= glEnvirUB Then
				'		For X = 0 To goEnvirs(lIdx).lPlayersInEnvirUB
				'			If goEnvirs(lIdx).lPlayersInEnvirIdx(X) = mlClientPlayer(Index) Then
				'				goEnvirs(lIdx).lPlayersInEnvirIdx(X) = -1
				'				goEnvirs(lIdx).oPlayersInEnvir(X) = Nothing
				'				'Exit For
				'			End If
				'		Next X
				'	End If
				'End If

			End If

			'moClients(Index).Disconnect()	'to ensure it is disconnected
			moClients(Index) = Nothing
			mlClientPlayer(Index) = -1
		End If


    End Sub

    Private Sub moClients_onSocketClosed(ByVal lIndex As Int32)
        moClients(lIndex) = Nothing
    End Sub
#End Region

#Region "Message Definitions and Handlers"
    Private Sub HandleUpdateUnitExpLevel(ByVal yData() As Byte)
        'System.BitConverter.GetBytes(GlobalMessageCode.eUpdateUnitExpLevel).CopyTo(yMsg, 0)
        'oUnit.GetGUIDAsString.CopyTo(yMsg, 2)
        'yMsg(8) = oUnit.ExpLevel
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim yNewVal As Byte = yData(8)

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        oEntity.Exp_Level = yNewVal
        oEntity.SetExpLevelMods()
    End Sub

    Private Sub HandlePlanetColonyLimitMsg(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim yVal As Byte = yData(lPos) : lPos += 1

        If iTypeID = ObjectType.ePlanet Then
            Dim lUB As Int32 = -1
            If goEnvirs Is Nothing = False Then lUB = Math.Min(glEnvirUB, goEnvirs.GetUpperBound(0))
            For X As Int32 = 0 To lUB
                If goEnvirs(X) Is Nothing = False AndAlso goEnvirs(X).ObjectID = lID AndAlso goEnvirs(X).ObjTypeID = iTypeID Then
                    goEnvirs(X).bEnvirAtColonyLimit = yVal <> 0
                    Exit For
                End If
            Next X

        End If
    End Sub

    Private Sub HandlePlayerConnectedPrimary(ByVal yData() As Byte)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lPrim As Int32 = System.BitConverter.ToInt32(yData, 6)

        If lPrim = Int32.MinValue Then
            'Ok, force disconnect the player
            Dim lUB As Int32 = -1
            If mlClientPlayer Is Nothing = False Then lUB = Math.Min(mlClientUB, mlClientPlayer.GetUpperBound(0))
            For X As Int32 = 0 To lUB
                If mlClientPlayer(X) = lPlayerID Then
                    Try
                        If moClients(X) Is Nothing = False Then
                            moClients(X).Disconnect()
                        End If
                    Catch
                    End Try
                    moClients(X) = Nothing
                End If
            Next X
        End If
    End Sub

    Private Sub HandlePrimaryUpdateCP(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lCPUsed As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lUB As Int32 = -1
        If goEnvirs Is Nothing = False Then lUB = Math.Min(glEnvirUB, goEnvirs.GetUpperBound(0))
        For X As Int32 = 0 To lUB
            If goEnvirs(X) Is Nothing = False Then
                If goEnvirs(X).ObjectID = lEnvirID AndAlso goEnvirs(X).ObjTypeID = iEnvirTypeID Then
                    Dim lTmpUB As Int32 = -1
                    If goEnvirs(X).lPlayersWhoHaveUnitsHereIdx Is Nothing = False Then lTmpUB = Math.Min(goEnvirs(X).lPlayersWhoHaveUnitsHereIdx.GetUpperBound(0), goEnvirs(X).lPlayersWhoHaveUnitsHereUB)
                    For Y As Int32 = 0 To goEnvirs(X).lPlayersWhoHaveUnitsHereUB
                        If goEnvirs(X).lPlayersWhoHaveUnitsHereIdx(Y) = lPlayerID Then
                            goEnvirs(X).lPlayersWhoHaveUnitsHereCP(Y) = lCPUsed
                            Exit For
                        End If
                    Next Y
                    Exit For
                End If
            End If
        Next X

    End Sub

    Private Sub HandleSetTetherPoint(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2

        Dim oCmdPlayer As Player = Nothing
        If lIndex > -1 Then
            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) = mlClientPlayer(lIndex) Then
                    oCmdPlayer = goPlayers(X)
                    Exit For
                End If
            Next X
            If oCmdPlayer Is Nothing Then Return
        End If

        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yAction As Byte = yData(lPos) : lPos += 1

        For X As Int32 = 0 To lCnt - 1
            Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            Dim oEntity As Epica_Entity = Nothing
            Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
            If lIdx > -1 Then oEntity = goEntity(lIdx)
            If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

            With oEntity
                If .lOwnerID <> mlClientPlayer(lIndex) AndAlso (.lOwnerID <> mlAliasedAs(lIndex) OrElse (mlAliasedRights(lIndex) And AliasingRights.eChangeBehavior) = 0) AndAlso _
                    (oCmdPlayer Is Nothing OrElse .Owner.lGuildID < 1 OrElse .Owner.lGuildID <> oCmdPlayer.lGuildID OrElse (.CurrentStatus And elUnitStatus.eGuildAsset) = 0) Then
                    gfrmDisplayForm.AddEventLine("Possible Cheat: Non-owner setting entity tether. PlayerID = " & mlClientPlayer(lIndex))
                    Return
                End If

                If yAction = 0 Then
                    'clearing
                    .lTetherPointX = Int32.MinValue
                    .lTetherPointZ = Int32.MinValue
                Else
                    'setting
                    If (.CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
                        .lTetherPointX = .LocX
                        .lTetherPointZ = .LocZ
                    End If
                End If
            End With
        Next X
        SendToPrimary(yData)
    End Sub

    Private Sub HandleRequestEntityDefenses(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        Dim oEntity As Epica_Entity = Nothing
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        Dim yResp(35) As Byte
        lPos = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestEntityDefenses).CopyTo(yResp, lPos) : lPos += 2
        oEntity.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
        For X As Int32 = 0 To 3
            System.BitConverter.GetBytes(oEntity.ArmorHP(X)).CopyTo(yResp, lPos) : lPos += 4
        Next X
        System.BitConverter.GetBytes(oEntity.StructuralHP).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(oEntity.ShieldHP).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(oEntity.CurrentStatus).CopyTo(yResp, lPos) : lPos += 4

        moPrimary.SendData(yResp)

    End Sub

    Private Sub HandleUpdatePlayerDetails(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lCurUB As Int32 = -1
        If glPlayerIdx Is Nothing = False Then lCurUB = Math.Min(glPlayerUB, glPlayerIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If glPlayerIdx(X) = lPlayerID Then
                Dim oPlayer As Player = goPlayers(X)
                If oPlayer Is Nothing = False Then
                    ReDim oPlayer.PlayerUserName(19)
                    Array.Copy(yData, lPos, oPlayer.PlayerUserName, 0, 20) : lPos += 20
                    ReDim oPlayer.PlayerPassword(19)
                    Array.Copy(yData, lPos, oPlayer.PlayerPassword, 0, 20) : lPos += 20
                    oPlayer.AccountStatus = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    oPlayer.PlayerName = oPlayer.PlayerUserName
                End If
                Exit For
            End If
        Next X

    End Sub

    Private Sub HandleForcedMoveSpeedMove(ByRef yData() As Byte)
        Dim lPos As Int32 = 2 'for msgcode

        Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDestA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'need to go through the guid list and append the current loc data... also, accumulate the minimum speed and store that at the specified loc
        Dim yForward(22 + (lCnt * 24)) As Byte
        Dim lDestPos As Int32 = 0
        Dim lSpeedPos As Int32 = 0
        Dim yMinSpeed As Byte = 255

        System.BitConverter.GetBytes(GlobalMessageCode.eForcedMoveSpeedMove).CopyTo(yForward, lDestPos) : lDestPos += 2
        System.BitConverter.GetBytes(lDestX).CopyTo(yForward, lDestPos) : lDestPos += 4
        System.BitConverter.GetBytes(lDestZ).CopyTo(yForward, lDestPos) : lDestPos += 4
        System.BitConverter.GetBytes(iDestA).CopyTo(yForward, lDestPos) : lDestPos += 2
        System.BitConverter.GetBytes(lEnvirID).CopyTo(yForward, lDestPos) : lDestPos += 4
        System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yForward, lDestPos) : lDestPos += 2
        lSpeedPos = lDestPos
        lDestPos += 1

        System.BitConverter.GetBytes(lCnt).CopyTo(yForward, lDestPos) : lDestPos += 4

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))

        For lItem As Int32 = 0 To lCnt - 1
            Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            Dim oEntity As Epica_Entity = Nothing
            Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
            If lIdx > -1 Then oEntity = goEntity(lIdx)
            If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Continue For

            ''Now, go thru our entity list
            'For X As Int32 = 0 To lCurUB
            '    If glEntityIdx(X) = lObjID Then
            '        'goEntity(X).ObjTypeID = iObjTypeID Then
            '        Dim oEntity As Epica_Entity = goEntity(X)
            '        If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iObjTypeID Then

            'Got our entity
            With oEntity
                If .Owner Is Nothing = False Then
                    If ((.CurrentStatus And elUnitStatus.eEngineOperational) <> 0) AndAlso .MaxSpeed > 0 AndAlso .Acceleration > 0 Then
                        .bPlayerMoveRequestPending = True
                        .bAIMoveRequestPending = False

                        If yMinSpeed > .MaxSpeed Then yMinSpeed = .MaxSpeed

                        .GetGUIDAsString.CopyTo(yForward, lDestPos) : lDestPos += 6
                        System.BitConverter.GetBytes(.LocX).CopyTo(yForward, lDestPos) : lDestPos += 4
                        System.BitConverter.GetBytes(.LocZ).CopyTo(yForward, lDestPos) : lDestPos += 4
                        System.BitConverter.GetBytes(.LocAngle).CopyTo(yForward, lDestPos) : lDestPos += 2

                        If .ParentEnvir Is Nothing = False Then .ParentEnvir.GetGUIDAsString.CopyTo(yForward, lDestPos)
                        lDestPos += 6

                        If .lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(.lEntityDefServerIndex) <> -1 Then
                            System.BitConverter.GetBytes(goEntityDefs(.lEntityDefServerIndex).ModelID).CopyTo(yForward, lDestPos) : lDestPos += 2
                        Else
                            System.BitConverter.GetBytes(CShort(-1)).CopyTo(yForward, lDestPos) : lDestPos += 2
                        End If
                    End If
                End If
            End With
            '            Exit For
            '        End If
            '    End If
            'Next X
        Next lItem

        yForward(lSpeedPos) = yMinSpeed

        moPathfinding.SendData(yForward)
    End Sub

    Private Sub HandleRouteMove(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        'k... here we go...
        Dim oEntity As Epica_Entity = Nothing
        'Dim lCurUB As Int32 = glEntityUB
        'If glEntityIdx Is Nothing = False Then lCurUB = Math.Min(lCurUB, glEntityIdx.GetUpperBound(0))
        'If goEntity Is Nothing = False Then lCurUB = Math.Min(lCurUB, goEntity.GetUpperBound(0))

        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)

        'For X As Int32 = 0 To lCurUB
        '    If glEntityIdx(X) = lObjID Then
        '        Dim oTmp As Epica_Entity = goEntity(X)
        '        If oTmp Is Nothing = False AndAlso oTmp.ObjTypeID = iObjTypeID Then
        '            oEntity = oTmp
        '            Exit For
        '        End If
        '    End If
        'Next X
        If oEntity Is Nothing Then
            'Send the request back to the Primary indicating that we need more time
            moPrimary.SendData(yData)
            Return
        End If
        If oEntity.ParentEnvir Is Nothing Then Return
        If oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        Dim yForward() As Byte = Nothing

        If iTargetTypeID = ObjectType.eWormhole Then
            'Ok, the parent is the next item
            Dim lJumpTargetEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iJumpTargetEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            If lJumpTargetEnvirID <> oEntity.ParentEnvir.ObjectID OrElse iJumpTargetEnvirTypeID <> oEntity.ParentEnvir.ObjTypeID Then Return

            'Ok, gotta treat this like a JumpTarget command
            If iJumpTargetEnvirTypeID = ObjectType.eSolarSystem Then
                Dim lDestX As Int32 = Int32.MinValue
                Dim lDestZ As Int32 = Int32.MinValue
                Dim iDestA As Int16 = 0
                For X As Int32 = 0 To goGalaxy.mlSystemUB
                    If goGalaxy.moSystems(X).ObjectID = lJumpTargetEnvirID Then
                        For Y As Int32 = 0 To goGalaxy.moSystems(X).WormholeUB
                            If goGalaxy.moSystems(X).moWormholes(Y).ObjectID = lTargetID Then
                                With goGalaxy.moSystems(X).moWormholes(Y)
                                    If .System1 Is Nothing = False AndAlso .System1.ObjectID = lJumpTargetEnvirID Then
                                        lDestX = .LocX1
                                        lDestZ = .LocY1
                                        iDestA = 0
                                    Else
                                        lDestX = .LocX2
                                        lDestZ = .LocY2
                                        iDestA = 0
                                    End If
                                End With
                                Exit For
                            End If
                        Next Y
                        Exit For
                    End If
                Next X

                ReDim yForward(47)
                Dim lDestPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eJumpTarget).CopyTo(yForward, lDestPos) : lDestPos += 2
                System.BitConverter.GetBytes(lDestX).CopyTo(yForward, lDestPos) : lDestPos += 4
                System.BitConverter.GetBytes(lDestZ).CopyTo(yForward, lDestPos) : lDestPos += 4
                System.BitConverter.GetBytes(iDestA).CopyTo(yForward, lDestPos) : lDestPos += 2
                System.BitConverter.GetBytes(lJumpTargetEnvirID).CopyTo(yForward, lDestPos) : lDestPos += 4
                System.BitConverter.GetBytes(iJumpTargetEnvirTypeID).CopyTo(yForward, lDestPos) : lDestPos += 2
                System.BitConverter.GetBytes(lTargetID).CopyTo(yForward, lDestPos) : lDestPos += 4
                System.BitConverter.GetBytes(iTargetTypeID).CopyTo(yForward, lDestPos) : lDestPos += 2

                oEntity.GetGUIDAsString.CopyTo(yForward, lDestPos) : lDestPos += 6
                System.BitConverter.GetBytes(oEntity.LocX).CopyTo(yForward, lDestPos) : lDestPos += 4
                System.BitConverter.GetBytes(oEntity.LocZ).CopyTo(yForward, lDestPos) : lDestPos += 4
                System.BitConverter.GetBytes(oEntity.LocAngle).CopyTo(yForward, lDestPos) : lDestPos += 2
                oEntity.ParentEnvir.GetGUIDAsString.CopyTo(yForward, lDestPos) : lDestPos += 6

                If oEntity.lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(oEntity.lEntityDefServerIndex) <> -1 Then
                    System.BitConverter.GetBytes(goEntityDefs(oEntity.lEntityDefServerIndex).ModelID).CopyTo(yForward, lDestPos) : lDestPos += 2
                Else
                    System.BitConverter.GetBytes(CShort(-1)).CopyTo(yForward, lDestPos) : lDestPos += 2
                End If

            Else
                gfrmDisplayForm.AddEventLine("Possible Cheat: Wormhole Parent Envir is not Solar System! In a Route.")
                Return
            End If
        Else
            Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lTargetParentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iTargetParentTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            'Now, is our target type something?
            If iTargetTypeID = ObjectType.eUnit OrElse iTargetTypeID = ObjectType.eFacility Then
                Dim oTarget As Epica_Entity = Nothing
                lIdx = LookupEntity(lTargetID, iTargetTypeID)
                If lIdx > -1 Then oTarget = goEntity(lIdx)

                'For X As Int32 = 0 To lCurUB
                '    If glEntityIdx(X) = lTargetID Then
                '        Dim oTmp As Epica_Entity = goEntity(X)
                '        If oTmp Is Nothing = False AndAlso oTmp.ObjTypeID = iTargetTypeID Then
                '            oTarget = oTmp
                '            Exit For
                '        End If
                '    End If
                'Next X
                If oTarget Is Nothing Then Return
                If oTarget.ObjectID <> lTargetID OrElse oTarget.ObjTypeID <> iTargetTypeID Then Return
                If oTarget.ParentEnvir Is Nothing = False Then
                    If oTarget.ParentEnvir.ObjectID <> lTargetParentID OrElse oTarget.ParentEnvir.ObjTypeID <> iTargetParentTypeID Then
                        lTargetParentID = oTarget.ParentEnvir.ObjectID
                        iTargetParentTypeID = oTarget.ParentEnvir.ObjTypeID

                        lLocX = oTarget.LocX
                        lLocZ = oTarget.LocZ
                    End If
                End If
            End If

            'Ok, now, we forward the message to the pathfinding server to work on...
            'MsgCode (2)
            'ObjGUID (6)
            'Object's Loc (8)
            'Object's Parent (6)
            'Dest (8)
            'Dest GUID (6)
            'Target GUID (6)

            ReDim yForward(43)
            lPos = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eRouteMoveCommand).CopyTo(yForward, lPos) : lPos += 2
            oEntity.GetGUIDAsString.CopyTo(yForward, lPos) : lPos += 6
            If oEntity.lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(oEntity.lEntityDefServerIndex) <> -1 Then
                System.BitConverter.GetBytes(goEntityDefs(oEntity.lEntityDefServerIndex).ModelID).CopyTo(yForward, lPos) : lPos += 2
            Else : System.BitConverter.GetBytes(-1S).CopyTo(yForward, lPos) : lPos += 2
            End If
            System.BitConverter.GetBytes(oEntity.LocX).CopyTo(yForward, lPos) : lPos += 4
            System.BitConverter.GetBytes(oEntity.LocZ).CopyTo(yForward, lPos) : lPos += 4
            oEntity.ParentEnvir.GetGUIDAsString.CopyTo(yForward, lPos) : lPos += 6
            System.BitConverter.GetBytes(lLocX).CopyTo(yForward, lPos) : lPos += 4
            System.BitConverter.GetBytes(lLocZ).CopyTo(yForward, lPos) : lPos += 4
            System.BitConverter.GetBytes(lTargetParentID).CopyTo(yForward, lPos) : lPos += 4
            System.BitConverter.GetBytes(iTargetParentTypeID).CopyTo(yForward, lPos) : lPos += 2
            System.BitConverter.GetBytes(lTargetID).CopyTo(yForward, lPos) : lPos += 4
            System.BitConverter.GetBytes(iTargetTypeID).CopyTo(yForward, lPos) : lPos += 2

        End If

        If yForward Is Nothing Then Return

        oEntity.bPlayerMoveRequestPending = False
        oEntity.bAIMoveRequestPending = True

        moPathfinding.SendData(yForward)

    End Sub

    Private Sub HandleSetPrimaryTarget(ByVal yData() As Byte, ByVal lIndex As Int32)
        'If glCurrentCycle - mlClientLastRequest(lIndex) > ml_REQUEST_THROTTLE Then
        If lIndex <> -1 Then
            If glCurrentCycle - muClientData(lIndex).mlClientLastRequest > ml_REQUEST_THROTTLE Then
                'mlClientLastRequest(lIndex) = glCurrentCycle
                muClientData(lIndex).mlClientLastRequest = glCurrentCycle
            Else : Return
            End If
        End If

        'Msg Code (2), TargetID (4), TargetTypeID (2), GUIDList...
        Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        Dim lObjID As Int32
        Dim iObjTypeID As Int16

        Dim lPos As Int32 = 8
        Dim lDestPos As Int32
        Dim lDestLen As Int32 = -1

        Dim X As Int32

        Dim oTarget As Epica_Entity = Nothing
        Dim lRng As Int32

        Dim yGridResult As Byte
        Dim lTemp As Int32
        Dim lRelTinyX As Int32
        Dim lRelTinyZ As Int32
        Dim lDist As Int32

        Dim oCmdPlayer As Player = Nothing
        If lIndex > -1 Then
            For X = 0 To glPlayerUB
                If glPlayerIdx(X) = mlClientPlayer(lIndex) Then
                    oCmdPlayer = goPlayers(X)
                    Exit For
                End If
            Next X
            If oCmdPlayer Is Nothing Then Return
        End If

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X = 0 To lCurUB
        '    If glEntityIdx(X) = lTargetID AndAlso goEntity(X).ObjTypeID = iTargetTypeID Then
        '        oTarget = goEntity(X)
        '        Exit For
        '    End If
        'Next X
        Dim lIdx As Int32 = LookupEntity(lTargetID, iTargetTypeID)
        If lIdx > -1 Then oTarget = goEntity(lIdx)
        If oTarget Is Nothing OrElse oTarget.ObjectID <> lTargetID OrElse oTarget.ObjTypeID <> iTargetTypeID Then Return

        If oTarget Is Nothing = False Then

            'Ok, set up our message header...
            '  MsgCode (2), TargetGUID (6), TargetLoc (10), TargetModelId (2), GUIDListCnt (4)
            Dim lCntPos As Int32 = 0
            Dim lEntCnt As Int32 = 0

            Dim yMoveMsg(39) As Byte
            Dim lMoveMsgUB As Int32 = 39
            System.BitConverter.GetBytes(GlobalMessageCode.eSetPrimaryTarget).CopyTo(yMoveMsg, lDestPos) : lDestPos += 2
            oTarget.GetGUIDAsString.CopyTo(yMoveMsg, lDestPos) : lDestPos += 6
            System.BitConverter.GetBytes(oTarget.LocX).CopyTo(yMoveMsg, lDestPos) : lDestPos += 4
            System.BitConverter.GetBytes(oTarget.LocZ).CopyTo(yMoveMsg, lDestPos) : lDestPos += 4
            System.BitConverter.GetBytes(oTarget.LocAngle).CopyTo(yMoveMsg, lDestPos) : lDestPos += 2
            If oTarget.lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(oTarget.lEntityDefServerIndex) <> -1 Then
                System.BitConverter.GetBytes(goEntityDefs(oTarget.lEntityDefServerIndex).ModelID).CopyTo(yMoveMsg, lDestPos) : lDestPos += 2
            Else
                System.BitConverter.GetBytes(CShort(-1)).CopyTo(yMoveMsg, lDestPos) : lDestPos += 2
            End If
            oTarget.ParentEnvir.GetGUIDAsString.CopyTo(yMoveMsg, lDestPos) : lDestPos += 6
            lCntPos = lDestPos
            lDestPos += 4

            While lPos + 6 < yData.Length '- 1
                lObjID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                Dim oEntity As Epica_Entity = Nothing
                lIdx = LookupEntity(lObjID, iObjTypeID)
                If lIdx > -1 Then oEntity = goEntity(lIdx)
                If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Continue While

                'For X = 0 To lCurUB
                '    If glEntityIdx(X) = lObjID AndAlso goEntity(X).ObjTypeID = iObjTypeID Then
                '        Dim oEntity As Epica_Entity = goEntity(X)
                '        If oEntity Is Nothing Then Exit For
                With oEntity

                    If lIndex <> -1 AndAlso .lOwnerID <> mlClientPlayer(lIndex) AndAlso (.lOwnerID <> mlAliasedAs(lIndex) OrElse (mlAliasedRights(lIndex) And AliasingRights.eMoveUnits) = 0) AndAlso _
                        (oCmdPlayer Is Nothing OrElse .Owner.lGuildID <> oCmdPlayer.lGuildID OrElse .Owner.lGuildID = -1 OrElse (.CurrentStatus And elUnitStatus.eGuildAsset) = 0) Then Continue While
                    If .ParentEnvir.ObjectID <> oTarget.ParentEnvir.ObjectID OrElse .ParentEnvir.ObjTypeID <> oTarget.ParentEnvir.ObjTypeID Then
                        If lIndex <> -1 Then
                            gfrmDisplayForm.AddEventLine("Possible cheat: SetPrimaryTarget Envirs Mismatch. PlayerID: " & mlClientPlayer(lIndex))
                        End If
                        Continue While
                    End If
                    'If .lPrimaryTargetServerIdx = oTarget.ServerIndex Then Exit For
                    If (.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then
                        'send a move lock reset msg
                        Dim yMsg(7) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eMoveLockViolate).CopyTo(yMsg, 0)
                        .GetGUIDAsString.CopyTo(yMsg, 2)
                        SendToPrimary(yMsg)
                        .CurrentStatus = .CurrentStatus Xor elUnitStatus.eMoveLock ' oEntity.CurrentStatus -= elUnitStatus.eMoveLock
                    End If

                    lMoveMsgUB += 28
                    ReDim Preserve yMoveMsg(lMoveMsgUB)

                    lEntCnt += 1
                    .GetGUIDAsString.CopyTo(yMoveMsg, lDestPos) : lDestPos += 6
                    System.BitConverter.GetBytes(.LocX).CopyTo(yMoveMsg, lDestPos) : lDestPos += 4
                    System.BitConverter.GetBytes(.LocZ).CopyTo(yMoveMsg, lDestPos) : lDestPos += 4
                    System.BitConverter.GetBytes(.LocAngle).CopyTo(yMoveMsg, lDestPos) : lDestPos += 2

                    'Assign the target...
                    .CurrentStatus = .CurrentStatus Or elUnitStatus.eTargetingByPlayer
                    .lPrimaryTargetServerIdx = oTarget.ServerIndex
                    '.lTetherPointX = Int32.MinValue
                    '.lTetherPointZ = Int32.MinValue
                    .bForceAggressionTest = True
                    SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)
                    oTarget.AddTargetedBy(.ServerIndex, .ObjectID)

                    'Now, create a move message to get us close to that unit (if I am not)
                    'this is easy... 
                    lRng = GetPreferredRange(oEntity) '\ gl_FINAL_GRID_SQUARE_SIZE

                    'check if I am already in range of the target...
                    lDist = Int32.MaxValue
                    yGridResult = gyLargeGridArray(oTarget.lGridIndex, .lGridIndex, .ParentEnvir.lGridsPerRow)
                    If yGridResult <> 255 Then
                        lTemp = giRelativeSmall(yGridResult, .lSmallSectorID, oTarget.lSmallSectorID)
                        If lTemp <> -1 Then
                            lRelTinyX = glBaseRelTinyX(lTemp) + oTarget.lTinyX + (gl_HALF_FINAL_PER_ROW - .lTinyX)
                            If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                lRelTinyZ = glBaseRelTinyZ(lTemp) + oTarget.lTinyZ + (gl_HALF_FINAL_PER_ROW - .lTinyZ)
                                If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                    lDist = gyDistances(lRelTinyX, lRelTinyZ) - .lModelRangeOffset
                                End If
                            End If
                        End If
                    End If

                    Dim lSide As Int32 = GetPreferredSideFacing(oEntity)
                    Dim lCurSide As Int32 = lSide
                    If lDist <> Int32.MaxValue Then
                        lCurSide = WhatSideCanFire(oEntity, lRelTinyX, lRelTinyZ)
                    End If

                    If lDist <> Int32.MaxValue AndAlso (lDist > lRng) AndAlso .Maneuver <> 0 AndAlso .MaxSpeed <> 0 Then 'If lDist <> Int32.MaxValue AndAlso (lDist > lRng OrElse lSide <> lCurSide) AndAlso .Maneuver <> 0 AndAlso .MaxSpeed <> 0 Then
                        'ok, we're not in range... now, get line angle degrees
                        Dim fAngle As Single = LineAngleDegrees(oTarget.LocX, oTarget.LocZ, .LocX, .LocZ)
                        Dim lTX As Int32 = oTarget.LocX + (lRng * gl_FINAL_GRID_SQUARE_SIZE)
                        Dim lTZ As Int32 = oTarget.LocZ
                        RotatePoint(oTarget.LocX, oTarget.LocZ, lTX, lTZ, fAngle)

                        'Now, determine our final dest's side angle
                        lTemp = (lSide - lCurSide) * 900
                        lTemp += CShort(fAngle)
                        If lTemp < 0 Then lTemp += 3600
                        If lTemp > 3599 Then lTemp -= 3600
                        lTemp -= 1800
                        If lTemp < 0 Then lTemp += 3600
                        If lTemp > 3599 Then lTemp -= 3600
                        System.BitConverter.GetBytes(lTX).CopyTo(yMoveMsg, lDestPos) : lDestPos += 4
                        System.BitConverter.GetBytes(lTZ).CopyTo(yMoveMsg, lDestPos) : lDestPos += 4
                        System.BitConverter.GetBytes(CShort(lTemp)).CopyTo(yMoveMsg, lDestPos) : lDestPos += 2

                        'AI Move sets these, we need to unset them
                        .bPlayerMoveRequestPending = True
                        .bAIMoveRequestPending = False
                    Else
                        System.BitConverter.GetBytes(.LocX).CopyTo(yMoveMsg, lDestPos) : lDestPos += 4
                        System.BitConverter.GetBytes(.LocZ).CopyTo(yMoveMsg, lDestPos) : lDestPos += 4
                        System.BitConverter.GetBytes(.LocAngle).CopyTo(yMoveMsg, lDestPos) : lDestPos += 2
                    End If

                    If .lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(.lEntityDefServerIndex) <> -1 Then
                        System.BitConverter.GetBytes(goEntityDefs(.lEntityDefServerIndex).ModelID).CopyTo(yMoveMsg, lDestPos) : lDestPos += 2
                    Else
                        System.BitConverter.GetBytes(CShort(-1)).CopyTo(yMoveMsg, lDestPos) : lDestPos += 2
                    End If

                End With

                '        Exit For
                '    End If
                'Next X
            End While

            System.BitConverter.GetBytes(lEntCnt).CopyTo(yMoveMsg, lCntPos)

            If yMoveMsg Is Nothing = False AndAlso yMoveMsg.Length > 0 Then moPathfinding.SendData(yMoveMsg)

        End If

    End Sub
    'Private Sub HandleSetPrimaryTarget(ByVal yData() As Byte, ByVal lIndex As Int32)
    '	'If glCurrentCycle - mlClientLastRequest(lIndex) > ml_REQUEST_THROTTLE Then
    '	If glCurrentCycle - muClientData(lIndex).mlClientLastRequest > ml_REQUEST_THROTTLE Then
    '		'mlClientLastRequest(lIndex) = glCurrentCycle
    '		muClientData(lIndex).mlClientLastRequest = glCurrentCycle
    '	Else : Return
    '	End If

    '	'Msg Code (2), TargetID (4), TargetTypeID (2), GUIDList...
    '	Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, 2)
    '	Dim iTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

    '	Dim lObjID As Int32
    '	Dim iObjTypeID As Int16

    '	Dim lPos As Int32 = 8
    '	Dim lDestPos As Int32
    '	Dim lDestLen As Int32 = -1

    '	Dim X As Int32

    '	Dim oTarget As Epica_Entity = Nothing
    '	Dim lRng As Int32

    '	Dim yGridResult As Byte
    '	Dim lTemp As Int32
    '	Dim lRelTinyX As Int32
    '	Dim lRelTinyZ As Int32
    '	Dim lDist As Int32

    '	Dim yMoveMsg() As Byte = Nothing

    '	Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
    '	For X = 0 To lCurUB
    '		If glEntityIdx(X) = lTargetID AndAlso goEntity(X).ObjTypeID = iTargetTypeID Then
    '			oTarget = goEntity(X)
    '			Exit For
    '		End If
    '	Next X

    '	If oTarget Is Nothing = False Then
    '		While lPos + 6 < yData.Length '- 1
    '			lObjID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '			iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

    '			For X = 0 To lCurUB
    '				If glEntityIdx(X) = lObjID AndAlso goEntity(X).ObjTypeID = iObjTypeID Then
    '					Dim oEntity As Epica_Entity = goEntity(X)
    '					If oEntity Is Nothing Then Exit For
    '					With oEntity

    '						If .lOwnerID <> mlClientPlayer(lIndex) AndAlso (.lOwnerID <> mlAliasedAs(lIndex) OrElse (mlAliasedRights(lIndex) And AliasingRights.eMoveUnits) = 0) Then Exit For

    '						'Assign the target...
    '						.CurrentStatus = .CurrentStatus Or elUnitStatus.eTargetingByPlayer
    '						.lPrimaryTargetServerIdx = oTarget.ServerIndex
    '						'If oTarget.lOwnerID = .lOwnerID Then
    '						'	gfrmDisplayForm.AddEventLine("Shooting at myself: SetPrimaryTarget.Primary")
    '						'End If
    '						.bForceAggressionTest = True
    '						AddEntityMoving(.ServerIndex, .ObjectID)
    '						oTarget.AddTargetedBy(.ServerIndex)

    '						'Now, create a move message to get us close to that unit (if I am not)
    '						'this is easy... 
    '						lRng = GetPreferredRange(oEntity) '\ gl_FINAL_GRID_SQUARE_SIZE

    '						'check if I am already in range of the target...
    '						yGridResult = gyLargeGridArray(oTarget.lGridIndex, .lGridIndex, .ParentEnvir.lGridsPerRow)
    '						If yGridResult <> 255 Then
    '							lTemp = giRelativeSmall(yGridResult, .lSmallSectorID, oTarget.lSmallSectorID)
    '							If lTemp <> -1 Then
    '								lRelTinyX = glBaseRelTinyX(lTemp) + oTarget.lTinyX + (gl_HALF_FINAL_PER_ROW - .lTinyX)
    '								If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
    '									lRelTinyZ = glBaseRelTinyZ(lTemp) + oTarget.lTinyZ + (gl_HALF_FINAL_PER_ROW - .lTinyZ)
    '									If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
    '										lDist = gyDistances(lRelTinyX, lRelTinyZ) - .lModelRangeOffset
    '									End If
    '								End If
    '							End If
    '						End If

    '						If lDist > lRng AndAlso .Maneuver <> 0 AndAlso .MaxSpeed <> 0 Then
    '							'ok, we're not in range... now, get line angle degrees
    '							Dim fAngle As Single = LineAngleDegrees(oTarget.LocX, oTarget.LocZ, .LocX, .LocZ)
    '							Dim lTX As Int32 = oTarget.LocX + (lRng * gl_FINAL_GRID_SQUARE_SIZE)
    '							Dim lTZ As Int32 = oTarget.LocZ
    '							RotatePoint(oTarget.LocX, oTarget.LocZ, lTX, lTZ, fAngle)

    '							lDestLen += 44
    '							ReDim Preserve yMoveMsg(lDestLen)
    '							System.BitConverter.GetBytes(CShort(42)).CopyTo(yMoveMsg, lDestPos) : lDestPos += 2
    '							System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequest).CopyTo(yMoveMsg, lDestPos) : lDestPos += 2
    '							System.BitConverter.GetBytes(lTX).CopyTo(yMoveMsg, lDestPos) : lDestPos += 4
    '							System.BitConverter.GetBytes(lTZ).CopyTo(yMoveMsg, lDestPos) : lDestPos += 4
    '							System.BitConverter.GetBytes(CShort(fAngle)).CopyTo(yMoveMsg, lDestPos) : lDestPos += 2

    '							.ParentEnvir.GetGUIDAsString.CopyTo(yMoveMsg, lDestPos) : lDestPos += 6
    '							.GetGUIDAsString.CopyTo(yMoveMsg, lDestPos) : lDestPos += 6

    '							System.BitConverter.GetBytes(.LocX).CopyTo(yMoveMsg, lDestPos) : lDestPos += 4
    '							System.BitConverter.GetBytes(.LocZ).CopyTo(yMoveMsg, lDestPos) : lDestPos += 4
    '							System.BitConverter.GetBytes(.LocAngle).CopyTo(yMoveMsg, lDestPos) : lDestPos += 2
    '							.ParentEnvir.GetGUIDAsString.CopyTo(yMoveMsg, lDestPos) : lDestPos += 6

    '							If .lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(.lEntityDefServerIndex) <> -1 Then
    '								System.BitConverter.GetBytes(goEntityDefs(.lEntityDefServerIndex).ModelID).CopyTo(yMoveMsg, lDestPos) : lDestPos += 2
    '							Else
    '								System.BitConverter.GetBytes(CShort(-1)).CopyTo(yMoveMsg, lDestPos) : lDestPos += 2
    '							End If

    '							'AI Move sets these, we need to unset them
    '							.bPlayerMoveRequestPending = True
    '							.bAIMoveRequestPending = False

    '						End If
    '					End With

    '					Exit For
    '				End If
    '			Next X
    '		End While

    '		If yMoveMsg Is Nothing = False AndAlso yMoveMsg.Length > 0 Then moPathfinding.SendLenAppendedData(yMoveMsg)

    '	End If

    'End Sub

    Private Sub HandleGetEntityDetails(ByVal yData() As Byte, ByVal lClientIndex As Int32)
        'If glCurrentCycle - mlClientLastRequest(lClientIndex) > ml_REQUEST_THROTTLE Then
        'If glCurrentCycle - muClientData(lClientIndex).mlLastGetEntityDetails > ml_REQUEST_THROTTLE Then
        '    'mlClientLastRequest(lClientIndex) = glCurrentCycle
        '    muClientData(lClientIndex).mlLastGetEntityDetails = glCurrentCycle
        'Else : Return
        'End If

        'msg code (2), ID (4), typeid (2), lenvir id (4), i envir typeid (2)
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
        Dim X As Int32
        'Dim Y As Int32

        Dim oEnvir As Envir = Nothing

        Dim yResp() As Byte = Nothing
        Dim yTemp() As Byte = Nothing

        Dim lTemp As Int32

        'Find our environment
        If iEnvirTypeID = ObjectType.ePlanet AndAlso lEnvirID >= 500000000 Then
            'instance
            Dim lCurUB As Int32 = -1
            If glInstanceIdx Is Nothing = False Then lCurUB = Math.Min(glInstanceUB, glInstanceIdx.GetUpperBound(0))
            For X = 0 To lCurUB
                If glInstanceIdx(X) = lEnvirID Then
                    oEnvir = goInstances(X)
                    Exit For
                End If
            Next X
        Else
            For X = 0 To glEnvirUB
                If goEnvirs(X).ObjectID = lEnvirID AndAlso goEnvirs(X).ObjTypeID = iEnvirTypeID Then
                    oEnvir = goEnvirs(X)
                    Exit For
                End If
            Next X
        End If


        'Not all objects will require an environment to be found...

        Select Case iTypeID
            Case ObjectType.eMineralCache, ObjectType.eComponentCache
                If oEnvir Is Nothing = False Then
                    For X = 0 To oEnvir.lCacheUB
                        If oEnvir.lCacheIdx(X) = lID Then
                            yTemp = oEnvir.oCache(X).GetObjectAsString
                            Exit For
                        End If
                    Next X

                    If yTemp Is Nothing = False Then
                        ReDim Preserve yResp(yTemp.Length + 1)
                        yTemp.CopyTo(yResp, 2)
                        System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityDetails).CopyTo(yResp, 0)
                        Erase yTemp
                    End If

                    If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eGetEntityDetails, MsgMonitor.eMM_AppType.ClientConnection, yResp.Length, mlClientPlayer(lClientIndex))
                    moClients(lClientIndex).SendData(yResp)
                    Erase yResp
                End If
            Case ObjectType.eUnit, ObjectType.eFacility
                Dim oEntity As Epica_Entity = Nothing
                Dim lIdx As Int32 = LookupEntity(lID, iTypeID)
                If lIdx > -1 Then oEntity = goEntity(lIdx)
                If oEntity Is Nothing = False AndAlso oEntity.ObjectID = lID AndAlso oEntity.ObjTypeID = iTypeID Then
                    If oEntity.lOwnerID = mlClientPlayer(lClientIndex) OrElse (mlAliasedAs(lClientIndex) = oEntity.lOwnerID AndAlso (mlAliasedRights(lClientIndex) And AliasingRights.eViewUnitsAndFacilities) <> 0) Then
                        ReDim yResp(43) '47
                        System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityDetails).CopyTo(yResp, 0)
                        With oEntity
                            .GetGUIDAsString.CopyTo(yResp, 2)
                            .UnitName.CopyTo(yResp, 8)

                            lTemp = 100
                            yResp(28) = CByte(lTemp)

                            yResp(29) = .Exp_Level
                            System.BitConverter.GetBytes(.iTargetingTactics).CopyTo(yResp, 30)
                            System.BitConverter.GetBytes(.iCombatTactics).CopyTo(yResp, 32)

                            System.BitConverter.GetBytes(.lTetherPointX).CopyTo(yResp, 36)
                            System.BitConverter.GetBytes(.lTetherPointZ).CopyTo(yResp, 40)

                            'Dim oDef As Epica_Entity_Def = Nothing
                            'If .lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(.lEntityDefServerIndex) = .lEntityDefServerID Then
                            '    oDef = goEntityDefs(.lEntityDefServerIndex)
                            '    If oDef Is Nothing = False Then
                            '        Dim lWPU As Int32 = oDef.lWarpointValue \ 100
                            '        System.BitConverter.GetBytes(lWPU).CopyTo(yResp, 44)
                            '    End If
                            'End If
                        End With

                        If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eGetEntityDetails, MsgMonitor.eMM_AppType.ClientConnection, yResp.Length, mlClientPlayer(lClientIndex))
                        moClients(lClientIndex).SendData(yResp)
                        Erase yResp
                    End If
                End If
        End Select



    End Sub

    Public Sub SendAIMoveRequestToPathfinding(ByVal oEntity As Epica_Entity, ByVal lDestX As Int32, ByVal lDestZ As Int32, ByVal iDestAngle As Int16)
        Dim yData(41) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequest).CopyTo(yData, 0)
        System.BitConverter.GetBytes(lDestX).CopyTo(yData, 2)
        System.BitConverter.GetBytes(lDestZ).CopyTo(yData, 6)
        System.BitConverter.GetBytes(iDestAngle).CopyTo(yData, 10)

        oEntity.ParentEnvir.GetGUIDAsString.CopyTo(yData, 12)
        oEntity.GetGUIDAsString.CopyTo(yData, 18)

        System.BitConverter.GetBytes(oEntity.LocX).CopyTo(yData, 24)
        System.BitConverter.GetBytes(oEntity.LocZ).CopyTo(yData, 28)
        System.BitConverter.GetBytes(oEntity.LocAngle).CopyTo(yData, 32)

        oEntity.ParentEnvir.GetGUIDAsString.CopyTo(yData, 34)

        If oEntity.lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(oEntity.lEntityDefServerIndex) <> -1 Then
            System.BitConverter.GetBytes(goEntityDefs(oEntity.lEntityDefServerIndex).ModelID).CopyTo(yData, 40)
        Else
            System.BitConverter.GetBytes(CShort(-1)).CopyTo(yData, 40)
        End If

        oEntity.bAIMoveRequestPending = True
        oEntity.bPlayerMoveRequestPending = False

        moPathfinding.SendData(yData)
        Erase yData
    End Sub

    Public Sub SendAIDockRequestToPathfinding(ByVal oEntity As Epica_Entity, ByVal lDockeeID As Int32, ByVal iDockeeTypeID As Int16)
        Dim yData(47) As Byte
        'Dim X As Int32
        'Dim lIdx As Int32 = -1

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X = 0 To lCurUB
        '    If glEntityIdx(X) = lDockeeID AndAlso goEntity(X).ObjTypeID = iDockeeTypeID Then
        '        lIdx = X
        '        Exit For
        '    End If
        'Next X
        Dim oDockee As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lDockeeID, iDockeeTypeID)
        If lIdx > -1 Then oDockee = goEntity(lIdx)
        If oDockee Is Nothing OrElse oDockee.ObjectID <> lDockeeID OrElse oDockee.ObjTypeID <> iDockeeTypeID Then Return

        'If (goEntity(lIdx).CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then Exit Sub

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestDock).CopyTo(yData, 0)
        oDockee.GetGUIDAsString.CopyTo(yData, 2)
        System.BitConverter.GetBytes(oDockee.LocX).CopyTo(yData, 8)
        System.BitConverter.GetBytes(oDockee.LocZ).CopyTo(yData, 12)
        System.BitConverter.GetBytes(oDockee.LocAngle).CopyTo(yData, 16)
        oDockee.ParentEnvir.GetGUIDAsString.CopyTo(yData, 18)

        With oEntity
            .GetGUIDAsString.CopyTo(yData, 24)
            System.BitConverter.GetBytes(.LocX).CopyTo(yData, 30)
            System.BitConverter.GetBytes(.LocZ).CopyTo(yData, 34)
            System.BitConverter.GetBytes(.LocAngle).CopyTo(yData, 38)
            .ParentEnvir.GetGUIDAsString.CopyTo(yData, 40)
            If .lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(.lEntityDefServerIndex) <> -1 Then
                System.BitConverter.GetBytes(goEntityDefs(.lEntityDefServerIndex).ModelID).CopyTo(yData, 46)
            Else
                System.BitConverter.GetBytes(CShort(-1)).CopyTo(yData, 46)
            End If
        End With

        oEntity.bAIMoveRequestPending = True
        oEntity.bPlayerMoveRequestPending = False

        moPathfinding.SendData(yData)
    End Sub

    Private Function CreateBurstEnvironmentResponse(ByVal yRequest() As Byte, ByRef oPlayer As Player) As Byte()
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yRequest, 2)  '2 because the first 2 bytes is msgid
        Dim lObjTypeID As Int16 = System.BitConverter.ToInt16(yRequest, 6)
        Dim lIdx As Int32 = -1
        Dim lEnvirIdx As Int32 = -1

        RemovePlayerFromAllEnvirs(oPlayer.ObjectID, -1, -1)

        Dim oEnvir As Envir = Nothing
        If lObjTypeID = ObjectType.ePlanet AndAlso lObjID >= 500000000 Then
            'instance
            Dim lCurUB As Int32 = -1
            If glInstanceIdx Is Nothing = False Then lCurUB = Math.Min(glInstanceUB, glInstanceIdx.GetUpperBound(0))
            For X As Int32 = 0 To lCurUB
                If glInstanceIdx(X) = lObjID Then
                    oEnvir = goInstances(X)
                    lEnvirIdx = X
                    Exit For
                End If
            Next X
        Else
            For X As Int32 = 0 To glEnvirUB
                'envirs are never removed... so we don't have to worry about null objects
                If goEnvirs(X).ObjectID = lObjID AndAlso goEnvirs(X).ObjTypeID = lObjTypeID Then
                    oEnvir = goEnvirs(X)
                    lEnvirIdx = X
                    Exit For
                End If
            Next X
        End If

        If oEnvir Is Nothing = False Then
            'Add the player to the environment
            oEnvir.lPlayersInEnvirCnt += 1

            'for setting which environment the player is in
            oPlayer.lEnvirIdx = lEnvirIdx

            'oEnvir.AddPlayerToEnvir(oPlayer)
            oEnvir.DoEnvirPlayerChange(Envir.eyChangePlayerEnvirCode.eAddToEnvir, oPlayer)

            Return oEnvir.GenerateBurstMessage(oPlayer.ObjectID)
        End If


        Return Nothing
    End Function

#Region " Old Single Entity Requests (Obsolete) "
    'Private Sub HandleMoveRequest(ByVal yData() As Byte, ByVal lSenderIndex As Int32)
    '    'this is a client requesting to move a unit from one location to another.
    '    ' yData will have our ObjID, DestX, DestZ
    '    'However, we will append the unit's current locX and locY and then forward it to the Pathfinding Server
    '    Dim X As Int32
    '    Dim lObjID As Int32
    '    Dim iObjTypeID As Int16
    '    Dim bFound As Boolean = False

    '    'first two bytes is the message code
    '    lObjID = System.BitConverter.ToInt32(yData, 2)
    '    iObjTypeID = System.BitConverter.ToInt16(yData, 6)

    '    If lObjID = -1 Then Exit Sub

    '    For X = 0 To glEntityUB
    '        If glEntityIdx(X) = lObjID AndAlso goEntity(X).ObjTypeID = iObjTypeID Then
    '            If goEntity(X).Owner Is Nothing = False Then
    '                If (lSenderIndex = -1 OrElse goEntity(X).Owner.ObjectID = mlClientPlayer(lSenderIndex)) AndAlso _
    '                   ((goEntity(X).CurrentStatus And elUnitStatus.eEngineOperational) <> 0) Then
    '                    bFound = True
    '                    Dim yNewMsg() As Byte
    '                    Dim lUB As Int32
    '                    lUB = yData.GetUpperBound(0)
    '                    ReDim yNewMsg(lUB + 16)
    '                    yData.CopyTo(yNewMsg, 0)
    '                    System.BitConverter.GetBytes(goEntity(X).LocX).CopyTo(yNewMsg, lUB) '+1
    '                    System.BitConverter.GetBytes(goEntity(X).LocZ).CopyTo(yNewMsg, lUB + 4) '5
    '                    System.BitConverter.GetBytes(goEntity(X).LocAngle).CopyTo(yNewMsg, lUB + 8) '9

    '                    System.BitConverter.GetBytes(goEntity(X).ParentEnvir.ObjectID).CopyTo(yNewMsg, lUB + 10) '11
    '                    System.BitConverter.GetBytes(goEntity(X).ParentEnvir.ObjTypeID).CopyTo(yNewMsg, lUB + 14) '15

    '                    goEntity(X).bPlayerMoveRequestPending = True
    '                    goEntity(X).bAIMoveRequestPending = False

    '                    'now send our message to the pathfinding server
    '                    SendToPathfinding(yNewMsg)
    '                    Erase yNewMsg
    '                End If

    '            End If

    '            Exit For
    '        End If
    '    Next X

    '    If bFound = False Then
    '        'ok, need to respond with a no good... we do this by replacing the first two bytes with a move no good
    '        If lSenderIndex <> -1 Then
    '            System.BitConverter.GetBytes(EpicaMessageCode.eMoveObjectRequestDeny).CopyTo(yData, 0)
    '            moClients(lSenderIndex).SendData(yData)
    '        End If
    '    End If

    'End Sub

    'Private Sub HandleAddWaypointRequest(ByVal yData() As Byte, ByVal lSenderIndex As Int32)
    '    'this is a client requesting to move add a waypoint to a unit from one location to another.
    '    ' yData will have our ObjID, DestX, DestZ
    '    'However, we will append the unit's current locX and locY and then forward it to the Pathfinding Server
    '    Dim X As Int32
    '    Dim lObjID As Int32
    '    Dim iObjTypeID As Int16
    '    Dim bFound As Boolean = False

    '    'first two bytes is the message code
    '    lObjID = System.BitConverter.ToInt32(yData, 2)
    '    iObjTypeID = System.BitConverter.ToInt16(yData, 6)

    '    If lObjID = -1 Then Exit Sub

    '    For X = 0 To glEntityUB
    '        If glEntityIdx(X) = lObjID AndAlso goEntity(X).ObjTypeID = iObjTypeID Then
    '            If goEntity(X).Owner Is Nothing = False Then
    '                'If goEntity(X).Owner.oSocket.SocketIndex = lSenderIndex Then
    '                If goEntity(X).Owner.ObjectID = mlClientPlayer(lSenderIndex) Then
    '                    bFound = True
    '                    Dim yNewMsg() As Byte
    '                    Dim lUB As Int32
    '                    lUB = yData.GetUpperBound(0)
    '                    ReDim yNewMsg(lUB + 16)
    '                    yData.CopyTo(yNewMsg, 0)
    '                    System.BitConverter.GetBytes(goEntity(X).LocX).CopyTo(yNewMsg, lUB + 1)
    '                    System.BitConverter.GetBytes(goEntity(X).LocZ).CopyTo(yNewMsg, lUB + 5)
    '                    System.BitConverter.GetBytes(goEntity(X).LocAngle).CopyTo(yNewMsg, lUB + 9)

    '                    System.BitConverter.GetBytes(goEntity(X).ParentEnvir.ObjectID).CopyTo(yNewMsg, lUB + 11)
    '                    System.BitConverter.GetBytes(goEntity(X).ParentEnvir.ObjTypeID).CopyTo(yNewMsg, lUB + 15)

    '                    goEntity(X).bPlayerMoveRequestPending = True
    '                    goEntity(X).bAIMoveRequestPending = False

    '                    'now send our message to the pathfinding server
    '                    SendToPathfinding(yNewMsg)
    '                    Erase yNewMsg
    '                End If

    '            End If

    '            Exit For
    '        End If
    '    Next X

    '    If bFound = False Then
    '        'ok, need to respond with a no good... we do this by replacing the first two bytes with a move no good
    '        System.BitConverter.GetBytes(EpicaMessageCode.eMoveObjectRequestDeny).CopyTo(yData, 0)
    '        moClients(lSenderIndex).SendData(yData)
    '    End If

    'End Sub
#End Region

    Private Sub HandleMoveCommand(ByVal yData() As Byte)
        'the pathfinding server is instructing us to move...
        Dim lObjID As Int32
        Dim iObjTypeID As Int16
        'Dim bFound As Boolean ' = False
        'Dim X As Int32

        Dim Dx As Int32
        Dim Dz As Int32
        Dim Da As Int16

        Dim iMsg As Int16 = System.BitConverter.ToInt16(yData, 0)

        Dim bClearMovedByPlayer As Boolean = False
        If iMsg = GlobalMessageCode.eFinalMoveCommand Then
            iMsg = GlobalMessageCode.eMoveObjectCommand
            bClearMovedByPlayer = True
        End If

        'the first two bytes is the message code
        lObjID = System.BitConverter.ToInt32(yData, 2)
        iObjTypeID = System.BitConverter.ToInt16(yData, 6)

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return
        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X = 0 To lCurUB
        '    If glEntityIdx(X) = lObjID AndAlso goEntity(X).ObjTypeID = iObjTypeID Then
        '        bFound = True
        '        oEntity = goEntity(X)
        '        Exit For
        '    End If
        'Next X

        If ((oEntity.CurrentStatus And elUnitStatus.eEngineOperational) <> 0) Then
            'if the entity has the MiningOrBuilding flag...
            'If (oEntity.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then
            '    'send a move lock reset msg
            '    Dim yMsg(7) As Byte
            '    System.BitConverter.GetBytes(GlobalMessageCode.eMoveLockViolate).CopyTo(yMsg, 0)
            '    oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
            '    SendToPrimary(yMsg)
            '    oEntity.CurrentStatus = oEntity.CurrentStatus Xor elUnitStatus.eMoveLock ' oEntity.CurrentStatus -= elUnitStatus.eMoveLock
            'End If

            'ok... found our unit... assign our values
            '  the DestX and DestZ should be the next values...
            Dx = System.BitConverter.ToInt32(yData, 8)
            Dz = System.BitConverter.ToInt32(yData, 12)

            Da = CShort(LineAngleDegrees(oEntity.LocX, oEntity.LocZ, Dx, Dz) * 10)
            Da = System.BitConverter.ToInt16(yData, 16)
            If Da <> -1 Then oEntity.TrueDestAngle = Da

            If oEntity.bPlayerMoveRequestPending = True Then oEntity.CurrentStatus = oEntity.CurrentStatus Or elUnitStatus.eMovedByPlayer
            oEntity.bPlayerMoveRequestPending = False
            oEntity.bAIMoveRequestPending = False

            'By resetting these, the unit uses them again
            With oEntity
                .yFormationManeuver = 0
                .yFormationMaxSpeed = 0
                .yFormationTurnAmount = 0
                .iFormationTurnAmount100 = 0
                .fFormationAcceleration = 0
            End With

            If iMsg = GlobalMessageCode.eEntityChangeEnvironment Then
                oEntity.yChangeEnvironments = yData(18)
            Else
                oEntity.yChangeEnvironments = 0
                If yData.GetUpperBound(0) >= 18 Then
                    oEntity.yFormationMaxSpeed = Math.Min(yData(18), oEntity.MaxSpeed)
                End If
            End If

            oEntity.SetDest(Dx, Dz, Da)

            If bClearMovedByPlayer = True Then
                oEntity.yResetMoveByPlayer = 255
                'If (oEntity.CurrentStatus And elUnitStatus.eMovedByPlayer) <> 0 Then oEntity.CurrentStatus -= elUnitStatus.eMovedByPlayer
            End If

            'finally, relay the value to the clients...
            If oEntity.ParentEnvir.lPlayersInEnvirCnt > 0 Then BroadcastToEnvironmentClients(yData, oEntity.ParentEnvir)
        End If

    End Sub

    Private Sub HandleMoveRequestDeny(ByVal yData() As Byte)
        'ok the pathfinding server is telling us that the path is invalid... respond to the client
        Dim lObjID As Int32
        Dim iObjTypeID As Int16
        'Dim X As Int32
        Dim bNeedToForward As Boolean = False

        'the first two bytes is the message code
        lObjID = System.BitConverter.ToInt32(yData, 2)
        iObjTypeID = System.BitConverter.ToInt16(yData, 6)

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        If oEntity.Owner Is Nothing = False Then
            bNeedToForward = oEntity.bPlayerMoveRequestPending
            oEntity.bPlayerMoveRequestPending = False
            oEntity.bAIMoveRequestPending = False
            If bNeedToForward = True Then
                If oEntity.Owner.oSocket Is Nothing Then
                    bNeedToForward = False
                    For Y As Int32 = 0 To oEntity.Owner.lAliasUB
                        If oEntity.Owner.lAliasIdx(Y) <> -1 AndAlso oEntity.Owner.oAliases(Y) Is Nothing = False AndAlso oEntity.Owner.oAliases(Y).oSocket Is Nothing = False Then
                            bNeedToForward = True
                            Exit For
                        End If
                    Next Y
                End If
            End If
            If bNeedToForward = True Then 'AndAlso oEntity.Owner.oSocket Is Nothing = False Then
                'resend it to the player(s)
                If oEntity.Owner.oSocket Is Nothing = False Then
                    If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eMoveObjectRequestDeny, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, oEntity.lOwnerID)
                    oEntity.Owner.oSocket.SendData(yData)
                End If
                For Y As Int32 = 0 To oEntity.Owner.lAliasUB
                    If oEntity.Owner.lAliasIdx(Y) <> -1 AndAlso oEntity.Owner.oAliases(Y) Is Nothing = False AndAlso oEntity.Owner.oAliases(Y).oSocket Is Nothing = False Then
                        If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eMoveObjectRequestDeny, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, oEntity.Owner.oAliases(Y).ObjectID)
                        oEntity.Owner.oAliases(Y).oSocket.SendData(yData)
                    End If
                Next Y
            End If
        End If

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X = 0 To lCurUB
        '    If glEntityIdx(X) = lObjID AndAlso goEntity(X).ObjTypeID = iObjTypeID Then
        '        If goEntity(X).Owner Is Nothing = False Then
        '            bNeedToForward = goEntity(X).bPlayerMoveRequestPending
        '            goEntity(X).bPlayerMoveRequestPending = False
        '            goEntity(X).bAIMoveRequestPending = False
        '            If bNeedToForward = True Then
        '                If goEntity(X).Owner.oSocket Is Nothing Then
        '                    bNeedToForward = False
        '                    For Y As Int32 = 0 To goEntity(X).Owner.lAliasUB
        '                        If goEntity(X).Owner.lAliasIdx(Y) <> -1 AndAlso goEntity(X).Owner.oAliases(Y) Is Nothing = False AndAlso goEntity(X).Owner.oAliases(Y).oSocket Is Nothing = False Then
        '                            bNeedToForward = True
        '                            Exit For
        '                        End If
        '                    Next Y
        '                End If
        '            End If
        '            If bNeedToForward = True Then 'AndAlso goEntity(X).Owner.oSocket Is Nothing = False Then
        '                'resend it to the player(s)
        '                If goEntity(X).Owner.oSocket Is Nothing = False Then
        '                    If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eMoveObjectRequestDeny, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, goEntity(X).lOwnerID)
        '                    goEntity(X).Owner.oSocket.SendData(yData)
        '                End If
        '                For Y As Int32 = 0 To goEntity(X).Owner.lAliasUB
        '                    If goEntity(X).Owner.lAliasIdx(Y) <> -1 AndAlso goEntity(X).Owner.oAliases(Y) Is Nothing = False AndAlso goEntity(X).Owner.oAliases(Y).oSocket Is Nothing = False Then
        '                        If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eMoveObjectRequestDeny, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, goEntity(X).Owner.oAliases(Y).ObjectID)
        '                        goEntity(X).Owner.oAliases(Y).oSocket.SendData(yData)
        '                    End If
        '                Next Y
        '            End If
        '        End If
        '        Exit For
        '    End If
        'Next X
    End Sub

    Public Function CreateStopObjectCommand(ByVal oEntity As Epica_Entity) As Byte()
        'Stop object consists of the unit's ID, the locx, locz, locangle
        Dim yBytes(17) As Byte      '0 to 17 bytes = 18 bytes total

        'Stop Object Command Message Is:
        '  SO Code - 2 bytes
        '  ObjectID - 4 bytes
        '  ObjTypeID - 2 bytes
        '  LocX - 4 Bytes
        '  LocZ - 4 bytes
        '  LocAngle - 2 bytes
        System.BitConverter.GetBytes(GlobalMessageCode.eStopObjectCommand).CopyTo(yBytes, 0)
        System.BitConverter.GetBytes(oEntity.ObjectID).CopyTo(yBytes, 2)
        System.BitConverter.GetBytes(oEntity.ObjTypeID).CopyTo(yBytes, 6)
        System.BitConverter.GetBytes(oEntity.LocX).CopyTo(yBytes, 8)
        System.BitConverter.GetBytes(oEntity.LocZ).CopyTo(yBytes, 12)
        System.BitConverter.GetBytes(oEntity.LocAngle).CopyTo(yBytes, 16)

        Return yBytes
    End Function

    'MSC 05/13/08 - added for new fire weapon optimization
    Public Function CreateSetTargetMessage(ByRef oAttacker As Epica_Entity) As Byte()
        Dim lEntityID As Int32
        If oAttacker.ObjTypeID = ObjectType.eFacility Then lEntityID = -oAttacker.ObjectID Else lEntityID = oAttacker.ObjectID

        Dim yResult(21) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityTarget).CopyTo(yResult, 0)
        System.BitConverter.GetBytes(lEntityID).CopyTo(yResult, 2)

        Dim lPos As Int32 = 6
        For X As Int32 = 0 To 3
            Dim lTargetID As Int32 = 0
            Dim lTargetIdx As Int32 = oAttacker.lTargetsServerIdx(X)
            If lTargetIdx > -1 AndAlso glEntityIdx(lTargetIdx) > 0 Then
                Dim oEntity As Epica_Entity = goEntity(lTargetIdx)
                If oEntity Is Nothing = False Then
                    If oEntity.ObjTypeID = ObjectType.eFacility Then
                        lTargetID = -oEntity.ObjectID
                    Else : lTargetID = oEntity.ObjectID
                    End If
                End If
            End If
            System.BitConverter.GetBytes(lTargetID).CopyTo(yResult, lPos) : lPos += 4
        Next X
        Return yResult
    End Function

    'MSC 05/13/08 - added for new fire weapon optimization
    'yHitType = 0 for a miss, 1 for armor hit, 2 for shield hit
    Public Function CreateFireWeaponMessage(ByVal lAttackerID As Int32, ByVal iAttackerTypeID As Int16, ByVal ySideShot As Byte, ByVal yWeaponType As Byte, ByVal yHitType As Byte, ByVal yAOE As Byte) As Byte()
        Dim yBytes(7) As Byte
        If yHitType = 0 Then
            System.BitConverter.GetBytes(GlobalMessageCode.eFireWeaponMiss).CopyTo(yBytes, 0)
        Else
            System.BitConverter.GetBytes(GlobalMessageCode.eFireWeapon).CopyTo(yBytes, 0)
        End If
        If iAttackerTypeID = ObjectType.eFacility Then
            System.BitConverter.GetBytes(-lAttackerID).CopyTo(yBytes, 2)
        Else
            System.BitConverter.GetBytes(lAttackerID).CopyTo(yBytes, 2)
        End If
        If yHitType = 2 Then yWeaponType = CByte(yWeaponType Or WeaponType.eShieldHitBitMask)
        yBytes(6) = yWeaponType

        'Prepare our special data here...
        Dim lAOE As Int32
        If yAOE <> 0 Then
            lAOE = yAOE
            lAOE \= 4
            lAOE = lAOE << 2
            lAOE = lAOE Or ySideShot
            If lAOE > 255 Then lAOE = 255
            If lAOE < 0 Then lAOE = 0
            ySideShot = CByte(lAOE)
        End If

        yBytes(7) = ySideShot
        Return yBytes
    End Function

    Public Function CreatePointDefenseMessage(ByVal iMsg As Int16, ByRef oAttacker As Epica_Entity, ByVal lTargetID As Int32, ByVal iTargetTypeID As Int16, ByVal yWeaponType As Byte, ByVal yHitType As Byte) As Byte()
        Dim yOut(10) As Byte
        System.BitConverter.GetBytes(iMsg).CopyTo(yOut, 0)

        If oAttacker.ObjTypeID = ObjectType.eFacility Then
            System.BitConverter.GetBytes(-oAttacker.ObjectID).CopyTo(yOut, 2)
        Else
            System.BitConverter.GetBytes(oAttacker.ObjectID).CopyTo(yOut, 2)
        End If
        If iTargetTypeID = ObjectType.eFacility Then
            System.BitConverter.GetBytes(-lTargetID).CopyTo(yOut, 6)
        Else
            System.BitConverter.GetBytes(lTargetID).CopyTo(yOut, 6)
        End If
        If yHitType = 2 Then yWeaponType = CByte(yWeaponType Or WeaponType.eShieldHitBitMask)
        yOut(10) = yWeaponType
        Return yOut
    End Function

    ''yHitType = 0 for a miss, 1 for armor hit, 2 for shield hit
    'Public Function CreateFireWeaponMessage(ByVal lAttackerID As Int32, ByVal iAttackerTypeID As Int16, ByVal lTargetID As Int32, _
    'ByVal iTargetTypeID As Int16, ByVal yWeaponType As Byte, ByVal yHitType As Byte) As Byte()

    '	Dim yBytes(14) As Byte	'0 to 14 bytes = 15 bytes total

    '	'fire weapon message is:
    '	'  FW code - 2 bytes
    '	'  AttackerID - 4 bytes
    '	'  AttackerTypeID - 2 bytes 
    '	'  TargetID - 4 bytes
    '	'  TargetTypeID - 2 bytes
    '	'  WeaponType - 1 byte
    '	If yHitType = 0 Then
    '		System.BitConverter.GetBytes(GlobalMessageCode.eFireWeaponMiss).CopyTo(yBytes, 0)
    '	Else
    '		System.BitConverter.GetBytes(GlobalMessageCode.eFireWeapon).CopyTo(yBytes, 0)
    '	End If
    '	System.BitConverter.GetBytes(lAttackerID).CopyTo(yBytes, 2)
    '	System.BitConverter.GetBytes(iAttackerTypeID).CopyTo(yBytes, 6)
    '	System.BitConverter.GetBytes(lTargetID).CopyTo(yBytes, 8)
    '	System.BitConverter.GetBytes(iTargetTypeID).CopyTo(yBytes, 12)

    '	If yHitType = 2 Then yWeaponType = CByte(yWeaponType Or WeaponType.eShieldHitBitMask)
    '	yBytes(14) = yWeaponType

    '	Return yBytes
    'End Function

    Public Function CreateBombardFireMessage(ByVal oAttacker As Epica_GUID, ByVal oPlanet As Epica_GUID, ByVal yWeaponType As Byte, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal yAOE As Byte) As Byte()
        Dim yBytes(23) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eBombardFireMsg).CopyTo(yBytes, 0)
        oAttacker.GetGUIDAsString.CopyTo(yBytes, 2)
        oPlanet.GetGUIDAsString.CopyTo(yBytes, 8)
        yBytes(14) = yWeaponType
        System.BitConverter.GetBytes(lLocX).CopyTo(yBytes, 15)
        System.BitConverter.GetBytes(lLocZ).CopyTo(yBytes, 19)
        yBytes(23) = yAOE

        Return yBytes
    End Function

    Private Shared Function AddEntity(ByVal lObjID As Int32, ByVal iObjTypeID As Int16) As Int32
        Dim lIdx As Int32 = -1
        Dim lFirstIndex As Int32 = -1

        SyncLock goEntity
            Dim lCurUB As Int32 = -1
            If glEntityIdx Is Nothing = False Then lCurUB = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
            For X As Int32 = 0 To lCurUB
                If glEntityIdx(X) = lObjID AndAlso goEntity(X).ObjTypeID = iObjTypeID Then
                    lIdx = X
                    If lFirstIndex <> -1 AndAlso glEntityIdx(lFirstIndex) = -2 Then glEntityIdx(lFirstIndex) = -1
                    Exit For
                ElseIf lFirstIndex = -1 AndAlso glEntityIdx(X) = -1 Then
                    lFirstIndex = X
                    glEntityIdx(X) = -2
                End If
            Next X

            If lIdx = -1 Then
                'No... check lFirstIndex
                If lFirstIndex <> -1 Then
                    'Ok, so set our object to that
                    lIdx = lFirstIndex
                Else
                    'Ok, we don't have a clear spot... so redim
                    glEntityUB += 1
                    lIdx = glEntityUB
                    If glEntityUB > glEntityIdx.GetUpperBound(0) Then
                        ReDim Preserve glEntityIdx(glEntityUB + 10000)
                        ReDim Preserve goEntity(glEntityUB + 10000)
                    End If
                End If
                goEntity(lIdx) = New Epica_Entity()
                glEntityIdx(lIdx) = -2
            End If
        End SyncLock

        Return lIdx
    End Function

    Private Sub ReceiveAddObject(ByVal yData() As Byte)
        'ok, the primary server sent us an Add Object Message... yData has the pertinent information...
        ' an Add Object has the following format: ObjID, ObjTypeID, ParentID, ParentTypeID... data after that
        ' is specific to the object being added
        Dim lObjID As Int32
        Dim iObjTypeID As Int16
        Dim lParentID As Int32
        Dim iParentTypeID As Int16
        Dim X As Int32
        Dim Y As Int32
        Dim lIdx As Int32 = -1
        Dim lFirstIndex As Int32 = -1
        Dim lTemp As Int32
        Dim lTemp2 As Int32
        Dim lPos As Int32

        Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yData, 0)

        lObjID = System.BitConverter.ToInt32(yData, 2)
        iObjTypeID = System.BitConverter.ToInt16(yData, 6)

        Select Case iObjTypeID
            Case ObjectType.eUnit, ObjectType.eFacility
                'first check if we already have this object...
                lParentID = System.BitConverter.ToInt32(yData, 8)
                iParentTypeID = System.BitConverter.ToInt16(yData, 12)

                'If iObjTypeID = ObjectType.eFacility Then
                '    gfrmDisplayForm.AddEventLine("Received Add Facility to " & lParentID & " type " & iParentTypeID)
                'Else
                '    gfrmDisplayForm.AddEventLine("Received Add Unit to " & lParentID & " type " & iParentTypeID)
                'End If

                'For X = 0 To glEntityUB
                '	If glEntityIdx(X) = lObjID AndAlso goEntity(X).ObjTypeID = iObjTypeID Then
                '		lIdx = X
                '		Exit For
                '	ElseIf lFirstIndex = -1 AndAlso glEntityIdx(X) = -1 Then
                '		lFirstIndex = X
                '	End If
                'Next X

                'If lIdx = -1 Then
                '	'No... check lFirstIndex
                '	If lFirstIndex <> -1 Then
                '		'Ok, so set our object to that
                '		lIdx = lFirstIndex
                '	Else
                '		'Ok, we don't have a clear spot... so redim
                '		glEntityUB += 1
                '		lIdx = glEntityUB
                '		ReDim Preserve glEntityIdx(glEntityUB)
                '		ReDim Preserve goEntity(glEntityUB)
                '	End If
                '	goEntity(lIdx) = New Epica_Entity()
                'End If
                lIdx = AddEntity(lObjID, iObjTypeID)
                If lIdx < 0 OrElse lIdx > glEntityUB Then
                    gfrmDisplayForm.AddEventLine("Invalid Index returned from AddEntity! Object not added: " & lObjID & ", " & iObjTypeID)
                    Return
                End If

                'set our index
                glEntityIdx(lIdx) = lObjID

                'Now, parse our values... first check the parent object
                With goEntity(lIdx)
                    .ObjectID = lObjID
                    .ObjTypeID = iObjTypeID

                    .ServerIndex = lIdx

                    If .ParentEnvir Is Nothing = False Then
                        If .lGridIndex <> -1 AndAlso .lGridEntityIdx <> -1 Then
                            If .ParentEnvir.oGrid(.lGridIndex).lEntityUB >= .lGridEntityIdx Then
                                If .ParentEnvir.oGrid(.lGridIndex).lEntities(.lGridEntityIdx) = .ServerIndex Then
                                    .ParentEnvir.oGrid(.lGridIndex).RemoveEntity(.lGridEntityIdx)
                                End If
                            End If
                        End If
                        If .ParentEnvir.ObjectID <> lParentID OrElse .ParentEnvir.ObjTypeID <> iParentTypeID Then
                            'ok, we are not good for the environment...
                            If iParentTypeID = ObjectType.ePlanet AndAlso lParentID >= 500000000 Then
                                'instance
                                Dim lCurUB As Int32 = -1
                                If glInstanceIdx Is Nothing = False Then lCurUB = Math.Min(glInstanceUB, glInstanceIdx.GetUpperBound(0))
                                For X = 0 To lCurUB
                                    If glInstanceIdx(X) = lParentID Then
                                        'ok, found the environment...
                                        .ParentEnvir = goInstances(X)
                                        Exit For
                                    End If
                                Next X
                            Else
                                For X = 0 To glEnvirUB
                                    If goEnvirs(X).ObjectID = lParentID AndAlso goEnvirs(X).ObjTypeID = iParentTypeID Then
                                        'ok, found the environment...
                                        .ParentEnvir = goEnvirs(X)
                                        Exit For
                                    End If
                                Next X
                            End If
                        End If
                    Else
                        If iParentTypeID = ObjectType.ePlanet AndAlso lParentID >= 500000000 Then
                            Dim lCurUB As Int32 = -1
                            If glInstanceIdx Is Nothing = False Then lCurUB = Math.Min(glInstanceUB, glInstanceIdx.GetUpperBound(0))
                            For X = 0 To lCurUB
                                If glInstanceIdx(X) = lParentID Then
                                    'ok, found the environment...
                                    .ParentEnvir = goInstances(X)
                                    Exit For
                                End If
                            Next X
                        Else
                            For X = 0 To glEnvirUB
                                If goEnvirs(X).ObjectID = lParentID AndAlso goEnvirs(X).ObjTypeID = iParentTypeID Then
                                    'ok, found the environment...
                                    .ParentEnvir = goEnvirs(X)
                                    Exit For
                                End If
                            Next X
                        End If
                    End If

                    'If .ParentEnvir Is Nothing = False Then
                    '    If .ParentEnvir.ObjectID <> lParentID OrElse .ParentEnvir.ObjTypeID <> iParentTypeID Then
                    '        gfrmDisplayForm.AddEventLine("ParentEnvir Mismatch! " & lObjID & ", " & iObjTypeID)
                    '        'Now, check all environments for this object (REMOVE ME WHEN DONE)
                    '        For X = 0 To glEnvirUB
                    '            For Y = 0 To goEnvirs(X).lGridUB
                    '                If goEnvirs(X).oGrid(Y) Is Nothing = False Then
                    '                    With goEnvirs(X).oGrid(Y)
                    '                        For Z As Int32 = 0 To .lEntityUB
                    '                            If .lEntities(Z) = lIdx Then
                    '                                Stop
                    '                            End If
                    '                        Next Z
                    '                    End With
                    '                End If
                    '            Next Y
                    '        Next X
                    '    End If
                    'Else
                    '    gfrmDisplayForm.AddEventLine("ParentEnvir is nothing! " & lObjID & ", " & iObjTypeID)
                    '    Stop
                    'End If

                    'Ok, now... get the UnitDefID
                    lTemp = System.BitConverter.ToInt32(yData, 14)
                    .lEntityDefServerID = lTemp
                    .lEntityDefServerIndex = -1
                    For X = 0 To glEntityDefUB
                        If glEntityDefIdx(X) = lTemp Then
                            If (iObjTypeID = ObjectType.eUnit AndAlso goEntityDefs(X).ObjTypeID = ObjectType.eUnitDef) OrElse _
                               (iObjTypeID = ObjectType.eFacility AndAlso goEntityDefs(X).ObjTypeID = ObjectType.eFacilityDef) Then
                                .lEntityDefServerIndex = X
                                .lModelRangeOffset = goEntityDefs(X).lModelRangeOffset
                                .Acceleration = goEntityDefs(X).BaseAcceleration
                                .Maneuver = goEntityDefs(X).BaseManeuver
                                .MaxSpeed = goEntityDefs(X).BaseMaxSpeed
                                .TurnAmount = goEntityDefs(X).BaseTurnAmount
                                .TurnAmountTimes100 = goEntityDefs(X).BaseTurnAmountTimes100

                                ReDim .lWpnAmmoCnt(goEntityDefs(X).WeaponDefUB)
                                For Y = 0 To .lWpnAmmoCnt.Length - 1
                                    .lWpnAmmoCnt(Y) = -1
                                Next Y

                                Exit For
                            End If
                        End If
                    Next X
                    If .lEntityDefServerIndex = -1 Then
                        X = 1
                    End If

                    Array.Copy(yData, 18, .UnitName, 0, 20)

                    lTemp = System.BitConverter.ToInt32(yData, 38)
                    If .ObjTypeID = ObjectType.eUnit AndAlso .Owner Is Nothing = False Then
                        'Ok, already has an owner, is the same player?
                        If .Owner.ObjectID <> lTemp Then
                            'no, so, we need to adjust the command points
                            .ParentEnvir.AdjustPlayerCommandPoints(.Owner.ObjectID, -(.CPUsage + .Owner.BadWarDecCPIncrease))
                        End If
                    End If

                    For X = 0 To glPlayerUB
                        If glPlayerIdx(X) = lTemp Then
                            .Owner = goPlayers(X)
                            '.sPlayerRelString = "PR_" & CStr(lTemp)
                            .lOwnerID = lTemp
                            Exit For
                        End If
                    Next X

                    If .Owner Is Nothing Then
                        .Owner = Nothing
                        'TODO: It is okay to have no owner, but in that case, it is a free-for-all unit...
                        ' now, I believe we should introduce dummy factions or "owners" that are pirates or something
                    End If

                    .InitializeWeaponCycles()   'we are suppose to call this when the entity def is assigned


                    Dim lTmpLocX As Int32 = System.BitConverter.ToInt32(yData, 42)
                    Dim lTmpLocZ As Int32 = System.BitConverter.ToInt32(yData, 46)
                    .LocX = lTmpLocX
                    .LocZ = lTmpLocZ
                    .MoveLocX = .LocX
                    .MoveLocZ = .LocZ
                    .LocX = lTmpLocX
                    .LocZ = lTmpLocZ

                    '.lTetherPointX = Int32.MinValue '.LocX
                    '.lTetherPointZ = Int32.MinValue '.LocZ

#If EXTENSIVELOGGING = 1 Then
                    gfrmDisplayForm.AddEventLine("Add Object Received(" & lIdx & "): " & lObjID & ", " & iObjTypeID & " to " & lParentID & ", " & iParentTypeID & " at " & .LocX & ", " & .LocZ)
#End If

                    .LocAngle = System.BitConverter.ToInt16(yData, 50)
                    .StructuralHP = System.BitConverter.ToInt32(yData, 52)
                    .Fuel_Cap = System.BitConverter.ToInt32(yData, 56)
                    .lLastShieldRechargeCycle = glCurrentCycle
                    .ShieldHP = System.BitConverter.ToInt32(yData, 60)
                    .Exp_Level = yData(64)
                    lTemp = System.BitConverter.ToInt32(yData, 65)  'unitgroupid
                    .CurrentStatus = System.BitConverter.ToInt32(yData, 69)

                    .ArmorHP(0) = System.BitConverter.ToInt32(yData, 73)
                    .ArmorHP(1) = System.BitConverter.ToInt32(yData, 77)
                    .ArmorHP(2) = System.BitConverter.ToInt32(yData, 81)
                    .ArmorHP(3) = System.BitConverter.ToInt32(yData, 85)

                    lPos = 89
                    If iObjTypeID = ObjectType.eFacility Then
                        'got some extra values that we currently don't use... but we'll need to
                        'the extra values are...
                        'CurrentWorkers (4 bytes)
                        'ParentColonyID (4 bytes)
                        'Width (2 bytes)
                        'Height (2 bytes)
                        lPos += 12
                    End If
                    .yProductionType = yData(lPos) : lPos += 1
                    .iCombatTactics = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .iTargetingTactics = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                    'Now, get our cnt of Ammo Amts
                    lTemp = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    For X = 0 To lTemp - 1
                        lTemp2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        For Y = 0 To goEntityDefs(.lEntityDefServerIndex).WeaponDefUB
                            If goEntityDefs(.lEntityDefServerIndex).WeaponDefs(Y).lEntityWpnDefID = lTemp2 Then
                                .lWpnAmmoCnt(Y) = System.BitConverter.ToInt32(yData, lPos)
                                Exit For
                            End If
                        Next Y
                        lPos += 4       'regardless of whether we find it or not, increment by 4
                    Next X
                    If iObjTypeID = ObjectType.eUnit Then
                        .lTetherPointX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lTetherPointZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    End If


                    .MoveLocX = .LocX
                    .MoveLocZ = .LocZ
                    .bForceAggressionTest = True
                    .lLastForceAggressionTest = glCurrentCycle - gl_FORCE_AGGRESSION_THRESHOLD
                    .bNewAddedEntity = True
                    .DestX = .LocX
                    .DestZ = .LocZ
                    .DestAngle = .LocAngle
                    .LastCycleMoved = glCurrentCycle

                    .SetExpLevelMods()

                    AddLookupEntity(.ObjectID, .ObjTypeID, lIdx)

                    'finally, add the object to its parent environment
                    If .ParentEnvir Is Nothing = False Then
                        .ParentEnvir.AddEntity(goEntity(lIdx))
                    End If
                    SyncLockMovementRegisters(MovementCommand.RemoveEntityMoving, .ServerIndex, .ObjectID)
                    .LocX = lTmpLocX
                    .LocZ = lTmpLocZ
                    .MoveLocX = lTmpLocX
                    .MoveLocZ = lTmpLocZ
                    .LocX = lTmpLocX
                    .LocZ = lTmpLocZ
                    .LastCycleMoved = glCurrentCycle
                    SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)
                End With

                'To force the entity to move to the next location from Pathfinding...
                If iMsgCode = GlobalMessageCode.eAddObjectCommand_CE Then
                    SendToPathfinding(CreateStopObjectCommand(goEntity(lIdx)))
                End If

                'Now, that we have our object... forward the command to connected players
                If goEntity(lIdx).ParentEnvir.lPlayersInEnvirCnt > 0 Then
                    BroadcastToEnvironmentClients_Ex(iMsgCode, goEntity(lIdx).GetObjectAsSmallString, goEntity(lIdx).ParentEnvir)
                End If

            Case ObjectType.ePlayer
                'gfrmDisplayForm.AddEventLine("Received Add Player")
                'ok... this is a GUID (4 and 2)
                For X = 0 To glPlayerUB
                    If glPlayerIdx(X) = lObjID Then
                        'we already have it...
                        lIdx = X
                        Exit For
                    ElseIf glPlayerIdx(X) = -1 AndAlso lFirstIndex = -1 Then
                        lFirstIndex = X
                    End If
                Next X

                If lIdx = -1 Then
                    'we don't have this one yet... so add it
                    If lFirstIndex = -1 Then
                        'need to redim
                        glPlayerUB += 1
                        ReDim Preserve goPlayers(glPlayerUB)
                        ReDim Preserve glPlayerIdx(glPlayerUB)
                        lIdx = glPlayerUB
                    Else
                        lIdx = lFirstIndex
                    End If

                    goPlayers(lIdx) = New Player()
                    glPlayerIdx(lIdx) = lObjID
                    With goPlayers(lIdx)
                        '2 bytes for message code... 6 for guid
                        .ObjectID = lObjID
                        .ObjTypeID = iObjTypeID
                        ' PlayerName (20)
                        Array.Copy(yData, 8, .PlayerName, 0, 20)
                        ' Empire Name (20)
                        Array.Copy(yData, 28, .EmpireName, 0, 20)
                        ' Race Name (20)
                        Array.Copy(yData, 48, .RaceName, 0, 20)
                        ' Player User Name (20)
                        Array.Copy(yData, 68, .PlayerUserName, 0, 20)
                        ' Player Password (20)
                        Array.Copy(yData, 88, .PlayerPassword, 0, 20)
                        ' SenateID (4)
                        .SenateID = System.BitConverter.ToInt32(yData, 108)
                        ' CommEncryptLevel (2)
                        .CommEncryptLevel = System.BitConverter.ToInt16(yData, 112)
                        ' EmpireTaxRate (1)
                        .EmpireTaxRate = yData(114)
                        ' lCredits (8)
                        .lStartEnvirID = System.BitConverter.ToInt32(yData, 123)
                        .lStartLocX = System.BitConverter.ToInt32(yData, 127)
                        .lStartLocZ = System.BitConverter.ToInt32(yData, 131)
                        .lPirateStartLocX = System.BitConverter.ToInt32(yData, 135)
                        .lPirateStartLocZ = System.BitConverter.ToInt32(yData, 139)
                        lPos = 143
                        .lGuildID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .lJoinedGuildOn = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .BadWarDecCPIncrease = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        lPos += 4        'for the BadWarDecMoralePenalty value
                        .yPlayerPhase = yData(lPos) : lPos += 1
                        lPos += 4       'tutorial step
                        .AccountStatus = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    End With
                End If
            Case ObjectType.eUnitDef, ObjectType.eFacilityDef       'not sure if this is accurate...
                'generally speaking... the server will cache unit def objects in order to make the processing
                '  of units faster than if the unit's data had to be transmitted each time.
                'NOTE - the primary server cannot allow the Unit Def object to be editted. A New Unit Def object
                '  will need to be created. (the old one can be deleted in the normal manner if there are no units
                '  still currently active that reference that unit def).
                'NOTE: in order for this to really work, the primary updates all domains of all unitdef objects
                'first, determine if this unit def already exists
                For X = 0 To glEntityDefUB
                    If glEntityDefIdx(X) = lObjID AndAlso goEntityDefs(X).ObjTypeID = iObjTypeID Then
                        'we already have it...
                        lIdx = X
                        Exit For
                    ElseIf lFirstIndex = -1 AndAlso glEntityDefIdx(X) = -1 Then
                        lFirstIndex = X
                    End If
                Next X

                If lIdx = -1 Then
                    'gfrmDisplayForm.AddEventLine("Received Entity Def")

                    'we dont have this one yet... so add it
                    If lFirstIndex = -1 Then
                        'need to redim
                        glEntityDefUB += 1
                        ReDim Preserve glEntityDefIdx(glEntityDefUB)
                        ReDim Preserve goEntityDefs(glEntityDefUB)
                        lIdx = glEntityDefUB
                    Else
                        lIdx = lFirstIndex
                    End If

                    goEntityDefs(lIdx) = New Epica_Entity_Def()
                    glEntityDefIdx(lIdx) = lObjID
                    With goEntityDefs(lIdx)
                        '2 bytes for message code... 6 bytes for GUID

                        lPos = 8
                        lPos += 4        'for ownerid
                        .Armor_MaxHP(0) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .Armor_MaxHP(1) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .Armor_MaxHP(2) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .Armor_MaxHP(3) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.Maneuver = yData(24)
                        .BaseManeuver = yData(lPos) : lPos += 1
                        '.MaxSpeed = yData(25)
                        .BaseMaxSpeed = yData(lPos) : lPos += 1
                        .FuelEfficiency = yData(lPos) : lPos += 1
                        .Structure_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .Hangar_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .HullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .Cargo_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .Shield_MaxHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .ShieldRecharge = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .ShieldRechargeFreq = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.MaxDoorSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        '.NumberOfDoors = yData(lPos) : lPos += 1
                        .MaxCrew = System.BitConverter.ToInt32(yData, lPos) : lPos += 4     'was FuelCap
                        .Weapon_Acc = yData(lPos) : lPos += 1
                        .ScanResolution = yData(lPos) : lPos += 1
                        .yDefOptRadarRange = yData(lPos) : lPos += 1
                        .yDefMaxRadarRange = yData(lPos) : lPos += 1
                        .DisruptionResist = yData(lPos) : lPos += 1
                        .PiercingResist = yData(lPos) : lPos += 1
                        .ImpactResist = yData(lPos) : lPos += 1
                        .BeamResist = yData(lPos) : lPos += 1
                        .ECMResist = yData(lPos) : lPos += 1
                        .FlameResist = yData(lPos) : lPos += 1
                        .ChemicalResist = yData(lPos) : lPos += 1
                        .DetectionResist = yData(lPos) : lPos += 1
                        .ModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        .ObjectID = lObjID
                        .ObjTypeID = iObjTypeID

                        'Def Name...
                        Array.Copy(yData, lPos, .DefName, 0, 20)
                        lPos += 20

                        If iObjTypeID = ObjectType.eFacilityDef Then
                            'got some extra members here that the Domain Server doesn't care about
                            ' at least, not yet anyway... the values not included are:
                            'WorkerFactor (4 bytes)
                            .WorkerFactor = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                            'MaxFacilitySize (1 byte)
                            lPos += 1
                            'ProdFactor (4 byte)
                            .ProdFactor = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                            'PowerFactor (4 byte)
                            lPos += 4
                        End If
                        'ProductionTypeID (1 byte)
                        .ProductionTypeID = yData(lPos) : lPos += 1
                        .RequiredProductionTypeID = yData(lPos) : lPos += 1
                        .yChassisType = yData(lPos) : lPos += 1
                        .yFXColor = yData(lPos) : lPos += 1
                        .ArmorIntegrity = yData(lPos) : lPos += 1
                        .JamImmunity = yData(lPos) : lPos += 1
                        .JamStrength = yData(lPos) : lPos += 1
                        .JamTargets = yData(lPos) : lPos += 1
                        .JamEffect = yData(lPos) : lPos += 1

                        'get our side crits
                        Dim lBits As BitArray
                        'Front
                        lTemp = System.BitConverter.ToInt32(yData, lPos)
                        lPos += 4
                        lBits = New BitArray(System.BitConverter.GetBytes(lTemp))
                        lTemp2 = -1
                        ReDim .lSide1Crits(-1)
                        For X = 0 To lBits.Length - 1
                            If lBits(X) = True Then
                                lTemp2 += 1
                                ReDim Preserve .lSide1Crits(lTemp2)
                                .lSide1Crits(lTemp2) = CInt(2 ^ X)
                            End If
                        Next X
                        lBits = Nothing
                        'Left
                        lTemp = System.BitConverter.ToInt32(yData, lPos)
                        lPos += 4
                        lBits = New BitArray(System.BitConverter.GetBytes(lTemp))
                        lTemp2 = -1
                        ReDim .lSide2Crits(-1)
                        For X = 0 To lBits.Length - 1
                            If lBits(X) = True Then
                                lTemp2 += 1
                                ReDim Preserve .lSide2Crits(lTemp2)
                                .lSide2Crits(lTemp2) = CInt(2 ^ X)
                            End If
                        Next X
                        lBits = Nothing
                        'Back
                        lTemp = System.BitConverter.ToInt32(yData, lPos)
                        lPos += 4
                        lBits = New BitArray(System.BitConverter.GetBytes(lTemp))
                        lTemp2 = -1
                        ReDim .lSide3Crits(-1)
                        For X = 0 To lBits.Length - 1
                            If lBits(X) = True Then
                                lTemp2 += 1
                                ReDim Preserve .lSide3Crits(lTemp2)
                                .lSide3Crits(lTemp2) = CInt(2 ^ X)
                            End If
                        Next X
                        lBits = Nothing
                        'Right
                        lTemp = System.BitConverter.ToInt32(yData, lPos)
                        lPos += 4
                        lBits = New BitArray(System.BitConverter.GetBytes(lTemp))
                        lTemp2 = -1
                        ReDim .lSide4Crits(-1)
                        For X = 0 To lBits.Length - 1
                            If lBits(X) = True Then
                                lTemp2 += 1
                                ReDim Preserve .lSide4Crits(lTemp2)
                                .lSide4Crits(lTemp2) = CInt(2 ^ X)
                            End If
                        Next X
                        lBits = Nothing


                    End With

                    'Now, the next 2 bytes indicates the number of weapon defs for this unit def
                    Dim iWDCnt As Int16 = System.BitConverter.ToInt16(yData, lPos) - 1S
                    lPos += 2

                    For X = 0 To iWDCnt
                        'the first byte is the arc id
                        goEntityDefs(lIdx).WeaponDefUB += 1
                        ReDim Preserve goEntityDefs(lIdx).WeaponDefs(goEntityDefs(lIdx).WeaponDefUB)
                        goEntityDefs(lIdx).WeaponDefs(goEntityDefs(lIdx).WeaponDefUB) = New WeaponDef()
                        With goEntityDefs(lIdx).WeaponDefs(goEntityDefs(lIdx).WeaponDefUB)
                            .ArcID = yData(lPos)
                            lPos += 1
                            'then, the next 59 bytes is the weapon def
                            .ObjectID = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4
                            .ObjTypeID = System.BitConverter.ToInt16(yData, lPos)
                            lPos += 2
                            Array.Copy(yData, lPos, .WeaponName, 0, 20)
                            lPos += 20
                            .yWeaponType = yData(lPos)
                            lPos += 1
                            .ROF_Delay = System.BitConverter.ToInt16(yData, lPos)
                            lPos += 2
                            .iDefRange = System.BitConverter.ToInt16(yData, lPos)
                            lPos += 2
                            .Accuracy = yData(lPos)
                            lPos += 1
                            .PiercingMinDmg = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4
                            .PiercingMaxDmg = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4
                            .ImpactMinDmg = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4
                            .ImpactMaxDmg = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4
                            .BeamMinDmg = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4
                            .BeamMaxDmg = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4
                            .ECMMinDmg = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4
                            .ECMMaxDmg = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4
                            .FlameMinDmg = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4
                            .FlameMaxDmg = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4
                            .ChemicalMinDmg = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4
                            .ChemicalMaxDmg = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4

                            .blOverallMaxDmg += CLng(.PiercingMaxDmg) + CLng(.ImpactMaxDmg) + CLng(.BeamMaxDmg) + CLng(.ECMMaxDmg) + CLng(.FlameMaxDmg) + CLng(.ChemicalMaxDmg)
                            If .ROF_Delay < 30 AndAlso .ROF_Delay > 0 Then
                                .blOverallMaxDmg = CLng(.blOverallMaxDmg * (30.0F / .ROF_Delay))
                            End If

                            .WpnGroup = yData(lPos)
                            lPos += 1
                            .lFirePowerRating = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4

                            lPos += 4   'Weapon Technology ID... skipped

                            .AOERange = yData(lPos) : lPos += 1
                            .WeaponSpeed = yData(lPos) : lPos += 1
                            If .WeaponSpeed = 0 Then .WeaponSpeed = 100
                            .Maneuver = yData(lPos) : lPos += 1

                            If .yWeaponType >= WeaponType.eShortGreenPulse AndAlso .yWeaponType <= WeaponType.eShortPurplePulse Then
                                Dim yVals(1) As Byte
                                yVals(0) = .WeaponSpeed : yVals(1) = .Maneuver
                                .fPulseDegradation = System.BitConverter.ToInt16(yVals, 0)
                                .fPulseDegradation *= 0.01F
                            ElseIf .yWeaponType >= WeaponType.eMissile_Color_1 AndAlso .yWeaponType <= WeaponType.eMissile_Color_9 Then
                                .lStructHP = .BeamMinDmg
                                .BeamMinDmg = 0
                            End If


                            'Entity Status Group
                            .lEntityStatusGroup = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4

                            .lAmmoCap = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4

                            .lEntityWpnDefID = System.BitConverter.ToInt32(yData, lPos)
                            lPos += 4

                        End With
                    Next X

                    goEntityDefs(lIdx).lExtendedFlags = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                    'Call optimize to finalize the object, or essentially unpackage it...
                    goEntityDefs(lIdx).OptimizeMe()
                End If
            Case ObjectType.eMineralCache, ObjectType.eComponentCache
                'first check if we already have this object...
                If iObjTypeID = ObjectType.eComponentCache Then
                    lParentID = System.BitConverter.ToInt32(yData, 8)
                    iParentTypeID = System.BitConverter.ToInt16(yData, 12)
                Else
                    lParentID = System.BitConverter.ToInt32(yData, 9)
                    iParentTypeID = System.BitConverter.ToInt16(yData, 13)
                End If
                Select Case iParentTypeID
                    Case ObjectType.eFacility, ObjectType.eUnit
                        Return
                    Case Else   'gotta be geography
                        lTemp = -1

                        Dim oEnvir As Envir = Nothing
                        If iParentTypeID = ObjectType.ePlanet AndAlso lParentID >= 500000000 Then
                            'instance
                            Dim lCurUB As Int32 = -1
                            If glInstanceIdx Is Nothing = False Then lCurUB = Math.Min(glInstanceUB, glInstanceIdx.GetUpperBound(0))
                            For X = 0 To lCurUB
                                If glInstanceIdx(X) = lParentID Then
                                    oEnvir = goInstances(X)
                                    Exit For
                                End If
                            Next X
                        Else
                            For X = 0 To glEnvirUB
                                If goEnvirs(X).ObjectID = lParentID AndAlso goEnvirs(X).ObjTypeID = iParentTypeID Then
                                    oEnvir = goEnvirs(X)
                                    Exit For
                                End If
                            Next X
                        End If

                        If oEnvir Is Nothing = False Then
                            With oEnvir
                                For X = 0 To .lCacheUB
                                    If .lCacheIdx(X) = lObjID AndAlso .oCache(X).ObjTypeID = iObjTypeID Then
                                        lIdx = X
                                        Exit For
                                    ElseIf lFirstIndex = -1 AndAlso .lCacheIdx(X) = -1 Then
                                        lFirstIndex = X
                                    End If
                                Next X

                                If lIdx = -1 Then
                                    'No... check lFirstIndex
                                    If lFirstIndex <> -1 Then
                                        'Ok, so set our object to that
                                        lIdx = lFirstIndex
                                    Else
                                        'Ok, we don't have a clear spot... so redim
                                        .lCacheUB += 1
                                        lIdx = .lCacheUB
                                        ReDim Preserve .lCacheIdx(.lCacheUB)
                                        ReDim Preserve .oCache(.lCacheUB)
                                    End If
                                    .oCache(lIdx) = New ObjectCache()
                                End If

                                'set our index
                                .lCacheIdx(lIdx) = lObjID
                            End With

                            'Now, fill our properties
                            With oEnvir.oCache(lIdx)
                                If iObjTypeID = ObjectType.eMineralCache Then
                                    .ObjectID = lObjID
                                    .ObjTypeID = iObjTypeID
                                    .CacheTypeID = yData(8)
                                    .LocX = System.BitConverter.ToInt32(yData, 15)
                                    .LocZ = System.BitConverter.ToInt32(yData, 19)
                                    .lDetail1 = System.BitConverter.ToInt32(yData, 23)
                                    .lQuantity = System.BitConverter.ToInt32(yData, 27)
                                    .lDetail2 = System.BitConverter.ToInt32(yData, 31)
                                Else
                                    .ObjectID = lObjID
                                    .ObjTypeID = iObjTypeID
                                    lPos = 14
                                    .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    .LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    .lQuantity = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                                    .lDetail1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4 'componentid
                                    'Not really used anymore...
                                    Dim iTempCacheType As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                                    .lDetail2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4 'component ownerid
                                    .CacheTypeID = yData(lPos) : lPos += 1
                                End If
                            End With

                            'Now, that we have our object... forward the command to connected players
                            If oEnvir.lPlayersInEnvirCnt > 0 Then
                                BroadcastToEnvironmentClients_Ex(iMsgCode, oEnvir.oCache(lIdx).GetObjectAsSmallString, oEnvir)
                            End If
                        End If
                End Select
            Case ObjectType.eWormhole
                If goGalaxy Is Nothing = False Then
                    Dim oWormhole As New Wormhole()
                    With oWormhole
                        .ObjectID = lObjID
                        .ObjTypeID = iObjTypeID
                        lPos = 8        'msgcode, objid, objtypeid
                        lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        For X = 0 To goGalaxy.mlSystemUB
                            If goGalaxy.moSystems(X).ObjectID = lTemp Then
                                .System1 = goGalaxy.moSystems(X)
                                Exit For
                            End If
                        Next X
                        lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        For X = 0 To goGalaxy.mlSystemUB
                            If goGalaxy.moSystems(X).ObjectID = lTemp Then
                                .System2 = goGalaxy.moSystems(X)
                                Exit For
                            End If
                        Next X
                        .LocX1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .LocY1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .LocX2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .LocY2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .StartCycle = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .WormholeFlags = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    End With
                    If oWormhole Is Nothing = False Then
                        If oWormhole.System1 Is Nothing = False Then oWormhole.System1.AddWormhole(oWormhole)
                        If oWormhole.System2 Is Nothing = False Then oWormhole.System2.AddWormhole(oWormhole)
                    End If
                Else : gfrmDisplayForm.AddEventLine("Received wormhole before getting galaxy!")
                End If
            Case ObjectType.eSolarSystem
                If goGalaxy Is Nothing = False Then
                    'ok, see if we have the system already
                    For X = 0 To goGalaxy.mlSystemUB
                        If goGalaxy.moSystems(X).ObjectID = lObjID Then
                            'we already have the system return
                            Return
                        End If
                    Next

                    'Ok, we don't have the system, let's add it
                    Dim oSystem As New SolarSystem()
                    With oSystem
                        .ObjectID = lObjID
                        .ObjTypeID = iObjTypeID
                        lPos = 8
                        .SystemName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
                        lPos += 4   'for parent galaxy
                        .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .LocY = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        .LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                        lTemp = yData(lPos) : lPos += 1
                        For X = 0 To glStarTypeUB
                            If goStarTypes(X).StarTypeID = lTemp Then
                                .StarType1Idx = X
                                Exit For
                            End If
                        Next X

                        lTemp = yData(lPos) : lPos += 1
                        For X = 0 To glStarTypeUB
                            If goStarTypes(X).StarTypeID = lTemp Then
                                .StarType2Idx = X
                                Exit For
                            End If
                        Next X

                        lTemp = yData(lPos) : lPos += 1
                        For X = 0 To glStarTypeUB
                            If goStarTypes(X).StarTypeID = lTemp Then
                                .StarType3Idx = X
                                Exit For
                            End If
                        Next X

                        .SystemType = yData(lPos) : lPos += 1
                        .FleetJumpPointX = yData(lPos) : lPos += 4
                        .FleetJumpPointZ = yData(lPos) : lPos += 4
                        goGalaxy.AddSystem(oSystem)
                    End With

                End If
            Case ObjectType.eMineral
                'SPECIAL CASE!!! We are not actually adding a mineral because, frankly, the region server does not care... we are
                '  requesting that the region server give us a valid location for the mineral cache passed in
                lPos = 8
                Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                Dim lMaxQty As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lMaxConc As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                Dim lEnvirUB As Int32 = glEnvirUB
                Dim pt As Point = Point.Empty
                For X = 0 To lEnvirUB
                    If goEnvirs(X) Is Nothing = False AndAlso goEnvirs(X).ObjectID = lEnvirID AndAlso goEnvirs(X).ObjTypeID = iEnvirTypeID Then
                        pt = goEnvirs(X).PlaceMineralCache()
                        Exit For
                    End If
                Next X
                If pt <> Point.Empty Then
                    Dim yMsg(29) As Byte
                    lPos = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lObjID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(iObjTypeID).CopyTo(yMsg, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
                    System.BitConverter.GetBytes(lMaxQty).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(lMaxConc).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(pt.X).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(pt.Y).CopyTo(yMsg, lPos) : lPos += 4
                    moPrimary.SendData(yMsg)
                End If


        End Select
    End Sub

    Private Sub ReceiveAddPlayerRel(ByVal yData() As Byte)
        'ok... this is a special message
        ' 2 byte message code... 4 Player Regards, 4 byte This Player, 1 byte Score
        Dim lPlayerRegards As Int32
        Dim lThisPlayer As Int32
        Dim X As Int32
        Dim oTmpRel As PlayerRel = New PlayerRel()
        Dim bP1Found As Boolean = False
        Dim bP2Found As Boolean = False

        lPlayerRegards = System.BitConverter.ToInt32(yData, 2)
        lThisPlayer = System.BitConverter.ToInt32(yData, 6)
        oTmpRel.WithThisScore = yData(10)
        oTmpRel.OtherTowardsMe = elRelTypes.eNeutral

        'now, find our players
        For X = 0 To glPlayerUB
            If glPlayerIdx(X) = lPlayerRegards Then
                oTmpRel.oPlayerRegards = goPlayers(X)
                bP1Found = True
            ElseIf glPlayerIdx(X) = lThisPlayer Then
                oTmpRel.oThisPlayer = goPlayers(X)

                Dim oOtherRel As PlayerRel = goPlayers(X).GetPlayerRel(lPlayerRegards)
                If oOtherRel Is Nothing = False Then
                    oOtherRel.OtherTowardsMe = oTmpRel.WithThisScore
                    oTmpRel.OtherTowardsMe = oOtherRel.WithThisScore
                Else
                    oTmpRel.OtherTowardsMe = goPlayers(X).GetPlayerRelScore(lPlayerRegards, False, -1)
                End If

                bP2Found = True
            End If
            If bP1Found AndAlso bP2Found Then
                Exit For
            End If
        Next X

        'oTmpRel.oPlayerRegards.colPlayerRels.Add(oTmpRel, "PR_" & lThisPlayer)
        oTmpRel.oPlayerRegards.SetPlayerRel(lThisPlayer, oTmpRel)

    End Sub

    Private Sub HandleSetPlayerRel(ByVal yData() As Byte)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, 6)
        Dim yRel As Byte = yData(10)

        Dim X As Int32
        Dim Y As Int32
        Dim bFound As Boolean = False

        Dim bProcessPeace As Boolean = False

        For X = 0 To glPlayerUB
            If glPlayerIdx(X) = lPlayerID Then

                For Y = 0 To glPlayerUB
                    If glPlayerIdx(Y) = lTargetID Then

                        Dim oTmpRel As PlayerRel = New PlayerRel()
                        oTmpRel.oPlayerRegards = goPlayers(X)
                        oTmpRel.oThisPlayer = goPlayers(Y)
                        oTmpRel.WithThisScore = yRel

                        Dim oTheirRel As PlayerRel = goPlayers(Y).GetPlayerRel(lPlayerID)
                        If oTheirRel Is Nothing = False Then
                            oTmpRel.OtherTowardsMe = oTheirRel.WithThisScore
                            oTheirRel.OtherTowardsMe = yRel
                        Else
                            oTmpRel.OtherTowardsMe = elRelTypes.eNeutral
                        End If

                        If yRel <= elRelTypes.eWar Then
                            'Because one player declared war, both do...
                            'oTmpRel.OtherTowardsMe = yRel
                            goPlayers(X).SetPlayerRel(lTargetID, oTmpRel)

                            'oTmpRel = Nothing
                            'oTmpRel = New PlayerRel()
                            'oTmpRel.oPlayerRegards = goPlayers(Y)
                            'oTmpRel.oThisPlayer = goPlayers(X)
                            'oTmpRel.WithThisScore = yRel
                            'oTmpRel.OtherTowardsMe = yRel
                            'goPlayers(Y).SetPlayerRel(lPlayerID, oTmpRel)
                        Else
                            If oTmpRel.OtherTowardsMe > elRelTypes.eWar Then bProcessPeace = True
                            goPlayers(X).SetPlayerRel(lTargetID, oTmpRel)
                        End If

                        bFound = True
                        Exit For
                    End If
                Next Y

                If bFound = True Then Exit For
            End If
        Next X
        If bFound = True Then
            For X = 0 To glEnvirUB
                goEnvirs(X).TestForAggression()
                goEnvirs(X).TestForFirstContact()
            Next X

            If yRel <= elRelTypes.eWar Then
                Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
                For X = 0 To lCurUB
                    If glEntityIdx(X) > 0 Then
                        Dim oEntity As Epica_Entity = goEntity(X)
                        If oEntity Is Nothing = False AndAlso (oEntity.Owner.ObjectID = lPlayerID OrElse oEntity.Owner.ObjectID = lTargetID) Then
                            If oEntity.ParentEnvir.PotentialAggression = True Then
                                oEntity.bForceAggressionTest = True
                                SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oEntity.ServerIndex, oEntity.ObjectID) 'AddEntityMoving(oEntity.ServerIndex, oEntity.ObjectID)
                            End If
                        End If
                    End If
                Next X
            ElseIf bProcessPeace = True Then
                'Be sure that our units no longer fight...
                Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
                For X = 0 To lCurUB
                    If glEntityIdx(X) > 0 Then
                        Dim oEntity As Epica_Entity = goEntity(X)
                        If oEntity Is Nothing = False AndAlso (oEntity.Owner.ObjectID = lPlayerID OrElse oEntity.Owner.ObjectID = lTargetID) Then
                            If (oEntity.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
                                If oEntity.lPrimaryTargetServerIdx <> -1 AndAlso glEntityIdx(oEntity.lPrimaryTargetServerIdx) > 0 Then
                                    Dim oTarget As Epica_Entity = goEntity(oEntity.lPrimaryTargetServerIdx)
                                    If oTarget Is Nothing = False Then
                                        If oTarget.Owner.ObjectID = lPlayerID OrElse oTarget.Owner.ObjectID = lTargetID Then
                                            oEntity.lPrimaryTargetServerIdx = -1
                                        End If
                                    End If

                                    For Y = 0 To oEntity.lTargetsServerIdx.GetUpperBound(0)
                                        If oEntity.lTargetsServerIdx(Y) <> -1 Then 'AndAlso glEntityIdx(goEntity(X).lTargetsServerIdx(Y)) <> -1 Then
                                            If glEntityIdx(oEntity.lTargetsServerIdx(Y)) > 0 Then
                                                oTarget = goEntity(oEntity.lTargetsServerIdx(Y))
                                                If oTarget Is Nothing = False Then
                                                    If oTarget.Owner.ObjectID = lPlayerID OrElse oTarget.Owner.ObjectID = lTargetID Then
                                                        oEntity.lTargetsServerIdx(Y) = -1
                                                    End If
                                                End If
                                            Else : oEntity.lTargetsServerIdx(Y) = -1
                                            End If
                                        End If
                                    Next Y
                                    If oEntity.ParentEnvir.PotentialAggression = True Then
                                        oEntity.bForceAggressionTest = True
                                        SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oEntity.ServerIndex, oEntity.ObjectID) 'AddEntityMoving(oEntity.ServerIndex, oEntity.ObjectID)
                                    End If
                                End If
                                oEntity.CurrentStatus = oEntity.CurrentStatus Xor elUnitStatus.eUnitEngaged ' oEntity.CurrentStatus -= elUnitStatus.eUnitEngaged
                                RemoveEntityInCombat(oEntity.ObjectID, oEntity.ObjTypeID)
                            End If
                        End If
                    End If
                Next X

            End If
        End If
    End Sub

    Public Function GetPlayerListFromPrimary() As Boolean
        'returns whether the player list was returned or not...
        Dim yRequest() As Byte
        Dim bRes As Boolean = False
        Dim sw As Stopwatch

        ReDim yRequest(1)
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerList).CopyTo(yRequest, 0)

        mbSyncWait = True
        sw = Stopwatch.StartNew
        moPrimary.SendData(yRequest)
        While mbSyncWait = True 'And (sw.ElapsedMilliseconds < mlTimeout)
            Application.DoEvents()
        End While
        If mbSyncWait = True Then
            mbSyncWait = False
            Return False
        End If

        'Next, request the Player Rels
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerRelList).CopyTo(yRequest, 0)
        mbSyncWait = True
        sw.Reset()
        sw.Start()
        moPrimary.SendData(yRequest)
        While mbSyncWait 'And (sw.ElapsedMilliseconds < mlTimeout)
            Application.DoEvents()
        End While
        If mbSyncWait = True Then
            mbSyncWait = False
            Return False
        End If
        sw.Stop()
        sw = Nothing

        Return True
    End Function

    Public Function GetStarTypes() As Boolean
        Dim yRequest(2) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestStarTypes).CopyTo(yRequest, 0)

        mbSyncWait = True
        Dim sw As Stopwatch = Stopwatch.StartNew
        moPrimary.SendData(yRequest)
        While mbSyncWait = True 'And (sw.ElapsedMilliseconds < mlTimeout)
            Application.DoEvents()
        End While
        sw.Stop()
        sw = Nothing
        If mbSyncWait = True Then
            mbSyncWait = False
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub HandleStarTypesMsg(ByVal yData() As Byte)
        'first two bytes is the message ID
        Dim lPos As Int32 = 2

        glStarTypeUB = -1
        ReDim goStarTypes(-1)

        While lPos < yData.Length - 1
            glStarTypeUB += 1
            ReDim Preserve goStarTypes(glStarTypeUB)
            goStarTypes(glStarTypeUB) = New StarType()
            With goStarTypes(glStarTypeUB)
                .StarTypeID = yData(lPos)
                lPos += 1

                lPos += 20

                .StarTypeAttrs = System.BitConverter.ToInt32(yData, lPos)
                lPos += 4
                lPos += 20

                .HeatIndex = yData(lPos)
                lPos += 1

                lPos += 4
                lPos += 4

                lPos += 1
                lPos += 4
                lPos += 4
                lPos += 4
                lPos += 4
                lPos += 4
                lPos += 4
                lPos += 4
                lPos += 1
                .StarRadius = System.BitConverter.ToInt32(yData, lPos)
                lPos += 4
            End With
        End While
        mbSyncWait = False
    End Sub

    Public Function GetGalaxyAndSystems() As Boolean
        Dim yRequest(2) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestGalaxyAndSystems).CopyTo(yRequest, 0)

        mbSyncWait = True
        Dim sw As Stopwatch = Stopwatch.StartNew
        moPrimary.SendData(yRequest)
        While mbSyncWait = True 'And (sw.ElapsedMilliseconds < mlTimeout)
            Application.DoEvents()
        End While
        sw.Stop()
        sw = Nothing
        If mbSyncWait = True Then
            mbSyncWait = False
            Return False
        Else
            Return True
        End If

    End Function

    Private Sub HandleGalaxyAndSystemsMsg(ByVal yData() As Byte)
        'Ok, time to light up the ovens... this message contains the Galaxy Objects and System Objects
        Dim lPos As Int32 = 2   'the position of our cursor, inc. by 2 for the msg ID
        Dim lObjID As Int32
        Dim iTypeID As Int16

        Dim lTemp As Int32

        Dim X As Int32

        Dim yTempName(19) As Byte       '20 byte names

        'Ok set out galaxy to new
        goGalaxy = New Galaxy()

        While lPos < yData.Length - 1
            'whether it is a galaxy or a system, the first 4 bytes is the ID
            lObjID = System.BitConverter.ToInt32(yData, lPos)
            lPos += 4
            'then the type id
            iTypeID = System.BitConverter.ToInt16(yData, lPos)
            lPos += 2

            'Now, we can determine whether it is a galaxy or system
            If iTypeID = ObjectType.eGalaxy Then
                'galaxy... its our galaxy definition... we will only load one galaxy at a time...
                With goGalaxy
                    .ObjectID = lObjID
                    .ObjTypeID = iTypeID

                    Array.Copy(yData, lPos, yTempName, 0, 20)
                    lPos += 20

                    .GalaxyName = System.Text.ASCIIEncoding.ASCII.GetString(yTempName)
                    lPos += 4
                End With
            ElseIf iTypeID = ObjectType.eSolarSystem Then
                'solar system
                goGalaxy.mlSystemUB += 1
                ReDim Preserve goGalaxy.moSystems(goGalaxy.mlSystemUB)
                goGalaxy.moSystems(goGalaxy.mlSystemUB) = New SolarSystem()
                With goGalaxy.moSystems(goGalaxy.mlSystemUB)
                    .ObjectID = lObjID
                    .ObjTypeID = iTypeID

                    Array.Copy(yData, lPos, yTempName, 0, 20)
                    lPos += 20
                    .SystemName = System.Text.ASCIIEncoding.ASCII.GetString(yTempName)

                    lTemp = System.BitConverter.ToInt32(yData, lPos)
                    'Systems are contained within Galaxies, but since we don't really care about multiple galaxies
                    '  at this time, then we'll just pass this up for now                    
                    lPos += 4

                    .LocX = System.BitConverter.ToInt32(yData, lPos)
                    lPos += 4
                    .LocY = System.BitConverter.ToInt32(yData, lPos)
                    lPos += 4
                    .LocZ = System.BitConverter.ToInt32(yData, lPos)
                    lPos += 4

                    lTemp = yData(lPos)
                    lPos += 1
                    'NOTE: We should have already called AND received our star types by now
                    For X = 0 To glStarTypeUB
                        If goStarTypes(X).StarTypeID = lTemp Then
                            .StarType1Idx = X
                            Exit For
                        End If
                    Next X

                    lTemp = yData(lPos)
                    lPos += 1
                    'NOTE: We should have already called AND received our star types by now
                    For X = 0 To glStarTypeUB
                        If goStarTypes(X).StarTypeID = lTemp Then
                            .StarType2Idx = X
                            Exit For
                        End If
                    Next X

                    lTemp = yData(lPos)
                    lPos += 1
                    'NOTE: We should have already called AND received our star types by now
                    For X = 0 To glStarTypeUB
                        If goStarTypes(X).StarTypeID = lTemp Then
                            .StarType3Idx = X
                            Exit For
                        End If
                    Next X

                    .SystemType = yData(lPos) : lPos += 1
                    .FleetJumpPointX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .FleetJumpPointZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                End With
            Else
                'This was bad, so we're just gonna slink away and forget it ever happened
                Application.Exit()
            End If
        End While
        mbSyncWait = False
    End Sub

    Public Function GetSystemDetails(ByVal lID As Int32) As Boolean
        Dim yData(5) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestSystemDetails).CopyTo(yData, 0)
        System.BitConverter.GetBytes(lID).CopyTo(yData, 2)

        'mbSyncWait = True
        'Dim sw As Stopwatch = Stopwatch.StartNew
        moPrimary.SendData(yData)
        'While mbSyncWait = True	'And (sw.ElapsedMilliseconds < mlTimeout)
        '	Application.DoEvents()
        'End While
        'sw.Stop()
        'sw = Nothing
        'If mbSyncWait = True Then
        '	mbSyncWait = False
        '	Return False
        'Else
        Return True
        'End If
    End Function

    'Private Sub HandleSystemDetailsMsg(ByVal yData() As Byte)
    '	Dim lPos As Int32 = 2	'msg id
    '	Dim oTmpPlnt As Planet
    '	Dim lID As Int32
    '	Dim iTypeID As Int16
    '	Dim sName As String
    '	Dim ySizeID As Byte
    '	Dim yMapTypeID As Byte
    '	Dim yTempName(19) As Byte

    '	Dim lInnerRadius As Int32
    '	Dim lOuterRadius As Int32
    '	Dim lRingDiffuse As Int32

    '	'ok, got our details, this is nothing more than a list of Planet objects... we would only get this message
    '	'  for our galaxy's current system
    '	If goGalaxy.CurrentSystemIdx <> -1 Then
    '		While lPos < yData.Length - 1


    '			'Object ID
    '			lID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '			'ObjTypeID
    '			iTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
    '			'Name
    '			Array.Copy(yData, lPos, yTempName, 0, 20)
    '			lPos += 20
    '			sName = System.Text.ASCIIEncoding.ASCII.GetString(yTempName)
    '			'Planet type id
    '			yMapTypeID = yData(lPos) : lPos += 1
    '			'Size ID
    '			ySizeID = yData(lPos) : lPos += 1

    '			'Now, create our planet
    '			oTmpPlnt = New Planet(lID, ySizeID, yMapTypeID)

    '			With oTmpPlnt
    '				.ObjTypeID = iTypeID
    '				.PlanetName = sName
    '				'ok, now the other attributes
    '				.PlanetRadius = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
    '				.ParentSystem = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx) : lPos += 4
    '				.LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '				.LocY = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '				.LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '				.Vegetation = yData(lPos) : lPos += 1
    '				.Atmosphere = yData(lPos) : lPos += 1
    '				.Hydrosphere = yData(lPos) : lPos += 1
    '				.Gravity = yData(lPos) : lPos += 1
    '				.SurfaceTemperature = yData(lPos) : lPos += 1
    '				.RotationDelay = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

    '				'Now, check for our ring...
    '				lInnerRadius = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '				lOuterRadius = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '				lRingDiffuse = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

    '				'finally, the axis angle
    '				.AxisAngle = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

    '				.RotateAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

    '				.PopulateData()
    '			End With
    '			goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).AddPlanet(oTmpPlnt)
    '			oTmpPlnt = Nothing
    '		End While
    '	End If
    'End Sub

    Private Sub HandleSystemDetailsMsg(ByRef yData() As Byte)
        Dim lPos As Int32 = 2       'for msgcode
        Dim lSystemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oSystem As SolarSystem = Nothing
        For X As Int32 = 0 To goGalaxy.mlSystemUB
            If goGalaxy.moSystems(X) Is Nothing = False AndAlso goGalaxy.moSystems(X).ObjectID = lSystemID Then
                oSystem = goGalaxy.moSystems(X)
                Exit For
            End If
        Next X
        If oSystem Is Nothing Then
            gfrmDisplayForm.AddEventLine("Error: System " & lSystemID & " not found in galaxy object!")
            Return
        End If

        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        For X As Int32 = 0 To lCnt - 1
            'Object ID
            Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'ObjTypeID
            Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            'Name
            Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
            'Planet type id
            Dim yMapTypeID As Byte = yData(lPos) : lPos += 1
            'Size ID
            Dim ySizeID As Byte = yData(lPos) : lPos += 1

            'Now, create our planet
            Dim oTmpPlnt As New Planet(lID, ySizeID, yMapTypeID)
            With oTmpPlnt
                .ObjTypeID = iTypeID
                .PlanetName = sName
                'ok, now the other attributes
                .PlanetRadius = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .ParentSystem = oSystem : lPos += 4
                .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .LocY = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .Vegetation = yData(lPos) : lPos += 1
                .Atmosphere = yData(lPos) : lPos += 1
                .Hydrosphere = yData(lPos) : lPos += 1
                .Gravity = yData(lPos) : lPos += 1
                .SurfaceTemperature = yData(lPos) : lPos += 1
                .RotationDelay = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                'Now, check for our ring...
                Dim lInnerRadius As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lOuterRadius As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lRingDiffuse As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                'finally, the axis angle
                .AxisAngle = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                .RotateAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

                .PopulateData()
            End With

            oSystem.AddPlanet(oTmpPlnt)
            oTmpPlnt = Nothing
        Next X

        Dim oThread As New Threading.Thread(AddressOf GeoSpawner.ReceivedSystemDetails)
        oThread.Start()
    End Sub

    Public Sub SendRegisterDomainMsg(ByVal lObjID As Int32, ByVal iObjTypeID As Int16)
        Dim yData() As Byte

        ReDim yData(9)
        System.BitConverter.GetBytes(GlobalMessageCode.eRegisterDomain).CopyTo(yData, 0)
        System.BitConverter.GetBytes(lObjID).CopyTo(yData, 2)
        System.BitConverter.GetBytes(iObjTypeID).CopyTo(yData, 6)

        'informing the Primary AND the Pathfinding
        moPrimary.SendData(yData)
        moPathfinding.SendData(yData)
    End Sub

    Public Sub SendRequestEnvirObjs(ByVal lObjID As Int32, ByVal iObjTypeID As Int16)
        Dim yRequest(7) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eDomainRequestEnvirObjects).CopyTo(yRequest, 0)
        System.BitConverter.GetBytes(lObjID).CopyTo(yRequest, 2)
        System.BitConverter.GetBytes(iObjTypeID).CopyTo(yRequest, 6)
        'mbSyncWait = True
        moPrimary.SendData(yRequest)
        'While mbSyncWait = True
        '	Application.DoEvents()
        'End While
    End Sub

    Public Function CreateEntityHPMsg(ByVal oEntity As Epica_Entity, ByVal oEntityDef As Epica_Entity_Def) As Byte()
        Dim yMsg(17) As Byte
        Dim X As Int32

        With oEntity
            'MsgID (2)
            System.BitConverter.GetBytes(GlobalMessageCode.eUnitHPUpdate).CopyTo(yMsg, 0)
            'Entity ID (4), Type ID (2)
            .GetGUIDAsString.CopyTo(yMsg, 2)

            ' Armor(4)
            For X = 0 To 3
                If oEntityDef.Armor_MaxHP(X) = 0 Then
                    yMsg(8 + X) = 255 '100
                Else
                    If .ArmorHP(X) < 0 Then .ArmorHP(X) = 0
                    yMsg(8 + X) = CByte((CSng(.ArmorHP(X)) / oEntityDef.Armor_MaxHP(X)) * 100)
                    If yMsg(8 + X) = 0 AndAlso .ArmorHP(X) > 0 Then yMsg(8 + X) = 1
                End If
            Next X
            ' Shld(1)
            If oEntityDef.Shield_MaxHP = 0 Then
                yMsg(12) = 0
            Else
                If .ShieldHP < 0 Then .ShieldHP = 0
                yMsg(12) = CByte((CSng(.ShieldHP) / oEntityDef.Shield_MaxHP) * 100)
                If yMsg(12) = 0 AndAlso .ShieldHP > 0 Then yMsg(12) = 1
            End If

            ' Structure (1)
            If oEntityDef.Structure_MaxHP = 0 Then
                'ensure that structurehp = 0
                yMsg(13) = 0
            Else
                If .StructuralHP < 0 Then .StructuralHP = 0
                yMsg(13) = CByte((CSng(.StructuralHP) / oEntityDef.Structure_MaxHP) * 100)
                If yMsg(13) = 0 AndAlso .StructuralHP > 0 Then yMsg(13) = 1
            End If

            'Now, only send what values are NOT working
            If oEntity.lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(oEntity.lEntityDefServerIndex) <> -1 Then
                X = goEntityDefs(oEntity.lEntityDefServerIndex).GetBaseCritLocs()
            Else
                X = 0
            End If
            System.BitConverter.GetBytes(oEntity.CurrentStatus Xor X).CopyTo(yMsg, 14)

        End With

        Return yMsg
    End Function

    Public Sub CheckForBroadcastEntityHP(ByRef oEntity As Epica_Entity, ByRef oEntityDef As Epica_Entity_Def)

        Dim bNeedToSend As Boolean = False
        Dim yArmor(3) As Byte
        Dim yStruct As Byte
        Dim yShield As Byte
        Dim X As Int32

        With oEntity
            'If a critical is hit, update
            Dim lTemp As Int32
            Dim lTempWhatWeCareAbout As Int32 = elUnitStatus.eAftWeapon1 Or elUnitStatus.eAftWeapon2 Or elUnitStatus.eCargoBayOperational Or elUnitStatus.eEngineOperational Or elUnitStatus.eFacilityPowered Or elUnitStatus.eForwardWeapon1 Or elUnitStatus.eForwardWeapon2 Or elUnitStatus.eFuelBayOperational Or elUnitStatus.eHangarOperational Or elUnitStatus.eLeftWeapon1 Or elUnitStatus.eLeftWeapon2 Or elUnitStatus.eRadarOperational Or elUnitStatus.eRightWeapon1 Or elUnitStatus.eRightWeapon2 Or elUnitStatus.eShieldOperational
            X = oEntityDef.GetBaseCritLocs()
            lTemp = (oEntity.CurrentStatus And lTempWhatWeCareAbout) Xor X

            bNeedToSend = (.PreviousStatus <> lTemp)
            .PreviousStatus = lTemp

            'PreviousShield = 0 orelse currentshield = 0 andalso previousshield <> currentshield
            If oEntityDef.Shield_MaxHP = 0 Then
                yShield = 0
            Else
                Dim lTmpShield As Int32 = .ShieldHP
                'If .ShieldHP < 0 Then .ShieldHP = 0
                If lTmpShield < 0 Then lTmpShield = 0
                If lTmpShield > oEntityDef.Shield_MaxHP Then lTmpShield = oEntityDef.Shield_MaxHP
                yShield = CByte((CSng(lTmpShield) / oEntityDef.Shield_MaxHP) * 100)
                If yShield = 0 AndAlso lTmpShield > 0 Then yShield = 1
            End If
            bNeedToSend = bNeedToSend OrElse (.PreviousShieldPerc <> yShield AndAlso (.PreviousShieldPerc = 0 OrElse yShield = 0))
            .PreviousShieldPerc = yShield

            'an armor side mod 4 <> previous value
            For X = 0 To 3
                If oEntityDef.Armor_MaxHP(X) = 0 Then
                    yArmor(X) = 0
                Else
                    If .ArmorHP(X) < 0 Then .ArmorHP(X) = 0
                    yArmor(X) = CByte((CSng(.ArmorHP(X)) / oEntityDef.Armor_MaxHP(X)) * 100)
                    If yArmor(X) = 0 AndAlso .ArmorHP(X) > 0 Then yArmor(X) = 1
                End If

                If bNeedToSend = False Then
                    If CInt(.PreviousArmorPerc(X)) \ 25 <> CInt(yArmor(X)) \ 25 Then
                        bNeedToSend = True
                    End If
                End If
                .PreviousArmorPerc(X) = yArmor(X)
            Next X

            'structure hp was > 50 and is now < 50
            If oEntityDef.Structure_MaxHP = 0 Then
                'ensure that structurehp = 0
                yStruct = 0
            Else
                If .StructuralHP < 0 Then .StructuralHP = 0
                yStruct = CByte((CSng(.StructuralHP) / oEntityDef.Structure_MaxHP) * 100)
                If yStruct = 0 AndAlso .StructuralHP > 0 Then yStruct = 1
            End If
            bNeedToSend = bNeedToSend OrElse (.PreviousStructPerc > 50 AndAlso yStruct < 50)
            .PreviousStructPerc = yStruct

            'Now, do we need to send this?
            If bNeedToSend = True Then
                Dim yMsg(17) As Byte

                'MsgID (2)
                System.BitConverter.GetBytes(GlobalMessageCode.eUnitHPUpdate).CopyTo(yMsg, 0)
                'Entity ID (4), Type ID (2)
                .GetGUIDAsString.CopyTo(yMsg, 2)

                ' Armor(4)
                For X = 0 To 3
                    yMsg(8 + X) = yArmor(X)
                Next X
                ' Shld(1)
                yMsg(12) = yShield
                ' Structure (1)
                yMsg(13) = yStruct

                'Now, what we care about for status...
                System.BitConverter.GetBytes(.PreviousStatus).CopyTo(yMsg, 14)

                BroadcastToEnvironmentClients(yMsg, .ParentEnvir)
            End If
        End With

    End Sub

    Public Sub SendServerReady()
        Dim yData(7) As Byte
        Dim iPort As Int16
        'Dim oINI As New InitFile()
        'Dim moHostInfo As System.Net.IPHostEntry = System.Net.Dns.GetHostByName(Environment.MachineName)
        'Dim moAddress As System.Net.IPAddress = moHostInfo.AddressList(0)          'and this
        Dim sIPAddy() As String = Split(gsExternalIP, ".") ' Split(moAddress.ToString(), ".")

        System.BitConverter.GetBytes(GlobalMessageCode.eDomainServerReady).CopyTo(yData, 0)

        iPort = CShort(glExternalPort) 'CShort(Val(oINI.GetString("SETTINGS", "ClientListenPort", "7710")))

        System.BitConverter.GetBytes(iPort).CopyTo(yData, 2)
        yData(4) = CByte(Val(sIPAddy(0)))
        yData(5) = CByte(Val(sIPAddy(1)))
        yData(6) = CByte(Val(sIPAddy(2)))
        yData(7) = CByte(Val(sIPAddy(3)))

        'oINI = Nothing
        'moAddress = Nothing
        'moHostInfo = Nothing

        moPrimary.SendData(yData)
    End Sub

    Private Sub HandleChangeEnvironment(ByVal yData() As Byte, ByVal lIndex As Int32)
        'Msg should contains the Player ID, the new Envir ID, the new Envir Type ID
        'Dim X As Int32
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, 6)
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 10)
        'Dim lIdx As Int32 = -1
        'Dim Y As Int32

        If mlClientPlayer(lIndex) = lPlayerID OrElse lPlayerID = mlAliasedAs(lIndex) Then
            If lPlayerID <> mlClientPlayer(lIndex) Then lPlayerID = mlClientPlayer(lIndex)

            ''But, all we care about is removing them from our index (for now)
            'For X = 0 To glPlayerUB
            '	If glPlayerIdx(X) = lPlayerID Then
            '		lIdx = goPlayers(X).lEnvirIdx
            '		If lIdx > -1 Then

            '			If goPlayers(X).yPlayerPhase = 0 Then
            '				'instance
            '				If lIdx <= glInstanceUB AndAlso glInstanceIdx(lIdx) <> -1 Then
            '					goInstances(lIdx).lPlayersInEnvirCnt -= 1
            '					For Y = 0 To goInstances(lIdx).lPlayersInEnvirUB
            '						If goInstances(lIdx).lPlayersInEnvirIdx(Y) = lPlayerID Then
            '							goInstances(lIdx).oPlayersInEnvir(Y) = Nothing
            '							goInstances(lIdx).lPlayersInEnvirIdx(Y) = -1
            '							'Exit For
            '						End If
            '					Next Y
            '				End If
            '			Else
            '				goEnvirs(lIdx).lPlayersInEnvirCnt -= 1
            '				For Y = 0 To goEnvirs(lIdx).lPlayersInEnvirUB
            '					If goEnvirs(lIdx).lPlayersInEnvirIdx(Y) = lPlayerID Then
            '						goEnvirs(lIdx).oPlayersInEnvir(Y) = Nothing
            '						goEnvirs(lIdx).lPlayersInEnvirIdx(Y) = -1
            '						'Exit For
            '					End If
            '				Next Y
            '			End If

            '		End If

            '		Exit For
            '	End If
            'Next X
        Else
            gfrmDisplayForm.AddEventLine("Possible Cheat: HandleChangeEnvironment for player " & lPlayerID & " mismatch. SocketPlayerID: " & mlClientPlayer(lIndex))
        End If

    End Sub

    Public Function CreateRemoveObjectMsg(ByVal lObjID As Int32, ByVal iObjTypeID As Int16, ByVal yRemoveType As RemovalType, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal lStatus As Int32, ByVal lKilledByID As Int32) As Byte()
        'ok, here's how this works...
        Dim yResp(24) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eRemoveObject).CopyTo(yResp, 0)
        System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
        System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
        yResp(8) = yRemoveType
        System.BitConverter.GetBytes(lLocX).CopyTo(yResp, 9)
        System.BitConverter.GetBytes(lLocZ).CopyTo(yResp, 13)
        System.BitConverter.GetBytes(lStatus).CopyTo(yResp, 17)
        System.BitConverter.GetBytes(lKilledByID).CopyTo(yResp, 21)

        Return yResp
    End Function

    Private Sub HandleClientRequestHPMsg(ByVal yData() As Byte, ByVal lSocketIndex As Int32)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        'Dim X As Int32
        Dim oDef As Epica_Entity_Def

        'TODO: Possible Bombard Exploit here where a player bombards the server with hp requests!!!

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        If oEntity.lEntityDefServerIndex <> -1 Then
            oDef = goEntityDefs(oEntity.lEntityDefServerIndex)
            If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eUnitHPUpdate, MsgMonitor.eMM_AppType.ClientConnection, 18, mlClientPlayer(lSocketIndex))
            moClients(lSocketIndex).SendData(CreateEntityHPMsg(oEntity, oDef))
            oDef = Nothing
        End If
        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X = 0 To lCurUB
        '    If glEntityIdx(X) = lObjID Then
        '        Dim oEntity As Epica_Entity = goEntity(X)
        '        If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iObjTypeID Then
        '            If oEntity.lEntityDefServerIndex <> -1 Then
        '                oDef = goEntityDefs(oEntity.lEntityDefServerIndex)
        '                If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eUnitHPUpdate, MsgMonitor.eMM_AppType.ClientConnection, 18, mlClientPlayer(lSocketIndex))
        '                moClients(lSocketIndex).SendData(CreateEntityHPMsg(oEntity, oDef))
        '                oDef = Nothing
        '            End If
        '            Exit For
        '        End If
        '    End If
        'Next X
    End Sub

    Public Sub SendDecelerationImminentMsg(ByVal oObj As Epica_Entity)
        'DI msg is a ObjID (4), ObjTypeID (2), DestX (4), DestZ (4), DestAngle (2), CurLocX (4), CurLocZ (4), CurAngle (2) 
        '  UnitManuever (1)
        ' we take this data and we calculate if there are any obstacles in the way...
        Dim yData(28) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eDecelerationImminent).CopyTo(yData, 0)
        oObj.GetGUIDAsString().CopyTo(yData, 2)

        'Now, the dest X, Z, A
        System.BitConverter.GetBytes(oObj.DestX).CopyTo(yData, 8)
        System.BitConverter.GetBytes(oObj.DestZ).CopyTo(yData, 12)
        System.BitConverter.GetBytes(oObj.DestAngle).CopyTo(yData, 16)

        'Now, the current X, Z, A
        System.BitConverter.GetBytes(oObj.LocX).CopyTo(yData, 18)
        System.BitConverter.GetBytes(oObj.LocZ).CopyTo(yData, 22)
        System.BitConverter.GetBytes(oObj.LocAngle).CopyTo(yData, 26)

        'the manuever property
        'If oObj.lEntityDefServerIndex <> -1 Then
        '    If glEntityDefIdx(oObj.lEntityDefServerIndex) <> -1 Then
        '        yData(28) = goEntityDefs(oObj.lEntityDefServerIndex).Maneuver
        '    End If
        'End If
        yData(28) = oObj.Maneuver

        SendToPathfinding(yData)

    End Sub

    Private Sub HandleRequestDock(ByVal yData() As Byte, ByRef oSocket As NetSock, ByVal lClientIndex As Int32)
        If lClientIndex <> -1 Then
            'If glCurrentCycle - mlClientLastRequest(lClientIndex) > ml_REQUEST_THROTTLE Then
            If glCurrentCycle - muClientData(lClientIndex).mlClientLastRequest > ml_REQUEST_THROTTLE Then
                'mlClientLastRequest(lClientIndex) = glCurrentCycle
                muClientData(lClientIndex).mlClientLastRequest = glCurrentCycle
            Else : Return
            End If
        End If

        Dim yForward() As Byte
        Dim lPos As Int32 = 2
        Dim lDestPos As Int32
        Dim lDestLen As Int32

        Dim lDockeeID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDockeeTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        'Dim X As Int32

        Dim oDockee As Epica_Entity = Nothing
        Dim oDocker As Epica_Entity
        Dim lDockerID As Int32
        Dim iDockerTypeID As Int16

        Dim bFound As Boolean = False

        Dim yFailMsg() As Byte

        Dim oCmdPlayer As Player = Nothing
        If lClientIndex > -1 Then
            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) = mlClientPlayer(lClientIndex) Then
                    oCmdPlayer = goPlayers(X)
                    Exit For
                End If
            Next X
            If oCmdPlayer Is Nothing Then Return
        End If

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X = 0 To lCurUB
        '    If glEntityIdx(X) = lDockeeID AndAlso goEntity(X) Is Nothing = False AndAlso goEntity(X).ObjTypeID = iDockeeTypeID Then
        '        oDockee = goEntity(X)
        '        Exit For
        '    End If
        'Next X 
        Dim lIdx As Int32 = LookupEntity(lDockeeID, iDockeeTypeID)
        If lIdx > -1 Then oDockee = goEntity(lIdx)
        If oDockee Is Nothing OrElse oDockee.ObjectID <> lDockeeID OrElse oDockee.ObjTypeID <> iDockeeTypeID Then Return

        ReDim yForward(23)

        lDestPos = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestDock).CopyTo(yForward, lDestPos) : lDestPos += 2
        oDockee.GetGUIDAsString.CopyTo(yForward, lDestPos) : lDestPos += 6
        System.BitConverter.GetBytes(oDockee.LocX).CopyTo(yForward, lDestPos) : lDestPos += 4
        System.BitConverter.GetBytes(oDockee.LocZ).CopyTo(yForward, lDestPos) : lDestPos += 4
        System.BitConverter.GetBytes(oDockee.LocAngle).CopyTo(yForward, lDestPos) : lDestPos += 2
        oDockee.ParentEnvir.GetGUIDAsString.CopyTo(yForward, lDestPos) : lDestPos += 6

        lDestLen = lDestPos - 1

        While lPos + 6 < yData.Length '- 1
            lDockerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iDockerTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            oDocker = Nothing
            lIdx = -1
            lIdx = LookupEntity(lDockerID, iDockerTypeID)
            If lIdx > -1 Then oDocker = goEntity(lIdx)
            If oDocker Is Nothing OrElse oDocker.ObjectID <> lDockerID OrElse oDocker.ObjTypeID <> iDockerTypeID Then Continue While

            'oDocker.lTetherPointX = Int32.MinValue
            'oDocker.lTetherPointZ = Int32.MinValue

            'For X = 0 To lCurUB
            '    If glEntityIdx(X) = lDockerID AndAlso goEntity(X) Is Nothing = False AndAlso goEntity(X).ObjTypeID = iDockerTypeID Then
            '        oDocker = goEntity(X)
            '        If oDocker Is Nothing Then Exit For
            If lClientIndex <> -1 AndAlso oDocker.lOwnerID <> mlClientPlayer(lClientIndex) AndAlso (oDocker.lOwnerID <> mlAliasedAs(lClientIndex) OrElse _
                (mlAliasedRights(lClientIndex) And AliasingRights.eDockUndockUnits) = 0) AndAlso _
                (oCmdPlayer Is Nothing OrElse oDocker.Owner.lGuildID = -1 OrElse oDocker.Owner.lGuildID <> oCmdPlayer.lGuildID OrElse (oDocker.CurrentStatus And elUnitStatus.eGuildAsset) = 0) Then Continue While
            If oDockee.lOwnerID <> oDocker.lOwnerID Then Continue While

            If (oDockee.CurrentStatus And elUnitStatus.eHangarOperational) = 0 Then
                ReDim yFailMsg(22)
                System.BitConverter.GetBytes(GlobalMessageCode.eRequestDockFail).CopyTo(yFailMsg, 0) '2
                oDocker.GetGUIDAsString.CopyTo(yFailMsg, 2) '6
                oDocker.ParentEnvir.GetGUIDAsString.CopyTo(yFailMsg, 8) '6
                System.BitConverter.GetBytes(oDocker.LocX).CopyTo(yFailMsg, 14)    '4
                System.BitConverter.GetBytes(oDocker.LocZ).CopyTo(yFailMsg, 18) '4
                yFailMsg(22) = DockRejectType.eHangarInoperable

                If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eRequestDockFail, MsgMonitor.eMM_AppType.ClientConnection, yFailMsg.Length, oDocker.lOwnerID)
                oSocket.SendData(yFailMsg)
            ElseIf (oDockee.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                ReDim yFailMsg(22)
                System.BitConverter.GetBytes(GlobalMessageCode.eRequestDockFail).CopyTo(yFailMsg, 0) '2
                oDocker.GetGUIDAsString.CopyTo(yFailMsg, 2) '6
                oDocker.ParentEnvir.GetGUIDAsString.CopyTo(yFailMsg, 8) '6
                System.BitConverter.GetBytes(oDocker.LocX).CopyTo(yFailMsg, 14)    '4
                System.BitConverter.GetBytes(oDocker.LocZ).CopyTo(yFailMsg, 18) '4
                yFailMsg(22) = DockRejectType.eDockeeMoving

                If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eRequestDockFail, MsgMonitor.eMM_AppType.ClientConnection, yFailMsg.Length, oDocker.lOwnerID)
                oSocket.SendData(yFailMsg)
            Else
                'NOTE: Door Size will never be higher than hangar cap
                bFound = True

                If (oDocker.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then
                    'send a move lock reset msg
                    Dim yMsg(7) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eMoveLockViolate).CopyTo(yMsg, 0)
                    oDocker.GetGUIDAsString.CopyTo(yMsg, 2)
                    SendToPrimary(yMsg)
                    oDocker.CurrentStatus = oDocker.CurrentStatus Xor elUnitStatus.eMoveLock ' oEntity.CurrentStatus -= elUnitStatus.eMoveLock
                End If

                lDestLen += 24
                ReDim Preserve yForward(lDestLen)

                'Now, put the GUID in of the docker
                oDocker.GetGUIDAsString.CopyTo(yForward, lDestPos) : lDestPos += 6

                System.BitConverter.GetBytes(oDocker.LocX).CopyTo(yForward, lDestPos) : lDestPos += 4
                System.BitConverter.GetBytes(oDocker.LocZ).CopyTo(yForward, lDestPos) : lDestPos += 4
                System.BitConverter.GetBytes(oDocker.LocAngle).CopyTo(yForward, lDestPos) : lDestPos += 2

                oDocker.ParentEnvir.GetGUIDAsString.CopyTo(yForward, lDestPos) : lDestPos += 6

                If oDocker.lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(oDocker.lEntityDefServerIndex) <> -1 Then
                    System.BitConverter.GetBytes(goEntityDefs(oDocker.lEntityDefServerIndex).ModelID).CopyTo(yForward, lDestPos) : lDestPos += 2
                Else
                    System.BitConverter.GetBytes(CShort(-1)).CopyTo(yForward, lDestPos) : lDestPos += 2
                End If

                oDocker.bPlayerMoveRequestPending = True
                oDocker.bAIMoveRequestPending = False
            End If

            '        Exit For
            '    End If
            'Next X
        End While

        If bFound = True Then SendToPathfinding(yForward)

    End Sub

    Private Sub HandleRequestUndock(ByVal yData() As Byte, ByVal lClientIndex As Int32)
        'If glCurrentCycle - mlClientLastRequest(lClientIndex) > ml_REQUEST_THROTTLE Then
        If glCurrentCycle - muClientData(lClientIndex).mlClientLastRequest > ml_REQUEST_THROTTLE Then
            'mlClientLastRequest(lClientIndex) = glCurrentCycle
            muClientData(lClientIndex).mlClientLastRequest = glCurrentCycle
        Else : Return
        End If

        Dim oCmdPlayer As Player = Nothing
        If lClientIndex > -1 Then
            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) = mlClientPlayer(lClientIndex) Then
                    oCmdPlayer = goPlayers(X)
                    Exit For
                End If
            Next X
            If oCmdPlayer Is Nothing Then Return
        End If

        'Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        'Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        With oEntity
            Dim bHasPermission As Boolean = (lClientIndex = -1 OrElse .Owner.ObjectID = mlClientPlayer(lClientIndex) OrElse _
              (.lOwnerID = mlAliasedAs(lClientIndex) AndAlso (mlAliasedRights(lClientIndex) And AliasingRights.eMoveUnits) <> 0)) OrElse _
              (oCmdPlayer Is Nothing = False AndAlso .Owner.lGuildID > 0 AndAlso .Owner.lGuildID = oCmdPlayer.lGuildID AndAlso (.CurrentStatus And elUnitStatus.eGuildAsset) <> 0)
            If bHasPermission = False Then Return
        End With


        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        ''Now, go thru our entity list
        'For X As Int32 = 0 To lCurUB
        '    If glEntityIdx(X) = lObjID Then
        '        'goEntity(X).ObjTypeID = iObjTypeID Then
        '        Dim oEntity As Epica_Entity = goEntity(X)
        '        If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iObjTypeID Then
        '            With oEntity
        '                Dim bHasPermission As Boolean = (lClientIndex = -1 OrElse .Owner.ObjectID = mlClientPlayer(lClientIndex) OrElse _
        '                  (.lOwnerID = mlAliasedAs(lClientIndex) AndAlso (mlAliasedRights(lClientIndex) And AliasingRights.eMoveUnits) <> 0)) OrElse _
        '                  (oCmdPlayer Is Nothing = False AndAlso .Owner.lGuildID > 0 AndAlso .Owner.lGuildID = oCmdPlayer.lGuildID AndAlso (.CurrentStatus And elUnitStatus.eGuildAsset) <> 0)

        '                If bHasPermission = False Then Return
        '            End With

        '            Exit For
        '        End If
        '    End If
        'Next X


        SendToPrimary(yData)

    End Sub

    Private Sub HandleSetMiningLoc(ByVal yData() As Byte, ByVal lClientIndex As Int32)
        If lClientIndex <> -1 Then
            'If glCurrentCycle - mlClientLastRequest(lClientIndex) > ml_REQUEST_THROTTLE Then
            If glCurrentCycle - muClientData(lClientIndex).mlClientLastRequest > ml_REQUEST_THROTTLE Then
                'mlClientLastRequest(lClientIndex) = glCurrentCycle
                muClientData(lClientIndex).mlClientLastRequest = glCurrentCycle
            Else : Return
            End If
        End If

        Dim lCacheID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iCacheTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        'Dim X As Int32
        Dim Y As Int32
        Dim lPos As Int32 = 8
        Dim lDestPos As Int32
        Dim lDestLen As Int32
        Dim bSet As Boolean = False
        Dim lObjID As Int32
        Dim iObjTypeID As Int16

        Dim yForward(15) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eSetMiningLoc).CopyTo(yForward, 0)
        lDestPos = 2
        System.BitConverter.GetBytes(lCacheID).CopyTo(yForward, lDestPos) : lDestPos += 4
        System.BitConverter.GetBytes(iCacheTypeID).CopyTo(yForward, lDestPos) : lDestPos += 2
        lDestLen = 16

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))

        While lPos + 6 < yData.Length ' - 1
            lObjID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            Dim oEntity As Epica_Entity = Nothing
            Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
            If lIdx > -1 Then oEntity = goEntity(lIdx)
            If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Continue While

            'For X = 0 To lCurUB
            '    If glEntityIdx(X) = lObjID AndAlso goEntity(X) Is Nothing = False AndAlso goEntity(X).ObjTypeID = iTypeID Then
            With oEntity 'goEntity(X)

                If lClientIndex <> -1 AndAlso .lOwnerID <> mlClientPlayer(lClientIndex) AndAlso (.lOwnerID <> mlAliasedAs(lClientIndex) OrElse (mlAliasedRights(lClientIndex) And AliasingRights.eMoveUnits) = 0) Then Continue While

                If bSet = False Then
                    For Y = 0 To .ParentEnvir.lCacheUB
                        If .ParentEnvir.lCacheIdx(Y) = lCacheID AndAlso .ParentEnvir.oCache(Y).ObjTypeID = iCacheTypeID Then
                            System.BitConverter.GetBytes(.ParentEnvir.oCache(Y).LocX).CopyTo(yForward, lDestPos) : lDestPos += 4
                            System.BitConverter.GetBytes(.ParentEnvir.oCache(Y).LocZ).CopyTo(yForward, lDestPos) : lDestPos += 4
                            bSet = True
                            Exit For
                        End If
                    Next Y
                End If

                If bSet = False Then Exit Sub

                If (.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then
                    'send a move lock reset msg
                    Dim yMsg(7) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eMoveLockViolate).CopyTo(yMsg, 0)
                    .GetGUIDAsString.CopyTo(yMsg, 2)
                    SendToPrimary(yMsg)
                    .CurrentStatus = .CurrentStatus Xor elUnitStatus.eMoveLock ' oEntity.CurrentStatus -= elUnitStatus.eMoveLock
                End If

                lDestLen += 24
                ReDim Preserve yForward(lDestLen)

                'Now, the entity's GUID
                .GetGUIDAsString.CopyTo(yForward, lDestPos) : lDestPos += 6
                'now the loc
                System.BitConverter.GetBytes(.LocX).CopyTo(yForward, lDestPos) : lDestPos += 4
                System.BitConverter.GetBytes(.LocZ).CopyTo(yForward, lDestPos) : lDestPos += 4
                System.BitConverter.GetBytes(.LocAngle).CopyTo(yForward, lDestPos) : lDestPos += 2

                .ParentEnvir.GetGUIDAsString.CopyTo(yForward, lDestPos) : lDestPos += 6

                If .lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(.lEntityDefServerIndex) <> -1 Then
                    System.BitConverter.GetBytes(goEntityDefs(.lEntityDefServerIndex).ModelID).CopyTo(yForward, lDestPos) : lDestPos += 2
                Else
                    System.BitConverter.GetBytes(CShort(-1)).CopyTo(yForward, lDestPos) : lDestPos += 2
                End If

                .bPlayerMoveRequestPending = True
                .bAIMoveRequestPending = False
            End With

            '        Exit For
            '    End If
            'Next X

        End While

        If bSet = True Then SendToPathfinding(yForward)

    End Sub

    Private Sub HandleShiftClickAddProduction(ByVal yData() As Byte, ByVal lClientIndex As Int32)
        If glCurrentCycle - muClientData(lClientIndex).mlClientLastRequest > ml_REQUEST_THROTTLE Then
            muClientData(lClientIndex).mlClientLastRequest = glCurrentCycle
        Else : Return
        End If

        'need to append the loc data
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        'Get the build location...
        Dim lBuildID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iBuildTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
        Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, 14)
        Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, 18)
        Dim bMine As Boolean = False

        'Dim X As Int32

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return


        If oEntity.lOwnerID <> mlClientPlayer(lClientIndex) AndAlso (oEntity.lOwnerID <> mlAliasedAs(lClientIndex) OrElse (mlAliasedRights(lClientIndex) And AliasingRights.eAddProduction) = 0) Then
            gfrmDisplayForm.AddEventLine("Possible Cheat: Build order given to unit belonging to different player. PlayerID = " & mlClientPlayer(lClientIndex))
            Return
        End If

        'Confirm that the build location is ok for building based on terrain
        If oEntity.ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
            Dim oDef As Epica_Entity_Def = Nothing
            For Y As Int32 = 0 To glEntityDefUB
                If glEntityDefIdx(Y) = lBuildID AndAlso goEntityDefs(Y).ObjTypeID = iBuildTypeID Then
                    If goEntityDefs(Y).ProductionTypeID = ProductionType.eMining Then bMine = True
                    oDef = goEntityDefs(Y)
                    Exit For
                End If
            Next Y
            If oDef Is Nothing Then
                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eSetEntityProdFailed, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, mlClientPlayer(lClientIndex))
                moClients(lClientIndex).SendData(yData)
                Return
            End If

            Dim bNaval As Boolean = (oDef.yChassisType And ChassisType.eNavalBased) <> 0

            'Ok, is the angle of the terrain okay?
            If bMine = False AndAlso CType(oEntity.ParentEnvir.oGeoObject, Planet).TerrainGradeBuildable(lLocX, lLocZ) = False Then
                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eSetEntityProdFailed, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, mlClientPlayer(lClientIndex))
                moClients(lClientIndex).SendData(yData)
                Return
            End If

            'Is the terrain height high enough?
            If bNaval = False Then
                If CType(oEntity.ParentEnvir.oGeoObject, Planet).GetHeightAtPoint(lLocX, lLocZ, False) < CType(oEntity.ParentEnvir.oGeoObject, Planet).WaterHeight Then
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                    If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eSetEntityProdFailed, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, mlClientPlayer(lClientIndex))
                    moClients(lClientIndex).SendData(yData)
                    Return
                End If
            Else
                If CType(oEntity.ParentEnvir.oGeoObject, Planet).GetHeightAtPoint(lLocX, lLocZ, False) > CType(oEntity.ParentEnvir.oGeoObject, Planet).WaterHeight Then
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                    If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eSetEntityProdFailed, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, mlClientPlayer(lClientIndex))
                    moClients(lClientIndex).SendData(yData)
                    Return
                End If
            End If

            Dim lHalfVal As Int32 = oDef.lModelSizeXZ \ 2
            Dim lMinHt As Int32 = Int32.MaxValue
            Dim lMaxHt As Int32 = Int32.MinValue
            With CType(oEntity.ParentEnvir.oGeoObject, Planet)
                Dim LocX As Int32 = CInt(lLocX)
                Dim LocZ As Int32 = CInt(lLocZ)
                Dim lTemp As Int32 = CInt(.GetHeightAtPoint(LocX - lHalfVal, LocZ - lHalfVal, False))
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX + lHalfVal, LocZ - lHalfVal, False))
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX - lHalfVal, LocZ + lHalfVal, False))
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX + lHalfVal, LocZ + lHalfVal, False))
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX, LocZ, False))
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
            End With
            If lMaxHt - lMinHt > 200 AndAlso oDef.ProductionTypeID <> ProductionType.eMining Then
                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eSetEntityProdFailed, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, mlClientPlayer(lClientIndex))
                moClients(lClientIndex).SendData(yData)
                Return
            End If
        End If

        moPrimary.SendData(yData)

    End Sub

    Private Sub HandleSetEntityProduction(ByVal yData() As Byte, ByVal lClientIndex As Int32)
        'If glCurrentCycle - mlClientLastRequest(lClientIndex) > ml_REQUEST_THROTTLE Then
        If glCurrentCycle - muClientData(lClientIndex).mlClientLastRequest > ml_REQUEST_THROTTLE Then
            'mlClientLastRequest(lClientIndex) = glCurrentCycle
            muClientData(lClientIndex).mlClientLastRequest = glCurrentCycle
        Else : Return
        End If

        'need to append the loc data
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        'Get the build location...
        Dim lBuildID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iBuildTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
        Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, 14)
        Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, 18)
        Dim bMine As Boolean = False

        'Dim X As Int32

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        'For X = 0 To lCurUB
        '    If glEntityIdx(X) = lObjID AndAlso goEntity(X) Is Nothing = False AndAlso goEntity(X).ObjTypeID = iObjTypeID Then
        '        Dim oEntity As Epica_Entity = goEntity(X)
        '        If oEntity Is Nothing Then Exit For

        If oEntity.lOwnerID <> mlClientPlayer(lClientIndex) AndAlso (oEntity.lOwnerID <> mlAliasedAs(lClientIndex) OrElse (mlAliasedRights(lClientIndex) And AliasingRights.eAddProduction) = 0) Then
            gfrmDisplayForm.AddEventLine("Possible Cheat: Build order given to unit belonging to different player. PlayerID = " & mlClientPlayer(lClientIndex))
            Return
        End If

        'Confirm that the build location is ok for building based on terrain
        If oEntity.ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
            Dim oDef As Epica_Entity_Def = Nothing
            For Y As Int32 = 0 To glEntityDefUB
                If glEntityDefIdx(Y) = lBuildID AndAlso goEntityDefs(Y).ObjTypeID = iBuildTypeID Then
                    If goEntityDefs(Y).ProductionTypeID = ProductionType.eMining Then bMine = True
                    oDef = goEntityDefs(Y)
                    Exit For
                End If
            Next Y
            If oDef Is Nothing Then
                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eSetEntityProdFailed, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, mlClientPlayer(lClientIndex))
                moClients(lClientIndex).SendData(yData)
                Return
            End If

            Dim bNaval As Boolean = (oDef.yChassisType And ChassisType.eNavalBased) <> 0

            'Ok, is the angle of the terrain okay?
            If bMine = False AndAlso CType(oEntity.ParentEnvir.oGeoObject, Planet).TerrainGradeBuildable(lLocX, lLocZ) = False Then
                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eSetEntityProdFailed, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, mlClientPlayer(lClientIndex))
                moClients(lClientIndex).SendData(yData)
                Return
            End If

            'Is the terrain height high enough?
            If bNaval = False Then
                If CType(oEntity.ParentEnvir.oGeoObject, Planet).GetHeightAtPoint(lLocX, lLocZ, False) < CType(oEntity.ParentEnvir.oGeoObject, Planet).WaterHeight Then
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                    If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eSetEntityProdFailed, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, mlClientPlayer(lClientIndex))
                    moClients(lClientIndex).SendData(yData)
                    Return
                End If
            Else
                If CType(oEntity.ParentEnvir.oGeoObject, Planet).GetHeightAtPoint(lLocX, lLocZ, False) > CType(oEntity.ParentEnvir.oGeoObject, Planet).WaterHeight Then
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                    If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eSetEntityProdFailed, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, mlClientPlayer(lClientIndex))
                    moClients(lClientIndex).SendData(yData)
                    Return
                End If
            End If

            Dim lHalfVal As Int32 = oDef.lModelSizeXZ \ 2
            Dim lMinHt As Int32 = Int32.MaxValue
            Dim lMaxHt As Int32 = Int32.MinValue
            With CType(oEntity.ParentEnvir.oGeoObject, Planet)
                Dim LocX As Int32 = CInt(lLocX)
                Dim LocZ As Int32 = CInt(lLocZ)
                Dim lTemp As Int32 = CInt(.GetHeightAtPoint(LocX - lHalfVal, LocZ - lHalfVal, False))
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX + lHalfVal, LocZ - lHalfVal, False))
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX - lHalfVal, LocZ + lHalfVal, False))
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX + lHalfVal, LocZ + lHalfVal, False))
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX, LocZ, False))
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
            End With
            If lMaxHt - lMinHt > 200 AndAlso oDef.ProductionTypeID <> ProductionType.eMining Then
                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yData, 0)
                If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(GlobalMessageCode.eSetEntityProdFailed, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, mlClientPlayer(lClientIndex))
                moClients(lClientIndex).SendData(yData)
                Return
            End If
        End If

        If (oEntity.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then
            'send a move lock reset msg
            Dim yMsg(7) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eMoveLockViolate).CopyTo(yMsg, 0)
            oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
            SendToPrimary(yMsg)
            oEntity.CurrentStatus = oEntity.CurrentStatus Xor elUnitStatus.eMoveLock ' oEntity.CurrentStatus -= elUnitStatus.eMoveLock
        End If

        ReDim Preserve yData(39)
        System.BitConverter.GetBytes(oEntity.LocX).CopyTo(yData, 24)
        System.BitConverter.GetBytes(oEntity.LocZ).CopyTo(yData, 28)

        oEntity.ParentEnvir.GetGUIDAsString.CopyTo(yData, 32)

        If oEntity.lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(oEntity.lEntityDefServerIndex) <> -1 Then
            System.BitConverter.GetBytes(goEntityDefs(oEntity.lEntityDefServerIndex).ModelID).CopyTo(yData, 38)
        Else : System.BitConverter.GetBytes(CShort(-1)).CopyTo(yData, 38)
        End If

        SendToPathfinding(yData)

        Return
        '    End If
        'Next X


    End Sub

    'comes from Primary Server...
    Private Sub HandleSetEntityProdSucceed(ByVal yData() As Byte)
        'we need to get the entity referred to by this message and set a special flag?
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        'Dim X As Int32

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        oEntity.CurrentStatus = oEntity.CurrentStatus Or elUnitStatus.eMoveLock

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X = 0 To lCurUB
        '    If glEntityIdx(X) = lObjID AndAlso goEntity(X) Is Nothing = False AndAlso goEntity(X).ObjTypeID = iObjTypeID Then
        '        goEntity(X).CurrentStatus = goEntity(X).CurrentStatus Or elUnitStatus.eMoveLock
        '        Exit For
        '    End If
        'Next X

    End Sub

    'Comes from Primary Server...
    Private Sub HandleSetEntityName(ByRef yData() As Byte)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        ReDim oEntity.UnitName(19)
        Array.Copy(yData, 8, oEntity.UnitName, 0, 20)

        'Try
        '    If iTypeID = ObjectType.eUnit OrElse iTypeID = ObjectType.eFacility Then
        '        Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        '        For X As Int32 = 0 To lCurUB
        '            If glEntityIdx(X) = lObjID AndAlso goEntity(X).ObjTypeID = iTypeID Then
        '                ReDim goEntity(X).UnitName(19)
        '                Array.Copy(yData, 8, goEntity(X).UnitName, 0, 20)
        '                Exit For
        '            End If
        '        Next X
        '    End If
        'Catch
        'End Try
    End Sub

    Private Sub HandleSetEntityStatus(ByVal yData() As Byte)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lStatus As Int32 = System.BitConverter.ToInt32(yData, 8)
        'Dim X As Int32

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        'ok, check if the status is set already
        If lStatus < 0 Then
            'ok, trying to remove a status
            If (oEntity.CurrentStatus And Math.Abs(lStatus)) <> 0 Then
                oEntity.CurrentStatus = oEntity.CurrentStatus Xor Math.Abs(lStatus) 'goEntity(X).CurrentStatus += lStatus		'because its a negative
            End If
        Else
            'ok, trying to add a status
            If (oEntity.CurrentStatus And lStatus) = 0 Then
                oEntity.CurrentStatus = oEntity.CurrentStatus Or lStatus
            End If
        End If

        'Now, check the status changed
        If lStatus = -(elUnitStatus.eMoveLock) Then
            'Ok, check if the entity's currentstatus is not moving
            If (oEntity.CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
                'Ok, then let's move it away from where it was building...
                'SendAIMoveRequestToPathfinding(goEntity(X), goEntity(X).LocX, goEntity(X).LocZ - 500, goEntity(X).LocAngle)
            End If
        ElseIf Math.Abs(lStatus) = elUnitStatus.eFacilityPowered Then
            'regardless of whether the entity is being turned on or off, we need to force the aggression test
            With oEntity
                If (.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
                    .CurrentStatus = .CurrentStatus Xor elUnitStatus.eUnitEngaged ' .CurrentStatus -= elUnitStatus.eUnitEngaged
                    RemoveEntityInCombat(.ObjectID, .ObjTypeID)
                End If
                If (.CurrentStatus And elUnitStatus.eSide1HasTarget) <> 0 Then .CurrentStatus = .CurrentStatus Xor elUnitStatus.eSide1HasTarget
                If (.CurrentStatus And elUnitStatus.eSide2HasTarget) <> 0 Then .CurrentStatus = .CurrentStatus Xor elUnitStatus.eSide2HasTarget '.CurrentStatus -= elUnitStatus.eSide2HasTarget
                If (.CurrentStatus And elUnitStatus.eSide3HasTarget) <> 0 Then .CurrentStatus = .CurrentStatus Xor elUnitStatus.eSide3HasTarget '.CurrentStatus -= elUnitStatus.eSide3HasTarget
                If (.CurrentStatus And elUnitStatus.eSide4HasTarget) <> 0 Then .CurrentStatus = .CurrentStatus Xor elUnitStatus.eSide4HasTarget '.CurrentStatus -= elUnitStatus.eSide4HasTarget
                'Not sure about this...
                '.lPrimaryTargetServerIdx = -1
                'For Y As Int32 = 0 To .lTargetsServerIdx.GetUpperBound(0)
                '    .lTargetsServerIdx(Y) = -1
                'Next
                .bForceAggressionTest = True
                SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)
                'RemoveEntityCombat(.ServerIndex, .ObjectID)
            End With
        End If
    End Sub

    Public Sub SendHangarCargoDestroyed(ByVal oEntity As Epica_Entity, ByVal lStatusID As Int32)
        Dim yData(19) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eHangarCargoBayDestroyed).CopyTo(yData, 0)
        oEntity.GetGUIDAsString.CopyTo(yData, 2)
        System.BitConverter.GetBytes(lStatusID).CopyTo(yData, 8)
        System.BitConverter.GetBytes(oEntity.LocX).CopyTo(yData, 12)
        System.BitConverter.GetBytes(oEntity.LocZ).CopyTo(yData, 16)
        SendToPrimary(yData)
    End Sub

    Private Sub HandleRemoveObject(ByVal yData() As Byte)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim yRemoveType As Byte = yData(8)

        Select Case iObjTypeID
            Case ObjectType.eUnit, ObjectType.eFacility
                Dim oEntity As Epica_Entity = Nothing
                Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
                If lIdx > -1 Then oEntity = goEntity(lIdx)
                If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return
                If oEntity.ParentEnvir Is Nothing = False Then oEntity.ParentEnvir.RemoveEntity(lIdx, CType(yRemoveType, RemovalType), True, False, -1)
            Case ObjectType.eMineralCache, ObjectType.eComponentCache
                Dim bFound As Boolean = False
                For X As Int32 = 0 To glEnvirUB
                    For Y As Int32 = 0 To goEnvirs(X).lCacheUB
                        If goEnvirs(X).lCacheIdx(Y) = lObjID AndAlso goEnvirs(X).oCache(Y).ObjTypeID = iObjTypeID Then
                            goEnvirs(X).lCacheIdx(Y) = -1
                            If goEnvirs(X).lPlayersInEnvirCnt > 0 Then goMsgSys.BroadcastToEnvironmentClients(yData, goEnvirs(X))
                            bFound = True
                            Exit For
                        End If
                    Next Y
                    If bFound = True Then Exit For
                Next X
        End Select
    End Sub

    Private Sub HandleGroupMoveMessage(ByVal yData() As Byte, ByVal lClientIndex As Int32)
        'Msg Structure...
        'MsgCode - 2, DestX - 4, DestZ - 4, DestA - 2, DestID - 4, DestTypeID - 2, GUID List...
        ' Honestly, we only care about the GUID List...

        If lClientIndex <> -1 Then
            'If glCurrentCycle - mlClientLastRequest(lClientIndex) > ml_REQUEST_THROTTLE Then
            If glCurrentCycle - muClientData(lClientIndex).mlClientLastRequest > ml_REQUEST_THROTTLE Then
                'mlClientLastRequest(lClientIndex) = glCurrentCycle
                muClientData(lClientIndex).mlClientLastRequest = glCurrentCycle
            Else : Return
            End If
        End If

        Dim yForward(17) As Byte
        Dim lPos As Int32 = 0
        Dim lObjID As Int32
        Dim iObjTypeID As Int16
        Dim bFound As Boolean = False
        Dim X As Int32
        Dim lDestPos As Int32
        Dim lDestLen As Int32
        Dim bEntityIncluded As Boolean = False

        Dim bCanEnterEnvir As Boolean = False

        Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, 16)

        Array.Copy(yData, 0, yForward, 0, 18)   'copy the code, dest coord and dest envir guid

        'Now, go through our GUID list...
        lPos = 18
        lDestPos = 18
        lDestLen = 17


        Dim oCmdPlayer As Player = Nothing
        If lClientIndex <> -1 Then
            For X = 0 To glPlayerUB
                If glPlayerIdx(X) = mlClientPlayer(lClientIndex) Then
                    oCmdPlayer = goPlayers(X)
                    Exit For
                End If
            Next X
            If oCmdPlayer Is Nothing Then Return
        End If

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        While lPos + 6 < yData.Length '- 1
            bFound = False
            lObjID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            Dim oEntity As Epica_Entity = Nothing
            Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
            If lIdx > -1 Then oEntity = goEntity(lIdx)
            If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Continue While

            ''Now, go thru our entity list
            'For X = 0 To lCurUB
            '    If glEntityIdx(X) = lObjID Then
            '        'goEntity(X).ObjTypeID = iObjTypeID Then
            '        Dim oEntity As Epica_Entity = goEntity(X)
            '        If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iObjTypeID Then

            With oEntity
                bCanEnterEnvir = False
                If .lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(.lEntityDefServerIndex) <> -1 Then
                    If iDestTypeID = ObjectType.ePlanet Then
                        If (goEntityDefs(.lEntityDefServerIndex).yChassisType And ChassisType.eAtmospheric) <> 0 OrElse _
                           (goEntityDefs(.lEntityDefServerIndex).yChassisType And (ChassisType.eGroundBased Or ChassisType.eNavalBased)) <> 0 Then
                            bCanEnterEnvir = True
                        End If
                    ElseIf iDestTypeID = ObjectType.eSolarSystem Then
                        If (goEntityDefs(.lEntityDefServerIndex).yChassisType And ChassisType.eSpaceBased) <> 0 Then
                            bCanEnterEnvir = True
                        End If
                    End If
                End If

                If bCanEnterEnvir = True AndAlso .Owner Is Nothing = False Then

                    Dim bHasPermission As Boolean = (lClientIndex = -1 OrElse .Owner.ObjectID = mlClientPlayer(lClientIndex) OrElse _
                      (.lOwnerID = mlAliasedAs(lClientIndex) AndAlso (mlAliasedRights(lClientIndex) And AliasingRights.eMoveUnits) <> 0)) OrElse _
                      (oCmdPlayer Is Nothing = False AndAlso .Owner.lGuildID > 0 AndAlso .Owner.lGuildID = oCmdPlayer.lGuildID AndAlso (.CurrentStatus And elUnitStatus.eGuildAsset) <> 0)

                    If bHasPermission = True AndAlso ((.CurrentStatus And elUnitStatus.eEngineOperational) <> 0) AndAlso .MaxSpeed > 0 AndAlso .Acceleration > 0 Then

                        If (.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then
                            'send a move lock reset msg
                            Dim yMsg(7) As Byte
                            System.BitConverter.GetBytes(GlobalMessageCode.eMoveLockViolate).CopyTo(yMsg, 0)
                            .GetGUIDAsString.CopyTo(yMsg, 2)
                            SendToPrimary(yMsg)
                            .CurrentStatus = oEntity.CurrentStatus Xor elUnitStatus.eMoveLock ' oEntity.CurrentStatus -= elUnitStatus.eMoveLock
                        End If

                        bFound = True
                        bEntityIncluded = True

                        lDestLen += 24
                        ReDim Preserve yForward(lDestLen)

                        'Now, put our GUID in the forward message
                        .GetGUIDAsString.CopyTo(yForward, lDestPos) : lDestPos += 6

                        'Now, put our current location there...
                        System.BitConverter.GetBytes(.LocX).CopyTo(yForward, lDestPos) : lDestPos += 4
                        System.BitConverter.GetBytes(.LocZ).CopyTo(yForward, lDestPos) : lDestPos += 4
                        System.BitConverter.GetBytes(.LocAngle).CopyTo(yForward, lDestPos) : lDestPos += 2

                        'And our current parent envir GUID
                        .ParentEnvir.GetGUIDAsString.CopyTo(yForward, lDestPos) : lDestPos += 6

                        If .lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(.lEntityDefServerIndex) <> -1 Then
                            System.BitConverter.GetBytes(goEntityDefs(.lEntityDefServerIndex).ModelID).CopyTo(yForward, lDestPos) : lDestPos += 2
                        Else
                            System.BitConverter.GetBytes(CShort(-1)).CopyTo(yForward, lDestPos) : lDestPos += 2
                        End If

                        '.lTetherPointX = Int32.MinValue
                        '.lTetherPointZ = Int32.MinValue

                        .bPlayerMoveRequestPending = True
                        .bAIMoveRequestPending = False
                    End If
                End If
            End With

            '            Exit For
            '        End If
            '    End If
            'Next X

            If bFound = False Then
                'ok... for now, do nothing... it simply won't be included...
            End If
        End While

        If yForward Is Nothing = False AndAlso yForward.Length > 0 AndAlso bEntityIncluded Then SendToPathfinding(yForward)

        If lClientIndex <> -1 Then
            'Ok, send to the primary server to notify it of any possible cancel mining routes
            SendToPrimary(yData)
        End If

    End Sub

    Private Sub HandleUndockCommand(ByVal yData() As Byte)
        'ok... a lot like the Add Object message
        Dim X As Int32
        Dim Y As Int32
        Dim lIdx As Int32 = -1
        Dim lFirstIndex As Int32 = -1
        Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yData, 0)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        'Ok... this parent object is the object we are undocking from...
        Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)

        Dim lTemp As Int32

        Dim lX As Int32
        Dim lZ As Int32

        Dim lLaunchFromID As Int32 = Int32.MinValue
        Dim iLaunchFromTypeID As Int16 = Int16.MinValue

        'Get the object we are undocking from's Parent
        Dim lPIdx As Int32 = LookupEntity(lParentID, iParentTypeID)
        Dim oParent As Epica_Entity = Nothing
        If lPIdx > -1 Then oParent = goEntity(lPIdx)
        If oParent Is Nothing = False AndAlso oParent.ObjectID = lParentID AndAlso oParent.ObjTypeID = iParentTypeID Then
            If oParent.ParentEnvir Is Nothing = False Then
                lParentID = oParent.ParentEnvir.ObjectID
                iParentTypeID = oParent.ParentEnvir.ObjTypeID
            End If
            lX = oParent.LocX
            lZ = oParent.LocZ
            If oParent.bInAILaunchAll = True Then
                lLaunchFromID = oParent.ObjectID
                iLaunchFromTypeID = oParent.ObjTypeID
            End If
        End If

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X = 0 To lCurUB
        '    If glEntityIdx(X) = lParentID AndAlso goEntity(X) Is Nothing = False AndAlso goEntity(X).ObjTypeID = iParentTypeID Then
        '        lParentID = goEntity(X).ParentEnvir.ObjectID
        '        iParentTypeID = goEntity(X).ParentEnvir.ObjTypeID
        '        lX = goEntity(X).LocX
        '        lZ = goEntity(X).LocZ
        '        If goEntity(X).bInAILaunchAll = True Then
        '            lLaunchFromID = goEntity(X).ObjectID
        '            iLaunchFromTypeID = goEntity(X).ObjTypeID
        '        End If
        '        Exit For
        '    End If
        'Next X

        lIdx = AddEntity(lObjID, iObjTypeID)
        If lIdx < 0 OrElse lIdx > glEntityUB Then
            gfrmDisplayForm.AddEventLine("Invalid Index returned from AddEntity! Object not added: " & lObjID & ", " & iObjTypeID)
            Return
        End If

        'set our index
        glEntityIdx(lIdx) = lObjID

        'Now, parse our values... first check the parent object
        With goEntity(lIdx)
            .ObjectID = lObjID
            .ObjTypeID = iObjTypeID

            .ServerIndex = lIdx

            If .ParentEnvir Is Nothing = False Then
                If .lGridIndex <> -1 AndAlso .lGridEntityIdx <> -1 Then
                    If .ParentEnvir.oGrid(.lGridIndex).lEntityUB >= .lGridEntityIdx Then
                        If .ParentEnvir.oGrid(.lGridIndex).lEntities(.lGridEntityIdx) = .ServerIndex Then
                            .ParentEnvir.oGrid(.lGridIndex).RemoveEntity(.lGridEntityIdx)
                        End If
                    End If
                End If
                If .ParentEnvir.ObjectID <> lParentID OrElse .ParentEnvir.ObjTypeID <> iParentTypeID Then
                    'ok, we are not good for the environment...
                    If iParentTypeID = ObjectType.ePlanet AndAlso lParentID >= 500000000 Then
                        'instance
                        Dim lTmpCurUB As Int32 = -1
                        If glInstanceIdx Is Nothing = False Then lTmpCurUB = Math.Min(glInstanceUB, glInstanceIdx.GetUpperBound(0))
                        For X = 0 To lTmpCurUB
                            If glInstanceIdx(X) = lParentID Then
                                'ok, found the environment...
                                .ParentEnvir = goInstances(X)
                                Exit For
                            End If
                        Next X
                    Else
                        For X = 0 To glEnvirUB
                            If goEnvirs(X).ObjectID = lParentID AndAlso goEnvirs(X).ObjTypeID = iParentTypeID Then
                                'ok, found the environment...
                                .ParentEnvir = goEnvirs(X)
                                Exit For
                            End If
                        Next X
                    End If
                End If
            Else
                If iParentTypeID = ObjectType.ePlanet AndAlso lParentID >= 500000000 Then
                    Dim lTmpCurUB As Int32 = -1
                    If glInstanceIdx Is Nothing = False Then lTmpCurUB = Math.Min(glInstanceUB, glInstanceIdx.GetUpperBound(0))
                    For X = 0 To lTmpCurUB
                        If glInstanceIdx(X) = lParentID Then
                            'ok, found the environment...
                            .ParentEnvir = goInstances(X)
                            Exit For
                        End If
                    Next X
                Else
                    For X = 0 To glEnvirUB
                        If goEnvirs(X).ObjectID = lParentID AndAlso goEnvirs(X).ObjTypeID = iParentTypeID Then
                            'ok, found the environment...
                            .ParentEnvir = goEnvirs(X)
                            Exit For
                        End If
                    Next X
                End If
            End If

            'If .ParentEnvir Is Nothing = False Then
            '	If .ParentEnvir.ObjectID <> lParentID OrElse .ParentEnvir.ObjTypeID <> iParentTypeID Then
            '		'ok, we are not good for the environment...
            '		For X = 0 To glEnvirUB
            '			If goEnvirs(X).ObjectID = lParentID AndAlso goEnvirs(X).ObjTypeID = iParentTypeID Then
            '				'ok, found the environment...
            '				.ParentEnvir = goEnvirs(X)
            '				Exit For
            '			End If
            '		Next X
            '	End If
            'Else
            '	For X = 0 To glEnvirUB
            '		If goEnvirs(X).ObjectID = lParentID AndAlso goEnvirs(X).ObjTypeID = iParentTypeID Then
            '			'ok, found the environment...
            '			.ParentEnvir = goEnvirs(X)
            '			Exit For
            '		End If
            '	Next X
            'End If

            'Ok, now... get the UnitDefID
            lTemp = System.BitConverter.ToInt32(yData, 14)
            .lEntityDefServerIndex = -1
            For X = 0 To glEntityDefUB
                If glEntityDefIdx(X) = lTemp Then
                    If (iObjTypeID = ObjectType.eUnit AndAlso goEntityDefs(X).ObjTypeID = ObjectType.eUnitDef) OrElse _
                       (iObjTypeID = ObjectType.eFacility AndAlso goEntityDefs(X).ObjTypeID = ObjectType.eFacilityDef) Then
                        .lEntityDefServerIndex = X
                        .lModelRangeOffset = goEntityDefs(X).lModelRangeOffset
                        .Acceleration = goEntityDefs(X).BaseAcceleration
                        .Maneuver = goEntityDefs(X).BaseManeuver
                        .MaxSpeed = goEntityDefs(X).BaseMaxSpeed
                        .TurnAmount = goEntityDefs(X).BaseTurnAmount
                        .TurnAmountTimes100 = goEntityDefs(X).BaseTurnAmountTimes100

                        ReDim .lWpnAmmoCnt(goEntityDefs(X).WeaponDefUB)
                        For Y = 0 To .lWpnAmmoCnt.Length - 1
                            .lWpnAmmoCnt(Y) = -1
                        Next Y
                        Exit For
                    End If
                End If
            Next X
            If .lEntityDefServerIndex = -1 Then
                X = 1
            End If

            Array.Copy(yData, 18, .UnitName, 0, 20)
            lTemp = System.BitConverter.ToInt32(yData, 38)
            For X = 0 To glPlayerUB
                If glPlayerIdx(X) = lTemp Then
                    .Owner = goPlayers(X)
                    '.sPlayerRelString = "PR_" & CStr(lTemp)
                    .lOwnerID = lTemp
                    Exit For
                End If
            Next X

            If .Owner Is Nothing Then
                .Owner = Nothing
                'TODO: It is okay to have no owner, but in that case, it is a free-for-all unit...
                ' now, I believe we should introduce dummy factions or "owners" that are pirates or something
            End If

            .InitializeWeaponCycles()   'we are suppose to call this when the entity def is assigned

            .LocX = lX 'System.BitConverter.ToInt32(yData, 42)
            .LocZ = lZ 'System.BitConverter.ToInt32(yData, 46)
            '.lTetherPointX = Int32.MinValue '.LocX
            '.lTetherPointZ = Int32.MinValue ' .LocZ

            .lLaunchedFromID = lLaunchFromID
            .iLaunchedFromTypeID = iLaunchFromTypeID

#If EXTENSIVELOGGING = 1 Then
            gfrmDisplayForm.AddEventLine("Add Object Received(" & lIdx & "): " & lObjID & ", " & iObjTypeID & " to " & lParentID & ", " & iParentTypeID & " at " & .LocX & ", " & .LocZ)
#End If

            .LocAngle = System.BitConverter.ToInt16(yData, 50)
            .StructuralHP = System.BitConverter.ToInt32(yData, 52)
            .Fuel_Cap = System.BitConverter.ToInt32(yData, 56)
            .lLastShieldRechargeCycle = glCurrentCycle
            .ShieldHP = System.BitConverter.ToInt32(yData, 60)
            .Exp_Level = yData(64)
            lTemp = System.BitConverter.ToInt32(yData, 65)      'UnitGroupID
            .CurrentStatus = System.BitConverter.ToInt32(yData, 69)

            .ArmorHP(0) = System.BitConverter.ToInt32(yData, 73)
            .ArmorHP(1) = System.BitConverter.ToInt32(yData, 77)
            .ArmorHP(2) = System.BitConverter.ToInt32(yData, 81)
            .ArmorHP(3) = System.BitConverter.ToInt32(yData, 85)

            Dim lPos As Int32 = 89
            'If iObjTypeID = ObjectType.eFacility Then
            '    'got some extra values that we currently don't use... but we'll need to
            '    'the extra values are...
            '    'CurrentWorkers (4 bytes)
            '    'ParentColonyID (4 bytes)
            '    'Width (2 bytes)
            '    'Height (2 bytes)
            '    lPos += 12
            'End If
            .yProductionType = yData(lPos) : lPos += 1
            .iCombatTactics = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .iTargetingTactics = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            Dim lTemp2 As Int32

            'Now, get our cnt of Ammo Amts
            lTemp = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            'TODO: It is assumed that the message length allows for the number of items here...
            For X = 0 To lTemp - 1
                lTemp2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                For Y = 0 To goEntityDefs(.lEntityDefServerIndex).WeaponDefUB
                    If goEntityDefs(.lEntityDefServerIndex).WeaponDefs(Y).lEntityWpnDefID = lTemp2 Then
                        .lWpnAmmoCnt(Y) = System.BitConverter.ToInt32(yData, lPos)
                        Exit For
                    End If
                Next Y
                lPos += 4       'regardless of whether we find it or not, increment by 4
            Next X

            Dim lFinalParentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iFinalParentTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Dim lBackupLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lBackupLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            If .ParentEnvir Is Nothing Then

                gfrmDisplayForm.AddEventLine("HandleUndockCommand: Had to rely on FinalGUID for " & lObjID & ", " & iObjTypeID & ". ParentGUID: " & lParentID & ", " & iParentTypeID & ". FinalGUID: " & lFinalParentID & ", " & iFinalParentTypeID)

                For X = 0 To glEnvirUB
                    If goEnvirs(X).ObjectID = lFinalParentID AndAlso goEnvirs(X).ObjTypeID = iFinalParentTypeID Then
                        .ParentEnvir = goEnvirs(X)
                        Exit For
                    End If
                Next X
                .LocX = lBackupLocX
                .LocZ = lBackupLocZ
            End If

            .MoveLocX = .LocX
            .MoveLocZ = .LocZ
            .bForceAggressionTest = True
            .lLastForceAggressionTest = glCurrentCycle - gl_FORCE_AGGRESSION_THRESHOLD
            .DestX = .LocX
            .DestZ = .LocZ
            .DestAngle = .LocAngle
            .LastCycleMoved = glCurrentCycle
            .SetExpLevelMods()

            AddLookupEntity(.ObjectID, .ObjTypeID, lIdx)

            'finally, add the object to its parent environment
            If .ParentEnvir Is Nothing = False Then
                .ParentEnvir.AddEntity(goEntity(lIdx))
            Else
                gfrmDisplayForm.AddEventLine("HandleUndockCommand.Parent Envir is nothing. " & .ObjectID & ", " & .ObjTypeID)
            End If

            SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)

        End With

        'Now, that we have our object... forward the command to connected players
        If goEntity(lIdx).ParentEnvir.lPlayersInEnvirCnt > 0 Then
            BroadcastToEnvironmentClients_Ex(iMsgCode, goEntity(lIdx).GetObjectAsSmallString, goEntity(lIdx).ParentEnvir)
        End If

        Dim yResp(13) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eUndockCommandFinished).CopyTo(yResp, 0)
        Array.Copy(yData, 2, yResp, 2, 12)
        'System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
        'System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
        moPrimary.SendData(yResp)
    End Sub

    Private Sub HandleDockCommand(ByVal yData() As Byte)
        Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)

        'Dim X As Int32

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lEntityID, iEntityTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lEntityID OrElse oEntity.ObjTypeID <> iEntityTypeID Then Return

        lIdx = -1
        Dim oParent As Epica_Entity = Nothing
        lIdx = LookupEntity(lParentID, iParentTypeID)
        If lIdx > -1 Then oParent = goEntity(lIdx)
        If oParent Is Nothing OrElse oParent.ObjectID <> lParentID OrElse oParent.ObjTypeID <> iParentTypeID Then Return

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X = 0 To lCurUB
        '    If glEntityIdx(X) = lEntityID AndAlso goEntity(X).ObjTypeID = iEntityTypeID Then
        '        oEntity = goEntity(X)
        '    ElseIf glEntityIdx(X) = lParentID AndAlso goEntity(X).ObjTypeID = iParentTypeID Then
        '        oParent = goEntity(X)
        '    End If
        '    If oEntity Is Nothing = False AndAlso oParent Is Nothing = False Then Exit For
        'Next X

        If oEntity Is Nothing = False AndAlso oParent Is Nothing = False Then
            '1) Send Save to Primary
            Try
                Dim yMsgVal() As Byte = oEntity.GetObjectUpdateMsg(GlobalMessageCode.eRemoveObject, RemovalType.eDocking)
                moPrimary.SendData(yMsgVal)
            Catch ex As Exception
                gfrmDisplayForm.AddEventLine("HandleDockCommand.SendObjectUpdateMsg Failed for " & oEntity.ObjectID & ", " & oEntity.ObjTypeID & " because: " & ex.Message)
            End Try

            '2) Remove entity
            If oEntity.ParentEnvir Is Nothing = False Then
                Try
                    oEntity.ParentEnvir.RemoveEntity(oEntity.ServerIndex, RemovalType.eDocking, True, False, -1)
                Catch ex As Exception
                    gfrmDisplayForm.AddEventLine("HandleDockCommand.RemoveEntity failed for " & oEntity.ObjectID & ", " & oEntity.ObjTypeID & " because: " & ex.Message)
                End Try
            Else
                gfrmDisplayForm.AddEventLine("HandleDockCommand.ParentEnvir for " & oEntity.ObjectID & ", " & oEntity.ObjTypeID & " is nothing")
            End If

            '3) send msg back to primary
            moPrimary.SendData(yData)

        Else
            gfrmDisplayForm.AddEventLine("HandleDockCommand: Entity or Parent is nothing. " & lEntityID & ", " & iEntityTypeID & " to " & lParentID & ", " & iParentTypeID)
        End If

    End Sub

    Private Sub HandleRequestBombard(ByVal yData() As Byte, ByVal lSocketIndex As Int32)
        'eRequestOrbitalBombard, lPlayerID, PlanetID, LocX, LocZ, BombType
        If lSocketIndex <> -1 Then
            'If glCurrentCycle - mlClientLastRequest(lSocketIndex) > ml_REQUEST_THROTTLE Then
            If glCurrentCycle - muClientData(lSocketIndex).mlClientLastRequest > ml_REQUEST_THROTTLE Then
                'mlClientLastRequest(lSocketIndex) = glCurrentCycle
                muClientData(lSocketIndex).mlClientLastRequest = glCurrentCycle
            Else : Return
            End If
        End If

        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lPlanetID As Int32 = System.BitConverter.ToInt32(yData, 6)
        Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, 10)
        Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, 14)
        Dim yType As Byte = yData(18)

        'Verify lPlayerID = mlClientPlayers() here
        If lPlayerID <> mlClientPlayer(lSocketIndex) AndAlso (lPlayerID <> mlAliasedAs(lSocketIndex) OrElse (mlAliasedRights(lSocketIndex) And AliasingRights.eChangeBehavior) = 0) Then
            gfrmDisplayForm.AddEventLine("Possible Cheat: Non-Owner, or non alias rights, setting HandleRequestBombard.  PlayerID = " & mlClientPlayer(lSocketIndex))
            Return
        End If

        AddBombRequest(lPlanetID, yType, lPlayerID, lLocX, lLocZ)
    End Sub

    Private Sub HandleRequestStopBombard(ByVal yData() As Byte, ByVal lSocketIndex As Int32)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lPlanetID As Int32 = System.BitConverter.ToInt32(yData, 6)

        Dim X As Int32

        'Verify lPlayerID = mlClientPlayers() here
        If lPlayerID <> mlClientPlayer(lSocketIndex) AndAlso (lPlayerID <> mlAliasedAs(lSocketIndex) OrElse (mlAliasedRights(lSocketIndex) And AliasingRights.eChangeBehavior) = 0) Then
            gfrmDisplayForm.AddEventLine("Possible Cheat: Non-Owner, or non alias rights, setting HandleRequestBombard.  PlayerID = " & mlClientPlayer(lSocketIndex))
            Return
        End If

        For X = 0 To glBombRequestUB
            If gyBombRequestUsed(X) > 0 Then
                If goBombRequests(X).lPlayerID = lPlayerID AndAlso goBombRequests(X).PlanetID = lPlanetID Then
                    gyBombRequestUsed(X) = 0
                    Exit For
                End If
            End If
        Next X

    End Sub

    Private Sub HandleSetEntityAI(ByVal yData() As Byte, ByVal lIndex As Int32)
        'If lIndex <> -1 Then
        '    If glCurrentCycle - mlClientLastRequest(lIndex) > ml_REQUEST_THROTTLE Then
        '        mlClientLastRequest(lIndex) = glCurrentCycle
        '    Else : Return
        '    End If
        'End If

        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim iTargetingTactics As Int16 = System.BitConverter.ToInt16(yData, 8)
        Dim iCombatTactics As Int32 = System.BitConverter.ToInt32(yData, 10)
        Dim lExt1 As Int32 = System.BitConverter.ToInt32(yData, 14)
        Dim lExt2 As Int32 = System.BitConverter.ToInt32(yData, 18)

        Dim X As Int32

        Dim oCmdPlayer As Player = Nothing
        If lIndex > -1 Then
            For X = 0 To glPlayerUB
                If glPlayerIdx(X) = mlClientPlayer(lIndex) Then
                    oCmdPlayer = goPlayers(X)
                    Exit For
                End If
            Next X
            If oCmdPlayer Is Nothing Then Return
        End If

        'TODO: Should verify the tactics being set

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        With oEntity
            If .lOwnerID <> mlClientPlayer(lIndex) AndAlso (.lOwnerID <> mlAliasedAs(lIndex) OrElse (mlAliasedRights(lIndex) And AliasingRights.eChangeBehavior) = 0) AndAlso _
                (oCmdPlayer Is Nothing OrElse .Owner.lGuildID < 1 OrElse .Owner.lGuildID <> oCmdPlayer.lGuildID OrElse (.CurrentStatus And elUnitStatus.eGuildAsset) = 0) Then
                gfrmDisplayForm.AddEventLine("Possible Cheat: Non-owner setting entity AI. PlayerID = " & mlClientPlayer(lIndex))
                Return
            End If

            If (iCombatTactics And eiBehaviorPatterns.eTactics_Maneuver) <> 0 Then
                If goEntityDefs(.lEntityDefServerIndex).yChassisType = ChassisType.eGroundBased OrElse goEntityDefs(.lEntityDefServerIndex).yChassisType = ChassisType.eNavalBased Then
                    iCombatTactics -= eiBehaviorPatterns.eTactics_Maneuver
                End If
            End If

            .iCombatTactics = iCombatTactics
            .iTargetingTactics = iTargetingTactics
            .lExtendedCT_1 = lExt1
            .lExtendedCT_2 = lExt2

            If (iCombatTactics And eiBehaviorPatterns.eEngagement_Hold_Fire) <> 0 Then
                If (.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
                    .CurrentStatus = .CurrentStatus Xor elUnitStatus.eUnitEngaged
                    RemoveEntityInCombat(.ObjectID, .ObjTypeID)
                End If
                If .lTargetsServerIdx Is Nothing = False Then
                    For lTmp As Int32 = 0 To .lTargetsServerIdx.GetUpperBound(0)
                        .lTargetsServerIdx(lTmp) = -1
                    Next lTmp
                    .lPrimaryTargetServerIdx = -1
                    Dim lTemp As Int32 = (.CurrentStatus And (elUnitStatus.eUnitEngaged Or elUnitStatus.eSide1HasTarget Or elUnitStatus.eSide2HasTarget Or elUnitStatus.eSide3HasTarget Or elUnitStatus.eSide4HasTarget))
                    If lTemp <> 0 Then
                        .CurrentStatus = .CurrentStatus Xor lTemp
                    End If
                End If
            End If

            'set our force aggression test because we need to make sure the unit takes the new orders into effect
            .bForceAggressionTest = True '(goEntity(X).bForceAggressionTest) Or ((goEntity(X).CurrentStatus And elUnitStatus.eUnitEngaged) <> 0)
            SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)
        End With

        'Now, we should have verified the validity of this message already, so send it to primary
        moPrimary.SendData(yData)

    End Sub

    Public Sub SendAILaunchAll(ByVal oEntity As Epica_Entity)
        Dim yMsg(7) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eAILaunchAll).CopyTo(yMsg, 0)
        oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
        moPrimary.SendData(yMsg)
    End Sub

    Private Sub HandleChatMessage(ByVal yData() As Byte, ByVal lIndex As Int32)
        'ok, let's create it...
        Dim X As Int32
        Dim Y As Int32
        Dim bFound As Boolean = False

        'Dim lLen As Int32
        'Dim sMsg As String
        ''Parse our message
        'lLen = yData.Length
        'For X = 2 To yData.Length - 1
        '    If yData(X) = 0 Then
        '        lLen = X
        '        Exit For
        '    End If
        'Next X
        'sMsg = Mid$(System.Text.ASCIIEncoding.ASCII.GetString(yData), 3, lLen)

        ''TODO: if we need to do any parsing of the message, do it here

        Dim yBroad(yData.Length + 4) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eChatMessage).CopyTo(yBroad, 0)
        System.BitConverter.GetBytes(mlClientPlayer(lIndex)).CopyTo(yBroad, 2)
        yBroad(6) = ChatMessageType.eLocalMessage
        Array.Copy(yData, 2, yBroad, 7, yData.Length - 2)

        For X = 0 To glEnvirUB
            For Y = 0 To goEnvirs(X).lPlayersInEnvirUB
                If goEnvirs(X).lPlayersInEnvirIdx(Y) = mlClientPlayer(lIndex) Then
                    'found it
                    bFound = True
                    BroadcastToEnvironmentClients(yBroad, goEnvirs(X))
                    Exit For
                End If
            Next Y
            If bFound = True Then Exit For
        Next X

    End Sub

    Public Sub SendReloadRequestMsg(ByVal oEntity As Epica_Entity, ByVal oWeapon As WeaponDef)
        Dim yMsg(13) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eReloadWpnMsg).CopyTo(yMsg, 0)
        oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(oWeapon.lEntityWpnDefID).CopyTo(yMsg, 8)
        System.BitConverter.GetBytes(ObjectType.eUnitWeaponDef).CopyTo(yMsg, 12)

        moPrimary.SendData(yMsg)
    End Sub

    Private Sub HandleReloadWpnMsg(ByVal yData() As Byte)
        '2 is msg
        'entity guid
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lWpnID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim lAmmoAmt As Int32 = System.BitConverter.ToInt32(yData, 12)

        'Now, find our entity
        Dim X As Int32

        Dim oEntity As Epica_Entity = Nothing

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X = 0 To lCurUB
        '    If glEntityIdx(X) = lObjID AndAlso goEntity(X).ObjTypeID = iObjTypeID Then
        '        oEntity = goEntity(X)
        '        Exit For
        '    End If
        'Next X
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        'Now, get that entity's weapon def corresponding to the GUID passed in
        With goEntityDefs(oEntity.lEntityDefServerIndex)
            For X = 0 To .WeaponDefUB
                If .WeaponDefs(X).ObjectID = lWpnID Then
                    oEntity.lWpnAmmoCnt(X) = lAmmoAmt
                    Exit For
                End If
            Next X
        End With
    End Sub

    Private Sub HandleRequestPlayerStartLoc(ByVal yData() As Byte)
        'MsgID, PlayerID, EnvirID, EnvirTypeID
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, 6)
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 10)
        Dim X As Int32
        Dim lLocX As Int32 = Int32.MinValue
        Dim lLocZ As Int32 = Int32.MinValue

        Dim oPlanet As Planet = Nothing
        Dim yResp(19) As Byte

        gfrmDisplayForm.AddEventLine("Received Request for Player Start Loc")

        Array.Copy(yData, 0, yResp, 0, 11)

        For X = 0 To glEnvirUB
            If goEnvirs(X).ObjectID = lEnvirID AndAlso goEnvirs(X).ObjTypeID = iEnvirTypeID Then
                oPlanet = CType(goEnvirs(X).oGeoObject, Planet)
                Exit For
            End If
        Next X

        If oPlanet Is Nothing = False Then
            Dim ptTemp As Point = oPlanet.GetPlayerStartLocation()
            If ptTemp.IsEmpty = False Then
                lLocX = ptTemp.X
                lLocZ = ptTemp.Y
            End If
        End If

        System.BitConverter.GetBytes(lLocX).CopyTo(yResp, 12)
        System.BitConverter.GetBytes(lLocZ).CopyTo(yResp, 16)

        moPrimary.SendData(yResp)

    End Sub

    Public Sub SendClearDestList(ByRef oEntity As Epica_Entity)
        Dim yMsg(7) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eClearDestList).CopyTo(yMsg, 0)
        oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
        moPathfinding.SendData(yMsg)
    End Sub

    Private Sub HandleSetFleetDest(ByRef yData() As Byte)
        'Ok, the only thing we care about right now... the Cnt and ID list
        '2 for msg code, 4 for origin system id, 4 for target system id
        Dim lPos As Int32 = 10
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'The end result
        Dim lDestPos As Int32 = 14
        Dim yForward(13 + (lCnt * 22)) As Byte

        Dim lActualCnt As Int32 = 0

        Array.Copy(yData, 0, yForward, 0, 13)

        'Ok, now, get all of the entity locs and place them into the yForward
        For X As Int32 = 0 To lCnt - 1
            Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            Dim oEntity As Epica_Entity = Nothing
            Dim lIdx As Int32 = LookupEntity(lObjID, ObjectType.eUnit)
            If lIdx > -1 Then oEntity = goEntity(lIdx)
            If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> ObjectType.eUnit Then Continue For

            lActualCnt += 1

            System.BitConverter.GetBytes(lObjID).CopyTo(yForward, lDestPos) : lDestPos += 4
            With oEntity
                System.BitConverter.GetBytes(.LocX).CopyTo(yForward, lDestPos) : lDestPos += 4
                System.BitConverter.GetBytes(.LocZ).CopyTo(yForward, lDestPos) : lDestPos += 4
                System.BitConverter.GetBytes(.LocAngle).CopyTo(yForward, lDestPos) : lDestPos += 2
                .ParentEnvir.GetGUIDAsString.CopyTo(yForward, lDestPos) : lDestPos += 6
                System.BitConverter.GetBytes(goEntityDefs(.lEntityDefServerIndex).ModelID).CopyTo(yForward, lDestPos) : lDestPos += 2
            End With
        Next X

        'Now, reduce our msg by difference of cnt and actual
        If lCnt <> lActualCnt Then
            ReDim Preserve yForward(13 + (lActualCnt * 22))
            'And put the actual cnt in
            System.BitConverter.GetBytes(lActualCnt).CopyTo(yForward, 10)
        End If

        moPathfinding.SendData(yForward)

    End Sub

    Private Sub HandleFleetInterSystemMoving(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lFleetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))

        'Now, its a list of lcnt number of UnitID's
        For X As Int32 = 0 To lCnt - 1
            Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim oEntity As Epica_Entity = Nothing
            Dim lIdx As Int32 = LookupEntity(lObjID, ObjectType.eUnit)
            If lIdx > -1 Then oEntity = goEntity(lIdx)
            If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> ObjectType.eUnit Then Return

            With oEntity
                Dim bSendToClients As Boolean = False
                If .ParentEnvir.ObjectID <> lParentID OrElse .ParentEnvir.ObjTypeID <> iParentTypeID Then
                    gfrmDisplayForm.AddEventLine("WARNING: Fleet Movement but unit was not in expected environment! ID: " & .ObjectID)
                    bSendToClients = True
                End If

                'Here's what needs to occur:
                '  1) Send Update message of the unit to the Primary (change environments) with eUpdateEntityAndSave
                SendToPrimary(.GetObjectUpdateMsg(GlobalMessageCode.eUpdateEntityAndSave, RemovalType.eChangingEnvironments))
                '  2) Call the ParentEnvir's Remove Entity
                .ParentEnvir.RemoveEntity(lIdx, RemovalType.eChangingEnvironments, bSendToClients, False, -1)
            End With

            'For lIdx As Int32 = 0 To lCurUB
            '    If glEntityIdx(lIdx) = lID AndAlso goEntity(lIdx).ObjTypeID = ObjectType.eUnit Then
            '        With goEntity(lIdx)
            '            Dim bSendToClients As Boolean = False
            '            If .ParentEnvir.ObjectID <> lParentID OrElse .ParentEnvir.ObjTypeID <> iParentTypeID Then
            '                gfrmDisplayForm.AddEventLine("WARNING: Fleet Movement but unit was not in expected environment! ID: " & .ObjectID)
            '                bSendToClients = True
            '            End If

            '            'Here's what needs to occur:
            '            '  1) Send Update message of the unit to the Primary (change environments) with eUpdateEntityAndSave
            '            SendToPrimary(.GetObjectUpdateMsg(GlobalMessageCode.eUpdateEntityAndSave, RemovalType.eChangingEnvironments))
            '            '  2) Call the ParentEnvir's Remove Entity
            '            .ParentEnvir.RemoveEntity(lIdx, RemovalType.eChangingEnvironments, bSendToClients, False)
            '        End With
            '        '  3) Set the glEntityIdx() to -1
            '        'glEntityIdx(lIdx) = -1 - done in RemoveEntity from the ParentEnvir.RemoveEntity

            '        'Move on to the next unit in the list
            '        Exit For
            '    End If
            'Next lIdx
        Next X

        'And then, forward the message to all connected clients
        For X As Int32 = 0 To glEnvirUB
            If goEnvirs(X).ObjectID = lParentID AndAlso goEnvirs(X).ObjTypeID = iParentTypeID Then
                If goEnvirs(X).lPlayersInEnvirCnt > 0 Then BroadcastToEnvironmentClients(yData, goEnvirs(X))
                Exit For
            End If
        Next X
    End Sub

    Private Sub HandleGetPirateStartLoc(ByRef yData() As Byte)
        'MsgCode (2), PlanetID (4), PlayerID (4), PlayerStartLocX (4), PlayerStartLocY (4)
        Dim lPos As Int32 = 2
        Dim lPlanetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lStartX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lStartY As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim ptResult As Point = New Point(Int32.MinValue, Int32.MinValue)

        For X As Int32 = 0 To glEnvirUB
            If goEnvirs(X).ObjectID = lPlanetID AndAlso goEnvirs(X).ObjTypeID = ObjectType.ePlanet Then
                ptResult = CType(goEnvirs(X).oGeoObject, Planet).GetPirateStartLocation(lPlayerID, lStartX, lStartY)
                Exit For
            End If
        Next X

        Dim yResp(13) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eGetPirateStartLoc).CopyTo(yResp, 0)
        System.BitConverter.GetBytes(lPlanetID).CopyTo(yResp, 2)
        System.BitConverter.GetBytes(ptResult.X).CopyTo(yResp, 6)
        System.BitConverter.GetBytes(ptResult.Y).CopyTo(yResp, 10)
        moPrimary.SendData(yResp)
    End Sub

    Private Sub HandlePlacePirateAssets(ByRef yData() As Byte)
        '- MsgCode, EnvirGUID(6), StartLoc (8), ItmCnt (4), Items...
        Dim lPos As Int32 = 2
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeId As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        If iEnvirTypeId = ObjectType.ePlanet Then
            For X As Int32 = 0 To glEnvirUB
                If goEnvirs(X).ObjectID = lEnvirID AndAlso goEnvirs(X).ObjTypeID = ObjectType.ePlanet Then
                    Dim yResp() As Byte = CType(goEnvirs(X).oGeoObject, Planet).PlacePirateFacilities(yData)
                    If yResp Is Nothing = False Then
                        moPrimary.SendData(yResp)
                    End If
                    Exit For
                End If
            Next X
        End If

    End Sub

    Private Sub HandleMoveEngineer(ByRef yData() As Byte)
        'Primary is telling us to move out from under the object represented by model Id... it assumed that the model is exactly over the engineer
        Dim lEngineerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iEngineerTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lObstacleID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iObstacleTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)

        'Ok, find our entity
        Dim oEntity As Epica_Entity = Nothing
        Dim oObstacle As Epica_Entity = Nothing

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X As Int32 = 0 To lCurUB
        '    If glEntityIdx(X) = lEngineerID AndAlso goEntity(X).ObjTypeID = iEngineerTypeID Then
        '        oEntity = goEntity(X)
        '        Exit For
        '    End If
        'Next X
        Dim lIdx As Int32 = LookupEntity(lEngineerID, iEngineerTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lEngineerID OrElse oEntity.ObjTypeID <> iEngineerTypeID Then Return

        lIdx = -1
        lIdx = LookupEntity(lObstacleID, iObstacleTypeID)
        If lIdx > -1 Then oObstacle = goEntity(lIdx)
        If oObstacle Is Nothing OrElse oObstacle.ObjectID <> lObstacleID OrElse oObstacle.ObjTypeID <> iObstacleTypeID Then Return

        'For X As Int32 = 0 To lCurUB
        '    If glEntityIdx(X) = lObstacleID AndAlso goEntity(X).ObjTypeID = iObstacleTypeID Then
        '        oObstacle = goEntity(X)
        '        Exit For
        '    End If
        'Next X

        If oEntity Is Nothing OrElse oObstacle Is Nothing Then Return
        If oObstacle.lEntityDefServerIndex = -1 OrElse glEntityDefIdx(oObstacle.lEntityDefServerIndex) = -1 Then Return
        Dim oDef As Epica_Entity_Def = goEntityDefs(oObstacle.lEntityDefServerIndex)

        'Ok, now... let's move our entity
        Dim lX As Int32 = oEntity.LocX + CInt(oObstacle.lModelRangeOffset * gl_FINAL_GRID_SQUARE_SIZE * 1.5F)
        Dim lZ As Int32 = oEntity.LocZ
        RotatePoint(oEntity.LocX, oEntity.LocZ, lX, lZ, oEntity.LocAngle \ 10)

        'try to find SOMEWHERE for the end point to be.
        If oEntity.ParentEnvir Is Nothing = False AndAlso oEntity.ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
            If oEntity.lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(oEntity.lEntityDefServerIndex) <> -1 Then
                Dim yChassis As Byte = goEntityDefs(oEntity.lEntityDefServerIndex).yChassisType
                If (yChassis And ChassisType.eGroundBased) <> 0 Then
                    Dim lHt As Int32 = CType(oEntity.ParentEnvir.oGeoObject, Planet).GetHeightAtPoint(lX, lZ, False)
                    If lHt < CType(oEntity.ParentEnvir.oGeoObject, Planet).WaterHeight Then
                        RotatePoint(oEntity.LocX, oEntity.LocZ, lX, lZ, 180)
                        lHt = CType(oEntity.ParentEnvir.oGeoObject, Planet).GetHeightAtPoint(lX, lZ, False)
                        If lHt < CType(oEntity.ParentEnvir.oGeoObject, Planet).WaterHeight Then
                            RotatePoint(oEntity.LocX, oEntity.LocZ, lX, lZ, 90)
                            lHt = CType(oEntity.ParentEnvir.oGeoObject, Planet).GetHeightAtPoint(lX, lZ, False)
                            If lHt < CType(oEntity.ParentEnvir.oGeoObject, Planet).WaterHeight Then
                                RotatePoint(oEntity.LocX, oEntity.LocZ, lX, lZ, 180)
                            End If
                        End If
                    End If
                ElseIf (yChassis And ChassisType.eNavalBased) <> 0 Then
                    Dim lHt As Int32 = CType(oEntity.ParentEnvir.oGeoObject, Planet).GetHeightAtPoint(lX, lZ, False)
                    If lHt > CType(oEntity.ParentEnvir.oGeoObject, Planet).WaterHeight Then
                        RotatePoint(oEntity.LocX, oEntity.LocZ, lX, lZ, 180)
                        lHt = CType(oEntity.ParentEnvir.oGeoObject, Planet).GetHeightAtPoint(lX, lZ, False)
                        If lHt > CType(oEntity.ParentEnvir.oGeoObject, Planet).WaterHeight Then
                            RotatePoint(oEntity.LocX, oEntity.LocZ, lX, lZ, 90)
                            lHt = CType(oEntity.ParentEnvir.oGeoObject, Planet).GetHeightAtPoint(lX, lZ, False)
                            If lHt > CType(oEntity.ParentEnvir.oGeoObject, Planet).WaterHeight Then
                                RotatePoint(oEntity.LocX, oEntity.LocZ, lX, lZ, 180)
                            End If
                        End If
                    End If
                End If
            End If
        End If

        Me.SendAIMoveRequestToPathfinding(oEntity, lX, lZ, oEntity.LocAngle)
    End Sub

    Private Sub HandleSetMaintenanceTarget(ByRef yData() As Byte, ByVal lIndex As Int32, ByVal iMsgCode As Int16)
        If lIndex <> -1 Then
            'If glCurrentCycle - mlClientLastRequest(lIndex) > ml_REQUEST_THROTTLE Then
            If glCurrentCycle - muClientData(lIndex).mlClientLastRequest > ml_REQUEST_THROTTLE Then
                'mlClientLastRequest(lIndex) = glCurrentCycle
                muClientData(lIndex).mlClientLastRequest = glCurrentCycle
            Else : Return
            End If
        End If

        'Ok, it has 2 guids
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        'Now, get both entities
        Dim oEntity As Epica_Entity = Nothing
        Dim oTarget As Epica_Entity = Nothing

        If lObjID = -1 OrElse lTargetID = -1 Then Return
        If iObjTypeID <> ObjectType.eUnit AndAlso iObjTypeID <> ObjectType.eFacility Then Return
        If iTargetTypeID <> ObjectType.eUnit AndAlso iTargetTypeID <> ObjectType.eFacility Then Return

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X As Int32 = 0 To lCurUB
        '    If glEntityIdx(X) = lObjectID AndAlso goEntity(X).ObjTypeID = iObjTypeID Then
        '        oObject = goEntity(X)
        '        Exit For
        '    End If
        'Next X
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return


        If oEntity Is Nothing = False AndAlso lObjID = lTargetID AndAlso iObjTypeID = iTargetTypeID Then
            Dim yForward(41) As Byte

            lPos = 0
            System.BitConverter.GetBytes(iMsgCode).CopyTo(yForward, lPos) : lPos += 2
            System.BitConverter.GetBytes(lObjID).CopyTo(yForward, lPos) : lPos += 4
            System.BitConverter.GetBytes(iObjTypeID).CopyTo(yForward, lPos) : lPos += 2
            System.BitConverter.GetBytes(lTargetID).CopyTo(yForward, lPos) : lPos += 4
            System.BitConverter.GetBytes(iTargetTypeID).CopyTo(yForward, lPos) : lPos += 2
            System.BitConverter.GetBytes(oEntity.LocX).CopyTo(yForward, lPos) : lPos += 4
            System.BitConverter.GetBytes(oEntity.LocZ).CopyTo(yForward, lPos) : lPos += 4
            System.BitConverter.GetBytes(oEntity.CurrentStatus).CopyTo(yForward, lPos) : lPos += 4
            System.BitConverter.GetBytes(oEntity.ArmorHP(0)).CopyTo(yForward, lPos) : lPos += 4
            System.BitConverter.GetBytes(oEntity.ArmorHP(1)).CopyTo(yForward, lPos) : lPos += 4
            System.BitConverter.GetBytes(oEntity.ArmorHP(2)).CopyTo(yForward, lPos) : lPos += 4
            System.BitConverter.GetBytes(oEntity.ArmorHP(3)).CopyTo(yForward, lPos) : lPos += 4

            moPrimary.SendData(yForward)
            Return
        End If

        'For X As Int32 = 0 To lCurUB
        '    If glEntityIdx(X) = lTargetID AndAlso goEntity(X).ObjTypeID = iTargetTypeID Then
        '        oTarget = goEntity(X)
        '        Exit For
        '    End If
        'Next X
        lIdx = -1
        lIdx = LookupEntity(lTargetID, iTargetTypeID)
        If lIdx > -1 Then oTarget = goEntity(lIdx)
        If oTarget Is Nothing OrElse oTarget.ObjectID <> lTargetID OrElse oTarget.ObjTypeID <> iTargetTypeID Then Return
        If oEntity Is Nothing OrElse oTarget Is Nothing OrElse (oEntity.lOwnerID <> mlClientPlayer(lIndex) AndAlso (oEntity.lOwnerID <> mlAliasedAs(lIndex) OrElse (mlAliasedRights(lIndex) And AliasingRights.eMoveUnits) = 0)) Then
            'Sending the response back to the client indicates a failed command
            If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(iMsgCode, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, mlClientPlayer(lIndex))
            moClients(lIndex).SendData(yData)
            Return
        End If

        'Is the target moving?
        If (oTarget.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
            If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(iMsgCode, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, mlClientPlayer(lIndex))
            moClients(lIndex).SendData(yData)
            Return
        End If

        If iTargetTypeID = ObjectType.eFacility AndAlso iObjTypeID = ObjectType.eUnit AndAlso iMsgCode = GlobalMessageCode.eSetRepairTarget Then
            If gb_MONITOR_MSGS = True Then goMsgMonitor.AddOutMsg(iMsgCode, MsgMonitor.eMM_AppType.ClientConnection, yData.Length, mlClientPlayer(lIndex))
            moClients(lIndex).SendData(yData)
            Return
        End If

        With oEntity
            If (.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then
                'send a move lock reset msg
                Dim yMsg(7) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eMoveLockViolate).CopyTo(yMsg, 0)
                .GetGUIDAsString.CopyTo(yMsg, 2)
                SendToPrimary(yMsg)
                .CurrentStatus = .CurrentStatus Xor elUnitStatus.eMoveLock ' oEntity.CurrentStatus -= elUnitStatus.eMoveLock
            End If
        End With

        'Ok, now, we reconstruct our message with what the pathfinding server will need to know
        ReDim Preserve yData(lPos + 25)

        System.BitConverter.GetBytes(oEntity.LocX).CopyTo(yData, lPos) : lPos += 4
        System.BitConverter.GetBytes(oEntity.LocZ).CopyTo(yData, lPos) : lPos += 4
        oEntity.ParentEnvir.GetGUIDAsString.CopyTo(yData, lPos) : lPos += 6

        System.BitConverter.GetBytes(oTarget.LocX).CopyTo(yData, lPos) : lPos += 4
        System.BitConverter.GetBytes(oTarget.LocZ).CopyTo(yData, lPos) : lPos += 4

        'And the model ID of each entity
        If oEntity.lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(oEntity.lEntityDefServerIndex) <> -1 Then
            System.BitConverter.GetBytes(goEntityDefs(oEntity.lEntityDefServerIndex).ModelID).CopyTo(yData, lPos)
        Else : System.BitConverter.GetBytes(CShort(-1)).CopyTo(yData, lPos)
        End If
        lPos += 2
        If oTarget.lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(oTarget.lEntityDefServerIndex) <> -1 Then
            System.BitConverter.GetBytes(goEntityDefs(oTarget.lEntityDefServerIndex).ModelID).CopyTo(yData, lPos)
        Else : System.BitConverter.GetBytes(CShort(-1)).CopyTo(yData, lPos)
        End If
        lPos += 2

        SendToPathfinding(yData)

    End Sub

    Private Sub HandleSetMaintenanceTargetFromPrimary(ByRef yData() As Byte, ByVal iMsgCode As Int16)
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeId As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim lX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lTargetID, iTargetTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lTargetID OrElse oEntity.ObjTypeID <> iTargetTypeID Then Return

        If Math.Abs(oEntity.LocX - lX) < 100 AndAlso Math.Abs(oEntity.LocZ - lZ) < 100 Then
            Dim yResp(yData.Length + 19) As Byte
            yData.CopyTo(yResp, 0)
            System.BitConverter.GetBytes(oEntity.CurrentStatus).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(oEntity.ArmorHP(0)).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(oEntity.ArmorHP(1)).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(oEntity.ArmorHP(2)).CopyTo(yResp, lPos) : lPos += 4
            System.BitConverter.GetBytes(oEntity.ArmorHP(3)).CopyTo(yResp, lPos) : lPos += 4
            SendToPrimary(yResp)
        End If

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))

        'For X As Int32 = 0 To lCurUB
        '    If glEntityIdx(X) = lTargetID AndAlso goEntity(X).ObjTypeID = iTargetTypeID Then

        '        If Math.Abs(goEntity(X).LocX - lX) < 100 AndAlso Math.Abs(goEntity(X).LocZ - lZ) < 100 Then
        '            Dim yResp(yData.Length + 19) As Byte
        '            yData.CopyTo(yResp, 0)
        '            System.BitConverter.GetBytes(goEntity(X).CurrentStatus).CopyTo(yResp, lPos) : lPos += 4
        '            System.BitConverter.GetBytes(goEntity(X).ArmorHP(0)).CopyTo(yResp, lPos) : lPos += 4
        '            System.BitConverter.GetBytes(goEntity(X).ArmorHP(1)).CopyTo(yResp, lPos) : lPos += 4
        '            System.BitConverter.GetBytes(goEntity(X).ArmorHP(2)).CopyTo(yResp, lPos) : lPos += 4
        '            System.BitConverter.GetBytes(goEntity(X).ArmorHP(3)).CopyTo(yResp, lPos) : lPos += 4
        '            SendToPrimary(yResp)
        '        End If

        '        Exit For
        '    End If
        'Next X

    End Sub

    Private Sub HandleRepairCompleted(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lProdID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lQty As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))

        'For X As Int32 = 0 To lCurUB
        '    If glEntityIdx(X) = lObjID AndAlso goEntity(X).ObjTypeID = iObjTypeID Then
        'what was repaired?
        If lProdID > 0 Then
            oEntity.CurrentStatus = oEntity.CurrentStatus Or lProdID
        Else
            Select Case lProdID
                Case -2 'q1
                    oEntity.ArmorHP(UnitArcs.eForwardArc) += lQty
                Case -3 'q2
                    oEntity.ArmorHP(UnitArcs.eLeftArc) += lQty
                Case -4 'q3
                    oEntity.ArmorHP(UnitArcs.eBackArc) += lQty
                Case -5 'q4
                    oEntity.ArmorHP(UnitArcs.eRightArc) += lQty
                Case -6 'structure
                    If oEntity.lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(oEntity.lEntityDefServerIndex) <> -1 Then
                        oEntity.StructuralHP = goEntityDefs(oEntity.lEntityDefServerIndex).Structure_MaxHP
                    End If
                Case Int32.MinValue
                    oEntity.StructuralHP = goEntityDefs(oEntity.lEntityDefServerIndex).Structure_MaxHP \ 10
            End Select
        End If

        '        Exit For
        '    End If
        'Next X
    End Sub

    Private Sub HandleUpdatePlayerTechValue(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPropID As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lNewValue As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        For X As Int32 = 0 To glPlayerUB
            If glPlayerIdx(X) = lPlayerID Then
                Select Case lPropID
                    Case 7          'CP Limit
                        goPlayers(X).lCPLimit = lNewValue
                End Select

                Exit For
            End If
        Next X
    End Sub

    Private Sub HandleFinalizeStopEvent(ByRef yData() As Byte)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        oEntity.bFinalizeStopEvent = False
        oEntity.bForceAggressionTest = True
        SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oEntity.ServerIndex, oEntity.ObjectID)
        moPathfinding.SendData(yData)
        'If oEntity.lTetherPointX = Int32.MinValue Then oEntity.lTetherPointX = oEntity.LocX
        'If oEntity.lTetherPointZ = Int32.MinValue Then oEntity.lTetherPointZ = oEntity.LocZ

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X As Int32 = 0 To lCurUB
        '    If glEntityIdx(X) = lObjID AndAlso goEntity(X) Is Nothing = False AndAlso goEntity(X).ObjTypeID = iTypeID Then
        '        Dim oEntity As Epica_Entity = goEntity(X)
        '        If oEntity Is Nothing = False Then
        '            'If (oEntity.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
        '            '	oEntity.bFinalizeStopEvent = True
        '            '	AddEntityMoving(oEntity.ServerIndex, oEntity.ObjectID)
        '            'Else
        '            oEntity.bFinalizeStopEvent = False
        '            oEntity.bForceAggressionTest = True
        '            SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oEntity.ServerIndex, oEntity.ObjectID) 'AddEntityMoving(oEntity.ServerIndex, oEntity.ObjectID)
        '            moPathfinding.SendData(yData)
        '            If oEntity.lTetherPointX = Int32.MinValue Then oEntity.lTetherPointX = oEntity.LocX
        '            If oEntity.lTetherPointZ = Int32.MinValue Then oEntity.lTetherPointZ = oEntity.LocZ
        '            'End If
        '        End If
        '        Exit For
        '    End If
        'Next X
    End Sub

    Private Sub HandleEntityDefCriticalHitChances(ByRef yData() As Byte)
        Dim lPos As Int32 = 2       'for msgcode
        Dim lDefID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'Ok, get our def
        For X As Int32 = 0 To glEntityDefUB
            If glEntityDefIdx(X) = lDefID AndAlso goEntityDefs(X).ObjTypeID = iTypeID Then
                'ok, found it
                Try
                    For Y As Int32 = 0 To lCnt - 1
                        Dim ySide As UnitArcs = CType(yData(lPos), UnitArcs) : lPos += 1
                        Dim lCritical As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        Dim yChance As Byte = yData(lPos) : lPos += 1
                        goEntityDefs(X).AddCriticalHitChance(ySide, lCritical, yChance)
                    Next Y
                Catch
                    'Do nothing?
                End Try

                Exit For
            End If
        Next X

    End Sub

    'ONLY FROM THE PRIMARY SERVER!
    Private Sub HandleSetPlayerSpecialAttribute(ByRef yData() As Byte)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iAttrID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lValue As Int32 = System.BitConverter.ToInt32(yData, 8)

        If iAttrID = ePlayerSpecialAttributeSetting.eBadWarDecCPPenalty Then

            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) = lPlayerID Then
                    goPlayers(X).BadWarDecCPIncrease = lValue
                    Exit For
                End If
            Next X

            ResyncPlayerCPPenalty(lPlayerID, lValue)

            'For X As Int32 = 0 To glEnvirUB
            '    Dim oEnvir As Envir = goEnvirs(X)
            '    If oEnvir Is Nothing = False Then
            '        oEnvir.ClearPlayersCP(lPlayerID)
            '    End If
            'Next X

            'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
            'For X As Int32 = 0 To lCurUB
            '    If glEntityIdx(X) > 0 Then
            '        Dim oEntity As Epica_Entity = goEntity(X)
            '        If oEntity Is Nothing = False Then
            '            If oEntity.ObjTypeID = ObjectType.eUnit Then
            '                If oEntity.lOwnerID = lPlayerID Then
            '                    'oEntity.ParentEnvir.AdjustPlayerCommandPoints(lPlayerID, lDiff)
            '                    oEntity.ParentEnvir.AdjustPlayerCommandPoints(lPlayerID, lValue + oEntity.CPUsage)
            '                End If
            '            End If
            '        End If
            '    End If
            'Next X

        End If
    End Sub

    Private Sub HandleSetCounterAttack(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Try
            Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim yType As Byte = yData(lPos) : lPos += 1
            Dim lEnemyID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            'Ok, now, get our player
            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) = lPlayerID Then
                    Dim ptLoc As Point = goPlayers(X).GetAlertLoc(lEnvirID, iEnvirTypeID, PlayerAlertType.eUnderAttack)
                    If ptLoc <> Point.Empty Then
                        lLocX = ptLoc.X
                        lLocZ = ptLoc.Y
                        Exit For
                    End If
                End If
            Next X

            'Ok, now... go through the Unit list which indicates the Unit ID's involved in the counter attack and order them to move to lLocX, lLocZ
            ' and set up their AI to work properly
            'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
            For X As Int32 = 0 To lCnt - 1
                Dim lCurrID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim oEntity As Epica_Entity = Nothing
                Dim lIdx As Int32 = LookupEntity(lCurrID, ObjectType.eUnit)
                If lIdx > -1 Then oEntity = goEntity(lIdx)
                If oEntity Is Nothing OrElse oEntity.ObjectID <> lCurrID OrElse oEntity.ObjTypeID <> ObjectType.eUnit Then Return

                'For Y As Int32 = 0 To lCurUB
                '    If glEntityIdx(Y) = lCurrID AndAlso goEntity(Y) Is Nothing = False AndAlso goEntity(Y).ObjTypeID = ObjectType.eUnit Then
                '        'Ok, found it
                '        Dim oTmpEntity As Epica_Entity = goEntity(Y)
                '        If oTmpEntity Is Nothing Then Exit For
                With oEntity
                    If .Owner.ObjectID = lPlayerID Then
                        If .bAIControlled = False Then
                            .Pre_AI_LocX = .LocX
                            .Pre_AI_LocZ = .LocZ
                            .bAIControlled = True
                        End If
                        SendAIMoveRequestToPathfinding(oEntity, lLocX, lLocZ, 0)
                    End If
                End With
                '        Exit For
                '    End If
                'Next Y
            Next X

        Catch ex As Exception
            gfrmDisplayForm.AddEventLine("HandleSetCounterAttack: " & ex.Message)
        End Try
    End Sub

    Private Sub HandlePlayerDiscoversWormhole(ByRef yData() As Byte)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lWormholeID As Int32 = System.BitConverter.ToInt32(yData, 6)
        Dim lSystemID As Int32 = System.BitConverter.ToInt32(yData, 10)
        Dim oWormhole As Wormhole = Nothing

        For X As Int32 = 0 To goGalaxy.mlSystemUB
            If goGalaxy.moSystems(X).ObjectID = lSystemID Then
                With goGalaxy.moSystems(X)
                    For Y As Int32 = 0 To .WormholeUB
                        If .moWormholes(Y) Is Nothing = False AndAlso .moWormholes(Y).ObjectID = lWormholeID Then
                            oWormhole = .moWormholes(Y)
                            Exit For
                        End If
                    Next Y
                End With
                Exit For
            End If
        Next X

        For X As Int32 = 0 To glPlayerUB
            If glPlayerIdx(X) = lPlayerID Then
                goPlayers(X).AddWormholeKnowledge(oWormhole)
                Exit For
            End If
        Next X
    End Sub

    Private Sub HandleJumpTarget(ByRef yData() As Byte, ByVal lClientIndex As Int32)
        'Msg Structure...
        'MsgCode (2), JumpTargetEnvirID (4), JumpTargetEnvirTypeID (2), JumpTargetID (4), JumpTargetTypeID (2)
        ' Honestly, we only care about the GUID List...
        Dim lJumpTargetEnvirID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iJumpTargetEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lJumpTargetID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iJumpTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
        Dim X As Int32
        Dim yForward(23) As Byte
        Dim lPos As Int32 = 0
        Dim lObjID As Int32
        Dim iObjTypeID As Int16
        Dim bFound As Boolean = False
        Dim lDestPos As Int32
        Dim lDestLen As Int32
        Dim bEntityIncluded As Boolean = False

        If lClientIndex <> -1 Then
            'If glCurrentCycle - mlClientLastRequest(lClientIndex) > ml_REQUEST_THROTTLE Then
            If glCurrentCycle - muClientData(lClientIndex).mlClientLastRequest > ml_REQUEST_THROTTLE Then
                'mlClientLastRequest(lClientIndex) = glCurrentCycle
                muClientData(lClientIndex).mlClientLastRequest = glCurrentCycle
            Else : Return
            End If
        End If

        Dim lDestX As Int32 = Int32.MinValue
        Dim lDestZ As Int32 = Int32.MinValue
        Dim iDestA As Int16
        If iJumpTargetTypeID = ObjectType.eWormhole Then
            If iJumpTargetEnvirTypeID = ObjectType.eSolarSystem Then
                For X = 0 To goGalaxy.mlSystemUB
                    If goGalaxy.moSystems(X).ObjectID = lJumpTargetEnvirID Then
                        For Y As Int32 = 0 To goGalaxy.moSystems(X).WormholeUB
                            If goGalaxy.moSystems(X).moWormholes(Y).ObjectID = lJumpTargetID Then
                                With goGalaxy.moSystems(X).moWormholes(Y)
                                    .StartCycle = 1
                                    .WormholeFlags = .WormholeFlags Or elWormholeFlag.eSystem2Detectable Or elWormholeFlag.eSystem1Detectable

                                    If .System1 Is Nothing = False AndAlso .System1.ObjectID = lJumpTargetEnvirID Then
                                        lDestX = .LocX1
                                        lDestZ = .LocY1
                                        iDestA = 0
                                    Else
                                        lDestX = .LocX2
                                        lDestZ = .LocY2
                                        iDestA = 0
                                    End If
                                End With
                                Exit For
                            End If
                        Next Y
                        Exit For
                    End If
                Next X
            Else
                gfrmDisplayForm.AddEventLine("Possible Cheat: Wormhole Parent Envir is not Solar System! " & mlClientPlayer(lClientIndex))
                Return
            End If
        Else
            gfrmDisplayForm.AddEventLine("Possible Cheat: Jump target is not wormhole! " & mlClientPlayer(lClientIndex))
            Return
        End If

        If lDestX = Int32.MinValue OrElse lDestZ = Int32.MinValue Then
            gfrmDisplayForm.AddEventLine("Unable to determine DestX and DestZ: " & mlClientPlayer(lClientIndex))
            Return
        End If

        'MsgCode (2), JumpTargetEnvirID (4), JumpTargetEnvirTypeID (2), JumpTargetID (4), JumpTargetTypeID (2)
        lDestPos = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eJumpTarget).CopyTo(yForward, lDestPos) : lDestPos += 2
        System.BitConverter.GetBytes(lDestX).CopyTo(yForward, lDestPos) : lDestPos += 4
        System.BitConverter.GetBytes(lDestZ).CopyTo(yForward, lDestPos) : lDestPos += 4
        System.BitConverter.GetBytes(iDestA).CopyTo(yForward, lDestPos) : lDestPos += 2
        Array.Copy(yData, 2, yForward, lDestPos, 12) : lDestPos += 12

        'Now, go through our GUID list...
        lPos = 14
        lDestLen = lDestPos - 1

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))

        While lPos + 6 < yData.Length '- 1
            bFound = False
            lObjID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            Dim oEntity As Epica_Entity = Nothing
            Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
            If lIdx > -1 Then oEntity = goEntity(lIdx)
            If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

            ''Now, go thru our entity list
            'For X = 0 To lCurUB
            '    If glEntityIdx(X) = lObjID AndAlso goEntity(X) Is Nothing = False AndAlso goEntity(X).ObjTypeID = iObjTypeID Then
            '        Dim oEntity As Epica_Entity = goEntity(X)
            '        If oEntity Is Nothing Then Exit For
            With oEntity
                If (goEntityDefs(.lEntityDefServerIndex).yChassisType And ChassisType.eSpaceBased) <> 0 Then

                    If .Owner Is Nothing = False AndAlso _
                      (lClientIndex = -1 OrElse .Owner.ObjectID = mlClientPlayer(lClientIndex) OrElse (.lOwnerID = mlAliasedAs(lClientIndex) AndAlso (mlAliasedRights(lClientIndex) And AliasingRights.eChangeEnvironment) <> 0)) AndAlso _
                      ((.CurrentStatus And elUnitStatus.eEngineOperational) <> 0) AndAlso .MaxSpeed > 0 AndAlso .Acceleration > 0 Then
                        bFound = True
                        bEntityIncluded = True

                        lDestLen += 24
                        ReDim Preserve yForward(lDestLen)

                        'Now, put our GUID in the forward message
                        .GetGUIDAsString.CopyTo(yForward, lDestPos) : lDestPos += 6

                        'Now, put our current location there...
                        System.BitConverter.GetBytes(.LocX).CopyTo(yForward, lDestPos) : lDestPos += 4
                        System.BitConverter.GetBytes(.LocZ).CopyTo(yForward, lDestPos) : lDestPos += 4
                        System.BitConverter.GetBytes(.LocAngle).CopyTo(yForward, lDestPos) : lDestPos += 2

                        'And our current parent envir GUID
                        .ParentEnvir.GetGUIDAsString.CopyTo(yForward, lDestPos) : lDestPos += 6

                        If .lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(.lEntityDefServerIndex) <> -1 Then
                            System.BitConverter.GetBytes(goEntityDefs(.lEntityDefServerIndex).ModelID).CopyTo(yForward, lDestPos) : lDestPos += 2
                        Else
                            System.BitConverter.GetBytes(CShort(-1)).CopyTo(yForward, lDestPos) : lDestPos += 2
                        End If

                        .bPlayerMoveRequestPending = True
                        .bAIMoveRequestPending = False
                    End If
                End If
            End With

            '        Exit For
            '    End If
            'Next X

            If bFound = False Then
                'ok... for now, do nothing... it simply won't be included...
            End If
        End While

        If yForward Is Nothing = False AndAlso yForward.Length > 0 AndAlso bEntityIncluded Then SendToPathfinding(yForward)

        If lClientIndex <> -1 Then
            'Ok, send to the primary server to notify it of any possible cancel mining routes
            SendToPrimary(yData)
        End If
    End Sub

    Private Sub HandlePrimaryJumpTarget(ByRef yData() As Byte)
        Dim lPos As Int32 = 2       'for msgcode
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lJumpID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iJumpTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lCurParentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iCurParentTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        oEntity.ParentEnvir.RemoveEntity(oEntity.ServerIndex, RemovalType.eJumping, True, False, -1)

        Dim yResp() As Byte = oEntity.GetObjectUpdateMsg(GlobalMessageCode.eRemoveObject, RemovalType.eJumping)
        lPos = yResp.GetUpperBound(0) + 1
        ReDim Preserve yResp(yResp.GetUpperBound(0) + 12)
        System.BitConverter.GetBytes(lJumpID).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(iJumpTypeID).CopyTo(yResp, lPos) : lPos += 2
        System.BitConverter.GetBytes(lCurParentID).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(iCurParentTypeID).CopyTo(yResp, lPos) : lPos += 2
        SendToPrimary(yResp)

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X As Int32 = 0 To lCurUB
        '    If glEntityIdx(X) = lEntityID AndAlso goEntity(X) Is Nothing = False AndAlso goEntity(X).ObjTypeID = iEntityTypeID Then
        '        Dim oEntity As Epica_Entity = goEntity(X)
        '        If oEntity Is Nothing Then Exit For
        '        oEntity.ParentEnvir.RemoveEntity(oEntity.ServerIndex, RemovalType.eJumping, True, False)

        '        Dim yResp() As Byte = oEntity.GetObjectUpdateMsg(GlobalMessageCode.eRemoveObject, RemovalType.eJumping)
        '        lPos = yResp.GetUpperBound(0) + 1
        '        ReDim Preserve yResp(yResp.GetUpperBound(0) + 12)
        '        System.BitConverter.GetBytes(lJumpID).CopyTo(yResp, lPos) : lPos += 4
        '        System.BitConverter.GetBytes(iJumpTypeID).CopyTo(yResp, lPos) : lPos += 2
        '        System.BitConverter.GetBytes(lCurParentID).CopyTo(yResp, lPos) : lPos += 4
        '        System.BitConverter.GetBytes(iCurParentTypeID).CopyTo(yResp, lPos) : lPos += 2
        '        SendToPrimary(yResp)

        '        Exit For
        '    End If
        'Next X

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

    Private Sub HandleLoginRequest(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2       'msgcode
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oEnc As New StrEncDec
        Dim yUserName(19) As Byte
        Array.Copy(yData, lPos, yUserName, 0, 20) : lPos += 20
        yUserName = oEnc.Decrypt(yUserName)
        Dim yPassword(19) As Byte
        Array.Copy(yData, lPos, yPassword, 0, 20) : lPos += 21
        yPassword = oEnc.Decrypt(yPassword)

        Dim lAliasID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yAliasName(19) As Byte
        Array.Copy(yData, lPos, yAliasName, 0, 20) : lPos += 20
        yAliasName = oEnc.Decrypt(yAliasName)
        Dim yAliasPassword(19) As Byte
        Array.Copy(yData, lPos, yAliasPassword, 0, 20) : lPos += 21
        yAliasPassword = oEnc.Decrypt(yAliasPassword)

        'sValue = GetStringFromBytes(yData, lPos, 20) : lPos += 21   '21
        'If sValue <> "" Then
        '    yTempArray = StringToBytes(sValue)
        '    yTempArray.CopyTo(yAliasPassword, 0)
        'End If

        Dim oPlayer As Player = Nothing
        Dim oAlias As Player = Nothing

        Try
            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) = lPlayerID Then
                    oPlayer = goPlayers(X)
                    If lAliasID = -1 OrElse lAliasID = lPlayerID Then Exit For
                End If
                If glPlayerIdx(X) = lAliasID Then
                    oAlias = goPlayers(X)
                End If

                If oPlayer Is Nothing = False AndAlso oAlias Is Nothing = False Then Exit For
            Next X

            If lAliasID = lPlayerID Then oAlias = Nothing

            'Now, check if oPlayer is nothing
            If oPlayer Is Nothing = False Then
                'Now, check the player's login credentials (must match)
                With oPlayer
                    For X As Int32 = 0 To 19
                        If yUserName(X) <> .PlayerUserName(X) OrElse yPassword(X) <> .PlayerPassword(X) Then
                            Err.Raise(-1, "HandleLoginRequest", "Possible Cheat: Player credentials do not match. PlayerID: " & lPlayerID & ", UserName: " & BytesToString(yUserName) & ", Password: " & BytesToString(yPassword))
                            Exit For
                        End If
                    Next X
                End With

                If oPlayer.AccountStatus <> AccountStatusType.eActiveAccount AndAlso oPlayer.AccountStatus <> AccountStatusType.eMondelisActive AndAlso (oPlayer.AccountStatus <> AccountStatusType.eOpenBetaAccount OrElse gb_IN_OPEN_BETA = False) AndAlso oPlayer.AccountStatus <> AccountStatusType.eTrialAccount Then
                    Err.Raise(-1, "HandleLoginRequest", "Possible Cheat: Account is not active. PlayerID: " & lPlayerID)
                End If
            Else
                'Invalid login
                Err.Raise(-1, "HandleLoginRequest", "Possible Cheat: Player Object not found. Player: " & lPlayerID)
            End If

            Dim lRights As Int32 = 0
            If oAlias Is Nothing = False Then
                'Now, check the alias login...
                Dim bFound As Boolean = False

                With oPlayer
                    If .ObjectID = 1 OrElse .ObjectID = 2 OrElse .ObjectID = 6 OrElse .ObjectID = 221 OrElse .ObjectID = 131 OrElse .ObjectID = 2067 OrElse .ObjectID = 7 OrElse .ObjectID = 2076 OrElse .ObjectID = 1780 Then
                        bFound = True
                        lRights = AliasingRights.eAddProduction Or AliasingRights.eAddResearch Or AliasingRights.eAlterAgents Or AliasingRights.eAlterAutoLaunchPower Or AliasingRights.eAlterColonyStats Or AliasingRights.eAlterDiplomacy Or AliasingRights.eAlterEmail Or AliasingRights.eAlterTrades Or AliasingRights.eCancelProduction Or AliasingRights.eCancelResearch Or AliasingRights.eChangeBehavior Or AliasingRights.eChangeEnvironment Or AliasingRights.eCreateBattleGroups Or AliasingRights.eCreateDesigns Or AliasingRights.eDockUndockUnits Or AliasingRights.eModifyBattleGroups Or AliasingRights.eMoveUnits Or AliasingRights.eTransferCargo Or AliasingRights.eViewAgents Or AliasingRights.eViewBattleGroups Or AliasingRights.eViewBudget Or AliasingRights.eViewColonyStats Or AliasingRights.eViewDiplomacy Or AliasingRights.eViewEmail Or AliasingRights.eViewMining Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns Or AliasingRights.eViewTrades Or AliasingRights.eViewTreasury Or AliasingRights.eViewUnitsAndFacilities
                    End If
                End With

                If bFound = False Then
                    For X As Int32 = 0 To oAlias.lAliasUB
                        If oAlias.lAliasIdx(X) = lPlayerID Then
                            With oAlias.uAliasLogin(X)
                                bFound = True
                                lRights = .lRights
                                For Y As Int32 = 0 To 19
                                    If .yUserName(Y) <> yAliasName(Y) OrElse .yPassword(Y) <> yAliasPassword(Y) Then
                                        'Invalid login
                                        bFound = False
                                        Exit For
                                    End If
                                Next Y
                            End With
                            Exit For
                        End If
                    Next X
                End If

                If bFound = False Then
                    Err.Raise(-1, "HandleLoginRequest", "Possible Cheat: Alias UserName/Password Mismatch. PlayerID: " & lPlayerID & ", Aliased Player: " & lAliasID)
                End If
            ElseIf lAliasID <> -1 AndAlso lAliasID <> lPlayerID Then
                Err.Raise(-1, "HandleLoginRequest", "Possible Cheat: Alias UserName/Password Mismatch. PlayerID: " & lPlayerID & ", Aliased Player: " & lAliasID)
            End If

            oPlayer.oSocket = moClients(lIndex)
            mlClientPlayer(lIndex) = oPlayer.ObjectID
            If oAlias Is Nothing = False Then
                mlAliasedAs(lIndex) = oAlias.ObjectID
                mlAliasedRights(lIndex) = lRights
                oPlayer.lAliasingPlayerID = oAlias.ObjectID
            Else
                mlAliasedAs(lIndex) = -1
                oPlayer.lAliasingPlayerID = -1
            End If

        Catch ex As Exception
            gfrmDisplayForm.AddEventLine(ex.Source & " - " & ex.Message)
            moClients(lIndex).Disconnect()
            moClients(lIndex) = Nothing
            mlClientPlayer(lIndex) = -1
            mlAliasedAs(lIndex) = -1
            mlAliasedRights(lIndex) = 0
        End Try
    End Sub

    Private Sub HandlePlayerInitialSetup(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'msgcode
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        For X As Int32 = 0 To glPlayerUB
            If glPlayerIdx(X) = lPlayerID Then
                ReDim goPlayers(X).PlayerUserName(19)
                ReDim goPlayers(X).PlayerPassword(19)
                Array.Copy(yData, lPos, goPlayers(X).PlayerUserName, 0, 20) : lPos += 20
                Array.Copy(yData, lPos, goPlayers(X).PlayerPassword, 0, 20) : lPos += 20
                Exit For
            End If
        Next X
    End Sub

    Private Sub HandlePlayerAliasConfig(ByRef yData() As Byte)
        'MsgCode 2, Type(1),AliasPlayer(20), AliasUN(20), AliasPW(20), Rights(4)
        Dim lPos As Int32 = 2       'for msgcode
        Dim yType As Byte = yData(lPos) : lPos += 1
        'first 4 bytes of playername is PlayerID
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'second 4 bytes of playername is OtherID
        Dim lOtherPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lPos += 12      'for remainder of playername
        Dim yUserName(19) As Byte
        Array.Copy(yData, lPos, yUserName, 0, 20) : lPos += 20
        Dim yPassword(19) As Byte
        Array.Copy(yData, lPos, yPassword, 0, 20) : lPos += 20
        Dim lRights As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'Ok, get our player
        Dim oPlayer As Player = Nothing
        For X As Int32 = 0 To glPlayerUB
            If glPlayerIdx(X) = lPlayerID Then
                oPlayer = goPlayers(X)
                Exit For
            End If
        Next X
        If oPlayer Is Nothing Then Return

        'Now, check our type
        If yType = 0 Then
            'Removal, lOtherPlayerID is in my alias list
            For X As Int32 = 0 To oPlayer.lAliasUB
                If oPlayer.lAliasIdx(X) = lOtherPlayerID Then
                    oPlayer.lAliasIdx(X) = -1
                    With oPlayer.uAliasLogin(X)
                        ReDim .yPassword(19)
                        ReDim .yUserName(19)
                        .lRights = 0
                    End With

                    'drop the player
                    If oPlayer.oAliases(X).oSocket Is Nothing = False Then oPlayer.oAliases(X).oSocket.Disconnect()
                    oPlayer.oAliases(X) = Nothing
                    Return
                End If
            Next X
        Else
            'Add or Update... get our other player
            Dim oOther As Player = Nothing
            For X As Int32 = 0 To glPlayerUB
                If glPlayerIdx(X) = lOtherPlayerID Then
                    oOther = goPlayers(X)
                    Exit For
                End If
            Next X
            If oOther Is Nothing Then Return

            'Now, do our add/update
            Dim bFound As Boolean = False
            For X As Int32 = 0 To oPlayer.lAliasUB
                If oPlayer.lAliasIdx(X) = lOtherPlayerID Then
                    'Ok, found it, update it
                    bFound = True
                    With oPlayer.uAliasLogin(X)
                        .yUserName = yUserName
                        .yPassword = yPassword
                        .lRights = lRights
                    End With

                    Exit For
                End If
            Next X
            If bFound = False Then
                For X As Int32 = 0 To oPlayer.lAliasUB
                    If oPlayer.lAliasIdx(X) = -1 Then
                        oPlayer.oAliases(X) = oOther
                        With oPlayer.uAliasLogin(X)
                            .yUserName = yUserName
                            .yPassword = yPassword
                            .lRights = lRights
                        End With
                        oPlayer.lAliasIdx(X) = oOther.ObjectID
                        bFound = True
                        Exit For
                    End If
                Next X
                If bFound = False Then
                    Dim lIdx As Int32 = oPlayer.lAliasUB + 1
                    ReDim Preserve oPlayer.oAliases(lIdx)
                    ReDim Preserve oPlayer.lAliasIdx(lIdx)
                    ReDim Preserve oPlayer.uAliasLogin(lIdx)
                    oPlayer.oAliases(lIdx) = oOther
                    oPlayer.lAliasIdx(lIdx) = oOther.ObjectID
                    With oPlayer.uAliasLogin(lIdx)
                        .yUserName = yUserName
                        .yPassword = yPassword
                        .lRights = lRights
                    End With
                    oPlayer.lAliasUB = lIdx
                End If

                For X As Int32 = 0 To mlClientUB
                    If mlClientPlayer(X) = oOther.ObjectID Then
                        If mlAliasedAs(X) = oPlayer.ObjectID Then mlAliasedRights(X) = lRights
                        Exit For
                    End If
                Next X
            End If

        End If
    End Sub

    Private Sub HandleCameraPosUpdate(ByRef yData() As Byte, ByVal lIndex As Int32)
        muClientData(lIndex).lCameraPosX = yData(2)
        muClientData(lIndex).lCameraPosZ = yData(3)
    End Sub

    Private Sub HandleRemovePlayerRel(ByRef yData() As Byte)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lWithPlayerID As Int32 = System.BitConverter.ToInt32(yData, 6)

        Dim oPlayer As Player = Nothing
        For X As Int32 = 0 To glPlayerUB
            If glPlayerIdx(X) = lPlayerID Then
                oPlayer = goPlayers(X)
                Exit For
            End If
        Next X
        If oPlayer Is Nothing Then Return

        'ok, now, what is withplayerid?
        oPlayer.RemovePlayerRel(lWithPlayerID, False)
    End Sub

    Private Sub HandlePrimarySetEntityProduction(ByRef yData() As Byte)
        'need to append the loc data
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        'Get the build location...
        Dim lBuildID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iBuildTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
        Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, 14)
        Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, 18)
        Dim bMine As Boolean = False

        'Dim X As Int32

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        'Confirm that the build location is ok for building based on terrain
        If oEntity.ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
            Dim oDef As Epica_Entity_Def = Nothing
            For Y As Int32 = 0 To glEntityDefUB
                If glEntityDefIdx(Y) = lBuildID AndAlso goEntityDefs(Y).ObjTypeID = iBuildTypeID Then
                    If goEntityDefs(Y).ProductionTypeID = ProductionType.eMining Then bMine = True
                    oDef = goEntityDefs(Y)
                    Exit For
                End If
            Next Y
            If oDef Is Nothing Then Return

            Dim bNaval As Boolean = (oDef.yChassisType And ChassisType.eNavalBased) <> 0

            'Ok, is the angle of the terrain okay?
            If bMine = False AndAlso CType(oEntity.ParentEnvir.oGeoObject, Planet).TerrainGradeBuildable(lLocX, lLocZ) = False Then
                Return    'exit for loop because we'll hit the ProdFailed msg
            End If

            'Is the terrain height high enough?
            If bNaval = False Then
                If CType(oEntity.ParentEnvir.oGeoObject, Planet).GetHeightAtPoint(lLocX, lLocZ, False) < CType(oEntity.ParentEnvir.oGeoObject, Planet).WaterHeight Then
                    Return
                End If
            Else
                If CType(oEntity.ParentEnvir.oGeoObject, Planet).GetHeightAtPoint(lLocX, lLocZ, False) > CType(oEntity.ParentEnvir.oGeoObject, Planet).WaterHeight Then
                    Return
                End If
            End If

            Dim lHalfVal As Int32 = oDef.lModelSizeXZ \ 2
            Dim lMinHt As Int32 = Int32.MaxValue
            Dim lMaxHt As Int32 = Int32.MinValue
            With CType(oEntity.ParentEnvir.oGeoObject, Planet)
                Dim LocX As Int32 = CInt(lLocX)
                Dim LocZ As Int32 = CInt(lLocZ)
                Dim lTemp As Int32 = CInt(.GetHeightAtPoint(LocX - lHalfVal, LocZ - lHalfVal, False))
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX + lHalfVal, LocZ - lHalfVal, False))
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX - lHalfVal, LocZ + lHalfVal, False))
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX + lHalfVal, LocZ + lHalfVal, False))
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
                lTemp = CInt(.GetHeightAtPoint(LocX, LocZ, False))
                If lTemp > lMaxHt Then lMaxHt = lTemp
                If lTemp < lMinHt Then lMinHt = lTemp
            End With
            If lMaxHt - lMinHt > 200 AndAlso oDef.ProductionTypeID <> ProductionType.eMining Then Return
        End If

        ReDim Preserve yData(39)
        System.BitConverter.GetBytes(oEntity.LocX).CopyTo(yData, 24)
        System.BitConverter.GetBytes(oEntity.LocZ).CopyTo(yData, 28)

        oEntity.ParentEnvir.GetGUIDAsString.CopyTo(yData, 32)

        If oEntity.lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(oEntity.lEntityDefServerIndex) <> -1 Then
            System.BitConverter.GetBytes(goEntityDefs(oEntity.lEntityDefServerIndex).ModelID).CopyTo(yData, 38)
        Else : System.BitConverter.GetBytes(CShort(-1)).CopyTo(yData, 38)
        End If

        SendToPathfinding(yData)

    End Sub

    Private Sub HandleAddFormation(ByRef yData() As Byte)
        Dim oFormation As FormationDef = New FormationDef
        oFormation.FillFromMsg(yData)

        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To glFormationDefUB
            If glFormationDefIdx(X) = oFormation.lFormationID Then
                goFormationDef(X) = oFormation
                Return
            ElseIf lIdx = -1 AndAlso glFormationDefIdx(X) = -1 Then
                lIdx = X
            End If
        Next X
        If lIdx = -1 Then
            ReDim Preserve glFormationDefIdx(glFormationDefUB + 1)
            ReDim Preserve goFormationDef(glFormationDefUB + 1)
            goFormationDef(glFormationDefUB + 1) = oFormation
            glFormationDefIdx(glFormationDefUB + 1) = oFormation.lFormationID
            glFormationDefUB += 1
        Else
            goFormationDef(lIdx) = oFormation
            glFormationDefIdx(lIdx) = oFormation.lFormationID
        End If
    End Sub

    Private Sub HandleMoveFormation(ByRef yData() As Byte, ByVal lClientIndex As Int32)
        'Msg Structure...
        'MsgCode - 2, FormationID - 4, DestX - 4, DestZ - 4, DestA - 2, DestID - 4, DestTypeID - 2, GUID List...
        ' Honestly, we only care about the GUID List...

        If lClientIndex <> -1 Then
            If glCurrentCycle - muClientData(lClientIndex).mlClientLastRequest > ml_REQUEST_THROTTLE Then
                muClientData(lClientIndex).mlClientLastRequest = glCurrentCycle
            Else : Return
            End If
        End If

        Dim yForward(32) As Byte
        Dim lPos As Int32 = 0
        Dim lObjID As Int32
        Dim iObjTypeID As Int16
        Dim bFound As Boolean = False
        'Dim X As Int32
        Dim lDestPos As Int32
        Dim lDestLen As Int32
        Dim bEntityIncluded As Boolean = False

        Dim bCanEnterEnvir As Boolean = False
        Dim lFormationID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, 21)

        Dim yCritType As CriteriaType = CriteriaType.eHullSize
        For X As Int32 = 0 To glFormationDefUB
            If glFormationDefIdx(X) = lFormationID Then
                yCritType = goFormationDef(X).yCriteria
                Exit For
            End If
        Next X

        Array.Copy(yData, 0, yForward, 0, 23)   'copy the code, dest coord and dest envir guid

        'Now, go through our GUID list...
        lPos = 23
        lDestPos = 33
        lDestLen = 32

        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))

        Dim lMinSpeed As Int32 = 255
        Dim lMinSpeedID As Int32 = -1
        Dim iMinSpeedTypeID As Int16 = -1

        For lItem As Int32 = 0 To lCnt - 1
            bFound = False
            lObjID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            Dim oEntity As Epica_Entity = Nothing
            Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
            If lIdx > -1 Then oEntity = goEntity(lIdx)
            If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return
            ''Now, go thru our entity list
            'For X = 0 To lCurUB
            '    If glEntityIdx(X) = lObjID Then
            '        'goEntity(X).ObjTypeID = iObjTypeID Then
            '        Dim oEntity As Epica_Entity = goEntity(X)
            '        If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = iObjTypeID Then

            With oEntity
                bCanEnterEnvir = False
                If .lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(.lEntityDefServerIndex) <> -1 Then
                    If iDestTypeID = ObjectType.ePlanet Then
                        If (goEntityDefs(.lEntityDefServerIndex).yChassisType And ChassisType.eAtmospheric) <> 0 OrElse _
                           (goEntityDefs(.lEntityDefServerIndex).yChassisType And (ChassisType.eGroundBased Or ChassisType.eNavalBased)) <> 0 Then
                            bCanEnterEnvir = True
                        End If
                    ElseIf iDestTypeID = ObjectType.eSolarSystem Then
                        If (goEntityDefs(.lEntityDefServerIndex).yChassisType And ChassisType.eSpaceBased) <> 0 Then
                            bCanEnterEnvir = True
                        End If
                    End If
                End If

                If bCanEnterEnvir = True AndAlso .Owner Is Nothing = False Then
                    If (lClientIndex = -1 OrElse .Owner.ObjectID = mlClientPlayer(lClientIndex) OrElse (.lOwnerID = mlAliasedAs(lClientIndex) AndAlso (mlAliasedRights(lClientIndex) And AliasingRights.eMoveUnits) <> 0)) AndAlso _
                      ((.CurrentStatus And elUnitStatus.eEngineOperational) <> 0) AndAlso .MaxSpeed > 0 AndAlso .Acceleration > 0 Then

                        'one more thing, make sure this unit flies...
                        If .lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(.lEntityDefServerIndex) <> -1 Then
                            If (goEntityDefs(.lEntityDefServerIndex).yChassisType And (ChassisType.eGroundBased Or ChassisType.eNavalBased)) <> 0 Then Exit For
                        End If

                        If (.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then
                            'send a move lock reset msg
                            Dim yMsg(7) As Byte
                            System.BitConverter.GetBytes(GlobalMessageCode.eMoveLockViolate).CopyTo(yMsg, 0)
                            .GetGUIDAsString.CopyTo(yMsg, 2)
                            SendToPrimary(yMsg)
                            .CurrentStatus = .CurrentStatus Xor elUnitStatus.eMoveLock ' oEntity.CurrentStatus -= elUnitStatus.eMoveLock
                        End If

                        bFound = True
                        bEntityIncluded = True

                        If .MaxSpeed < lMinSpeed Then
                            lMinSpeedID = .ObjectID
                            iMinSpeedTypeID = .ObjTypeID
                            lMinSpeed = .MaxSpeed
                        End If

                        lDestLen += 22
                        ReDim Preserve yForward(lDestLen)

                        'Now, put our GUID in the forward message
                        .GetGUIDAsString.CopyTo(yForward, lDestPos) : lDestPos += 6

                        'Now, put our current location there...
                        System.BitConverter.GetBytes(.LocX).CopyTo(yForward, lDestPos) : lDestPos += 4
                        System.BitConverter.GetBytes(.LocZ).CopyTo(yForward, lDestPos) : lDestPos += 4
                        System.BitConverter.GetBytes(.LocAngle).CopyTo(yForward, lDestPos) : lDestPos += 2

                        If .lEntityDefServerIndex <> -1 AndAlso glEntityDefIdx(.lEntityDefServerIndex) <> -1 Then
                            System.BitConverter.GetBytes(goEntityDefs(.lEntityDefServerIndex).ModelID).CopyTo(yForward, lDestPos) : lDestPos += 2
                        Else
                            System.BitConverter.GetBytes(CShort(-1)).CopyTo(yForward, lDestPos) : lDestPos += 2
                        End If

                        'crit value...
                        Dim lValue As Int32 = 0
                        Dim oDef As Epica_Entity_Def = Nothing
                        If .lEntityDefServerIndex <> -1 Then oDef = goEntityDefs(.lEntityDefServerIndex)

                        Select Case yCritType
                            Case CriteriaType.eCargoBayCap
                                If oDef Is Nothing = False Then lValue = oDef.Cargo_Cap
                            Case CriteriaType.eCombatRating
                                If oDef Is Nothing = False Then lValue = oDef.CombatRating
                            Case CriteriaType.eEntityName
                                lValue = .UnitName(0)
                            Case CriteriaType.eHangarCap
                                If oDef Is Nothing = False Then lValue = oDef.Hangar_Cap
                            Case CriteriaType.eHullSize
                                If oDef Is Nothing = False Then lValue = oDef.HullSize
                            Case CriteriaType.eManeuver
                                If oDef Is Nothing = False Then lValue = oDef.BaseManeuver
                            Case CriteriaType.eMaxRadarRange
                                If oDef Is Nothing = False Then lValue = oDef.yDefOptRadarRange
                            Case CriteriaType.eMostFrontArmorHP
                                If oDef Is Nothing = False Then lValue = oDef.Armor_MaxHP(UnitArcs.eForwardArc)
                            Case CriteriaType.eMostShieldHP
                                If oDef Is Nothing = False Then lValue = oDef.Shield_MaxHP
                            Case CriteriaType.eSpeed
                                If oDef Is Nothing = False Then lValue = oDef.BaseMaxSpeed
                            Case CriteriaType.eWeaponSlots
                                If oDef Is Nothing = False Then lValue = oDef.WeaponDefUB + 1
                        End Select

                        System.BitConverter.GetBytes(lValue).CopyTo(yForward, lDestPos) : lDestPos += 4

                        .bPlayerMoveRequestPending = True
                        .bAIMoveRequestPending = False
                        '.lTetherPointX = Int32.MinValue
                        '.lTetherPointZ = Int32.MinValue
                    End If
                End If
            End With

            '            Exit For
            '        End If
            '    End If
            'Next X

            If bFound = False Then
                'ok... for now, do nothing... it simply won't be included...
                lCnt -= 1
            End If
        Next lItem

        System.BitConverter.GetBytes(lCnt).CopyTo(yForward, 23)
        System.BitConverter.GetBytes(lMinSpeedID).CopyTo(yForward, 27)
        System.BitConverter.GetBytes(iMinSpeedTypeID).CopyTo(yForward, 31)

        If yForward Is Nothing = False AndAlso yForward.Length > 0 AndAlso bEntityIncluded Then SendToPathfinding(yForward)

        If lClientIndex <> -1 Then
            'Ok, send to the primary server to notify it of any possible cancel mining routes
            SendToPrimary(yData)
        End If
    End Sub

    Private Sub HandlePFMoveFormation(ByRef yData() As Byte)
        Dim lPos As Int32 = 4 ' for msgcode and placeholder for max speed and placeholder for maneuver
        Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim iDestA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lObjID(lCnt - 1) As Int32
        Dim iObjTypeID(lCnt - 1) As Int16
        Dim lDestX(lCnt - 1) As Int32
        Dim lDestZ(lCnt - 1) As Int32
        Dim oEntity(lCnt - 1) As Epica_Entity

        Dim lMinSpeed As Int32 = 255
        Dim lMinManeuver As Int32 = 255

        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))

        For X As Int32 = 0 To lCnt - 1
            lObjID(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iObjTypeID(X) = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            lDestX(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            lDestZ(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            Dim lIdx As Int32 = LookupEntity(lObjID(X), iObjTypeID(X))
            If lIdx > -1 Then oEntity(X) = goEntity(lIdx)
            If oEntity(X) Is Nothing OrElse oEntity(X).ObjectID <> lObjID(X) OrElse oEntity(X).ObjTypeID <> iObjTypeID(X) Then oEntity(X) = Nothing
            If (oEntity(X).CurrentStatus And elUnitStatus.eEngineOperational) <> 0 Then
                If oEntity(X).MaxSpeed < lMinSpeed Then lMinSpeed = oEntity(X).MaxSpeed
                If oEntity(X).Maneuver < lMinManeuver Then lMinManeuver = oEntity(X).Maneuver
            End If
            'For Y As Int32 = 0 To lCurUB
            '    If glEntityIdx(Y) = lObjID(X) Then
            '        Dim oTemp As Epica_Entity = goEntity(Y)
            '        If oTemp Is Nothing = False AndAlso oTemp.ObjTypeID = iObjTypeID(X) Then
            '            oEntity(X) = oTemp
            '            If (oEntity(X).CurrentStatus And elUnitStatus.eEngineOperational) <> 0 Then
            '                If oEntity(X).MaxSpeed < lMinSpeed Then lMinSpeed = oEntity(X).MaxSpeed
            '                If oEntity(X).Maneuver < lMinManeuver Then lMinManeuver = oEntity(X).Maneuver
            '            End If
            '            Exit For
            '        End If
            '    End If
            'Next Y
        Next X

        'Public fFormationAcceleration As Single = 0.0F
        Dim fAcc As Single = lMinManeuver * 0.01F ' / 100.0F
        Dim yTurnAmt As Byte = CByte(lMinManeuver)
        Dim iTurnAmt100 As Int16 = yTurnAmt * 100S

        For X As Int32 = 0 To lCnt - 1
            If oEntity(X) Is Nothing = False AndAlso (oEntity(X).CurrentStatus And elUnitStatus.eEngineOperational) <> 0 Then
                'if the entity has the MiningOrBuilding flag...
                With oEntity(X)
                    'If (.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then
                    '    'send a move lock reset msg
                    '    Dim yMsg(7) As Byte
                    '    System.BitConverter.GetBytes(GlobalMessageCode.eMoveLockViolate).CopyTo(yMsg, 0)
                    '    .GetGUIDAsString.CopyTo(yMsg, 2)
                    '    SendToPrimary(yMsg)
                    '    .CurrentStatus = .CurrentStatus Xor elUnitStatus.eMoveLock
                    'End If

                    .TrueDestAngle = iDestA
                    .CurrentStatus = oEntity(X).CurrentStatus Or elUnitStatus.eMovedByPlayer
                    .bPlayerMoveRequestPending = False
                    .bAIMoveRequestPending = False

                    'By setting these, the unit uses them instead of its actual values...
                    .yFormationManeuver = yTurnAmt
                    .yFormationMaxSpeed = CByte(lMinSpeed)
                    .yFormationTurnAmount = yTurnAmt
                    .iFormationTurnAmount100 = iTurnAmt100
                    .fFormationAcceleration = fAcc

                    .SetDest(lDestX(X), lDestZ(X), iDestA)
                End With
            End If
        Next X

        Dim oEnvir As Envir = Nothing
        For X As Int32 = 0 To glEnvirUB
            If goEnvirs(X).ObjectID = lDestID AndAlso goEnvirs(X).ObjTypeID = iDestTypeID Then
                oEnvir = goEnvirs(X)
                Exit For
            End If
        Next X

        yData(2) = CByte(lMinSpeed)
        yData(3) = CByte(lMinManeuver)

        If oEnvir Is Nothing = False AndAlso oEnvir.lPlayersInEnvirCnt > 0 Then BroadcastToEnvironmentClients(yData, oEnvir)
    End Sub

    Private Sub HandleSetIronCurtain(ByRef yData() As Byte)
        Dim lPos As Int32 = 2 'for msgcode
        Dim lPlanetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yValue As Byte = yData(lPos) : lPos += 1

        Dim oPlayer As Player = Nothing

        Dim lCurUB As Int32 = Math.Min(glPlayerUB, glPlayerIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If glPlayerIdx(X) = lPlayerID Then
                oPlayer = goPlayers(X)
                Exit For
            End If
        Next X
        If oPlayer Is Nothing = False Then
            Dim bForceAgg As Boolean = False
            For X As Int32 = 0 To glEnvirUB
                If oPlayer.iIronCurtainStatus = 255 Then
                    If goEnvirs(X).lPlayersInEnvirCnt > 0 Then BroadcastToEnvironmentClients(yData, goEnvirs(X))
                    If goEnvirs(X).PotentialAggression = True Then bForceAgg = True
                Else
                    If goEnvirs(X).ObjTypeID = ObjectType.ePlanet AndAlso goEnvirs(X).ObjectID = lPlanetID Then
                        If goEnvirs(X).lPlayersInEnvirCnt > 0 Then BroadcastToEnvironmentClients(yData, goEnvirs(X))
                        If goEnvirs(X).PotentialAggression = True Then bForceAgg = True
                        Exit For
                    End If
                End If
            Next X
            oPlayer.iIronCurtainStatus = yValue
            If yValue = 0 Then
                If oPlayer.lIronCurtainPlanetID <> -1 Then
                    oPlayer.lIronCurtainPlanetID = -1
                Else : Return           'nothing to do, the value is already set
                End If
            ElseIf yValue = 255 Then
                oPlayer.lIronCurtainPlanetID = Int32.MinValue
                lCurUB = -1
            Else
                'If oPlayer.lIronCurtainPlanetID < 1 Then
                '	oPlayer.lIronCurtainPlanetID = lPlanetID
                'Else : Return			'nothing to do, the value is already set
                'End If
                oPlayer.lIronCurtainPlanetID = lPlanetID
            End If

            'All units belonging to this player need to force an aggression test (that sucks)
            If bForceAgg = True Then
                lCurUB = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
                For X As Int32 = 0 To lCurUB
                    If glEntityIdx(X) > 0 Then
                        Dim oEntity As Epica_Entity = goEntity(X)
                        If oEntity Is Nothing = False Then

                            If yValue = 255 Then
                                If oEntity.lOwnerID = lPlayerID Then
                                    With oEntity
                                        If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then .CurrentStatus = .CurrentStatus Xor elUnitStatus.eUnitMoving
                                        If (.CurrentStatus And elUnitStatus.eMovedByPlayer) <> 0 Then .CurrentStatus = .CurrentStatus Xor elUnitStatus.eMovedByPlayer
                                        If (.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then .CurrentStatus = .CurrentStatus Xor elUnitStatus.eUnitEngaged

                                        RemoveEntityInCombat(.ObjectID, .ObjTypeID)

                                        For Y As Int32 = 0 To oEntity.lTargetedByUB
                                            Try
                                                Dim lTmpIdx As Int32 = oEntity.lTargetedByIdx(Y)
                                                If lTmpIdx > -1 AndAlso glEntityIdx(lTmpIdx) = oEntity.lTargetedByID(Y) Then
                                                    Dim oTmpEntity As Epica_Entity = goEntity(lTmpIdx)
                                                    If oTmpEntity Is Nothing = False Then
                                                        oTmpEntity.bForceAggressionTest = True
                                                        SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oTmpEntity.ServerIndex, oTmpEntity.ObjectID)
                                                    End If
                                                End If
                                            Catch
                                            End Try
                                        Next Y

                                        SyncLockMovementRegisters(MovementCommand.RemoveEntityMoving, .ServerIndex, .ObjectID)
                                        .yResetMoveByPlayer = 0

                                        goMsgSys.SendClearDestList(oEntity)

                                        Dim yOutMsg() As Byte = CreateStopObjectCommand(oEntity)
                                        SendToPathfinding(yOutMsg)
                                        If .ParentEnvir.lPlayersInEnvirCnt > 0 Then BroadcastToEnvironmentClients(yOutMsg, .ParentEnvir)
                                    End With
                                End If
                            End If

                            'If oEntity.lOwnerID = lPlayerID AndAlso oEntity.ParentEnvir Is Nothing = False Then
                            If oEntity.ParentEnvir Is Nothing = False Then
                                If yValue = 255 OrElse (oEntity.ParentEnvir.ObjectID = lPlanetID AndAlso oEntity.ParentEnvir.ObjTypeID = ObjectType.ePlanet) Then
                                    If yValue <> 0 Then
                                        oEntity.ClearTargetLists(False, 0)
                                    End If
                                    oEntity.bForceAggressionTest = True
                                    SyncLockMovementRegisters(MovementCommand.AddEntityMoving, X, oEntity.ObjectID) 'AddEntityMoving(X, oEntity.ObjectID)
                                End If
                            End If
                        End If
                    End If
                Next X
            End If
        End If
    End Sub

    Private Sub HandleAlertDestReached(ByRef yData() As Byte, ByVal lClientIndex As Int32)
        If lClientIndex <> -1 Then
            'If glCurrentCycle - mlClientLastRequest(lClientIndex) > ml_REQUEST_THROTTLE Then
            If glCurrentCycle - muClientData(lClientIndex).mlClientLastRequest > ml_REQUEST_THROTTLE Then
                'mlClientLastRequest(lClientIndex) = glCurrentCycle
                muClientData(lClientIndex).mlClientLastRequest = glCurrentCycle
            Else : Return
            End If
        End If

        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim yValue As Byte = yData(lPos) : lPos += 1

        If glEntityIdx Is Nothing = False Then
            Dim oEntity As Epica_Entity = Nothing
            Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
            If lIdx > -1 Then oEntity = goEntity(lIdx)
            If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return
            If oEntity.lOwnerID = mlClientPlayer(lClientIndex) OrElse (oEntity.lOwnerID = mlAliasedAs(lClientIndex) AndAlso (mlAliasedRights(lClientIndex) And AliasingRights.eMoveUnits) <> 0) Then
                moPathfinding.SendData(yData)
            End If
            'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
            'For X As Int32 = 0 To lCurUB
            '    If glEntityIdx(X) = lObjID Then
            '        Dim oEntity As Epica_Entity = goEntity(X)
            '        If oEntity Is Nothing = False Then
            '            If oEntity.ObjTypeID = iObjTypeID Then
            '                If oEntity.lOwnerID = mlClientPlayer(lClientIndex) OrElse (oEntity.lOwnerID = mlAliasedAs(lClientIndex) AndAlso (mlAliasedRights(lClientIndex) And AliasingRights.eMoveUnits) <> 0) Then
            '                    moPathfinding.SendData(yData)
            '                End If
            '                Exit For
            '            End If
            '        End If
            '    End If
            'Next X
        End If

    End Sub

    Private Sub HandleCreatePlanetInstance(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'msg id

        'Createplanetinstance is msgcode, playerid
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lInstanceID As Int32 = lPlayerID + 500000000

        If Planet.moTutorialPlanet Is Nothing Then
            Dim oPlanet As Planet = New Planet(0, 0, PlanetType.eGeoPlastic)
            With oPlanet
                .ObjTypeID = ObjectType.ePlanet
                .PlanetName = "Tutorial One"
                .PlanetRadius = 1000
                .ParentSystem = Nothing
                .LocX = 100000
                .LocY = 0
                .LocZ = 0
                .Vegetation = 0
                .Atmosphere = 100
                .Hydrosphere = 0
                .Gravity = 40
                .SurfaceTemperature = 90
                .RotationDelay = 120
                .AxisAngle = 680
                .RotateAngle = 0
                .PopulateInstanceData()
            End With
            Planet.moTutorialPlanet = oPlanet
        End If

        'Now, we need to create a new environment...
        Dim oNewInstance As New Envir()
        With oNewInstance
            .ObjectID = lInstanceID
            .ObjTypeID = ObjectType.ePlanet
            .oGeoObject = Planet.moTutorialPlanet
            .SetEnvirGridValues()
        End With

        Dim lCurUB As Int32 = -1
        Dim lIdx As Int32 = -1
        Dim bFound As Boolean = False
        If glInstanceIdx Is Nothing = False Then lCurUB = Math.Min(glInstanceUB, glInstanceIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If glInstanceIdx(X) = lInstanceID Then
                lIdx = X
                bFound = True
                Exit For
            ElseIf glInstanceIdx(X) = -1 AndAlso lIdx = -1 Then
                lIdx = X
            End If
        Next X
        If bFound = False Then
            If lIdx = -1 Then
                lIdx = glInstanceUB + 1
                ReDim Preserve goInstances(lIdx)
                ReDim Preserve glInstanceIdx(lIdx)
                glInstanceIdx(lIdx) = -1
                glInstanceUB += 1
            End If
            goInstances(lIdx) = oNewInstance
            glInstanceIdx(lIdx) = oNewInstance.ObjectID
        End If

        'now, return the instance to the primary
        Dim yResp(9) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eCreatePlanetInstance).CopyTo(yResp, 0)
        System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, 2)
        System.BitConverter.GetBytes(lInstanceID).CopyTo(yResp, 6)
        moPrimary.SendData(yResp)
    End Sub

    Private Sub HandleSaveAndUnloadInstance(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lInstanceID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'Region server sends all necessary save msgs
        'Units and Structures
        'Region server clears the objects in the instance 
        '==== Begin outbound msg ====
        Dim lSingleMsgLen As Int32
        Dim X As Int32
        Dim yTemp() As Byte
        Dim yCache(200000) As Byte
        Dim yFinal() As Byte = Nothing

        lPos = 0
        lSingleMsgLen = -1
        Dim lCurUB As Int32 = -1
        If glEntityIdx Is Nothing = False Then lCurUB = Math.Min(glEntityIdx.GetUpperBound(0), glEntityUB)
        For X = 0 To lCurUB
            If glEntityIdx(X) > 0 Then
                Dim oEntity As Epica_Entity = goEntity(X)
                If oEntity Is Nothing = False AndAlso oEntity.ParentEnvir Is Nothing = False AndAlso oEntity.ParentEnvir.ObjectID = lInstanceID AndAlso oEntity.ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
                    'Be sure the entity is disengaged...
                    With oEntity
                        If (.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
                            .CurrentStatus = oEntity.CurrentStatus Xor elUnitStatus.eUnitEngaged
                            RemoveEntityInCombat(.ObjectID, .ObjTypeID)
                        End If
                        If (.CurrentStatus And elUnitStatus.eSide1HasTarget) <> 0 Then .CurrentStatus = oEntity.CurrentStatus Xor elUnitStatus.eSide1HasTarget
                        If (.CurrentStatus And elUnitStatus.eSide2HasTarget) <> 0 Then .CurrentStatus = oEntity.CurrentStatus Xor elUnitStatus.eSide2HasTarget
                        If (.CurrentStatus And elUnitStatus.eSide3HasTarget) <> 0 Then .CurrentStatus = oEntity.CurrentStatus Xor elUnitStatus.eSide3HasTarget
                        If (.CurrentStatus And elUnitStatus.eSide4HasTarget) <> 0 Then .CurrentStatus = oEntity.CurrentStatus Xor elUnitStatus.eSide4HasTarget

                        yTemp = .GetObjectUpdateMsg(GlobalMessageCode.eUpdateEntityAndSave, 0)

                        lSingleMsgLen = yTemp.Length

                        'Ok, before we continue, check if we need to increase our cache
                        If lPos + lSingleMsgLen + 2 > yCache.Length Then
                            'increase it
                            ReDim Preserve yCache(yCache.Length + 200000)
                        End If
                        System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                        lPos += 2

                        yTemp.CopyTo(yCache, lPos)
                        lPos += lSingleMsgLen

                        .ParentEnvir.RemoveEntity(.ServerIndex, 0, False, False, -1)
                    End With
                End If
            End If
        Next X
        If lPos <> 0 Then
            ReDim yFinal(lPos - 1)
            Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
        End If
        If yFinal Is Nothing = False Then moPrimary.SendLenAppendedData(yFinal)
        '=== End of Massive Save ===

        'Region Server clears the instance itself
        lCurUB = -1
        If glInstanceIdx Is Nothing = False Then lCurUB = Math.Min(glInstanceUB, glInstanceIdx.GetUpperBound(0))
        For X = 0 To lCurUB
            If glInstanceIdx(X) = lInstanceID Then
                Dim oInstance As Envir = goInstances(X)
                If oInstance Is Nothing = False Then
                    oInstance.ClearEnvirVariables()
                End If
                glInstanceIdx(X) = -1
                goInstances(X) = Nothing
            End If
        Next X

        'Region server sends the SaveAndUnloadInstancedPlanet command back to primary, indicating it is done
        moPrimary.SendData(yData)

    End Sub

    Private Sub HandleAILaunchAll(ByRef yData() As Byte)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return
        oEntity.bInAILaunchAll = False
        'Dim lCurUB As Int32 = -1
        'If glEntityIdx Is Nothing = False Then lCurUB = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X As Int32 = 0 To lCurUB
        '    If glEntityIdx(X) = lObjID Then
        '        Dim oEntity As Epica_Entity = goEntity(X)
        '        If oEntity Is Nothing = False Then
        '            If oEntity.ObjTypeID = iObjTypeID Then
        '                oEntity.bInAILaunchAll = False
        '                Exit For
        '            End If
        '        End If
        '    End If
        'Next X
    End Sub

    Private Sub HandleRegisterDomain(ByRef yData() As Byte)
        Dim lPos As Int32 = 2

        'this is essentially an add system message, get our system GUID
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        If iTypeID <> ObjectType.eSolarSystem Then
            gfrmDisplayForm.AddEventLine("Non-Solar System Assignments Not Implemented")
            Return
        End If

        Dim oSystem As SolarSystem = Nothing
        For X As Int32 = 0 To goGalaxy.mlSystemUB
            If goGalaxy.moSystems(X) Is Nothing = False AndAlso goGalaxy.moSystems(X).ObjectID = lID Then
                oSystem = goGalaxy.moSystems(X)
                Exit For
            End If
        Next X
        If oSystem Is Nothing = False Then
            'check if we already have an environment for this system...
            For X As Int32 = 0 To glEnvirUB
                If goEnvirs(X).ObjTypeID = ObjectType.eSolarSystem AndAlso goEnvirs(X).ObjectID = oSystem.ObjectID Then
                    Return
                End If
            Next X
        Else
            oSystem = New SolarSystem()
            oSystem.ObjectID = lID
            oSystem.ObjTypeID = iTypeID
            goGalaxy.AddSystem(oSystem)
        End If
        With oSystem
            .SystemName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
            lPos += 4       'for parent galaxy
            .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .LocY = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            Dim lTemp As Int32 = yData(lPos) : lPos += 1
            For X As Int32 = 0 To glStarTypeUB
                If goStarTypes(X).StarTypeID = lTemp Then
                    .StarType1Idx = X
                    Exit For
                End If
            Next X

            lTemp = yData(lPos) : lPos += 1
            For X As Int32 = 0 To glStarTypeUB
                If goStarTypes(X).StarTypeID = lTemp Then
                    .StarType2Idx = X
                    Exit For
                End If
            Next X

            lTemp = yData(lPos) : lPos += 1
            For X As Int32 = 0 To glStarTypeUB
                If goStarTypes(X).StarTypeID = lTemp Then
                    .StarType3Idx = X
                    Exit For
                End If
            Next X
        End With

        'ok, we're here, add it to our list of systems to load...
        GeoSpawner.AddToLoadSystemQueue(lID)

    End Sub

    Private Sub HandleClientSetEntityStatus(ByRef yData() As Byte, ByVal lClientIndex As Int32)
        If lClientIndex <> -1 Then
            'If glCurrentCycle - mlClientLastRequest(lClientIndex) > ml_REQUEST_THROTTLE Then
            If glCurrentCycle - muClientData(lClientIndex).mlClientLastRequest > ml_REQUEST_THROTTLE Then
                'mlClientLastRequest(lClientIndex) = glCurrentCycle
                muClientData(lClientIndex).mlClientLastRequest = glCurrentCycle
            Else : Return
            End If
        End If
        If mlAliasedAs(lClientIndex) > -1 AndAlso mlAliasedAs(lClientIndex) <> mlClientPlayer(lClientIndex) Then Return

        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lStatus As Int32 = System.BitConverter.ToInt32(yData, 8)

        If Math.Abs(lStatus) <> elUnitStatus.eGuildAsset Then
            gfrmDisplayForm.AddEventLine("Possible Cheat: HandleClientSetEntityStatus, status is not GuildAsset. Player: " & mlClientPlayer(lClientIndex))
            Return
        End If

        Dim oEntity As Epica_Entity = Nothing
        Dim lIdx As Int32 = LookupEntity(lObjID, iObjTypeID)
        If lIdx > -1 Then oEntity = goEntity(lIdx)
        If oEntity Is Nothing OrElse oEntity.ObjectID <> lObjID OrElse oEntity.ObjTypeID <> iObjTypeID Then Return

        'Dim lCurUB As Int32 = -1
        'If glEntityIdx Is Nothing = False Then lCurUB = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'For X As Int32 = 0 To lCurUB
        '    If glEntityIdx(X) = lObjID Then
        '        Dim oEntity As Epica_Entity = goEntity(X)
        '        If oEntity Is Nothing = False Then
        '            If oEntity.ObjTypeID = iObjTypeID Then
        If oEntity.lOwnerID <> mlClientPlayer(lClientIndex) Then
            gfrmDisplayForm.AddEventLine("Possible Cheat: HandleClientSetEntityStatus, player/owner mismatch. Player: " & mlClientPlayer(lClientIndex))
            Return
        End If

        If lStatus < 0 Then
            If (oEntity.CurrentStatus And Math.Abs(lStatus)) <> 0 Then oEntity.CurrentStatus = oEntity.CurrentStatus Xor Math.Abs(lStatus)
        Else
            oEntity.CurrentStatus = oEntity.CurrentStatus Or lStatus
        End If

        moPrimary.SendData(yData)
        BroadcastToEnvironmentClients(yData, oEntity.ParentEnvir)

        '                Exit For
        '            End If
        '        End If
        '    End If
        'Next X
    End Sub

    Private Sub HandlePlayerIsDead(ByVal yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lCurUB As Int32 = -1
        If glPlayerIdx Is Nothing = False Then lCurUB = Math.Min(glPlayerUB, glPlayerIdx.GetUpperBound(0))
        Dim oPlayer As Player = Nothing
        For X As Int32 = 0 To lCurUB
            If glPlayerIdx(X) = lPlayerID Then
                oPlayer = goPlayers(X)
                Exit For
            End If
        Next X
        If oPlayer Is Nothing = False Then
            oPlayer.BadWarDecCPIncrease = 0
            oPlayer.RemovePlayerRel(-1, False)
        End If

        lCurUB = -1
        If glEntityIdx Is Nothing = False Then lCurUB = Math.Min(glEntityIdx.GetUpperBound(0), glEntityUB)
        For X As Int32 = 0 To lCurUB
            If glEntityIdx(X) > 0 Then
                Dim oEntity As Epica_Entity = goEntity(X)
                If oEntity Is Nothing = False Then
                    If oEntity.lOwnerID = lPlayerID Then
                        'remove it
                        If oEntity.ParentEnvir Is Nothing = False Then oEntity.ParentEnvir.RemoveEntity(oEntity.ServerIndex, RemovalType.eObjectDestroyed, False, False, -1)
                    End If
                End If
            End If
        Next X

        For X As Int32 = 0 To glEnvirUB
            Dim oEnvir As Envir = goEnvirs(X)
            If oEnvir Is Nothing = False Then
                For Y As Int32 = 0 To oEnvir.lPlayersWhoHaveUnitsHereUB
                    If oEnvir.lPlayersWhoHaveUnitsHereIdx(Y) = lPlayerID Then
                        oEnvir.lPlayersWhoHaveUnitsHereCP(Y) = 0
                        oEnvir.lPlayersWhoHaveUnitsHereIdx(Y) = -1
                        Exit For
                    End If
                Next Y
            End If
        Next X

        BroadcastToClients(yData)
    End Sub
#End Region
 
End Class
