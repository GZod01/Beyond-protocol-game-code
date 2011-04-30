Option Strict On

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

    'Pathfinding Server connects to the Primary server
    Private WithEvents moPrimary As NetSock

    'Pathfinding Server receives connections from the Domains
    Private WithEvents moDomainListener As NetSock
    Private moDomains() As NetSock
    Private mlDomainUB As Int32 = -1

    'Pathfinding Server connects to the Operator server
    Private WithEvents moOperator As NetSock

    Private mbAcceptingDomains As Boolean = False
    Private mbConnectedToPrimary As Boolean = False

    Private mbSyncWait As Boolean = False
    Private mlTimeout As Int32

    Public bHavePrimaryConnInfo As Boolean = False

    Public Sub New()
        'do our variable initialization here
        Dim oINI As InitFile = New InitFile()
        mlTimeout = CInt(Val(oINI.GetString("SETTINGS", "ConnectTimeout", "15000")))
        oINI = Nothing
    End Sub

    'Clients do not communicate directly to the pathfinding server...
    '  Clients send messages to the Domain Servers and the Domain Server forwards the message
    '  to the Pathfinding Server for processing. This is to avoid a Network bottleneck here
    Public Property AcceptingDomains() As Boolean
        Get
            Return mbAcceptingDomains
        End Get
        Set(ByVal Value As Boolean)
            If mbAcceptingDomains <> Value Then
                mbAcceptingDomains = Value

                If Value = True Then
                    'Ok, we are now accepting them
                    Dim oINI As New InitFile()
                    Dim lPort As Int32

                    Try
                        lPort = CInt(Val(oINI.GetString("SETTINGS", "DomainListenPort", "0")))
                        If lPort = 0 Then Err.Raise(-1, "AcceptingDomains", "Could not get Domain Listen Port Number from INI File.")
                        moDomainListener = Nothing
                        moDomainListener = New NetSock()
                        moDomainListener.PortNumber = lPort
                        moDomainListener.Listen()
                    Catch
                        MsgBox("An error was caught in " & Err.Source & ":" & vbCrLf & Err.Description, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, Err.Source)
                        mbAcceptingDomains = False
                    Finally
                        oINI = Nothing
                    End Try
                Else
                    'ok, we are no longer accepting them
                    moDomainListener.StopListening()
                End If
            End If
        End Set
    End Property

    Public Function ConnectToPrimary() As Boolean
        Dim sIP As String
        Dim lPort As Int32
        Dim oINI As InitFile
        Dim lTimeout As Int32

        Try
            oINI = New InitFile()
            lTimeout = CInt(Val(oINI.GetString("SETTINGS", "ConnectTimeout", "15000")))

            lPort = CInt(Val(oINI.GetString("PRIMARY", "PortNumber", "0")))
            If lPort = 0 Then Err.Raise(-1, "ConnectToPrimary", "Unable to get Primary Port Number from INI.")
            sIP = oINI.GetString("PRIMARY", "IPAddress", "")
            If sIP = "" Then Err.Raise(-1, "ConnectToPrimary", "Unable to get Primary IP Address from INI.")
            moPrimary = New NetSock()
            moPrimary.Connect(sIP, lPort)

            Dim sw As Stopwatch = Stopwatch.StartNew
            While mbConnectedToPrimary = False AndAlso (sw.ElapsedMilliseconds < lTimeout)
                Application.DoEvents()
            End While
            sw.Stop()
            sw = Nothing

            If mbConnectedToPrimary = False Then
                Err.Raise(-1, "ConnectToPrimary", "Connection to Primary timed out.")
            End If
        Catch
            MsgBox("An error occurred while attempting to connect to the Primary Server." & vbCrLf & vbCrLf & _
                Err.Source & " (" & Err.Number & "): " & Err.Description, MsgBoxStyle.Critical Or MsgBoxStyle.OkOnly, "Error")
        Finally
            oINI = Nothing
        End Try

        Return mbConnectedToPrimary
    End Function

    Public Function SendPrimaryMyInfo() As Boolean
        Dim yData() As Byte
        Dim sIPBytes() As String
        Dim sTemp As String
        Dim ipEntry As System.Net.IPHostEntry
        Dim ipAddr As System.Net.IPAddress

        Dim lPort As Int32
        Dim oINI As New InitFile()
        Dim bRes As Boolean = False

        'Message will be Message Code(2), IPByte 1-4(4), Port(4)
        gfrmDisplayForm.AddEventLine("Sending Connection Info")

        lPort = CInt(Val(oINI.GetString("SETTINGS", "DomainListenPort", "0")))
        If lPort <> 0 Then

            ipEntry = System.Net.Dns.GetHostByName(Environment.MachineName)
            ipAddr = ipEntry.AddressList(0)
            sTemp = ipAddr.ToString()
            'stemp now is XXX.XXX.XXX.XXX
            sIPBytes = Split(sTemp, ".")
            If sIPBytes.GetUpperBound(0) = 3 Then
                ReDim yData(9)      '0 to 9 = 10 bytes
                System.BitConverter.GetBytes(GlobalMessageCode.ePathfindingConnectionInfo).CopyTo(yData, 0)
                yData(2) = CByte(Val(sIPBytes(0)))
                yData(3) = CByte(Val(sIPBytes(1)))
                yData(4) = CByte(Val(sIPBytes(2)))
                yData(5) = CByte(Val(sIPBytes(3)))
                System.BitConverter.GetBytes(lPort).CopyTo(yData, 6)

                moPrimary.SendData(yData)
                bRes = True
            End If
        End If

        ipAddr = Nothing
        ipEntry = Nothing
        oINI = Nothing
        Return bRes

    End Function

    Public Sub SendToPrimary(ByVal yData() As Byte)
        moPrimary.SendData(yData)
    End Sub

#Region "Primary Server Events - Pathfinding Connects to Primary"
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
        Dim iMsgID As Int16 = System.BitConverter.ToInt16(Data, 0)
        Select Case iMsgID
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
                mbSyncWait = False
            Case GlobalMessageCode.eServerShutdown
                gfrmDisplayForm.AddEventLine("Server Shutting Down")
                gfrmDisplayForm.Close()
            Case GlobalMessageCode.ePFRequestEntitys
                ReceiveEntities(Data)
            Case GlobalMessageCode.eAddObjectCommand
                ReceiveAddObject(Data)
            Case GlobalMessageCode.eRemoveObject
				'gfrmDisplayForm.AddEventLine("Remove Object Received")
				HandleRemoveObject(Data)
			Case GlobalMessageCode.eAddFormation
				HandleAddFormation(Data)
		End Select
    End Sub

    Private Sub moPrimary_onDisconnect(ByVal Index As Integer) Handles moPrimary.onDisconnect
        'Index is required for backwards compatibility, on the Server, we do not use it
    End Sub

    Private Sub moPrimary_onError(ByVal Index As Integer, ByVal Description As String) Handles moPrimary.onError
        'Index is required for backwards compatibility, on the Server, we do not use it
    End Sub

    Private Sub moPrimary_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moPrimary.onSendComplete
        'Index is required for backwards compatibility, on the Server, we do not use it
    End Sub
#End Region

#Region "Domain Listener - The Listening Socket for Domain Connections"
    Private Sub moDomainListener_onConnect(ByVal Index As Integer) Handles moDomainListener.onConnect
        'Index is required for backwards compatibility, on the Server, we do not use it
    End Sub

    Private Sub moDomainListener_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moDomainListener.onConnectionRequest
        'Index is required for backwards compatibility, on the Server, we do not use it
        Dim X As Int32
        Dim lIdx As Int32

        'let's find an unused socket first
        lIdx = -1
        For X = 0 To mlDomainUB
            If moDomains(X) Is Nothing Then
                'found one, use this one
                lIdx = X
                Exit For
            End If
        Next X

        If lIdx = -1 Then
            'Ok, need a new one
            mlDomainUB += 1
            ReDim Preserve moDomains(mlDomainUB)
            lIdx = mlDomainUB
        End If

        'Now, set up the socket
        moDomains(lIdx) = New NetSock(oClient)
        moDomains(lIdx).SocketIndex = lIdx

        gfrmDisplayForm.AddEventLine("Domain " & lIdx & " Connected")

        'Now, set up our events and we're done
        AddHandler moDomains(lIdx).onConnect, AddressOf moDomains_onConnect
        AddHandler moDomains(lIdx).onDataArrival, AddressOf moDomains_onDataArrival
        AddHandler moDomains(lIdx).onDisconnect, AddressOf moDomains_onDisconnect
        AddHandler moDomains(lIdx).onError, AddressOf moDomains_onError

        moDomains(lIdx).MakeReadyToReceive()

    End Sub

    Private Sub moDomainListener_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moDomainListener.onDataArrival
        'Index is required for backwards compatibility, on the Server, we do not use it
        'This should never happen, the Server Socket is our Listening socket... it should not receive Data
    End Sub

    Private Sub moDomainListener_onDisconnect(ByVal Index As Integer) Handles moDomainListener.onDisconnect
        'Index is required for backwards compatibility, on the Server, we do not use it
        ' shouldn't really happen I guess... ignore
    End Sub

    Private Sub moDomainListener_onError(ByVal Index As Integer, ByVal Description As String) Handles moDomainListener.onError
        'Index is required for backwards compatibility, on the Server, we do not use it
    End Sub

    Private Sub moDomainListener_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moDomainListener.onSendComplete
        'Index is required for backwards compatibility, on the server, we do not use it
        'Not really interested in this event
    End Sub
#End Region

#Region "Server Socket Array for the Domain Connections (moDomains)"
    Private Sub moDomains_onConnect(ByVal Index As Integer)
        'Don't really care
    End Sub

    Private Sub moDomains_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer)
        'Data() is our byte array, let's parse it out...
        Dim iMsgID As Int16
        Dim yResponse() As Byte
        Dim X As Int32

        Dim lObjID As Int32
        Dim iObjTypeID As Int16

        iMsgID = System.BitConverter.ToInt16(Data, 0)
        'iMsgID has our msg code
		Select Case iMsgID
			Case GlobalMessageCode.eRouteMoveCommand
				HandleRouteMove(Data, Index)
			Case GlobalMessageCode.eMoveObjectRequest
				HandleGroupMoveRequest(Data, Index, False)
			Case GlobalMessageCode.eAddWaypointMsg
				HandleGroupMoveRequest(Data, Index, True)
			Case GlobalMessageCode.eStopObjectCommand
				'Stop Object Command is:
				' MessageCode(2), ObjectID(4), ObjTypeID(2), LocX(4), LocZ(4), LocAngle(2)
				'gfrmDisplayForm.AddEventLine("Stop Object Command")
				lObjID = System.BitConverter.ToInt32(Data, 2)
                iObjTypeID = System.BitConverter.ToInt16(Data, 6)
                Dim lCurUB As Int32 = Math.Min(glMoverUB, glMoverIdx.GetUpperBound(0))
                For X = 0 To lCurUB
                    If glMoverIdx(X) = lObjID Then
                        If goMovers(X) Is Nothing = False AndAlso goMovers(X).ObjTypeID = iObjTypeID Then
                            yResponse = goMovers(X).ProcessStopMessage(System.BitConverter.ToInt32(Data, 8), _
                             System.BitConverter.ToInt32(Data, 12), System.BitConverter.ToInt16(Data, 16))
                            If yResponse Is Nothing = False AndAlso yResponse.Length > 0 Then
                                moDomains(Index).SendData(yResponse)
                            Else
                                If Mover.FinalizeStopEvent(goMovers(X)) = True Then
                                    'We send an update of the entity's loc to the primary server for asynchronous persistence
                                    Dim yTemp(17) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.eUpdateEntityLoc).CopyTo(yTemp, 0)
                                    goMovers(X).GetGUIDAsString.CopyTo(yTemp, 2)
                                    System.BitConverter.GetBytes(goMovers(X).lCurLocX).CopyTo(yTemp, 8)
                                    System.BitConverter.GetBytes(goMovers(X).lCurLocZ).CopyTo(yTemp, 12)
                                    System.BitConverter.GetBytes(goMovers(X).lCurAngle).CopyTo(yTemp, 16)
                                    goMsgSys.SendToPrimary(yTemp)

                                    goMovers(X).bMoving = False
                                Else
                                    'get the next dest in the list
                                    Dim uDest As DestLoc = CType(goMovers(X).colDests(1), DestLoc)
                                    'remove that dest from the list
                                    goMovers(X).colDests.Remove(1)

                                    Dim yTemp(17) As Byte     '0 to 17 = 18 bytes
                                    System.BitConverter.GetBytes(GlobalMessageCode.eFinalMoveCommand).CopyTo(yTemp, 0)
                                    goMovers(X).GetGUIDAsString.CopyTo(yTemp, 2)
                                    System.BitConverter.GetBytes(uDest.DestX).CopyTo(yTemp, 8)
                                    System.BitConverter.GetBytes(uDest.DestZ).CopyTo(yTemp, 12)
                                    System.BitConverter.GetBytes(uDest.DestAngle).CopyTo(yTemp, 16)
                                    goMovers(X).bMoving = True
                                    goMovers(X).lCurDestX = uDest.DestX
                                    goMovers(X).lCurDestZ = uDest.DestZ
                                    moDomains(Index).SendData(yTemp)
                                End If
                            End If
                            Exit For
                        End If
                    End If
                Next X
			Case GlobalMessageCode.eDecelerationImminent
				HandleDecelerationImminentMsg(Data, moDomains(Index))
			Case GlobalMessageCode.eEntityChangeEnvironment
				HandleEntityChangeEnvironmentMsg(Data)
			Case GlobalMessageCode.eRequestDock
				HandleDockRequest(Data, moDomains(Index))
			Case GlobalMessageCode.eSetMiningLoc
				HandleSetMiningLoc(Data, Index)
			Case GlobalMessageCode.eSetEntityProduction
				HandleSetEntityProduction(Data, Index)
			Case GlobalMessageCode.eSetFleetDest
				HandleSetFleetDest(Data, Index)
			Case GlobalMessageCode.eSetDismantleTarget, GlobalMessageCode.eSetRepairTarget
				HandleSetMaintenanceTarget(Data, Index, iMsgID)
			Case GlobalMessageCode.eFinalizeStopEvent
				HandleFinalizeStopEvent(Data, Index)
			Case GlobalMessageCode.eJumpTarget
				HandleJumpTarget(Data, Index)
			Case GlobalMessageCode.eSetPrimaryTarget
				HandleSetPrimaryTarget(Data, Index)
			Case GlobalMessageCode.eMoveFormation
				HandleMoveFormation(Data, Index)
			Case GlobalMessageCode.eAlertDestinationReached
				HandleAlertDestReached(Data, Index)
			Case GlobalMessageCode.eClearDestList
				HandleClearDestList(Data)
			Case GlobalMessageCode.eForcedMoveSpeedMove
                HandleForcedMoveSpeedMove(Data, Index)
		End Select

    End Sub

    Private Sub moDomains_onDisconnect(ByVal Index As Integer)
        On Error Resume Next
        moDomains(Index).Disconnect()   'to ensure it is disconnected
        moDomains(Index) = Nothing
        gfrmDisplayForm.AddEventLine("Domain " & Index & " disconnected")
    End Sub

    Private Sub moDomains_onError(ByVal Index As Integer, ByVal Description As String)
        gfrmDisplayForm.AddEventLine("Domain Connection Error (" & Index & "): " & Description)
    End Sub
#End Region

#Region "  Operator Events  "
    Private mbConnectingToOperator As Boolean = False
    Private mbConnectedToOperator As Boolean = False
    Public Function ConnectToOperator() As Boolean

        Try
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
                System.BitConverter.GetBytes(ConnectionType.ePathfindingServerApp).CopyTo(yMsg, lPos) : lPos += 4

                Dim oMyProcess As System.Diagnostics.Process = System.Diagnostics.Process.GetCurrentProcess()
                Dim lProcessID As Int32 = oMyProcess.Id
                System.BitConverter.GetBytes(lProcessID).CopyTo(yMsg, lPos) : lPos += 4
                oMyProcess = Nothing

                System.BitConverter.GetBytes(1I).CopyTo(yMsg, lPos) : lPos += 4     'indicates 1 connection specifics

                'Get our Port number data
                Dim oIni As New InitFile()
                Dim lRegionPort As Int32 = CInt(Val(oIni.GetString("SETTINGS", "DomainListenPort", "0")))
                If lRegionPort = 0 Then Err.Raise(-1, "moOperator.onConnect", "moOperator.onConnect: Unable to determine Region Listen Port from INI")
                oIni = Nothing

                'we'll indicate our region listener... so use our local ip address
                StringToBytes(Mid$(.Address.ToString(), 1, 20)).CopyTo(yMsg, lPos) : lPos += 20
                System.BitConverter.GetBytes(lRegionPort).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(ConnectionType.eRegionServerApp).CopyTo(yMsg, lPos) : lPos += 4

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

        Select Case iMsgID
            Case GlobalMessageCode.ePathfindingConnectionInfo
                HandlePrimaryConnInfo(Data)
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

    Public Sub SendReadyStateToOperator()
        Dim yMsg(1) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eDomainServerReady).CopyTo(yMsg, 0)
        moOperator.SendData(yMsg)
    End Sub

    Private Sub HandlePrimaryConnInfo(ByRef yData() As Byte)
        'this contains the primary server's connection data
        Dim lPos As Int32 = 2       'for msgcode
        Dim sIP As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
        Dim lPort As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim oINI As New InitFile()

        oINI.WriteString("PRIMARY", "PortNumber", lPort.ToString)
        oINI.WriteString("PRIMARY", "IPAddress", sIP)
        oINI = Nothing

        gfrmDisplayForm.AddEventLine("Primary Connection Info Received: " & sIP & ":" & lPort)

        bHavePrimaryConnInfo = True
    End Sub
#End Region


    Public Sub CloseAllConnections()
        Dim X As Int32

        AcceptingDomains = False

        For X = 0 To mlDomainUB
            If moDomains(X) Is Nothing = False Then
                moDomains(X).Disconnect()
            End If
            moDomains(X) = Nothing
        Next X
        moPrimary.Disconnect()
        moPrimary = Nothing
    End Sub

    Public Function GetStarTypes() As Boolean
        Dim yRequest(2) As Byte
        Dim sw As Stopwatch

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestStarTypes).CopyTo(yRequest, 0)

        mbSyncWait = True
        sw = Stopwatch.StartNew()
        moPrimary.SendData(yRequest)
        While mbSyncWait = True And (sw.ElapsedMilliseconds < mlTimeout)
            Application.DoEvents()
        End While
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

        'set up our galaxy
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

    Public Function GetGalaxyAndSystems() As Boolean
        Dim yRequest(2) As Byte
        Dim sw As Stopwatch

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestGalaxyAndSystems).CopyTo(yRequest, 0)

        mbSyncWait = True
        sw = Stopwatch.StartNew
        moPrimary.SendData(yRequest)
        While mbSyncWait = True And (sw.ElapsedMilliseconds < mlTimeout)
            Application.DoEvents()
        End While
        If mbSyncWait = True Then
            mbSyncWait = False
            Return False
        Else
            Return True
        End If

	End Function

	Private Sub HandleSystemDetailsMsg(ByRef yData() As Byte)
		Dim lPos As Int32 = 2		'for msgcode
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

			glPlanetVPUB += 1
			ReDim Preserve goPlanetVP(glPlanetVPUB)
			goPlanetVP(glPlanetVPUB).lID = oTmpPlnt.ObjectID
			goPlanetVP(glPlanetVPUB).lParentID = oTmpPlnt.ParentSystem.ObjectID
			goPlanetVP(glPlanetVPUB).oPlanetRef = oTmpPlnt

			oSystem.AddPlanet(oTmpPlnt)
			oTmpPlnt = Nothing

		Next X
	End Sub

	'Private Sub HandleSystemDetailsMsg(ByVal yData() As Byte)
	'    Dim lPos As Int32 = 2   'msg id
	'    Dim oTmpPlnt As Planet
	'    Dim lID As Int32
	'    Dim iTypeID As Int16
	'    Dim sName As String
	'    Dim ySizeID As Byte
	'    Dim yMapTypeID As Byte
	'    Dim yTempName(19) As Byte

	'    Dim lInnerRadius As Int32
	'    Dim lOuterRadius As Int32
	'    Dim lRingDiffuse As Int32

	'    'ok, got our details, this is nothing more than a list of Planet objects... we would only get this message
	'    '  for our galaxy's current system
	'    If goGalaxy.CurrentSystemIdx <> -1 Then
	'        While lPos < yData.Length - 1


	'            'Object ID
	'            lID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
	'            'ObjTypeID
	'            iTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
	'            'Name
	'            Array.Copy(yData, lPos, yTempName, 0, 20)
	'            lPos += 20
	'            sName = System.Text.ASCIIEncoding.ASCII.GetString(yTempName)
	'            'Planet type id
	'            yMapTypeID = yData(lPos) : lPos += 1
	'            'Size ID
	'            ySizeID = yData(lPos) : lPos += 1

	'            'Now, create our planet
	'            oTmpPlnt = New Planet(lID, ySizeID, yMapTypeID)

	'            With oTmpPlnt
	'                .ObjTypeID = iTypeID
	'                .PlanetName = sName
	'                'ok, now the other attributes
	'                .PlanetRadius = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
	'                .ParentSystem = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx) : lPos += 4
	'                .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
	'                .LocY = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
	'                .LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
	'                .Vegetation = yData(lPos) : lPos += 1
	'                .Atmosphere = yData(lPos) : lPos += 1
	'                .Hydrosphere = yData(lPos) : lPos += 1
	'                .Gravity = yData(lPos) : lPos += 1
	'                .SurfaceTemperature = yData(lPos) : lPos += 1
	'                .RotationDelay = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

	'                'Now, check for our ring...
	'                lInnerRadius = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
	'                lOuterRadius = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
	'                lRingDiffuse = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

	'                'finally, the axis angle
	'                .AxisAngle = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

	'                .RotateAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

	'                .PopulateData()
	'            End With

	'            glPlanetVPUB += 1
	'            ReDim Preserve goPlanetVP(glPlanetVPUB)
	'            goPlanetVP(glPlanetVPUB).lID = oTmpPlnt.ObjectID
	'            goPlanetVP(glPlanetVPUB).lParentID = oTmpPlnt.ParentSystem.ObjectID
	'            goPlanetVP(glPlanetVPUB).oPlanetRef = oTmpPlnt

	'            goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).AddPlanet(oTmpPlnt)
	'            oTmpPlnt = Nothing
	'        End While
	'    End If
	'End Sub

    Public Function GetSystemDetails(ByVal lID As Int32) As Boolean
        Dim sw As Stopwatch
        Dim yData(5) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eRequestSystemDetails).CopyTo(yData, 0)
        System.BitConverter.GetBytes(lID).CopyTo(yData, 2)

        mbSyncWait = True
        sw = Stopwatch.StartNew
        moPrimary.SendData(yData)
        While mbSyncWait = True And (sw.ElapsedMilliseconds < mlTimeout)
            Application.DoEvents()
        End While
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

#Region " Old Single Entity Message Handling (Obsolete) "
    'Private Function HandleAddWaypointMsg(ByVal yData() As Byte) As Byte()
    '    'Ok, got an AddWaypoint msg... this acts just like a MoveRequest msg... but we add it to the end of any
    '    '  dest list that may already be created for this object
    '    'process the request... the Domain Server should pass us the original request:
    '    ' MessageCode(2), ObjectID(4), ObjTypeID(2), DestX(4), DestZ(4), DestAngle(2)
    '    'with the appended location data:
    '    ' CurLocX(4), CurLocZ(4), CurAngle(2)
    '    'added: CurrentEnvirGUID (6), DestEnvirGUID (6)
    '    Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
    '    Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

    '    Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, 8)
    '    Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, 12)
    '    Dim iDestAngle As Int16 = System.BitConverter.ToInt16(yData, 16)
    '    Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, 18)
    '    Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, 22)

    '    Dim lCurX As Int32 = System.BitConverter.ToInt32(yData, 24)
    '    Dim lCurZ As Int32 = System.BitConverter.ToInt32(yData, 28)
    '    Dim iCurAngle As Int16 = System.BitConverter.ToInt16(yData, 32)
    '    Dim lCurID As Int32 = System.BitConverter.ToInt32(yData, 34)
    '    Dim iCurTypeID As Int16 = System.BitConverter.ToInt16(yData, 38)

    '    Return ExecuteAddWaypointRequest(lObjID, iObjTypeID, lCurX, lCurZ, iCurAngle, lCurID, iCurTypeID, _
    '      lDestX, lDestZ, iDestAngle, lDestID, iDestTypeID)

    'End Function

    'Private Function HandleMoveRequestMsg(ByVal yData() As Byte) As Byte()
    '    'process the request... the Domain Server should pass us the original request:
    '    ' MessageCode(2), ObjectID(4), ObjTypeID(2), DestX(4), DestZ(4), DestAngle(2)
    '    'with the appended location data:
    '    ' CurLocX(4), CurLocZ(4), CurAngle(2)
    '    'added: CurrentEnvirGUID (6), DestEnvirGUID (6)

    '    Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
    '    Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

    '    Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, 8)
    '    Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, 12)
    '    Dim iDestAngle As Int16 = System.BitConverter.ToInt16(yData, 16)
    '    Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, 18)
    '    Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, 22)

    '    Dim lCurX As Int32 = System.BitConverter.ToInt32(yData, 24)
    '    Dim lCurZ As Int32 = System.BitConverter.ToInt32(yData, 28)
    '    Dim iCurAngle As Int16 = System.BitConverter.ToInt16(yData, 32)
    '    'Because we are replacing our dest list, we use these values regardless
    '    Dim lCurID As Int32 = System.BitConverter.ToInt32(yData, 34)
    '    Dim iCurTypeID As Int16 = System.BitConverter.ToInt16(yData, 38)

    '    Return ExecuteMoveRequest(lObjID, iObjTypeID, lCurX, lCurZ, iCurAngle, lCurID, iCurTypeID, lDestX, lDestZ, _
    '      iDestAngle, lDestID, iDestTypeID)

    'End Function
#End Region

    Private Sub HandleDecelerationImminentMsg(ByVal yData() As Byte, ByRef oSocket As NetSock)
        'DI msg is a ObjID (4), ObjTypeID (2), DestX (4), DestZ (4), DestAngle (2), CurLocX (4), CurLocZ (4), CurAngle (2) 
        '  UnitManuever (1)
        'we take this data and we calculate if there are any obstacles in the way...

        'TODO: UnitManeuver is passed down to ensure that we have reached the correct loc... we should probably send down the speed too
        '  that way we can determine the correct time to begin our new dest...

        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim X As Int32
        Dim yResp() As Byte

        Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, 12)
        Dim iDestAngle As Int16 = System.BitConverter.ToInt16(yData, 16)

        Dim lCurUB As Int32 = Math.Min(glMoverUB, glMoverIdx.GetUpperBound(0))
        For X = 0 To lCurUB
            If glMoverIdx(X) = lObjID Then
                If goMovers(X) Is Nothing Then
                    glMoverIdx(X) = -1
                    Continue For
                End If
                If goMovers(X).ObjTypeID = iObjTypeID Then
                    yResp = goMovers(X).ProcessDecelerateImminentMsg(lDestX, lDestZ, iDestAngle)
                    If yResp Is Nothing = False AndAlso yResp.Length > 0 Then oSocket.SendData(yResp)
                    Exit For
                End If
            End If
        Next X

    End Sub

    Private Sub HandleEntityChangeEnvironmentMsg(ByVal yData() As Byte)
        'This message will be the ObjectID and ObjectTypeID...
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim X As Int32
        Dim lIdx As Int32 = -1
        Dim yResp() As Byte
        Dim uDest As DestLoc

        Dim lCurUB As Int32 = Math.Min(glMoverUB, glMoverIdx.GetUpperBound(0))
        For X = 0 To lCurUB
            If glMoverIdx(X) = lObjID AndAlso goMovers(X) Is Nothing = False Then
                If goMovers(X).ObjTypeID = iObjTypeID Then
                    lIdx = X
                    Exit For
                End If
            End If
        Next X

        If lIdx = -1 Then
            'ok, just plain bad... we'll ignore it for now... but do we do when this happens?
            'TODO: what happens here?
            gfrmDisplayForm.AddEventLine("HandleEntityChangeEnvirMsg but could not find Mover!")
        Else
            If goMovers(lIdx).colDests.Count > 0 Then
                'get the next dest and remove it from the list...
                uDest = CType(goMovers(lIdx).colDests(1), DestLoc)
                goMovers(lIdx).colDests.Remove(1)

                ReDim yResp(24)

                'Now, send it as a special message...
                System.BitConverter.GetBytes(GlobalMessageCode.eEntityChangeEnvironment).CopyTo(yResp, 0)
                goMovers(X).GetGUIDAsString.CopyTo(yResp, 2)

                With uDest
                    System.BitConverter.GetBytes(.DestX).CopyTo(yResp, 8)
                    System.BitConverter.GetBytes(.DestZ).CopyTo(yResp, 12)
                    System.BitConverter.GetBytes(.DestAngle).CopyTo(yResp, 16)
                    System.BitConverter.GetBytes(.lDestID).CopyTo(yResp, 18)
                    System.BitConverter.GetBytes(.iDestTypeID).CopyTo(yResp, 22)
                    yResp(24) = .yChangeEnvironment     'should always be no change... but...
                End With

                moPrimary.SendData(yResp)
            End If
        End If
    End Sub

    'Private Sub HandleDockRequest(ByVal yData() As Byte, ByRef oSocket As NetSock)
    '    Dim lDockerID As Int32 = System.BitConverter.ToInt32(yData, 2)
    '    Dim iDockerTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
    '    Dim lDockeeID As Int32 = System.BitConverter.ToInt32(yData, 8)
    '    Dim iDockeeTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)


    '    Dim lIdx As Int32 = -1

    '    Dim lCurX As Int32 = System.BitConverter.ToInt32(yData, 14)
    '    Dim lCurZ As Int32 = System.BitConverter.ToInt32(yData, 18)
    '    Dim iCurAngle As Int16 = System.BitConverter.ToInt16(yData, 22)
    '    Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, 24)
    '    Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, 28)
    '    Dim iDestAngle As Int16 = System.BitConverter.ToInt16(yData, 32)

    '    Dim lCurID As Int32 = System.BitConverter.ToInt32(yData, 34)
    '    Dim iCurTypeID As Int16 = System.BitConverter.ToInt16(yData, 38)

    '    Dim colNewDests As Collection

    '    Dim X As Int32

    '    Dim fAngle As Single = LineAngleDegrees(lDestX, lDestZ, lCurX, lCurZ)
    '    Dim fX As Single = lDestX + 300
    '    Dim fZ As Single = lDestZ

    '    Dim yResp() As Byte

    '    RotatePoint(lDestX, lDestZ, fX, fZ, fAngle)

    '    colNewDests = GetDestList(lCurX, lCurZ, iCurAngle, lCurID, iCurTypeID, fX, fZ, iDestAngle, lCurID, iCurTypeID)

    '    If colNewDests Is Nothing Then
    '        'return invalid dest
    '        System.BitConverter.GetBytes(EpicaMessageCode.eMoveObjectRequestDeny).CopyTo(yData, 0)
    '        oSocket.SendData(yData)
    '    Else
    '        For X = 0 To glMoverUB
    '            If glMoverIdx(X) = lDockerID Then
    '                If goMovers(X).ObjTypeID = iDockerTypeID Then
    '                    lIdx = X

    '                    'Ok, clear our current list
    '                    goMovers(X).colDests = Nothing
    '                    Exit For
    '                End If
    '            End If
    '        Next X

    '        If lIdx = -1 Then
    '            glMoverUB += 1
    '            ReDim Preserve goMovers(glMoverUB)
    '            ReDim Preserve glMoverIdx(glMoverUB)
    '            lIdx = glMoverUB
    '            goMovers(lIdx) = New Mover()
    '            glMoverIdx(lIdx) = lDockerID
    '            With goMovers(lIdx)
    '                .lCurAngle = iCurAngle
    '                .lCurLocX = lCurX
    '                .lCurLocZ = lCurZ
    '                .lCurLocY = 0
    '                .ObjectID = lDockerID
    '                .ObjTypeID = iDockerTypeID
    '            End With
    '        End If

    '        'Now, finally, do our move to dest in the final envir
    '        Dim uDest As DestLoc
    '        uDest.DestAngle = fAngle * 10
    '        uDest.DestX = lDestX
    '        uDest.DestZ = lDestZ
    '        uDest.iDestTypeID = iDockeeTypeID
    '        uDest.lDestID = lDockeeID
    '        uDest.yChangeEnvironment = ChangeEnvironmentType.eDocking
    '        uDest.ySpecialOp = SpecialOp.eNoSpecialOp
    '        colNewDests.Add(uDest)

    '        'Add the exact same dest again..
    '        colNewDests.Add(uDest)


    '        goMovers(lIdx).colDests = colNewDests

    '        'Now, we kinda cheat here...
    '        yResp = goMovers(lIdx).ProcessStopMessage(lCurX, lCurZ, iCurAngle)

    '        If yResp.Length > 0 Then
    '            oSocket.SendData(yResp)
    '        Else
    '            'Ok, send invalid move
    '            goMovers(lIdx).bMoving = False
    '            System.BitConverter.GetBytes(EpicaMessageCode.eMoveObjectRequestDeny).CopyTo(yData, 0)
    '            oSocket.SendData(yData)
    '        End If
    '    End If
    'End Sub

    Private Sub HandleDockRequest(ByVal yData() As Byte, ByRef oSocket As NetSock)
        Dim lPos As Int32 = 2
        Dim lDockeeID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDockeeTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDestA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim lObjID As Int32
        Dim iObjTypeID As Int16
        Dim lLocX As Int32
        Dim lLocZ As Int32
        Dim iLocA As Int16
        Dim lParentID As Int32
        Dim iParentTypeID As Int16
        Dim iModelID As Int16

        'Dim X As Int32

        Dim yResp() As Byte

        Dim yFinal() As Byte = Nothing
		'        Dim lLen As Int32
        Dim lDestPos As Int32

        Dim fAngle As Single
        Dim lX As Int32
        Dim lZ As Int32

        Dim uDockOp As DestLoc

        While lPos + 24 < yData.Length '- 1
            lObjID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            lLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            lLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iLocA = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            lParentID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iParentTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            iModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            'MSC: PF only cares about the model portion of the modelid
            iModelID = (iModelID And 255S)

            fAngle = LineAngleDegrees(lDestX, lDestZ, lLocX, lLocZ)
            lX = lDestX + 300       'TODO: where does 300 come from???
            lZ = lDestZ

            RotatePoint(lDestX, lDestZ, lX, lZ, fAngle)

            yResp = ExecuteMoveRequest(lObjID, iObjTypeID, lLocX, lLocZ, iLocA, lParentID, iParentTypeID, _
              lX, lZ, iDestA, lDestID, iDestTypeID, iModelID)

            If yResp Is Nothing = False AndAlso yResp.Length > 0 Then
                'now, add our final dest
                With uDockOp
                    .DestAngle = CShort(fAngle * 10)
                    .DestX = lDestX
                    .DestZ = lDestZ
                    .iDestTypeID = iDockeeTypeID
                    .lDestID = lDockeeID
                    .yChangeEnvironment = ChangeEnvironmentType.eDocking
                    .ySpecialOp = SpecialOp.eNoSpecialOp
                    .lSpecialOpID = 0
                    .iSpecialOpTypeID = 0
                End With

                'unfortunately, gotta find this object again...
                'TODO: remove the redundant search, possibly expand the Execute routines
                Dim lMoverIdx As Int32 = GetOrAddMover(lObjID, iObjTypeID, lLocX, lLocZ, iLocA, iModelID, lParentID, iParentTypeID)
                If lMoverIdx < 0 Then Return '?
                Dim oMover As Mover = goMovers(lMoverIdx)
                If oMover Is Nothing Then Return '?
                oMover.colDests.Add(uDockOp)
                'Dim lCurUB As Int32 = Math.Min(glMoverUB, glMoverIdx.GetUpperBound(0))
                'For X = 0 To lCurUB
                '    If glMoverIdx(X) = lObjID Then

                '        If goMovers(X).ObjTypeID = iObjTypeID Then
                '            goMovers(X).colDests.Add(uDockOp)
                '            ''for some reason, add it twice???
                '            'goMovers(X).colDests.Add(uDockOp)
                '            Exit For
                '        End If
                '    End If
                'Next X

                'Now, store this with our final...
				'lLen += yResp.Length + 1
				'            ReDim Preserve yFinal(lLen)
				'            System.BitConverter.GetBytes(CShort(yResp.Length)).CopyTo(yFinal, lDestPos) : lDestPos += 2
				'            yResp.CopyTo(yFinal, lDestPos) : lDestPos += yResp.Length
				lDestPos = AppendLenAppendedMsg(yResp, yFinal, lDestPos, 1000)
			End If

        End While

		yFinal = FinalizeLenAppendedMsg(yFinal, lDestPos)

		'Now, send all responses in one packet
		If yFinal Is Nothing = False AndAlso yFinal.Length > 0 Then
			'ReDim Preserve yFinal(yFinal.Length - 2)
			oSocket.SendLenAppendedData(yFinal)
		End If

	End Sub

	Public Shared Function AppendLenAppendedMsg(ByRef yData() As Byte, ByRef yFinal() As Byte, ByVal lPos As Int32, ByVal lCacheGrowth As Int32) As Int32
		If yFinal Is Nothing Then ReDim yFinal(-1)
		If yData Is Nothing Then Return 0

		Dim lSingleMsgLen As Int32 = yData.Length
		'Ok, before we continue, check if we need to increase our cache
		If lPos + lSingleMsgLen + 2 > yFinal.Length Then
			'increase it
			ReDim Preserve yFinal(yFinal.Length + lCacheGrowth)
		End If
		System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yFinal, lPos)
		lPos += 2
		'GetAddObjectMessage(goUnit(X), EpicaMessageCode.eAddObjectCommand).CopyTo(yCache, lPos)
		yData.CopyTo(yFinal, lPos)
		lPos += lSingleMsgLen

		Return lPos
	End Function
	Public Shared Function FinalizeLenAppendedMsg(ByRef yMsg() As Byte, ByVal lPos As Int32) As Byte()
		Dim yFinal() As Byte = Nothing
		If lPos <> 0 Then
			ReDim yFinal(lPos - 1)
			Array.Copy(yMsg, 0, yFinal, 0, lPos)
		End If
		Return yFinal
	End Function

    'Private Sub HandleSetMiningLoc(ByVal yData() As Byte, ByVal lDomainIndex As Int32)
    '    Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
    '    Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
    '    Dim lCacheID As Int32 = System.BitConverter.ToInt32(yData, 8)
    '    Dim lObjX As Int32 = System.BitConverter.ToInt32(yData, 12)
    '    Dim lObjZ As Int32 = System.BitConverter.ToInt32(yData, 16)
    '    Dim iObjA As Int16 = System.BitConverter.ToInt16(yData, 20)
    '    Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, 22)
    '    Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 26)
    '    Dim lCacheX As Int32 = System.BitConverter.ToInt32(yData, 28)
    '    Dim lCacheZ As Int32 = System.BitConverter.ToInt32(yData, 32)

    '    Dim yResp() As Byte

    '    Dim X As Int32
    '    Dim colNewDests As Collection
    '    Dim uMineOp As DestLoc

    '    Dim lIdx As Int32 = -1

    '    colNewDests = GetDestList(lObjX, lObjZ, iObjA, lEnvirID, iEnvirTypeID, lCacheX, lCacheZ, 0, lEnvirID, iEnvirTypeID)

    '    If colNewDests Is Nothing Then
    '        'return invalid dest
    '        System.BitConverter.GetBytes(EpicaMessageCode.eMoveObjectRequestDeny).CopyTo(yData, 0)
    '        moDomains(lDomainIndex).SendData(yData)
    '    Else
    '        For X = 0 To glMoverUB
    '            If glMoverIdx(X) = lObjID Then
    '                If goMovers(X).ObjTypeID = iObjTypeID Then
    '                    lIdx = X

    '                    'Ok, clear our current list
    '                    goMovers(X).colDests = Nothing
    '                    Exit For
    '                End If
    '            End If
    '        Next X

    '        If lIdx = -1 Then
    '            glMoverUB += 1
    '            ReDim Preserve goMovers(glMoverUB)
    '            ReDim Preserve glMoverIdx(glMoverUB)
    '            lIdx = glMoverUB
    '            goMovers(lIdx) = New Mover()
    '            glMoverIdx(lIdx) = lObjID
    '            With goMovers(lIdx)
    '                .lCurAngle = iObjA
    '                .lCurLocX = lObjX
    '                .lCurLocZ = lObjZ
    '                .lCurLocY = 0
    '                .ObjectID = lObjID
    '                .ObjTypeID = iObjTypeID
    '            End With
    '        End If

    '        'Finally, add our final dest item
    '        With uMineOp
    '            .DestAngle = 0
    '            .DestX = lCacheX
    '            .DestZ = lCacheZ
    '            .iDestTypeID = iEnvirTypeID
    '            .lDestID = lEnvirID
    '            .yChangeEnvironment = 0
    '            .ySpecialOp = SpecialOp.eBeginMiningOp
    '            .lSpecialOpID = lCacheID
    '            .iSpecialOpTypeID = ObjectType.eMineralCache
    '        End With
    '        colNewDests.Add(uMineOp)

    '        goMovers(lIdx).colDests = colNewDests

    '        'Now, we kinda cheat here...
    '        yResp = goMovers(lIdx).ProcessStopMessage(lObjX, lObjZ, iObjA)

    '        If yResp.Length > 0 Then
    '            moDomains(lDomainIndex).SendData(yResp)
    '        Else
    '            'Ok, send invalid move
    '            goMovers(lIdx).bMoving = False
    '            System.BitConverter.GetBytes(EpicaMessageCode.eMoveObjectRequestDeny).CopyTo(yData, 0)
    '            moDomains(lDomainIndex).SendData(yResp)
    '        End If
    '    End If

    'End Sub

    Private Sub HandleSetMiningLoc(ByVal yData() As Byte, ByVal lDomainIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lCacheID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iCacheTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lCacheX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lCacheZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lObjID As Int32
        Dim iObjTypeID As Int16
        Dim lLocX As Int32
        Dim lLocZ As Int32
        Dim iLocA As Int16
        Dim lParentID As Int32
        Dim iParentTypeID As Int16
        Dim iModelID As Int16

        'Dim X As Int32

        Dim yResp() As Byte

        Dim uMineOp As DestLoc

        Dim yFinal() As Byte = Nothing
		'        Dim lLen As Int32
        Dim lDestPos As Int32

        While lPos + 24 < yData.Length '- 1
            lObjID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            lLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            lLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iLocA = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            lParentID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iParentTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            'For backwards compatibility...
            iModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            'MSC: PF only cares about the model portion of the modelid
            iModelID = (iModelID And 255S)

            yResp = ExecuteMoveRequest(lObjID, iObjTypeID, lLocX, lLocZ, iLocA, lParentID, iParentTypeID, _
              lCacheX, lCacheZ, 0, lParentID, iParentTypeID, iModelID)

            If yResp Is Nothing = False AndAlso yResp.Length > 0 Then
                'now, add our final dest
                With uMineOp
                    .DestAngle = 0
                    .DestX = lCacheX
                    .DestZ = lCacheZ
                    .iDestTypeID = iParentTypeID
                    .lDestID = lParentID
                    .yChangeEnvironment = 0
                    .ySpecialOp = SpecialOp.eBeginMiningOp
                    .lSpecialOpID = lCacheID
                    .iSpecialOpTypeID = iCacheTypeID
                End With

                'unfortunately, gotta find this object again...
                'TODO: remove the redundant search, possibly expand the Execute routines
                Dim lMoverIdx As Int32 = GetOrAddMover(lObjID, iObjTypeID, lLocX, lLocZ, iLocA, iModelID, lParentID, iParentTypeID)
                If lMoverIdx < 0 Then Return '?
                Dim oMover As Mover = goMovers(lMoverIdx)
                If oMover Is Nothing Then Return
                oMover.colDests.Add(uMineOp)
                'Dim lCurUB As Int32 = Math.Min(glMoverUB, glMoverIdx.GetUpperBound(0))
                'For X = 0 To lCurUB
                '    If glMoverIdx(X) = lObjID Then
                '        If goMovers(X).ObjTypeID = iObjTypeID Then
                '            goMovers(X).colDests.Add(uMineOp)

                '            Exit For
                '        End If
                '    End If
                'Next X


                'Now, store this with our final...
				'lLen += yResp.Length + 1 '
				'            ReDim Preserve yFinal(lLen)
				'            System.BitConverter.GetBytes(CShort(yResp.Length)).CopyTo(yFinal, lDestPos) : lDestPos += 2
				'            yResp.CopyTo(yFinal, lDestPos) : lDestPos += yResp.Length
				lDestPos = AppendLenAppendedMsg(yResp, yFinal, lDestPos, 1000)
            End If

		End While

		yFinal = FinalizeLenAppendedMsg(yFinal, lDestPos)

        'Now, send all responses in one packet
        If yFinal Is Nothing = False AndAlso yFinal.Length > 0 Then moDomains(lDomainIndex).SendLenAppendedData(yFinal)
    End Sub

    Private Sub HandleSetEntityProduction(ByVal yData() As Byte, ByVal lDomainIndex As Int32)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lBuildObjID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iBuildObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
        Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, 14)
        Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, 18)
        Dim iDestA As Int16 = System.BitConverter.ToInt16(yData, 22)
        Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, 24)
        Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, 28)
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, 32)
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 36)

        Dim iModelID As Int16 = System.BitConverter.ToInt16(yData, 38)
        'MSC: PF only cares about the model portion of the modelid
        iModelID = (iModelID And 255S)

        Dim yResp() As Byte

        ' Dim X As Int32
        Dim colNewDests As Collection
        Dim uMineOp As DestLoc

        Dim lIdx As Int32 = -1
        'Dim lFirstIdx As Int32 = -1

        If iEnvirTypeID = ObjectType.ePlanet Then
            colNewDests = GetDestList(lLocX, lLocZ, 0, lEnvirID, iEnvirTypeID, lDestX, lDestZ, iDestA, lEnvirID, iEnvirTypeID, Mover.GetPathLocs(iModelID))
        Else
            colNewDests = GetDestList(lLocX, lLocZ, 0, lEnvirID, iEnvirTypeID, lDestX, lDestZ, iDestA, lEnvirID, iEnvirTypeID, 0)
        End If

        If colNewDests Is Nothing Then
            'return invalid dest
            System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yData, 0)
            moDomains(lDomainIndex).SendData(yData)
        Else
            lIdx = GetOrAddMover(lObjID, iObjTypeID, lLocX, lLocZ, 0, iModelID, lEnvirID, iEnvirTypeID)
            If lIdx = -1 Then Return

            'Finally, add our final dest item
            With uMineOp
                .DestAngle = iDestA
                .DestX = lDestX
                .DestZ = lDestZ
                .iDestTypeID = iEnvirTypeID
                .lDestID = lEnvirID
                .yChangeEnvironment = 0
                .ySpecialOp = SpecialOp.eBeginConstruction
                .lSpecialOpID = lBuildObjID
                .iSpecialOpTypeID = iBuildObjTypeID
            End With
            colNewDests.Add(uMineOp)

            goMovers(lIdx).colDests = colNewDests

            'Now, we kinda cheat here...
            yResp = goMovers(lIdx).ProcessStopMessage(lLocX, lLocZ, 0)

            If yResp.Length > 0 Then
                moDomains(lDomainIndex).SendData(yResp)
            Else
                'Ok, send invalid move
                goMovers(lIdx).bMoving = False
                System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yData, 0)
                moDomains(lDomainIndex).SendData(yData)
            End If
        End If

    End Sub

    Private Sub HandleJumpTarget(ByVal yData() As Byte, ByVal lDomainIndex As Int32)
        'MsgCode (2), JumpTargetX (4), JumpTargetZ (4), JumpTargetA (2), JumpTargetEnvirID (4), JumpTargetEnvirTypeID (2)
        '  JumpTargetID (4), JumpTargetTypeID (2)
        '  LIST
        '    GUID (6)
        '    LocX (4), LocZ (4), LocA (2), LocGUID (6)
        Dim lPos As Int32 = 2
        Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDestA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lJumpTargetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iJumpTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

		iDestA = -1

        Dim lObjID As Int32
        Dim iObjTypeID As Int16
        Dim lLocX As Int32
        Dim lLocZ As Int32
        Dim iLocA As Int16
        Dim lLocID As Int32
        Dim iLocTypeID As Int16
        Dim iModelID As Int16

        Dim yResp() As Byte
        Dim yFinal() As Byte = Nothing

        Dim lLen As Int32 = 0
        Dim lDestPos As Int32 = 0
        Dim lSizeChange As Int32
        Dim bFirst As Boolean = True

        'TODO: Is this what we want to do?
        On Error Resume Next

        Dim lDestCell As Int32 = Int32.MinValue
        If iDestTypeID = ObjectType.ePlanet Then
            For X As Int32 = 0 To glPlanetVPUB
                If goPlanetVP(X).lID = lDestID Then
                    lDestCell = goPlanetVP(X).oPlanetRef.GetCellSpacing
                    Exit For
                End If
            Next X
            If lDestCell <> Int32.MinValue Then
                Dim lExtent As Int32 = lDestCell * TerrainClass.Width
                Dim lHlfExt As Int32 = lExtent \ 2
                If lDestX < -lHlfExt Then
                    lDestX = -lHlfExt
                ElseIf lDestX > lHlfExt Then
                    lDestX = lHlfExt
                End If
            End If
        End If

        'Now, begin our list
        While lPos + 24 < yData.Length '- 1
            lObjID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            lLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            lLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iLocA = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            lLocID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iLocTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            iModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            'MSC: PF only cares about the model portion of the modelid
            iModelID = (iModelID And 255S)

            If iLocTypeID = ObjectType.ePlanet Then
                Dim lCell As Int32 = Int32.MinValue
                For X As Int32 = 0 To glPlanetVPUB
                    If goPlanetVP(X).lID = lLocID Then
                        lCell = goPlanetVP(X).oPlanetRef.GetCellSpacing()
                        Exit For
                    End If
                Next X

                If lCell <> Int32.MinValue Then
                    Dim lExtent As Int32 = lCell * TerrainClass.Width
                    Dim lHlfExt As Int32 = lExtent \ 2
                    If lLocX < -lHlfExt Then
                        lLocX = -lHlfExt
                    ElseIf lLocX > lHlfExt Then
                        lLocX = lHlfExt
                    End If
                End If
            End If

            ''Find our mover...
            Dim lIdx As Int32 = GetOrAddMover(lObjID, iObjTypeID, lLocX, lLocZ, iLocA, iModelID, lLocID, iLocTypeID)
            If lIdx = -1 Then Return

            With goMovers(lIdx)
                .lCurAngle = iLocA
                .lCurLocX = lLocX
                .lCurLocZ = lLocZ
                .lCurLocY = 0
                .lEnvirID = lLocID
                .iEnvirTypeID = iLocTypeID
                .rcArea.X = .lCurLocX
                .rcArea.Y = .lCurLocZ
            End With

            If goMovers(lIdx).colDests Is Nothing = False Then goMovers(lIdx).colDests.Clear()
            goMovers(lIdx).colDests = Nothing
            Dim colDestList As Collection = GetDestList(lLocX, lLocZ, iLocA, lLocID, iLocTypeID, lDestX, lDestZ, iDestA, lDestID, iDestTypeID, Mover.GetPathLocs(iModelID))
            If colDestList Is Nothing Then
                ReDim yResp(7)
                System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yResp, 0)
                System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
                System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
            Else
                Dim uFinalDest As DestLoc
                With uFinalDest
                    .DestAngle = iDestA
                    .DestX = lDestX
                    .DestZ = lDestZ
                    .iDestTypeID = iDestTypeID
                    .lDestID = lDestID
                    .lSpecialOpID = lJumpTargetID
                    .iSpecialOpTypeID = iJumpTargetTypeID
                    .yChangeEnvironment = ChangeEnvironmentType.eNoChangeEnvironment
                    .ySpecialOp = SpecialOp.eJumpTarget
                End With
                colDestList.Add(uFinalDest)
                goMovers(lIdx).colDests = colDestList
                'If goMovers(lIdx).bMoving = False Then yResp = goMovers(lIdx).ProcessStopMessage(lLocX, lLocZ, iLocA) Else yResp = Nothing
                yResp = goMovers(lIdx).ProcessStopMessage(lLocX, lLocZ, iLocA)
            End If

            'Now, store the array in yFinal...
            If yResp Is Nothing = False AndAlso yResp.Length > 0 Then
				'lLen += yResp.Length + 1
				'            ReDim Preserve yFinal(lLen)
				'            System.BitConverter.GetBytes(CShort(yResp.Length)).CopyTo(yFinal, lDestPos) : lDestPos += 2
				'            yResp.CopyTo(yFinal, lDestPos) : lDestPos += yResp.Length
				lDestPos = AppendLenAppendedMsg(yResp, yFinal, lDestPos, 1000)

                lSizeChange = gl_FINAL_GRID_SQUARE_SIZE
                If iModelID < goModels.Length Then
                    lSizeChange = goModels(iModelID).lRectSize
                Else
                    For X As Int32 = 0 To glModelUB
                        If goModels(X) Is Nothing = False AndAlso goModels(X).lModelID = iModelID Then
                            lSizeChange = goModels(X).lRectSize
                            Exit For
                        End If
                    Next X
                End If

            End If

        End While

		'Now, send all responses in one packet
		yFinal = FinalizeLenAppendedMsg(yFinal, lDestPos)
        If yFinal Is Nothing = False AndAlso yFinal.Length > 0 Then
			'ReDim Preserve yFinal(yFinal.Length - 2)
            moDomains(lDomainIndex).SendLenAppendedData(yFinal)
        End If
    End Sub

    Private Sub HandleGroupMoveRequest(ByVal yData() As Byte, ByVal lDomainIndex As Int32, ByVal bAdd As Boolean)
        'MsgCode (2), DestX (4), DestZ (4), DestA (2), DestID (4), DestTypeID (2)
        '  LIST
        '    GUID (6)
		'    LocX (4), LocZ (4), LocA (2), LocGUID (6)
		Dim lPos As Int32 = 2
        Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDestA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

		iDestA = -1

        Dim lObjID As Int32
        Dim iObjTypeID As Int16
        Dim lLocX As Int32
        Dim lLocZ As Int32
        Dim iLocA As Int16
        Dim lLocID As Int32
        Dim iLocTypeID As Int16
        Dim iModelID As Int16

        Dim lLeftXMod As Int32
        Dim lRightXMod As Int32
        Dim lFrontZMod As Int32
        Dim lBackZMod As Int32
        Dim lPlaceID As Int32 = 3   '0 = front, 1 = left, 2 = back, 3 = right

        Dim yResp() As Byte
        Dim yFinal() As Byte = Nothing

        Dim lTX As Int32 = lDestX
        Dim lTZ As Int32 = lDestZ

        Dim lLen As Int32 = 0
        Dim lDestPos As Int32 = 0
        Dim lSizeChange As Int32
        Dim bFirst As Boolean = True

        Dim uReuse(-1) As LocationalGrouping

        'TODO: Is this what we want to do?
        On Error Resume Next

        'TODO: Implement Formations here... we would need to sort by our Model ID's Size

        Dim lDestCell As Int32 = Int32.MinValue
		If iDestTypeID = ObjectType.ePlanet Then
			If lDestID >= 500000000 Then
				lDestCell = Planet.GetInstancePlanet.GetCellSpacing
			Else
				For X As Int32 = 0 To glPlanetVPUB
					If goPlanetVP(X).lID = lDestID Then
						lDestCell = goPlanetVP(X).oPlanetRef.GetCellSpacing
						Exit For
					End If
				Next X
			End If
			If lDestCell <> Int32.MinValue Then
				Dim lExtent As Int32 = lDestCell * TerrainClass.Width
				Dim lHlfExt As Int32 = lExtent \ 2
				If lDestX < -lHlfExt Then
                    lDestX = -lHlfExt
				ElseIf lDestX > lHlfExt Then
                    lDestX = lHlfExt
				End If
			End If
		End If

        'Now, begin our list
        While lPos + 24 < yData.Length '- 1
            lObjID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			'If lObjID = 200228 Then Stop
            lLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            lLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iLocA = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            lLocID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iLocTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            iModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            'MSC: PF only cares about the model portion of the modelid
            iModelID = (iModelID And 255S)

            If iLocTypeID = ObjectType.ePlanet Then
				Dim lCell As Int32 = Int32.MinValue
				If lLocID >= 500000000 Then
					lCell = Planet.GetInstancePlanet.GetCellSpacing
				Else
					For X As Int32 = 0 To glPlanetVPUB
						If goPlanetVP(X).lID = lLocID Then
							lCell = goPlanetVP(X).oPlanetRef.GetCellSpacing()
							Exit For
						End If
					Next X
				End If
                If lCell <> Int32.MinValue Then
                    Dim lExtent As Int32 = lCell * TerrainClass.Width
                    Dim lHlfExt As Int32 = lExtent \ 2
                    If lLocX < -lHlfExt Then
                        lLocX = -lHlfExt
                    ElseIf lLocX > lHlfExt Then
                        lLocX = lHlfExt
                    End If
                End If
            End If

            'Now... do our stuff... find our mover
            'MSC - 06/25/2007 - Added this for better processing of move requests to fewer hits to the pather
            If iDestTypeID = ObjectType.ePlanet AndAlso iLocTypeID = iDestTypeID AndAlso lLocID = lDestID Then
                Dim lIdx As Int32 = GetOrAddMover(lObjID, iObjTypeID, lLocX, lLocZ, iLocA, iModelID, lLocID, iLocTypeID)
                If lIdx = -1 Then Continue While

                With goMovers(lIdx)
                    .lCurAngle = iLocA
                    .lCurLocX = lLocX
                    .lCurLocZ = lLocZ
                    .lCurLocY = 0
                    .lEnvirID = lLocID
                    .iEnvirTypeID = iLocTypeID
                    .rcArea.X = .lCurLocX
                    .rcArea.Y = .lCurLocZ

                    If .oFormation Is Nothing = False Then .oFormation.DetachItem(.ObjectID, .ObjTypeID)
                    .oFormation = Nothing
                End With 

                If bAdd = True AndAlso goMovers(lIdx).colDests Is Nothing Then goMovers(lIdx).colDests = New Collection()

                'Get our vert X
                Dim lVX As Int32
                Dim lVZ As Int32

                If bAdd = True AndAlso goMovers(lIdx).colDests.Count <> 0 Then
                    'get the last dest
                    Dim uDest As DestLoc = CType(goMovers(lIdx).colDests.Item(goMovers(lIdx).colDests.Count), DestLoc)
                    'lVX and lVZ are based from that
                    lVX = uDest.DestX \ lDestCell
                    lVZ = uDest.DestZ \ lDestCell
                Else
                    lVX = lLocX \ lDestCell
                    lVZ = lLocZ \ lDestCell
                End If

                Dim colNewDests As Collection = Nothing
                Dim bReused As Boolean = False
                For X As Int32 = 0 To uReuse.GetUpperBound(0)
                    If uReuse(X).VertX = lVX AndAlso uReuse(X).VertZ = lVZ Then
                        colNewDests = uReuse(X).colDestList
                        bReused = True
                        Exit For
                    End If
                Next X

                If colNewDests Is Nothing Then
                    colNewDests = GetDestList(lLocX, lLocZ, iLocA, lLocID, iLocTypeID, lTX, lTZ, iDestA, lDestID, iDestTypeID, Mover.GetPathLocs(iModelID))
                    If colNewDests Is Nothing Then Continue While 'TODO: Respond with unable to get there
                    ReDim Preserve uReuse(uReuse.GetUpperBound(0) + 1)
                    With uReuse(uReuse.GetUpperBound(0))
                        .VertX = lVX : .VertZ = lVZ : .colDestList = New Collection 'colNewDests
                        For Each uDest As DestLoc In colNewDests
                            .colDestList.Add(uDest)
                        Next
                    End With
                End If

                If bAdd = True AndAlso goMovers(lIdx).colDests.Count <> 0 Then
                    'Now, move our colNewDests over to colDests
                    For Each uDest As DestLoc In colNewDests
                        goMovers(lIdx).colDests.Add(uDest)
                    Next
                Else
                    'Normal move, so we just replace the collection
                    If bReused = True Then
                        If goMovers(lIdx).colDests Is Nothing Then goMovers(lIdx).colDests = New Collection
                        goMovers(lIdx).colDests.Clear()
                        For Each uDest As DestLoc In colNewDests
                            goMovers(lIdx).colDests.Add(uDest)
                        Next
                    Else : goMovers(lIdx).colDests = colNewDests
                    End If
                End If

                'Now, we kinda cheat here...
                yResp = goMovers(lIdx).ProcessStopMessage(lLocX, lLocZ, iLocA)
                If yResp Is Nothing OrElse yResp.Length = 0 Then
                    'Ok, send invalid move
                    goMovers(lIdx).bMoving = False
                    ReDim yResp(7)
                    System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yResp, 0)
                    System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
                    System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
                End If
            Else
                'Do it like normal
                If bAdd = True Then
                    yResp = ExecuteAddWaypointRequest(lObjID, iObjTypeID, lLocX, lLocZ, iLocA, lLocID, iLocTypeID, lTX, lTZ, iDestA, lDestID, iDestTypeID, iModelID)
                Else
                    yResp = ExecuteMoveRequest(lObjID, iObjTypeID, lLocX, lLocZ, iLocA, lLocID, iLocTypeID, lTX, lTZ, iDestA, lDestID, iDestTypeID, iModelID)
                End If
            End If

            'Now, store the array in yFinal...
            If yResp Is Nothing = False AndAlso yResp.Length > 0 Then
				'lLen += yResp.Length + 1 '
				'            ReDim Preserve yFinal(lLen)
				'            System.BitConverter.GetBytes(CShort(yResp.Length)).CopyTo(yFinal, lDestPos) : lDestPos += 2
				'            yResp.CopyTo(yFinal, lDestPos) : lDestPos += yResp.Length
				lDestPos = AppendLenAppendedMsg(yResp, yFinal, lDestPos, 1000)

                lSizeChange = gl_FINAL_GRID_SQUARE_SIZE
                If iModelID < goModels.Length Then
                    lSizeChange = goModels(iModelID).lRectSize
                Else
                    For X As Int32 = 0 To glModelUB
                        If goModels(X) Is Nothing = False AndAlso goModels(X).lModelID = iModelID Then
                            lSizeChange = goModels(X).lRectSize
                            Exit For
                        End If
                    Next X
                End If

                'Increment our space usage
                If bFirst = True Then
                    bFirst = False
                    lFrontZMod = lSizeChange
                    lLeftXMod = -lSizeChange
                    lBackZMod = -lSizeChange
                    lRightXMod = lSizeChange
                Else
                    Select Case lPlaceID
                        Case 0 : lFrontZMod += lSizeChange
                        Case 1 : lLeftXMod -= lSizeChange
                        Case 2 : lBackZMod -= lSizeChange
                        Case 3 : lRightXMod += lSizeChange
                    End Select
                End If

                lPlaceID += 1
                If lPlaceID = 8 Then lPlaceID = 0

                'Get our next coordinate
                Select Case lPlaceID
                    Case 0 : lTX = lDestX : lTZ = lDestZ + lFrontZMod
                    Case 1 : lTX = lDestX + lLeftXMod : lTZ = lDestZ
                    Case 2 : lTX = lDestX : lTZ = lDestZ + lBackZMod
                    Case 3 : lTX = lDestX + lRightXMod : lTZ = lDestZ 'else = 3 or otherwise...
                    Case 4 : lTX = lDestX + lLeftXMod : lTZ = lDestZ + lFrontZMod
                    Case 5 : lTX = lDestX + lRightXMod : lTZ = lDestZ + lFrontZMod
                    Case 6 : lTX = lDestX + lLeftXMod : lTZ = lDestZ + lBackZMod
                    Case Else : lTX = lDestX + lRightXMod : lTZ = lDestZ + lBackZMod
                End Select

            End If

		End While

		yFinal = FinalizeLenAppendedMsg(yFinal, lDestPos)

        'Now, send all responses in one packet
        If yFinal Is Nothing = False AndAlso yFinal.Length > 0 Then
			'ReDim Preserve yFinal(yFinal.Length - 2)
            moDomains(lDomainIndex).SendLenAppendedData(yFinal)
        End If
    End Sub

    Public Sub SendRequestEntities()
        Dim yMsg(1) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.ePFRequestEntitys).CopyTo(yMsg, 0)
        Me.moPrimary.SendData(yMsg)
    End Sub

    Private Sub ReceiveEntities(ByVal yData() As Byte)
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim X As Int32

        Dim lIdx As Int32
        Dim lPos As Int32 = 6

        'Should only be called when the PF server starts up... so, we'll fill it as we go
        glMoverUB = lCnt - 1
        ReDim goMovers(glMoverUB)
        lIdx = 0

        For X = 0 To lCnt - 1
            goMovers(lIdx) = New Mover()
            With goMovers(lIdx)
                .lGlobalIdx = lIdx
                .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .lEnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .iEnvirTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .lCurLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lCurLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .iModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 4

                .rcArea = Mover.GetRectWidthHeight(.iModelID)
                .rcArea.X = .lCurLocX
                .rcArea.Y = .lCurLocZ
            End With
        Next X
    End Sub

    Private Sub ReceiveAddObject(ByVal yData() As Byte)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

		If iObjTypeID = ObjectType.eSolarSystem Then
			Dim oSystem As SolarSystem = goGalaxy.GetSystem(lObjID)
			If oSystem Is Nothing = False Then
				'already have the system, return
				Return
			End If
			oSystem = New SolarSystem()
			With oSystem
				.ObjectID = lObjID
				.ObjTypeID = iObjTypeID
				Dim lPos As Int32 = 8
				.SystemName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
				lPos += 4	'galaxyid
				.LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.LocY = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.StarType1Idx = StarType.GetStarTypeIdx(yData(lPos)) : lPos += 1
				.StarType2Idx = StarType.GetStarTypeIdx(yData(lPos)) : lPos += 1
                .StarType3Idx = StarType.GetStarTypeIdx(yData(lPos)) : lPos += 1
                .SystemType = yData(lPos) : lPos += 1
			End With
			goGalaxy.AddSystem(oSystem)
			'Now, get the system details
			Dim yMsg(5) As Byte

			System.BitConverter.GetBytes(GlobalMessageCode.eRequestSystemDetails).CopyTo(yMsg, 0)
			System.BitConverter.GetBytes(oSystem.ObjectID).CopyTo(yMsg, 2)
			moPrimary.SendData(yMsg)
        Else
            Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, 8)
            Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
            Dim lCurLocX As Int32 = System.BitConverter.ToInt32(yData, 14)
            Dim lCurLocZ As Int32 = System.BitConverter.ToInt32(yData, 18)
            Dim iModelID As Int16 = System.BitConverter.ToInt16(yData, 22)
            'MSC: PF only cares about the model portion of the modelid
            iModelID = (iModelID And 255S)
            Dim lIdx As Int32 = GetOrAddMover(lObjID, iObjTypeID, lCurLocX, lCurLocZ, 0, iModelID, lEnvirID, iEnvirTypeID)
        End If

    End Sub

    Private Sub HandleRemoveObject(ByVal yData() As Byte)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        Dim lCurUB As Int32 = Math.Min(glMoverUB, glMoverIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If glMoverIdx(X) = lObjID AndAlso goMovers(X) Is Nothing = False AndAlso goMovers(X).ObjTypeID = iObjTypeID Then
                glMoverIdx(X) = -1
                'goMovers(X) = Nothing
                Return
            End If
        Next X
    End Sub

    Private Sub HandleClearDestList(ByRef yData() As Byte)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        Dim lCurUB As Int32 = Math.Min(glMoverUB, glMoverIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If glMoverIdx(X) = lObjID AndAlso goMovers(X) Is Nothing = False AndAlso goMovers(X).ObjTypeID = iObjTypeID Then
                If goMovers(X).colDests Is Nothing = False Then goMovers(X).colDests.Clear()
                Exit For
            End If
        Next X
    End Sub

    Private Sub HandleSetFleetDest(ByRef yData() As Byte, ByVal lDomainIndex As Int32)
        'MsgCode (2), origin systemID (4), targetsystemid (4), lCnt (4), List: ID (4), X (4), Z (4), Angle (2), EnvirGUID (6)
        Dim lPos As Int32 = 2
        Dim lOriginID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'Ok, let's find our systems
        Dim oOrigin As SolarSystem = Nothing
        Dim oTarget As SolarSystem = Nothing

        For X As Int32 = 0 To goGalaxy.mlSystemUB
            If goGalaxy.moSystems(X).ObjectID = lOriginID Then
                oOrigin = goGalaxy.moSystems(X)
            ElseIf goGalaxy.moSystems(X).ObjectID = lTargetID Then
                oTarget = goGalaxy.moSystems(X)
            End If
            If oOrigin Is Nothing = False AndAlso oTarget Is Nothing = False Then Exit For
        Next X

        If oOrigin Is Nothing OrElse oTarget Is Nothing Then Return

        'Ok, now... get our X,Z angles... we put it in ORIGIN, DEST because we want the LEAVE FROM Angle
        'Dim fAngle As Single = LineAngleDegrees(oOrigin.LocX, oOrigin.LocZ, oTarget.LocX, oTarget.LocZ)

        'Now... set our DX and DZ values
        'Dim fDX As Single = 5000000
        'Dim fDZ As Single = 0
        'and rotate them by fAngle
        'RotatePoint(0, 0, fDX, fDZ, fAngle)

        'Dim fAbsX As Single = Math.Abs(fDX)
        'Dim fAbsZ As Single = Math.Abs(fDZ)

        ''Ok, that tells us where we are leaving from... we need to normalize the distance so that either X or Z hits +/-5000000
        'If fAbsX <> 5000000 AndAlso fAbsZ <> 5000000 Then
        '    If fAbsX > fAbsZ Then
        '        'ok, make X 5000000 and adjust Z
        '        fDZ *= (5000000 / fAbsX)
        '        If fDX < 0 Then fDX = -5000000 Else fDX = 5000000
        '    Else
        '        'Make Z 5000000 and adjust X
        '        fDX *= (5000000 / fAbsZ)
        '        If fDZ < 0 Then fDZ = -5000000 Else fDZ = 5000000
        '    End If
        'End If

        ''Now... determine destX and DestZ
        'Dim lDX As Int32 = CInt(fDX)
        'Dim lDZ As Int32 = CInt(fDZ)
        'Dim iDA As Int16 = CShort(fAngle * 10)


        'Now... determine destX and DestZ (Using Fleet Jump Point)
        Dim fAngle As Single = LineAngleDegrees(oOrigin.FleetJumpPointX, oOrigin.FleetJumpPointZ, 0, 0)
        If fAngle >= 90.0F Then
            fAngle -= 90.0F
        Else
            fAngle = 360.0F - fAngle - 90.0F
        End If
        Dim lDX As Int32 = oOrigin.FleetJumpPointX
        Dim lDZ As Int32 = oOrigin.FleetJumpPointZ
        Dim iDA As Int16 = CShort(fAngle * 10)
        'End new Fleet Jump Point

        'Ok, now, let's go through our mover list...
        'List: ID (4), X (4), Z (4), Angle (2), EnvirGUID (6)
        Dim lLen As Int32 = 0
        Dim lDestPos As Int32 = 0
        Dim lSizeChange As Int32
        Dim yFinal() As Byte = Nothing

        For lIdx As Int32 = 0 To lCnt - 1
            Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iObjTypeID As Int16 = ObjectType.eUnit
            Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iLocA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Dim lLocID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iLocTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Dim iModelID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            'MSC - the pathfinding only cares about the first byte of the modelid - the model
            iModelID = (iModelID And 255S)

            'Now... do our stuff
            'Dim yResp() As Byte = ExecuteMoveRequest(lObjID, iObjTypeID, lLocX, lLocZ, iLocA, lLocID, iLocTypeID, lDX, lDZ, iDA, lOriginID, ObjectType.eSolarSystem, iModelID)
            Dim yResp() As Byte = ExecuteFleetRegroupRequest(lObjID, iObjTypeID, lLocX, lLocZ, iLocA, lLocID, iLocTypeID, lDX, lDZ, iDA, lOriginID, ObjectType.eSolarSystem, iModelID)


            'Now, store the array in yFinal...
            If yResp Is Nothing = False AndAlso yResp.Length > 0 Then
                'lLen += yResp.Length + 1
                '            ReDim Preserve yFinal(lLen)
                '            System.BitConverter.GetBytes(CShort(yResp.Length)).CopyTo(yFinal, lDestPos) : lDestPos += 2
                '            yResp.CopyTo(yFinal, lDestPos) : lDestPos += yResp.Length
                lDestPos = AppendLenAppendedMsg(yResp, yFinal, lDestPos, 1000)

                lSizeChange = gl_FINAL_GRID_SQUARE_SIZE
                If iModelID < goModels.Length Then
                    lSizeChange = goModels(iModelID).lRectSize
                Else
                    For X As Int32 = 0 To glModelUB
                        If goModels(X).lModelID = iModelID Then
                            lSizeChange = goModels(X).lRectSize
                            Exit For
                        End If
                    Next X
                End If
            End If

        Next lIdx

        'Now, send all responses in one packet
        yFinal = FinalizeLenAppendedMsg(yFinal, lDestPos)
        If yFinal Is Nothing = False AndAlso yFinal.Length > 0 Then
            'ReDim Preserve yFinal(yFinal.Length - 2)
            moDomains(lDomainIndex).SendLenAppendedData(yFinal)
        End If
    End Sub

    Private Sub HandleSetMaintenanceTarget(ByRef yData() As Byte, ByVal lDomainIndex As Int32, ByVal iMsgCode As Int16)
        'alright, we have this message...
        'MsgCode(2), ObjGUID (6), TargetGUID (6), ObjLoc (8), oEnvirGUID (6) TargetLoc (8), ObjModel (2), TargetModel (2)
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeId As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjModelID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim iTargetModelID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        'Now...
        Dim colNewDests As Collection = New Collection
        If iEnvirTypeId = ObjectType.ePlanet Then
            colNewDests = GetDestList(lLocX, lLocZ, 0, lEnvirID, iEnvirTypeId, lDestX, lDestZ, 0, lEnvirID, iEnvirTypeId, Mover.GetPathLocs(iObjModelID))
        Else
            colNewDests = GetDestList(lLocX, lLocZ, 0, lEnvirID, iEnvirTypeId, lDestX, lDestZ, 0, lEnvirID, iEnvirTypeId, 0)
        End If
        If colNewDests Is Nothing Then
            'return invalid dest
            System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yData, 0)
            moDomains(lDomainIndex).SendData(yData)
        Else
            'Find a mover for the object
            Dim lIdx As Int32 = GetOrAddMover(lObjID, iObjTypeID, lLocX, lLocZ, 0, iObjModelID, lEnvirID, iEnvirTypeId)
            If lIdx = -1 Then Return
             

            'Finally, add our final dest item
            Dim uFinalOp As DestLoc
            With uFinalOp
                .DestAngle = 0
                .DestX = lDestX
                .DestZ = lDestZ
                .iDestTypeID = iEnvirTypeId
                .lDestID = lEnvirID
                .yChangeEnvironment = 0
                If iMsgCode = GlobalMessageCode.eSetRepairTarget Then
                    .ySpecialOp = SpecialOp.eBeginRepair
                ElseIf iMsgCode = GlobalMessageCode.eSetDismantleTarget Then
                    .ySpecialOp = SpecialOp.eBeginDismantle
                End If
                .lSpecialOpID = lTargetID
                .iSpecialOpTypeID = iTargetTypeID
            End With
            colNewDests.Add(uFinalOp)

            goMovers(lIdx).colDests = colNewDests

            'Now, we kinda cheat here...
            Dim yResp() As Byte = goMovers(lIdx).ProcessStopMessage(lLocX, lLocZ, 0)

            If yResp Is Nothing = False AndAlso yResp.Length > 0 Then
                moDomains(lDomainIndex).SendData(yResp)
            Else
                'Ok, send invalid move
                goMovers(lIdx).bMoving = False
                System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yData, 0)
                moDomains(lDomainIndex).SendData(yData)
            End If
        End If
    End Sub

    Private Sub HandleFinalizeStopEvent(ByRef yData() As Byte, ByVal lDomainIndex As Int32)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        Dim lCurUB As Int32 = Math.Min(glMoverUB, glMoverIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If glMoverIdx(X) = lObjID AndAlso goMovers(X) Is Nothing = False AndAlso goMovers(X).ObjTypeID = iTypeID Then
                goMovers(X).yInAForcedMoveSpeed = 0
                If Mover.FinalizeStopEvent(goMovers(X)) = True Then
                    'We send an update of the entity's loc to the primary server for asynchronous persistence
                    Dim yTemp(17) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eUpdateEntityLoc).CopyTo(yTemp, 0)
                    goMovers(X).GetGUIDAsString.CopyTo(yTemp, 2)
                    System.BitConverter.GetBytes(goMovers(X).lCurLocX).CopyTo(yTemp, 8)
                    System.BitConverter.GetBytes(goMovers(X).lCurLocZ).CopyTo(yTemp, 12)
                    System.BitConverter.GetBytes(goMovers(X).lCurAngle).CopyTo(yTemp, 16)
                    goMsgSys.SendToPrimary(yTemp)

                    goMovers(X).bMoving = False
                Else
                    'get the next dest in the list
                    Dim uDest As DestLoc = CType(goMovers(X).colDests(1), DestLoc)
                    'remove that dest from the list
                    goMovers(X).colDests.Remove(1)

                    Dim yTemp(17) As Byte     '0 to 17 = 18 bytes
                    System.BitConverter.GetBytes(GlobalMessageCode.eFinalMoveCommand).CopyTo(yTemp, 0)
                    'System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectCommand).CopyTo(yTemp, 0)
                    goMovers(X).GetGUIDAsString.CopyTo(yTemp, 2)
                    System.BitConverter.GetBytes(uDest.DestX).CopyTo(yTemp, 8)
                    System.BitConverter.GetBytes(uDest.DestZ).CopyTo(yTemp, 12)
                    System.BitConverter.GetBytes(uDest.DestAngle).CopyTo(yTemp, 16)
                    goMovers(X).bMoving = True
                    goMovers(X).lCurDestX = uDest.DestX
                    goMovers(X).lCurDestZ = uDest.DestZ
                    moDomains(lDomainIndex).SendData(yTemp)
                End If
            End If
        Next X
	End Sub

	Private Sub HandleRouteMove(ByRef yData() As Byte, ByVal lDomainIndex As Int32)
		'ok, let's parse our msg
		Dim lPos As Int32 = 2 'for msgcode
		Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim iModelID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim lObjX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lObjZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		'Always the destination environment to move to
		Dim lDestEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iDestEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		'if set, indicates that when I get to my destination (destX,destz) I need to dock with the target
		Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

		'gfrmDisplayForm.LogSpecificEvent("HandleRouteMove: " & lObjID & ", " & iObjTypeID & ". Dock target: " & lTargetID & ", " & iTargetTypeID)

		'Ok, 
		Dim lDestCell As Int32 = Int32.MinValue
		If iDestEnvirTypeID = ObjectType.ePlanet Then
			If lDestEnvirID >= 500000000 Then
				lDestCell = Planet.GetInstancePlanet.GetCellSpacing
			Else
				For X As Int32 = 0 To glPlanetVPUB
					If goPlanetVP(X).lID = lDestEnvirID Then
						lDestCell = goPlanetVP(X).oPlanetRef.GetCellSpacing
						Exit For
					End If
				Next X
			End If
			If lDestCell <> Int32.MinValue Then
				Dim lExtent As Int32 = lDestCell * TerrainClass.Width
				Dim lHlfExt As Int32 = lExtent \ 2
				If lDestX < -lHlfExt Then
                    lDestX = -lHlfExt 'lExtent
				ElseIf lDestX > lHlfExt Then
                    lDestX = lHlfExt 'lExtent
				End If
			End If
		End If

		If iEnvirTypeID = ObjectType.ePlanet Then
			Dim lCell As Int32 = Int32.MinValue
			If lEnvirID >= 500000000 Then
				lCell = Planet.GetInstancePlanet.GetCellSpacing
			Else
				For X As Int32 = 0 To glPlanetVPUB
					If goPlanetVP(X).lID = lEnvirID Then
						lCell = goPlanetVP(X).oPlanetRef.GetCellSpacing()
						Exit For
					End If
				Next X
			End If

			If lCell <> Int32.MinValue Then
				Dim lExtent As Int32 = lCell * TerrainClass.Width
				Dim lHlfExt As Int32 = lExtent \ 2
				If lObjX < -lHlfExt Then
                    lObjX = -lHlfExt 'lExtent
				ElseIf lObjX > lHlfExt Then
                    lObjX = lHlfExt 'lExtent
				End If
			End If
		End If

        'Find our mover...
        Dim lIdx As Int32 = GetOrAddMover(lObjID, iObjTypeID, lObjX, lObjZ, 0, iModelID, lEnvirID, iEnvirTypeID)
        If lIdx < 0 Then Return
 
        With goMovers(lIdx)
            .lCurAngle = 0
            .lCurLocX = lObjX
            .lCurLocZ = lObjZ
            .lCurLocY = 0
            .lEnvirID = lEnvirID
            .iEnvirTypeID = iEnvirTypeID
            .rcArea.X = .lCurLocX
            .rcArea.Y = .lCurLocZ
        End With 

		goMovers(lIdx).bInARoute = True

		If goMovers(lIdx).colDests Is Nothing = False Then goMovers(lIdx).colDests.Clear()
		goMovers(lIdx).colDests = Nothing

		Dim colDestList As Collection = GetDestList(lObjX, lObjZ, 0, lEnvirID, iEnvirTypeID, lDestX, lDestZ, 0, lDestEnvirID, iDestEnvirTypeID, Mover.GetPathLocs(iModelID))
		If colDestList Is Nothing Then
			Dim yResp(7) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequestDeny).CopyTo(yResp, 0)
			System.BitConverter.GetBytes(lObjID).CopyTo(yResp, 2)
			System.BitConverter.GetBytes(iObjTypeID).CopyTo(yResp, 6)
			moDomains(lDomainIndex).SendLenAppendedData(yResp)

			'gfrmDisplayForm.LogSpecificEvent("HandleRouteMove.Failed: " & lObjID & ", " & iObjTypeID & ". Dock target: " & lTargetID & ", " & iTargetTypeID)
		Else
			If iTargetTypeID = ObjectType.eFacility OrElse iTargetTypeID = ObjectType.eUnit Then
				Dim uDockOp As DestLoc
				With uDockOp
					.DestAngle = 0
					.DestX = lDestX
					.DestZ = lDestZ
					.iDestTypeID = iTargetTypeID
					.lDestID = lTargetID
					.yChangeEnvironment = ChangeEnvironmentType.eDocking
					.ySpecialOp = SpecialOp.eNoSpecialOp
					.lSpecialOpID = 0
					.iSpecialOpTypeID = 0
				End With
				colDestList.Add(uDockOp)
			End If
			goMovers(lIdx).colDests = colDestList

			Dim yResp() As Byte = goMovers(lIdx).ProcessStopMessage(lObjX, lObjZ, 0)
			If yResp Is Nothing = False Then moDomains(lDomainIndex).SendData(yResp)
		End If
	End Sub

	Private Sub HandleSetPrimaryTarget(ByRef yData() As Byte, ByVal lIndex As Int32)
		'MsgCode (2), TargetGUID (6), LocX (4), LocZ (4), Angle (2), ModelID (2), GUIDListCnt (4)
		'  GUID LIST:
		'  GUID (6)
		'  Loc (10)
		'  Dest (10)
		'  ModelID (2)
		Dim lPos As Int32 = 2 'for msgcode
		Dim lTargetID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iTargetTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim lTargetX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lTargetZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iTargetA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim iTargetModel As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

		Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

		Dim lGUIDListUB As Int32 = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4

		Dim lCell As Int32 = Int32.MinValue
		If iEnvirTypeID = ObjectType.ePlanet Then
			If lEnvirID >= 500000000 Then
				lCell = Planet.GetInstancePlanet.GetCellSpacing()
			Else
				For X As Int32 = 0 To glPlanetVPUB
					If goPlanetVP(X).lID = lEnvirID Then
						lCell = goPlanetVP(X).oPlanetRef.GetCellSpacing()
						Exit For
					End If
				Next X
			End If
		End If

		Dim lLen As Int32 = 0
		Dim yFinal() As Byte = Nothing
		Dim lDestPos As Int32 = 0

		For lEntityIdx As Int32 = 0 To lGUIDListUB
			Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim iLocA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim iDestA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			Dim iModelID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

			If lLocX = lDestX AndAlso lLocZ = lDestZ Then Continue For

			'Adjust for map wrap
			If lCell <> Int32.MinValue Then
				Dim lExtent As Int32 = lCell * TerrainClass.Width
				Dim lHlfExt As Int32 = lExtent \ 2
				If lLocX < -lHlfExt Then
                    lLocX = -lHlfExt 'lExtent
				ElseIf lLocX > lHlfExt Then
                    lLocX = lHlfExt 'lExtent
				End If
			End If

            'Dim oMover As Mover = GetOrAddMover(lObjID, iObjTypeID, lLocX, lLocZ, iLocA, iModelID, lEnvirID, iEnvirTypeID)
            Dim lMoverIdx As Int32 = GetOrAddMover(lObjID, iObjTypeID, lLocX, lLocZ, iLocA, iModelID, lEnvirID, iEnvirTypeID)
            If lMoverIdx < 0 Then Continue For
            Dim oMover As Mover = goMovers(lMoverIdx)
			If oMover Is Nothing Then Continue For

			'Now, do our anti-stacking here...
			Dim ptTemp As Point = Mover.GetNearestAntiStackingPoint(lObjID, iObjTypeID, lDestX, lDestZ)

			'Now... do our stuff... find our mover
			'TODO: Possible performance issue here for land-based unit pathing
			Dim yResp() As Byte = ExecuteMoveRequestNoLookup(oMover, lLocX, lLocZ, iLocA, lEnvirID, iEnvirTypeID, ptTemp.X, ptTemp.Y, iDestA, lEnvirID, iEnvirTypeID, iModelID)

			'Now, store the array in yFinal...
			If yResp Is Nothing = False AndAlso yResp.Length > 0 Then
				If yFinal Is Nothing Then ReDim yFinal(-1)
				lDestPos = AppendLenAppendedMsg(yResp, yFinal, lDestPos, 2000)
				'lLen += yResp.Length + 1
				'ReDim Preserve yFinal(lLen)
				'System.BitConverter.GetBytes(CShort(yResp.Length)).CopyTo(yFinal, lDestPos) : lDestPos += 2
				'yResp.CopyTo(yFinal, lDestPos) : lDestPos += yResp.Length
			End If
		Next lEntityIdx
		'Now, send all responses in one packet
		If yFinal Is Nothing = False AndAlso yFinal.Length > 0 Then
			yFinal = FinalizeLenAppendedMsg(yFinal, lDestPos)
			moDomains(lIndex).SendLenAppendedData(yFinal)
		End If
	End Sub

	Private Sub HandleAddFormation(ByRef yData() As Byte)
		Dim oFormation As New FormationDefinition()
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
			ReDim Preserve goFormationDef(glFormationDefUB + 1)
			ReDim Preserve glFormationDefIdx(glFormationDefUB + 1)
			goFormationDef(glFormationDefUB + 1) = oFormation
			glFormationDefIdx(glFormationDefUB + 1) = oFormation.lFormationID
			glFormationDefUB += 1
		End If
	End Sub

	Public Sub SendLenReadyMsgToDomain(ByRef yData() As Byte, ByVal lDomainIndex As Int32)
		moDomains(lDomainIndex).SendLenAppendedData(yData)
	End Sub

	Public Sub SendMsgToDomain(ByRef yData() As Byte, ByVal lDomainIndex As Int32)
		moDomains(lDomainIndex).SendData(yData)
	End Sub

	Private Sub HandleMoveFormation(ByRef yData() As Byte, ByVal lDomain As Int32)
		Dim lFormationID As Int32 = System.BitConverter.ToInt32(yData, 2)
		For X As Int32 = 0 To glFormationDefUB
			If glFormationDefIdx(X) = lFormationID Then
				goFormationDef(X).HandleMoveFormation(yData, lDomain)
				Exit For
			End If
		Next X
	End Sub

	Private Sub HandleAlertDestReached(ByRef yData() As Byte, ByVal lDomain As Int32)
		Dim lPos As Int32 = 2
		Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim yValue As Byte = yData(lPos) : lPos += 1

		If glMoverIdx Is Nothing = False Then
			Dim lCurUB As Int32 = Math.Min(glMoverUB, glMoverIdx.GetUpperBound(0))
            For X As Int32 = 0 To lCurUB
                If glMoverIdx(X) = lObjID AndAlso goMovers(X) Is Nothing = False AndAlso goMovers(X).ObjTypeID = iObjTypeID Then
                    If yValue = 0 Then
                        goMovers(X).bAlertAtFinalDest = False
                    Else : goMovers(X).bAlertAtFinalDest = True
                    End If
                    Exit For
                End If
            Next X
		End If
		
	End Sub

	Private Sub HandleForcedMoveSpeedMove(ByRef yData() As Byte, ByVal lDomain As Int32)
		Dim lPos As Int32 = 2	'for msgcode
		Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iDestA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim lDestID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iDestTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim ySpeed As Byte = yData(lPos) : lPos += 1
		Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim yFinal() As Byte = Nothing
		Dim lDestPos As Int32 = 0

		For X As Int32 = 0 To lCnt - 1
			Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim iLocA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			Dim lLocID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim iLocTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			Dim iModelID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            'Dim oMover As Mover = GetOrAddMover(lObjID, iObjTypeID, lLocX, lLocZ, iLocA, iModelID, lLocID, iLocTypeID)
            Dim lMoverIdx As Int32 = GetOrAddMover(lObjID, iObjTypeID, lLocX, lLocZ, iLocA, iModelID, lLocID, iLocTypeID)
            If lMoverIdx < 0 Then Continue For
            Dim oMover As Mover = goMovers(lMoverIdx)
            If oMover Is Nothing = False Then
                oMover.yInAForcedMoveSpeed = ySpeed

                Dim yResp() As Byte = ExecuteMoveRequestNoLookup(oMover, lLocX, lLocZ, iLocA, lLocID, iLocTypeID, lDestX, lDestZ, iDestA, lDestID, iDestTypeID, iModelID)
                'Now, store the array in yFinal...
                If yResp Is Nothing = False AndAlso yResp.Length > 0 Then
                    If yFinal Is Nothing Then ReDim yFinal(-1)
                    lDestPos = AppendLenAppendedMsg(yResp, yFinal, lDestPos, 2000)
                End If

            End If
		Next X

		'Now, send all responses in one packet
		If yFinal Is Nothing = False AndAlso yFinal.Length > 0 Then
			yFinal = FinalizeLenAppendedMsg(yFinal, lDestPos)
			moDomains(lDomain).SendLenAppendedData(yFinal)
		End If
 
	End Sub
End Class

