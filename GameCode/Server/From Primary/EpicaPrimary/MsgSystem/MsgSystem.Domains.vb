Option Strict On

Partial Class MsgSystem
	Private moDomains() As NetSock
	Private mlDomainUB As Int32 = -1
	Private moServers() As DomainServer		'more description to the domain server connections
	Public bReceivedDomains As Boolean = False

	Private WithEvents moDomainListener As NetSock

	Public Property AcceptingDomains() As Boolean
		Get
			Return mbAcceptingDomains
		End Get
		Set(ByVal Value As Boolean)
			Dim lPort As Int32
			If mbAcceptingDomains <> Value Then
				mbAcceptingDomains = Value
				If mbAcceptingDomains Then
					Dim oINI As New InitFile()
					Try
						lPort = CInt(Val(oINI.GetString("SETTINGS", "DomainListenerPort", "0")))
						If lPort = 0 Then Err.Raise(-1, "AcceptingDomains", "Could not get Domain Listen Port Number from INI")
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
					'ok, stop listening
					moDomainListener.StopListening()
				End If
			End If
		End Set
	End Property

	Public Sub BroadcastToDomains(ByRef yMsg() As Byte)
		For X As Int32 = 0 To mlDomainUB
			moDomains(X).SendData(yMsg)
		Next X
	End Sub

	Public Function GetDomainUB() As Int32
		Return mlDomainUB
	End Function

#Region "Domain Listener Handling"
	Private Sub moDomainListener_onConnect(ByVal Index As Integer) Handles moDomainListener.onConnect
		'
	End Sub

	Private Sub moDomainListener_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moDomainListener.onConnectionRequest
		Dim X As Int32
		Dim lIdx As Int32 = -1

		If AcceptingDomains Then
			For X = 0 To mlDomainUB
				If moDomains(X) Is Nothing Then
					lIdx = X
					Exit For
				End If
			Next X

			If lIdx = -1 Then
				mlDomainUB += 1
				ReDim Preserve moDomains(mlDomainUB)
				ReDim Preserve moServers(mlDomainUB)
				lIdx = mlDomainUB
			End If

			moDomains(lIdx) = New NetSock(oClient)
			moDomains(lIdx).SocketIndex = lIdx
			moServers(lIdx) = New DomainServer()
			moServers(lIdx).DomainSocket = moDomains(lIdx)

			moDomains(lIdx).lSpecificID = lIdx
			moDomains(lIdx).lAppType = MsgMonitor.eMM_AppType.RegionServer

			LogEvent(LogEventType.Informational, "Domain Server Connected")

			'add event handlers
			AddHandler moDomains(lIdx).onConnect, AddressOf moDomains_onConnect
			AddHandler moDomains(lIdx).onDataArrival, AddressOf moDomains_onDataArrival
			AddHandler moDomains(lIdx).onDisconnect, AddressOf moDomains_onDisconnect
			AddHandler moDomains(lIdx).onError, AddressOf moDomains_onError

			'and then tell the socket to expect data
			moDomains(lIdx).MakeReadyToReceive()
		Else
			LogEvent(LogEventType.Informational, "Domain Connection Denied: Not Accepting!")
		End If
	End Sub

	Private Sub moDomainListener_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moDomainListener.onDataArrival
		'
	End Sub

	Private Sub moDomainListener_onDisconnect(ByVal Index As Integer) Handles moDomainListener.onDisconnect
		'
	End Sub

	Private Sub moDomainListener_onError(ByVal Index As Integer, ByVal Description As String) Handles moDomainListener.onError
		'
	End Sub

	Private Sub moDomainListener_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moDomainListener.onSendComplete
		'
	End Sub
#End Region

#Region "Domain Server Connections Handling"
	Private Sub moDomains_onConnect(ByVal Index As Integer)
		'
	End Sub

	Private Sub moDomains_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket)
		'
	End Sub

	Private Sub moDomains_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer)
		'ok, gotta get the data...
		'Data() contains our data... let's dissect it...
		Dim iMsgID As Int16

		'LogEvent(LogEventType.Informational, "Message Received From Domain " & Index)

		iMsgID = System.BitConverter.ToInt16(Data, 0)
		If mb_MONITOR_MSG_ACTIVITY = True Then moMonitor.AddInMsg(iMsgID, MsgMonitor.eMM_AppType.RegionServer, Data.Length, Index)

        Select Case iMsgID
            'Case GlobalMessageCode.eReportWarpoints
            '    HandleDomainRewardWarpoints(Data, Index)
            'Case GlobalMessageCode.eRewardWarpoints
            '    HandleDomainRewardWarpoints(Data, Index)
            Case GlobalMessageCode.eRequestObject
                'LogEvent(LogEventType.Informational, "Domain Requesting Object")
                HandleDomainRequestObject(Data, Index)
            Case GlobalMessageCode.eRequestPlayerList
                LogEvent(LogEventType.Informational, "Domain Requesting Player List")
                HandleGetPlayerList(moDomains(Index))
            Case GlobalMessageCode.eRequestPlayerRelList
                LogEvent(LogEventType.Informational, "Domain Requesting Player Rel List")
                HandleGetPlayerRelList(moDomains(Index))
            Case GlobalMessageCode.eDomainRequestEnvirObjects
                'LogEvent(LogEventType.Informational, "Domain Requesting Envir Objects")
                HandleDomainRequestEnvirObjs(Data, moDomains(Index), moServers(Index))
            Case GlobalMessageCode.eUndockCommandFinished
                HandleUndockCommandFinished(Data, Index)
            Case GlobalMessageCode.eRegisterDomain
                'LogEvent(LogEventType.Informational, "Domain Registering")
                HandleRegisterDomain(Data, Index)
            Case GlobalMessageCode.ePathfindingConnectionInfo
                LogEvent(LogEventType.Informational, "Sending Pathfinding Connection to Domain " & Index)
                moDomains(Index).SendData(myPathfindingData)
            Case GlobalMessageCode.eRequestGalaxyAndSystems
                LogEvent(LogEventType.Informational, "Domain Requesting Galaxy Map")
                HandleRequestGalaxy(Index, ConnectionType.eRegionServerApp)
            Case GlobalMessageCode.eRequestStarTypes
                LogEvent(LogEventType.Informational, "Domain Requesting Star Types")
                HandleRequestStarTypes(Index, ConnectionType.eRegionServerApp)
            Case GlobalMessageCode.eRequestSystemDetails
                LogEvent(LogEventType.Informational, "Domain Requesting System Details - " & System.BitConverter.ToInt32(Data, 2).ToString)
                HandleRequestSystemDetails(Index, ConnectionType.eRegionServerApp, Data)
            Case GlobalMessageCode.eDomainServerReady
                LogEvent(LogEventType.Informational, "Domain Indicating Ready State")
                HandleServerReady(Data, Index)
            Case GlobalMessageCode.eMoveObjectRequest, GlobalMessageCode.eAddWaypointMsg
                CheckForCargoRouteCancel(Data)
            Case GlobalMessageCode.eRemoveObject
                'LogEvent(GlobalVars.LogEventType.Informational, "Remove Object Message Received")
                HandleRemoveObject(Data)
            Case GlobalMessageCode.eRequestUndock
                HandleRequestUndock(Data)
            Case GlobalMessageCode.eMoveLockViolate
                HandleMoveLockViolate(Data)
            Case GlobalMessageCode.eHangarCargoBayDestroyed
                HandleHangarCargoDestroyed(Data)
            Case GlobalMessageCode.eDockCommand
                HandleDockCommand(Data)
            Case GlobalMessageCode.eSetEntityAI
                HandleSetEntityAI(Data)
            Case GlobalMessageCode.eReloadWpnMsg
                HandleReloadWpnMsg(Data, Index)
            Case GlobalMessageCode.eRequestPlayerStartLoc
                HandleRequestPlayerStartLoc(Data)
            Case GlobalMessageCode.eServerShutdown
                HandleDomainServerShutdown(Index)
            Case GlobalMessageCode.eUpdateEntityAndSave
                HandleUpdateAndSave(Data, Index)
            Case GlobalMessageCode.ePlayerAlert
                HandlePlayerAlert(Data)
            Case GlobalMessageCode.eGetPirateStartLoc
                HandleGetPirateStartLoc(Data, Index)
            Case GlobalMessageCode.ePlacePirateAssets
                HandlePlacePirateAssets(Data, Index)
            Case GlobalMessageCode.eUpdateEntityAttrs
                HandleUpdateEntityAttrs(Data)
            Case GlobalMessageCode.eSetPlayerRel
                HandleFirstContact(Data)
            Case GlobalMessageCode.eMsgMonitorData
                HandleMsgMonitorMsg(Data, Index, MsgMonitor.eMM_AppType.RegionServer)
            Case GlobalMessageCode.eUpdateUnitExpLevel
                HandleUpdateUnitExpLevel(Data)
            Case GlobalMessageCode.eSetRepairTarget, GlobalMessageCode.eSetDismantleTarget
                HandleSetMaintenanceTarget(Data, iMsgID, False)
            Case GlobalMessageCode.ePlayerDiscoversWormhole
                HandlePlayerDiscoversWormhole(Data)
            Case GlobalMessageCode.eNewsItem
                HandleNewsItem(Data)
            Case GlobalMessageCode.eCreatePlanetInstance
                HandleCreatePlanetInstance(Data, Index)
            Case GlobalMessageCode.eSaveAndUnloadInstance
                HandleSaveAndUnloadInstance(Data, Index)
            Case GlobalMessageCode.eAILaunchAll
                HandleAILaunchAll(Data, Index)
            Case GlobalMessageCode.eAddObjectCommand
                HandleMineralCachePlacement(Data, Index)
            Case GlobalMessageCode.eSetEntityStatus
                HandleRegionSetEntityStatus(Data, Index)
            Case GlobalMessageCode.eShiftClickAddProduction
                HandleRegionShiftClickAddProduction(Data, Index)
            Case GlobalMessageCode.eRequestEntityDefenses
                HandleDomainRequestEntityDefenses(Data, Index)
            Case GlobalMessageCode.eRouteMoveCommand
                HandleDomainRouteMoveCommand(Data, Index)
            Case GlobalMessageCode.eMicroExplosion
                HandleMicroExplosion(Data, Index)
            Case GlobalMessageCode.eSetTetherPoint
                HandleSetTetherPoint(Data, Index)
        End Select
	End Sub

	Private Sub moDomains_onDisconnect(ByVal Index As Integer)
		LogEvent(LogEventType.Informational, "Domain " & Index & " disconnected.")
	End Sub

	Private Sub moDomains_onError(ByVal Index As Integer, ByVal Description As String)
		LogEvent(LogEventType.Informational, "Domain Connection Error (" & Index & "): " & Description)
	End Sub

	Private Sub moDomains_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer)
		'
	End Sub

#End Region

#Region "  Message Handlers  "
    Private Sub CheckForCargoRouteCancel(ByRef yData() As Byte)
        'MsgCode - 2, DestX - 4, DestZ - 4, DestA - 2, DestID - 4, DestTypeID - 2, GUID List...
        'We only really need to GUID List...
        Dim lPos As Int32 = 2
        Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'don't care about destA
        lPos += 2
        Dim lDestEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iDestEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lCnt As Int32 = (yData.Length - 18) \ 6

        Try
            For X As Int32 = 0 To lCnt - 1
                Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                If iObjTypeID = ObjectType.eUnit Then
                    Dim oUnit As Unit = GetEpicaUnit(lObjID)
                    If oUnit Is Nothing = False Then
                        With oUnit
                            .DestEnvirID = lDestEnvirID
                            .DestEnvirTypeID = iDestEnvirTypeID
                            .DestX = lDestX
                            .DestZ = lDestZ
                            .ClearDropoffAndSourceDetails()
                        End With
                    End If
                End If
            Next X
        Catch
            'Do nothing
        End Try
    End Sub
    Private Sub HandleAILaunchAll(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2 'for msgcode
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim oEntity As Epica_Entity = CType(GetEpicaObject(lObjID, iObjTypeID), Epica_Entity)
        If oEntity Is Nothing = False Then
            With oEntity
                If .Owner.AccountStatus <> AccountStatusType.eActiveAccount AndAlso .Owner.AccountStatus <> AccountStatusType.eOpenBetaAccount AndAlso .Owner.AccountStatus <> AccountStatusType.eTrialAccount AndAlso .Owner.AccountStatus <> AccountStatusType.eMondelisActive Then Return

                Dim lLastIdx As Int32 = -1
                For X As Int32 = 0 To .lHangarUB
                    If .lHangarIdx(X) <> -1 Then
                        lLastIdx = X
                    End If
                Next X

                For X As Int32 = 0 To .lHangarUB
                    If .lHangarIdx(X) <> -1 Then
                        If lLastIdx = X Then
                            AddToQueue(glCurrentCycle, QueueItemType.eAILaunchAll, .oHangarContents(X).ObjectID, .oHangarContents(X).ObjTypeID, .ObjectID, .ObjTypeID, 1, 0, 0, 0)
                        Else
                            AddToQueue(glCurrentCycle, QueueItemType.eAILaunchAll, .oHangarContents(X).ObjectID, .oHangarContents(X).ObjTypeID, .ObjectID, .ObjTypeID, 0, 0, 0, 0)
                        End If

                    End If
                Next X
            End With
        End If

    End Sub
    Private Sub HandleCreatePlanetInstance(ByRef yData() As Byte, ByVal lIndex As Int32)
        'region server is telling us the instance id
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lInstanceID As Int32 = System.BitConverter.ToInt32(yData, 6)

        Dim oPlanet As Planet = GetEpicaPlanet(lInstanceID)
        If oPlanet Is Nothing Then
            oPlanet = New Planet()
            With oPlanet
                .Atmosphere = 100
                .AxisAngle = 680
                .Gravity = 40
                .Hydrosphere = 0
                .InMyDomain = True
                .InnerRingRadius = -1
                .LocX = 100000
                .LocY = 0
                .LocZ = 0
                .lPrimaryComposition = 157
                .ObjectID = lInstanceID
                .ObjTypeID = ObjectType.ePlanet
                .oDomain = moServers(lIndex)
                .OuterRingRadius = -1
                .OwnerID = lPlayerID
                .ParentSystem = Nothing
                .PlanetName = StringToBytes("Tutorial One")
                .PlanetRadius = 1000
                .PlanetSizeID = 0
                .PlanetTypeID = PlanetType.eGeoPlastic
                .PlayerSpawns = 1
                .RingDiffuse = 0
                .RotationDelay = 120
                .SpawnLocked = False
                .SurfaceTemperature = 90
                .Vegetation = 0
                .ySentGNSLowRes = 255
            End With

            Dim lCurUB As Int32 = -1
            If glPlanetIdx Is Nothing = False Then lCurUB = Math.Min(glPlanetUB, glPlanetIdx.GetUpperBound(0))
            Dim bFound As Boolean = False
            For X As Int32 = 0 To lCurUB
                If glPlanetIdx(X) = -1 Then
                    goPlanet(X) = oPlanet
                    glPlanetIdx(X) = oPlanet.ObjectID
                    bFound = True
                    Exit For
                End If
            Next X
            If bFound = False Then
                Dim lTmpIdx As Int32 = glPlanetUB + 1
                ReDim Preserve goPlanet(lTmpIdx)
                ReDim Preserve glPlanetIdx(lTmpIdx)
                goPlanet(lTmpIdx) = oPlanet
                glPlanetIdx(lTmpIdx) = oPlanet.ObjectID
                glPlanetUB += 1
            End If
        End If

        'Ok, region server is also telling us the instance is ready... let's load it up...
        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing = False Then
            With oPlayer
                If .lStartedEnvirID < 1 Then
                    If .yPlayerPhase = eyPlayerPhase.eInitialPhase Then .blCredits = 15000000 Else .blCredits = 500000
                    .lLastViewedEnvir = lInstanceID
                    .iLastViewedEnvirType = ObjectType.ePlanet
                    .lStartedEnvirID = lInstanceID
                    .iStartedEnvirTypeID = ObjectType.ePlanet
                    .lStartLocX = 0
                    .lStartLocZ = 0
                    .lIronCurtainPlanet = .lStartedEnvirID

                    .DataChanged()

                    'Now, set up our starting unit
                    Dim lIdx As Int32 = -1
                    Dim oUnitDef As Epica_Entity_Def = GetEpicaUnitDef(27)
                    Dim oTmp As Unit = New Unit()
                    Dim X As Int32

                    'spawn the player's engineer
                    With oTmp
                        .bProducing = False
                        '.Cargo_Cap = oUnitDef.Cargo_Cap
                        .CurrentProduction = Nothing
                        .CurrentSpeed = 0

                        .CurrentStatus = 0
                        For X = 0 To oUnitDef.lSideCrits.Length - 1
                            .CurrentStatus = .CurrentStatus Or oUnitDef.lSideCrits(X)
                        Next X

                        .EntityDef = oUnitDef
                        oUnitDef.DefName.CopyTo(.EntityName, 0)
                        .ExpLevel = 0
                        .Fuel_Cap = oUnitDef.Fuel_Cap
                        .iCombatTactics = 514
                        .iTargetingTactics = 0
                        .LocAngle = 0
                        .LocX = 8100
                        .LocZ = -200
                        .ObjTypeID = ObjectType.eUnit
                        .Owner = oPlayer
                        .ParentObject = oPlanet
                        .Q1_HP = oUnitDef.Q1_MaxHP
                        .Q2_HP = oUnitDef.Q2_MaxHP
                        .Q3_HP = oUnitDef.Q3_MaxHP
                        .Q4_HP = oUnitDef.Q4_MaxHP
                        .Shield_HP = oUnitDef.Shield_MaxHP
                        .Structure_HP = oUnitDef.Structure_MaxHP
                        .yProductionType = oUnitDef.ProductionTypeID
                        .DataChanged()
                    End With

                    If oTmp.SaveObject() = True Then
                        'Now, find a suitable place...
                        lIdx = -1
                        'SyncLock goUnit
                        '    For X = 0 To glUnitUB
                        '        If glUnitIdx(X) = -1 Then
                        '            lIdx = X
                        '            Exit For
                        '        End If
                        '    Next X

                        '    If lIdx = -1 Then
                        '        glUnitUB += 1
                        '        ReDim Preserve glUnitIdx(glUnitUB)
                        '        ReDim Preserve goUnit(glUnitUB)
                        '        lIdx = glUnitUB
                        '    End If

                        '    goUnit(lIdx) = oTmp
                        '    glUnitIdx(lIdx) = oTmp.ObjectID
                        'End SyncLock
                        lIdx = AddUnitToGlobalArray(oTmp)
                    Else
                        LogEvent(LogEventType.CriticalError, "Unable to save the first unit during initializeplayer!")
                    End If

                    'Send our add object to the region server
                    CType(oTmp.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oTmp, GlobalMessageCode.eAddObjectCommand))

                    If oPlayer.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                        'then, let's create 15 tanks
                        oUnitDef = GetEpicaUnitDef(26)

                        For lUnitIdx As Int32 = 0 To 14
                            oTmp = New Unit()

                            'Ok, populate our values
                            With oTmp
                                .bProducing = False
                                '.Cargo_Cap = oUnitDef.Cargo_Cap
                                .CurrentProduction = Nothing
                                .CurrentSpeed = 0

                                .CurrentStatus = 0
                                For X = 0 To oUnitDef.lSideCrits.Length - 1
                                    .CurrentStatus = .CurrentStatus Or oUnitDef.lSideCrits(X)
                                Next X

                                .EntityDef = oUnitDef
                                oUnitDef.DefName.CopyTo(.EntityName, 0)
                                .ExpLevel = 0
                                .Fuel_Cap = oUnitDef.Fuel_Cap
                                .iCombatTactics = eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Normal
                                .iTargetingTactics = 0
                                .LocAngle = 0
                                .LocX = 8200 + ((lUnitIdx Mod 5) * 100)
                                .LocZ = (lUnitIdx \ 5) * 100
                                .ObjTypeID = ObjectType.eUnit
                                .Owner = oPlayer
                                .ParentObject = oPlanet
                                .Q1_HP = oUnitDef.Q1_MaxHP
                                .Q2_HP = oUnitDef.Q2_MaxHP
                                .Q3_HP = oUnitDef.Q3_MaxHP
                                .Q4_HP = oUnitDef.Q4_MaxHP
                                .Shield_HP = oUnitDef.Shield_MaxHP
                                .Structure_HP = oUnitDef.Structure_MaxHP
                                .yProductionType = oUnitDef.ProductionTypeID
                                .DataChanged()
                            End With

                            If oTmp.SaveObject() = True Then
                                'Now, find a suitable place...
                                lIdx = -1
                                'SyncLock goUnit
                                '    For X = 0 To glUnitUB
                                '        If glUnitIdx(X) = -1 Then
                                '            lIdx = X
                                '            Exit For
                                '        End If
                                '    Next X

                                '    If lIdx = -1 Then
                                '        glUnitUB += 1
                                '        ReDim Preserve glUnitIdx(glUnitUB)
                                '        ReDim Preserve goUnit(glUnitUB)
                                '        lIdx = glUnitUB
                                '    End If

                                '    goUnit(lIdx) = oTmp
                                '    glUnitIdx(lIdx) = oTmp.ObjectID
                                'End SyncLock
                                lIdx = AddUnitToGlobalArray(oTmp)
                            Else
                                LogEvent(LogEventType.CriticalError, "Unable to save the first unit during initializeplayer!")
                            End If

                            'Send our add object to the region server
                            CType(oTmp.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oTmp, GlobalMessageCode.eAddObjectCommand))
                        Next lUnitIdx

                        'Spawn the player's Pirate Turrets
                        'Turret1 120, -6500
                        'Turret2 -550, -6500
                        'Turret3 600, -7000
                        Dim oPirate As Player = GetEpicaPlayer(gl_HARDCODE_PIRATE_PLAYER_ID)
                        oUnitDef = GetEpicaFacilityDef(46)
                        Dim oFac As Facility
                        For lFacIdx As Int32 = 0 To 2
                            oFac = New Facility
                            With oFac
                                .bProducing = False
                                .CurrentProduction = Nothing
                                .CurrentSpeed = 0
                                .EntityDef = CType(oUnitDef, FacilityDef)
                                .CurrentStatus = 0
                                For Y As Int32 = 0 To oUnitDef.lSideCrits.Length - 1
                                    .CurrentStatus = .CurrentStatus Or oUnitDef.lSideCrits(Y)
                                Next Y

                                oUnitDef.DefName.CopyTo(.EntityName, 0)
                                .ExpLevel = 0
                                .Fuel_Cap = oUnitDef.Fuel_Cap
                                '.Hangar_Cap = oDefs(lDefIdx).Hangar_Cap
                                .iCombatTactics = eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Maximize_Damage
                                .iTargetingTactics = 0
                                .LocAngle = 0

                                Select Case lFacIdx
                                    Case 0
                                        .LocX = 120 : .LocZ = -6500
                                    Case 1
                                        .LocX = -550 : .LocZ = -6500
                                    Case 2
                                        .LocX = 600 : .LocZ = -7000
                                End Select

                                .Owner = oPirate
                                .ParentObject = oPlanet
                                .ObjTypeID = ObjectType.eFacility

                                .Q1_HP = oUnitDef.Q1_MaxHP
                                .Q2_HP = oUnitDef.Q2_MaxHP
                                .Q3_HP = oUnitDef.Q3_MaxHP
                                .Q4_HP = oUnitDef.Q4_MaxHP
                                .Shield_HP = oUnitDef.Shield_MaxHP
                                .Structure_HP = oUnitDef.Structure_MaxHP
                                .yProductionType = oUnitDef.ProductionTypeID

                                .DataChanged()

                                If .SaveObject() = True Then
                                    'Now, find a suitable place...
                                    lIdx = -1
                                    'SyncLock goFacility
                                    '    For Y As Int32 = 0 To glFacilityUB
                                    '        If glFacilityIdx(Y) = -1 Then
                                    '            lIdx = Y
                                    '            Exit For
                                    '        End If
                                    '    Next Y

                                    '    If lIdx = -1 Then
                                    '        ReDim Preserve glFacilityIdx(glFacilityUB + 1)
                                    '        ReDim Preserve goFacility(glFacilityUB + 1)
                                    '        glFacilityUB += 1
                                    '        lIdx = glFacilityUB
                                    '    End If

                                    '    goFacility(lIdx) = oFac
                                    '    glFacilityIdx(lIdx) = oFac.ObjectID
                                    'End SyncLock

                                    lIdx = AddFacilityToGlobalArray(oFac)

                                    .CurrentStatus = goFacility(lIdx).CurrentStatus Or elUnitStatus.eFacilityPowered
                                    .ServerIndex = lIdx
                                    .DataChanged()
                                    .SaveObject()
                                Else
                                    LogEvent(LogEventType.CriticalError, "HandleCreatePlanetInstance spawn turrets failed.")
                                End If
                            End With
                            CType(oFac.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oFac, GlobalMessageCode.eAddObjectCommand))
                            oFac = Nothing
                        Next lFacIdx

                        'spawn the player's pirate medium facility
                        oUnitDef = GetEpicaFacilityDef(1558)
                        oFac = New Facility
                        With oFac
                            .bProducing = False
                            .CurrentProduction = Nothing
                            .CurrentSpeed = 0
                            .EntityDef = CType(oUnitDef, FacilityDef)
                            .CurrentStatus = 0
                            For Y As Int32 = 0 To oUnitDef.lSideCrits.Length - 1
                                .CurrentStatus = .CurrentStatus Or oUnitDef.lSideCrits(Y)
                            Next Y

                            oUnitDef.DefName.CopyTo(.EntityName, 0)
                            .ExpLevel = 0
                            .Fuel_Cap = oUnitDef.Fuel_Cap
                            .iCombatTactics = eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Maximize_Damage
                            .iTargetingTactics = 0
                            .LocAngle = 0

                            'Medium Facility loc is -3600, -13000
                            .LocX = -3600 : .LocZ = -13000

                            .Owner = oPirate
                            .ParentObject = oPlanet
                            .ObjTypeID = ObjectType.eFacility

                            .Q1_HP = oUnitDef.Q1_MaxHP
                            .Q2_HP = oUnitDef.Q2_MaxHP
                            .Q3_HP = oUnitDef.Q3_MaxHP
                            .Q4_HP = oUnitDef.Q4_MaxHP
                            .Shield_HP = oUnitDef.Shield_MaxHP
                            .Structure_HP = oUnitDef.Structure_MaxHP \ 3
                            .yProductionType = oUnitDef.ProductionTypeID

                            .DataChanged()

                            If .SaveObject() = True Then
                                'Now, find a suitable place...
                                lIdx = -1
                                'SyncLock goFacility
                                '    For Y As Int32 = 0 To glFacilityUB
                                '        If glFacilityIdx(Y) = -1 Then
                                '            lIdx = Y
                                '            Exit For
                                '        End If
                                '    Next Y

                                '    If lIdx = -1 Then
                                '        ReDim Preserve glFacilityIdx(glFacilityUB + 1)
                                '        ReDim Preserve goFacility(glFacilityUB + 1)
                                '        glFacilityUB += 1
                                '        lIdx = glFacilityUB
                                '    End If

                                '    goFacility(lIdx) = oFac
                                '    glFacilityIdx(lIdx) = oFac.ObjectID
                                'End SyncLock
                                lIdx = AddFacilityToGlobalArray(oFac)

                                .CurrentStatus = goFacility(lIdx).CurrentStatus Or elUnitStatus.eFacilityPowered
                                .ServerIndex = lIdx
                                .DataChanged()
                                .SaveObject()
                            Else
                                LogEvent(LogEventType.CriticalError, "HandleCreatePlanetInstance spawn turrets failed.")
                            End If
                            CType(oFac.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oFac, GlobalMessageCode.eAddObjectCommand))
                        End With

                        'spawn the enochine cache
                        'enochine is -13000, 14500
                        Dim lCacheIdx As Int32 = AddMineralCache(lInstanceID, ObjectType.ePlanet, MineralCacheType.eMineable, 100, 1000000, -13000, 14500, GetEpicaMineral(157))
                        If lCacheIdx = -1 Then
                            LogEvent(LogEventType.CriticalError, "HandleCreatePlanetInstance spawn Enochine failed.")
                        Else
                            If glMineralCacheIdx(lCacheIdx) > -1 Then
                                Dim oCache As MineralCache = goMineralCache(lCacheIdx)
                                If oCache Is Nothing = False Then
                                    moDomains(lIndex).SendData(Me.GetAddObjectMessage(oCache, GlobalMessageCode.eAddObjectCommand_CE))
                                Else
                                    LogEvent(LogEventType.CriticalError, "HandleCreatePlanetInstance spawn Enochine failed. oCache=nothing")
                                End If
                            Else
                                LogEvent(LogEventType.CriticalError, "HandleCreatePlanetInstance spawn Enochine failed.. glMineralCacheIdx=-1.")
                            End If
                        End If
                    End If
                Else
                    'Ok, everything is already spawned... we need to load it...
                    AddToQueue(glCurrentCycle + 30, QueueItemType.eLoadInstancedPlanet, lInstanceID, lPlayerID, -1, -1, 0, 0, 0, 0)
                End If
                .lLastViewedEnvir = lInstanceID
                .iLastViewedEnvirType = ObjectType.ePlanet
            End With


            Dim yResp(5) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerLoginResponse).CopyTo(yResp, 0)
            System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yResp, 2)
            oPlayer.SendPlayerMessage(yResp, False, 0)

            oPlayer.SaveObject(False)

        End If
    End Sub
    Private Sub HandleDockCommand(ByVal yData() As Byte)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)

        Dim oEntity As Epica_Entity = CType(GetEpicaObject(lObjID, iObjTypeID), Epica_Entity)
        'Dim oParent As Epica_Entity = CType(oEntity.ParentObject, Epica_Entity)
        Dim oParent As Epica_Entity = CType(GetEpicaObject(lParentID, iParentTypeID), Epica_Entity)
        'Dim X As Int32

        'Region Server is telling us that the docking procedure is done... 

        'LogEvent(LogEventType.ExtensiveLogging, "HandleDockCommand received for " & lObjID & ", " & iObjTypeID & " to " & lParentID & ", " & iParentTypeID)

        If oEntity Is Nothing = False Then
            If (oEntity.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then
                oEntity.CurrentStatus = oEntity.CurrentStatus Xor elUnitStatus.eMoveLock
            End If

            'need to put the object in the oEnvir
            oParent.AddHangarRef(CType(oEntity, Epica_GUID))

            'now, reduce the envir obj's Hangar
            'If oEntity.ObjTypeID = ObjectType.eUnit Then
            '    oParent.Hangar_Cap -= CType(oEntity, Unit).EntityDef.HullSize
            'Else : oParent.Hangar_Cap -= CType(oEntity, Facility).EntityDef.HullSize
            'End If

            oEntity.ParentObject = oParent

            'Special things occur when a unit docks with a facility...
            If oEntity.ObjTypeID = ObjectType.eUnit AndAlso oParent.ObjTypeID = ObjectType.eFacility Then

                'Ok, before anything else, check if the production type can produce AMMO
                'oParent.ProduceResupply(oEntity)

                'Now... check our entity's Refinery Index is not -1
                If CType(oParent, Facility).AutoLaunch = True Then
                    With CType(oEntity, Unit)
                        If .lCurrentRouteIdx > -1 AndAlso .lCurrentRouteIdx <= .lRouteUB Then
                            'LogEvent(LogEventType.ExtensiveLogging, "Route HandleDockCommand: " & lObjID & ", " & iObjTypeID & " to " & lParentID & ", " & iParentTypeID)
                            CType(oEntity, Unit).CheckRouteArrival()
                        ElseIf oParent.yProductionType = ProductionType.eMining Then
                            'LogEvent(LogEventType.ExtensiveLogging, "HandleDockCommand forced launch because parent is mining facility.")
                            AddToQueue(glCurrentCycle, QueueItemType.eUndockAndReturnToRefinery_QIT, .ObjectID, .ObjTypeID, oParent.ObjectID, oParent.ObjTypeID, 0, 0, 0, 0)
                        ElseIf .lRouteUB <> -1 Then
                            LogEvent(LogEventType.Warning, "HandleDockCommand entity has route but may not be on route. (" & lObjID & ", " & iObjTypeID & ")")
                        End If
                    End With

                    'If oParent.yProductionType = ProductionType.eRefining Then
                    '	'X = CType(oEntity, Unit).lRefineryIndex
                    '	'If X = -1 Then
                    '	'For X = 0 To glFacilityUB
                    '	'	If glFacilityIdx(X) = oParent.ObjectID Then
                    '	'		CType(oEntity, Unit).lRefineryIndex = X
                    '	'		Exit For
                    '	'	End If
                    '	'Next X
                    '	'End If
                    '	CType(oEntity, Unit).CheckRouteArrival()

                    '	'If X <> -1 Then
                    '	'	If glFacilityIdx(X) <> -1 AndAlso glFacilityIdx(X) = oParent.ObjectID Then
                    '	'		Dim oFac As Facility = goFacility(X)
                    '	'		If oFac Is Nothing = False Then
                    '	'			'ok, unload all our resources...
                    '	'			oEntity.TransferCargo(oParent) 

                    '	'			'Now, add to our queue
                    '	'			If CType(oEntity, Unit).bRoutePaused = False Then AddToQueue(glCurrentCycle, EngineCode.QueueItemType.eUndockAndReturnToMine_QIT, oEntity.ObjectID, oEntity.ObjTypeID, oParent.ObjectID, oParent.ObjTypeID)
                    '	'		End If
                    '	'		'Else
                    '	'		'	CType(oEntity, Unit).lRefineryIndex = -1
                    '	'	End If
                    '	'End If

                    'ElseIf oParent.yProductionType = ProductionType.eMining Then
                    '	'Ok, this facility is a mine...

                    '	CType(oEntity, Unit).CheckRouteArrival()

                    '	''ok... set up our entity's lMiningFacIndex
                    '	'For X = 0 To glFacilityUB
                    '	'	If glFacilityIdx(X) = oParent.ObjectID Then
                    '	'		'CType(oEntity, Unit).lMiningFacIndex = X
                    '	'		CType(oEntity, Unit).lCacheIndex = -1

                    '	'		'Now, tell that facility to transfer
                    '	'		oParent.TransferCargo(oEntity)

                    '	'		Exit For
                    '	'	End If
                    '	'Next X

                    '	''Now, check if the entity is full
                    '	'If oEntity.Cargo_Cap = 0 AndAlso CType(oEntity, Unit).bRoutePaused = False Then
                    '	'	AddToQueue(glCurrentCycle, EngineCode.QueueItemType.eUndockAndReturnToRefinery_QIT, oEntity.ObjectID, oEntity.ObjTypeID, oParent.ObjectID, oParent.ObjTypeID)
                    '	'End If

                    'End If
                End If
            End If
        Else
            LogEvent(LogEventType.Warning, "HandleDockCommand entity is nothing (" & lObjID & ", " & iObjTypeID & ").")
        End If
    End Sub
    Private Sub HandleDomainRequestEnvirObjs(ByVal yData() As Byte, ByRef oDomainSocket As NetSock, ByRef oDomainServer As DomainServer)
        Dim lObjID As Int32
        Dim iObjTypeID As Int16
        Dim X As Int32
        Dim oTemp As Object

        Dim yCache(200000) As Byte
        Dim yFinal() As Byte
        Dim lPos As Int32
        Dim lSingleMsgLen As Int32

        Dim yTemp() As Byte

        lObjID = System.BitConverter.ToInt32(yData, 2)
        iObjTypeID = System.BitConverter.ToInt16(yData, 6)

        If oDomainServer.bReceivedDefs = False OrElse oDomainSocket.SocketIndex = 0 Then System.GC.Collect()

        If oDomainServer.bReceivedDefs = False Then
            System.GC.Collect()

            'first transmit all unit defs
            For X = 0 To glUnitDefUB
                If glUnitDefIdx(X) <> -1 Then
                    oDomainSocket.SendData(GetAddObjectMessage(goUnitDef(X), GlobalMessageCode.eAddObjectCommand))
                    Dim yTmpMsg() As Byte = goUnitDef(X).GetCriticalHitChanceMsg()
                    If yTmpMsg Is Nothing = False Then oDomainSocket.SendData(yTmpMsg)
                    Threading.Thread.Sleep(1)
                End If
            Next X
            For X = 0 To glFacilityDefUB
                If glFacilityDefIdx(X) <> -1 Then
                    oDomainSocket.SendData(GetAddObjectMessage(goFacilityDef(X), GlobalMessageCode.eAddObjectCommand))
                    Dim yTmpMsg() As Byte = goFacilityDef(X).GetCriticalHitChanceMsg()
                    If yTmpMsg Is Nothing = False Then oDomainSocket.SendData(yTmpMsg)
                    Threading.Thread.Sleep(1)
                End If
            Next X
            oDomainServer.bReceivedDefs = True
        End If

        'then, transmit all units
        lPos = 0
        lSingleMsgLen = -1
        For X = 0 To glUnitUB
            If glUnitIdx(X) <> -1 Then
                Dim oUnit As Unit = goUnit(X)
                If oUnit Is Nothing = False Then
                    oTemp = oUnit.ParentObject
                    If oTemp Is Nothing = False Then
                        If CType(oTemp, Epica_GUID).ObjectID = lObjID AndAlso CType(oTemp, Epica_GUID).ObjTypeID = iObjTypeID Then
                            yTemp = GetAddObjectMessage(oUnit, GlobalMessageCode.eAddObjectCommand)
                            lSingleMsgLen = yTemp.Length
                            'Ok, before we continue, check if we need to increase our cache
                            If lPos + lSingleMsgLen + 2 > yCache.Length Then
                                'increase it
                                ReDim Preserve yCache(yCache.Length + 200000)
                            End If
                            System.BitConverter.GetBytes(CShort(lSingleMsgLen)).CopyTo(yCache, lPos)
                            lPos += 2
                            'GetAddObjectMessage(goUnit(X), EpicaMessageCode.eAddObjectCommand).CopyTo(yCache, lPos)
                            yTemp.CopyTo(yCache, lPos)
                            lPos += lSingleMsgLen
                        End If
                    End If
                End If
            End If
        Next X
        If lPos <> 0 Then
            ReDim yFinal(lPos - 1)
            Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
            oDomainSocket.SendLenAppendedData(yFinal)
            Threading.Thread.Sleep(1)
            If MsgSystem.mb_MONITOR_MSG_ACTIVITY = True Then goMsgSys.moMonitor.AddOutMsg(GlobalMessageCode.eAddObjectCommand, MsgMonitor.eMM_AppType.RegionServer, yFinal.Length, oDomainSocket.SocketIndex)
        End If

        'next, transmit all facilities
        'next, transmit all facilities
        ReDim yCache(200000)
        ReDim yFinal(0)
        lPos = 0
        lSingleMsgLen = -1
        For X = 0 To glFacilityUB
            If glFacilityIdx(X) <> -1 AndAlso goFacility(X) Is Nothing = False Then
                oTemp = goFacility(X).ParentObject
                If oTemp Is Nothing = False Then
                    If CType(oTemp, Epica_GUID).ObjectID = lObjID AndAlso CType(oTemp, Epica_GUID).ObjTypeID = iObjTypeID Then
                        'Ok, one more test, we dont send NPC Tradeposts...
                        With goFacility(X)
                            If .Owner Is Nothing = False AndAlso .Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID AndAlso (.yProductionType = ProductionType.eTradePost OrElse iObjTypeID = ObjectType.eSolarSystem) Then Continue For
                        End With

                        yTemp = GetAddObjectMessage(goFacility(X), GlobalMessageCode.eAddObjectCommand)
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
                End If
            End If
        Next X
        If lPos <> 0 Then
            ReDim yFinal(lPos - 1)
            Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
            oDomainSocket.SendLenAppendedData(yFinal)
            Threading.Thread.Sleep(1)
            If MsgSystem.mb_MONITOR_MSG_ACTIVITY = True Then goMsgSys.moMonitor.AddOutMsg(GlobalMessageCode.eAddObjectCommand, MsgMonitor.eMM_AppType.RegionServer, yFinal.Length, oDomainSocket.SocketIndex)
        End If
        'For X = 0 To glFacilityUB
        '	If glFacilityIdx(X) <> -1 Then
        '		oTemp = goFacility(X).ParentObject
        '		If oTemp Is Nothing = False Then
        '			If CType(oTemp, Epica_GUID).ObjectID = lObjID AndAlso CType(oTemp, Epica_GUID).ObjTypeID = iObjTypeID Then
        '				oDomainSocket.SendData(GetAddObjectMessage(goFacility(X), GlobalMessageCode.eAddObjectCommand))
        '			End If
        '		End If
        '	End If
        'Next X

        'after that, transmit all caches
        ReDim yCache(200000)
        ReDim yFinal(0)
        lPos = 0
        lSingleMsgLen = -1
        For X = 0 To glMineralCacheUB
            If glMineralCacheIdx(X) <> -1 Then
                oTemp = goMineralCache(X).ParentObject
                If oTemp Is Nothing = False Then
                    If CType(oTemp, Epica_GUID).ObjectID = lObjID AndAlso CType(oTemp, Epica_GUID).ObjTypeID = iObjTypeID Then
                        yTemp = GetAddObjectMessage(goMineralCache(X), GlobalMessageCode.eAddObjectCommand)
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
                End If
            End If
        Next X
        If lPos <> 0 Then
            ReDim yFinal(lPos - 1)
            Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
            oDomainSocket.SendLenAppendedData(yFinal)
            Threading.Thread.Sleep(1)
            If MsgSystem.mb_MONITOR_MSG_ACTIVITY = True Then goMsgSys.moMonitor.AddOutMsg(GlobalMessageCode.eAddObjectCommand, MsgMonitor.eMM_AppType.RegionServer, yFinal.Length, oDomainSocket.SocketIndex)
        End If
        'For X = 0 To glMineralCacheUB
        '	If glMineralCacheIdx(X) <> -1 Then
        '		oTemp = goMineralCache(X).ParentObject
        '		If oTemp Is Nothing = False Then
        '			If CType(oTemp, Epica_GUID).ObjectID = lObjID AndAlso CType(oTemp, Epica_GUID).ObjTypeID = iObjTypeID Then
        '				oDomainSocket.SendData(GetAddObjectMessage(goMineralCache(X), GlobalMessageCode.eAddObjectCommand))
        '			End If
        '		End If
        '	End If
        'Next X

        'And the Component Caches
        ReDim yCache(200000)
        ReDim yFinal(0)
        lPos = 0
        lSingleMsgLen = -1
        For X = 0 To glComponentCacheUB
            If glComponentCacheIdx(X) <> -1 Then
                oTemp = goComponentCache(X).ParentObject
                If oTemp Is Nothing = False Then
                    If CType(oTemp, Epica_GUID).ObjectID = lObjID AndAlso CType(oTemp, Epica_GUID).ObjTypeID = iObjTypeID Then
                        yTemp = GetAddObjectMessage(goComponentCache(X), GlobalMessageCode.eAddObjectCommand)
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
                End If
            End If
        Next X
        If lPos <> 0 Then
            ReDim yFinal(lPos - 1)
            Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
            oDomainSocket.SendLenAppendedData(yFinal)
            Threading.Thread.Sleep(1)
            If MsgSystem.mb_MONITOR_MSG_ACTIVITY = True Then goMsgSys.moMonitor.AddOutMsg(GlobalMessageCode.eAddObjectCommand, MsgMonitor.eMM_AppType.RegionServer, yFinal.Length, oDomainSocket.SocketIndex)
        End If
        'For X = 0 To glComponentCacheUB
        '	If glComponentCacheIdx(X) <> -1 Then
        '		oTemp = goComponentCache(X).ParentObject
        '		If oTemp Is Nothing = False Then
        '			If CType(oTemp, Epica_GUID).ObjectID = lObjID AndAlso CType(oTemp, Epica_GUID).ObjTypeID = iObjTypeID Then
        '				oDomainSocket.SendData(GetAddObjectMessage(goComponentCache(X), GlobalMessageCode.eAddObjectCommand))
        '			End If
        '		End If
        '	End If
        'Next X


        'NOW FOR the Formation Defs
        ReDim yCache(200000)
        ReDim yFinal(0)
        lPos = 0
        lSingleMsgLen = -1
        For X = 0 To glFormationDefUB
            If glFormationDefIdx(X) <> -1 Then
                yTemp = goFormationDefs(X).GetAsAddMsg()
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
            oDomainSocket.SendLenAppendedData(yFinal)
            Threading.Thread.Sleep(1)
            If MsgSystem.mb_MONITOR_MSG_ACTIVITY = True Then goMsgSys.moMonitor.AddOutMsg(GlobalMessageCode.eAddObjectCommand, MsgMonitor.eMM_AppType.RegionServer, yFinal.Length, oDomainSocket.SocketIndex)
        End If

        'finally, reply with the original message
        oDomainSocket.SendData(yData)
    End Sub
    Private Sub HandleDomainRequestObject(ByVal Data() As Byte, ByVal lIndex As Int32)
        'Request object... what we need is a ObjectID and an Object Type ID... therefore...
        Dim lObjID As Int32
        Dim iObjTypeID As Int16
        Dim yResponse() As Byte
        Dim yTemp() As Byte = Nothing
        Dim oObj As Object

        'TODO: Figure out a better way for this I'm guessing...
        On Error Resume Next

        lObjID = System.BitConverter.ToInt32(Data, 2)
        iObjTypeID = System.BitConverter.ToInt16(Data, 6)

        oObj = GetEpicaObject(lObjID, iObjTypeID)
        If oObj Is Nothing = False Then
            yTemp = CType(CallByName(oObj, "GetObjAsString", CallType.Get), Byte())
        End If

        ReDim yResponse(yTemp.Length + 2)
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestObjectResponse).CopyTo(yResponse, 0)
        yTemp.CopyTo(yResponse, 2)

        moDomains(lIndex).SendData(yResponse)
    End Sub
    'Private Sub HandleDomainRewardWarpoints(ByVal yData() As Byte, ByVal lIndex As Int32)
    '    Dim lPos As Int32 = 0
    '    Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
    '    Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '    Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
    '    Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '    Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

    '    'locs...
    '    lPos += 8

    '    Dim lKillerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

    '    Dim lWPPos As Int32 = lPos
    '    Dim lWarpoints As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

    '    Dim lKilledOwnerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '    Dim iModelID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
    '    Dim lUnitCR As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

    '    If lKilledOwnerID = gl_HARDCODE_PIRATE_PLAYER_ID OrElse lKillerID = gl_HARDCODE_PIRATE_PLAYER_ID Then Return

    '    Dim oKilled As Player = GetEpicaPlayer(lKilledOwnerID)
    '    Dim oKiller As Player = GetEpicaPlayer(lKillerID)

    '    If oKilled Is Nothing OrElse oKiller Is Nothing Then Return

    '    'TODO: determine actual warpoint value based on whether they are guildies or allies or whatever... aliased... etc...

    '    Dim oKilledRel As PlayerRel = oKilled.GetPlayerRel(oKiller.ObjectID)
    '    Dim oKillerRel As PlayerRel = oKiller.GetPlayerRel(oKilled.ObjectID)

    '    'Ok, killer...
    '    If oKillerRel.lPlayersWPV = 0 Then
    '        lWarpoints = 0
    '    Else
    '        Dim fWPVAttrition As Single = ((oKillerRel.lPlayersWPV * 0.001F) + 1.0F) / ((oKilledRel.lPlayersWPV * 0.001F) + 1.0F)
    '        lWarpoints = CInt(lWarpoints * fWPVAttrition)
    '    End If
    '    If oKillerRel.blTotalWarpointsGained > (oKiller.blWarpointsAllTime \ 10) AndAlso oKillerRel.blTotalWarpointsGained <> oKiller.blWarpointsAllTime Then
    '        lWarpoints = 0
    '    End If
    '    'same guild gets 0
    '    If oKiller.lGuildID = oKilled.lGuildID AndAlso oKiller.lGuildID > 0 Then lWarpoints = 0
    '    'if killed is in faction with killer, 0 wp
    '    For X As Int32 = 0 To oKiller.lSlotID.GetUpperBound(0)
    '        If oKiller.lSlotID(X) = oKilled.ObjectID Then
    '            If (oKiller.ySlotState(X) And eySlotState.Accepted) <> 0 Then lWarpoints = 0
    '            Exit For
    '        End If
    '        If oKilled.lSlotID(X) = oKiller.ObjectID Then
    '            If (oKilled.ySlotState(X) And eySlotState.Accepted) <> 0 Then lWarpoints = 0
    '            Exit For
    '        End If
    '    Next X
    '    'if killed and killer have aliases to each other, 0 wp
    '    If oKilled.lAliasingPlayerID = oKiller.ObjectID OrElse oKiller.lAliasingPlayerID = oKilled.ObjectID Then lWarpoints = 0

    '    If iMsgCode = GlobalMessageCode.eRewardWarpoints Then
    '        oKillerRel.RegenerateWPV()
    '        oKilledRel.RegenerateWPV()

    '        oKiller.blWarpointsAllTime += lWarpoints

    '        oKiller.blWarpoints += lWarpoints
    '        oKillerRel.blTotalWarpointsGained += lWarpoints

    '        'So if a player kills some unit, the killer will lose (PlayerUnitCR/30) WPV. 
    '        oKillerRel.lPlayersWPV -= (lUnitCR \ 30)
    '        If oKillerRel.lPlayersWPV < 0 Then oKillerRel.lPlayersWPV = 0

    '        ''For every kill, the killed receives 500 WPV.
    '        'oKilledRel.lPlayersWPV += 500
    '        oKilledRel.lPlayersWPV += 1
    '        If oKilledRel.lPlayersWPV > 10000 Then oKilledRel.lPlayersWPV = 10000


    '        oKilledRel.bForceSaveMe = True
    '        oKillerRel.bForceSaveMe = True
    '    End If

    '    System.BitConverter.GetBytes(lWarpoints).CopyTo(yData, lWPPos)
    '    oKiller.SendPlayerMessage(yData, False, AliasingRights.eViewUnitsAndFacilities)
    'End Sub
    Private Sub HandleDomainRouteMoveCommand(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        'Domain server is indicating that it could not find this entity...
        If iObjTypeID = ObjectType.eUnit Then
            Dim oUnit As Unit = GetEpicaUnit(lObjID)
            If oUnit Is Nothing = False Then
                'Ok, it is there... add a retry event
                AddToQueue(glCurrentCycle + 150, QueueItemType.eRetryRouteItem, oUnit.ObjectID, -1, -1, -1, -1, -1, -1, -1)
            End If
        End If
    End Sub
    Private Sub HandleDomainServerShutdown(ByVal lDomainIdx As Int32)
        Dim X As Int32
        Dim bReady As Boolean = True

        moServers(lDomainIdx).bReportingShutdown = True

        'Are all servers shutdown?
        For X = 0 To mlDomainUB
            If moServers(X).bReportingShutdown = False Then
                bReady = False
                Exit For
            End If
        Next X

        If bReady = True Then
            BeginPrimarySave()
        End If
    End Sub
    Private Sub HandleFirstContact(ByRef yData() As Byte)
        Dim lPlayer1ID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lPlayer2ID As Int32 = System.BitConverter.ToInt32(yData, 6)
        Dim yScore As Byte = yData(10)
        Dim lEnvirID As Int32 = -1
        Dim iEnvirTypeID As Int16 = -1
        Dim lEntityID As Int32 = -1
        Dim iEntityTypeID As Int16 = -1

        Try
            lEnvirID = System.BitConverter.ToInt32(yData, 11)
            iEnvirTypeID = System.BitConverter.ToInt16(yData, 15)
            lEntityID = System.BitConverter.ToInt32(yData, 17)
            iEntityTypeID = System.BitConverter.ToInt16(yData, 21)
        Catch ex As Exception
            'do nothing
        End Try

        'First contact is a situation where the region server looked for a relationship but did not find one...
        '  in this case, it means the relationship does not exist. Now, yScore tells us what to set the new relationship to

        If lPlayer1ID = lPlayer2ID Then Return

        Dim oPlayer1 As Player = GetEpicaPlayer(lPlayer1ID)
        Dim oPlayer2 As Player = GetEpicaPlayer(lPlayer2ID)

        If oPlayer1 Is Nothing OrElse oPlayer2 Is Nothing Then Return


        Dim uWP(0) As PlayerComm.WPAttachment
        Dim sLoc As String = ""
        With uWP(0)
            .AttachNumber = 1
            .EnvirID = lEnvirID
            .EnvirTypeID = iEnvirTypeID
            Dim oEntity As Epica_Entity = GetEpicaEntity(lEntityID, iEntityTypeID)
            If oEntity Is Nothing = False Then
                .LocX = oEntity.LocX
                .LocZ = oEntity.LocZ
                .sWPName = "First Contact"
                .yWPNameBytes = StringToBytes(.sWPName)
            End If
            If .EnvirTypeID = ObjectType.ePlanet Then
                Dim oPlanet As Planet = GetEpicaPlanet(.EnvirID)
                If oPlanet Is Nothing = False Then sLoc = ", " & BytesToString(oPlanet.PlanetName) & ","
            ElseIf .EnvirTypeID = ObjectType.eSolarSystem Then
                Dim oSystem As SolarSystem = GetEpicaSystem(.EnvirID)
                If oSystem Is Nothing = False Then sLoc = ", " & BytesToString(oSystem.SystemName) & ","
            End If
        End With

        Dim oPC As PlayerComm = oPlayer1.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "We have made first contact with " & oPlayer2.sPlayerNameProper & ". The point where contact was made" & sLoc & " has been attached to this message.", "First Contact with " & oPlayer2.sPlayerNameProper, oPlayer1.ObjectID, GetDateAsNumber(Now), False, oPlayer1.sPlayerNameProper, uWP)
        If oPC Is Nothing = False Then
            oPlayer1.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
        End If
        oPC = Nothing

        oPC = oPlayer2.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "We have made first contact with " & oPlayer1.sPlayerNameProper & ". The point where contact was made" & sLoc & " has been attached to this message.", "First Contact with " & oPlayer1.sPlayerNameProper, oPlayer2.ObjectID, GetDateAsNumber(Now), False, oPlayer2.sPlayerNameProper, uWP)
        If oPC Is Nothing = False Then
            oPlayer2.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
        End If
        oPC = Nothing

        Dim yResp(27) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityName).CopyTo(yResp, 0)
        oPlayer2.GetGUIDAsString.CopyTo(yResp, 2)
        oPlayer2.PlayerName.CopyTo(yResp, 8)
        oPlayer1.SendPlayerMessage(yResp, False, 0)
        oPlayer1.GetGUIDAsString.CopyTo(yResp, 2)
        oPlayer1.PlayerName.CopyTo(yResp, 8)
        oPlayer2.SendPlayerMessage(yResp, False, 0)

        Dim yMsg2(22) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerRel).CopyTo(yMsg2, 0)
        System.BitConverter.GetBytes(lPlayer2ID).CopyTo(yMsg2, 2)
        System.BitConverter.GetBytes(lPlayer1ID).CopyTo(yMsg2, 6)
        yMsg2(10) = yScore
        System.BitConverter.GetBytes(lEnvirID).CopyTo(yMsg2, 11)
        System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yMsg2, 15)
        System.BitConverter.GetBytes(lEntityID).CopyTo(yMsg2, 17)
        System.BitConverter.GetBytes(iEntityTypeID).CopyTo(yMsg2, 21)

        Dim oTmpRel As PlayerRel = New PlayerRel
        oTmpRel.oPlayerRegards = oPlayer1
        oTmpRel.oThisPlayer = oPlayer2
        oTmpRel.WithThisScore = yScore
        oTmpRel.TargetScore = yScore
        oPlayer1.SetPlayerRel(lPlayer2ID, oTmpRel, True)
        oPlayer1.SendPlayerMessage(yMsg2, True, AliasingRights.eViewAgents Or AliasingRights.eViewDiplomacy)

        oTmpRel = Nothing
        oTmpRel = New PlayerRel
        oTmpRel.oPlayerRegards = oPlayer2
        oTmpRel.oThisPlayer = oPlayer1
        oTmpRel.WithThisScore = yScore
        oTmpRel.TargetScore = yScore
        oPlayer2.SetPlayerRel(lPlayer1ID, oTmpRel, True)
        oPlayer2.SendPlayerMessage(yData, True, AliasingRights.eViewAgents Or AliasingRights.eViewDiplomacy)

        For X As Int32 = 0 To mlDomainUB
            moDomains(X).SendData(yData)
            moDomains(X).SendData(yMsg2)
        Next X
    End Sub
    Private Sub HandleGetPirateStartLoc(ByRef yData() As Byte, ByVal Index As Int32)
        'MsgCode (2), PlayerID (4), X (4), Y (4)
        'Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lPlanetID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lX As Int32 = System.BitConverter.ToInt32(yData, 6)
        Dim lY As Int32 = System.BitConverter.ToInt32(yData, 10)

        If lX = Int32.MinValue OrElse lY = Int32.MinValue Then
            LogEvent(LogEventType.Warning, "GetPirateStartLoc returned no loc for Planet ID: " & lPlanetID)
            Return
        End If
        If goAureliusAI Is Nothing = False Then goAureliusAI.HandleGetPirateStartLoc(lPlanetID, lX, lY)

        'Dim yResp() As Byte = CreatePirateSpawn(lPlayerID, lX, lY)
        'If yResp Is Nothing = False Then moDomains(Index).SendData(yResp)
    End Sub
    Private Sub HandleGetPlayerList(ByRef oDomainSocket As NetSock)
        Dim X As Int32
        Dim yData() As Byte

        ''then, transmit all units
        Dim lPos As Int32 = 0
        Dim lSingleMsgLen As Int32 = -1
        Dim yCache(200000) As Byte
        Dim yTemp() As Byte

        For X = 0 To glPlayerUB
            If glPlayerIdx(X) <> -1 Then
                yTemp = GetAddObjectMessage(goPlayer(X), GlobalMessageCode.eAddObjectCommand)
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

                ReDim yTemp(11)
                Dim lTempPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerTechValue).CopyTo(yTemp, lTempPos) : lTempPos += 2
                System.BitConverter.GetBytes(goPlayer(X).ObjectID).CopyTo(yTemp, lTempPos) : lTempPos += 4
                System.BitConverter.GetBytes(PlayerSpecialAttributeID.eCPLimit).CopyTo(yTemp, lTempPos) : lTempPos += 2
                System.BitConverter.GetBytes(CInt(goPlayer(X).oSpecials.iCPLimit)).CopyTo(yTemp, lTempPos) : lTempPos += 4
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

                ReDim yTemp(13)
                For Y As Int32 = 0 To goPlayer(X).lWormholeUB
                    If goPlayer(X).oWormholes(Y) Is Nothing = False Then
                        lTempPos = 0
                        System.BitConverter.GetBytes(GlobalMessageCode.ePlayerDiscoversWormhole).CopyTo(yTemp, lTempPos) : lTempPos += 2
                        System.BitConverter.GetBytes(goPlayer(X).ObjectID).CopyTo(yTemp, lTempPos) : lTempPos += 4
                        System.BitConverter.GetBytes(goPlayer(X).oWormholes(Y).ObjectID).CopyTo(yTemp, lTempPos) : lTempPos += 4
                        System.BitConverter.GetBytes(goPlayer(X).oWormholes(Y).System1.ObjectID).CopyTo(yTemp, lTempPos) : lTempPos += 4

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
                Next Y

                'oDomainSocket.SendData(GetAddObjectMessage(goPlayer(X), GlobalMessageCode.eAddObjectCommand))
                'Dim yMsg(11) As Byte
                'Dim lPos As Int32 = 0
                'System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerTechValue).CopyTo(yMsg, lPos) : lPos += 2
                'System.BitConverter.GetBytes(goPlayer(X).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                'System.BitConverter.GetBytes(PlayerSpecialAttributeID.eCPLimit).CopyTo(yMsg, lPos) : lPos += 2
                'System.BitConverter.GetBytes(CInt(goPlayer(X).oSpecials.iCPLimit)).CopyTo(yMsg, lPos) : lPos += 4
                'oDomainSocket.SendData(yMsg)

                'ReDim yMsg(13)
                'For Y As Int32 = 0 To goPlayer(X).lWormholeUB
                '	If goPlayer(X).oWormholes(Y) Is Nothing = False Then
                '		lPos = 0
                '		System.BitConverter.GetBytes(GlobalMessageCode.ePlayerDiscoversWormhole).CopyTo(yMsg, lPos) : lPos += 2
                '		System.BitConverter.GetBytes(goPlayer(X).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                '		System.BitConverter.GetBytes(goPlayer(X).oWormholes(Y).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                '		System.BitConverter.GetBytes(goPlayer(X).oWormholes(Y).System1.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                '		oDomainSocket.SendData(yMsg)
                '	End If
                'Next Y
            End If
        Next X

        For X = 0 To glPlayerUB
            If glPlayerIdx(X) <> -1 Then
                For Y As Int32 = 0 To goPlayer(X).lAliasUB
                    If goPlayer(X).lAliasIdx(Y) <> -1 Then
                        ReDim yTemp(66)

                        Dim lTempPos As Int32 = 0
                        System.BitConverter.GetBytes(GlobalMessageCode.ePlayerAliasConfig).CopyTo(yTemp, lTempPos) : lTempPos += 2
                        yTemp(lTempPos) = 1 : lTempPos += 1
                        System.BitConverter.GetBytes(goPlayer(X).ObjectID).CopyTo(yTemp, lTempPos) : lTempPos += 4
                        System.BitConverter.GetBytes(goPlayer(X).lAliasIdx(Y)).CopyTo(yTemp, lTempPos) : lTempPos += 4
                        lTempPos += 12
                        With goPlayer(X).uAliasLogin(Y)
                            .yUserName.CopyTo(yTemp, lTempPos) : lTempPos += 20
                            .yPassword.CopyTo(yTemp, lTempPos) : lTempPos += 20
                            System.BitConverter.GetBytes(.lRights).CopyTo(yTemp, lTempPos) : lTempPos += 4
                        End With
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

                        'Dim yMsg(65) As Byte
                        'Dim lPos As Int32 = 0
                        'System.BitConverter.GetBytes(GlobalMessageCode.ePlayerAliasConfig).CopyTo(yMsg, lPos) : lPos += 2
                        'System.BitConverter.GetBytes(goPlayer(X).ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                        'System.BitConverter.GetBytes(goPlayer(X).lAliasIdx(Y)).CopyTo(yMsg, lPos) : lPos += 4
                        'lPos += 12
                        'With goPlayer(X).uAliasLogin(Y)
                        '	.yUserName.CopyTo(yMsg, lPos) : lPos += 20
                        '	.yPassword.CopyTo(yMsg, lPos) : lPos += 20
                        '	System.BitConverter.GetBytes(.lRights).CopyTo(yMsg, lPos) : lPos += 4
                        'End With
                        'oDomainSocket.SendData(yMsg)
                    End If
                Next Y
            End If
        Next X


        If lPos <> 0 Then
            Dim yFinal(lPos - 1) As Byte
            Array.Copy(yCache, 0, yFinal, 0, lPos - 1)
            oDomainSocket.SendLenAppendedData(yFinal)
            Threading.Thread.Sleep(1)
            If MsgSystem.mb_MONITOR_MSG_ACTIVITY = True Then goMsgSys.moMonitor.AddOutMsg(GlobalMessageCode.eAddObjectCommand, MsgMonitor.eMM_AppType.RegionServer, yFinal.Length, oDomainSocket.SocketIndex)
        End If

        'Now, send back one final message indicating the end of the list
        ReDim yData(1)
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerList).CopyTo(yData, 0)
        oDomainSocket.SendData(yData)
    End Sub
	Private Sub HandleGetPlayerRelList(ByRef oDomainSocket As NetSock)
		Dim X As Int32
		Dim Y As Int32
		Dim yData() As Byte

		For X = 0 To glPlayerUB
            If glPlayerIdx(X) <> -1 Then

                Dim oPlayer As Player = goPlayer(X)
                If oPlayer Is Nothing = False AndAlso oPlayer.InMyDomain = True Then
                    goMsgSys.SendPlayerCPLimitUpdate(oPlayer.ObjectID, oPlayer.oSpecials.iCPLimit)
                End If

                For Y = 0 To goPlayer(X).PlayerRelUB
                    Dim oTmpRel As PlayerRel = goPlayer(X).GetPlayerRelByIndex(Y)
                    If oTmpRel Is Nothing = False Then oDomainSocket.SendData(GetAddPlayerRelMessage(oTmpRel))
                Next Y
            End If
		Next X

		'and finally, send back an End of List message
		ReDim yData(1)
		System.BitConverter.GetBytes(GlobalMessageCode.eRequestPlayerRelList).CopyTo(yData, 0)
		oDomainSocket.SendData(yData)
	End Sub
    Private Sub HandleHangarCargoDestroyed(ByVal yData() As Byte)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        'Indicates what was destroyed... (i hope)
        Dim lStatus As Int32 = System.BitConverter.ToInt32(yData, 8)

        Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, 12)
        Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, 16)

        Dim oEntity As Epica_Entity = CType(GetEpicaObject(lObjID, iObjTypeID), Epica_Entity)

        If oEntity Is Nothing Then Exit Sub

        If (lStatus And elUnitStatus.eHangarOperational) <> 0 Then
            'hangar was destroyed
            If (oEntity.CurrentStatus And elUnitStatus.eHangarOperational) <> 0 Then oEntity.CurrentStatus = oEntity.CurrentStatus Xor elUnitStatus.eHangarOperational ' oEntity.CurrentStatus -= elUnitStatus.eHangarOperational
            'oEntity.Hangar_Cap = 0
            oEntity.HangarDestroyed(lLocX, lLocZ)
        End If

        If (lStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
            If (oEntity.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then oEntity.CurrentStatus = oEntity.CurrentStatus Xor elUnitStatus.eCargoBayOperational 'oEntity.CurrentStatus -= elUnitStatus.eCargoBayOperational
            'oEntity.Cargo_Cap = 0
            oEntity.CargoDestroyed(lLocX, lLocZ)
        End If

    End Sub
    Private Sub HandleMicroExplosion(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lOwnerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'Ok get the unit
        If iTypeID = ObjectType.eUnit Then

            Dim oPlayer As Player = GetEpicaPlayer(lOwnerID)
            Dim bInitial As Boolean = oPlayer.bSentMicroExplosionInitial
            oPlayer.bSentMicroExplosionInitial = True

            Dim oUnit As Unit = GetEpicaUnit(lObjID)
            Dim sUnitName As String = ""
            If oUnit Is Nothing = False Then
                sUnitName = BytesToString(oUnit.EntityName)
            End If

            'Now, get the parent environment
            Dim oPlanet As Planet = Nothing
            If iParentTypeID = ObjectType.ePlanet Then
                'Ok, this planet is the dest
                oPlanet = GetEpicaPlanet(lParentID)
            ElseIf iParentTypeID = ObjectType.eSolarSystem Then
                Dim oSys As SolarSystem = GetEpicaSystem(lParentID)
                If oSys Is Nothing = False Then
                    oPlanet = oSys.GetNearestPlanet(lX, lZ)
                End If
            End If

            Dim oSB As New System.Text.StringBuilder
            If bInitial = False Then
                'ok, player has not received the initial msg
                If sUnitName = "" Then sUnitName = "a unit"
                oSB.AppendLine("Fleet command is reporting a message received from " & sUnitName & ".")
                oSB.AppendLine()
                oSB.AppendLine("""... Commander, we are experiencing massive levels of quantum flux permeating throughout the hull. Hull integrity is failing... (lost signal)""")
                oSB.AppendLine()

                If oPlanet Is Nothing = False Then
                    oSB.AppendLine("Following the signal loss, we detected what is believed to be an escape pod landing on the surface of " & BytesToString(oPlanet.PlanetName) & ".")
                    oSB.AppendLine()
                End If

                oSB.AppendLine("This could be a major defect in the designs of our units. Our analysts are researching into possible causes now.")
            Else
                oSB.AppendLine("Another report has come in from a unit encountering massive levels of quantum flux causing a catastrophic event.")
                If sUnitName <> "" Then oSB.AppendLine(vbCrLf & "The unit in question, " & sUnitName & ", seemed to be destroyed almost instantly.")

                If oPlanet Is Nothing = False Then
                    oSB.AppendLine(vbCrLf & "Witnesses say that an object was seen falling from the explosion towards the surface of " & BytesToString(oPlanet.PlanetName) & ".")
                End If
            End If

            Dim uAttach(0) As PlayerComm.WPAttachment
            With uAttach(0)
                .AttachNumber = 0
                .EnvirID = lParentID
                .EnvirTypeID = iParentTypeID
                .LocX = lX
                .LocZ = lZ
                .sWPName = sUnitName
            End With
            Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Strange Event", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, uAttach)
            If oPC Is Nothing = False Then
                oPlayer.SendPlayerMessage(GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If

            'Now, create our mineral cache
            If oPlanet Is Nothing = False Then
                Dim yMsg(21) As Byte
                Dim lMineralID As Int32 = 41991
                lPos = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eAddObjectCommand).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(lMineralID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(ObjectType.eMineral).CopyTo(yMsg, lPos) : lPos += 2
                oPlanet.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                System.BitConverter.GetBytes(1I).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(1I).CopyTo(yMsg, lPos) : lPos += 4

                oPlanet.oDomain.DomainSocket.SendData(yMsg)
            End If
        End If

    End Sub
    Private Sub HandleMineralCachePlacement(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2   'for msgcode
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lMaxQty As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lMaxConc As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oMineral As Mineral = GetEpicaMineral(lObjID)
        If oMineral Is Nothing Then Return

        'Dim lQrtrVal As Int32 = lMaxConc \ 4
        'Dim lConc As Int32 = CInt(Rnd() * lQrtrVal) + (lQrtrVal * 3)
        'lQrtrVal = lMaxQty \ 4
        'Dim lQty As Int32 = CInt(Rnd() * lQrtrVal) + (lQrtrVal * 3)
        Dim lCacheIdx As Int32 = AddMineralCache(lEnvirID, iEnvirTypeID, MineralCacheType.eMineable, lMaxConc, lMaxQty, lLocX, lLocZ, oMineral)
        If lCacheIdx <> -1 Then
            moDomains(lIndex).SendData(goMsgSys.GetAddObjectMessage(goMineralCache(lCacheIdx), GlobalMessageCode.eAddObjectCommand))
        End If

    End Sub
    Private Sub HandleMoveLockViolate(ByVal yData() As Byte)
        'the domain server is reporting to me that an entity represented in this message
        '  was flagged as building/mining/being docked to, and was ordered to move...
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        Dim oObj As Epica_Entity = CType(GetEpicaObject(lObjID, iObjTypeID), Epica_Entity)
        'Dim X As Int32

        gfrmDisplayForm.AddEventLine("Move Lock Violation")

        If oObj Is Nothing = False Then
            If (oObj.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then
                'Ok, supposedly it was right...
                oObj.CurrentStatus = oObj.CurrentStatus Xor elUnitStatus.eMoveLock ' oObj.CurrentStatus -= elUnitStatus.eMoveLock
            End If

            'ok, determine what the object was doing that put it in move lock
            If oObj.bProducing = True Then
                'it was building something
                oObj.bProducing = False
                'At the minimum, refund the credits
                If oObj.CurrentProduction Is Nothing = False AndAlso oObj.CurrentProduction.ProdCost Is Nothing = False Then
                    oObj.Owner.blCredits += oObj.CurrentProduction.ProdCost.CreditCost
                    'For X = 0 To glColonyUB
                    '    If glColonyIdx(X) <> -1 AndAlso goColony(X).Owner.ObjectID = oObj.Owner.ObjectID Then
                    '        If CType(goColony(X).ParentObject, Epica_GUID).ObjectID = CType(oObj.ParentObject, Epica_GUID).ObjectID Then
                    '            If CType(goColony(X).ParentObject, Epica_GUID).ObjTypeID = CType(oObj.ParentObject, Epica_GUID).ObjTypeID Then
                    '                With oObj.CurrentProduction.ProdCost
                    '                    goColony(X).AddNonWorkers(.ColonistCost)
                    '                End With
                    '                Exit For
                    '            End If
                    '        End If
                    '    End If
                    'Next X
                End If
                If oObj.CurrentProduction.ProductionTypeID = ObjectType.eFacilityDef Then
                    Dim oFacDef As FacilityDef = GetEpicaFacilityDef(oObj.CurrentProduction.ProductionID)
                    If oFacDef.ProductionTypeID = ProductionType.eCommandCenterSpecial OrElse oFacDef.ProductionTypeID = ProductionType.eTradePost Then
                        With CType(oObj.ParentObject, Epica_GUID)
                            Dim lColonyIdx As Int32 = oObj.Owner.GetColonyFromParent(.ObjectID, .ObjTypeID)
                            If lColonyIdx > -1 AndAlso glColonyIdx(lColonyIdx) <> -1 Then
                                Dim oColony As Colony = goColony(lColonyIdx)
                                If oColony Is Nothing = False Then
                                    If oFacDef.ProductionTypeID = ProductionType.eCommandCenterSpecial Then
                                        oColony.bCCInProduction = False
                                    ElseIf oFacDef.ProductionTypeID = ProductionType.eTradePost Then
                                        oColony.bTradepostInProduction = False
                                    End If
                                End If
                            End If
                        End With
                    End If
                End If
                oObj.CurrentProduction = Nothing
            ElseIf oObj.bMining = True Then
                oObj.bMining = False

                If oObj.lCacheIndex <> -1 AndAlso glMineralCacheIdx(oObj.lCacheIndex) = oObj.lCacheID Then
                    If goMineralCache(oObj.lCacheIndex) Is Nothing = False Then goMineralCache(oObj.lCacheIndex).BeingMinedBy = Nothing
                End If
            ElseIf oObj.lDockeeCnt > 0 Then
                'TODO: Cancel any outstanding dock requests
                oObj.lDockeeCnt = 0
            End If

            'clear any dock requests frm the queue
            CancelDockingQueueItem(oObj.ObjectID, oObj.ObjTypeID)

        End If
    End Sub
    Private Sub HandleMsgMonitorMsg(ByRef yData() As Byte, ByVal lSocketIndex As Int32, ByVal lSrcType As MsgMonitor.eMM_AppType)
        'Ok, another application is telling us its monitor msgs...
        'MsgCode (2), Type (4), Specific (4), Outs(...), Ins (...)

        Dim lPos As Int32 = 2
        Dim lConnType As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lSpecific As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'Now, for our outs and ins
        Dim blOuts(GlobalMessageCode.eLastMsgCode - 1) As Int64
        Dim blIns(GlobalMessageCode.eLastMsgCode - 1) As Int64

        For X As Int32 = 0 To blOuts.GetUpperBound(0)
            blOuts(X) = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
        Next X
        For X As Int32 = 0 To blIns.GetUpperBound(0)
            blIns(X) = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
        Next X

        MsgMonitor.AddExternalMsgMonData(lSrcType, lSocketIndex, CType(lConnType, MsgMonitor.eMM_AppType), lSpecific, blOuts, blIns)

    End Sub
    Private Sub HandleNewsItem(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        Dim yType As Byte = yData(lPos) : lPos += 1

        If yType = NewsItemType.eSpaceCombat OrElse yType = NewsItemType.eSpaceCombatUpdate OrElse yType = NewsItemType.eSpaceCombatEnd OrElse yType = NewsItemType.ePlanetCombat OrElse yType = NewsItemType.ePlanetCombatEnd OrElse yType = NewsItemType.ePlanetCombatUpdate Then
            Dim lEnvirPos As Int32 = lPos
            Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Dim lFightLocID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iFightLocTypeID As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            Dim lPlayerID As Int32 = -1

            Dim lGalaxyID As Int32 = -1

            If iFightLocTypeID = ObjectType.ePlanet Then
                Dim oPlanet As Planet = GetEpicaPlanet(lFightLocID)
                If oPlanet Is Nothing = False Then oPlanet.PlanetName.CopyTo(yData, lPos)

                'Now, determine our total kills
                lPos += 20 'for the name
                lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                lPos += 20  'for non-related data

                Dim lUnits As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Dim lFacs As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                lUnits += System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                lFacs += System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                If lUnits + lFacs > 10000 Then
                    lEnvirID = oPlanet.ParentSystem.ObjectID
                    iEnvirTypeID = oPlanet.ParentSystem.ObjTypeID
                End If

                lGalaxyID = oPlanet.ParentSystem.ParentGalaxy.ObjectID
                oPlanet = Nothing
            ElseIf iFightLocTypeID = ObjectType.eSolarSystem Then
                Dim oSystem As SolarSystem = GetEpicaSystem(lFightLocID)
                If oSystem Is Nothing = False Then oSystem.SystemName.CopyTo(yData, lPos)

                lPos += 20 'for the name
                lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                lPos += 36 'for other data

                lGalaxyID = oSystem.ParentGalaxy.ObjectID
                oSystem = Nothing
            End If

            'Now, does our player lists include an emperor or king?
            Dim oThisPlayer As Player = GetEpicaPlayer(lPlayerID)
            If oThisPlayer Is Nothing = False Then
                Dim yTestTitle As Byte = oThisPlayer.yPlayerTitle
                If (yTestTitle And Player.PlayerRank.ExRankShift) <> 0 Then yTestTitle = yTestTitle Xor Player.PlayerRank.ExRankShift
                If yTestTitle >= Player.PlayerRank.King Then
                    lEnvirID = lGalaxyID
                    iEnvirTypeID = ObjectType.eGalaxy
                End If
            End If

            'If iEnvirTypeID <> ObjectType.eGalaxy Then
            Dim lFirstUB As Int32 = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
            Dim oPlayers(lFirstUB) As Player
            For X As Int32 = 0 To lFirstUB
                lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                oPlayers(X) = GetEpicaPlayer(lPlayerID)
                If oPlayers(X) Is Nothing = False Then
                    Dim yTestTitle As Byte = oPlayers(X).yPlayerTitle
                    If (yTestTitle And Player.PlayerRank.ExRankShift) <> 0 Then yTestTitle = yTestTitle Xor Player.PlayerRank.ExRankShift
                    If yTestTitle >= Player.PlayerRank.King Then
                        lEnvirID = lGalaxyID
                        iEnvirTypeID = ObjectType.eGalaxy
                        'Exit For
                    End If
                End If
            Next X
            'If iEnvirTypeID <> ObjectType.eGalaxy Then
            Dim lSecondUB As Int32 = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4

            Dim lStartIdx As Int32 = oPlayers.GetUpperBound(0) + 1
            ReDim Preserve oPlayers(oPlayers.Length + lSecondUB)


            For X As Int32 = 0 To lSecondUB
                lPlayerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                oPlayers(X + lStartIdx) = GetEpicaPlayer(lPlayerID)
                If oPlayers(X + lStartIdx) Is Nothing = False Then
                    Dim yTestTitle As Byte = oPlayers(X + lStartIdx).yPlayerTitle
                    If (yTestTitle And Player.PlayerRank.ExRankShift) <> 0 Then yTestTitle = yTestTitle Xor Player.PlayerRank.ExRankShift
                    If yTestTitle >= Player.PlayerRank.King Then
                        lEnvirID = lGalaxyID
                        iEnvirTypeID = ObjectType.eGalaxy
                        'Exit For
                    End If
                End If
            Next X
            'End If
            'End If

            System.BitConverter.GetBytes(lEnvirID).CopyTo(yData, lEnvirPos) : lEnvirPos += 4
            System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yData, lEnvirPos) : lEnvirPos += 2

            'Begin preparing our new message... 75 bytes for the first portion of the message (doesn't change)
            '  an additional 41 bytes for the player specifics for this news item
            '  and... we add another 25 bytes for every entry in the oplayers array
            Dim yFinal(123 + ((lFirstUB + lSecondUB + 2) * 25)) As Byte
            Dim lFinalPos As Int32 = 0
            Array.Copy(yData, 0, yFinal, 0, 75) : lFinalPos += 75

            'PlayerName
            'PlayerGender
            'EmpireName
            If oThisPlayer Is Nothing = False Then
                oThisPlayer.PlayerName.CopyTo(yFinal, lFinalPos) : lFinalPos += 20
                yFinal(lFinalPos) = oThisPlayer.yGender : lFinalPos += 1
                oThisPlayer.EmpireName.CopyTo(yFinal, lFinalPos) : lFinalPos += 20
            Else : lFinalPos = 41
            End If

            System.BitConverter.GetBytes(lFirstUB + 1).CopyTo(yFinal, lFinalPos) : lFinalPos += 4
            For X As Int32 = 0 To lFirstUB
                If oPlayers(X) Is Nothing = False Then
                    With oPlayers(X)
                        System.BitConverter.GetBytes(.ObjectID).CopyTo(yFinal, lFinalPos) : lFinalPos += 4
                        .PlayerName.CopyTo(yFinal, lFinalPos) : lFinalPos += 20
                        yFinal(lFinalPos) = .yGender : lFinalPos += 1
                    End With
                Else : lFinalPos += 25
                End If
            Next X
            System.BitConverter.GetBytes(lSecondUB + 1).CopyTo(yFinal, lFinalPos) : lFinalPos += 4
            For X As Int32 = 0 To lSecondUB
                If oPlayers(lStartIdx + X) Is Nothing = False Then
                    With oPlayers(lStartIdx + X)
                        System.BitConverter.GetBytes(.ObjectID).CopyTo(yFinal, lFinalPos) : lFinalPos += 4
                        .PlayerName.CopyTo(yFinal, lFinalPos) : lFinalPos += 20
                        yFinal(lFinalPos) = .yGender : lFinalPos += 1
                    End With
                Else : lFinalPos += 25
                End If
            Next X

            SendToEmailSrvr(yFinal) 'moEmailSrvr.SendData(yFinal)
        End If
    End Sub
    Private Sub HandlePlacePirateAssets(ByRef yData() As Byte, ByVal Index As Int32)
        If goAureliusAI Is Nothing = False Then goAureliusAI.HandlePlacePirateAsset(yData)
        ''MscCode, EnvirGUID, StartLoc, PlayerID, ItemCnt, Items (GUID, Loc)
        'Dim lPos As Int32 = 2
        'Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        'Dim lSX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'Dim lSZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'Dim bDeleteEntities As Boolean = False

        'If iEnvirTypeID <> ObjectType.ePlanet Then bDeleteEntities = True

        'Dim oPlanet As Planet = GetEpicaPlanet(lEnvirID)
        'If oPlanet Is Nothing Then bDeleteEntities = True
        'Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        'If oPlayer Is Nothing Then bDeleteEntities = True



        'Dim lMoveToX As Int32 = Int32.MinValue
        'Dim lMoveToY As Int32 = Int32.MinValue
        'Dim lColonyIdx As Int32 = -1
        'If oPlayer Is Nothing = False Then lColonyIdx = oPlayer.GetColonyFromParent(lEnvirID, iEnvirTypeID)

        'If lColonyIdx = -1 OrElse glColonyIdx(lColonyIdx) = -1 Then
        '	bDeleteEntities = True
        'Else
        '	Dim oColony As Colony = goColony(lColonyIdx)
        '	If oColony Is Nothing = False Then
        '		For X As Int32 = 0 To oColony.ChildrenUB
        '			If oColony.lChildrenIdx(X) <> -1 Then
        '				If oColony.oChildren(X).yProductionType = ProductionType.eCommandCenterSpecial OrElse lMoveToX = Int32.MinValue Then
        '					'TODO: Make this smarter
        '					lMoveToX = oColony.oChildren(X).LocX + 1000
        '					lMoveToY = oColony.oChildren(X).LocZ + 1000

        '					If oColony.oChildren(X).yProductionType = ProductionType.eCommandCenterSpecial Then Exit For
        '				End If
        '			End If
        '		Next X
        '	Else : bDeleteEntities = True
        '	End If
        'End If


        'If bDeleteEntities = True Then
        '	For X As Int32 = 0 To lCnt - 1
        '		Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        '		Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        '		Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        '		Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        '		Dim oEntity As Epica_Entity = Nothing
        '		If iTypeID = ObjectType.eFacility Then
        '			oEntity = GetEpicaFacility(lObjID)
        '		Else : oEntity = GetEpicaUnit(lObjID)
        '		End If
        '		If oEntity Is Nothing = False Then oEntity.DeleteEntity(-1)
        '	Next X
        'End If

        'For X As Int32 = 0 To lCnt - 1
        '	Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        '	Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        '	Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        '	Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        '	Dim oEntity As Epica_Entity = Nothing
        '	If iTypeID = ObjectType.eFacility Then
        '		oEntity = GetEpicaFacility(lObjID)
        '		CType(oEntity, Facility).lPirate_For_PlayerID = lPlayerID
        '	Else : oEntity = GetEpicaUnit(lObjID)
        '	End If
        '	If oEntity Is Nothing = False Then
        '		If lLocX = Int32.MinValue OrElse lLocZ = Int32.MinValue Then
        '			oEntity.DeleteEntity(-1)
        '		Else
        '			oEntity.LocX = lLocX
        '			oEntity.LocZ = lLocZ
        '			oEntity.ParentObject = oPlanet
        '			oEntity.DataChanged()

        '			moDomains(Index).SendData(GetAddObjectMessage(oEntity, GlobalMessageCode.eAddObjectCommand))

        '			If oEntity.ObjTypeID = ObjectType.eUnit Then
        '				Dim yMove(23) As Byte
        '				System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequest).CopyTo(yMove, 0)
        '				System.BitConverter.GetBytes(lMoveToX).CopyTo(yMove, 2)
        '				System.BitConverter.GetBytes(lMoveToY).CopyTo(yMove, 6)
        '				System.BitConverter.GetBytes(0S).CopyTo(yMove, 10)
        '				System.BitConverter.GetBytes(lEnvirID).CopyTo(yMove, 12)
        '				System.BitConverter.GetBytes(iEnvirTypeID).CopyTo(yMove, 16)
        '				oEntity.GetGUIDAsString.CopyTo(yMove, 18)
        '				moDomains(Index).SendData(yMove)
        '			End If
        '		End If
        '	End If
        'Next X

    End Sub
    Private Sub HandlePlayerAlert(ByRef yData() As Byte)
        'MsgCode, Type(1), EntityGUID(6), EnvirGUID(6), PlayerID(4), EnemyID(4)

        'We only really care about the player ID
        Dim lPos As Int32 = 2       'for msg code
        Dim yType As Byte = yData(lPos) : lPos += 1
        lPos += 6   'for entityguid
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        For X As Int32 = 0 To glPlayerUB
            If glPlayerIdx(X) = lPlayerID AndAlso goPlayer(X).InMyDomain = True Then

                If (yType = PlayerAlertType.eUnderAttack OrElse yType = PlayerAlertType.eEngagedEnemy) AndAlso (iEnvirTypeID = ObjectType.ePlanet OrElse iEnvirTypeID = ObjectType.eSolarSystem) Then
                    goPlayer(X).oBudget.SetEnvironmentInConflict(lEnvirID, iEnvirTypeID, True)
                End If

                Dim yFinal(yData.GetUpperBound(0) + 20) As Byte
                yData.CopyTo(yFinal, 0) : lPos = yData.Length - 1

                If iEnvirTypeID = ObjectType.ePlanet Then
                    Dim oPlanet As Planet = GetEpicaPlanet(lEnvirID)
                    If oPlanet Is Nothing = False Then oPlanet.PlanetName.CopyTo(yFinal, lPos)
                ElseIf iEnvirTypeID = ObjectType.ePlanet Then
                    Dim oSystem As SolarSystem = GetEpicaSystem(lEnvirID)
                    If oSystem Is Nothing = False Then oSystem.SystemName.CopyTo(yFinal, lPos)
                End If
                goPlayer(X).SendPlayerMessage(yFinal, True, 0)
                Exit For
            End If
        Next X

    End Sub
    Private Sub HandlePlayerDiscoversWormhole(ByVal yData() As Byte)
        'MsgCode, PlayerID, WormholeID
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lWormholeID As Int32 = System.BitConverter.ToInt32(yData, 6)
        Dim lSystemID As Int32 = System.BitConverter.ToInt32(yData, 10)

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        If oPlayer Is Nothing Then Return
        Dim oWormhole As Wormhole = GetEpicaWormhole(lWormholeID)
        If oWormhole Is Nothing Then Return
        'oWormhole.StartCycle = 1
        'With oWormhole
        '    If lSystemID = .System1.ObjectID Then
        '        .WormholeFlags = .WormholeFlags Or elWormholeFlag.eSystem2Detectable
        '    Else
        '        .WormholeFlags = .WormholeFlags Or elWormholeFlag.eSystem1Detectable
        '    End If
        'End With

        oWormhole.QueueMeToSave()
        Dim oSystem As SolarSystem = GetEpicaSystem(lSystemID)
        If oSystem Is Nothing Then Return
        oPlayer.AddWormholeKnowledge(oWormhole, True, oSystem, True)
    End Sub
    Private Sub HandleRegionSetEntityStatus(ByRef yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lStatus As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oEntity As Epica_Entity = GetEpicaEntity(lObjID, iObjTypeID)
        If oEntity Is Nothing = False Then
            If lStatus < 0 Then
                lStatus = Math.Abs(lStatus)
                If (oEntity.CurrentStatus And lStatus) <> 0 Then oEntity.CurrentStatus = oEntity.CurrentStatus Xor lStatus
            Else
                'If lStatus = elUnitStatus.eGuildAsset Then
                '    Dim lCost As Int32 = 0
                '    If iObjTypeID = ObjectType.eUnit Then
                '        lCost = CType(oEntity, Unit).EntityDef.WarpointUpkeep
                '    ElseIf iObjTypeID = ObjectType.eFacility Then
                '        lCost = CType(oEntity, Facility).EntityDef.WarpointUpkeep
                '    End If

                '    If oEntity.Owner.blWarpoints > lCost Then
                '        oEntity.Owner.blWarpoints -= lCost
                '    Else
                '        lStatus = -lStatus
                '        Dim yResp(11) As Byte
                '        System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityStatus).CopyTo(yResp, 0)
                '        oEntity.GetGUIDAsString.CopyTo(yResp, 2)
                '        System.BitConverter.GetBytes(lStatus).CopyTo(yResp, 8)
                '        moDomains(lIndex).SendData(yResp)
                '        Return
                '    End If

                'End If
                oEntity.CurrentStatus = oEntity.CurrentStatus Or lStatus
            End If
        End If
    End Sub
    Private Sub HandleRegionShiftClickAddProduction(ByVal yData() As Byte, ByVal lIndex As Int32)
        'ok, a region server is stating that a player is ordering a unit to queue a production item...
        '  it has validated the player and the terrain.... so let's get the data to play with
        Dim lPos As Int32 = 2

        Dim lBuilderID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iBuilderTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lProdID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iProdTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iLocA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        If iBuilderTypeID <> ObjectType.eUnit Then Return

        Dim oED As Epica_Entity_Def = Nothing
        If iProdTypeID = ObjectType.eUnitDef OrElse iProdTypeID = ObjectType.eFacilityDef Then
            oED = CType(GetEpicaObject(lProdID, iProdTypeID), Epica_Entity_Def)
        End If
        If oED Is Nothing = True Then Return

        Dim oUnit As Unit = GetEpicaUnit(lBuilderID)
        If oUnit Is Nothing Then Return

        Dim oProd As New EntityProduction()
        With oProd
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
        oUnit.AddToProdQueue(oProd)

        'Now, send that to the player
        oUnit.Owner.SendPlayerMessage(yData, False, AliasingRights.eViewUnitsAndFacilities)
    End Sub
    Private Sub HandleRegisterDomain(ByVal yData() As Byte, ByVal Index As Integer)
        Dim lObjID As Int32
        Dim iTypeID As Int16

        'Ok, the domain server is saying it has domain over the object referenced

        lObjID = System.BitConverter.ToInt32(yData, 2)
        iTypeID = System.BitConverter.ToInt16(yData, 6)
        If iTypeID = ObjectType.ePlanet Then
            GetEpicaPlanet(lObjID).oDomain = moServers(Index)
        ElseIf iTypeID = ObjectType.eSolarSystem Then
            GetEpicaSystem(lObjID).oDomain = moServers(Index)
        End If

    End Sub
    Private Sub HandleReloadWpnMsg(ByVal yData() As Byte, ByVal lIndex As Int32)
        '2 bytes msg code
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lWpnID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iWpnTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)

        Dim oEntity As Epica_Entity
        Dim oEntityDef As Epica_Entity_Def

        Dim bAdded As Boolean = False
        Dim lVal As Int32 = 0
        Dim oAmmo As AmmunitionCache

        Dim lAmmoCap As Int32 = -1

        Dim X As Int32

        oEntity = CType(GetEpicaObject(lObjID, iObjTypeID), Epica_Entity)
        If oEntity Is Nothing = False AndAlso (oEntity.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
            'Ok, get the entitydef
            If oEntity.ObjTypeID = ObjectType.eUnit Then
                oEntityDef = CType(oEntity, Unit).EntityDef
            Else : oEntityDef = CType(oEntity, Facility).EntityDef
            End If

            'Now, let's get the actual WeaponDef lWpnID is the UnitDefWeaponID (or FacilityDefWeaponID)
            Dim oWpnDef As WeaponDef = Nothing
            For X = 0 To oEntityDef.WeaponDefUB
                If oEntityDef.WeaponDefs(X).ObjectID = lWpnID Then
                    oWpnDef = oEntityDef.WeaponDefs(X).oWeaponDef
                    lAmmoCap = oEntityDef.WeaponDefs(X).mlAmmoCap
                    If lAmmoCap < 1 Then Return
                    Exit For
                End If
            Next X
            If oWpnDef Is Nothing OrElse oWpnDef.RelatedWeapon Is Nothing Then Return

            Dim lTrueWpnID As Int32 = oWpnDef.RelatedWeapon.ObjectID

            For lCIdx As Int32 = 0 To oEntity.lCargoUB
                If oEntity.lCargoIdx(lCIdx) <> -1 AndAlso oEntity.oCargoContents(lCIdx).ObjTypeID = ObjectType.eAmmunition Then
                    oAmmo = CType(oEntity.oCargoContents(lCIdx), AmmunitionCache)

                    If oAmmo.oWeaponTech Is Nothing = False AndAlso oAmmo.oWeaponTech.ObjectID = lTrueWpnID Then
                        'Ok, this ammo cache matches the weapon we are needing to reload...
                        lVal = oAmmo.Quantity

                        If lVal > 0 Then
                            'If we are getting this message than the region server is reporting that the entity
                            '  is out of ammo for this weapon, so iAmmoCap is the amount of ammo we need to load

                            If lVal < lAmmoCap Then
                                'Ok, cargo doesn't have enough to fully reload the weapon, so take all of it (what cargo does have)
                                lAmmoCap = lVal
                            End If

                            'lAmmoCap is what will be reloaded into the weapon in the end...
                            oAmmo.Quantity -= lAmmoCap

                            'Now, add to our queue
                            AddToQueue(glCurrentCycle + oWpnDef.AmmoReloadDelay, EngineCode.QueueItemType.eReloadRequest, lObjID, iObjTypeID, lWpnID, lAmmoCap, 0, 0, 0, 0)
                            bAdded = True
                        End If

                        'WE exit for because the cargo bay should not have any other ammo caches for this weapon
                        Exit For
                    End If

                    oAmmo = Nothing
                End If
            Next lCIdx

        End If

        'Now, check if we added a queue item
        If bAdded = False Then
            'We did not, so... that means no more ammo is available
            'TODO: Notify the player that the unit is out of ammo
        End If
    End Sub
    Private Sub HandleRemoveObject(ByVal yData() As Byte)
        Dim X As Int32

        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int32 = System.BitConverter.ToInt16(yData, 6)
        Dim yRemoveType As RemovalType = CType(yData(8), RemovalType)

        Dim oEntity As Epica_Entity = Nothing
        Dim oTmpDef As Epica_Entity_Def
        Dim lIdx As Int32 = -1
        Dim lCacheIdx As Int32 = -1

        If iObjTypeID = ObjectType.eFacility Then
            For X = 0 To glFacilityUB
                If glFacilityIdx(X) = lObjID Then
                    oEntity = goFacility(X)
                    lIdx = X
                    Exit For
                End If
            Next X
        ElseIf iObjTypeID = ObjectType.eUnit Then
            For X = 0 To glUnitUB
                If glUnitIdx(X) = lObjID Then
                    oEntity = goUnit(X)
                    lIdx = X
                    Exit For
                End If
            Next X
        End If
        If oEntity Is Nothing Then Return

        Select Case yRemoveType
            Case RemovalType.eObjectDestroyed
                'Now, create the mineral deposits... oEntity should have an EntityDef with it... that entitydef should
                '  have the mineral cache details for when a unit of that type is destroyed

                'store its final resting place
                oEntity.LocX = System.BitConverter.ToInt32(yData, 9)
                oEntity.LocZ = System.BitConverter.ToInt32(yData, 13)
                oEntity.lDeathStatus = System.BitConverter.ToInt32(yData, 17)

                If oEntity.Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID Then
                    Dim oTmpParent As Epica_GUID = CType(oEntity.ParentObject, Epica_GUID)
                    If oTmpParent Is Nothing = False Then
                        If oTmpParent.ObjectID < 500000000 Then
                            If goAureliusAI Is Nothing = False Then goAureliusAI.RemoveAIEntity(oTmpParent.ObjectID, oTmpParent.ObjTypeID, oEntity.ObjectID, oEntity.ObjTypeID)
                        End If
                    End If
                End If

                'ok, is the entity a facility?
                If oEntity.ObjTypeID = ObjectType.eFacility Then

                    'If oEntity.Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID Then
                    '	With CType(oEntity, Facility)
                    '		If .EntityDef Is Nothing = False AndAlso .EntityDef.ObjectID = 1558 Then
                    '			If .ParentObject Is Nothing = False Then
                    '				Dim lTemp As Int32 = CType(.ParentObject, Epica_GUID).ObjectID
                    '				lTemp -= 500000000
                    '				AddToQueue(glCurrentCycle + 30, QueueItemType.eGenerateNewbieAgent, lTemp, ObjectType.ePlayer, -1, -1, 0, 0, 0, 0)
                    '			End If
                    '		End If
                    '	End With
                    'End If

                    If oEntity.yProductionType = ProductionType.eCommandCenterSpecial AndAlso oEntity.Owner.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                        CType(oEntity, Facility).ParentColony.Population = 0
                    End If

                    'is someone online that can rebuild?
                    'If oEntity.Owner.HasOnlineAliases(AliasingRights.eAddProduction) = False Then
                    'no, ok, is the facility part of a colony?
                    With CType(oEntity, Facility)
                        If .ParentColony Is Nothing = False Then
                            .ParentColony.AddRebuildItem(.EntityDef.ObjectID, .EntityDef.ObjTypeID, CType(.ParentObject, Epica_GUID).ObjectID, CType(.ParentObject, Epica_GUID).ObjTypeID, .LocX, .LocZ, .LocAngle)
                            If .yProductionType = ProductionType.eColonists Then
                                'Ok, we take the colony's powered and unpowered population over the powered and unpowered residence
                                Dim lSumHousing As Int32 = .ParentColony.PoweredHousing + .ParentColony.UnpoweredHousing
                                Dim fMult As Single
                                If lSumHousing = 0 Then
                                    fMult = 1
                                Else
                                    fMult = Math.Min(1.0F, CSng(.ParentColony.Population / lSumHousing))
                                End If
                                Dim lPopLoss As Int32 = .EntityDef.ProdFactor \ 2
                                lPopLoss = CInt(lPopLoss * fMult)
                                .ParentColony.Population -= lPopLoss
                            End If
                        End If
                    End With
                    'End If
                Else
                    With CType(oEntity, Unit)
                        If .oRebuilderFor Is Nothing = False Then
                            'ok, this guy was a rebuilder...
                            .oRebuilderFor.lRebuilderQueueIdx = AddToQueueResult(glCurrentCycle, QueueItemType.eBeginRebuilderAI, .oRebuilderFor.ObjectID, 0, 0, 0, 0, 0, 0, 0)
                            .oRebuilderFor.lRebuilderUnitID = -1
                        End If
                    End With
                End If
                Dim lKilledByID As Int32 = -1
                Try
                    lKilledByID = System.BitConverter.ToInt32(yData, 21)
                Catch
                End Try

                DestroyEntity(oEntity, False, lKilledByID, True, "HandleRemoveObject")
                'If oEntity.ObjTypeID = ObjectType.eUnit Then
                '    oTmpDef = CType(oEntity, Unit).EntityDef
                'Else : oTmpDef = CType(oEntity, Facility).EntityDef
                'End If

                'For X = 0 To oTmpDef.lEntityDefMineralUB
                '    lCacheIdx = AddMineralCache(CType(oEntity.ParentObject, Epica_GUID).ObjectID, CType(oEntity.ParentObject, Epica_GUID).ObjTypeID, _
                '       MineralCacheType.ePickupable, oTmpDef.EntityDefMinerals(X).lQuantity, _
                '       oTmpDef.EntityDefMinerals(X).lQuantity, oEntity.LocX, oEntity.LocZ, oTmpDef.EntityDefMinerals(X).oMineral)

                '    Dim iTemp As Int16 = CType(oEntity.ParentObject, Epica_GUID).ObjTypeID
                '    If iTemp = ObjectType.ePlanet Then
                '        CType(oEntity.ParentObject, Planet).oDomain.DomainSocket.SendData(GetAddObjectMessage(goMineralCache(lCacheIdx), GlobalMessageCode.eAddObjectCommand))
                '    ElseIf iTemp = ObjectType.eSolarSystem Then
                '        CType(oEntity.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(GetAddObjectMessage(goMineralCache(lCacheIdx), GlobalMessageCode.eAddObjectCommand))
                '    End If
                'Next X

                'If oEntity.ObjTypeID = ObjectType.eFacility Then
                '    CType(oEntity, Facility).RemoveMe()
                'End If

                'If iObjTypeID = ObjectType.eUnit Then
                '    With goUnit(lIdx)
                '        .Owner.CreateAndSendPlayerAlert(PlayerAlertType.eUnitLost, .ObjectID, .ObjTypeID, CType(.ParentObject, Epica_GUID).ObjectID, CType(.ParentObject, Epica_GUID).ObjTypeID, -1, BytesToString(.EntityName), .LocX, .LocZ)
                '    End With

                '    glUnitIdx(lIdx) = -1
                '    goUnit(lIdx).DeleteEntity(lIdx)
                '    goUnit(lIdx) = Nothing
                'ElseIf iObjTypeID = ObjectType.eFacility Then
                '    With goFacility(lIdx)
                '        .Owner.CreateAndSendPlayerAlert(PlayerAlertType.eFacilityLost, .ObjectID, .ObjTypeID, CType(.ParentObject, Epica_GUID).ObjectID, CType(.ParentObject, Epica_GUID).ObjTypeID, -1, BytesToString(.EntityName), .LocX, .LocZ)
                '        If .Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID AndAlso (.yProductionType And ProductionType.eProduction) <> 0 Then
                '            AddToQueue(glCurrentCycle + 9000, QueueItemType.ePirateSelfDestruct, .lPirate_For_PlayerID, ObjectType.eFacility, gl_HARDCODE_PIRATE_PLAYER_ID, CType(.ParentObject, Epica_GUID).ObjectID)
                '            AddToQueue(glCurrentCycle + 18000, QueueItemType.ePirateSelfDestruct, .lPirate_For_PlayerID, ObjectType.eUnit, gl_HARDCODE_PIRATE_PLAYER_ID, CType(.ParentObject, Epica_GUID).ObjectID)
                '        End If
                '    End With

                '    glFacilityIdx(lIdx) = -1
                '    goFacility(lIdx).DeleteEntity(lIdx)
                '    goFacility(lIdx) = Nothing
                'End If

            Case RemovalType.eChangingEnvironments, RemovalType.eDocking, RemovalType.eJumping
                Dim lJumpID As Int32 = -1
                Dim iJumpTypeID As Int16 = -1
                Dim lCurEnvirID As Int32 = -1
                Dim iCurEnvirTypeID As Int16 = -1

                With oEntity
                    'when change environments is called, the PF will indicate our new location
                    If yRemoveType <> RemovalType.eChangingEnvironments Then
                        .LocX = System.BitConverter.ToInt32(yData, 9)
                        .LocZ = System.BitConverter.ToInt32(yData, 13)
                    End If

                    'Skip LocAngle
                    .Fuel_Cap = System.BitConverter.ToInt16(yData, 19)
                    .ExpLevel = yData(23)
                    .iTargetingTactics = System.BitConverter.ToInt16(yData, 24)
                    .iCombatTactics = System.BitConverter.ToInt32(yData, 26)

                    If (.iCombatTactics And eiBehaviorPatterns.eEngagement_Evade) <> 0 Then
                        .iCombatTactics -= eiBehaviorPatterns.eEngagement_Evade
                        .iCombatTactics = .iCombatTactics Or eiBehaviorPatterns.eEngagement_Stand_Ground
                    End If

                    'If yRemoveType <> RemovalType.eDocking Then
                    '    .ParentObject = GetObject(System.BitConverter.ToInt32(yData, 26), System.BitConverter.ToInt16(yData, 32))
                    'End If
                    'If .ParentObject Is Nothing = False Then
                    '	Dim lPID As Int32 = System.BitConverter.ToInt32(yData, 30)
                    '	Dim iPTypeID As Int16 = System.BitConverter.ToInt16(yData, 34)
                    '	With CType(.ParentObject, Epica_GUID)
                    '		If .ObjectID <> lPID OrElse .ObjTypeID <> iPTypeID Then
                    '			LogEvent(LogEventType.Warning, "HandleRemoveObject_CE Parent Mismatch. Region Parent: " & lPID & ", " & iPTypeID & ". Primary: " & .ObjectID & ", " & .ObjTypeID)
                    '		End If
                    '	End With
                    'End If

                    'Add object occurs in the EntityChangeEnvironment handler

                    '.Owner = GetEpicaPlayer(System.BitConverter.ToInt32(yData, 36))

                    .CurrentStatus = System.BitConverter.ToInt32(yData, 40)
                    .Shield_HP = System.BitConverter.ToInt32(yData, 44)
                    .Q1_HP = System.BitConverter.ToInt32(yData, 48)
                    .Q2_HP = System.BitConverter.ToInt32(yData, 52)
                    .Q3_HP = System.BitConverter.ToInt32(yData, 56)
                    .Q4_HP = System.BitConverter.ToInt32(yData, 60)
                    .Structure_HP = System.BitConverter.ToInt32(yData, 64)

                    Dim lCnt As Int32 = System.BitConverter.ToInt16(yData, 68)
                    Dim lPos As Int32 = 70
                    Dim lUDWID As Int32
                    Dim Y As Int32

                    If oEntity.ObjTypeID = ObjectType.eUnit Then
                        oTmpDef = CType(oEntity, Unit).EntityDef
                    Else : oTmpDef = CType(oEntity, Facility).EntityDef
                    End If

                    'TODO: This assumes that the message will be long enough for this list
                    For X = 0 To lCnt - 1
                        lUDWID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        For Y = 0 To oTmpDef.WeaponDefUB
                            If oTmpDef.WeaponDefs(Y).ObjectID = lUDWID Then
                                If .lCurrentAmmo Is Nothing = False AndAlso .lCurrentAmmo.Length > Y Then
                                    .lCurrentAmmo(Y) = System.BitConverter.ToInt32(yData, lPos)
                                End If
                                Exit For
                            End If
                        Next Y

                        'regardless of found or not, increment position by 4
                        lPos += 4
                    Next X

                    'ok, if this is jumping
                    If yRemoveType = RemovalType.eJumping Then
                        lJumpID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        iJumpTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                        lCurEnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                        iCurEnvirTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    End If
                End With

                If oEntity.ObjTypeID = ObjectType.eUnit Then
                    CType(oEntity, Unit).DataChanged()
                ElseIf oEntity.ObjTypeID = ObjectType.eFacility Then
                    CType(oEntity, Facility).DataChanged()
                Else : oEntity.DataChanged()
                End If

                If lJumpID <> -1 AndAlso iJumpTypeID <> -1 AndAlso lCurEnvirID <> -1 AndAlso iCurEnvirTypeID <> -1 Then
                    If iJumpTypeID = ObjectType.eWormhole Then
                        Dim oWormhole As Wormhole = GetEpicaWormhole(lJumpID)
                        Dim oSocket As NetSock = Nothing
                        If oWormhole Is Nothing = False Then
                            If iCurEnvirTypeID = ObjectType.eSolarSystem Then
                                If oWormhole.System1 Is Nothing = False AndAlso oWormhole.System1.ObjectID = lCurEnvirID Then
                                    If oWormhole.System2 Is Nothing = False Then
                                        oEntity.ParentObject = oWormhole.System2
                                        oEntity.LocX = oWormhole.LocX2
                                        oEntity.LocZ = oWormhole.LocY2
                                        oSocket = oWormhole.System2.oDomain.DomainSocket
                                    Else
                                        oEntity.LocX = oWormhole.LocX1
                                        oEntity.LocZ = oWormhole.LocY1
                                        oEntity.ParentObject = oWormhole.System1
                                        oSocket = oWormhole.System1.oDomain.DomainSocket
                                    End If
                                ElseIf oWormhole.System2 Is Nothing = False AndAlso oWormhole.System2.ObjectID = lCurEnvirID Then
                                    If oWormhole.System1 Is Nothing = False Then
                                        oEntity.ParentObject = oWormhole.System1
                                        oEntity.LocX = oWormhole.LocX1
                                        oEntity.LocZ = oWormhole.LocY1
                                        oSocket = oWormhole.System1.oDomain.DomainSocket
                                    Else
                                        oEntity.ParentObject = oWormhole.System2
                                        oEntity.LocX = oWormhole.LocX2
                                        oEntity.LocZ = oWormhole.LocY2
                                        oSocket = oWormhole.System2.oDomain.DomainSocket
                                    End If
                                End If
                            End If
                        End If

                        oEntity.LocX += CInt(Rnd() * 2000) - 1000
                        oEntity.LocZ += CInt(Rnd() * 2000) - 1000

                        If oSocket Is Nothing = False Then
                            If oEntity.ObjTypeID = ObjectType.eUnit Then
                                CType(oEntity, Unit).DataChanged()
                            ElseIf oEntity.ObjTypeID = ObjectType.eFacility Then
                                CType(oEntity, Facility).DataChanged()
                            Else : oEntity.DataChanged()
                            End If

                            oSocket.SendData(GetAddObjectMessage(oEntity, GlobalMessageCode.eAddObjectCommand))
                        End If

                        oEntity.CheckUpdateUnitGroup()
                    Else
                        'TODO: What else?
                    End If
                End If

        End Select

    End Sub
    Private Sub HandleDomainRequestEntityDefenses(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim oEntity As Epica_Entity = GetEpicaEntity(lObjID, iObjTypeID)
        If oEntity Is Nothing = False Then
            With oEntity
                .Q1_HP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .Q2_HP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .Q3_HP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .Q4_HP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .Structure_HP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .Shield_HP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                'We want to trust the primary's version of the Facility Powered and Move Lock flags
                Dim bFacPowered As Boolean = (.CurrentStatus And elUnitStatus.eFacilityPowered) <> 0
                Dim bMoveLock As Boolean = (.CurrentStatus And elUnitStatus.eMoveLock) <> 0
                .CurrentStatus = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                If (.CurrentStatus And elUnitStatus.eFacilityPowered) <> 0 Then
                    'Ok, its powered, but is it supposed to be?
                    If bFacPowered = False Then .CurrentStatus = .CurrentStatus Xor elUnitStatus.eFacilityPowered ' .CurrentStatus -= elUnitStatus.eFacilityPowered
                Else
                    'ok, its not powered, but is it supposed to be?
                    If bFacPowered = True Then .CurrentStatus = .CurrentStatus Or elUnitStatus.eFacilityPowered
                End If
                If (.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then
                    'ok, its move locked, but is it supposed to be
                    If bMoveLock = False Then .CurrentStatus = .CurrentStatus Xor elUnitStatus.eMoveLock ' .CurrentStatus -= elUnitStatus.eMoveLock
                Else
                    'ok, its not move locked, but is it supposed to be?
                    If bMoveLock = True Then .CurrentStatus = .CurrentStatus Or elUnitStatus.eMoveLock
                End If
            End With

        End If


    End Sub
    Public Sub HandleRequestPlayerStartLoc(ByVal yData() As Byte)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, 6)
        Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 10)
        Dim lStartX As Int32 = System.BitConverter.ToInt32(yData, 12)
        Dim lStartZ As Int32 = System.BitConverter.ToInt32(yData, 16)

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        Dim oPlanet As Planet

        If oPlayer Is Nothing Then Return

        If iEnvirTypeID = ObjectType.ePlanet Then
            oPlanet = GetEpicaPlanet(lEnvirID)
        Else : Return         'TODO: We always assume the player starts on a planet
        End If

        If oPlanet.ParentSystem Is Nothing = False AndAlso oPlanet.ParentSystem.SystemType = 255 Then
            If lStartX = Int32.MinValue OrElse lStartZ = Int32.MinValue Then
                Select Case lEnvirID
                    Case 509
                        lStartX = -1690
                        lStartZ = 0
                    Case 513
                        lStartX = 5355
                        lStartZ = -9866
                    Case 520
                        lStartX = -1860
                        lStartZ = -4350
                    Case 528
                        lStartX = -1860
                        lStartZ = 0
                    Case 538
                        lStartX = 16800
                        lStartZ = 10400
                    Case 539
                        lStartX = 54700
                        lStartZ = -51272
                    Case Else
                        lStartX = 0
                        lStartZ = 0
                End Select
            End If
        End If

        If lStartX = Int32.MinValue OrElse lStartZ = Int32.MinValue Then
            'Ok, not good... 
            LogEvent(LogEventType.Informational, "Region server reported no room for planet ID: " & lEnvirID)
            If oPlanet Is Nothing = False Then
                oPlanet.PlayerSpawns += 1
                oPlanet.SpawnLocked = True
            End If
            Dim lOrigPlanetID As Int32 = oPlayer.lStartedEnvirID
            If lOrigPlanetID = -1 Then lOrigPlanetID = lEnvirID
            oPlayer.lStartedEnvirID = lEnvirID '-1
            InitializePlayer(oPlayer, lOrigPlanetID)
            Return
        End If

        'If oPlayer.oSocket Is Nothing = False Then
        'Set up our player...
        With oPlayer
            'If .PlayerIsDead = False Then
            '    .blCredits = 500000


            'Else
            .DeathBudgetBalance = 0
            If .blCredits = 0 Then .blCredits = 3000000
            'End If
            .lLastViewedEnvir = lEnvirID
            .iLastViewedEnvirType = iEnvirTypeID

            If .lStartedEnvirID > -1 AndAlso .iStartedEnvirTypeID > -1 Then
                If .iStartedEnvirTypeID = ObjectType.ePlanet Then
                    Dim oLastStart As Planet = GetEpicaPlanet(.lStartedEnvirID)
                    If oLastStart Is Nothing = False Then
                        If oLastStart.ParentSystem Is Nothing = False Then
                            If oLastStart.ParentSystem.SystemType = 255 Then
                                Try
                                    Dim yGNS(48) As Byte
                                    Dim lPos As Int32 = 0
                                    System.BitConverter.GetBytes(GlobalMessageCode.eNewsItem).CopyTo(yGNS, lPos) : lPos += 2
                                    yGNS(lPos) = NewsItemType.ePlanetFall : lPos += 1
                                    oPlanet.GetGUIDAsString.CopyTo(yGNS, lPos) : lPos += 6
                                    StringToBytes(.sPlayerNameProper).CopyTo(yGNS, lPos) : lPos += 20
                                    oPlanet.PlanetName.CopyTo(yGNS, lPos) : lPos += 20
                                    SendToEmailSrvr(yGNS)
                                Catch ex As Exception
                                    LogEvent(LogEventType.CriticalError, "PlayerStartLoc.SendToGNS: " & ex.Message)
                                End Try
                            End If
                        End If
                    End If
                End If
            End If

            .lStartedEnvirID = lEnvirID
            .iStartedEnvirTypeID = iEnvirTypeID
            .lStartLocX = lStartX
            .lStartLocZ = lStartZ
            .lIronCurtainPlanet = .lStartedEnvirID
            .DataChanged()
        End With

        If oPlanet.ParentSystem Is Nothing = False Then
            If oPlanet.ParentSystem.SystemType <> 255 Then
                oPlanet.PlayerSpawns += 1
            Else        'aurelium
                oPlayer.lIronCurtainPlanet = oPlanet.ObjectID
                oPlayer.yIronCurtainState = eIronCurtainState.IronCurtainIsUpOnSelectedPlanet
                AddToQueue(glCurrentCycle + 27000, QueueItemType.eIronCurtainFall, oPlayer.ObjectID, -1, -1, -1, -1, -1, -1, -1)

                Dim yMsg(10) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eSetIronCurtain).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(oPlanet.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = 1 : lPos += 1
                oPlanet.oDomain.DomainSocket.SendData(yMsg)
            End If
        End If

        'Now, set up our starting unit
        Dim lIdx As Int32 = -1
        Dim oUnitDef As Epica_Entity_Def = GetEpicaUnitDef(27)
        Dim oTmp As Unit = New Unit()
        Dim X As Int32

        'Ok, populate our values
        With oTmp
            .bProducing = False
            '.Cargo_Cap = oUnitDef.Cargo_Cap
            .CurrentProduction = Nothing
            .CurrentSpeed = 0

            .CurrentStatus = 0
            For X = 0 To oUnitDef.lSideCrits.Length - 1
                .CurrentStatus = .CurrentStatus Or oUnitDef.lSideCrits(X)
            Next X

            .EntityDef = oUnitDef
            oUnitDef.DefName.CopyTo(.EntityName, 0)
            .ExpLevel = 0
            .Fuel_Cap = oUnitDef.Fuel_Cap
            '.Hangar_Cap = oUnitDef.Hangar_Cap
            .iCombatTactics = 514
            .iTargetingTactics = 0
            .LocAngle = 0
            .LocX = lStartX
            .LocZ = lStartZ
            .ObjTypeID = ObjectType.eUnit
            .Owner = oPlayer
            .ParentObject = GetEpicaObject(lEnvirID, iEnvirTypeID)
            .Q1_HP = oUnitDef.Q1_MaxHP
            .Q2_HP = oUnitDef.Q2_MaxHP
            .Q3_HP = oUnitDef.Q3_MaxHP
            .Q4_HP = oUnitDef.Q4_MaxHP
            .Shield_HP = oUnitDef.Shield_MaxHP
            .Structure_HP = oUnitDef.Structure_MaxHP
            .yProductionType = oUnitDef.ProductionTypeID
            .DataChanged()
        End With

        If oTmp.SaveObject() = True Then
            'Now, find a suitable place...
            'SyncLock goUnit
            '    For X = 0 To glUnitUB
            '        If glUnitIdx(X) = -1 Then
            '            lIdx = X
            '            Exit For
            '        End If
            '    Next X

            '    If lIdx = -1 Then
            '        glUnitUB += 1
            '        ReDim Preserve glUnitIdx(glUnitUB)
            '        ReDim Preserve goUnit(glUnitUB)
            '        lIdx = glUnitUB
            '    End If

            '    goUnit(lIdx) = oTmp
            '    glUnitIdx(lIdx) = oTmp.ObjectID
            'End SyncLock
            lIdx = AddUnitToGlobalArray(oTmp)
        Else
            LogEvent(LogEventType.CriticalError, "Unable to save the first unit during initializeplayer!")
        End If

        'Send our add object to the region server
        If iEnvirTypeID = ObjectType.ePlanet Then
            CType(oTmp.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oTmp, GlobalMessageCode.eAddObjectCommand))
        ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
            CType(oTmp.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oTmp, GlobalMessageCode.eAddObjectCommand))
        End If

        If oPlayer.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
            'then, let's create 15 tanks
            oUnitDef = GetEpicaUnitDef(26)

            For lUnitIdx As Int32 = 0 To 15
                oTmp = New Unit()

                'Ok, populate our values
                With oTmp
                    .bProducing = False
                    '.Cargo_Cap = oUnitDef.Cargo_Cap
                    .CurrentProduction = Nothing
                    .CurrentSpeed = 0

                    .CurrentStatus = 0
                    For X = 0 To oUnitDef.lSideCrits.Length - 1
                        .CurrentStatus = .CurrentStatus Or oUnitDef.lSideCrits(X)
                    Next X

                    .EntityDef = oUnitDef
                    oUnitDef.DefName.CopyTo(.EntityName, 0)
                    .ExpLevel = 0
                    .Fuel_Cap = oUnitDef.Fuel_Cap
                    .iCombatTactics = eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Normal
                    .iTargetingTactics = 0
                    .LocAngle = 0
                    .LocX = lStartX
                    .LocZ = lStartZ
                    .ObjTypeID = ObjectType.eUnit
                    .Owner = oPlayer
                    .ParentObject = GetEpicaObject(lEnvirID, iEnvirTypeID)
                    .Q1_HP = oUnitDef.Q1_MaxHP
                    .Q2_HP = oUnitDef.Q2_MaxHP
                    .Q3_HP = oUnitDef.Q3_MaxHP
                    .Q4_HP = oUnitDef.Q4_MaxHP
                    .Shield_HP = oUnitDef.Shield_MaxHP
                    .Structure_HP = oUnitDef.Structure_MaxHP
                    .yProductionType = oUnitDef.ProductionTypeID
                    .DataChanged()
                End With

                If oTmp.SaveObject() = True Then
                    'Now, find a suitable place...
                    'SyncLock goUnit
                    '    For X = 0 To glUnitUB
                    '        If glUnitIdx(X) = -1 Then
                    '            lIdx = X
                    '            Exit For
                    '        End If
                    '    Next X

                    '    If lIdx = -1 Then
                    '        glUnitUB += 1
                    '        ReDim Preserve glUnitIdx(glUnitUB)
                    '        ReDim Preserve goUnit(glUnitUB)
                    '        lIdx = glUnitUB
                    '    End If

                    '    goUnit(lIdx) = oTmp
                    '    glUnitIdx(lIdx) = oTmp.ObjectID
                    'End SyncLock
                    lIdx = AddUnitToGlobalArray(oTmp)
                Else
                    LogEvent(LogEventType.CriticalError, "Unable to save the first unit during initializeplayer!")
                End If

                'Send our add object to the region server
                If iEnvirTypeID = ObjectType.ePlanet Then
                    CType(oTmp.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oTmp, GlobalMessageCode.eAddObjectCommand_CE))
                ElseIf iEnvirTypeID = ObjectType.eSolarSystem Then
                    CType(oTmp.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oTmp, GlobalMessageCode.eAddObjectCommand_CE))
                End If
            Next lUnitIdx
        End If

        'Now... send back a login response...
        oPlayer.PlayerIsDead = False
        If oPlayer.PlayerIsDead = False Then
            Dim yResp(5) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerLoginResponse).CopyTo(yResp, 0)
            System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, 2)
            oPlayer.SendPlayerMessage(yResp, False, 0)
            'oPlayer.oSocket.SendData(yResp)
        Else
            oPlayer.PlayerIsDead = False
            oPlayer.DeathBudgetBalance = 0
            oPlayer.DeathBudgetFundsRemaining = CInt(oPlayer.blCredits)

            'Dim yResp(20) As Byte
            'Dim lPos As Int32 = 0
            'System.BitConverter.GetBytes(GlobalMessageCode.ePlayerIsDead).CopyTo(yResp, 0) : lPos += 2
            'yResp(lPos) = 2 : lPos += 1
            'System.BitConverter.GetBytes(oPlayer.lStartedEnvirID).CopyTo(yResp, lPos) : lPos += 4
            'System.BitConverter.GetBytes(oPlayer.iStartedEnvirTypeID).CopyTo(yResp, lPos) : lPos += 2

            'Dim lGalaxyID As Int32
            'Dim lSystemID As Int32
            'Dim lPlanetID As Int32

            'lPlanetID = oPlayer.lStartedEnvirID
            'lSystemID = CType(oTmp.ParentObject, Planet).ParentSystem.ObjectID
            'lGalaxyID = CType(oTmp.ParentObject, Planet).ParentSystem.ParentGalaxy.ObjectID

            'System.BitConverter.GetBytes(lGalaxyID).CopyTo(yResp, lPos) : lPos += 4
            'System.BitConverter.GetBytes(lSystemID).CopyTo(yResp, lPos) : lPos += 4
            'System.BitConverter.GetBytes(lPlanetID).CopyTo(yResp, lPos) : lPos += 4
            ''oPlayer.oSocket.SendData(yResp)
            'oPlayer.SendPlayerMessage(yResp, False, 0)

            Dim yResp(5) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerLoginResponse).CopyTo(yResp, 0)
            System.BitConverter.GetBytes(lPlayerID).CopyTo(yResp, 2)
            oPlayer.SendPlayerMessage(yResp, False, 0)

            'ReDim yResp(2)
            'System.BitConverter.GetBytes(GlobalMessageCode.ePlayerIsDead).CopyTo(yResp, 0)
            'yResp(2) = 0
            'oPlayer.oSocket.SendData(yResp)
        End If

        'And save the player
        oPlayer.SaveObject(False)
        'End If
    End Sub
    Private Sub HandleRequestUndock(ByVal yData() As Byte)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, 8)
        Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
        Dim yCancelRequest As Byte = yData(14)

        'Do this smarter now...
        If lObjID = -1 OrElse iObjTypeID = -1 Then
            Dim oEntity As Epica_Entity = CType(GetEpicaObject(lParentID, iParentTypeID), Epica_Entity)
            If oEntity Is Nothing = False Then
                With oEntity
                    If .Owner.AccountStatus <> AccountStatusType.eActiveAccount AndAlso .Owner.AccountStatus <> AccountStatusType.eOpenBetaAccount AndAlso .Owner.AccountStatus <> AccountStatusType.eTrialAccount AndAlso .Owner.AccountStatus <> AccountStatusType.eMondelisActive Then Return

                    Dim lCP As Int32 = 0
                    Dim lCnt As Int32 = 0

                    For X As Int32 = 0 To .lHangarUB
                        If .lHangarIdx(X) <> -1 Then
                            If yCancelRequest = 0 Then
                                AddToQueue(glCurrentCycle, QueueItemType.eHandleUndockRequest_QIT, .oHangarContents(X).ObjectID, .oHangarContents(X).ObjTypeID, .ObjectID, .ObjTypeID, 0, 0, 0, 0)
                                lCP += CInt(CType(.oHangarContents(X), Epica_Entity).ExpLevel)
                                lCnt += 1
                            Else
                                CancelDockingQueueItem(.oHangarContents(X).ObjectID, .oHangarContents(X).ObjTypeID)
                            End If
                        End If
                    Next X
                    If yCancelRequest <> 0 Then
                        CancelDockingQueueItem(.ObjectID, .ObjTypeID)
                        LogEvent(LogEventType.ExtensiveLogging, "CancelLaunchAll: " & .Owner.ObjectID & ". Cnt: " & lCnt)
                    Else
                        Dim lCurCP As Int32 = .Owner.oBudget.GetEnvirCPUsage(lParentID, iParentTypeID)
                        LogEvent(LogEventType.ExtensiveLogging, "LaunchAll: " & .Owner.ObjectID & ". CurCP: " & lCurCP & ". Addtl CP: " & (lCP \ 25) & ". Cnt: " & lCnt)
                    End If
                End With
            End If
        Else

            Dim oEntity As Epica_Entity = Nothing
            If iObjTypeID = ObjectType.eUnit Then
                oEntity = GetEpicaUnit(lObjID)
            ElseIf iObjTypeID = ObjectType.eFacility Then
                oEntity = GetEpicaFacility(lObjID)
            ElseIf iObjTypeID = ObjectType.eUnitDef Then
                'Ok, all units of this unit def
                Dim lCnt As Int32 = 1
                If yData.GetUpperBound(0) >= 15 Then lCnt = yData(15)

                oEntity = GetEpicaEntity(lParentID, iParentTypeID)
                If oEntity Is Nothing Then Return

                'Ok, go through the hangar
                With oEntity
                    For X As Int32 = 0 To .lHangarUB
                        If .lHangarIdx(X) <> -1 Then

                            If .oHangarContents(X).ObjTypeID = ObjectType.eUnit Then
                                Dim oTmp As Unit = CType(.oHangarContents(X), Unit)
                                If oTmp Is Nothing = False Then
                                    If oTmp.EntityDef.ObjectID = lObjID AndAlso oTmp.EntityDef.ObjTypeID = iObjTypeID Then
                                        If yCancelRequest <> 0 Then
                                            CancelDockingQueueItem(oTmp.ObjectID, oTmp.ObjTypeID)
                                        Else
                                            AddToQueue(glCurrentCycle, QueueItemType.eHandleUndockRequest_QIT, .oHangarContents(X).ObjectID, .oHangarContents(X).ObjTypeID, .ObjectID, .ObjTypeID, 0, 0, 0, 0)
                                        End If

                                        lCnt -= 1
                                        If lCnt < 1 Then Exit For
                                    End If
                                End If
                            End If
                        End If
                    Next X
                End With
                Return

            End If
            If oEntity Is Nothing = False Then
                With CType(oEntity.ParentObject, Epica_GUID)
                    If .ObjectID <> lParentID OrElse .ObjTypeID <> iParentTypeID Then
                        LogEvent(LogEventType.Warning, "HandleRequestUndock: Parent being undocked from is not what I believe it is")
                    End If
                End With
            End If

            'ok, add it to our queue
            If yCancelRequest <> 0 Then
                CancelDockingQueueItem(lObjID, iObjTypeID)
            Else : AddToQueue(glCurrentCycle, EngineCode.QueueItemType.eHandleUndockRequest_QIT, lObjID, iObjTypeID, lParentID, iParentTypeID, 0, 0, 0, 0)
            End If
        End If
    End Sub
    Private Sub HandleSaveAndUnloadInstance(ByRef yData() As Byte, ByVal lIndex As Int32)
        'Ok, a region server is telling us that we have all of the updated data for the instance. 
        'we now save our instance data and unload it
        Dim lInstanceID As Int32 = System.BitConverter.ToInt32(yData, 2)

        Dim oPlanet As Planet = GetEpicaPlanet(lInstanceID)
        If oPlanet Is Nothing Then Return
        If oPlanet.oDomain Is Nothing Then Return
        If oPlanet.oDomain.DomainSocket Is Nothing Then Return

        Dim lCurUB As Int32 = -1

        'Save and remove our colony
        If glColonyIdx Is Nothing = False Then lCurUB = Math.Min(glColonyUB, glColonyIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If glColonyIdx(X) <> -1 Then
                Dim oColony As Colony = goColony(X)
                If oColony Is Nothing = False AndAlso oColony.ParentObject Is Nothing = False Then
                    With CType(oColony.ParentObject, Epica_GUID)
                        If .ObjectID <> lInstanceID OrElse .ObjTypeID <> ObjectType.ePlanet Then Continue For
                    End With

                    If oColony.SaveObject() = False Then
                        LogEvent(LogEventType.CriticalError, "HandleSaveAndUnloadInstance could not save Colony!")
                    End If
                    glColonyIdx(X) = -1
                    goColony(X) = Nothing
                End If
            End If
        Next X


        'Ok, now remove all units
        lCurUB = -1
        If glUnitIdx Is Nothing = False Then lCurUB = Math.Min(glUnitUB, glUnitIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If glUnitIdx(X) <> -1 Then
                Dim oUnit As Unit = goUnit(X)
                If oUnit Is Nothing = False AndAlso oUnit.ParentObject Is Nothing = False Then
                    With CType(oUnit.ParentObject, Epica_GUID)
                        If .ObjTypeID = ObjectType.eFacility Then
                            Dim oFac As Facility = CType(oUnit.ParentObject, Facility)
                            If oFac Is Nothing = False Then
                                If oFac.ParentObject Is Nothing = False Then
                                    With CType(oFac.ParentObject, Epica_GUID)
                                        If .ObjectID <> lInstanceID OrElse .ObjTypeID <> ObjectType.ePlanet Then Continue For
                                    End With
                                End If
                            End If
                        ElseIf .ObjTypeID = ObjectType.eUnit Then
                            Dim oTmpUnit As Unit = CType(oUnit.ParentObject, Unit)
                            If oTmpUnit Is Nothing = False Then
                                If oTmpUnit.ParentObject Is Nothing = False Then
                                    With CType(oTmpUnit.ParentObject, Epica_GUID)
                                        If .ObjectID <> lInstanceID OrElse .ObjTypeID <> ObjectType.ePlanet Then Continue For
                                    End With
                                End If
                            End If
                        ElseIf .ObjectID <> lInstanceID OrElse .ObjTypeID <> ObjectType.ePlanet Then
                            Continue For
                        End If
                    End With

                    'Ok, we are good... save the unit
                    If oUnit.SaveObject() = False Then
                        LogEvent(LogEventType.CriticalError, "HandleSaveAndUnloadInstance could not save unit!")
                    End If
                    glUnitIdx(X) = -1
                End If
            End If
        Next X

        'Now, the facilities
        lCurUB = -1
        If glFacilityIdx Is Nothing = False Then lCurUB = Math.Min(glFacilityUB, glFacilityIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If glFacilityIdx(X) <> -1 Then
                Dim oFac As Facility = goFacility(X)
                If oFac Is Nothing = False AndAlso oFac.ParentObject Is Nothing = False Then
                    With CType(oFac.ParentObject, Epica_GUID)
                        If .ObjectID <> lInstanceID OrElse .ObjTypeID <> ObjectType.ePlanet Then
                            Continue For
                        End If
                    End With

                    'Ok, we are good... save the unit
                    If oFac.SaveObject() = False Then
                        LogEvent(LogEventType.CriticalError, "HandleSaveAndUnloadInstance could not save Facility!")
                    End If
                    RemoveLookupFacility(oFac.ObjectID, oFac.ObjTypeID)
                    glFacilityIdx(X) = -1
                    goFacility(X) = Nothing
                End If
            End If
        Next X

        'Finally, the mineral cache
        lCurUB = -1
        If glMineralCacheIdx Is Nothing = False Then lCurUB = Math.Min(glMineralCacheIdx.GetUpperBound(0), glMineralCacheUB)
        For X As Int32 = 0 To lCurUB
            If glMineralCacheIdx(X) <> -1 Then
                Dim oCache As MineralCache = goMineralCache(X)
                If oCache Is Nothing = False AndAlso oCache.ParentObject Is Nothing = False Then
                    With CType(oCache.ParentObject, Epica_GUID)
                        If .ObjectID <> lInstanceID OrElse .ObjTypeID <> ObjectType.ePlanet Then Continue For
                    End With

                    If oCache.SaveObject() = False Then
                        LogEvent(LogEventType.CriticalError, "HandleSaveAndUnloadInstance could not save mineral cache!")
                    End If
                    glMineralCacheIdx(X) = -1
                    goMineralCache(X) = Nothing
                End If
            End If
        Next X

        'Now, the planet
        lCurUB = -1
        If glPlanetIdx Is Nothing = False Then lCurUB = Math.Min(glPlanetIdx.GetUpperBound(0), glPlanetUB)
        For X As Int32 = 0 To lCurUB
            If glPlanetIdx(X) = lInstanceID Then
                glPlanetIdx(X) = -1
            End If
        Next X
    End Sub
    Private Sub HandleServerReady(ByVal yData() As Byte, ByVal Index As Int32)
        Dim bReady As Boolean
        Dim X As Int32

        'Ok, the server is telling us it is ready... it has an IP and Port in it
        moServers(Index).bRegistered = True
        moServers(Index).ClientListenPort = System.BitConverter.ToInt16(yData, 2)
        Array.Copy(yData, 4, moServers(Index).ClientListenIP, 0, 4)

        'Now, once we have set the domain, we need to go through and check if all domains are ready
        bReady = True
        For X = 0 To mlDomainUB
            If moServers(X).bRegistered = False Then
                bReady = False
                Exit For
            End If
        Next X

        If bReady Then
            'Ok, all domains are ready...
            'TODO: what's the next step? Probably tell all domains to begin working... Need to do that...
            'AcceptingClients = True
        End If
    End Sub
    Private Sub HandleSetEntityAI(ByVal yData() As Byte)
        Dim lID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim iTargetingTactics As Int16 = System.BitConverter.ToInt16(yData, 8)
        Dim iCombatTactics As Int32 = System.BitConverter.ToInt32(yData, 10)
        Dim lExt1 As Int32 = System.BitConverter.ToInt32(yData, 14)
        Dim lExt2 As Int32 = System.BitConverter.ToInt32(yData, 18)

        Dim oEntity As Epica_Entity = CType(GetEpicaObject(lID, iTypeID), Epica_Entity)

        If oEntity Is Nothing = False Then
            oEntity.iCombatTactics = iCombatTactics
            oEntity.iTargetingTactics = iTargetingTactics
            oEntity.lExtendedCT_1 = lExt1
            oEntity.lExtendedCT_2 = lExt2
            oEntity.bNeedsSaved = True
        End If
    End Sub
    Private Sub HandleUndockCommandFinished(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        Dim oEntity As Epica_Entity = GetEpicaEntity(lObjID, iObjTypeID)
        If oEntity Is Nothing = False Then

            Select Case oEntity.uUndockQueueItem.lEventType
                Case QueueItemType.eUndockAndReturnToMine_QIT, QueueItemType.eUndockAndReturnToRefinery_QIT
                    'If oUnit.ObjTypeID = ObjectType.eUnit Then oUnit.ProcessNextRouteItem()
                    If oEntity.ObjTypeID = ObjectType.eUnit Then CType(oEntity, Unit).ProcessNextRouteItem()
                Case QueueItemType.eUndock_To_Reinforce
                    If oEntity.ObjTypeID = ObjectType.eUnit Then
                        Dim lFleetID As Int32 = CType(oEntity, Unit).lFleetID
                        If lFleetID > 0 Then
                            Dim oGroup As UnitGroup = GetEpicaUnitGroup(lFleetID)
                            With oEntity.uUndockQueueItem
                                If oGroup Is Nothing = False Then oGroup.FinalizeLaunchToReinforce(oEntity, .lExtended1, .lExtended2, .lExtended3, .lExtended4)
                            End With
                        End If
                    End If
                Case Else
                    Dim lParentID As Int32 = System.BitConverter.ToInt32(yData, 8)
                    Dim iParentTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)
                    Dim oParent As Epica_Entity = GetEpicaEntity(lParentID, iParentTypeID)
                    If oParent Is Nothing = False Then
                        oParent.SendUndockFirstWaypoint(oEntity)
                    End If
            End Select
        End If
    End Sub
    Private Sub HandleUpdateAndSave(ByVal yData() As Byte, ByVal lDomainIdx As Int32)
        'Domain Server is ONLY going to tell us about facilities and units...
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)

        Dim oObject As Epica_Entity = CType(GetEpicaObject(lObjID, iObjTypeID), Epica_Entity)
        Dim lPos As Int32 = 8

        Dim lParentID As Int32
        Dim iParentTypeID As Int16
        Dim bFacPowered As Boolean = False
        Dim bMoveLock As Boolean = False

        If oObject Is Nothing Then Return

        'The region server really only truly has jurisdiction over these:
        With oObject
            .LocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .LocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .LocAngle = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .Fuel_Cap = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ExpLevel = yData(lPos) : lPos += 1
            .iTargetingTactics = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .iCombatTactics = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lExtendedCT_1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lExtendedCT_2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            lParentID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            iParentTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            'For Battlegroups in inter-system movement
            If .ObjTypeID = ObjectType.eUnit Then
                'However, if the ParentType is in a Unit or Facility, then we don't care
                If iParentTypeID <> ObjectType.eUnit AndAlso iParentTypeID <> ObjectType.eFacility Then
                    'Is the unit in a fleet?
                    Dim lFleetID As Int32 = CType(oObject, Unit).lFleetID
                    If lFleetID <> -1 Then
                        'find that battlegroup
                        For lIdx As Int32 = 0 To glUnitGroupUB
                            If glUnitGroupIdx(lIdx) = lFleetID Then
                                'Ok, a target system and a parent object of a galaxy indicates inter-system movement
                                If goUnitGroup(lIdx).lInterSystemTargetID <> -1 AndAlso CType(goUnitGroup(lIdx).oParentObject, Epica_GUID).ObjTypeID = ObjectType.eGalaxy Then
                                    'Ok, we're going to override the ParentID and TypeID
                                    lParentID = CType(goUnitGroup(lIdx).oParentObject, Epica_GUID).ObjectID
                                    iParentTypeID = ObjectType.eGalaxy
                                End If
                                Exit For
                            End If
                        Next lIdx
                    End If
                End If
            End If

            'trust the primary's version of the parent object first and foremost
            If CType(.ParentObject, Epica_GUID).ObjectID <> lParentID OrElse CType(.ParentObject, Epica_GUID).ObjTypeID <> iParentTypeID Then
                LogEvent(LogEventType.Warning, "HandleUpdateAndSave: Region Srvr Reported Invalid Parent Object! EntityID: " & .ObjectID & ", " & .ObjTypeID & ". Parent: " & CType(.ParentObject, Epica_GUID).ObjectID & ", " & CType(.ParentObject, Epica_GUID).ObjTypeID & ". Rgn Parent: " & lParentID & ", " & iParentTypeID)
            End If
            .ParentObject = GetEpicaObject(lParentID, iParentTypeID)
            '.Owner = GetEpicaPlayer(System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
            lPos += 4

            'We want to trust the primary's version of the Facility Powered and Move Lock flags
            bFacPowered = (.CurrentStatus And elUnitStatus.eFacilityPowered) <> 0
            bMoveLock = (.CurrentStatus And elUnitStatus.eMoveLock) <> 0
            .CurrentStatus = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            If (.CurrentStatus And elUnitStatus.eFacilityPowered) <> 0 Then
                'Ok, its powered, but is it supposed to be?
                If bFacPowered = False Then .CurrentStatus = .CurrentStatus Xor elUnitStatus.eFacilityPowered ' .CurrentStatus -= elUnitStatus.eFacilityPowered
            Else
                'ok, its not powered, but is it supposed to be?
                If bFacPowered = True Then .CurrentStatus = .CurrentStatus Or elUnitStatus.eFacilityPowered
            End If
            If (.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then
                'ok, its move locked, but is it supposed to be
                If bMoveLock = False Then .CurrentStatus = .CurrentStatus Xor elUnitStatus.eMoveLock ' .CurrentStatus -= elUnitStatus.eMoveLock
            Else
                'ok, its not move locked, but is it supposed to be?
                If bMoveLock = True Then .CurrentStatus = .CurrentStatus Or elUnitStatus.eMoveLock
            End If

            .Shield_HP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Q1_HP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Q2_HP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Q3_HP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Q4_HP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Structure_HP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            'Now, for the ammo cnts
            Dim lCnt As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Dim lUDWID As Int32
            Dim X As Int32
            Dim Y As Int32
            Dim oTmpDef As Epica_Entity_Def

            If .ObjTypeID = ObjectType.eUnit Then
                oTmpDef = CType(oObject, Unit).EntityDef
            Else : oTmpDef = CType(oObject, Facility).EntityDef
            End If

            'TODO: This assumes that the message will be long enough for this list
            For X = 0 To lCnt - 1
                lUDWID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                For Y = 0 To oTmpDef.WeaponDefUB
                    If oTmpDef.WeaponDefs(Y).ObjectID = lUDWID Then
                        If .lCurrentAmmo Is Nothing = False AndAlso .lCurrentAmmo.Length > Y Then
                            .lCurrentAmmo(Y) = System.BitConverter.ToInt32(yData, lPos)
                        End If
                        Exit For
                    End If
                Next Y

                'regardless of found or not, increment position by 4
                lPos += 4
            Next X

            .DataChanged()
            .bNeedsSaved = True
            .CheckUpdateUnitGroup()
        End With


    End Sub
    Private Sub HandleUpdateEntityAttrs(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        Dim lStatus As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lShield As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lArmorHP(3) As Int32

        For X As Int32 = 0 To 3
            lArmorHP(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Next X

        Dim lStructHP As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lFuelCap As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yExpLevel As Byte = yData(lPos) : lPos += 1

        Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lDestX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim lDestZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oObj As Epica_Entity = Nothing

        If iObjTypeID = ObjectType.eUnit Then
            oObj = GetEpicaUnit(lObjID)
        ElseIf iObjTypeID = ObjectType.eFacility Then
            oObj = GetEpicaFacility(lObjID)
        End If

        If oObj Is Nothing = False Then
            With oObj
                .Shield_HP = lShield
                .Structure_HP = lStructHP
                .Fuel_Cap = lFuelCap
                .ExpLevel = yExpLevel
                .Q1_HP = lArmorHP(0)
                .Q2_HP = lArmorHP(1)
                .Q3_HP = lArmorHP(2)
                .Q4_HP = lArmorHP(3)

                Dim lTemp As Int32 = .CurrentStatus Xor lStatus
                'lTemp now has the differences
                If lTemp <> 0 Then
                    'Ok, there are differences, we need to make sure jurisdiction is properly maintained...
                    Dim lJurisdiction As Int32 = elUnitStatus.eFacilityPowered Or elUnitStatus.eMoveLock
                    'Jurisdiction now has the things the primary is responsible for... now, see if lTemp has these items
                    lJurisdiction = (lJurisdiction And lTemp)
                    'Now, Jurisdiction has the items that lTemp contains that are the Primary's responsibility, remove them
                    lTemp -= lJurisdiction
                    'Now, OR ltemp because it has everything we need
                    .CurrentStatus = .CurrentStatus Or lTemp
                End If

                .LocX = lLocX
                .LocZ = lLocZ
                .bNeedsSaved = True
            End With

            If iObjTypeID = ObjectType.eUnit Then
                With CType(oObj, Unit)
                    .DestX = lDestX
                    .DestZ = lDestZ
                End With
            End If
        End If

    End Sub
    Private Sub HandleUpdateUnitExpLevel(ByRef yData() As Byte)
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
        Dim yLvl As Byte = yData(8)

        If iObjTypeID = ObjectType.eUnit Then
            Dim oUnit As Unit = GetEpicaUnit(lObjID)
            If oUnit Is Nothing = False Then
                oUnit.ExpLevel = yLvl
                oUnit.bNeedsSaved = True
            End If
        ElseIf iObjTypeID = ObjectType.eFacility Then
            Dim oFacility As Facility = GetEpicaFacility(lObjID)
            If oFacility Is Nothing = False Then
                oFacility.ExpLevel = yLvl
                oFacility.bNeedsSaved = True
            End If
        End If
    End Sub

    Private Sub HandleSetTetherPoint(ByVal yData() As Byte, ByVal lIndex As Int32)
        Dim lPos As Int32 = 2

        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yAction As Byte = yData(lPos) : lPos += 1

        For X As Int32 = 0 To lCnt - 1
            Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            Dim oUnit As Unit = GetEpicaUnit(lObjID)
            If oUnit Is Nothing = False Then
                With oUnit
                    If yAction = 0 Then
                        'clearing
                        .TetherPointX = Int32.MinValue
                        .TetherPointZ = Int32.MinValue
                    Else
                        'setting
                        If (.CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
                            .TetherPointX = .LocX
                            .TetherPointZ = .LocZ
                        End If
                    End If
                End With
            End If
        Next X
    End Sub

#End Region

End Class
