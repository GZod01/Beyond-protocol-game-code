Option Strict On

Partial Class MsgSystem

	Private WithEvents moPathfinding As NetSock
	Private WithEvents moPathfindingListener As NetSock
	Private myPathfindingData() As Byte

	Public Property AcceptingPathfinding() As Boolean
		Get
			Return mbAcceptingPathfinding
		End Get
		Set(ByVal Value As Boolean)
			Dim lPort As Int32
			If mbAcceptingPathfinding <> Value Then
				mbAcceptingPathfinding = Value
				If mbAcceptingPathfinding Then
					Dim oINI As New InitFile()
					Try
						lPort = CInt(Val(oINI.GetString("SETTINGS", "PathfindingListenerPort", "0")))
						If lPort = 0 Then Err.Raise(-1, "AcceptingPathfinding", "Could not get Pathfinding Listen Port Number from INI")
						'moPathfinding = Nothing
						If moPathfindingListener Is Nothing Then moPathfindingListener = New NetSock()
						moPathfindingListener.PortNumber = lPort
						moPathfindingListener.Listen()
					Catch
						MsgBox("An error was caught in " & Err.Source & ":" & vbCrLf & Err.Description, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, Err.Source)
						mbAcceptingPathfinding = False
					Finally
						oINI = Nothing
					End Try
				Else
					'ok, stop listening
					moPathfindingListener.StopListening()
				End If
			End If
		End Set
	End Property

#Region "Pathfinding Listener"
	Private Sub moPathfindingListener_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moPathfindingListener.onConnectionRequest
		LogEvent(LogEventType.Informational, "Pathfinding Server Connecting...")
		moPathfinding = New NetSock(oClient)
		moPathfinding.MakeReadyToReceive()
		moPathfinding.lAppType = MsgMonitor.eMM_AppType.PathfindingServer
		moPathfinding.lSpecificID = -1
		AcceptingPathfinding = False
	End Sub
#End Region

#Region "Pathfinding Server Handling"
	Private Sub moPathfinding_onConnect(ByVal Index As Integer) Handles moPathfinding.onConnect
		'
	End Sub

	Private Sub moPathfinding_onConnectionRequest(ByVal Index As Integer, ByVal oClient As System.Net.Sockets.Socket) Handles moPathfinding.onConnectionRequest
		'If AcceptingPathfinding Then
		'    AcceptingPathfinding = False
		'    moPathfinding.StopListening()
		'    'moPathfinding = New NetSock(oClient)
		'    moPathfinding = New NetSock(oClient)
		'    moPathfinding.MakeReadyToReceive()
		'End If
	End Sub

	Private Sub moPathfinding_onDataArrival(ByVal Index As Integer, ByVal Data() As Byte, ByVal TotalBytes As Integer) Handles moPathfinding.onDataArrival
		Dim iMsgID As Int16

		iMsgID = System.BitConverter.ToInt16(Data, 0)
		If mb_MONITOR_MSG_ACTIVITY = True Then moMonitor.AddInMsg(iMsgID, MsgMonitor.eMM_AppType.PathfindingServer, Data.Length, -1)

		Select Case iMsgID
			Case GlobalMessageCode.eUpdateEntityLoc
				HandleUpdateEntityLoc(Data)
			Case GlobalMessageCode.eEntityChangeEnvironment
				'LogEvent(GlobalVars.LogEventType.Informational, "Pathfinding Server Changing Entity Environment")
				HandleEntityChangeEnvironment(Data)
			Case GlobalMessageCode.eSetMiningLoc
				HandleSetMiningLoc(Data)
			Case GlobalMessageCode.eSetEntityProduction
				HandleSetEntityProduction(-1, Data)
			Case GlobalMessageCode.ePathfindingConnectionInfo
				myPathfindingData = Data
				LogEvent(LogEventType.Informational, "Received Pathfinding Connection Info")
			Case GlobalMessageCode.eRequestStarTypes
				LogEvent(LogEventType.Informational, "Pathfinding Server Requesting Star Types")
				HandleRequestStarTypes(Index, ConnectionType.ePathfindingServerApp)
			Case GlobalMessageCode.eRequestGalaxyAndSystems
				LogEvent(GlobalVars.LogEventType.Informational, "Pathfinding Server Requesting Galaxy and Systems")
				HandleRequestGalaxy(Index, ConnectionType.ePathfindingServerApp)
			Case GlobalMessageCode.eRequestSystemDetails
				LogEvent(LogEventType.Informational, "Pathfinding Server Requesting System Details - " & System.BitConverter.ToInt32(Data, 2).ToString)
				HandleRequestSystemDetails(Index, ConnectionType.ePathfindingServerApp, Data)
			Case GlobalMessageCode.ePFRequestEntitys
				HandlePFRequestEntities()
			Case GlobalMessageCode.eSetRepairTarget, GlobalMessageCode.eSetDismantleTarget
				HandleSetMaintenanceTarget(Data, iMsgID, True)
			Case GlobalMessageCode.eJumpTarget
				HandleJumpTarget(Data)
			Case GlobalMessageCode.eAlertDestinationReached
				HandleAlertDestReached(Data)
		End Select
	End Sub

	Private Sub moPathfinding_onDisconnect(ByVal Index As Integer) Handles moPathfinding.onDisconnect
		LogEvent(LogEventType.Informational, "Pathfinding Server Disconnected")
	End Sub

	Private Sub moPathfinding_onError(ByVal Index As Integer, ByVal Description As String) Handles moPathfinding.onError
		LogEvent(LogEventType.Informational, "Pathfinding Server Error: " & Description)
	End Sub

	Private Sub moPathfinding_onSendComplete(ByVal Index As Integer, ByVal DataSize As Integer) Handles moPathfinding.onSendComplete
		'
	End Sub
#End Region

#Region "  Message Handlers  "
	'Only called from the pathfinding server when an entity has finalized its stop event
	Private Sub HandleUpdateEntityLoc(ByRef yData() As Byte)
		Dim lPos As Int32 = 2
		Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim lX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

		Dim oObj As Epica_Entity = Nothing

		If iObjTypeID = ObjectType.eUnit Then
			oObj = GetEpicaUnit(lObjID)
		ElseIf iObjTypeID = ObjectType.eFacility Then
			oObj = GetEpicaFacility(lObjID)
		End If

		If oObj Is Nothing = False Then
			With oObj
				.LocX = lX
				.LocZ = lZ
				.LocAngle = iA
			End With

			If oObj.ObjTypeID = ObjectType.eUnit Then
				With CType(oObj, Unit)

					If .Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID AndAlso .yProductionType = ProductionType.eFacility Then
						Dim oTmp As Epica_GUID = CType(.ParentObject, Epica_GUID)
						If oTmp Is Nothing Then Return
						If goAureliusAI Is Nothing = False Then goAureliusAI.HandlePirateStopped(oTmp.ObjectID, oTmp.ObjTypeID, .ObjectID, .ObjTypeID)
					End If

                    If .lRouteUB <> -1 AndAlso .lCurrentRouteIdx > -1 Then
                        If .bRoutePaused = False Then
                            'there is no action to do if the unit is moving to a point...
                            'verify that the current action is a non-dock event
                            Dim oTmp As Epica_GUID = .uRoute(.lCurrentRouteIdx).oDest
                            If oTmp Is Nothing = True OrElse (oTmp.ObjTypeID <> ObjectType.eUnit AndAlso oTmp.ObjTypeID <> ObjectType.eFacility) Then
                                .ProcessNextRouteItem()
                            End If
                        End If
                    End If
				End With
			End If
		End If
	End Sub
	Private Sub HandleEntityChangeEnvironment(ByVal yData() As Byte)
		'2 byte msg code, 4 Obj Id, 2 typeid, 4 X, 4 Z, 2 A, 4 EnvirID, 2 EnvirTypeID, 1 ChangeEnvirType
		Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
		Dim lX As Int32 = System.BitConverter.ToInt32(yData, 8)
		Dim lZ As Int32 = System.BitConverter.ToInt32(yData, 12)
		Dim iA As Int16 = System.BitConverter.ToInt16(yData, 16)
		Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, 18)
		Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, 22)
		Dim yChangeType As Byte = yData(24)

		Dim X As Int32
		Dim oEntity As Epica_Entity = Nothing
		Dim oEnvir As Object
		Dim oDef As Object = Nothing

		'Dim sResult As String

		If iObjTypeID = ObjectType.eUnit Then
			For X = 0 To glUnitUB
				If glUnitIdx(X) = lObjID Then
					oEntity = goUnit(X)
					If oEntity Is Nothing = False Then oDef = CType(oEntity, Unit).EntityDef
					Exit For
				End If
			Next X
		ElseIf iObjTypeID = ObjectType.eFacility Then
			For X = 0 To glFacilityUB
				If glFacilityIdx(X) = lObjID Then
					oEntity = goFacility(X)
					If oEntity Is Nothing = False Then oDef = CType(oEntity, Facility).EntityDef
					Exit For
				End If
			Next X
		End If

		If oEntity Is Nothing = False Then

			oEntity.RallyPointX = Int32.MinValue
			oEntity.RallyPointZ = Int32.MinValue

			'Update the entity's stats
			oEnvir = GetEpicaObject(lEnvirID, iEnvirTypeID)

			If yChangeType = ChangeEnvironmentType.eDocking Then
				oEntity.LocX = lX : oEntity.LocZ = lZ
				If oEntity.ObjTypeID = ObjectType.eUnit Then
					CType(oEntity, Unit).DataChanged()
				Else : CType(oEntity, Facility).DataChanged()
				End If

				'LogEvent(LogEventType.ExtensiveLogging, "Adding Dock Event " & lObjID & ", " & iObjTypeID & " to " & lEnvirID & ", " & iEnvirTypeID)

				''send movelock which is a SetEntityProdSucceed
				'Dim yMoveLock(7) As Byte
				'System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdSucceed).CopyTo(yMoveLock, 0)
				'oEntity.GetGUIDAsString.CopyTo(yMoveLock, 2)
				'Dim iTemp As Int16 = CType(oEntity.ParentObject, Epica_GUID).ObjTypeID
				'If iTemp = ObjectType.ePlanet Then
				'	CType(oEntity.ParentObject, Planet).oDomain.DomainSocket.SendData(yMoveLock)
				'ElseIf iTemp = ObjectType.eSolarSystem Then
				'	CType(oEntity.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMoveLock)
				'End If

				AddToQueue(glCurrentCycle, EngineCode.QueueItemType.eHandleDockRequest_QIT, lObjID, iObjTypeID, lEnvirID, iEnvirTypeID, 0, 0, 0, 0)
				'Dim lQueueIdx As Int32 = AddToQueueResult(glCurrentCycle, EngineCode.QueueItemType.eHandleDockRequest_QIT, lObjID, iObjTypeID, lEnvirID, iEnvirTypeID, 0, 0, 0, 0)
				'If lQueueIdx > -1 Then
				'	'LogEvent(GlobalVars.LogEventType.ExtensiveLogging, "Dock Event Added")
				'Else : LogEvent(LogEventType.Warning, "Dock Event Not added for " & lObjID & ", " & iObjTypeID & " to " & lEnvirID & ", " & iEnvirTypeID)
				'End If
			ElseIf yChangeType = ChangeEnvironmentType.eSystemToSystem Then
				'Ok, this is a special case...
				If iObjTypeID = ObjectType.eUnit Then
					Dim lFleetID As Int32 = CType(oEntity, Unit).lFleetID
					If lFleetID <> -1 Then
						Dim oUnitGroup As UnitGroup = GetEpicaUnitGroup(lFleetID)
						If oUnitGroup Is Nothing = False Then
							'TODO: Place a movelock on the unit, if it moves, it will need to regroup???
							oUnitGroup.UnitReady(oEntity.ObjectID)
						End If
					End If
				End If
			Else
				'sResult = "Entity's Parent Before: " & CType(oEntity.ParentObject, Epica_GUID).ObjectID.ToString & " type " & CType(oEntity.ParentObject, Epica_GUID).ObjTypeID.ToString

                'ok, the unit is changing environments... get their current envir...
                Dim lParentX As Int32 = lX
                Dim lParentZ As Int32 = lZ
                If oEntity.ParentObject Is Nothing = False Then
                    Dim iTemp As Int16 = CType(oEntity.ParentObject, Epica_GUID).ObjTypeID

                    If iTemp = ObjectType.ePlanet Then
                        With CType(oEntity.ParentObject, Planet)
                            lParentX = .LocX
                            lParentZ = .LocZ

                            If Math.Abs(lParentX - lX) > 20000 Then
                                lX = lParentX
                            End If
                            If Math.Abs(lParentZ - lZ) > 20000 Then
                                lZ = lParentZ
                            End If
                        End With
                    End If
                End If

				oEntity.ParentObject = oEnvir
				oEntity.LocX = lX
				oEntity.LocZ = lZ
				If oEntity.ObjTypeID = ObjectType.eUnit Then
					CType(oEntity, Unit).DataChanged()
				Else : CType(oEntity, Facility).DataChanged()
				End If
				oEntity.CheckUpdateUnitGroup()

				If oEnvir Is Nothing = False Then
					'Ensure the server has the definition object, and add the object
					If CType(oEnvir, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
						CType(oEnvir, Planet).oDomain.DomainSocket.SendData(GetAddObjectMessage(oDef, GlobalMessageCode.eAddObjectCommand))
						Dim yTmpMsg() As Byte = CType(oDef, Epica_Entity_Def).GetCriticalHitChanceMsg()
						If yTmpMsg Is Nothing = False Then CType(oEnvir, Planet).oDomain.DomainSocket.SendData(yTmpMsg)
						CType(oEnvir, Planet).oDomain.DomainSocket.SendData(GetAddObjectMessage(oEntity, GlobalMessageCode.eAddObjectCommand_CE))
					ElseIf CType(oEnvir, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
						CType(oEnvir, SolarSystem).oDomain.DomainSocket.SendData(GetAddObjectMessage(oDef, GlobalMessageCode.eAddObjectCommand))
						Dim yTmpMsg() As Byte = CType(oDef, Epica_Entity_Def).GetCriticalHitChanceMsg()
						If yTmpMsg Is Nothing = False Then CType(oEnvir, SolarSystem).oDomain.DomainSocket.SendData(yTmpMsg)
						CType(oEnvir, SolarSystem).oDomain.DomainSocket.SendData(GetAddObjectMessage(oEntity, GlobalMessageCode.eAddObjectCommand_CE))
					End If
				End If
			End If
		End If

	End Sub
	Private Sub HandleSetMiningLoc(ByVal yData() As Byte)
		Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, 6)
		Dim lCacheID As Int32 = System.BitConverter.ToInt32(yData, 8)
		Dim iCacheTypeID As Int16 = System.BitConverter.ToInt16(yData, 12)

		Dim oUnit As Unit = GetEpicaUnit(lObjID)
		If oUnit Is Nothing Then Return

		If iCacheTypeID = ObjectType.eComponentCache Then
			Dim oCache As ComponentCache = Nothing
			For X As Int32 = 0 To glComponentCacheUB
				If glComponentCacheIdx(X) = lCacheID Then
					With goComponentCache(X)
						oUnit.LocX = .LocX
						oUnit.LocZ = .LocZ

						If (oUnit.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
							Dim lQty As Int32 = .Quantity
							lQty = Math.Min(lQty, oUnit.Cargo_Cap)

							'Get the domain socket
							Dim oDomainSocket As NetSock = Nothing
							If CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
								oDomainSocket = CType(.ParentObject, Planet).oDomain.DomainSocket
							ElseIf CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
								oDomainSocket = CType(.ParentObject, SolarSystem).oDomain.DomainSocket
							End If

							If lQty = .Quantity Then
								'Put the cache into the unit and remove it from the region server owning it

								If oDomainSocket Is Nothing = False Then
									Dim yRemove(8) As Byte
									System.BitConverter.GetBytes(GlobalMessageCode.eRemoveObject).CopyTo(yRemove, 0)
									.GetGUIDAsString.CopyTo(yRemove, 2)
									yRemove(8) = 0
									oDomainSocket.SendData(yRemove)
								End If

								.ParentObject = oUnit
								oUnit.AddCargoRef(CType(goComponentCache(X), Epica_GUID))
							Else
								'add a cache to the unit and update region server owning it
								oUnit.AddComponentCacheToCargo(.ComponentID, .ComponentTypeID, lQty, .ComponentOwnerID)
								.Quantity -= lQty

								If oDomainSocket Is Nothing = False Then
									.DataChanged()
									oDomainSocket.SendData(GetAddObjectMessage(goComponentCache(X), GlobalMessageCode.eAddObjectCommand))
								End If
							End If
						End If
					End With
					Exit For
				End If
			Next X
		ElseIf iCacheTypeID = ObjectType.eMineralCache Then
			For X As Int32 = 0 To glMineralCacheUB
				If glMineralCacheIdx(X) = lCacheID Then
					Dim oCache As MineralCache = goMineralCache(X)
					If oCache Is Nothing = False Then
						oUnit.lCacheIndex = X
						oUnit.lCacheID = oCache.ObjectID
						oUnit.LocX = oCache.LocX
						oUnit.LocZ = oCache.LocZ

						AddToQueue(glCurrentCycle + 5, EngineCode.QueueItemType.eHandleBeginMining_QIT, _
						   lObjID, iObjTypeID, lCacheID, iCacheTypeID, 0, 0, 0, 0)
					End If

					Exit For
				End If
			Next X
		End If

	End Sub
	Private Sub HandlePFRequestEntities()
		Dim lPos As Int32
		Dim X As Int32
		Dim yMsg() As Byte = Nothing


		lPos = 0

		'This message is going to be:
		'Cnt - 4 bytes
		'  Obj_GUID (6)
		'  Parent_GUID (6)
		'  LocX (4)
		'  LocZ (4)
		'  iModelID (2)

		Dim yCache(200000) As Byte
		Dim yFinal(0) As Byte
		Dim yTemp() As Byte

		lPos = 0
		Dim lSingleMsgLen As Int32 = -1
		For X = 0 To glFacilityUB
			If glFacilityIdx(X) <> -1 Then
				Dim oFac As Facility = goFacility(X)
				If oFac Is Nothing Then Continue For
				If oFac.ParentObject Is Nothing Then Continue For
				Dim iTemp As Int16 = CType(oFac.ParentObject, Epica_GUID).ObjTypeID
				If iTemp <> ObjectType.ePlanet OrElse iTemp <> ObjectType.eSolarSystem Then Continue For

				yTemp = GetPathfindingNewEntityMsg(CType(oFac, Epica_GUID), CType(oFac.ParentObject, Epica_GUID), oFac.LocX, oFac.LocZ, oFac.EntityDef.ModelID)
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
			moPathfinding.SendLenAppendedData(yFinal)
			Threading.Thread.Sleep(1)
			If MsgSystem.mb_MONITOR_MSG_ACTIVITY = True Then goMsgSys.moMonitor.AddOutMsg(GlobalMessageCode.eAddObjectCommand, MsgMonitor.eMM_AppType.RegionServer, yFinal.Length, moPathfinding.SocketIndex)
		End If

		ReDim yCache(200000)
		ReDim yFinal(0)
		lPos = 0
		lSingleMsgLen = -1
		For X = 0 To glUnitUB
			If glUnitIdx(X) <> -1 Then
				Dim oUnit As Unit = goUnit(X)
				If oUnit Is Nothing Then Continue For
				If oUnit.ParentObject Is Nothing Then Continue For
				Try
					Dim iTemp As Int16 = CType(oUnit.ParentObject, Epica_GUID).ObjTypeID
					If iTemp <> ObjectType.ePlanet OrElse iTemp <> ObjectType.eSolarSystem Then Continue For
				Catch
					Continue For
				End Try

				yTemp = GetPathfindingNewEntityMsg(CType(oUnit, Epica_GUID), CType(oUnit.ParentObject, Epica_GUID), oUnit.LocX, oUnit.LocZ, oUnit.EntityDef.ModelID)
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
			moPathfinding.SendLenAppendedData(yFinal)
			Threading.Thread.Sleep(1)
			If MsgSystem.mb_MONITOR_MSG_ACTIVITY = True Then goMsgSys.moMonitor.AddOutMsg(GlobalMessageCode.eAddObjectCommand, MsgMonitor.eMM_AppType.RegionServer, yFinal.Length, moPathfinding.SocketIndex)
		End If

		ReDim yCache(200000)
		ReDim yFinal(0)
		lPos = 0
		lSingleMsgLen = -1
		For X = 0 To glFormationDefUB
			If glFormationDefIdx(X) <> -1 Then
				Dim oFormation As FormationDef = goFormationDefs(X)
				If oFormation Is Nothing Then Continue For

				yTemp = oFormation.GetAsAddMsg()
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
			moPathfinding.SendLenAppendedData(yFinal)
			Threading.Thread.Sleep(1)
			If MsgSystem.mb_MONITOR_MSG_ACTIVITY = True Then goMsgSys.moMonitor.AddOutMsg(GlobalMessageCode.eAddFormation, MsgMonitor.eMM_AppType.RegionServer, yFinal.Length, moPathfinding.SocketIndex)
		End If


		'For X = 0 To glFacilityUB
		'	If glFacilityIdx(X) <> -1 Then
		'		Me.SendPathfindingNewEntity(CType(goFacility(X), Epica_GUID), CType(goFacility(X).ParentObject, Epica_GUID), goFacility(X).LocX, goFacility(X).LocZ, goFacility(X).EntityDef.ModelID)
		'	End If
		'Next X
		'For X = 0 To glUnitUB
		'	If glUnitIdx(X) <> -1 Then
		'		Me.SendPathfindingNewEntity(CType(goUnit(X), Epica_GUID), CType(goUnit(X).ParentObject, Epica_GUID), goUnit(X).LocX, goUnit(X).LocZ, goUnit(X).EntityDef.ModelID)
		'	End If
		'Next X
		'For X = 0 To glFormationDefUB
		'	If glFormationDefIdx(X) <> -1 Then
		'		moPathfinding.SendData(goFormationDefs(X).GetAsAddMsg())
		'	End If
		'Next X

		ReDim yMsg(5)
		System.BitConverter.GetBytes(GlobalMessageCode.ePFRequestEntitys).CopyTo(yMsg, 0)
		System.BitConverter.GetBytes(0I).CopyTo(yMsg, 2)
		If yMsg Is Nothing = False Then Me.moPathfinding.SendData(yMsg)

	End Sub
	Private Sub HandleJumpTarget(ByRef yData() As Byte)
		Dim lPos As Int32 = 2		'for msgcode
		Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim lJumpID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iJumpTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim lX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iA As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim lCurParentID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iCurParentTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 4

		'Pathfinding server is telling us that a unit is attempting to jump wormhole
		If iEntityTypeID <> ObjectType.eUnit Then Return

		Dim oEntity As Epica_Entity = GetEpicaUnit(lEntityID)
		If oEntity Is Nothing Then Return

		Dim oPlayer As Player = oEntity.Owner
		If oPlayer Is Nothing Then Return

		If iJumpTypeID = ObjectType.eWormhole Then
			If (oPlayer.oSpecials.lSuperSpecials And Player_Specials.SuperSpecialID.eWormholesTraversable) <> 0 Then
				If iCurParentTypeID = ObjectType.eSolarSystem Then
					'so, let's get the wormhole and the system
					Dim oWormhole As Wormhole = GetEpicaWormhole(lJumpID)
                    If oWormhole Is Nothing = False Then

                        Dim lFlags As Int32 = oWormhole.WormholeFlags Or elWormholeFlag.eSystem1Detectable Or elWormholeFlag.eSystem2Detectable
                        If oWormhole.WormholeFlags <> lFlags Then
                            oWormhole.WormholeFlags = lFlags
                            oWormhole.StartCycle = 1
                            oWormhole.QueueMeToSave()

                            If oWormhole.System1 Is Nothing = False Then
                                If oWormhole.System1.SystemName Is Nothing = False AndAlso oWormhole.System1.SystemType = SolarSystem.elSystemType.SpawnSystem Then
                                    oWormhole.System1.SystemType = CByte(SolarSystem.elSystemType.UnlockedSystem)
                                    oWormhole.System1.DataChanged()
                                    gbGalaxyMsgReady = False
                                    oWormhole.System1.SaveObject()

                                    Dim sName As String = BytesToString(oWormhole.System1.SystemName)
                                    If sName.Contains("(S)") = True Then
                                        sName = sName.Replace("(S)", "").Trim
                                        ReDim oWormhole.System1.SystemName(19)
                                        For X As Int32 = 0 To 19
                                            oWormhole.System1.SystemName(X) = 0
                                        Next X
                                        StringToBytes(sName).CopyTo(oWormhole.System1.SystemName, 0)
                                    End If

                                    oWormhole.System1.SaveObject()
                                End If
                            End If
                        End If
                        If oWormhole.System1 Is Nothing = False AndAlso oWormhole.System1.ObjectID = lCurParentID Then
                            'check sys 2
                            If oWormhole.System2 Is Nothing = False AndAlso oWormhole.System2.SystemType = SolarSystem.elSystemType.RespawnSystem Then
                                Return
                            End If
                        Else
                            'check sys 1
                            If oWormhole.System1 Is Nothing = False AndAlso oWormhole.System1.SystemType = SolarSystem.elSystemType.RespawnSystem Then
                                Return
                            End If
                        End If

                        'TODO: check if the wormhole's life cycle has passed
                        Dim yRequest(19) As Byte
                        lPos = 0
                        System.BitConverter.GetBytes(GlobalMessageCode.eJumpTarget).CopyTo(yRequest, lPos) : lPos += 2
                        oEntity.GetGUIDAsString.CopyTo(yRequest, lPos) : lPos += 6
                        oWormhole.GetGUIDAsString.CopyTo(yRequest, lPos) : lPos += 6
                        If oWormhole.System1 Is Nothing = False AndAlso oWormhole.System1.ObjectID = lCurParentID Then
                            oWormhole.System1.GetGUIDAsString.CopyTo(yRequest, lPos) : lPos += 6
                            oWormhole.System1.oDomain.DomainSocket.SendData(yRequest)
                        ElseIf oWormhole.System2 Is Nothing = False AndAlso oWormhole.System2.ObjectID = lCurParentID Then
                            oWormhole.System2.GetGUIDAsString.CopyTo(yRequest, lPos) : lPos += 6
                            oWormhole.System2.oDomain.DomainSocket.SendData(yRequest)
                        Else : LogEvent(LogEventType.PossibleCheat, "Wormhole Jump but wormhole does not connect to parent. S1:" & oWormhole.System1.ObjectID.ToString & " S2:" & oWormhole.System2.ObjectID.ToString & " E:" & oEntity.ObjectID.ToString & " L:" & lCurParentID.ToString & " P:" & oPlayer.ObjectID)
                        End If
                    Else : LogEvent(LogEventType.PossibleCheat, "Wormhole Jump, wormhole is nothing. " & oPlayer.ObjectID)
                    End If
				Else : LogEvent(LogEventType.PossibleCheat, "Wormhole Jump but parent is not system: " & oPlayer.ObjectID)
				End If
			Else : LogEvent(LogEventType.PossibleCheat, "Wormhole Jump but player cannot traverse: " & oPlayer.ObjectID)
			End If
		Else
			'TODO: Figure out other jump types
		End If
	End Sub
	Private Sub HandleAlertDestReached(ByRef yData() As Byte)
		Dim lPos As Int32 = 2
		Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

		If iObjTypeID = ObjectType.eFacility OrElse iObjTypeID = ObjectType.eUnit Then
			Dim oEntity As Epica_Entity = GetEpicaEntity(lObjID, iObjTypeID)
			If oEntity Is Nothing = False Then
				If oEntity.Owner Is Nothing = False Then
					Dim yForward(41) As Byte

					Array.Copy(yData, 0, yForward, 0, 22)
					oEntity.EntityName.CopyTo(yForward, lPos) : lPos += 20
					oEntity.Owner.SendPlayerMessage(yForward, False, AliasingRights.eViewUnitsAndFacilities)
				End If
			End If
		End If
	End Sub

#End Region

End Class
