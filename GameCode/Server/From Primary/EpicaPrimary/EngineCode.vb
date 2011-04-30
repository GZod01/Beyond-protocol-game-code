Option Strict On

Module EngineCode
    'This Module contains all of our engine code that the Primary Server is responsible for
    Public otMainLoop As Threading.Thread

    Public gb_Main_Loop_Running As Boolean
    Public gb_In_Main_Loop As Boolean

    Private mswCredits As Stopwatch         'MSC 12/18/2006

    Public glNextColonyGrowthUpdate As Int32
    Public Const gl_COLONY_GROWTH_INTERVAL As Int32 = 90        'once every 3 seconds
    Public glNextProductionUpdate As Int32
    Public Const gl_PRODUCTION_INTERVAL As Int32 = 30       'once every second
    Private mblLastCensus As Int64                           'NOTE: We will want to remove the Census calcs later...

	Public goAgentEngine As AgentEngine

	Public goAureliusAI As AureliusAI

    Private mlPreviousGTCDeadline As Int32 = 0
    'Private mlPreviousSentimentCheck As Int32 = 0

#Region " Production Manager Code "
	Private mlEntityProducing(-1) As Int32				'Index in the appropriate array
	Private miEntityProducingType(-1) As Int32			'ObjTypeID of the entity (unit or facility) to determine which array to look
	Private mlEntityProducingID(-1) As Int32			'for verification purposes
    Private mlEntityProducingUB As Int32 = -1
    Private mlEntityProducingSyncLock(-1) As Int32
	Public Sub AddEntityProducing(ByVal lEntityIdx As Int32, ByVal iEntityTypeID As Int16, ByVal lEntityID As Int32)
        Dim lIdx As Int32 = -1
        Dim lCurUB As Int32 = -1
        If mlEntityProducing Is Nothing = False Then lCurUB = Math.Min(mlEntityProducingUB, mlEntityProducing.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If mlEntityProducing(X) = lEntityIdx AndAlso miEntityProducingType(X) = iEntityTypeID AndAlso mlEntityProducingID(X) = lEntityID Then
                'Already in place, so return
                Return
            ElseIf lIdx = -1 AndAlso mlEntityProducing(X) < 0 Then
                lIdx = X
            End If
        Next X

		'Ok, if we are here...
        If lIdx = -1 Then
            If True = True Then
                SyncLock mlEntityProducingSyncLock
                    lIdx = mlEntityProducingUB + 1
                    If mlEntityProducing Is Nothing OrElse lIdx > mlEntityProducing.GetUpperBound(0) Then
                        System.GC.Collect()
                        ReDim Preserve mlEntityProducing(mlEntityProducingUB + 1000)
                        ReDim Preserve miEntityProducingType(mlEntityProducingUB + 1000)
                        ReDim Preserve mlEntityProducingID(mlEntityProducingUB + 1000)
                    End If
                    mlEntityProducingUB += 1
                    lIdx = mlEntityProducingUB
                End SyncLock
            Else
                ReDim Preserve mlEntityProducing(mlEntityProducingUB + 1)
                ReDim Preserve miEntityProducingType(mlEntityProducingUB + 1)
                ReDim Preserve mlEntityProducingID(mlEntityProducingUB + 1)
                mlEntityProducingUB += 1
                lIdx = mlEntityProducingUB
            End If
        End If

		mlEntityProducingID(lIdx) = lEntityID
		miEntityProducingType(lIdx) = iEntityTypeID
		mlEntityProducing(lIdx) = lEntityIdx
	End Sub

	Public Sub GetEnvirConstObjects(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lOwnerID As Int32, ByRef oSocket As NetSock)

		Dim lCurUB As Int32 = Math.Min(mlEntityProducingUB, mlEntityProducing.GetUpperBound(0))
		For X As Int32 = 0 To lCurUB ' glUnitUB
			Try
				If miEntityProducingType(X) = ObjectType.eUnit Then
					Dim lIdx As Int32 = mlEntityProducing(X)
					If lIdx <> -1 AndAlso glUnitIdx(lIdx) = mlEntityProducingID(X) Then
						Dim oUnit As Unit = goUnit(lIdx)
						If oUnit Is Nothing = False Then
							If oUnit.Owner.ObjectID = lOwnerID AndAlso oUnit.bProducing = True Then
								Dim bGood As Boolean = False
								With CType(oUnit.ParentObject, Epica_GUID)
									If .ObjectID = lEnvirID AndAlso .ObjTypeID = iEnvirTypeID Then
										bGood = True
									End If
								End With

                                If bGood = True AndAlso oUnit.CurrentProduction Is Nothing = False Then

                                    Dim yProdQueueList() As Byte = oUnit.GetProdQueueList()
                                    If yProdQueueList Is Nothing Then Continue For

                                    Dim yResp(29 + yProdQueueList.Length) As Byte

                                    System.BitConverter.GetBytes(GlobalMessageCode.eEntityProductionStatus).CopyTo(yResp, 0)
                                    oUnit.GetGUIDAsString.CopyTo(yResp, 2)
                                    Dim blTemp As Int64
                                    With oUnit.CurrentProduction

                                        Dim blPointsProd As Int64 = .PointsProduced
                                        Dim lLastUpdate As Int32 = .lLastUpdateCycle
                                        Dim lProd As Int32 = oUnit.mlProdPoints
                                        Dim lFinish As Int32 = .lFinishCycle

                                        blTemp = blPointsProd + ((glCurrentCycle - lLastUpdate) * lProd)

                                        Dim fTmp As Single = CSng((blTemp / .PointsRequired) * 100)
                                        System.BitConverter.GetBytes(fTmp).CopyTo(yResp, 8)
                                        System.BitConverter.GetBytes(lFinish).CopyTo(yResp, 16)
                                        System.BitConverter.GetBytes(.ProductionID).CopyTo(yResp, 20)
                                        System.BitConverter.GetBytes(.ProductionTypeID).CopyTo(yResp, 24)
                                        System.BitConverter.GetBytes(.ProductionItemModelID).CopyTo(yResp, 26)
                                        System.BitConverter.GetBytes(.iProdA).CopyTo(yResp, 28)
                                    End With
                                    yProdQueueList.CopyTo(yResp, 30)

                                    System.BitConverter.GetBytes(glCurrentCycle).CopyTo(yResp, 12)
                                    oSocket.SendData(yResp)
                                End If
							End If
						End If
					End If
				End If
			Catch
			End Try

		Next X

	End Sub

    Private Sub CheckProduction()
        'Ok, go through our array
		Dim lIdx As Int32

		Dim lCurUB As Int32 = Math.Min(mlEntityProducingUB, mlEntityProducing.GetUpperBound(0))
		For X As Int32 = 0 To lCurUB
			Try
				If mlEntityProducing(X) > -1 Then
					lIdx = mlEntityProducing(X)

					If miEntityProducingType(X) = ObjectType.eUnit Then
						'Ok, its a unit producing, so look in the unit list
						If glUnitIdx(lIdx) <> -1 Then
							If glUnitIdx(lIdx) = mlEntityProducingID(X) Then
								Dim oUnit As Unit = goUnit(lIdx)
								If oUnit Is Nothing Then Continue For
								With oUnit
									If .bProducing = True Then
                                        If .CurrentProduction Is Nothing = False Then
                                            If ((.Owner.yPlayerPhase = eyPlayerPhase.eInitialPhase AndAlso ((.Owner.lTutorialStep > 230 AndAlso .Owner.lTutorialStep < 285) OrElse .Owner.lTutorialStep = 223 OrElse .Owner.lTutorialStep = 221)) OrElse gb_IS_TEST_SERVER = True) AndAlso .Owner.ObjectID <> gl_HARDCODE_PIRATE_PLAYER_ID Then 'If gb_IS_TEST_SERVER = True AndAlso .Owner.ObjectID <> gl_HARDCODE_PIRATE_PLAYER_ID Then
                                                'We clear the item before calling production completed
                                                mlEntityProducing(X) = -1
                                                ProductionCompleted(.CurrentProduction, CType(oUnit, Epica_Entity))
                                            Else
                                                If .CurrentProduction.lFinishCycle < glCurrentCycle Then
                                                    If .VerifyCompletion = True Then
                                                        'We clear the item before calling production completed
                                                        mlEntityProducing(X) = -1
                                                        ProductionCompleted(.CurrentProduction, CType(oUnit, Epica_Entity))
                                                    Else : .RecalcProduction()
                                                    End If
                                                End If
                                            End If

                                        Else
                                            .bProducing = False
                                            'No production clears the item fora unit
                                            mlEntityProducing(X) = -1
                                        End If
                                    Else
                                        'Not producing DOES clear the item for a unit
                                        mlEntityProducing(X) = -1
                                    End If
								End With
							End If
						Else : mlEntityProducing(X) = -1
						End If
					Else
						'It has to be a facility producing...
						If glFacilityIdx(lIdx) <> -1 Then
							If glFacilityIdx(lIdx) = mlEntityProducingID(X) Then
								Dim oFac As Facility = goFacility(lIdx)
								If oFac Is Nothing Then Continue For
								With oFac
									'Ok, first, check if current production is nothing
									If .CurrentProduction Is Nothing = False Then
										'Active or Producing being false does not clear the slot
										If .Active = True Then
											If .bProducing = True Then
                                                'Ok, is the production ready?
                                                If ((.Owner.yPlayerPhase = eyPlayerPhase.eInitialPhase AndAlso ((.Owner.lTutorialStep > 230 AndAlso .Owner.lTutorialStep < 285) OrElse .Owner.lTutorialStep = 223 OrElse .Owner.lTutorialStep = 221)) OrElse gb_IS_TEST_SERVER = True) AndAlso .Owner.ObjectID <> gl_HARDCODE_PIRATE_PLAYER_ID Then 'If gb_IS_TEST_SERVER = True AndAlso .Owner.ObjectID <> gl_HARDCODE_PIRATE_PLAYER_ID Then
                                                    Try
                                                        ProductionCompleted(.CurrentProduction, CType(oFac, Epica_Entity))
                                                    Catch
                                                        .bProducing = False
                                                        mlEntityProducing(X) = -1
                                                    End Try
                                                Else
                                                    If .CurrentProduction.lFinishCycle < glCurrentCycle Then
                                                        'Yes, verify it...
                                                        If .VerifyCompletion = True Then
                                                            'ok, we have a completed production...
                                                            Try
                                                                ProductionCompleted(.CurrentProduction, CType(oFac, Epica_Entity))
                                                            Catch
                                                                .bProducing = False
                                                                mlEntityProducing(X) = -1
                                                            End Try
                                                            'we don't clear the slot until the next refresh cycle... and that is only
                                                            '  if CurrentProduction is nothing
                                                        Else
                                                            'Attempt to recalc the production to ensure it has the latest new time
                                                            .RecalcProduction()
                                                        End If
                                                    End If
                                                End If


                                            End If
                                        End If
                                    Else
                                        'ok, clear the slot
                                        .bProducing = False
                                        'clear our slot
                                        mlEntityProducing(X) = -1
                                    End If
                                End With
							End If
						Else
							'Release this slot
							mlEntityProducing(X) = -1
						End If

						End If
				End If
			Catch
				'do nohting
			End Try
		Next X
	End Sub

	Public Function BuildingCommandCenterOrTradepost(ByVal lOwnerID As Int32, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal yCheckProdType As Byte) As Boolean
		Dim lCurUB As Int32 = mlEntityProducingUB
		For X As Int32 = 0 To lCurUB
			If mlEntityProducing(X) <> -1 AndAlso miEntityProducingType(X) = ObjectType.eUnit Then
				Dim lIdx As Int32 = mlEntityProducing(X)
				'Ok, its a unit producing, so look in the unit list
				If glUnitIdx(lIdx) <> -1 Then
					Dim oUnit As Unit = goUnit(lIdx)
					If oUnit Is Nothing Then Continue For
					'is the unit mine?
					If oUnit.Owner.ObjectID = lOwnerID Then
						If oUnit.CurrentProduction Is Nothing = False Then
							If oUnit.CurrentProduction.ProductionTypeID = ObjectType.eFacilityDef Then
								Try
									Dim oDef As FacilityDef = GetEpicaFacilityDef(oUnit.CurrentProduction.ProductionID)
									If oDef Is Nothing = False Then
										'is the facility def produce what we are looking for?
										If oDef.ProductionTypeID = yCheckProdType Then
											'ok, is the unit in the correct environment?
											With CType(oUnit.ParentObject, Epica_GUID)
												If .ObjectID = lEnvirID AndAlso .ObjTypeID = iEnvirTypeID Then
													Return True
												End If
											End With
										End If
									End If
								Catch
								End Try
							End If
						End If
					End If
				End If
			End If
		Next X
		Return False
	End Function

	Public Sub FastForwardProduction(ByVal lPlayerID As Int32, ByVal lProdID As Int32, ByVal iProdTypeID As Int16)
		'Ok, go through our array
		Dim lIdx As Int32

		Dim bFinishAll As Boolean = (lProdID = -1 AndAlso iProdTypeID = -1)

		Dim lCurUB As Int32 = Math.Min(mlEntityProducingUB, mlEntityProducing.GetUpperBound(0))
		For X As Int32 = 0 To lCurUB
			Try
				If mlEntityProducing(X) > -1 Then
					lIdx = mlEntityProducing(X)

					If miEntityProducingType(X) = ObjectType.eUnit Then
						'Ok, its a unit producing, so look in the unit list
						If glUnitIdx(lIdx) <> -1 Then
							If glUnitIdx(lIdx) = mlEntityProducingID(X) Then
								Dim oUnit As Unit = goUnit(lIdx)
								If oUnit Is Nothing Then Continue For
								With oUnit
									If .Owner.ObjectID = lPlayerID AndAlso .bProducing = True Then
										If .CurrentProduction Is Nothing = False Then
											If bFinishAll = True OrElse (.CurrentProduction.ProductionID = lProdID AndAlso .CurrentProduction.ProductionTypeID = iProdTypeID) Then
												.CurrentProduction.PointsProduced = .CurrentProduction.PointsRequired
												.CurrentProduction.lFinishCycle = glCurrentCycle + 1
											End If
										End If 
									End If
								End With
							End If
						End If
					Else
						'It has to be a facility producing...
						If glFacilityIdx(lIdx) <> -1 Then
							If glFacilityIdx(lIdx) = mlEntityProducingID(X) Then
								Dim oFac As Facility = goFacility(lIdx)
								If oFac Is Nothing Then Continue For
								With oFac
									'Ok, first, check if current production is nothing
									If .Owner.ObjectID = lPlayerID AndAlso .CurrentProduction Is Nothing = False Then
										'Active or Producing being false does not clear the slot
										If .Active = True AndAlso .bProducing = True Then
											'Ok, is the production ready?
											If bFinishAll = True OrElse (.CurrentProduction.ProductionID = lProdID AndAlso .CurrentProduction.ProductionTypeID = iProdTypeID) Then
												.CurrentProduction.PointsProduced = .CurrentProduction.PointsRequired * Math.Max(1, .CurrentProduction.lProdCount)
												.CurrentProduction.lFinishCycle = glCurrentCycle + 1
											End If

											If bFinishAll = True Then
												For Z As Int32 = 0 To .ProductionUB
													If .ProductionIdx(Z) <> -1 Then
														.Production(Z).PointsRequired = 1
													End If
												Next Z
											End If
										End If
									End If
								End With
							End If 
						End If

					End If
				End If
			Catch
				'do nohting
			End Try
		Next X
	End Sub
#End Region

#Region " Mining Manager Code "
	Private mlEntityMining(-1) As Int32		'index of the appropriate array
	Private mlEntityMiningID(-1) As Int32	'actual ID for verification purposes
	Private mlEntityMiningUB As Int32 = -1
    Private mlEntityMiningSyncLock(-1) As Int32
	Public Sub AddFacilityMining(ByVal lEntityIdx As Int32, ByVal lFacilityID As Int32)
		Dim lIdx As Int32 = -1
		For X As Int32 = 0 To mlEntityMiningUB
			If mlEntityMining(X) = lEntityIdx AndAlso mlEntityMiningID(X) = lFacilityID Then
				'Already in place, so return
				Return
			ElseIf lIdx = -1 AndAlso mlEntityMining(X) < 0 Then
				lIdx = X
			End If
		Next X

		'Ok, if we are here...
        If lIdx = -1 Then
            If True = True Then
                SyncLock mlEntityMiningSyncLock
                    lIdx = mlEntityMiningUB + 1
                    If mlEntityMining Is Nothing OrElse lIdx > mlEntityMining.GetUpperBound(0) Then
                        System.GC.Collect()
                        ReDim Preserve mlEntityMining(mlEntityMiningUB + 1000)
                        ReDim Preserve mlEntityMiningID(mlEntityMiningUB + 1000)
                    End If
                    mlEntityMiningUB += 1
                    lIdx = mlEntityMiningUB
                End SyncLock
            Else
                ReDim Preserve mlEntityMining(mlEntityMiningUB + 1)
                ReDim Preserve mlEntityMiningID(mlEntityMiningUB + 1)
                mlEntityMiningUB += 1
                lIdx = mlEntityMiningUB
            End If
        End If
		mlEntityMiningID(lIdx) = lFacilityID
		mlEntityMining(lIdx) = lEntityIdx
	End Sub
 
	Private Sub CheckFacilityMining()
		'Ok, go through our array
		Dim lIdx As Int32

		Dim lCurUB As Int32 = Math.Min(mlEntityMiningUB, mlEntityMining.GetUpperBound(0))
		For X As Int32 = 0 To lCurUB
			Try
				If mlEntityMining(X) > -1 Then
					lIdx = mlEntityMining(X)

					'It has to be a facility producing...
					If glFacilityIdx(lIdx) <> -1 Then
						If mlEntityMiningID(X) < 1 Then Continue For
						If mlEntityMiningID(X) <> glFacilityIdx(lIdx) Then
							mlEntityMining(X) = -1
							Continue For
						End If
						Dim oFac As Facility = goFacility(lIdx)
						If oFac Is Nothing Then Continue For
						'If oFac.Active = False Then Continue For

						Dim lFacCargoCap As Int32 = oFac.Cargo_Cap
                        If (oFac.yProductionType = ProductionType.eMining) Then 'AndAlso ((oFac.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0) Then 'AndAlso (lFacCargoCap <> 0) Then
                            'Ok, set our temp cache for use later...
                            Dim lCacheIndex As Int32 = oFac.lCacheIndex
                            If lCacheIndex <> -1 AndAlso glMineralCacheIdx(lCacheIndex) <> -1 Then
                                'ok, let's mine... Mining Facilities mine double the concentration 
                                Dim oCache As MineralCache = goMineralCache(lCacheIndex)
                                If oCache Is Nothing OrElse oCache.ObjectID <> oFac.lCacheID Then
                                    mlEntityMining(X) = -1
                                    oFac.lCacheIndex = -1
                                    Continue For
                                End If
                                Dim lNewQty As Int32 = CInt((oCache.Concentration * 2) + oFac.Owner.oSpecials.yMineralConcentrationBonus)

                                oFac.RecalcProduction()
                                lNewQty = CInt(Math.Ceiling(lNewQty * (oFac.mlProdPoints * 0.001F)))

                                'Verify Quantity vs. Cache.Quantity and Cargo Hold
                                If lNewQty > oCache.Quantity Then lNewQty = oCache.Quantity

                                'Now, verify that quantity versus the mineral cache/facility order
                                If oFac.oMiningBid Is Nothing = False Then
                                    Dim lRemainder As Int32 = oFac.oMiningBid.DeliverMinerals(lNewQty)
                                    lNewQty -= lRemainder
                                Else : lNewQty = 0
                                End If

                                'If lNewQty > lFacCargoCap Then lNewQty = lFacCargoCap

                                'Don't do anything if there is no minerals mined...
                                If lNewQty > 0 Then
                                    'Reduce the cache quantity
                                    oCache.Quantity -= lNewQty

                                    'Test for concentration reduction... mining facility has half chance
                                    If oCache.CacheTypeID = MineralCacheType.eMineable Then oCache.HandleDepletion(2.5F)

                                    'Now add the quantity to my cargo
                                    'oFac.AddMineralCacheToCargo(oCache.oMineral.ObjectID, lNewQty)
                                End If
                            End If

                            'Now, if there are units inside my hangar, let's transfer from cargo to their cargo
                            If (oFac.CurrentStatus And elUnitStatus.eHangarOperational) <> 0 Then
                                For lHIdx As Int32 = 0 To oFac.lHangarUB
                                    If oFac.lHangarIdx(lHIdx) <> -1 Then
                                        Dim oTmpUnit As Epica_Entity = CType(oFac.oHangarContents(lHIdx), Epica_Entity)

                                        oFac.TransferCargo(oTmpUnit, False)
                                        'Should no longer be necessary
                                        'If oTmpUnit.Cargo_Cap = 0 AndAlso oTmpUnit.ObjTypeID = ObjectType.eUnit AndAlso CType(oTmpUnit, Unit).bRoutePaused = False Then
                                        '	'ok, do the undock and return to refinery
                                        '	AddToQueue(glCurrentCycle, EngineCode.QueueItemType.eUndockAndReturnToRefinery_QIT, oTmpUnit.ObjectID, _
                                        '	  oTmpUnit.ObjTypeID, oFac.ObjectID, oFac.ObjTypeID, 0, 0, 0, 0)
                                        'End If

                                        'Now, if the facility is empty, no more cargo to transfer...
                                        If oFac.Cargo_Cap = oFac.EntityDef.Cargo_Cap Then Exit For
                                    End If
                                Next lHIdx
                            End If

                        End If

					Else
						'Release this slot
						mlEntityMining(X) = -1
					End If
				End If
			Catch
				'do nothing
			End Try
        Next X

        'Orbital Mining Platforms
        For X As Int32 = 0 To glPlanetUB
            If glPlanetIdx(X) > -1 Then
                Dim oPlanet As Planet = goPlanet(X)
                If oPlanet Is Nothing = False Then
                    If oPlanet.RingMineralID > 0 AndAlso oPlanet.RingMineralConcentration > 0 AndAlso oPlanet.RingMinerID > 0 Then
                        Dim oFac As Facility = oPlanet.RingMiner()
                        If oFac Is Nothing = False Then

                            Dim lFacCargoCap As Int32 = oFac.Cargo_Cap

                            Dim lNewQty As Int32 = oPlanet.RingMineralConcentration
                            If lNewQty < 1 Then Continue For
                            If lNewQty > lFacCargoCap Then lNewQty = lFacCargoCap

                            oFac.AddMineralCacheToCargo(oPlanet.RingMineralID, lNewQty)
                            ''Now, if there are units inside my hangar, let's transfer from cargo to their cargo
                            'If (oFac.CurrentStatus And elUnitStatus.eHangarOperational) <> 0 Then
                            '    For lHIdx As Int32 = 0 To oFac.lHangarUB
                            '        If oFac.lHangarIdx(lHIdx) <> -1 Then
                            '            Dim oTmpUnit As Epica_Entity = CType(oFac.oHangarContents(lHIdx), Epica_Entity)

                            '            oFac.TransferCargo(oTmpUnit, False)

                            '            'Now, if the facility is empty, no more cargo to transfer...
                            '            If oFac.Cargo_Cap = oFac.EntityDef.Cargo_Cap Then Exit For
                            '        End If
                            '    Next lHIdx
                            'End If

                        End If
                    End If
                End If
            End If
        Next X

	End Sub
#End Region

#Region " Action Queue Functionality "
    Public Enum QueueItemType As Integer
        eHandleDockRequest_QIT = 0
        eHandleUndockRequest_QIT
        eHandleBeginMining_QIT

        eUndockAndReturnToRefinery_QIT
        eUndockAndReturnToMine_QIT

        eReloadRequest

        eSystemToSystemMove

        eTradeEventHalfTime
        eTradeEventFinal

        ePirateSelfDestruct

        eLaunchToAttack

        eReinforcements_Built           'indicates that a unit group needs reinforcements and the facility built them
		eUndock_To_Reinforce			'indicates that a unit is launching to reinforce a unit group

		eGenerateAgent					'generate a new agent

		eRemoveComponentCache
		eRemoveMineralCache

		eBeginRebuilderAI
		eCancelRebuilderAI

		eIronCurtainRaise

		eCheckColonyResearchQueue

		eIronCurtainFall

		eLoadInstancedPlanet
        eSaveAndUnloadInstancedPlanet
		eGenerateNewbieAgent

		eAILaunchAll
		eResetAILaunchAll

		eTutorialFourHourTimerExpire
		eSpawnNextPirateWave

        eReassessMinimumMineralValues

        eReprocessPlayerRel

        eSaveObject

        ePerformLinkTest

        eProcessNextRouteItem

        eReprocessFacilityProdQueue

        eRetryRouteItem
	End Enum

    Public Structure QueueItem
        Public lEventType As Int32      'from QueueItemType
        Public lObjectID As Int32
        Public lObjTypeID As Int32
        Public lTargetID As Int32
        Public lTargetTypeID As Int32

        Public lExtended1 As Int32
        Public lExtended2 As Int32
        Public lExtended3 As Int32
        Public lExtended4 As Int32

        Public lCycle As Int32      'used for addtoqueue (delta queue)
    End Structure

	Private moQueue(-1) As QueueItem
	Private mlQueueUB As Int32 = -1
	Private mlQueueStart(-1) As Int32

	Public Sub ProcessQueue()
		Dim X As Int32
		Dim oParent As Epica_Entity
		Dim oEntity As Epica_Entity
		Dim oPDef As Epica_Entity_Def
		Dim iTemp As Int16

		Dim lQUB As Int32 = -1

		If mlQueueStart Is Nothing = False Then
			lQUB = Math.Min(mlQueueUB, mlQueueStart.GetUpperBound(0))
		End If

		Try
			For X = 0 To lQUB
				If mlQueueStart(X) <> Int32.MinValue AndAlso mlQueueStart(X) <= glCurrentCycle Then
					'ok got a queue item to execute...
					Dim uQueueItem As QueueItem = moQueue(X)
					With uQueueItem
                        Select Case .lEventType
                            Case QueueItemType.eProcessNextRouteItem
                                Dim lUnitID As Int32 = .lObjectID
                                mlQueueStart(X) = Int32.MinValue
                                Dim oUnit As Unit = GetEpicaUnit(lUnitID)
                                If oUnit Is Nothing = False Then
                                    oUnit.ProcessNextRouteItem()
                                End If
                            Case QueueItemType.eSaveObject
                                Dim lObjID As Int32 = .lObjectID
                                Dim iObjTypeID As Int16 = CShort(.lObjTypeID)
                                If GetAndSaveEpicaObject(lObjID, iObjTypeID) = False Then
                                    'ok what do we do... need to save again? so delay me
                                    mlQueueStart(X) = glCurrentCycle + 5
                                Else : mlQueueStart(X) = Int32.MinValue
                                End If
                            Case QueueItemType.eHandleBeginMining_QIT
                                'ok... begin mining... if we can...
                                Dim oCache As MineralCache = GetEpicaMineralCache(.lTargetID)
                                oEntity = Nothing
                                oEntity = CType(GetEpicaObject(.lObjectID, CShort(.lObjTypeID)), Epica_Entity)
                                If oEntity Is Nothing = False Then
                                    If oCache Is Nothing = False Then
                                        If oCache.BeingMinedBy Is Nothing Then
                                            If oEntity.yProductionType = ProductionType.eMining OrElse oCache.CacheTypeID <> MineralCacheType.eMineable Then
                                                oCache.BeingMinedBy = oEntity
                                                oEntity.CurrentStatus = oEntity.CurrentStatus Or elUnitStatus.eMoveLock
                                                oEntity.bMining = True

                                                Dim yMsg(11) As Byte
                                                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityStatus).CopyTo(yMsg, 0)
                                                oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
                                                System.BitConverter.GetBytes(elUnitStatus.eMoveLock).CopyTo(yMsg, 8)

                                                iTemp = CType(oEntity.ParentObject, Epica_GUID).ObjTypeID
                                                If iTemp = ObjectType.ePlanet Then
                                                    CType(oEntity.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
                                                ElseIf iTemp = ObjectType.eSolarSystem Then
                                                    CType(oEntity.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
                                                End If
                                            Else : mlQueueStart(X) = Int32.MinValue
                                            End If
                                        Else
                                            mlQueueStart(X) = glCurrentCycle + 5    'check every 5 cycles (150ms)
                                        End If
                                    Else
                                        mlQueueStart(X) = Int32.MinValue
                                    End If
                                Else : mlQueueStart(X) = Int32.MinValue
                                End If
                            Case QueueItemType.eHandleDockRequest_QIT, QueueItemType.eHandleUndockRequest_QIT, QueueItemType.eUndockAndReturnToMine_QIT, QueueItemType.eUndockAndReturnToRefinery_QIT, QueueItemType.eUndock_To_Reinforce, QueueItemType.eAILaunchAll
                                'update our event now...
                                mlQueueStart(X) = Int32.MaxValue

                                'LogEvent(LogEventType.ExtensiveLogging, "Processing Dock/Undock for " & .lObjectID & ", " & .lObjTypeID & " to/from " & .lTargetID & ", " & .lTargetTypeID & ". EventType: " & .lEventType)

                                If .lTargetTypeID = ObjectType.eFacility OrElse .lTargetTypeID = ObjectType.eUnit Then
                                    oParent = Nothing : oParent = CType(GetEpicaObject(.lTargetID, CShort(.lTargetTypeID)), Epica_Entity)
                                    If oParent Is Nothing = False Then
                                        If .lEventType = QueueItemType.eAILaunchAll AndAlso oParent.ParentObject Is Nothing = False Then
                                            Dim lTmpParentID As Int32 = CType(oParent.ParentObject, Epica_GUID).ObjectID
                                            Dim iTmpParentTypeID As Int16 = CType(oParent.ParentObject, Epica_GUID).ObjTypeID
                                            Dim lCPUsage As Int32 = oParent.Owner.oBudget.GetEnvirCPUsage(lTmpParentID, iTmpParentTypeID)
                                            If lCPUsage > oParent.Owner.oSpecials.iCPLimit Then
                                                'Ok, cannot launch
                                                mlQueueStart(X) = Int32.MinValue
                                                CancelDockingQueueItem(.lTargetID, .lTargetTypeID)
                                                AddToQueue(glCurrentCycle + 1800, QueueItemType.eResetAILaunchAll, .lTargetID, .lTargetTypeID, 0, 0, 0, 0, 0, 0)
                                                Continue For
                                            End If
                                        End If

                                        oPDef = Nothing
                                        If oParent.ObjTypeID = ObjectType.eUnit Then
                                            oPDef = CType(oParent, Unit).EntityDef
                                        Else : oPDef = CType(oParent, Facility).EntityDef
                                        End If
                                        If oPDef Is Nothing = False Then

                                            oEntity = Nothing : oEntity = CType(GetEpicaObject(.lObjectID, CShort(.lObjTypeID)), Epica_Entity)
                                            If oEntity Is Nothing = False Then
                                                'ok, check the entity's ai settings
                                                If oParent.yProductionType <> ProductionType.eMining AndAlso .lEventType <> QueueItemType.eHandleDockRequest_QIT AndAlso (oEntity.iCombatTactics And eiBehaviorPatterns.eDockDuringBattle) <> 0 Then
                                                    'ok, get the budget item for the parent of the parent
                                                    Dim oPP As Epica_GUID = CType(oParent.ParentObject, Epica_GUID)
                                                    If oPP Is Nothing = False Then
                                                        If oEntity.Owner.oBudget.EnvironmentInConflict(oPP.ObjectID, oPP.ObjTypeID) = True Then
                                                            mlQueueStart(X) = glCurrentCycle + 1800     'every minute...
                                                            Continue For
                                                        End If
                                                    End If
                                                End If
                                                Dim oED As Epica_Entity_Def
                                                If oEntity.ObjTypeID = ObjectType.eUnit Then
                                                    oED = CType(oEntity, Unit).EntityDef
                                                Else : oED = CType(oEntity, Facility).EntityDef
                                                End If

                                                Dim lResult As Int32 = CanDoDockOp(oParent, oEntity, oPDef, oED, X, .lEventType)
                                                If lResult <> -3 Then
                                                    If lResult = -2 Then
                                                        Dim yFailMsg(22) As Byte
                                                        System.BitConverter.GetBytes(GlobalMessageCode.eRequestDockFail).CopyTo(yFailMsg, 0) '2

                                                        oEntity.GetGUIDAsString.CopyTo(yFailMsg, 2) '6
                                                        CType(oEntity.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yFailMsg, 8) '6
                                                        System.BitConverter.GetBytes(oEntity.LocX).CopyTo(yFailMsg, 14) '4
                                                        System.BitConverter.GetBytes(oEntity.LocZ).CopyTo(yFailMsg, 18) '4
                                                        yFailMsg(22) = DockRejectType.eDoorSizeExceeded
                                                        oEntity.Owner.SendPlayerMessage(yFailMsg, True, AliasingRights.eDockUndockUnits)
                                                    ElseIf lResult <> -1 Then
                                                        If lResult = Int32.MaxValue Then
                                                            LogEvent(LogEventType.Warning, "ProcessQueue.AttemptDockCommand. Result is MaxValue. Using +30 instead.")
                                                            lResult = glCurrentCycle + 30
                                                            lResult = Int32.MinValue
                                                        End If
                                                        mlQueueStart(X) = lResult   'glCurrentCycle + lResult
                                                    Else
                                                        'Launch it... or dock it
                                                        If .lEventType = QueueItemType.eHandleDockRequest_QIT Then
                                                            SendDockCommand(.lObjectID, CShort(.lObjTypeID), oParent)
                                                        Else
                                                            'yReturn = 0, 1 or 2. 
                                                            '0 if the operation fails and needs to try again
                                                            '1 if the operation succeeds
                                                            '2 if the operation fails and needs to be removed
                                                            oEntity.uUndockQueueItem = uQueueItem
                                                            Dim yReturn As Byte = LaunchEntity(.lObjectID, CShort(.lObjTypeID), oParent)
                                                            If yReturn = 0 Then
                                                                mlQueueStart(X) = glCurrentCycle
                                                                'ElseIf yReturn = 1 Then
                                                                '    If .lEventType = QueueItemType.eUndockAndReturnToMine_QIT Then
                                                                '        'If oEntity.ObjTypeID = ObjectType.eUnit Then CType(oEntity, Unit).HandleReturnToMiningSite()
                                                                '        If oEntity.ObjTypeID = ObjectType.eUnit Then CType(oEntity, Unit).ProcessNextRouteItem()
                                                                '    ElseIf .lEventType = QueueItemType.eUndockAndReturnToRefinery_QIT Then
                                                                '        'If oEntity.ObjTypeID = ObjectType.eUnit Then CType(oEntity, Unit).HandleReturnToRefinery()
                                                                '        If oEntity.ObjTypeID = ObjectType.eUnit Then CType(oEntity, Unit).ProcessNextRouteItem()
                                                                '    ElseIf .lEventType = QueueItemType.eUndock_To_Reinforce Then
                                                                '        If oEntity.ObjTypeID = ObjectType.eUnit Then
                                                                '            Dim lFleetID As Int32 = CType(oEntity, Unit).lFleetID
                                                                '            If lFleetID > 0 Then
                                                                '                Dim oGroup As UnitGroup = GetEpicaUnitGroup(lFleetID)
                                                                '                If oGroup Is Nothing = False Then oGroup.FinalizeLaunchToReinforce(oEntity, .lExtended1, .lExtended2, .lExtended3, .lExtended4)
                                                                '            End If
                                                                '        End If
                                                                '    Else
                                                                '        'Send the entity to its first waypoint away from its parent
                                                                '        oParent.SendUndockFirstWaypoint(oEntity)
                                                                '    End If
                                                            End If
                                                        End If

                                                        If .lEventType = QueueItemType.eAILaunchAll AndAlso .lExtended1 = 1 Then
                                                            AddToQueue(glCurrentCycle + 1800, QueueItemType.eResetAILaunchAll, .lTargetID, .lTargetTypeID, 0, 0, 0, 0, 0, 0)
                                                        End If
                                                    End If
                                                End If

                                            ElseIf .lEventType = QueueItemType.eHandleUndockRequest_QIT AndAlso .lObjectID = -1 AndAlso .lObjTypeID = -1 Then
                                                Dim lResult As Int32 = ProcessLaunchAll(oParent)
                                                If lResult <> -1 Then
                                                    mlQueueStart(X) = lResult   ' glCurrentCycle + lResult
                                                End If
                                            End If
                                        Else
                                            LogEvent(LogEventType.Warning, "ProcessQueue.Dock/Undock: Entity is not found. " & .lObjectID & ", " & .lObjTypeID)
                                        End If
                                    Else
                                        LogEvent(LogEventType.Warning, "ProcessQueue.Dock/Undock: Parent is not found. " & .lTargetID & ", " & .lTargetTypeID)
                                        mlQueueStart(X) = Int32.MinValue
                                    End If
                                End If
                                If mlQueueStart(X) = Int32.MaxValue Then mlQueueStart(X) = Int32.MinValue
                            Case QueueItemType.eRetryRouteItem
                                Dim lObjID As Int32 = .lObjectID
                                mlQueueStart(X) = Int32.MinValue

                                'Ok, get the unit
                                Dim oUnit As Unit = GetEpicaUnit(lObjID)
                                If oUnit Is Nothing = False Then
                                    oUnit.lRetryCount += 1
                                    If oUnit.lRetryCount > 20 Then
                                        oUnit.lRetryCount = 0
                                    Else
                                        oUnit.CurrentRouteItemAction()
                                    End If
                                End If
                            Case QueueItemType.eReloadRequest
                                'Ok, an entity has finished reloading a weapon...
                                mlQueueStart(X) = Int32.MinValue

                                oEntity = CType(GetEpicaObject(.lObjectID, CShort(.lObjTypeID)), Epica_Entity)
                                If oEntity Is Nothing = False Then
                                    iTemp = CType(oEntity.ParentObject, Epica_GUID).ObjTypeID
                                    If iTemp = ObjectType.ePlanet Then
                                        CType(oEntity.ParentObject, Planet).oDomain.DomainSocket.SendData(GetReloadMsg(oEntity, .lTargetID, .lTargetTypeID))
                                    ElseIf iTemp = ObjectType.eSolarSystem Then
                                        CType(oEntity.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(GetReloadMsg(oEntity, .lTargetID, .lTargetTypeID))
                                    End If
                                End If
                            Case QueueItemType.eCheckColonyResearchQueue
                                Dim lColonyID As Int32 = .lObjectID

                                Dim lNextQueue As Int32 = Int32.MinValue
                                Dim oColony As Colony = GetEpicaColony(lColonyID)
                                If oColony Is Nothing = False Then
                                    Dim lResult As Int32 = oColony.ProcessColonyProdQueues()
                                    If lResult > 0 Then
                                        lNextQueue = glCurrentCycle + lResult
                                    End If
                                End If
                                oColony = Nothing
                                mlQueueStart(X) = lNextQueue
                            Case QueueItemType.eSystemToSystemMove
                                'Ok, the system to system move is over
                                mlQueueStart(X) = Int32.MaxValue

                                Dim oUnitGroup As UnitGroup = GetEpicaUnitGroup(.lObjectID)
                                If oUnitGroup Is Nothing = False Then
                                    Dim lResult As Int32 = oUnitGroup.TestUnitGroupArrival(.lTargetID)
                                    If lResult > 0 Then
                                        mlQueueStart(X) = glCurrentCycle + lResult
                                    End If
                                End If
                                If mlQueueStart(X) = Int32.MaxValue Then mlQueueStart(X) = Int32.MinValue
                            Case QueueItemType.eTradeEventHalfTime
                                If .lTargetID = 1 Then
                                    goGTC.TradeHalfTime(.lObjectID)
                                Else
                                    Dim oTrade As DirectTrade = goGTC.GetDirectTrade(.lObjectID)
                                    If oTrade Is Nothing = False Then oTrade.HandleHalfTime()
                                End If
                                mlQueueStart(X) = Int32.MinValue
                            Case QueueItemType.eTradeEventFinal
                                If .lTargetID = 1 Then
                                    goGTC.TradeFinalTime(.lObjectID)
                                Else
                                    Dim oTrade As DirectTrade = goGTC.GetDirectTrade(.lObjectID)
                                    If oTrade Is Nothing = False Then oTrade.FinalizeTrade()
                                End If
                                mlQueueStart(X) = Int32.MinValue
                            Case QueueItemType.ePirateSelfDestruct
                                mlQueueStart(X) = Int32.MinValue

                                If .lObjTypeID = ObjectType.eFacility Then
                                    'Ok, a pirate base is self-destructing
                                    For Y As Int32 = 0 To glFacilityUB
                                        If glFacilityIdx(Y) <> -1 Then
                                            Dim oFac As Facility = goFacility(Y)
                                            If oFac Is Nothing = False Then
                                                If oFac.Owner.ObjectID = .lTargetID AndAlso oFac.lPirate_For_PlayerID = .lObjectID Then
                                                    glFacilityIdx(Y) = -1
                                                    'kill me
                                                    With oFac
                                                        RemoveLookupFacility(.ObjectID, .ObjTypeID)
                                                        iTemp = CType(.ParentObject, Epica_GUID).ObjTypeID
                                                        Dim yMsg(8) As Byte
                                                        System.BitConverter.GetBytes(GlobalMessageCode.eRemoveObject).CopyTo(yMsg, 0)
                                                        .GetGUIDAsString.CopyTo(yMsg, 2)
                                                        yMsg(8) = 0
                                                        If iTemp = ObjectType.ePlanet Then
                                                            CType(.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
                                                        ElseIf iTemp = ObjectType.eSolarSystem Then
                                                            CType(.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
                                                        End If

                                                        .DeleteEntity(Y)
                                                    End With
                                                End If
                                            End If
                                        End If
                                    Next Y
                                Else
                                    'ok, pirate units are self-destructing
                                    For Y As Int32 = 0 To glUnitUB
                                        If glUnitIdx(Y) <> -1 Then
                                            Dim oUnit As Unit = goUnit(Y)
                                            If oUnit Is Nothing = False AndAlso oUnit.Owner.ObjectID = .lTargetID AndAlso oUnit.lPirate_For_PlayerID = .lObjectID Then
                                                'kill me
                                                With oUnit
                                                    iTemp = CType(.ParentObject, Epica_GUID).ObjTypeID
                                                    Dim yMsg(8) As Byte
                                                    System.BitConverter.GetBytes(GlobalMessageCode.eRemoveObject).CopyTo(yMsg, 0)
                                                    .GetGUIDAsString.CopyTo(yMsg, 2)
                                                    yMsg(8) = 0
                                                    If iTemp = ObjectType.ePlanet Then
                                                        CType(.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
                                                    ElseIf iTemp = ObjectType.eSolarSystem Then
                                                        CType(.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
                                                    End If

                                                    .DeleteEntity(Y)
                                                End With
                                                glUnitIdx(Y) = -1
                                            End If
                                        End If
                                    Next Y

                                    Dim oTmpPlayer As Player = GetEpicaPlayer(.lObjectID)
                                    If oTmpPlayer Is Nothing = False Then
                                        oTmpPlayer.bInPirateSpawn = False
                                    End If


                                    'TODO: Unremark this when Alpha is over
                                    ''Now, test all of the players on this planet to ensure that all players starting here have spawned their pirates
                                    'Dim bLockPlanet As Boolean = True
                                    'For Y As Int32 = 0 To glPlayerUB
                                    '    If glPlayerIdx(Y) <> 0 AndAlso goPlayer(Y).lStartedEnvirID = .lTargetTypeID Then
                                    '        If goPlayer(Y).PirateStartLocX = Int32.MinValue OrElse goPlayer(Y).PirateStartLocZ = Int32.MinValue Then
                                    '            'Ok, not done
                                    '            bLockPlanet = False
                                    '            Exit For
                                    '        End If
                                    '    End If
                                    'Next Y
                                    ''Ok, if we are here, then ensure that no units or facilities exist for the planet for the pirate
                                    'If bLockPlanet = True Then
                                    '    For Y As Int32 = 0 To glUnitUB
                                    '        If glUnitIdx(Y) <> -1 AndAlso goUnit(Y).Owner.ObjectID = .lTargetID Then
                                    '            If CType(goUnit(Y).ParentObject, Epica_GUID).ObjectID = .lTargetTypeID AndAlso CType(goUnit(Y).ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                                    '                bLockPlanet = False
                                    '                Exit For
                                    '            End If
                                    '        End If
                                    '    Next Y
                                    'End If
                                    'If bLockPlanet = True Then
                                    '    For Y As Int32 = 0 To glFacilityUB
                                    '        If glFacilityIdx(Y) <> -1 AndAlso goFacility(Y).Owner.ObjectID = .lTargetID Then
                                    '            If CType(goFacility(Y).ParentObject, Epica_GUID).ObjectID = .lTargetTypeID AndAlso CType(goFacility(Y).ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                                    '                bLockPlanet = False
                                    '                Exit For
                                    '            End If
                                    '        End If
                                    '    Next Y
                                    'End If

                                    'If bLockPlanet = True Then
                                    '    Dim oPlanet As Planet = GetEpicaPlanet(.lTargetTypeID)
                                    '    If oPlanet Is Nothing = False Then oPlanet.SpawnLocked = True
                                    'End If
                                End If
                            Case QueueItemType.eLaunchToAttack
                                'TODO: Launch the Entity referenced by ObjectID, ObjectTypeID by looking up its Parent Object
                                '  Ensure that the parent object is a launchable item
                                'After successful Launch, send message to domain of TARGETGUID (which should be the new parent object)
                                '  for eSetCounterAttack passing the necessary parameters (Type 1 for specific unit list)

                                'Units and facilities launch to the max cp - set AI setting to Engage - set target loc (the SetCounterTarget)

                                oEntity = Nothing
                                If .lObjTypeID = ObjectType.eFacility Then
                                    oEntity = GetEpicaFacility(.lObjectID)
                                ElseIf .lObjTypeID = ObjectType.eUnit Then
                                    oEntity = GetEpicaUnit(.lObjectID)
                                End If
                                If oEntity Is Nothing = False AndAlso (oEntity.CurrentStatus And elUnitStatus.eFacilityPowered) <> 0 Then
                                    For lHngr As Int32 = 0 To oEntity.lHangarUB
                                        If oEntity.lHangarIdx(lHngr) > -1 Then
                                            If oEntity.oHangarContents(lHngr).ObjTypeID = ObjectType.eUnit Then
                                                Dim oUnit As Unit = CType(oEntity.oHangarContents(lHngr), Unit)
                                                oUnit.bLaunchedForCounterAttack = True

                                            End If
                                        End If
                                    Next lHngr
                                End If

                                'obj2id is the envirid is the environment that is having the issues (which should be my resulting environment after launch)
                                'obj2typeid is the envirtypeid
                                'ext1 is the locx
                                'ext2 is the locz

                                mlQueueStart(X) = Int32.MinValue
                            Case QueueItemType.eReinforcements_Built
                                'objectguid is entity producing reinforcements
                                oEntity = Nothing
                                If .lObjTypeID = ObjectType.eUnit Then
                                    oEntity = GetEpicaUnit(.lObjectID)
                                ElseIf .lObjTypeID = ObjectType.eFacility Then
                                    oEntity = GetEpicaFacility(.lObjectID)
                                End If
                                'targetID1 is unitgroupid to reinforce
                                Dim oUnitGroup As UnitGroup = GetEpicaUnitGroup(.lTargetID)

                                Dim bFound As Boolean = False
                                If oEntity Is Nothing = False AndAlso oUnitGroup Is Nothing = False Then
                                    If oUnitGroup.GetElementCount < oUnitGroup.oOwner.oSpecials.yMaxBattleGroupUnits Then
                                        'check entity's hangar
                                        For lHIdx As Int32 = 0 To oEntity.lHangarUB
                                            If oEntity.lHangarIdx(lHIdx) <> -1 AndAlso oEntity.oHangarContents(lHIdx).ObjTypeID = ObjectType.eUnit Then
                                                'targetID2 is unitdefid being built
                                                Dim oUnit As Unit = CType(oEntity.oHangarContents(lHIdx), Unit)
                                                If .lTargetTypeID = oUnit.EntityDef.ObjectID AndAlso oUnit.lFleetID < 1 Then
                                                    'TODO: Remove this unnecessary loop
                                                    Dim lUnitIdx As Int32 = -1
                                                    For Z As Int32 = 0 To glUnitUB
                                                        If glUnitIdx(Z) = oUnit.ObjectID Then
                                                            lUnitIdx = Z
                                                            Exit For
                                                        End If
                                                    Next Z
                                                    If lUnitIdx <> -1 Then
                                                        oUnitGroup.AddUnit(lUnitIdx, True)
                                                        'AddToQueueExtended(glCurrentCycle, QueueItemType.eUndock_To_Reinforce, oUnit.ObjectID, oUnit.ObjTypeID, oEntity.ObjectID, oEntity.ObjTypeID, .lExtended1, .lExtended2, .lExtended3, .lExtended4)
                                                        AddToQueue(glCurrentCycle, QueueItemType.eUndock_To_Reinforce, oUnit.ObjectID, oUnit.ObjTypeID, oEntity.ObjectID, oEntity.ObjTypeID, .lExtended1, .lExtended2, .lExtended3, .lExtended4)
                                                        bFound = True
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        Next lHIdx
                                    End If
                                End If

                                If bFound = True Then
                                    mlQueueStart(X) = Int32.MinValue
                                Else : mlQueueStart(X) = glCurrentCycle + 30
                                End If
                            Case QueueItemType.eGenerateAgent
                                'ok, generate an agent, the ObjectID is the owner
                                Dim oPlayer As Player = GetEpicaPlayer(.lObjectID)
                                If oPlayer Is Nothing = False Then
                                    Dim oAgent As Agent = Agent.GenerateNewAgent(oPlayer)
                                    If oAgent Is Nothing = False Then
                                        If oPlayer.lConnectedPrimaryID > -1 OrElse oPlayer.HasOnlineAliases(AliasingRights.eViewAgents) = True Then
                                            oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oAgent, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                                        End If
                                    End If
                                End If
                                mlQueueStart(X) = Int32.MinValue
                            Case QueueItemType.eGenerateNewbieAgent
                                'ok, generate an agent, the ObjectID is the owner
                                Dim oPlayer As Player = GetEpicaPlayer(.lObjectID)
                                If oPlayer Is Nothing = False Then
                                    Dim oAgent As Agent = Agent.GenerateNewbieAgent(oPlayer)
                                    If oAgent Is Nothing = False Then
                                        If oPlayer.lConnectedPrimaryID > -1 OrElse oPlayer.HasOnlineAliases(AliasingRights.eViewAgents) = True Then
                                            oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oAgent, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                                        End If
                                    End If
                                End If
                                mlQueueStart(X) = Int32.MinValue
                            Case QueueItemType.eRemoveMineralCache, QueueItemType.eRemoveComponentCache
                                Dim lTmpObjID As Int32 = .lObjectID
                                Dim lTmpObjTypeID As Int32 = .lObjTypeID

                                mlQueueStart(X) = Int32.MinValue

                                If lTmpObjTypeID = ObjectType.eMineralCache Then
                                    Dim oCache As MineralCache = GetEpicaMineralCache(lTmpObjID)
                                    If oCache Is Nothing = False Then
                                        If oCache.ParentObject Is Nothing = False Then
                                            iTemp = CType(oCache.ParentObject, Epica_GUID).ObjTypeID
                                            If iTemp = ObjectType.ePlanet OrElse iTemp = ObjectType.eSolarSystem Then
                                                If goGTC Is Nothing = False Then goGTC.AddMineralCacheToNPCTradepost(oCache)
                                                oCache.Quantity = 0
                                            End If
                                        End If
                                    End If
                                Else
                                    Dim oCache As ComponentCache = GetEpicaComponentCache(lTmpObjID)
                                    If oCache Is Nothing = False Then
                                        If oCache.ParentObject Is Nothing = False Then
                                            iTemp = CType(oCache.ParentObject, Epica_GUID).ObjTypeID
                                            If iTemp = ObjectType.ePlanet OrElse iTemp = ObjectType.eSolarSystem Then
                                                If goGTC Is Nothing = False Then goGTC.AddComponentCacheToNPCTradepost(oCache)
                                                oCache.Quantity = 0
                                            End If
                                        End If
                                    End If
                                End If

                            Case QueueItemType.eIronCurtainRaise
                                Dim lPlayerID As Int32 = .lObjectID
                                mlQueueStart(X) = Int32.MinValue
                                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                                If oPlayer Is Nothing = False Then

                                    If oPlayer.AccountStatus <> AccountStatusType.eActiveAccount AndAlso oPlayer.AccountStatus <> AccountStatusType.eMondelisActive AndAlso oPlayer.dtAccountWentInactive <> Date.MinValue Then
                                        If oPlayer.oGuild Is Nothing = False Then
                                            If oPlayer.dtAccountWentInactive.AddDays(7) < Now Then
                                                oPlayer.oGuild.RemoveMember(oPlayer.ObjectID)
                                                oPlayer.lGuildID = -1
                                                oPlayer.lGuildRankID = -1
                                                oPlayer.lInvitedToJoinUB = -1
                                            End If
                                        End If
                                        If oPlayer.dtAccountWentInactive.AddDays(3) < Now Then
                                            oPlayer.ClearFactionSlots()
                                            Continue For
                                        End If
                                    End If

                                    If oPlayer.HasOnlineAliases(0) = False AndAlso (oPlayer.oSocket Is Nothing OrElse oPlayer.oSocket.IsConnected = False) AndAlso oPlayer.lConnectedPrimaryID < 0 Then
                                        If oPlayer.lIronCurtainPlanet < 1 Then oPlayer.lIronCurtainPlanet = oPlayer.lStartedEnvirID

                                        If oPlayer.yIronCurtainState = eIronCurtainState.RaisingIronCurtainOnEverything OrElse (oPlayer.lStatusFlags And elPlayerStatusFlag.AlwaysRaiseFullInvul) <> 0 Then
                                            oPlayer.bInFullLockDown = False
                                            oPlayer.InitiateFullLockdown()
                                            oPlayer.yIronCurtainState = eIronCurtainState.IronCurtainIsUpOnEverything
                                        Else
                                            Dim oPlanet As Planet = GetEpicaPlanet(oPlayer.lIronCurtainPlanet)
                                            If oPlanet Is Nothing = False AndAlso oPlanet.InMyDomain = True Then
                                                'ok, now, make our msg
                                                Dim yMsg(10) As Byte
                                                Dim lPos As Int32 = 0
                                                System.BitConverter.GetBytes(GlobalMessageCode.eSetIronCurtain).CopyTo(yMsg, lPos) : lPos += 2
                                                System.BitConverter.GetBytes(oPlanet.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                                                System.BitConverter.GetBytes(lPlayerID).CopyTo(yMsg, lPos) : lPos += 4
                                                yMsg(lPos) = 1 : lPos += 1
                                                oPlanet.oDomain.DomainSocket.SendData(yMsg)
                                            End If
                                            oPlayer.yIronCurtainState = eIronCurtainState.IronCurtainIsUpOnSelectedPlanet
                                        End If
                                    End If
                                End If

                            Case QueueItemType.eIronCurtainFall
                                Dim lPlayerID As Int32 = .lObjectID
                                Dim bForceful As Boolean = .lExtended4 = Int32.MaxValue
                                mlQueueStart(X) = Int32.MinValue
                                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                                If oPlayer Is Nothing = False Then
                                    If bForceful = True OrElse oPlayer.HasOnlineAliases(0) = True OrElse (oPlayer.oSocket Is Nothing = False AndAlso oPlayer.oSocket.IsConnected = True) Then
                                        oPlayer.lCurrent15MinutesRemaining = 0
                                        oPlayer.bInFullLockDown = False
                                        oPlayer.CancelIronCurtain()
                                        oPlayer.yIronCurtainState = eIronCurtainState.IronCurtainIsDown
                                        If bForceful = False OrElse oPlayer.HasOnlineAliases(0) = True OrElse (oPlayer.oSocket Is Nothing = False AndAlso oPlayer.oSocket.IsConnected = True) Then
                                            If (oPlayer.lStatusFlags And elPlayerStatusFlag.FullInvulnerabilityRaised) <> 0 Then
                                                oPlayer.lStatusFlags = oPlayer.lStatusFlags Xor elPlayerStatusFlag.FullInvulnerabilityRaised
                                            End If
                                        End If
                                    End If
                                End If

                            Case QueueItemType.eBeginRebuilderAI
                                Dim lObjID As Int32 = .lObjectID
                                If lObjID <> -1 Then
                                    Dim oColony As Colony = GetEpicaColony(lObjID)
                                    If oColony Is Nothing = False Then
                                        If oColony.lRebuilderQueueIdx <> -1 Then
                                            'ok, do it... it has not yet been cancelled
                                            oColony.BeginRebuildProcess()
                                        End If
                                    End If
                                End If
                                mlQueueStart(X) = Int32.MinValue
                            Case QueueItemType.eCancelRebuilderAI
                                Dim lObjID As Int32 = .lObjectID
                                If lObjID <> -1 Then
                                    Dim oColony As Colony = GetEpicaColony(lObjID)
                                    If oColony Is Nothing = False Then
                                        oColony.ClearRebuildCommander()
                                    End If
                                End If
                                mlQueueStart(X) = Int32.MinValue
                            Case QueueItemType.eLoadInstancedPlanet
                                Dim lTempID As Int32 = .lObjectID
                                mlQueueStart(X) = Int32.MinValue
                                LoadInstancedPlanet(lTempID)
                            Case QueueItemType.eSaveAndUnloadInstancedPlanet
                                Dim lTempID As Int32 = .lObjectID
                                mlQueueStart(X) = Int32.MinValue
                                SaveAndUnloadInstancedPlanet(lTempID)
                            Case QueueItemType.eResetAILaunchAll
                                Dim lID As Int32 = .lObjectID
                                Dim lObjTypeID As Int32 = .lObjTypeID
                                mlQueueStart(X) = Int32.MinValue

                                oEntity = Nothing
                                oEntity = GetEpicaEntity(lID, CShort(lObjTypeID))
                                If oEntity Is Nothing = False Then
                                    Dim yMsg(7) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.eAILaunchAll).CopyTo(yMsg, 0)
                                    oEntity.GetGUIDAsString.CopyTo(yMsg, 2)

                                    iTemp = CType(oEntity.ParentObject, Epica_GUID).ObjTypeID
                                    If iTemp = ObjectType.ePlanet Then
                                        CType(oEntity.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
                                    ElseIf iTemp = ObjectType.eSolarSystem Then
                                        CType(oEntity.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
                                    End If
                                End If
                            Case QueueItemType.eTutorialFourHourTimerExpire
                                Dim lID As Int32 = .lObjectID
                                mlQueueStart(X) = Int32.MinValue
                                Dim oPlayer As Player = GetEpicaPlayer(lID)
                                If oPlayer Is Nothing = False AndAlso oPlayer.yPlayerPhase = eyPlayerPhase.eSecondPhase Then
                                    Dim lSeconds As Int32 = oPlayer.TotalPlayTime
                                    If oPlayer.lConnectedPrimaryID > -1 Then
                                        lSeconds += CInt(Now.Subtract(oPlayer.LastLogin).TotalSeconds)
                                    End If
                                    If lSeconds >= 14400 Then
                                        'Ok, kill the player
                                        oPlayer.ForceKillMe()
                                    End If
                                End If
                            Case QueueItemType.eSpawnNextPirateWave
                                Dim lPlayerID As Int32 = .lObjectID
                                Dim lPlanetID As Int32 = .lObjTypeID
                                mlQueueStart(X) = Int32.MinValue
                                AureliusAI.SpawnNextWave(lPlayerID, lPlanetID)
                            Case QueueItemType.eReassessMinimumMineralValues
                                If ReassessMinimumMineralValues() = False Then
                                    mlQueueStart(X) = glCurrentCycle + 30
                                Else
                                    mlQueueStart(X) = glCurrentCycle + mlReassessInterval
                                End If
                            Case QueueItemType.eReprocessPlayerRel
                                mlQueueStart(X) = ReprocessPlayerRel(.lObjectID, .lTargetID)
                            Case QueueItemType.ePerformLinkTest
                                Dim lPlayerID As Int32 = .lObjectID
                                Dim lTechID As Int32 = .lObjTypeID
                                mlQueueStart(X) = Int32.MinValue
                                Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                                If oPlayer Is Nothing = False Then
                                    Dim oTech As PlayerSpecialTech = CType(oPlayer.GetTech(lTechID, ObjectType.eSpecialTech), PlayerSpecialTech)
                                    If oTech Is Nothing = False Then
                                        If oTech.bInTheTank = True Then
                                            If oTech.ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                                                oTech.yFlags = oTech.yFlags Or PlayerSpecialTech.eySpecialTechFlags.eTankThrowBackEvent
                                                oPlayer.oSpecials.PerformLinkTest()
                                            End If
                                        End If
                                    End If
                                End If
                            Case QueueItemType.eReprocessFacilityProdQueue
                                Dim lFacID As Int32 = .lObjectID
                                mlQueueStart(X) = Int32.MaxValue

                                Dim oFac As Facility = GetEpicaFacility(lFacID)
                                'Ok, simply check the required resources...
                                If oFac Is Nothing = False AndAlso oFac.ParentColony Is Nothing = False AndAlso oFac.CurrentProduction Is Nothing = False Then
                                    Dim yRes As eResourcesResult = oFac.ParentColony.HasRequiredResources(oFac.CurrentProduction.ProdCost, oFac, eyHasRequiredResourcesFlags.DoNotWait Or eyHasRequiredResourcesFlags.DoNotAlertLowRes Or eyHasRequiredResourcesFlags.NoReduction)
                                    If yRes = eResourcesResult.Sufficient Then
                                        oFac.CurrentProduction.bPaidFor = True
                                        If oFac.GetNextProduction(True) = True Then
                                            oFac.bProducing = True
                                            AddEntityProducing(oFac.ServerIndex, ObjectType.eFacility, oFac.ObjectID)
                                        End If
                                        mlQueueStart(X) = Int32.MinValue
                                    ElseIf yRes = eResourcesResult.Insufficient_Wait Then
                                        mlQueueStart(X) = glCurrentCycle + 1800     '1 minute
                                    Else
                                        mlQueueStart(X) = Int32.MinValue
                                    End If
                                Else
                                    mlQueueStart(X) = Int32.MinValue
                                End If
                            Case Else
                                LogEvent(LogEventType.Warning, "ProcessQueue: Default Case for Queue Type")
                                mlQueueStart(X) = Int32.MinValue
                        End Select
                    End With

                End If
            Next X
        Catch 'ex As Exception
            'LogEvent(LogEventType.ExtensiveLogging, "ProcessQueue: " & ex.Message)
        End Try

        ProcessDeltaQueue()
    End Sub

    'Ok, we manage 2 QueueMgr arrays... one is the "write" array, the other is the "read" array
    '  At the end of ProcessQueue, the "write" array and "read" array are swapped and the "read" array is applied to moQueue
    '  This allows us to ensure queue items are added properly
    Private moDeltaQueue1(-1) As QueueItem
    Private mlDeltaQueue1UB As Int32 = -1
    Private moDeltaQueue2(-1) As QueueItem
    Private mlDeltaQueue2UB As Int32 = -1

    Private mlDeltaQueue1Write As Int32 = 0
    Private mlDeltaQueue2Write As Int32 = 0
    Private myCurrentDeltaQueue As Byte = 0

    Private mlReassessInterval As Int32 = 2592000    '24 hours
    Private mlLastReassessIndex As Int32 = -1
    Private Function ReassessMinimumMineralValues() As Boolean
        Dim lUB As Int32 = -1
        If glFacilityIdx Is Nothing = False Then lUB = Math.Min(glFacilityIdx.GetUpperBound(0), glFacilityUB)
        Dim lCnt As Int32 = 0
        If mlLastReassessIndex = -1 Then mlLastReassessIndex = 0
        For X As Int32 = mlLastReassessIndex To lUB
            If lCnt > 500 Then
                mlLastReassessIndex = X
                Return False
            End If
            If glFacilityIdx(X) > -1 Then
                Dim oFac As Facility = goFacility(X)
                If oFac Is Nothing = False Then
                    If oFac.oMiningBid Is Nothing = False Then
                        'ok, time to reassess....
                        oFac.oMiningBid.ReassessMineralBids()
                        lCnt += 1
                    End If
                End If
            End If
        Next X

        mlLastReassessIndex = -1

        Dim oINI As New InitFile()
        oINI.WriteString("MineralAssess", "LastCalc", GetDateAsNumber(Now).ToString)
        oINI = Nothing
        Return True
    End Function

    Private Function CanDoDockOp(ByRef oParent As Epica_Entity, ByRef oEntity As Epica_Entity, ByVal oPDef As Epica_Entity_Def, ByVal oED As Epica_Entity_Def, ByVal lQIdx As Int32, ByVal lEventType As Int32) As Int32
        Dim bHangarCapGood As Boolean = False

        'verify the hangar cap...
        bHangarCapGood = oParent.Hangar_Cap >= oED.HullSize OrElse (lEventType = QueueItemType.eHandleUndockRequest_QIT OrElse lEventType = QueueItemType.eAILaunchAll OrElse lEventType = QueueItemType.eLaunchToAttack OrElse lEventType = QueueItemType.eUndock_To_Reinforce)

        If oEntity.ObjTypeID = ObjectType.eUnit Then
            If CType(oEntity, Unit).bUnitInSellOrder = True Then Return -3
        End If

        'ok, if our doorsize is not good...
        If oParent.ObjTypeID = ObjectType.eFacility AndAlso oParent.yProductionType <> ProductionType.eMining AndAlso (oParent.CurrentStatus And elUnitStatus.eFacilityPowered) = 0 Then
            LogEvent(LogEventType.PossibleCheat, "Dock/Undock from unpowered parent. Owner: " & oParent.Owner.ObjectID)
            'Target is unpowered
            Dim yFailMsg(22) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestDockFail).CopyTo(yFailMsg, 0) '2
            oEntity.GetGUIDAsString.CopyTo(yFailMsg, 2) '6
            CType(oEntity.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yFailMsg, 8) '6
            System.BitConverter.GetBytes(oEntity.LocX).CopyTo(yFailMsg, 14) '4
            System.BitConverter.GetBytes(oEntity.LocZ).CopyTo(yFailMsg, 18) '4
            yFailMsg(22) = DockRejectType.eHangarInoperable
            oEntity.Owner.SendPlayerMessage(yFailMsg, True, AliasingRights.eDockUndockUnits)
        ElseIf (oParent.CurrentStatus And elUnitStatus.eHangarOperational) = 0 Then
            'Hangar is inoperable, cancel the dock event...
            Dim yFailMsg(22) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestDockFail).CopyTo(yFailMsg, 0) '2
            oEntity.GetGUIDAsString.CopyTo(yFailMsg, 2) '6
            CType(oEntity.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yFailMsg, 8) '6
            System.BitConverter.GetBytes(oEntity.LocX).CopyTo(yFailMsg, 14) '4
            System.BitConverter.GetBytes(oEntity.LocZ).CopyTo(yFailMsg, 18) '4
            yFailMsg(22) = DockRejectType.eHangarInoperable

            oEntity.Owner.SendPlayerMessage(yFailMsg, True, AliasingRights.eDockUndockUnits)
        ElseIf CType(oParent.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem AndAlso (oED.yChassisType And ChassisType.eSpaceBased) = 0 Then
            'cannot undock this because the entity cannot exist in this environment
            Dim yFailMsg(22) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestDockFail).CopyTo(yFailMsg, 0) '2
            oEntity.GetGUIDAsString.CopyTo(yFailMsg, 2) '6
            CType(oEntity.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yFailMsg, 8) '6
            System.BitConverter.GetBytes(oEntity.LocX).CopyTo(yFailMsg, 14) '4
            System.BitConverter.GetBytes(oEntity.LocZ).CopyTo(yFailMsg, 18) '4
            yFailMsg(22) = DockRejectType.eDockerCannotEnterEnvir

            oEntity.Owner.SendPlayerMessage(yFailMsg, True, AliasingRights.eDockUndockUnits)
        ElseIf bHangarCapGood = False Then
            'Hangar cap not good right now... see if it will ever be
            If oPDef.Hangar_Cap >= oED.HullSize Then
                'Yes, it can be... so we'll delay this item
                mlQueueStart(lQIdx) = glCurrentCycle + 5    'check every 5 cycles (150ms)
            Else
                'No, it cannot, so send the response and do not reset timer
                Dim yFailMsg(22) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eRequestDockFail).CopyTo(yFailMsg, 0) '2

                oEntity.GetGUIDAsString.CopyTo(yFailMsg, 2) '6
                CType(oEntity.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yFailMsg, 8) '6
                System.BitConverter.GetBytes(oEntity.LocX).CopyTo(yFailMsg, 14) '4
                System.BitConverter.GetBytes(oEntity.LocZ).CopyTo(yFailMsg, 18) '4
                yFailMsg(22) = DockRejectType.eInsufficientHangarCap

                oEntity.Owner.SendPlayerMessage(yFailMsg, True, AliasingRights.eDockUndockUnits)
            End If
        Else
            'Returns -2 if failed, -1 for success, > 0 for delay timer
            Dim lResult As Int32 = oParent.AttemptDockCommand(oED.HullSize, oED.Maneuver)
            Return lResult
        End If
        Return -3
    End Function

    Public Sub AddToQueue(ByVal lCycle As Int32, ByVal lType As QueueItemType, ByVal lObj1ID As Int32, ByVal lObj1Type As Int32, ByVal lObj2ID As Int32, ByVal lObj2Type As Int32, ByVal lExt1 As Int32, ByVal lExt2 As Int32, ByVal lExt3 As Int32, ByVal lExt4 As Int32)
        If myCurrentDeltaQueue = 0 Then
            mlDeltaQueue1Write += 1

            Dim lIdx As Int32 = -1
            SyncLock moDeltaQueue1
                lIdx = mlDeltaQueue1UB + 1
                If lIdx > moDeltaQueue1.GetUpperBound(0) Then
                    ReDim Preserve moDeltaQueue1(mlDeltaQueue1UB + 1000)
                End If
                mlDeltaQueue1UB += 1
            End SyncLock
            With moDeltaQueue1(lIdx)
                .lEventType = lType
                .lObjectID = lObj1ID
                .lObjTypeID = lObj1Type
                .lTargetID = lObj2ID
                .lTargetTypeID = lObj2Type
                .lExtended1 = lExt1
                .lExtended2 = lExt2
                .lExtended3 = lExt3
                .lExtended4 = lExt4
                .lCycle = lCycle
            End With

            'SyncLock moDeltaQueue1
            '	Try

            '		Dim lIdx As Int32 = mlDeltaQueue1UB + 1
            '		If lIdx > moDeltaQueue1.GetUpperBound(0) Then
            '			ReDim Preserve moDeltaQueue1(mlDeltaQueue1UB + 1000)
            '		End If
            '		mlDeltaQueue1UB += 1

            '		With moDeltaQueue1(lIdx)
            '			.lEventType = lType
            '			.lObjectID = lObj1ID
            '			.lObjTypeID = lObj1Type
            '			.lTargetID = lObj2ID
            '			.lTargetTypeID = lObj2Type
            '			.lExtended1 = lExt1
            '			.lExtended2 = lExt2
            '			.lExtended3 = lExt3
            '			.lExtended4 = lExt4
            '			.lCycle = lCycle
            '		End With
            '	Catch ex As Exception
            '		LogEvent(LogEventType.CriticalError, "AddToQueue: " & ex.Message)
            '	End Try
            'End SyncLock
            mlDeltaQueue1Write -= 1
        Else
            mlDeltaQueue2Write += 1

            Dim lIdx As Int32 = -1
            SyncLock moDeltaQueue2
                lIdx = mlDeltaQueue2UB + 1
                If lIdx > moDeltaQueue2.GetUpperBound(0) Then
                    ReDim Preserve moDeltaQueue2(mlDeltaQueue2UB + 1000)
                End If
                mlDeltaQueue2UB += 1
            End SyncLock
            With moDeltaQueue2(lIdx)
                .lEventType = lType
                .lObjectID = lObj1ID
                .lObjTypeID = lObj1Type
                .lTargetID = lObj2ID
                .lTargetTypeID = lObj2Type
                .lExtended1 = lExt1
                .lExtended2 = lExt2
                .lExtended3 = lExt3
                .lExtended4 = lExt4
                .lCycle = lCycle
            End With
            'SyncLock moDeltaQueue2
            '	Try

            '		Dim lIdx As Int32 = mlDeltaQueue2UB + 1
            '		If lIdx > moDeltaQueue2.GetUpperBound(0) Then
            '			ReDim Preserve moDeltaQueue2(mlDeltaQueue2UB + 1000)
            '		End If
            '		mlDeltaQueue2UB += 1

            '		With moDeltaQueue2(lIdx)
            '			.lEventType = lType
            '			.lObjectID = lObj1ID
            '			.lObjTypeID = lObj1Type
            '			.lTargetID = lObj2ID
            '			.lTargetTypeID = lObj2Type
            '			.lExtended1 = lExt1
            '			.lExtended2 = lExt2
            '			.lExtended3 = lExt3
            '			.lExtended4 = lExt4
            '			.lCycle = lCycle
            '		End With
            '	Catch ex As Exception
            '		LogEvent(LogEventType.CriticalError, "AddToQueue: " & ex.Message)
            '	End Try
            'End SyncLock
            mlDeltaQueue2Write -= 1

        End If
    End Sub
    'DO NOT USE AddToQueueResult IF YOU DO NOT HAVE TO!!!!
    Public Function AddToQueueResult(ByVal lCycle As Int32, ByVal lType As QueueItemType, ByVal lObj1ID As Int32, ByVal lObj1Type As Int32, ByVal lObj2ID As Int32, ByVal lObj2Type As Int32, ByVal lExt1 As Int32, ByVal lExt2 As Int32, ByVal lExt3 As Int32, ByVal lExt4 As Int32) As Int32

        Dim X As Int32
        Dim lIdx As Int32 = -1
        Dim lMaxQueueUB As Int32 = mlQueueUB
        If mlQueueStart Is Nothing = False Then lMaxQueueUB = Math.Min(mlQueueUB, mlQueueStart.GetUpperBound(0))

        Dim bDoSortCheck As Boolean = (lType = QueueItemType.eHandleDockRequest_QIT OrElse lType = QueueItemType.eHandleUndockRequest_QIT OrElse lType = QueueItemType.eUndockAndReturnToMine_QIT OrElse lType = QueueItemType.eUndockAndReturnToRefinery_QIT)

        Try
            For X = 0 To lMaxQueueUB
                If mlQueueStart(X) <> Int32.MinValue Then
                    With moQueue(X)
                        If .lEventType = lType AndAlso .lObjectID = lObj1ID AndAlso .lObjTypeID = lObj1Type AndAlso .lTargetID = lObj2ID AndAlso .lTargetTypeID = lObj2Type Then
                            'ok, same event ?
                            .lExtended1 = lExt1
                            .lExtended2 = lExt2
                            .lExtended3 = lExt3
                            .lExtended4 = lExt4
                            Return X
                        ElseIf bDoSortCheck = True Then
                            If .lEventType = QueueItemType.eHandleDockRequest_QIT OrElse .lEventType = QueueItemType.eHandleUndockRequest_QIT OrElse .lEventType = QueueItemType.eUndockAndReturnToMine_QIT OrElse .lEventType = QueueItemType.eUndockAndReturnToRefinery_QIT Then
                                If .lTargetID = lObj2ID AndAlso .lTargetTypeID = lObj2Type Then
                                    'clear it out so it happens after me
                                    lIdx = -1
                                End If
                            End If
                        End If
                    End With
                ElseIf lIdx = -1 Then
                    lIdx = X
                End If
            Next X
        Catch
            lIdx = -1
        End Try

        Dim lAttempts As Int32 = 0
        Dim bDone As Boolean = False
        While lAttempts < 10 AndAlso bDone = False
            lAttempts += 1
            SyncLock moQueue
                Try
                    If lIdx = -1 Then
                        lIdx = mlQueueUB + 1
                        If lIdx > mlQueueStart.GetUpperBound(0) Then
                            ReDim Preserve moQueue(mlQueueUB + 1000)
                            ReDim Preserve mlQueueStart(mlQueueUB + 1000)
                        End If
                        mlQueueUB += 1
                    End If

                    mlQueueStart(lIdx) = lCycle
                    With moQueue(lIdx)
                        .lEventType = lType
                        .lObjectID = lObj1ID
                        .lObjTypeID = lObj1Type
                        .lTargetID = lObj2ID
                        .lTargetTypeID = lObj2Type
                        .lExtended1 = lExt1
                        .lExtended2 = lExt2
                        .lExtended3 = lExt3
                        .lExtended4 = lExt4
                    End With
                    bDone = True
                Catch ex As Exception
                    LogEvent(LogEventType.CriticalError, "AddToQueue: " & ex.Message)
                    bDone = False
                End Try
            End SyncLock
        End While

        Return lIdx
    End Function
    'Public Function AddToQueue(ByVal lCycle As Int32, ByVal lType As QueueItemType, ByVal lObj1ID As Int32, ByVal lObj1Type As Int32, ByVal lObj2ID As Int32, ByVal lObj2Type As Int32, ByVal lExt1 As Int32, ByVal lExt2 As Int32, ByVal lExt3 As Int32, ByVal lExt4 As Int32) As Int32

    '	Dim X As Int32
    '	Dim lIdx As Int32 = -1
    '	Dim lMaxQueueUB As Int32 = mlQueueUB
    '	If mlQueueStart Is Nothing = False Then lMaxQueueUB = Math.Min(mlQueueUB, mlQueueStart.GetUpperBound(0))

    '	Dim bDoSortCheck As Boolean = (lType = QueueItemType.eHandleDockRequest_QIT OrElse lType = QueueItemType.eHandleUndockRequest_QIT OrElse lType = QueueItemType.eUndockAndReturnToMine_QIT OrElse lType = QueueItemType.eUndockAndReturnToRefinery_QIT)

    '	Try
    '		For X = 0 To lMaxQueueUB
    '			If mlQueueStart(X) <> Int32.MinValue Then
    '				With moQueue(X)
    '					If .lEventType = lType AndAlso .lObjectID = lObj1ID AndAlso .lObjTypeID = lObj1Type AndAlso .lTargetID = lObj2ID AndAlso .lTargetTypeID = lObj2Type Then
    '						'ok, same event ?
    '						.lExtended1 = lExt1
    '						.lExtended2 = lExt2
    '						.lExtended3 = lExt3
    '						.lExtended4 = lExt4
    '						Return X
    '					ElseIf bDoSortCheck = True Then
    '						If .lEventType = QueueItemType.eHandleDockRequest_QIT OrElse .lEventType = QueueItemType.eHandleUndockRequest_QIT OrElse .lEventType = QueueItemType.eUndockAndReturnToMine_QIT OrElse .lEventType = QueueItemType.eUndockAndReturnToRefinery_QIT Then
    '							If .lTargetID = lObj2ID AndAlso .lTargetTypeID = lObj2Type Then
    '								'clear it out so it happens after me
    '								lIdx = -1
    '							End If
    '						End If
    '					End If
    '				End With
    '			ElseIf lIdx = -1 Then
    '				lIdx = X
    '			End If
    '		Next X
    '	Catch
    '		lIdx = -1
    '	End Try

    '	Dim lAttempts As Int32 = 0
    '	Dim bDone As Boolean = False
    '	While lAttempts < 10 AndAlso bDone = False
    '		lAttempts += 1
    '		SyncLock moQueue
    '			Try
    '				If lIdx = -1 Then
    '					lIdx = mlQueueUB + 1
    '					If lIdx > mlQueueStart.GetUpperBound(0) Then
    '						ReDim Preserve moQueue(mlQueueUB + 1000)
    '						ReDim Preserve mlQueueStart(mlQueueUB + 1000)
    '					End If
    '					mlQueueUB += 1
    '				End If

    '				mlQueueStart(lIdx) = lCycle
    '				With moQueue(lIdx)
    '					.lEventType = lType
    '					.lObjectID = lObj1ID
    '					.lObjTypeID = lObj1Type
    '					.lTargetID = lObj2ID
    '					.lTargetTypeID = lObj2Type
    '					.lExtended1 = lExt1
    '					.lExtended2 = lExt2
    '					.lExtended3 = lExt3
    '					.lExtended4 = lExt4
    '				End With
    '				bDone = True
    '			Catch ex As Exception
    '				LogEvent(LogEventType.CriticalError, "AddToQueue: " & ex.Message)
    '				bDone = False
    '			End Try
    '		End SyncLock
    '	End While

    '	Return lIdx
    'End Function
    'Public Function AddToQueueExtended(ByVal lCycle As Int32, ByVal lType As QueueItemType, ByVal lObj1ID As Int32, ByVal lObj1Type As Int32, ByVal lObj2ID As Int32, ByVal lObj2Type As Int32, ByVal lExt1 As Int32, ByVal lExt2 As Int32, ByVal lExt3 As Int32, ByVal lExt4 As Int32) As Int32

    '	Dim X As Int32
    '	Dim lIdx As Int32 = -1
    '	Dim lMaxQueueUB As Int32 = mlQueueUB
    '	If mlQueueStart Is Nothing = False Then lMaxQueueUB = Math.Min(mlQueueUB, mlQueueStart.GetUpperBound(0))

    '	Try
    '		For X = 0 To lMaxQueueUB
    '			If mlQueueStart(X) <> Int32.MinValue Then
    '				With moQueue(X)
    '					If .lEventType = lType AndAlso .lObjectID = lObj1ID AndAlso .lObjTypeID = lObj1Type AndAlso .lTargetID = lObj2ID AndAlso .lTargetTypeID = lObj2Type Then
    '						'ok, same event ?
    '						'If lCycle = mlQueueStart(X) Then
    '						.lExtended1 = lExt1
    '						.lExtended2 = lExt2
    '						.lExtended3 = lExt3
    '						.lExtended4 = lExt4
    '						Return X
    '						'End If
    '					End If
    '				End With
    '			ElseIf lIdx = -1 Then
    '				lIdx = X
    '			End If
    '		Next X
    '	Catch
    '		lIdx = -1
    '	End Try

    '	SyncLock moQueue
    '		Try
    '			If lIdx = -1 Then
    '				mlQueueUB += 1
    '				ReDim Preserve moQueue(mlQueueUB)
    '				ReDim Preserve mlQueueStart(mlQueueUB)
    '				lIdx = mlQueueUB
    '			End If

    '			mlQueueStart(lIdx) = lCycle
    '			With moQueue(lIdx)
    '				.lEventType = lType
    '				.lObjectID = lObj1ID
    '				.lObjTypeID = lObj1Type
    '				.lTargetID = lObj2ID
    '				.lTargetTypeID = lObj2Type
    '				.lExtended1 = lExt1
    '				.lExtended2 = lExt2
    '				.lExtended3 = lExt3
    '				.lExtended4 = lExt4
    '			End With
    '		Catch ex As Exception
    '			LogEvent(LogEventType.CriticalError, "AddToQueue: " & ex.Message)
    '		End Try
    '	End SyncLock

    '	Return lIdx
    'End Function

    Public Sub CancelIronCurtainEvents(ByVal lPlayerID As Int32)
        Dim lQUB As Int32 = -1

        If mlQueueStart Is Nothing = False Then
            lQUB = Math.Min(mlQueueUB, mlQueueStart.GetUpperBound(0))
        End If

        Try
            For X As Int32 = 0 To lQUB
                If mlQueueStart(X) <> Int32.MinValue Then
                    Dim uQueueItem As QueueItem = moQueue(X)
                    If uQueueItem.lEventType = QueueItemType.eIronCurtainRaise Then
                        If uQueueItem.lObjectID = lPlayerID Then
                            mlQueueStart(X) = Int32.MinValue
                        End If
                    ElseIf uQueueItem.lEventType = QueueItemType.eIronCurtainFall Then
                        If uQueueItem.lObjectID = lPlayerID Then
                            Dim lRemainder As Int32 = mlQueueStart(X) - glCurrentCycle
                            mlQueueStart(X) = Int32.MinValue

                            Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
                            If oPlayer Is Nothing = False Then
                                oPlayer.lCurrent15MinutesRemaining = lRemainder
                            End If
                        End If
                    End If
                End If
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "CancelIronCurtainEvent for PlayerID " & lPlayerID & ". Error: " & GetEntireErrorMsg(ex)) 'ex.Message)
        End Try

    End Sub
    Public Sub AlterQueueItem(ByVal lQueueItemIdx As Int32, ByVal lCycle As Int32, ByVal lType As QueueItemType, ByVal lObj1ID As Int32, ByVal iObj1Type As Int16, ByVal lObj2ID As Int32, ByVal iObj2Type As Int16)
        If mlQueueStart.GetUpperBound(0) >= lQueueItemIdx Then
            With moQueue(lQueueItemIdx)
                .lObjTypeID = iObj1Type
                .lObjectID = lObj1ID
                .lTargetTypeID = iObj2Type
                .lTargetID = lObj2ID
                .lEventType = lType
            End With
            mlQueueStart(lQueueItemIdx) = lCycle
        End If
    End Sub
    Public Sub CancelDockingQueueItem(ByVal lObjID As Int32, ByVal lObjTypeID As Int32)
        For X As Int32 = 0 To mlQueueUB
            If mlQueueStart(X) <> Int32.MinValue Then
                If moQueue(X).lObjectID = lObjID AndAlso moQueue(X).lObjTypeID = lObjTypeID OrElse _
                   (moQueue(X).lTargetID = lObjID AndAlso moQueue(X).lTargetTypeID = lObjTypeID AndAlso _
                   moQueue(X).lObjectID = -1 AndAlso moQueue(X).lObjTypeID = -1) Then
                    If moQueue(X).lEventType = QueueItemType.eHandleUndockRequest_QIT OrElse _
                     moQueue(X).lEventType = QueueItemType.eHandleDockRequest_QIT OrElse _
                     moQueue(X).lEventType = QueueItemType.eUndockAndReturnToMine_QIT OrElse _
                     moQueue(X).lEventType = QueueItemType.eUndockAndReturnToRefinery_QIT OrElse _
                     moQueue(X).lEventType = QueueItemType.eAILaunchAll Then

                        mlQueueStart(X) = Int32.MinValue
                    End If
                End If
            End If
        Next X
    End Sub
    Public Function GetQueueDelay(ByVal lObjID As Int32, ByVal iObjTypeID As Int16) As Int32
        For X As Int32 = 0 To mlQueueUB
            If mlQueueStart(X) <> Int32.MinValue Then
                If moQueue(X).lObjectID = lObjID AndAlso moQueue(X).lObjTypeID = iObjTypeID Then
                    Return mlQueueStart(X) - glCurrentCycle
                ElseIf moQueue(X).lTargetID = lObjID AndAlso moQueue(X).lTargetTypeID = iObjTypeID Then
                    If moQueue(X).lObjectID = -1 AndAlso moQueue(X).lObjTypeID = -1 Then Return mlQueueStart(X) - glCurrentCycle
                End If
            End If
        Next X
        Return -1
    End Function

    Private Sub ProcessDeltaQueue()
        'ok, clear our non delta queue data and then change our current delta to the other
        If myCurrentDeltaQueue = 0 Then
            mlDeltaQueue2UB = -1
            mlDeltaQueue2Write = 0
            myCurrentDeltaQueue = 1

            'Now, wait for our counters
            Dim lWaits As Int32 = 0
            While mlDeltaQueue1Write <> 0
                Threading.Thread.Sleep(1)
                lWaits += 1
                If lWaits > 10 Then Exit While
            End While
        Else
            mlDeltaQueue1UB = -1
            mlDeltaQueue1Write = 0
            myCurrentDeltaQueue = 0

            'Now, wait for our counters
            Dim lWaits As Int32 = 0
            While mlDeltaQueue2Write <> 0
                Threading.Thread.Sleep(1)
                lWaits += 1
                If lWaits > 10 Then Exit While
            End While
        End If

        'Ok, our delta queue is now the other delta, so we can work with the non-delta...
        If myCurrentDeltaQueue = 0 Then
            'working with queue 2
            For X As Int32 = 0 To mlDeltaQueue2UB
                With moDeltaQueue2(X)
                    DeltaQueueAddToQueue(.lCycle, .lEventType, .lObjectID, .lObjTypeID, .lTargetID, .lTargetTypeID, .lExtended1, .lExtended2, .lExtended3, .lExtended4)
                End With
            Next X
        Else
            'working with queue 1
            For X As Int32 = 0 To mlDeltaQueue1UB
                With moDeltaQueue1(X)
                    DeltaQueueAddToQueue(.lCycle, .lEventType, .lObjectID, .lObjTypeID, .lTargetID, .lTargetTypeID, .lExtended1, .lExtended2, .lExtended3, .lExtended4)
                End With
            Next X
        End If
    End Sub

    Private Sub DeltaQueueAddToQueue(ByVal lCycle As Int32, ByVal lType As Int32, ByVal lObj1ID As Int32, ByVal lObj1Type As Int32, ByVal lObj2ID As Int32, ByVal lObj2Type As Int32, ByVal lExt1 As Int32, ByVal lExt2 As Int32, ByVal lExt3 As Int32, ByVal lExt4 As Int32)

        Dim X As Int32
        Dim lIdx As Int32 = -1
        Dim lMaxQueueUB As Int32 = mlQueueUB
        If mlQueueStart Is Nothing = False Then lMaxQueueUB = Math.Min(mlQueueUB, mlQueueStart.GetUpperBound(0))

        Dim bDoSortCheck As Boolean = (lType = QueueItemType.eHandleDockRequest_QIT OrElse lType = QueueItemType.eHandleUndockRequest_QIT OrElse lType = QueueItemType.eUndockAndReturnToMine_QIT OrElse lType = QueueItemType.eUndockAndReturnToRefinery_QIT)

        Try
            For X = 0 To lMaxQueueUB
                If mlQueueStart(X) <> Int32.MinValue Then
                    With moQueue(X)
                        If .lEventType = lType AndAlso .lObjectID = lObj1ID AndAlso .lObjTypeID = lObj1Type AndAlso .lTargetID = lObj2ID AndAlso .lTargetTypeID = lObj2Type Then
                            'ok, same event ?
                            If lType = QueueItemType.eReprocessPlayerRel Then
                                .lCycle = Math.Min(lCycle, .lCycle)
                                mlQueueStart(X) = .lCycle
                            End If
                            .lExtended1 = lExt1
                            .lExtended2 = lExt2
                            .lExtended3 = lExt3
                            .lExtended4 = lExt4
                            Return
                        ElseIf bDoSortCheck = True Then
                            If .lEventType = QueueItemType.eHandleDockRequest_QIT OrElse .lEventType = QueueItemType.eHandleUndockRequest_QIT OrElse .lEventType = QueueItemType.eUndockAndReturnToMine_QIT OrElse .lEventType = QueueItemType.eUndockAndReturnToRefinery_QIT Then
                                If .lTargetID = lObj2ID AndAlso .lTargetTypeID = lObj2Type Then
                                    'clear it out so it happens after me
                                    lIdx = -1
                                End If
                            End If
                        End If
                    End With
                ElseIf lIdx = -1 Then
                    lIdx = X
                End If
            Next X
        Catch
            lIdx = -1
        End Try

        Dim lAttempts As Int32 = 0
        Dim bDone As Boolean = False
        While lAttempts < 10 AndAlso bDone = False
            lAttempts += 1
            SyncLock moQueue
                Try
                    If lIdx = -1 Then
                        lIdx = mlQueueUB + 1
                        If lIdx > mlQueueStart.GetUpperBound(0) Then
                            ReDim Preserve moQueue(mlQueueUB + 1000)
                            ReDim Preserve mlQueueStart(mlQueueUB + 1000)
                        End If
                        mlQueueUB += 1
                    End If

                    mlQueueStart(lIdx) = lCycle
                    With moQueue(lIdx)
                        .lEventType = lType
                        .lObjectID = lObj1ID
                        .lObjTypeID = lObj1Type
                        .lTargetID = lObj2ID
                        .lTargetTypeID = lObj2Type
                        .lExtended1 = lExt1
                        .lExtended2 = lExt2
                        .lExtended3 = lExt3
                        .lExtended4 = lExt4
                    End With
                    bDone = True
                Catch ex As Exception
                    LogEvent(LogEventType.CriticalError, "AddToQueue: " & ex.Message)
                    bDone = False
                End Try
            End SyncLock
        End While

    End Sub

    Public Function GetQueueSize() As Int32
        Return mlQueueUB
    End Function
    Public Function GetQueueItemCount() As Int32
        Dim lQUB As Int32 = -1

        If mlQueueStart Is Nothing = False Then
            lQUB = Math.Min(mlQueueUB, mlQueueStart.GetUpperBound(0))
        End If
        Dim lCnt As Int32 = 0
        Try
            For X As Int32 = 0 To lQUB
                If mlQueueStart(X) <> Int32.MinValue Then
                    lCnt += 1
                End If
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "GetQueueItemCount: " & ex.Message)
        End Try
        Return lCnt
    End Function

#End Region

    Private Function ReprocessPlayerRel(ByVal lPlayerID As Int32, ByVal lTargetID As Int32) As Int32
        Dim lResult As Int32 = Int32.MinValue

        Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
        Dim oTarget As Player = GetEpicaPlayer(lTargetID)
        If oPlayer Is Nothing = False AndAlso oTarget Is Nothing = False Then
            Dim oRel As PlayerRel = oPlayer.GetPlayerRel(lTargetID)
            Dim oTheirRel As PlayerRel = oTarget.GetPlayerRel(lPlayerID)
            If oRel Is Nothing = False AndAlso oTheirRel Is Nothing = False Then
                'ok, now... are we changing score?
                If oRel.TargetScore < oRel.WithThisScore Then
                    'yes, we are... determine the lower of the two...
                    Dim lNextVal As Int32 = oRel.WithThisScore
                    lNextVal -= 5
                    If lNextVal > oTheirRel.WithThisScore Then
                        lNextVal = oTheirRel.WithThisScore
                    End If

                    If oPlayer.oGuild Is Nothing = False Then
                        For X As Int32 = 0 To oPlayer.oGuild.lMemberUB
                            Dim oMem As GuildMember = oPlayer.oGuild.oMembers(X)
                            If oMem Is Nothing = False AndAlso (oMem.yMemberState And GuildMemberState.Approved) <> 0 Then
                                If oMem.oMember Is Nothing = False Then
                                    Dim yGldRel As Byte = oMem.oMember.GetPlayerRelScore(lTargetID)
                                    If yGldRel <= lNextVal Then
                                        lNextVal = Math.Min(lNextVal, yGldRel)  'Math.Max(oRel.TargetScore, yGldRel)
                                    End If
                                End If
                            End If
                        Next X
                    End If
                    'For X As Int32 = 0 To 4
                    '    Dim lTmp As Int32 = oPlayer.lSlotID(X)
                    '    If lTmp > -1 Then
                    '        Dim oFaction As Player = GetEpicaPlayer(lTmp)
                    '        If oFaction Is Nothing = False Then
                    '            Dim yFacRel As Byte = oFaction.GetPlayerRelScore(lTargetID)
                    '            If yFacRel <= lNextVal Then
                    '                lNextVal = Math.Min(lNextVal, yFacRel)
                    '            End If
                    '        End If
                    '    End If
                    'Next X
                    'For X As Int32 = 0 To 2
                    '    Dim lTmp As Int32 = oPlayer.lFactionID(X)
                    '    If lTmp > -1 Then
                    '        Dim oFaction As Player = GetEpicaPlayer(lTmp)
                    '        If oFaction Is Nothing = False Then
                    '            Dim yFacRel As Byte = oFaction.GetPlayerRelScore(lTargetID)
                    '            If yFacRel <= lNextVal Then
                    '                lNextVal = Math.Min(lNextVal, yFacRel)
                    '            End If
                    '        End If
                    '    End If
                    'Next X

                    If oRel.TargetScore < 1 Then oRel.TargetScore = 1
                    If lNextVal < oRel.TargetScore Then lNextVal = oRel.TargetScore

                    If lNextVal < 1 Then lNextVal = 1
                    If lNextVal > 255 Then lNextVal = 255

                    If gb_IS_TEST_SERVER = True OrElse (oPlayer.yPlayerPhase = eyPlayerPhase.eSecondPhase AndAlso oTarget.yPlayerPhase = eyPlayerPhase.eSecondPhase) Then
                        lResult = glCurrentCycle + 150
                    Else
                        lResult = glCurrentCycle + gl_DELAY_FOUR_HOURS
                    End If

                    oRel.lCycleOfNextScoreUpdate = lResult

                    'Ok, lNextVal is the new score...
                    If oRel.WithThisScore > elRelTypes.eWar AndAlso oTheirRel.WithThisScore > elRelTypes.eWar AndAlso lNextVal <= elRelTypes.eWar Then
                        ProcessWarDeclared(oPlayer, oTarget, lNextVal)
                    Else
                        'Ok, need to set the rels
                        oRel.WithThisScore = CByte(lNextVal)
                        Try
                            oRel.SaveObject()
                            oRel.StorePlayerRelHistory()
                        Catch
                        End Try

                        Dim yMsg() As Byte = goMsgSys.GetAddPlayerRelMessage(oRel)


                        Dim yResp(27) As Byte
                        System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityName).CopyTo(yResp, 0)
                        oTarget.GetGUIDAsString.CopyTo(yResp, 2)
                        oTarget.PlayerName.CopyTo(yResp, 8)
                        oPlayer.SendPlayerMessage(yResp, False, 0)
                        ReDim yResp(27)
                        System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityName).CopyTo(yResp, 0)
                        oPlayer.GetGUIDAsString.CopyTo(yResp, 2)
                        oPlayer.PlayerName.CopyTo(yResp, 8)
                        oTarget.SendPlayerMessage(yResp, False, 0)

                        System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerRel).CopyTo(yMsg, 0)
                        oPlayer.SendPlayerMessage(yMsg, False, AliasingRights.eViewDiplomacy Or AliasingRights.eViewAgents)
                        oTarget.SendPlayerMessage(yMsg, True, AliasingRights.eViewEmail Or AliasingRights.eViewDiplomacy Or AliasingRights.eViewAgents)

                        oPlayer.ReverifySlots()
                        oTarget.ReverifySlots()
                    End If

                    'If lNextVal <> oRel.TargetScore Then

                    'End If

                End If

            End If
        End If

        Return lResult
    End Function
    Public Sub ProcessWarDeclared(ByRef oPlayer As Player, ByRef oTarget As Player, ByVal lNewVal As Int32)
        Dim yPTestTitle As Byte = oPlayer.yPlayerTitle
        Dim yTTestTitle As Byte = oTarget.yPlayerTitle
        If (yPTestTitle And Player.PlayerRank.ExRankShift) <> 0 Then yPTestTitle = yPTestTitle Xor Player.PlayerRank.ExRankShift
        If (yTTestTitle And Player.PlayerRank.ExRankShift) <> 0 Then yTTestTitle = yTTestTitle Xor Player.PlayerRank.ExRankShift
        oTarget.bDeclaredWarOn = yPTestTitle > yTTestTitle

        'Ok, we are declaring war... let's do it cascadingly
        Dim oCascadeWarDec As New CascadeWarDec
        oCascadeWarDec.HandleCascadeWarDec(oPlayer, oTarget, CByte(lNewVal))
        oCascadeWarDec = Nothing

        oPlayer.ReverifySlots()
        oTarget.ReverifySlots()

        'Ok, now, go through the player's guild
        If oPlayer.oGuild Is Nothing = False Then
            For X As Int32 = 0 To oPlayer.oGuild.lMemberUB
                Try
                    If oPlayer.oGuild.oMembers(X) Is Nothing = False AndAlso (oPlayer.oGuild.oMembers(X).yMemberState And GuildMemberState.Approved) <> 0 Then
                        Dim oTmp As Player = oPlayer.oGuild.oMembers(X).oMember
                        If oTmp Is Nothing = False Then
                            Dim oTmpRel As PlayerRel = oTmp.GetPlayerRel(oTarget.ObjectID)
                            If oTmpRel Is Nothing = False Then
                                If oTmpRel.TargetScore <= lNewVal Then
                                    oTmpRel.lCycleOfNextScoreUpdate = glCurrentCycle + 1
                                    AddToQueue(oTmpRel.lCycleOfNextScoreUpdate, QueueItemType.eReprocessPlayerRel, oTmp.ObjectID, oTarget.ObjectID, -1, -1, -1, -1, -1, -1)
                                End If
                            End If
                        End If
                    End If
                Catch
                End Try
            Next X
        End If

        If oPlayer.lFactionID Is Nothing = False Then
            For X As Int32 = 0 To oPlayer.lFactionID.GetUpperBound(0)
                If oPlayer.lFactionID(X) > 0 Then
                    Dim oTmp As Player = GetEpicaPlayer(oPlayer.lFactionID(X))
                    If oTmp Is Nothing = False Then
                        Dim oTmpRel As PlayerRel = oTmp.GetPlayerRel(oTarget.ObjectID)
                        If oTmpRel Is Nothing = False Then
                            If oTmpRel.TargetScore <= lNewVal Then
                                oTmpRel.lCycleOfNextScoreUpdate = glCurrentCycle + 1
                                AddToQueue(oTmpRel.lCycleOfNextScoreUpdate, QueueItemType.eReprocessPlayerRel, oTmp.ObjectID, oTarget.ObjectID, -1, -1, -1, -1, -1, -1)
                            End If
                        End If
                    End If
                End If
            Next X
        End If
        If oPlayer.lSlotID Is Nothing = False Then
            For X As Int32 = 0 To oPlayer.lSlotID.GetUpperBound(0)
                If oPlayer.lSlotID(X) > 0 Then
                    Dim oTmp As Player = GetEpicaPlayer(oPlayer.lSlotID(X))
                    If oTmp Is Nothing = False Then
                        Dim oTmpRel As PlayerRel = oTmp.GetPlayerRel(oTarget.ObjectID)
                        If oTmpRel Is Nothing = False Then
                            If oTmpRel.TargetScore <= lNewVal Then
                                oTmpRel.lCycleOfNextScoreUpdate = glCurrentCycle + 1
                                AddToQueue(oTmpRel.lCycleOfNextScoreUpdate, QueueItemType.eReprocessPlayerRel, oTmp.ObjectID, oTarget.ObjectID, -1, -1, -1, -1, -1, -1)
                            End If
                        End If
                    End If
                End If
            Next X
        End If

        'Ok, now, go through the player's guild
        If oTarget.oGuild Is Nothing = False Then
            For X As Int32 = 0 To oTarget.oGuild.lMemberUB
                Try
                    If oTarget.oGuild.oMembers(X) Is Nothing = False AndAlso (oTarget.oGuild.oMembers(X).yMemberState And GuildMemberState.Approved) <> 0 Then
                        Dim oTmp As Player = oTarget.oGuild.oMembers(X).oMember
                        If oTmp Is Nothing = False Then
                            Dim oTmpRel As PlayerRel = oTmp.GetPlayerRel(oPlayer.ObjectID)
                            If oTmpRel Is Nothing = False Then
                                If oTmpRel.TargetScore <= lNewVal Then
                                    oTmpRel.lCycleOfNextScoreUpdate = glCurrentCycle + 1
                                    AddToQueue(oTmpRel.lCycleOfNextScoreUpdate, QueueItemType.eReprocessPlayerRel, oTmp.ObjectID, oPlayer.ObjectID, -1, -1, -1, -1, -1, -1)
                                End If
                            End If
                        End If
                    End If
                Catch
                End Try
            Next X
        End If

        If oTarget.lFactionID Is Nothing = False Then
            For X As Int32 = 0 To oTarget.lFactionID.GetUpperBound(0)
                If oTarget.lFactionID(X) > 0 Then
                    Dim oTmp As Player = GetEpicaPlayer(oTarget.lFactionID(X))
                    If oTmp Is Nothing = False Then
                        Dim oTmpRel As PlayerRel = oTmp.GetPlayerRel(oPlayer.ObjectID)
                        If oTmpRel Is Nothing = False Then
                            If oTmpRel.TargetScore <= lNewVal Then
                                oTmpRel.lCycleOfNextScoreUpdate = glCurrentCycle + 1
                                AddToQueue(oTmpRel.lCycleOfNextScoreUpdate, QueueItemType.eReprocessPlayerRel, oTmp.ObjectID, oPlayer.ObjectID, -1, -1, -1, -1, -1, -1)
                            End If
                        End If
                    End If
                End If
            Next X
        End If
        If oTarget.lSlotID Is Nothing = False Then
            For X As Int32 = 0 To oTarget.lSlotID.GetUpperBound(0)
                If oTarget.lSlotID(X) > 0 Then
                    Dim oTmp As Player = GetEpicaPlayer(oTarget.lSlotID(X))
                    If oTmp Is Nothing = False Then
                        Dim oTmpRel As PlayerRel = oTmp.GetPlayerRel(oPlayer.ObjectID)
                        If oTmpRel Is Nothing = False Then
                            If oTmpRel.TargetScore <= lNewVal Then
                                oTmpRel.lCycleOfNextScoreUpdate = glCurrentCycle + 1
                                AddToQueue(oTmpRel.lCycleOfNextScoreUpdate, QueueItemType.eReprocessPlayerRel, oTmp.ObjectID, oPlayer.ObjectID, -1, -1, -1, -1, -1, -1)
                            End If
                        End If
                    End If
                End If
            Next X
        End If

        'Ok, now, let's see what happens
        If oPlayer.oGuild Is Nothing = False Then
            For X As Int32 = 0 To oPlayer.oGuild.lMemberUB
                Dim oMem As GuildMember = oPlayer.oGuild.oMembers(X)
                If oMem Is Nothing = False AndAlso (oMem.yMemberState And GuildMemberState.Approved) <> 0 Then
                    If oTarget.oGuild Is Nothing = False Then
                        For Y As Int32 = 0 To oTarget.oGuild.lMemberUB
                            Dim oTMem As GuildMember = oTarget.oGuild.oMembers(Y)
                            If oTMem Is Nothing = False AndAlso (oTMem.yMemberState And GuildMemberState.Approved) <> 0 Then
                                'Dim oTmpRel As PlayerRel = oTMem.oMember.GetPlayerRel(oMem.lMemberID)
                                'If oTmpRel Is Nothing = False AndAlso oTmpRel.TargetScore <= lNextVal Then
                                '    oTmpRel.lCycleOfNextScoreUpdate = glCurrentCycle + 1
                                '    AddToQueue(oTmpRel.lCycleOfNextScoreUpdate, QueueItemType.eReprocessPlayerRel, oTMem.lMemberID, oMem.lMemberID, -1, -1, -1, -1, -1, -1)
                                'End If

                                Dim oNewRel As PlayerRel = oMem.oMember.GetPlayerRel(oTMem.lMemberID)
                                If oNewRel Is Nothing = False AndAlso oNewRel.TargetScore <= lNewVal Then
                                    oNewRel.lCycleOfNextScoreUpdate = glCurrentCycle + 1
                                    AddToQueue(oNewRel.lCycleOfNextScoreUpdate, QueueItemType.eReprocessPlayerRel, oMem.lMemberID, oTMem.lMemberID, -1, -1, -1, -1, -1, -1)
                                End If
                            End If
                        Next Y
                    End If
                    Dim oNother As PlayerRel = oMem.oMember.GetPlayerRel(oTarget.ObjectID)
                    If oNother Is Nothing = False AndAlso oNother.TargetScore <= lNewVal Then
                        oNother.lCycleOfNextScoreUpdate = glCurrentCycle + 1
                        AddToQueue(oNother.lCycleOfNextScoreUpdate, QueueItemType.eReprocessPlayerRel, oMem.lMemberID, oTarget.ObjectID, -1, -1, -1, -1, -1, -1)
                    End If
                End If
            Next X
        End If
        If oTarget.oGuild Is Nothing = False Then
            For X As Int32 = 0 To oTarget.oGuild.lMemberUB
                Dim oMem As GuildMember = oTarget.oGuild.oMembers(X)
                If oMem Is Nothing = False AndAlso (oMem.yMemberState And GuildMemberState.Approved) <> 0 Then
                    If oPlayer.oGuild Is Nothing = False Then
                        For Y As Int32 = 0 To oPlayer.oGuild.lMemberUB
                            Dim oTMem As GuildMember = oPlayer.oGuild.oMembers(Y)
                            If oTMem Is Nothing = False AndAlso (oTMem.yMemberState And GuildMemberState.Approved) <> 0 Then
                                Dim oNewRel As PlayerRel = oMem.oMember.GetPlayerRel(oTMem.lMemberID)
                                If oNewRel Is Nothing = False AndAlso oNewRel.TargetScore <= lNewVal Then
                                    oNewRel.lCycleOfNextScoreUpdate = glCurrentCycle + 1
                                    AddToQueue(oNewRel.lCycleOfNextScoreUpdate, QueueItemType.eReprocessPlayerRel, oMem.lMemberID, oTMem.lMemberID, -1, -1, -1, -1, -1, -1)
                                End If
                            End If
                        Next Y
                    End If
                    Dim oNother As PlayerRel = oMem.oMember.GetPlayerRel(oPlayer.ObjectID)
                    If oNother Is Nothing = False AndAlso oNother.TargetScore <= lNewVal Then
                        oNother.lCycleOfNextScoreUpdate = glCurrentCycle + 1
                        AddToQueue(oNother.lCycleOfNextScoreUpdate, QueueItemType.eReprocessPlayerRel, oMem.lMemberID, oPlayer.ObjectID, -1, -1, -1, -1, -1, -1)
                    End If
                End If
            Next X
        End If
    End Sub

    Private Function GetReloadMsg(ByVal oEntity As Epica_Entity, ByVal lWpnID As Int32, ByVal lAmt As Int32) As Byte()
        Dim yMsg(15) As Byte

        System.BitConverter.GetBytes(GlobalMessageCode.eReloadWpnMsg).CopyTo(yMsg, 0)
        oEntity.GetGUIDAsString.CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(lWpnID).CopyTo(yMsg, 8)
        System.BitConverter.GetBytes(lAmt).CopyTo(yMsg, 12)

        Return yMsg

    End Function

    Private Sub ProductionAndMining()
        Dim X As Int32

        'mlPreviousSentimentCheck = 0

        'Determine if we need to update credits...
        Dim bUpdateCredits As Boolean = False
        If mswCredits.ElapsedMilliseconds > 5000 Then
            Dim dtNow As DateTime = DateTime.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime
            Dim lNow As Int32 = GetDateAsNumber(dtNow)

            For X = 0 To glPlayerUB
                If glPlayerIdx(X) <> -1 AndAlso glPlayerIdx(X) <> gl_HARDCODE_PIRATE_PLAYER_ID Then
                    goPlayer(X).oBudget.Reset()

                    goPlayer(X).ProcessTransportEvents(dtNow)

                    'If mlPreviousSentimentCheck = 0 Then goPlayer(X).lTemporaryWarpointCheckFlag = goPlayer(X).CheckWPUpkeep(lNow, dtNow) Else goPlayer(X).lTemporaryWarpointCheckFlag = elWarpointCheckFlag.eNoCheckRequired
                    'If goPlayer(X).lTemporaryWarpointCheckFlag <> elWarpointCheckFlag.eNoCheckRequired Then
                    '    mlPreviousSentimentCheck = glCurrentCycle
                    'End If
                    'goPlayer(X).lTemporaryWarpointUpkeepCost = 0
                    'goPlayer(X).lCurrentWarpointUpkeepCost = 0

                    If goPlayer(X).InMyDomain = True Then
                        If goPlayer(X).BadWarDecCPIncreaseEndCycle <> 0 AndAlso goPlayer(X).BadWarDecCPIncreaseEndCycle < glCurrentCycle Then
                            'ok, end this player's suffering
                            goPlayer(X).BadWarDecCPIncrease = 0
                            goPlayer(X).BadWarDecCPIncreaseEndCycle = 0
                            goPlayer(X).ClearCPPenaltyList()
                            goPlayer(X).SendCPPenaltyList()
                            'Now, send to our regions...
                            Dim yMsg(11) As Byte
                            System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerSpecialAttribute).CopyTo(yMsg, 0)
                            System.BitConverter.GetBytes(goPlayer(X).ObjectID).CopyTo(yMsg, 2)
                            System.BitConverter.GetBytes(ePlayerSpecialAttributeSetting.eBadWarDecCPPenalty).CopyTo(yMsg, 6)
                            System.BitConverter.GetBytes(goPlayer(X).BadWarDecCPIncrease).CopyTo(yMsg, 8)
                            'goMsgSys.BroadcastToDomains(yMsg)
                            goMsgSys.SendMsgToOperator(yMsg)

                            'goPlayer(X).SendPlayerMessage(yMsg, False, AliasingRights.eViewUnitsAndFacilities)
                            Dim oPC As PlayerComm = goPlayer(X).AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "The command point penalty has been removed.", "Command Penalty End", goPlayer(X).ObjectID, GetDateAsNumber(Now), False, goPlayer(X).sPlayerNameProper, Nothing)
                            If oPC Is Nothing = False Then
                                goPlayer(X).SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                            End If
                        End If
                        If goPlayer(X).BadWarDecMoralePenaltyEndCycle <> 0 AndAlso goPlayer(X).BadWarDecMoralePenaltyEndCycle < glCurrentCycle Then
                            'ok, end this player's suffering
                            goPlayer(X).BadWarDecMoralePenalty = 0
                            goPlayer(X).BadWarDecMoralePenaltyEndCycle = 0
                            Dim yMsg(11) As Byte
                            System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerSpecialAttribute).CopyTo(yMsg, 0)
                            System.BitConverter.GetBytes(goPlayer(X).ObjectID).CopyTo(yMsg, 2)
                            System.BitConverter.GetBytes(ePlayerSpecialAttributeSetting.eBadWarDecMoralePenalty).CopyTo(yMsg, 6)
                            System.BitConverter.GetBytes(goPlayer(X).BadWarDecCPIncrease).CopyTo(yMsg, 8)
                            'goPlayer(X).SendPlayerMessage(yMsg, False, AliasingRights.eViewUnitsAndFacilities)
                            goMsgSys.SendMsgToOperator(yMsg)
                        End If
                    End If

                End If
            Next X

            RandomPirateOp()

            bUpdateCredits = True
            mswCredits.Reset()
            mswCredits.Start()
        End If

        'Now, tell our production manager to do its thing
        CheckProduction()

        If bUpdateCredits = True Then

            'MSC - 01/03/09 - moved here so that HandleColonyTaxIncome could access the system's facility point counts
            'TODO: There is code throughout about SpaceFac lookup in the Player, use that instead of this clumsy fac array lookup
            Dim lCurUB As Int32 = -1
            If glFacilityIdx Is Nothing = False Then lCurUB = Math.Min(glFacilityIdx.GetUpperBound(0), glFacilityUB)
            For X = 0 To lCurUB
                If glFacilityIdx(X) > -1 Then
                    Dim oFac As Facility = goFacility(X)
                    If oFac Is Nothing = False AndAlso oFac.ParentObject Is Nothing = False Then
                        With CType(oFac.ParentObject, Epica_GUID)
                            If .ObjTypeID = ObjectType.eSolarSystem Then
                                If oFac.Owner.ObjectID <> gl_HARDCODE_PIRATE_PLAYER_ID AndAlso oFac.Owner.oBudget Is Nothing = False Then
                                    If oFac.yProductionType = ProductionType.eSpaceStationSpecial Then
                                        oFac.Owner.oBudget.AddFacilityToEnvir(.ObjectID, .ObjTypeID, 50)
                                    Else
                                        oFac.Owner.oBudget.AddFacilityToEnvir(.ObjectID, .ObjTypeID, 1)
                                    End If
                                End If
                            End If
                        End With
                    End If
                End If
            Next X


            For X = 0 To glColonyUB
                If glColonyIdx(X) <> -1 Then
                    Dim oColony As Colony = goColony(X)
                    If oColony Is Nothing = False Then oColony.HandleColonyTaxIncome()
                End If
            Next X

            'credits and Mining
            'TODO: Do to Mining what we did to Production...

            'now, mining
            Dim lCacheIndex As Int32
            Dim bFound As Boolean = False
            Dim lNewQty As Int32
            For X = 0 To glUnitUB
                If glUnitIdx(X) <> -1 Then
                    Dim oUnit As Unit = goUnit(X)
                    Try
                        If oUnit Is Nothing = True Then
                            'glUnitIdx(X) = -1
                            Continue For
                        End If

                        If oUnit.ParentObject Is Nothing = False Then
                            With CType(oUnit.ParentObject, Epica_GUID)
                                If .ObjTypeID = ObjectType.ePlanet OrElse .ObjTypeID = ObjectType.eSolarSystem Then
                                    oUnit.Owner.oBudget.AddUnitToEnvir(.ObjectID, .ObjTypeID, (oUnit.EntityDef.yChassisType And (ChassisType.eGroundBased Or ChassisType.eNavalBased)) = 0, oUnit.ExpLevel)
                                Else
                                    Dim oTmpParent As Epica_GUID = oUnit.GetRootParentEnvir()
                                    If oTmpParent Is Nothing = False Then
                                        oUnit.Owner.oBudget.AddDockedUnitToEnvir(oTmpParent.ObjectID, oTmpParent.ObjTypeID, (oUnit.EntityDef.yChassisType And (ChassisType.eGroundBased Or ChassisType.eNavalBased)) = 0)
                                    End If
                                    oTmpParent = Nothing
                                End If
                            End With
                        Else : Continue For
                        End If

                        'Check if the player needs to update warpoints
                        'If oUnit.Owner Is Nothing = False AndAlso oUnit.bUnitInSellOrder = False Then
                        '    If (oUnit.CurrentStatus And elUnitStatus.eGuildAsset) = 0 Then
                        '        oUnit.Owner.lCurrentWarpointUpkeepCost += oUnit.EntityDef.WarpointUpkeep
                        '    End If
                        '    If (oUnit.Owner.lTemporaryWarpointCheckFlag And elWarpointCheckFlag.eCheckAllNormalUnits) <> 0 Then
                        '        oUnit.Owner.lTemporaryWarpointUpkeepCost += oUnit.EntityDef.WarpointUpkeep
                        '    End If
                        '    If (oUnit.CurrentStatus And elUnitStatus.eGuildAsset) <> 0 AndAlso (oUnit.Owner.lTemporaryWarpointCheckFlag And elWarpointCheckFlag.eCheckAllGuildShare) <> 0 Then
                        '        oUnit.Owner.lTemporaryWarpointUpkeepCost += oUnit.EntityDef.WarpointUpkeep
                        '    End If
                        'End If

                        If oUnit.bMining = True AndAlso oUnit.lCacheIndex <> -1 AndAlso _
                          ((oUnit.CurrentStatus And elUnitStatus.eMoveLock) <> 0) _
                           AndAlso ((oUnit.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0) Then
                            'ok, last check, does the cache index exist and is it still there?
                            If glMineralCacheIdx(oUnit.lCacheIndex) = oUnit.lCacheID Then
                                lCacheIndex = oUnit.lCacheIndex

                                Dim lUnitCargoCap As Int32 = oUnit.Cargo_Cap

                                If lUnitCargoCap = 0 Then
                                    'oUnit.HandleReturnToRefinery()
                                    oUnit.ProcessNextRouteItem()
                                Else
                                    'Ok, let's do some mining... first check if concentration > quantity
                                    'check if the concentration is greater than my cargo hold
                                    Dim oCache As MineralCache = goMineralCache(lCacheIndex)
                                    If oCache Is Nothing = False Then
                                        If oCache.Concentration > lUnitCargoCap Then
                                            'it is, so we scoop up remains of cargo hold
                                            lNewQty = lUnitCargoCap
                                        Else : lNewQty = CInt(oCache.Concentration + oUnit.Owner.oSpecials.yMineralConcentrationBonus)
                                        End If

                                        'Ok, reduce our quantity
                                        oCache.Quantity -= lNewQty

                                        'Now, check for concentration reduction
                                        If oCache.CacheTypeID = MineralCacheType.eMineable Then oCache.HandleDepletion(2.4F)

                                        'Now, add the quantity of goods to the unit...
                                        oUnit.AddMineralCacheToCargo(oCache.oMineral.ObjectID, lNewQty)
                                        'If goUnit(X).AddMineralCacheToCargo(goMineralCache(lCacheIndex).oMineral.ObjectID, lNewQty) Is Nothing = False Then
                                        '    'Now, reduce the unit's cargo cap
                                        '    goUnit(X).Cargo_Cap -= lNewQty
                                        'End If

                                        If oUnit.Cargo_Cap < 1 Then oUnit.ProcessNextRouteItem() ' oUnit.HandleReturnToRefinery()
                                    End If
                                End If
                            Else
                                'No, cache is no longer there, send miner back to refinery
                                'oUnit.HandleReturnToRefinery()
                                oUnit.ProcessNextRouteItem()
                            End If
                        End If
                    Catch
                    End Try
                End If
            Next X

            For X = 0 To glPlayerUB
                If glPlayerIdx(X) <> -1 AndAlso glPlayerIdx(X) <> gl_HARDCODE_PIRATE_PLAYER_ID Then
                    goPlayer(X).oBudget.FinalizeUnitCosts()
                    'If goPlayer(X).lTemporaryWarpointUpkeepCost > 0 Then goPlayer(X).ProcessWarpointReduction()
                End If
            Next X

            CheckFacilityMining()

            'Send our player credits update
            Dim yData(23) As Byte   '35
            For X = 0 To glPlayerUB
                If glPlayerIdx(X) <> -1 Then
                    Dim oPlayer As Player = goPlayer(X)
                    If oPlayer Is Nothing = False Then
                        If oPlayer.InMyDomain = True Then

                            'MSC 1/7/09 - don't trust the osocket query, the lConnectedPrimaryID has been working so far. We'll try the oSocket later
                            'If (oPlayer.oSocket Is Nothing = False OrElse oPlayer.HasOnlineAliases(AliasingRights.eViewTreasury) = True) Then
                            If (oPlayer.lConnectedPrimaryID > -1 OrElse oPlayer.HasOnlineAliases(AliasingRights.eViewTreasury) = True) Then
                                System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerCredits).CopyTo(yData, 0)
                                With oPlayer
                                    .GetGUIDAsString.CopyTo(yData, 2)
                                    System.BitConverter.GetBytes(.blCredits).CopyTo(yData, 8)
                                    System.BitConverter.GetBytes(.oBudget.GetCashFlow()).CopyTo(yData, 16)
                                    'System.BitConverter.GetBytes(.blWarpoints).CopyTo(yData, 24)
                                    'System.BitConverter.GetBytes(.lCurrentWarpointUpkeepCost).CopyTo(yData, 32)
                                    .SendPlayerMessage(yData, False, AliasingRights.eViewTreasury)
                                End With

                                'TODO: This doesn't work right yet, but it SHOULD fix the problem of invulnerability not dropping
                                'If oPlayer.lCurrent15MinutesRemaining > 0 Then
                                '    oPlayer.lCurrent15MinutesRemaining -= 166       '5 seconds in cycles, basically
                                '    If oPlayer.lCurrent15MinutesRemaining < 0 Then oPlayer.lCurrent15MinutesRemaining = 0
                                'Else
                                '    If oPlayer.yIronCurtainState <> eIronCurtainState.IronCurtainIsDown Then
                                '        AddToQueue(glCurrentCycle + 1, QueueItemType.eIronCurtainFall, oPlayer.ObjectID, -1, -1, -1, -1, -1, -1, -1)
                                '    End If
                                'End If
                            End If

                            'here, we will check for automatic mineral movement
                            If oPlayer.oSpecials.yAutomaticMinMove > 0 Then
                                'ok, player has automatic mineral movement
                                oPlayer.HandleAutomaticMineralMovement()
                            End If
                        ElseIf (oPlayer.oBudget Is Nothing = False AndAlso oPlayer.oBudget.mlItemUB > -1) OrElse oPlayer.blCrossPrimaryAdjustCredits <> 0 Then
                            'If player has budget items on this primary, send an update to the operator...
                            Dim yMsg(21) As Byte
                            System.BitConverter.GetBytes(GlobalMessageCode.eCrossPrimaryBudgetAdjust).CopyTo(yMsg, 0)
                            System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yMsg, 2)
                            System.BitConverter.GetBytes(oPlayer.oBudget.GetCashFlow()).CopyTo(yMsg, 6)
                            System.BitConverter.GetBytes(oPlayer.blCrossPrimaryAdjustCredits).CopyTo(yMsg, 14)
                            oPlayer.blCrossPrimaryAdjustCredits = 0
                            goMsgSys.SendMsgToOperator(yMsg)
                        End If
                    End If
                End If
            Next X

            'MSC 12/18/2006
            'mlLastCreditUpdate = timeGetTime
        End If

    End Sub

    Private mbHadSevenEmperors As Boolean = False
    Public Sub ProcessColonyGrowth()
        Dim X As Int32
        Dim blCensus As Int64

        If glCurrentCycle >= glNextColonyGrowthUpdate Then
            For X = 0 To glColonyUB
                If glColonyIdx(X) <> -1 Then
                    Dim oColony As Colony = goColony(X)
                    If oColony Is Nothing = False Then oColony.ProcessGrowth()
                    If glColonyIdx(X) <> -1 Then blCensus += CInt(oColony.Population)
                End If
            Next X

            glNextColonyGrowthUpdate = glCurrentCycle + gl_COLONY_GROWTH_INTERVAL
            If mblLastCensus <> blCensus Then gfrmDisplayForm.SetCensus(blCensus)
            mblLastCensus = blCensus

            'TODO: Not sure if this the best idea for this... figure out where to place this test
            Dim lEmperors As Int32 = 0
            For X = 0 To glPlayerUB
                If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).InMyDomain = True Then
                    goPlayer(X).ReTestTitle()
                    If goPlayer(X).yPlayerTitle = Player.PlayerRank.Emperor Then lEmperors += 1
                End If
            Next X
            If lEmperors > 6 AndAlso mbHadSevenEmperors = False Then
                Dim oComm As OleDb.OleDbCommand = Nothing
                Try
                    oComm = New OleDb.OleDbCommand("UPDATE tblTheSeven SET HadSevenEmperors = 1", goCN)
                    If oComm.ExecuteNonQuery() = 0 Then
                        oComm.Dispose()
                        oComm = Nothing
                        oComm = New OleDb.OleDbCommand("INSERT INTO tblTheSeven (HadSevenEmperors) VALUES (1)", goCN)
                        If oComm.ExecuteNonQuery() = 0 Then Throw New Exception("No records effected.")
                    End If
                Catch ex As Exception
                    LogEvent(LogEventType.CriticalError, "Could not update the seven emperors with 1 value")
                End Try
                mbHadSevenEmperors = True
                If oComm Is Nothing = False Then oComm.Dispose()
                oComm = Nothing
            End If
        End If

    End Sub

    Public Sub StartMainLoop()
        gb_Main_Loop_Running = True
        otMainLoop = New Threading.Thread(AddressOf MainLoop)
        otMainLoop.Start()
    End Sub

    Public Sub MainLoop()
        Dim lSleepTime As Int32
        Dim swMainLoop As Stopwatch = Stopwatch.StartNew
        mswCredits = Stopwatch.StartNew

        Dim swCycler As Stopwatch = Stopwatch.StartNew()
        'Dim lSixMinuteSaver As Int32
        Dim l24HourTimer As Int32 = 0
        Dim lLastGuildBillboardUpdate As Int32 = 0
        Dim lFullCycles As Int32 = 0

        Dim lLastPeriodicUpdate As Int32 = 0

        While gb_Main_Loop_Running = True
            swMainLoop.Reset()
            swMainLoop.Start()

            gb_In_Main_Loop = True

            'glCurrentCycle += 1
            Dim lLastCycleVal As Int32 = glCurrentCycle
            glCurrentCycle = CInt(swCycler.ElapsedMilliseconds \ 30)
            If glCurrentCycle > lLastCycleVal + 2 Then lFullCycles += 1
            gfrmDisplayForm.SetStatusText("FC: " & lFullCycles.ToString)

            Dim lVal As Int32 = 2592000
            'If gb_IS_TEST_SERVER = True Then lVal = 300
            If glCurrentCycle - l24HourTimer > lVal Then
                l24HourTimer = glCurrentCycle
                CheckRetestPlayerLinks()
            End If

            'Make sure any items that are waiting to be processed are processed
            ProcessQueue()

            'Do our work here...
            ProcessColonyGrowth()
            ProductionAndMining()

            If glCurrentCycle - mlPreviousGTCDeadline > 900 Then
                goGTC.CheckBuyOrderDeadlines()
                mlPreviousGTCDeadline = glCurrentCycle
            End If

            'If glCurrentCycle - mlPreviousSentimentCheck > 10368000 Then
            '	mlPreviousSentimentCheck = glCurrentCycle
            '	DoSentimentCheck()
            'End If

            goAgentEngine.ProcessQueue()

            gb_In_Main_Loop = False

            'If glCurrentCycle - lSixMinuteSaver > 10800 Then	'6 minutes
            '	lSixMinuteSaver = glCurrentCycle
            '	SaveSixMinuteSnapshot(mlLastCensus)
            'End If

            If goAureliusAI Is Nothing = False Then goAureliusAI.ProcessAureliusAI()

            Guild.GuildCheck()

            If glCurrentCycle - lLastGuildBillboardUpdate > 1800 Then       'once per minute
                lLastGuildBillboardUpdate = glCurrentCycle
                For X As Int32 = 0 To glPlanetUB
                    If glPlanetIdx(X) > -1 AndAlso glPlanetIdx(X) < 500000000 Then
                        goPlanet(X).ProcessBids()
                    End If
                Next X
            End If

            If glCurrentCycle - lLastPeriodicUpdate > 54000 Then        '30 minutes
                lLastPeriodicUpdate = glCurrentCycle
                PeriodicProdSave()
            End If

            'If goMsgSys.CheckOperatorConnection() = False Then
            '	goMsgSys.OperatorFailure()
            'End If

            'give our threads/processes time to work...
            'Application.DoEvents()
            Threading.Thread.Sleep(1)

            'Do an asynchronous pass
            AsyncPersistence()

            If goMsgSys.swEmailSrvrReconnect Is Nothing = False Then
                If goMsgSys.swEmailSrvrReconnect.IsRunning = True Then
                    If goMsgSys.swEmailSrvrReconnect.ElapsedMilliseconds > 10000 Then
                        goMsgSys.swEmailSrvrReconnect.Reset()
                        'goMsgSys.swEmailSrvrReconnect.Start()
                        goMsgSys.SendOperatorRequestServerDetails(MsgSystem.ConnectionType.eEmailServerApp)
                    End If
                End If
            End If

            'Now, tell us to sleep here...
            lSleepTime = CInt(30L - swMainLoop.ElapsedMilliseconds)
            If lSleepTime > 0 Then
                Threading.Thread.Sleep(lSleepTime)
            End If
        End While

        swCycler.Stop()
    End Sub

    Public Sub ProductionCompleted(ByVal CurrentProduction As EntityProduction, ByRef oCreator As Epica_Entity)
        'now, find out our next production
        Dim lIndex As Int32 = -1
        Dim yLowestNum As Byte = 255
        Dim yMsg() As Byte
        Dim iTemp As Int16
        Dim bBypassAlert As Boolean = False

        Dim bNoResetPointsProduced As Boolean = False

        Dim lAliasRight As AliasingRights = 0

        Select Case CurrentProduction.ProductionTypeID
            Case ObjectType.eUnitDef
                'creating a new unit
                lAliasRight = AliasingRights.eAddProduction Or AliasingRights.eCancelProduction Or AliasingRights.eMoveUnits Or AliasingRights.eViewUnitsAndFacilities

                'TODO: when academies, races, etc... are in place, we need to check for them here...

                lIndex = AddUnit(oCreator, GetEpicaUnitDef(CurrentProduction.ProductionID), 0)
                If lIndex = -1 Then
                    GlobalVars.gfrmDisplayForm.AddEventLine("Production Completed, but AddUnit return -1!!!")
                Else
                    'GlobalVars.gfrmDisplayForm.AddEventLine("Production Completed for Unit ID: " & lIndex)

                    'Ok, see if we need to launch this unit
                    Dim bLaunch As Boolean = False
                    Dim oUnit As Unit = goUnit(lIndex)

                    'Units and facilities are the only things that can produce... the unit's parent MUST be a producer!
                    If oUnit Is Nothing = False AndAlso oUnit.ParentObject Is Nothing = False Then
                        With CType(oUnit.ParentObject, Epica_Entity)
                            Dim bRebuilder As Boolean = False
                            If .ObjTypeID = ObjectType.eFacility Then
                                'Ok, check its autolaunch
                                With CType(oUnit.ParentObject, Facility)
                                    bLaunch = .AutoLaunch

                                    'ok, while we are here, check the facility's parent colony
                                    If oUnit.yProductionType = ProductionType.eFacility AndAlso .ParentColony Is Nothing = False Then
                                        If .ParentColony.bWaitingForRebuilder = True Then
                                            .ParentColony.bWaitingForRebuilder = False
                                            'ok, found our rebuilder...
                                            oUnit.oRebuilderFor = .ParentColony
                                            .ParentColony.lRebuilderUnitID = oUnit.ObjectID
                                            bLaunch = True
                                            bRebuilder = True
                                        End If
                                    End If
                                End With
                            Else
                                'bLaunch = (oCreator.iCombatTactics And eiBehaviorPatterns.eTactics_LaunchChildren) <> 0
                                bLaunch = True
                            End If

                            'Ok, make sure it has enough room in the Hangar for the entity
                            If bLaunch = False Then bLaunch = .Hangar_Cap < oUnit.EntityDef.HullSize

                            'Regardless, update the hangar capacity of the Parent (even if we are to launch it)
                            '.Hangar_Cap -= oUnit.EntityDef.HullSize
                            'Add the unit to the hangar...
                            .AddHangarRef(CType(oUnit, Epica_GUID))

                            If .Owner.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                                If .Owner.oBudget Is Nothing = False Then
                                    Dim lEnvir As Int32 = .Owner.ObjectID + 500000000
                                    If .Owner.oBudget.GetEnvirCPUsage(lEnvir, ObjectType.ePlanet) > 310 Then bLaunch = False
                                End If
                            End If

                            If bLaunch = True Then
                                'TODO: For now, we will bypass launch if the hangar is inoperable, but eventually,
                                ' this should not be the case... the following IF statement will need to be replaced
                                If bRebuilder = True OrElse .Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID OrElse (.CurrentStatus And elUnitStatus.eHangarOperational) = 0 Then
                                    For lHDIdx As Int32 = 0 To .lHangarUB
                                        If .lHangarIdx(lHDIdx) = oUnit.ObjectID AndAlso .oHangarContents(lHDIdx).ObjTypeID = ObjectType.eUnit Then
                                            'Now, remove the item from the Hangar...
                                            .lHangarIdx(lHDIdx) = -1
                                            .oHangarContents(lHDIdx) = Nothing
                                        End If
                                    Next lHDIdx

                                    'add our hangar cap back
                                    '.Hangar_Cap += oUnit.EntityDef.HullSize

                                    'set our parent object

                                    'Ensure that our parentobject is the creator... the region server requires this
                                    oUnit.LocX = .LocX
                                    oUnit.LocZ = .LocZ
                                    oUnit.DataChanged()

                                    'Send off our data... add... to our parent's parent...
                                    iTemp = CType(.ParentObject, Epica_GUID).ObjTypeID
                                    If iTemp = ObjectType.ePlanet Then
                                        'CType(.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oUnit, GlobalMessageCode.eRequestUndock))
                                        CType(.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetUndockCommandMsg(oUnit, CType(.ParentObject, Epica_GUID).ObjectID, CType(.ParentObject, Epica_GUID).ObjTypeID, oCreator.LocX, oCreator.LocZ))
                                    ElseIf iTemp = ObjectType.eSolarSystem Then
                                        'CType(.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oUnit, GlobalMessageCode.eRequestUndock))
                                        CType(.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetUndockCommandMsg(oUnit, CType(.ParentObject, Epica_GUID).ObjectID, CType(.ParentObject, Epica_GUID).ObjTypeID, oCreator.LocX, oCreator.LocZ))
                                    End If

                                    'Now, set the ParentObject to the Creator's ParentObject because by us sending that msg
                                    '  the region server will make it so
                                    oUnit.ParentObject = .ParentObject
                                    'One more call to datachanged
                                    oUnit.DataChanged()

                                    'Tell it where to go
                                    If bRebuilder = False OrElse oUnit.oRebuilderFor Is Nothing Then
                                        .SendUndockFirstWaypoint(CType(oUnit, Epica_Entity))
                                    Else
                                        'ok, let's do this
                                        oUnit.oRebuilderFor.DoNextRebuildOrder(oUnit)
                                    End If

                                Else
                                    'Ok, add the queue event
                                    AddToQueue(glCurrentCycle, QueueItemType.eHandleUndockRequest_QIT, oUnit.ObjectID, oUnit.ObjTypeID, .ObjectID, .ObjTypeID, 0, 0, 0, 0)
                                End If
                            End If
                        End With

                        If oCreator.Owner Is Nothing = False AndAlso oCreator.Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID Then
                            With CType(oCreator, Facility)
                                oUnit.lPirate_For_PlayerID = .lPirate_For_PlayerID
                                .lPirate_Counter += 1
                                .lPirate_ItemUB += 1
                                ReDim Preserve .lPirate_Items(.lPirate_ItemUB)
                                .lPirate_Items(.lPirate_ItemUB) = oUnit.ObjectID
                                If .lPirate_Counter >= .lPirate_Counter_Max Then
                                    If SendPirateRaidForce(CType(oCreator, Facility)) = True Then
                                        .lPirate_Counter_Max += 2
                                        If .lPirate_Counter_Max > 16 Then .lPirate_Counter_Max = 16
                                        .lPirate_Counter = 0
                                    End If
                                    .lPirate_ItemUB = -1
                                    ReDim .lPirate_Items(-1)
                                End If

                                'Now, if the facility is ready to create another group
                                If .lPirate_Counter < .lPirate_Counter_Max Then
                                    AddPirateProductionItem(CType(oCreator, Facility))
                                End If
                            End With
                        End If
                    End If

                    ''Now, send the add object message to the right domain server...
                    'iTemp = CType(goUnit(lIndex).ParentObject, Epica_GUID).ObjTypeID
                    'If iTemp = ObjectType.ePlanet Then
                    '    CType(goUnit(lIndex).ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(goUnit(lIndex), EpicaMessageCode.eAddObjectCommand))
                    '    goMsgSys.SendPathfindingNewEntity(CType(goUnit(lIndex), Epica_GUID), CType(goUnit(lIndex).ParentObject, Epica_GUID), goUnit(lIndex).LocX, goUnit(lIndex).LocZ, goUnit(lIndex).EntityDef.ModelID)
                    'ElseIf iTemp = ObjectType.eSolarSystem Then
                    '    CType(goUnit(lIndex).ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(goUnit(lIndex), EpicaMessageCode.eAddObjectCommand))
                    '    goMsgSys.SendPathfindingNewEntity(CType(goUnit(lIndex), Epica_GUID), CType(goUnit(lIndex).ParentObject, Epica_GUID), goUnit(lIndex).LocX, goUnit(lIndex).LocZ, goUnit(lIndex).EntityDef.ModelID)
                    'End If

                    ''Now, send a move object message... (which also adds the object to the PF server for us)
                    'oCreator.SendUndockFirstWaypoint(CType(goUnit(lIndex), Epica_Entity))
                End If
            Case ObjectType.eFacilityDef
                'creating a new facility...
                'Use lProdX, lProdZ, iProdA for locational data
                lAliasRight = AliasingRights.eAddProduction Or AliasingRights.eAlterAutoLaunchPower Or AliasingRights.eCancelProduction Or AliasingRights.eViewColonyStats Or AliasingRights.eViewUnitsAndFacilities

                With CurrentProduction
                    Dim oFacDef As FacilityDef = GetEpicaFacilityDef(.ProductionID)
                    If oFacDef Is Nothing Then Return
                    lIndex = AddFacility(oCreator, oFacDef, .lProdX, .lProdZ, .iProdA)

                    'HandleActOfWarTest(CType(oFacDef, Epica_Entity_Def), .lProdX, .lProdZ, oCreator)
                End With

                If lIndex = -1 Then
                    GlobalVars.gfrmDisplayForm.AddEventLine("Production Completed, but AddFacility return -1!!!")
                Else
                    Dim oFac As Facility = goFacility(lIndex)
                    If oFac Is Nothing = False Then

                        If oFac.yProductionType = ProductionType.eTradePost Then
                            If goGTC Is Nothing = False Then goGTC.ReRegisterTradeDeliveries(oFac)
                        End If

                        oFac.Owner.TestCustomTitlePermissions_Facility(oFac)

                        'Now, send the add object message to the right domain server...
                        If oCreator.yProductionType <> ProductionType.eSpaceStationSpecial Then
                            iTemp = CType(oCreator.ParentObject, Epica_GUID).ObjTypeID
                            If iTemp = ObjectType.ePlanet Then

                                'If (oFac.yProductionType And ProductionType.eProduction) <> 0 Then
                                '	If oCreator.Owner.PirateStartLocX = Int32.MinValue Then
                                '		oCreator.Owner.PirateStartLocX = 0

                                '		'Dim yPSLMsg(17) As Byte
                                '		'System.BitConverter.GetBytes(GlobalMessageCode.eGetPirateStartLoc).CopyTo(yPSLMsg, 0)
                                '		'System.BitConverter.GetBytes(CType(oCreator.ParentObject, Epica_GUID).ObjectID).CopyTo(yPSLMsg, 2)
                                '		'System.BitConverter.GetBytes(oCreator.Owner.ObjectID).CopyTo(yPSLMsg, 6)
                                '		'System.BitConverter.GetBytes(oCreator.Owner.lStartLocX).CopyTo(yPSLMsg, 10)
                                '		'System.BitConverter.GetBytes(oCreator.Owner.lStartLocZ).CopyTo(yPSLMsg, 14)
                                '		'CType(oCreator.ParentObject, Planet).oDomain.DomainSocket.SendData(yPSLMsg)

                                '		oCreator.Owner.bInPirateSpawn = True
                                '	End If
                                'End If

                                CType(oCreator.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oFac, GlobalMessageCode.eAddObjectCommand))
                                goMsgSys.SendPathfindingNewEntity(CType(oFac, Epica_GUID), CType(oFac.ParentObject, Epica_GUID), oFac.LocX, oFac.LocZ, oFac.EntityDef.ModelID)
                            ElseIf iTemp = ObjectType.eSolarSystem Then
                                CType(oCreator.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oFac, GlobalMessageCode.eAddObjectCommand))
                                goMsgSys.SendPathfindingNewEntity(CType(oFac, Epica_GUID), CType(oFac.ParentObject, Epica_GUID), oFac.LocX, oFac.LocZ, oFac.EntityDef.ModelID)
                            End If
                        End If

                        'Notify GNS of any space station completions
                        If oFac.yProductionType = ProductionType.eSpaceStationSpecial AndAlso oFac.Owner.yPlayerPhase = eyPlayerPhase.eFullLivePhase Then
                            If CType(oCreator.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
                                Dim oPlanet As Planet = CType(oCreator.ParentObject, SolarSystem).GetNearestPlanet(oFac.LocX, oFac.LocZ)
                                Dim yGNS(73) As Byte
                                Dim lPos As Int32 = 0
                                System.BitConverter.GetBytes(GlobalMessageCode.eNewsItem).CopyTo(yGNS, lPos) : lPos += 2
                                yGNS(lPos) = NewsItemType.eNewSpaceStation : lPos += 1
                                CType(oCreator.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yGNS, lPos) : lPos += 6
                                System.BitConverter.GetBytes(oCreator.Owner.ObjectID).CopyTo(yGNS, lPos) : lPos += 4
                                oPlanet.PlanetName.CopyTo(yGNS, lPos) : lPos += 20

                                oCreator.Owner.PlayerName.CopyTo(yGNS, lPos) : lPos += 20
                                yGNS(lPos) = oCreator.Owner.yGender : lPos += 1
                                oCreator.Owner.EmpireName.CopyTo(yGNS, lPos) : lPos += 20

                                goMsgSys.SendToEmailSrvr(yGNS)
                            End If
                        End If

                        If oCreator.Owner.DeathBudgetEndTime > glCurrentCycle AndAlso oCreator.Owner.DeathBudgetFundsRemaining > 0 Then
                            If oFac.yProductionType = ProductionType.eColonists Then
                                'if productiontype = colonists then add colonists
                                If oFac.yProductionType = ProductionType.eColonists Then
                                    oFac.ParentColony.Population += oFac.EntityDef.ProdFactor
                                    oFac.Owner.blDBPopulation += oFac.EntityDef.ProdFactor
                                Else
                                    oFac.ParentColony.Population += oFac.EntityDef.WorkerFactor
                                    oFac.Owner.blDBPopulation += oFac.EntityDef.WorkerFactor
                                End If
                            End If
                        End If

                        If oFac.ParentColony Is Nothing = False Then
                            AddToQueue(glCurrentCycle + 30, QueueItemType.eCheckColonyResearchQueue, oFac.ParentColony.ObjectID, 0, 0, 0, 0, 0, 0, 0)
                        End If

                        If oCreator.ObjTypeID = ObjectType.eUnit Then
                            CType(oCreator, Unit).bProdQueueMoveSent = False
                        End If
                    End If
                End If
            Case ObjectType.eEnlisted
                'creating new Crew... general crew.
                lAliasRight = AliasingRights.eAddProduction Or AliasingRights.eViewColonyStats
                If CType(oCreator, Facility).ParentColony Is Nothing = False Then
                    CType(oCreator, Facility).ParentColony.ColonyEnlisted += 10
                Else
                    'Ok, add it to the Cargo, if there is room...
                    If oCreator.Cargo_Cap > 10 Then
                        Dim bFound As Boolean = False
                        For lCIdx As Int32 = 0 To oCreator.lCargoUB
                            If oCreator.lCargoIdx(lCIdx) <> -1 AndAlso oCreator.oCargoContents(lCIdx).ObjTypeID = ObjectType.eEnlisted Then
                                bFound = True
                                oCreator.oCargoContents(lCIdx).ObjectID += 10
                                oCreator.lCargoIdx(lCIdx) = oCreator.oCargoContents(lCIdx).ObjectID
                                Exit For
                            End If
                        Next lCIdx

                        If bFound = False Then
                            Dim oGUID As New Epica_GUID
                            oGUID.ObjectID = 9
                            oGUID.ObjTypeID = ObjectType.eEnlisted
                            oCreator.AddCargoRef(oGUID)
                        End If

                        'oCreator.Cargo_Cap -= 10
                    End If
                End If
                If CurrentProduction.lProdCount > 1 Then bBypassAlert = True
            Case ObjectType.eOfficers
                'creating officers
                lAliasRight = AliasingRights.eAddProduction Or AliasingRights.eViewColonyStats
                If CType(oCreator, Facility).ParentColony Is Nothing = False Then
                    CType(oCreator, Facility).ParentColony.ColonyOfficers += 10

                    If oCreator.Owner.yPlayerPhase <> eyPlayerPhase.eInitialPhase Then
                        'Now, does an agent emerge?
                        Dim lAgentCnt As Int32 = 0
                        With oCreator.Owner
                            For X As Int32 = 0 To .mlAgentUB
                                If .mlAgentIdx(X) <> -1 Then lAgentCnt += 1
                            Next X
                        End With
                        If lAgentCnt < 50 Then
                            If lAgentCnt < 10 OrElse CInt(Rnd() * 100) < 10 Then ' OrElse gb_IS_TEST_SERVER = True Then
                                'LogEvent(LogEventType.Informational, "Agent Generation Added to Queue")
                                AddToQueue(glCurrentCycle + 900, QueueItemType.eGenerateAgent, oCreator.Owner.ObjectID, ObjectType.ePlayer, -1, -1, 0, 0, 0, 0)
                            End If
                        End If
                    End If
                Else
                    'Ok, add it to the Cargo, if there is room...
                    If oCreator.Cargo_Cap > 10 Then
                        Dim bFound As Boolean = False
                        For lCIdx As Int32 = 0 To oCreator.lCargoUB
                            If oCreator.lCargoIdx(lCIdx) <> -1 AndAlso oCreator.oCargoContents(lCIdx).ObjTypeID = ObjectType.eOfficers Then
                                bFound = True
                                oCreator.oCargoContents(lCIdx).ObjectID += 10
                                oCreator.lCargoIdx(lCIdx) = oCreator.oCargoContents(lCIdx).ObjectID
                                Exit For
                            End If
                        Next lCIdx

                        If bFound = False Then
                            Dim oGUID As New Epica_GUID
                            oGUID.ObjectID = 9
                            oGUID.ObjTypeID = ObjectType.eOfficers
                            oCreator.AddCargoRef(oGUID)
                        End If

                        'oCreator.Cargo_Cap -= 10
                    End If
                End If
            Case ObjectType.eCredits
                'creating credits
            Case ObjectType.eMorale
                'creating morale
            Case ObjectType.eFood
                'creating food
            Case ObjectType.eMineralTech
                'researching a mineral (or alloy in its mineral form)
                lAliasRight = AliasingRights.eAddResearch Or AliasingRights.eViewResearch
                For X As Int32 = 0 To oCreator.Owner.lPlayerMineralUB
                    If oCreator.Owner.oPlayerMinerals(X).lMineralID = CurrentProduction.ProductionID Then
                        Dim bFurtherStudy As Boolean = oCreator.Owner.oPlayerMinerals(X).ResearchComplete()
                        If CurrentProduction.lProdCount > 1 AndAlso bFurtherStudy = False Then
                            CurrentProduction.lProdCount = 0
                        End If
                        Exit For
                    End If
                Next X
                'Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHangarTech, ObjectType.eHullTech, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.ePrototype, ObjectType.eSpecialTech
            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.ePrototype, ObjectType.eSpecialTech
                Dim oTech As Epica_Tech = Nothing

                oTech = oCreator.Owner.GetTech(CurrentProduction.ProductionID, CurrentProduction.ProductionTypeID)

                If oTech Is Nothing = False Then
                    'Now, check if the Creator's production type is Production
                    If (oCreator.yProductionType And ProductionType.eProduction) <> 0 OrElse oCreator.yProductionType = ProductionType.eRefining Then
                        lAliasRight = AliasingRights.eAddProduction Or AliasingRights.eCancelProduction
                        'if it is, create a new ComponentCache instead of researching the component
                        'Ok, this requires that the tech be researched already
                        If oTech.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                            'Ok... create a component cache if necessary...
                            Dim lQty As Int32 = 1
                            If CurrentProduction.ProductionTypeID = ObjectType.eArmorTech AndAlso CurrentProduction.lProdCount <> 1 Then
                                bBypassAlert = True

                                If oCreator.ObjTypeID = ObjectType.eFacility Then
                                    'Now, determine the number of armor plates built in the time that has passed since last update... 
                                    'that determines our qty
                                    Dim lQtyProduced As Int32 = CInt(CurrentProduction.PointsProduced \ Math.Max(1, CurrentProduction.PointsRequired))
                                    If lQtyProduced > CurrentProduction.lProdCount Then lQtyProduced = CurrentProduction.lProdCount
                                    lQtyProduced = Math.Min(lQtyProduced, CurrentProduction.lProdCount)
                                    lQtyProduced = Math.Max(lQtyProduced, 1)

                                    Dim bBypassCheck As Boolean = False
                                    If oCreator.Owner.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                                        If oCreator.Owner.lTutorialStep = 221 Then
                                            bBypassCheck = True
                                            lQtyProduced = CurrentProduction.lProdCount
                                            CurrentProduction.lProdCount = 0
                                        End If
                                    End If

                                    If bBypassCheck = False AndAlso lQtyProduced > 1 AndAlso CType(oCreator, Facility).ParentColony Is Nothing = False Then
                                        Dim oTmpProdCost As ProductionCost = CurrentProduction.ProdCost.CreateClone(lQtyProduced - 1)   'minus one because we already bought the first one
                                        If CType(oCreator, Facility).ParentColony.HasRequiredResources(oTmpProdCost, CType(oCreator, Facility), eyHasRequiredResourcesFlags.DoNotWait Or eyHasRequiredResourcesFlags.DoNotAlertLowRes) <> eResourcesResult.Sufficient Then
                                            lQtyProduced = 1
                                        End If
                                        CurrentProduction.lProdCount -= (lQtyProduced - 1)
                                    End If

                                    CurrentProduction.PointsProduced = Math.Max(0, CurrentProduction.PointsProduced - (CurrentProduction.PointsRequired * lQtyProduced))
                                    bNoResetPointsProduced = True
                                    lQty = lQtyProduced
                                End If
                            End If

                            'TODO: When components take cargo space, need to adjust this accordingly
                            Dim oReceiver As Facility = Nothing
                            If oCreator.Cargo_Cap > 0 Then
                                oReceiver = CType(oCreator, Facility)
                            Else : oReceiver = CType(oCreator, Facility).ParentColony.GetFacilityWithCargoSpace(lQty)
                            End If

                            If oReceiver Is Nothing = False Then
                                oReceiver.AddComponentCacheToCargo(oTech.ObjectID, oTech.ObjTypeID, lQty, oTech.Owner.ObjectID)
                            Else
                                'TODO: What do we do? add it to the producer if we can...
                                oCreator.AddComponentCacheToCargo(oTech.ObjectID, oTech.ObjTypeID, lQty, oTech.Owner.ObjectID)
                                'And then set our production count to 0
                                CurrentProduction.lProdCount = 0
                            End If
                        Else
                            'Shouldn't be here
                            LogEvent(LogEventType.PossibleCheat, "ProductionCompleted on Tech that's not researched by a Producer. OwnerID = " & oCreator.Owner.ObjectID)
                            bBypassAlert = True
                        End If
                    Else
                        lAliasRight = AliasingRights.eAddResearch Or AliasingRights.eViewResearch
                        'Ok, this is research... check if the creator is primary
                        If oTech.IsPrimaryResearcher(oCreator.ObjectID) = True Then
                            If oTech.ResearchComplete() = False Then
                                'Ok, this research failed... increment the production count so that it goes back into it
                                CurrentProduction.lProdCount += 1
                                bBypassAlert = True
                                CurrentProduction.lProdCount = -1
                            Else
                                'Anytime the component researches true, we send it as a new add object
                                If oCreator.Owner Is Nothing = False Then
                                    oCreator.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oTech, GlobalMessageCode.eAddObjectCommand), True, AliasingRights.eViewTechDesigns Or AliasingRights.eAddProduction Or AliasingRights.eViewResearch Or AliasingRights.eAddResearch)
                                    oCreator.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oTech.GetResearchCost, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewTechDesigns Or AliasingRights.eAddProduction Or AliasingRights.eViewResearch Or AliasingRights.eAddResearch)
                                    oCreator.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oTech.GetProductionCost, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewTechDesigns Or AliasingRights.eAddProduction Or AliasingRights.eViewResearch Or AliasingRights.eAddResearch)
                                End If
                            End If
                        Else
                            'Ok, not the primary...
                            If oTech.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                                'Ok, that's bad, we should not be researching this....
                                CurrentProduction.lProdCount = 0
                            Else
                                oCreator.CurrentProduction.lFinishCycle = Int32.MaxValue
                                Return
                            End If
                        End If
                    End If
                End If
            Case ObjectType.eMineral
                'A refinery is building a mineral alloy... get the mineral
                lAliasRight = AliasingRights.eAddProduction Or AliasingRights.eCancelProduction

                With CType(oCreator, Facility)
                    If .ParentColony Is Nothing = False Then
                        If .ParentColony.Cargo_Cap > 9 Then
                            .ParentColony.AdjustColonyMineralCache(CurrentProduction.ProductionID, 10)
                        Else
                            .Owner.SendPlayerMessage(.ParentColony.GetLowResourcesMsg(ProductionType.eWareHouse, 0, 0, CurrentProduction.ProductionID, CurrentProduction.ProductionTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget)
                            CurrentProduction.lProdCount = 0
                            .ClearQueue()
                        End If
                    Else
                        If oCreator.Cargo_Cap > 9 Then
                            oCreator.AddMineralCacheToCargo(CurrentProduction.ProductionID, 10)
                        Else
                            Dim bFound As Boolean = False
                            With CType(oCreator, Facility).ParentColony
                                For lFac As Int32 = 0 To .ChildrenUB
                                    If .lChildrenIdx(lFac) <> -1 AndAlso .lChildrenIdx(lFac) <> oCreator.ObjectID AndAlso _
                                      (.oChildren(lFac).yProductionType = ProductionType.eWareHouse OrElse _
                                       .oChildren(lFac).yProductionType = ProductionType.eRefining) AndAlso .oChildren(lFac).Cargo_Cap > 0 Then
                                        .oChildren(lFac).AddMineralCacheToCargo(CurrentProduction.ProductionID, 10)
                                        bFound = True
                                        Exit For
                                    End If
                                Next lFac
                            End With
                            If bFound = False Then
                                oCreator.Owner.SendPlayerMessage(CType(oCreator, Facility).ParentColony.GetLowResourcesMsg(ProductionType.eWareHouse, 0, 0, CurrentProduction.ProductionID, CurrentProduction.ProductionTypeID), True, AliasingRights.eViewTreasury Or AliasingRights.eViewColonyStats Or AliasingRights.eViewBudget)
                                CurrentProduction.lProdCount = 0
                            End If
                        End If
                    End If
                End With
                If CurrentProduction.lProdCount > 1 Then bBypassAlert = True
            Case ObjectType.eRepairItem
                'ok, we are repairing an item...
                lAliasRight = AliasingRights.eAddProduction Or AliasingRights.eCancelProduction
                oCreator.RepairCompleted(CurrentProduction)
                'If oCreator.ObjTypeID = ObjectType.eFacility Then
                '	CType(oCreator, Facility).RepairCompleted(CurrentProduction)
                'Else
                '	'TODO: Define units repairing...
                'End If
            Case ObjectType.eAmmunition
                'oCreator.HandleAmmoProduced(CurrentProduction)
                'If CurrentProduction.lProdCount <> 1 Then bBypassAlert = True
        End Select

        'Notify the owner (if they are online) that production was completed
        If bBypassAlert = False AndAlso (oCreator.Owner.lConnectedPrimaryID > -1 OrElse oCreator.Owner.HasOnlineAliases(lAliasRight) = True) Then
            ReDim yMsg(20)
            System.BitConverter.GetBytes(GlobalMessageCode.eProductionCompleteNotice).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(CurrentProduction.ProductionID).CopyTo(yMsg, 2)
            System.BitConverter.GetBytes(CurrentProduction.ProductionTypeID).CopyTo(yMsg, 6)
            oCreator.GetGUIDAsString.CopyTo(yMsg, 8)

            If CType(oCreator.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                'Ok, send the creator's parent object's parent object
                CType(CType(oCreator.ParentObject, Facility).ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, 14)
            Else : CType(oCreator.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, 14)
            End If

            yMsg(20) = oCreator.yProductionType

            oCreator.Owner.SendPlayerMessage(yMsg, False, lAliasRight)
        End If

        'Decrement our ProdCount
        CurrentProduction.lProdCount -= 1
        'Set to false because the item just built was paid for but the next item in the ProdCount isn't
        CurrentProduction.bPaidFor = False

        'now, check if this is a multiple production item
        If oCreator.ObjTypeID = ObjectType.eFacility Then
            If CurrentProduction.lProdCount <= 0 Then
                'ok, time to remove this production and go to the next
                CType(oCreator, Facility).GetNextProduction(False)

                AddToQueue(glCurrentCycle + 30, QueueItemType.eCheckColonyResearchQueue, CType(oCreator, Facility).ParentColony.ObjectID, 0, 0, 0, 0, 0, 0, 0)
            Else
                Dim yRes As eResourcesResult = CType(oCreator, Facility).ParentColony.HasRequiredResources(CurrentProduction.ProdCost, CType(oCreator, Facility), eyHasRequiredResourcesFlags.NoFlags)
                If yRes = eResourcesResult.Sufficient Then
                    CurrentProduction.bPaidFor = True
                    If bNoResetPointsProduced = False Then CurrentProduction.PointsProduced = 0
                    CurrentProduction.lLastUpdateCycle = glCurrentCycle
                ElseIf yRes = eResourcesResult.Insufficient_Clear Then
                    CurrentProduction.lProdCount = 0
                    CType(oCreator, Facility).GetNextProduction(False)
                End If
            End If
        Else

            oCreator.bProducing = False
            CType(oCreator, Epica_Entity).CurrentProduction = Nothing
            If (oCreator.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then oCreator.CurrentStatus = oCreator.CurrentStatus Xor elUnitStatus.eMoveLock

            ReDim yMsg(11)
            System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityStatus).CopyTo(yMsg, 0)
            oCreator.GetGUIDAsString.CopyTo(yMsg, 2)
            System.BitConverter.GetBytes((elUnitStatus.eMoveLock) * -1).CopyTo(yMsg, 8)

            'now, send it to domain
            iTemp = CType(oCreator.ParentObject, Epica_GUID).ObjTypeID
            If iTemp = ObjectType.ePlanet Then
                CType(oCreator.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
            ElseIf iTemp = ObjectType.eSolarSystem Then
                CType(oCreator.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
            End If

            If CurrentProduction.ProductionTypeID = ObjectType.eFacilityDef AndAlso oCreator.ObjTypeID <> ObjectType.eFacility Then
                If oCreator.RallyPointX = Int32.MinValue OrElse oCreator.RallyPointZ = Int32.MinValue Then
                    If lIndex <> -1 Then
                        Dim oFac As Facility = goFacility(lIndex)
                        If oFac Is Nothing = False Then
                            ReDim yMsg(13)
                            System.BitConverter.GetBytes(GlobalMessageCode.eMoveEngineer).CopyTo(yMsg, 0)
                            oCreator.GetGUIDAsString.CopyTo(yMsg, 2)
                            oFac.GetGUIDAsString.CopyTo(yMsg, 8)
                            If iTemp = ObjectType.ePlanet Then
                                CType(oCreator.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
                            ElseIf iTemp = ObjectType.eSolarSystem Then
                                CType(oCreator.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
                            End If
                        End If
                    End If
                Else
                    oCreator.SendUndockFirstWaypoint(oCreator)
                End If
            End If

            'now for the rebuilder
            If oCreator.ObjTypeID = ObjectType.eUnit Then
                With CType(oCreator, Unit)
                    If .oRebuilderFor Is Nothing = False Then
                        .oRebuilderFor.DoNextRebuildOrder(CType(oCreator, Unit))
                    Else
                        .bProdQueueMoveSent = False
                        .CheckProdQueue()
                    End If

                    If .EntityDef.ObjectID = 28344 Then        'super space engineer
                        .Owner.DeathBudgetEndTime = glCurrentCycle + 2592000       '24 hrs
                        If .Owner.DeathBudgetBalance < 3000000 Then
                            .Owner.blCredits = 3000000
                            .Owner.DeathBudgetFundsRemaining = 3000000
                        Else
                            .Owner.blCredits = .Owner.DeathBudgetBalance
                            .Owner.DeathBudgetFundsRemaining = .Owner.DeathBudgetBalance
                        End If
                        .Owner.DeathBudgetBalance = 0
                        .Owner.lIronCurtainPlanet = CType(oCreator.ParentObject, Epica_GUID).ObjectID
                        .Owner.lStartedEnvirID = .Owner.lIronCurtainPlanet
                        .Owner.iStartedEnvirTypeID = ObjectType.ePlanet
                        DestroyEntity(oCreator, True, -1, False, "ProductionComplete of Respawn Capsule")
                    End If
                End With
            End If
        End If
    End Sub

	Public Function LaunchEntity(ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16, ByVal oParent As Epica_Entity) As Byte
		'now, because of launch... only the Region server knows where the parent exists... check the parent's hangar
		Dim oObj As Epica_Entity
		Dim bUndockAll As Boolean = (lEntityID = -1 AndAlso iEntityTypeID = -1)
		Dim bCanUndock As Boolean = False
		Dim iTemp As Int16

		'LogEvent(LogEventType.ExtensiveLogging, "LaunchEntity: " & lEntityID & ", " & iEntityTypeID & " to " & oParent.ObjectID & ", " & oParent.ObjTypeID)

		If ((oParent.CurrentStatus And elUnitStatus.eHangarOperational) <> 0) Then
			For lHIdx As Int32 = 0 To oParent.lHangarUB
				If oParent.lHangarIdx(lHIdx) <> -1 AndAlso (bUndockAll = True OrElse (oParent.oHangarContents(lHIdx).ObjectID = lEntityID AndAlso oParent.oHangarContents(lHIdx).ObjTypeID = iEntityTypeID)) Then
					'Set our oObj
					oObj = CType(oParent.oHangarContents(lHIdx), Epica_Entity)

                    Dim yObjChassis As Byte
                    If oObj.ObjTypeID = ObjectType.eUnit Then
                        yObjChassis = CType(oObj, Unit).EntityDef.yChassisType
                    Else
                        yObjChassis = CType(oObj, Facility).EntityDef.yChassisType
                    End If
                    Dim yParentChassis As Byte
                    If oParent.ObjTypeID = ObjectType.eUnit Then
                        yParentChassis = CType(oParent, Unit).EntityDef.yChassisType
                    Else
                        yParentChassis = CType(oParent, Facility).EntityDef.yChassisType
                    End If

					'Now, check our undock
					bCanUndock = False
					If CType(oParent.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
						If oObj.ObjTypeID = ObjectType.eUnit Then
                            If (yObjChassis And ChassisType.eSpaceBased) <> 0 Then
                                bCanUndock = True
                            End If
						Else
                            If (yObjChassis And ChassisType.eSpaceBased) <> 0 Then
                                bCanUndock = True
                            End If
						End If
                    Else
                        'on the planet... 
                        If (yObjChassis And ChassisType.eGroundBased) <> 0 Then
                            'a land unit cannot undock from a naval parent
                            bCanUndock = (yParentChassis And ChassisType.eNavalBased) = 0
                        ElseIf (yObjChassis And ChassisType.eNavalBased) <> 0 Then
                            'a naval unit cannot undock from a non-naval parent
                            bCanUndock = (yParentChassis And ChassisType.eNavalBased) <> 0
                        Else
                            bCanUndock = True
                        End If
                    End If

                    If bCanUndock = True Then

                        If oParent.ObjTypeID = ObjectType.eFacility AndAlso goGTC Is Nothing = False Then ' = ProductionType.eTradePost Then
                            'ok, let's go see if this item is for sale by the colony's tradepost
                            Dim oColony As Colony = CType(oParent, Facility).ParentColony
                            If oColony Is Nothing = False Then
                                For X As Int32 = 0 To oColony.ChildrenUB
                                    Dim oFac As Facility = oColony.oChildren(X)
                                    If oFac Is Nothing = False AndAlso oFac.yProductionType = ProductionType.eTradePost Then
                                        goGTC.CancelSellOrder(oObj.ObjectID, oObj.ObjTypeID, oFac.ObjectID)
                                        Exit For
                                    End If
                                Next X
                            End If

                        End If

                        'Now, remove the item from the Hangar...
                        oParent.lHangarIdx(lHIdx) = -1
                        oParent.oHangarContents(lHIdx) = Nothing

                        'Be sure our parent object is the parent object... the Region Server will know what to do
                        oObj.ParentObject = oParent
                        oObj.LocX = oParent.LocX
                        oObj.LocZ = oParent.LocZ

                        If oObj.ObjTypeID = ObjectType.eUnit Then
                            'oParent.Hangar_Cap += CType(oObj, Unit).EntityDef.HullSize
                            CType(oObj, Unit).DataChanged()
                        Else
                            'oParent.Hangar_Cap += CType(oObj, Facility).EntityDef.HullSize
                            CType(oObj, Facility).DataChanged()
                        End If

                        'Send off our data...
                        iTemp = CType(oParent.ParentObject, Epica_GUID).ObjTypeID
                        If iTemp = ObjectType.ePlanet Then
                            'CType(oParent.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oObj, GlobalMessageCode.eRequestUndock))
                            CType(oParent.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetUndockCommandMsg(CType(oObj, Unit), CType(oParent.ParentObject, Epica_GUID).ObjectID, CType(oParent.ParentObject, Epica_GUID).ObjTypeID, oParent.LocX, oParent.LocZ))
                        ElseIf iTemp = ObjectType.eSolarSystem Then
                            'CType(oParent.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oObj, GlobalMessageCode.eRequestUndock))
                            CType(oParent.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetUndockCommandMsg(CType(oObj, Unit), CType(oParent.ParentObject, Epica_GUID).ObjectID, CType(oParent.ParentObject, Epica_GUID).ObjTypeID, oParent.LocX, oParent.LocZ))
                        End If

                        'we don't update the object until after the message is sent... reason is, 
                        oObj.ParentObject = oParent.ParentObject
                        If oObj.ObjTypeID = ObjectType.eUnit Then
                            CType(oObj, Unit).DataChanged()
                        Else : CType(oObj, Facility).DataChanged()
                        End If

                        Dim bSendFirstWP As Boolean = True
                        If oObj.ObjTypeID = ObjectType.eUnit Then
                            With CType(oObj, Unit)
                                If .lRouteUB <> -1 AndAlso .lCurrentRouteIdx > -1 AndAlso .bRoutePaused = False Then
                                    bSendFirstWP = False
                                End If
                            End With
                        End If
                        If bSendFirstWP = True Then oParent.SendUndockFirstWaypoint(oObj)
                        oObj.CheckUpdateUnitGroup()

                        If bUndockAll = True AndAlso oParent.GetHangarEntityCount <> 0 Then
                            Return 0
                        Else : Return 1
                        End If
                    End If
				End If
			Next lHIdx

			LogEvent(LogEventType.Warning, "LaunchEntity:Not In Parent. " & lEntityID & ", " & iEntityTypeID & " from " & oParent.ObjectID & ", " & oParent.ObjTypeID & ".")
		Else : Return 2
		End If

		Return 1
	End Function

	Private Sub SendDockCommand(ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16, ByRef oParent As Epica_Entity)
		Dim yData() As Byte
		Dim iTemp As Int16

		'Now, go ahead store it
		Dim oEntity As Epica_Entity = CType(GetEpicaObject(lEntityID, iEntityTypeID), Epica_Entity)
		If oEntity Is Nothing = False Then

			'LogEvent(LogEventType.ExtensiveLogging, "SendDockCommand: " & oEntity.ObjectID & ", " & oEntity.ObjTypeID & " to " & oParent.ObjectID & ", " & oParent.ObjTypeID)

			oEntity.ParentObject = oParent
			oEntity.CheckUpdateUnitGroup()
            If (oEntity.CurrentStatus And elUnitStatus.eMoveLock) <> 0 Then oEntity.CurrentStatus = oEntity.CurrentStatus Xor elUnitStatus.eMoveLock
            'oParent.AddHangarRef(CType(oEntity, Epica_GUID))
		Else
			LogEvent(LogEventType.CriticalError, "SendDockCommand: oEntity is nothing. GUID: " & lEntityID & ", " & iEntityTypeID)
			Return
		End If

		iTemp = CType(oParent.ParentObject, Epica_GUID).ObjTypeID

		ReDim yData(11)
		System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityStatus).CopyTo(yData, 0)
		oEntity.GetGUIDAsString.CopyTo(yData, 2)
		System.BitConverter.GetBytes((elUnitStatus.eMoveLock) * -1).CopyTo(yData, 8)
		If iTemp = ObjectType.ePlanet Then
			CType(oParent.ParentObject, Planet).oDomain.DomainSocket.SendData(yData)
		ElseIf iTemp = ObjectType.eSolarSystem Then
			CType(oParent.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yData)
		Else
			LogEvent(LogEventType.CriticalError, "Cannot send Dock Command to region server, Parent Type mismatch!")
		End If

		ReDim yData(13)
		System.BitConverter.GetBytes(GlobalMessageCode.eDockCommand).CopyTo(yData, 0)
		oEntity.GetGUIDAsString.CopyTo(yData, 2)
		oParent.GetGUIDAsString.CopyTo(yData, 8)

		If iTemp = ObjectType.ePlanet Then
			CType(oParent.ParentObject, Planet).oDomain.DomainSocket.SendData(yData)
		ElseIf iTemp = ObjectType.eSolarSystem Then
			CType(oParent.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yData)
		Else
			LogEvent(LogEventType.CriticalError, "Cannot send Dock Command to region server, Parent Type mismatch!")
		End If

	End Sub

	Private Function ProcessLaunchAll(ByVal oParent As Epica_Entity) As Int32
		'now, because of launch... only the Region server knows where the parent exists... check the parent's hangar
		Dim oObj As Epica_Entity
		Dim bCanUndock As Boolean = False
		Dim iTemp As Int16
		Dim lHullSize As Int32
		Dim yMan As Byte

		If ((oParent.CurrentStatus And elUnitStatus.eHangarOperational) <> 0) Then
			For lHIdx As Int32 = 0 To oParent.lHangarUB
				If oParent.lHangarIdx(lHIdx) <> -1 Then
					'Set our oObj
					oObj = CType(oParent.oHangarContents(lHIdx), Epica_Entity)

					'Now, check our undock
					bCanUndock = False
					If oObj.ObjTypeID = ObjectType.eUnit Then
						With CType(oObj, Unit).EntityDef
							If (CType(oParent.ParentObject, Epica_GUID).ObjTypeID <> ObjectType.eSolarSystem) OrElse (.yChassisType And ChassisType.eSpaceBased) <> 0 Then
								bCanUndock = True
							End If
							If bCanUndock = True Then
								lHullSize = .HullSize
								yMan = .Maneuver
							End If
						End With
					Else
						With CType(oObj, Facility).EntityDef
							If (CType(oParent.ParentObject, Epica_GUID).ObjTypeID <> ObjectType.eSolarSystem) OrElse (.yChassisType And ChassisType.eSpaceBased) <> 0 Then
								bCanUndock = True
							End If
							If bCanUndock = True Then
								lHullSize = .HullSize
								yMan = .Maneuver
							End If
						End With
					End If

					If bCanUndock = True Then
						'Now, check for the launch ability
						Dim lValue As Int32
						'Returns -2 if failed, -1 for success, > 0 for delay timer
						If oParent.ObjTypeID = ObjectType.eUnit Then
							lValue = oParent.AttemptDockCommand(lHullSize, yMan)
						Else
							lValue = oParent.AttemptDockCommand(lHullSize, yMan)
						End If

						If lValue = -2 Then
							'because it failed, we'll just evacuate
							Return -1
						ElseIf lValue <> -1 Then
							'Not ready to launch yet, so we'll throw back the value to wait
							Return lValue
						End If

						'Now, remove the item from the Hangar...
						oParent.lHangarIdx(lHIdx) = -1
						oParent.oHangarContents(lHIdx) = Nothing

						If oObj.ObjTypeID = ObjectType.eUnit Then
							'oParent.Hangar_Cap += CType(oObj, Unit).EntityDef.HullSize
							CType(oObj, Unit).DataChanged()
						Else
							'oParent.Hangar_Cap += CType(oObj, Facility).EntityDef.HullSize
							CType(oObj, Facility).DataChanged()
						End If

						'Be sure our parent object is the parent object... the Region Server will know what to do
						oObj.ParentObject = oParent
						oObj.LocX = oParent.LocX
						oObj.LocZ = oParent.LocZ

						'Send off our data...
						iTemp = CType(oParent.ParentObject, Epica_GUID).ObjTypeID
						If iTemp = ObjectType.ePlanet Then
							'CType(oParent.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oObj, GlobalMessageCode.eRequestUndock))
							CType(oParent.ParentObject, Planet).oDomain.DomainSocket.SendData(goMsgSys.GetUndockCommandMsg(CType(oObj, Unit), CType(oParent.ParentObject, Epica_GUID).ObjectID, CType(oParent.ParentObject, Epica_GUID).ObjTypeID, oParent.LocX, oParent.LocZ))
						ElseIf iTemp = ObjectType.eSolarSystem Then
							'CType(oParent.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oObj, GlobalMessageCode.eRequestUndock))
							CType(oParent.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(goMsgSys.GetUndockCommandMsg(CType(oObj, Unit), CType(oParent.ParentObject, Epica_GUID).ObjectID, CType(oParent.ParentObject, Epica_GUID).ObjTypeID, oParent.LocX, oParent.LocZ))
						End If

						'we don't update the object until after the message is sent... reason is, 
						oObj.ParentObject = oParent.ParentObject
						If oObj.ObjTypeID = ObjectType.eUnit Then
							CType(oObj, Unit).DataChanged()
						Else : CType(oObj, Facility).DataChanged()
						End If

						oParent.SendUndockFirstWaypoint(oObj)
						oObj.CheckUpdateUnitGroup()

						'Now, because we are here, we'll return 1 to indicate to wait 1 cycle
						'  only once the hangar is completely empty will we return -1
						Return 1
					End If
				End If
			Next lHIdx

		End If

		'Either the hangar is completely empty or no other entities can launch
		Return -1
	End Function

    'Private miPrevTypeID As Int16 = -1
    '   Private mlPrevIdx As Int32 = -1
    '   Private mlSaveCount As Int32 = 0
    '   Private mlTotalCount As Int32 = 0
    '   Private moAsyncStart As Stopwatch

    'Private Sub AsyncPersistence()
    '	Dim sSQL As String = ""

    '	'So, we cycle through those and save as needed
    '	If miPrevTypeID = -1 Then
    '		mlPrevIdx = -1
    '           miPrevTypeID = ObjectType.eFacility
    '           mlSaveCount = 0
    '           mlTotalCount = 0
    '           'LogEvent(LogEventType.Informational, "Async: Starting Facility")
    '		Return
    '       ElseIf miPrevTypeID = ObjectType.eFacility Then
    '           If moAsyncStart Is Nothing OrElse moAsyncStart.IsRunning = False Then
    '               moAsyncStart = Stopwatch.StartNew()
    '           End If
    '           Try
    '               For X As Int32 = mlPrevIdx + 1 To glFacilityUB
    '                   If glFacilityIdx(X) <> -1 Then
    '                       Dim oFac As Facility = goFacility(X)
    '                       If oFac Is Nothing = False Then
    '                           mlTotalCount += 1
    '                           'Ok, check if its active status changed
    '                           If oFac.Active <> oFac.PreviousActive OrElse oFac.bNeedsSaved = True Then
    '                               mlSaveCount += 1
    '                               'Ok, found one
    '                               mlPrevIdx = X
    '                               oFac.SaveObject()
    '                               'sSQL = "UPDATE tblStructure SET CurrentStatus = " & goFacility(X).CurrentStatus & " WHERE StructureID = " & goFacility(X).ObjectID
    '                               'ExecuteSQL(sSQL)
    '                               'Now, simply return, we've done our update
    '                               oFac.PreviousActive = oFac.Active
    '                               Return
    '                           End If
    '                           'Else : glFacilityIdx(X) = -1
    '                       End If
    '                   End If
    '               Next X
    '               'If we are here, update our Previous Type and set our previous idx
    '               miPrevTypeID = ObjectType.ePlayer : mlPrevIdx = -1
    '               'LogEvent(LogEventType.Informational, "Async: Starting Player. FacStats: " & mlSaveCount & "/" & mlTotalCount)
    '               mlSaveCount = 0
    '               mlTotalCount = 0
    '           Catch
    '               mlPrevIdx += 1
    '           End Try
    '	ElseIf miPrevTypeID = ObjectType.ePlayer Then
    '		Try
    '			For X As Int32 = mlPrevIdx + 1 To glPlayerUB
    '                   If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).InMyDomain = True AndAlso goPlayer(X).yPlayerPhase = eyPlayerPhase.eFullLivePhase AndAlso goPlayer(X).PlayerIsDead = False Then
    '                       'Ok, credits always change, so just update it
    '                       mlPrevIdx = X
    '                       mlSaveCount += 1
    '                       mlTotalCount += 1
    '                       goPlayer(X).SaveObject(True)
    '                       'sSQL = "UPDATE tblPlayer SET Credits = " & goPlayer(X).blCredits & " WHERE PlayerID = " & goPlayer(X).ObjectID
    '                       'ExecuteSQL(sSQL)
    '                       Return
    '                   End If
    '			Next X
    '			'If we are here, update our Previous Type and set our previous idx
    '			miPrevTypeID = ObjectType.eMineralCache : mlPrevIdx = -1
    '               'LogEvent(LogEventType.Informational, "Async: Starting Mineral Cache. PlayerStats: " & mlSaveCount & "/" & mlTotalCount)
    '               mlSaveCount = 0
    '               mlTotalCount = 0
    '		Catch
    '		End Try
    '	ElseIf miPrevTypeID = ObjectType.eMineralCache Then
    '		Try
    '			'We use Mineral as an indicator of any 
    '			For X As Int32 = mlPrevIdx + 1 To glMineralCacheUB
    '				If glMineralCacheIdx(X) <> -1 Then
    '					Dim oCache As MineralCache = goMineralCache(X)
    '                       If oCache Is Nothing = False AndAlso oCache.IsDeleted = False Then
    '                           mlTotalCount += 1
    '                           If oCache.bNeedsAsync = True Then
    '                               mlPrevIdx = X
    '                               mlSaveCount += 1
    '                               If oCache.SaveObject() = False Then
    '                                   oCache.DeleteMe()
    '                                   glMineralCacheIdx(X) = -1
    '                               End If
    '                               Return
    '                           End If
    '                       Else : glMineralCacheIdx(X) = -1
    '                       End If
    '				End If
    '			Next X
    '			'If we are here, update our Previous Type and set our previous idx
    '			miPrevTypeID = ObjectType.eColony : mlPrevIdx = -1
    '               'LogEvent(LogEventType.Informational, "Async: Starting Colony. MinCacheStats: " & mlSaveCount & "/" & mlTotalCount)
    '               mlSaveCount = 0
    '               mlTotalCount = 0
    '		Catch
    '			mlPrevIdx += 1
    '		End Try
    '	ElseIf miPrevTypeID = ObjectType.eColony Then
    '		Try
    '			For X As Int32 = mlPrevIdx + 1 To glColonyUB
    '				If glColonyIdx(X) <> -1 Then
    '					Dim oColony As Colony = goColony(X)
    '                       If oColony Is Nothing = False Then
    '                           mlTotalCount += 1
    '                           mlSaveCount += 1
    '                           mlPrevIdx = X
    '                           oColony.SaveObject()
    '                           Return
    '                           'Else : glColonyIdx(X) = -1
    '                       End If
    '				End If
    '			Next X
    '			'If we are here, update our Previous Type and set our previous idx
    '			miPrevTypeID = ObjectType.eUnit : mlPrevIdx = -1
    '               'LogEvent(LogEventType.Informational, "Async: Starting Unit. ColonyStats: " & mlSaveCount & "/" & mlTotalCount)
    '               mlSaveCount = 0
    '               mlTotalCount = 0
    '		Catch
    '			mlPrevIdx += 1
    '		End Try
    '	ElseIf miPrevTypeID = ObjectType.eUnit Then
    '		Try
    '			For X As Int32 = mlPrevIdx + 1 To glUnitUB
    '				If glUnitIdx(X) <> -1 Then
    '					Dim oUnit As Unit = goUnit(X)
    '                       If oUnit Is Nothing = False Then
    '                           mlTotalCount += 1
    '                           If oUnit.bNeedsSaved = True Then
    '                               mlSaveCount += 1
    '                               mlPrevIdx = X
    '                               oUnit.DoAsyncPersistence()
    '                               Return
    '                           End If
    '                       Else
    '                           'LogEvent(LogEventType.Warning, "glUnitIdx was not -1 but object is nothing: " & glUnitIdx(X))
    '                           'glUnitIdx(X) = -1
    '                       End If
    '				End If
    '			Next X
    '			'If we are here, update our Previous Type and set our previous idx
    '			miPrevTypeID = ObjectType.eComponentCache : mlPrevIdx = -1
    '               'LogEvent(LogEventType.Informational, "Async: Starting Component Cache. UnitStats: " & mlSaveCount & "/" & mlTotalCount)
    '               mlSaveCount = 0
    '               mlTotalCount = 0
    '		Catch
    '			mlPrevIdx += 1
    '		End Try
    '	ElseIf miPrevTypeID = ObjectType.eComponentCache Then
    '		Try
    '			For X As Int32 = mlPrevIdx + 1 To glComponentCacheUB
    '				If glComponentCacheIdx(X) <> -1 Then
    '					Dim oCache As ComponentCache = goComponentCache(X)
    '                       If oCache Is Nothing = False Then
    '                           mlTotalCount += 1
    '                           If oCache.bNeedsSaved = True Then
    '                               mlSaveCount += 1
    '                               mlPrevIdx = X
    '                               If oCache.SaveObject() = False Then
    '                                   oCache.DeleteMe()
    '                                   glComponentCacheIdx(X) = -1
    '                                   Return
    '                               End If
    '                           End If
    '                       End If
    '				End If
    '			Next X
    '			'if we are here, update our previous type and set our previous idx
    '               miPrevTypeID = ObjectType.eAgent : mlPrevIdx = -1
    '               ' LogEvent(LogEventType.Informational, "Async: Starting Agent. CompCacheStats: " & mlSaveCount & "/" & mlTotalCount)
    '               mlSaveCount = 0
    '               mlTotalCount = 0
    '		Catch
    '			mlPrevIdx += 1
    '           End Try
    '       ElseIf miPrevTypeID = ObjectType.eAgent Then
    '           Try
    '               For X As Int32 = mlPrevIdx + 1 To glAgentUB
    '                   If glAgentIdx(X) <> -1 Then
    '                       Dim oAgent As Agent = goAgent(X)
    '                       If oAgent Is Nothing = False Then
    '                           mlPrevIdx = X
    '                           mlTotalCount += 1
    '                           mlSaveCount += 1
    '                           oAgent.SaveObject()
    '                           Return
    '                       End If
    '                   End If
    '               Next X
    '               miPrevTypeID = ObjectType.eMission : mlPrevIdx = -1
    '               ' LogEvent(LogEventType.Informational, "Async: Starting PlayerMissions. AgentStats: " & mlSaveCount & "/" & mlTotalCount)
    '               mlSaveCount = 0
    '               mlTotalCount = 0
    '           Catch
    '               mlPrevIdx += 1
    '           End Try
    '       ElseIf miPrevTypeID = ObjectType.eMission Then
    '           Try
    '               For X As Int32 = mlPrevIdx + 1 To glPlayerMissionUB
    '                   If glPlayerMissionIdx(X) > -1 Then
    '                       Dim oPM As PlayerMission = goPlayerMission(X)
    '                       If oPM Is Nothing = False Then
    '                           mlTotalCount += 1
    '                           If oPM.lCurrentPhase <> oPM.lPreviousAsyncPhase Then
    '                               mlSaveCount += 1
    '                               oPM.lPreviousAsyncPhase = oPM.lCurrentPhase
    '                               mlPrevIdx = X
    '                               oPM.SaveObject()
    '                               Return
    '                           End If
    '                       End If
    '                   End If
    '               Next X
    '               miPrevTypeID = ObjectType.eGuild : mlPrevIdx = -1
    '               ' LogEvent(LogEventType.Informational, "Async: Starting Guilds. MissionStats: " & mlSaveCount & "/" & mlTotalCount)
    '               mlSaveCount = 0
    '               mlTotalCount = 0
    '           Catch
    '               mlPrevIdx += 1
    '           End Try
    '       ElseIf miPrevTypeID = ObjectType.eGuild Then
    '           Try
    '               For X As Int32 = mlPrevIdx + 1 To glGuildUB
    '                   If glGuildIdx(X) > -1 Then
    '                       Dim oGuild As Guild = goGuild(X)
    '                       If oGuild Is Nothing = False Then
    '                           mlTotalCount += 1
    '                           mlSaveCount += 1
    '                           mlPrevIdx = X
    '                           oGuild.SaveObject()
    '                           Return
    '                       End If
    '                   End If
    '               Next X
    '           Catch
    '               mlPrevIdx += 1
    '           End Try
    '           miPrevTypeID = ObjectType.eFacility : mlPrevIdx = -1
    '           'LogEvent(LogEventType.Informational, "Async: Starting Facilities. GuildStats: " & mlSaveCount & "/" & mlTotalCount)
    '           mlSaveCount = 0
    '           mlTotalCount = 0

    '           If moAsyncStart Is Nothing = False Then
    '               moAsyncStart.Stop()
    '               LogEvent(LogEventType.Informational, "Async Series Ended. Duration: " & (moAsyncStart.ElapsedMilliseconds \ 1000).ToString("#,##0") & " seconds")
    '               moAsyncStart.Reset()
    '           End If
    '       Else
    '           miPrevTypeID = -1
    '	End If
    'End Sub
    Private mlPrevFacIdx As Int32 = -1
    Private mlPrevPlayerIdx As Int32 = -1
    Private mlPrevMineralCacheIdx As Int32 = -1
    Private mlPrevColonyIdx As Int32 = -1
    Private mlPrevUnitIdx As Int32 = -1
    Private mlPrevComponentCacheIdx As Int32 = -1
    Private mlPrevAgentIdx As Int32 = -1
    Private mlPrevMissionIdx As Int32 = -1
    Private mlPrevGuildIdx As Int32 = -1
    Private Sub AsyncPersistence()
        Dim sSQL As String = ""

        'So, we cycle through those and save as needed

        'FACILITY ASYNC ======================
        Try
            Dim bFound As Boolean = False
            For X As Int32 = mlPrevFacIdx + 1 To glFacilityUB
                If glFacilityIdx(X) <> -1 Then
                    Dim oFac As Facility = goFacility(X)
                    If oFac Is Nothing = False Then
                        'Ok, check if its active status changed
                        If oFac.Active <> oFac.PreviousActive OrElse oFac.bNeedsSaved = True Then
                            'Ok, found one
                            mlPrevFacIdx = X
                            oFac.SaveObject()
                            oFac.PreviousActive = oFac.Active
                            bFound = True
                            Exit For
                        End If
                        'Else : glFacilityIdx(X) = -1
                    End If
                End If
            Next X
            'check if all facs updated - set our idx back to the beginning
            If bFound = False Then mlPrevFacIdx = -1
        Catch
            mlPrevFacIdx += 1
        End Try

        'PLAYER ASYNC ========================
        Try
            Dim bFound As Boolean = False
            For X As Int32 = mlPrevPlayerIdx + 1 To glPlayerUB
                If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).InMyDomain = True AndAlso goPlayer(X).yPlayerPhase = eyPlayerPhase.eFullLivePhase AndAlso goPlayer(X).PlayerIsDead = False Then
                    'Ok, credits always change, so just update it
                    mlPrevPlayerIdx = X
                    goPlayer(X).SaveObject(True)
                    bFound = True
                    Exit For
                End If
            Next X
            'If we are here, update our Previous Type and set our previous idx
            If bFound = False Then mlPrevPlayerIdx = -1
        Catch
        End Try

        'MINERAL CACHE ASYNC =================
        Try
            'We use Mineral as an indicator of any 
            Dim bFound As Boolean = False
            Dim lCurUB As Int32 = -1
            If glMineralCacheIdx Is Nothing = False Then lCurUB = Math.Min(glMineralCacheUB, glMineralCacheIdx.GetUpperBound(0))
            For X As Int32 = mlPrevMineralCacheIdx + 1 To lCurUB
                If glMineralCacheIdx(X) <> -1 Then
                    Dim oCache As MineralCache = goMineralCache(X)
                    If oCache Is Nothing = False AndAlso oCache.IsDeleted = False Then
                        If oCache.bNeedsAsync = True Then
                            mlPrevMineralCacheIdx = X
                            If oCache.SaveObject() = False Then
                                oCache.DeleteMe()
                                glMineralCacheIdx(X) = -1
                            End If

                            bFound = True
                            Exit For
                        End If
                    Else : glMineralCacheIdx(X) = -1
                    End If
                End If
            Next X
            'If we are here, update our Previous Type and set our previous idx
            If bFound = False Then mlPrevMineralCacheIdx = -1
        Catch
            mlPrevMineralCacheIdx += 1
        End Try

        'COLONY ASYNC =========================
        Try
            Dim bFound As Boolean = False
            For X As Int32 = mlPrevColonyIdx + 1 To glColonyUB
                If glColonyIdx(X) <> -1 Then
                    Dim oColony As Colony = goColony(X)
                    If oColony Is Nothing = False Then
                        mlPrevColonyIdx = X
                        oColony.SaveObject()
                        bFound = True
                        Exit For
                    End If
                End If
            Next X
            If bFound = False Then mlPrevColonyIdx = -1
        Catch
            mlPrevColonyIdx += 1
        End Try

        'UNIT ASYNC ===========================
        Try
            Dim bFound As Boolean = False
            For X As Int32 = mlPrevUnitIdx + 1 To glUnitUB
                If glUnitIdx(X) <> -1 Then
                    Dim oUnit As Unit = goUnit(X)
                    If oUnit Is Nothing = False Then
                        If oUnit.bNeedsSaved = True Then
                            mlPrevUnitIdx = X
                            oUnit.DoAsyncPersistence()
                            bFound = True
                            Exit For
                        End If
                    Else
                        'LogEvent(LogEventType.Warning, "glUnitIdx was not -1 but object is nothing: " & glUnitIdx(X))
                        'glUnitIdx(X) = -1
                    End If
                End If
            Next X
            If bFound = False Then mlPrevUnitIdx = -1
        Catch
            mlPrevUnitIdx += 1
        End Try

        'COMPONENT CACHE ASYNC ====================
        Try
            Dim bFound As Boolean = False
            For X As Int32 = mlPrevComponentCacheIdx + 1 To glComponentCacheUB
                If glComponentCacheIdx(X) <> -1 Then
                    Dim oCache As ComponentCache = goComponentCache(X)
                    If oCache Is Nothing = False Then
                        If oCache.bNeedsSaved = True Then
                            mlPrevComponentCacheIdx = X
                            If oCache.SaveObject() = False Then
                                oCache.DeleteMe()
                                glComponentCacheIdx(X) = -1
                            End If
                            bFound = True
                            Exit For
                        End If
                    End If
                End If
            Next X
            If bFound = False Then mlPrevComponentCacheIdx = -1
        Catch
            mlPrevComponentCacheIdx += 1
        End Try

        'AGENT ASYNC =============================
        Try
            Dim bFound As Boolean = False
            For X As Int32 = mlPrevAgentIdx + 1 To glAgentUB
                If glAgentIdx(X) <> -1 Then
                    Dim oAgent As Agent = goAgent(X)
                    If oAgent Is Nothing = False Then
                        mlPrevAgentIdx = X
                        bFound = True
                        oAgent.SaveObject()
                        Exit For
                    End If
                End If
            Next X
            If bFound = False Then mlPrevAgentIdx = -1
        Catch
            mlPrevAgentIdx += 1
        End Try

        'MISSION ASYNC ============================
        Try
            Dim bFound As Boolean = False
            For X As Int32 = mlPrevMissionIdx + 1 To glPlayerMissionUB
                If glPlayerMissionIdx(X) > -1 Then
                    Dim oPM As PlayerMission = goPlayerMission(X)
                    If oPM Is Nothing = False Then
                        If oPM.lCurrentPhase <> oPM.lPreviousAsyncPhase Then
                            oPM.lPreviousAsyncPhase = oPM.lCurrentPhase
                            mlPrevMissionIdx = X
                            oPM.SaveObject()
                            bFound = True
                            Exit For
                        End If
                    End If
                End If
            Next X
            If bFound = False Then mlPrevMissionIdx = -1
        Catch
            mlPrevMissionIdx += 1
        End Try

        'GUILD ASYNC ===========================
        Try
            Dim bFound As Boolean = False
            For X As Int32 = mlPrevGuildIdx + 1 To glGuildUB
                If glGuildIdx(X) > -1 Then
                    Dim oGuild As Guild = goGuild(X)
                    If oGuild Is Nothing = False Then
                        mlPrevGuildIdx = X
                        oGuild.SaveObject()
                        bFound = True
                        Exit For
                    End If
                End If
            Next X
            If bFound = False Then mlPrevGuildIdx = -1
        Catch
            mlPrevGuildIdx += 1
        End Try

    End Sub

	Private Sub ExecuteSQL(ByVal sSQL As String)
		Dim oComm As OleDb.OleDbCommand = Nothing
		Try
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
		Catch ex As Exception
			LogEvent(LogEventType.CriticalError, "Async SQL Error (" & sSQL & "): " & ex.Message)
		Finally
			If oComm Is Nothing = False Then oComm.Dispose()
			oComm = Nothing
		End Try
	End Sub

	Private Sub DoSentimentCheck()
		'WarSentiment is a player wide effect that affects all colonies.
		'Desire to go to war is positive, Desire for peace is negative

		For X As Int32 = 0 To glPlayerUB
			If glPlayerIdx(X) <> -1 AndAlso goPlayer(X).InMyDomain = True AndAlso goPlayer(X).lStartLocX <> Int32.MinValue AndAlso goPlayer(X).lStartLocZ <> Int32.MinValue Then
				If goPlayer(X).PlayerRelUB > -1 Then
					If goPlayer(X).PlayerIsAtWar() = True Then
						'If any war, then subtract one.. people want peace
						goPlayer(X).lWarSentiment -= 1
						If Math.Abs(goPlayer(X).lWarSentiment) > 40 Then goPlayer(X).lWarSentiment = -40
					Else
						'If no war, then add one... peaople want war
						goPlayer(X).lWarSentiment += 1
						If Math.Abs(goPlayer(X).lWarSentiment) > 40 Then goPlayer(X).lWarSentiment = 40
					End If
				End If
			End If
		Next X

	End Sub

    'Private Sub HandleActOfWarTest(ByRef oED As Epica_Entity_Def, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByRef oCreator As Epica_Entity)
    '	If oED Is Nothing OrElse oCreator Is Nothing Then Return

    '	Dim oAlertPlayers() As Player = Nothing
    '	Dim lAlertPlayerUB As Int32 = -1
    '	Dim lMaxWpnRng As Int32
    '	Dim lParentID As Int32 = CType(oCreator.ParentObject, Epica_GUID).ObjectID
    '	Dim iParentTypeID As Int16 = CType(oCreator.ParentObject, Epica_GUID).ObjTypeID

    '	Dim bAlertWarDecPot As Boolean = False

    '	'get our max wpn range
    '	For Y As Int32 = 0 To oED.WeaponDefUB
    '		If oED.WeaponDefs(Y).oWeaponDef.Range > lMaxWpnRng Then lMaxWpnRng = oED.WeaponDefs(Y).oWeaponDef.Range
    '	Next Y

    '	'First, verify that no facilities will be a problem and we'll check entity def's wpn ranges too
    '	lMaxWpnRng *= gl_FINAL_GRID_SQUARE_SIZE
    '	Dim rcWpnRng As Rectangle = Rectangle.FromLTRB(lLocX - lMaxWpnRng, lLocZ - lMaxWpnRng, lLocX + lMaxWpnRng, lLocZ + lMaxWpnRng)
    '	For Y As Int32 = 0 To glFacilityUB
    '		If glFacilityIdx(Y) <> -1 Then
    '			Dim oFac As Facility = goFacility(Y)
    '			If oFac Is Nothing = False AndAlso oFac.ParentObject Is Nothing = False Then
    '				With CType(oFac.ParentObject, Epica_GUID)
    '					If .ObjectID <> lParentID OrElse .ObjTypeID <> iParentTypeID Then Continue For
    '				End With
    '				Dim rcTest As Rectangle = goModelDefs.GetModelDef(oFac.EntityDef.ModelID).rcSpacing
    '				rcTest.Location = New Point(oFac.LocX - (rcTest.Width \ 2), oFac.LocZ - (rcTest.Height \ 2))
    '				If oFac.Owner.ObjectID <> oCreator.Owner.ObjectID AndAlso rcWpnRng.IntersectsWith(rcTest) = True Then
    '					Dim yRelScore As Byte = oFac.Owner.GetPlayerRelScore(oCreator.Owner.ObjectID)
    '					If yRelScore < elRelTypes.ePeace AndAlso yRelScore > elRelTypes.eWar Then
    '						bAlertWarDecPot = True
    '						Dim bFoundPlayer As Boolean = False
    '						For Z As Int32 = 0 To lAlertPlayerUB
    '							If oAlertPlayers(Z) Is Nothing = False AndAlso oAlertPlayers(Z).ObjectID = oFac.Owner.ObjectID Then
    '								bFoundPlayer = True
    '								Exit For
    '							End If
    '						Next Z
    '						If bFoundPlayer = False Then
    '							lAlertPlayerUB += 1
    '							ReDim Preserve oAlertPlayers(lAlertPlayerUB)
    '							oAlertPlayers(lAlertPlayerUB) = oFac.Owner
    '						End If
    '					End If
    '				End If
    '			End If
    '		End If
    '	Next Y

    '	If bAlertWarDecPot = True Then
    '		'Ok, we need to alert our players
    '		Dim sBodyText As String = oCreator.Owner.sPlayerNameProper & " is building a facility of war that will have weapons in range of your facilities."

    '		If iParentTypeID = ObjectType.ePlanet Then
    '			sBodyText &= vbCrLf & vbCrLf & "Location: " & BytesToString(CType(oCreator.ParentObject, Planet).PlanetName)
    '		ElseIf iParentTypeID = ObjectType.eSolarSystem Then
    '			sBodyText &= vbCrLf & vbCrLf & "Location: " & BytesToString(CType(oCreator.ParentObject, SolarSystem).SystemName)
    '		End If

    '		For Y As Int32 = 0 To lAlertPlayerUB
    '			If oAlertPlayers(Y) Is Nothing = False Then
    '				Dim oPC As PlayerComm = oAlertPlayers(Y).AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, _
    '				  0, sBodyText, "Potential Threat Detected", oAlertPlayers(Y).ObjectID, GetDateAsNumber(Now), False, _
    '				  oAlertPlayers(Y).sPlayerNameProper)
    '				If oPC Is Nothing = False Then
    '					oPC.AddEmailAttachment(1, lParentID, iParentTypeID, lLocX, lLocZ, "Location")

    '					If oAlertPlayers(Y).oSocket Is Nothing = False Then
    '						If oAlertPlayers(Y).bInPlayerRequestDetails = False Then oAlertPlayers(Y).oSocket.SendData(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand))
    '					Else : goMsgSys.SendOutboundEmail(oPC, oAlertPlayers(Y), GlobalMessageCode.eActOfWarNotice, 1, _
    '					  oCreator.Owner.ObjectID, 0, 0, 0, 0, vbCrLf & "How do you wish to proceed?" & vbCrLf & _
    '					  "Set Relationship to <enter value>")	  '1 indicates that facility placement is cause
    '					End If
    '				End If
    '			End If
    '		Next Y
    '	End If

    'End Sub

	Private Sub CheckRetestPlayerLinks()
		For lIdx As Int32 = 0 To glPlayerUB
			If glPlayerIdx(lIdx) > 0 Then
                With goPlayer(lIdx).oSpecials
                    Dim lActiveLinks As Int32 = 0
                    Dim lMaxLinks As Int32 = .yMaxLinks
                    Dim lAddLinkChance As Int32 = 100 \ .yMultipleLinkChance
                    For X As Int32 = 0 To .mlSpecialTechUB
                        If .moSpecialTech(X).bLinked = True AndAlso .moSpecialTech(X).ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eResearched _
                            AndAlso .moSpecialTech(X).LinkAttempts < .moSpecialTech(X).oTech.MaxLinkChanceAttempts Then lActiveLinks += 1
                    Next X
                    If lActiveLinks >= lMaxLinks Then Continue For

                    'Ok, we're here, so force a perform link test
                    .PerformLinkTest()
                End With
			End If
		Next lIdx
	End Sub

    Private Sub RandomPirateOp()
        'ok, we pick a random unit or facility...
        Try
            Dim oRandom As New Random(Now.Second * Now.Minute * Now.Hour * Now.Day)

            Dim oTarget As Epica_Entity = Nothing
            Dim lHullSize As Int32 = 0
            Dim lCrew As Int32 = 0

            'If oRandom.Next(0, 100) < 50 Then
            'unit
            Dim lIdx As Int32 = -1
            For X As Int32 = 0 To 100
                lIdx = oRandom.Next(0, glUnitUB)
                If glUnitIdx(lIdx) > -1 Then

                    oTarget = goUnit(lIdx)
                    lHullSize = goUnit(lIdx).EntityDef.HullSize
                    If lHullSize < 1000 Then Continue For

                    If goUnit(lIdx).EntityDef.oPrototype Is Nothing = False Then
                        lCrew = goUnit(lIdx).EntityDef.oPrototype.GetTotalCrew()
                    End If

                    Exit For
                End If
            Next X 
 

            If oTarget Is Nothing Then Return

            'Ok, let's do this...
            If lCrew > 0 Then
                If lHullSize > 0 Then
                    Dim lRoll As Int32 = oRandom.Next(0, 100)
                    Dim lToHit As Int32 = CInt((lCrew / (lHullSize \ 100)) * 100)
                    If lRoll > lToHit Then
                        'sabotage
                        Dim uAttach() As PlayerComm.WPAttachment = Nothing
                        Dim lID As Int32 = -1
                        Dim iTypeID As Int16 = -1
                        With CType(oTarget.ParentObject, Epica_GUID)
                            lID = .ObjectID
                            iTypeID = .ObjTypeID
                            If .ObjTypeID = ObjectType.eFacility Then
                                With CType(CType(oTarget.ParentObject, Facility).ParentObject, Epica_GUID)
                                    lID = .ObjectID
                                    iTypeID = .ObjTypeID
                                End With
                            End If
                        End With
                        If lID > -1 AndAlso iTypeID > -1 Then
                            ReDim uAttach(0)
                            With uAttach(0)
                                .AttachNumber = 0
                                .EnvirID = lID
                                .EnvirTypeID = iTypeID
                                .LocX = oTarget.LocX
                                .LocZ = oTarget.LocZ
                                .sWPName = "Location"
                            End With
                        End If

                        LogEvent(LogEventType.Informational, "Pirate Raid on " & BytesToString(oTarget.EntityName) & ". OwnerID: " & oTarget.Owner.ObjectID & ". ParentID: " & lID & ", " & iTypeID & ", ToHit: " & lToHit & ", Roll: " & lRoll)

                        Dim oPC As PlayerComm = oTarget.Owner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                         BytesToString(oTarget.EntityName) & " has been destroyed by pirates. It appears the pirates were able to overrun our crew and security personnel within the unit and then planted explosives on it. The unit was completely destroyed. We may need to consider placing additional crew on board future designs.", _
                         "Pirate Raid", oTarget.Owner.ObjectID, GetDateAsNumber(Now), False, oTarget.Owner.sPlayerNameProper, uAttach)
                        If oPC Is Nothing = False Then
                            oTarget.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                        End If
                        DestroyEntity(oTarget, True, gl_HARDCODE_PIRATE_PLAYER_ID, True, "RandomPirateOp")
                    End If
                End If
            End If


        Catch
        End Try
    End Sub

End Module
