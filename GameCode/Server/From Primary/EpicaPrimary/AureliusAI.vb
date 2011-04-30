Option Strict On

Public Class AureliusAI
	Public Shared AureliusPlayerID As Int32 = 6
	Private Const ml_AI_PROCESS_INTERVAL As Int32 = 300

	Private Const l_JEEP_DEF_ID As Int32 = 19
	Private Const l_SM_TANK_DEF_ID As Int32 = 20
	Private Const l_LG_TANK_DEF_ID As Int32 = 21
	Private Const l_BLACK_HORNET_DEF_ID As Int32 = 2562
    Private Const l_PLUNDERER_DEF_ID As Int32 = 2563
    Private Const l_FURY_DEF_ID As Int32 = 2564
    Private Const l_SMITER_DEF_ID As Int32 = 2565

	Private Class AureliusAIEnvir

		Public Enum eyFacilityProductionCapable As Byte
			eFactory = 1
			eSpaceport = 2
			eCommandCenter = 4
		End Enum

		Public oObject As Object		'pointer to the geography class object for this envir
		Public lEnvirID As Int32		'ID of the envir
		Public iTypeID As Int16			'typeid of this envir

        Private mlAIFacIdx() As Int32       'indices into the global array of facilities
		Private mlAIFacID() As Int32		'ID of the facility (for verification purposes)
		Private mlAIFacProdID() As Int32	'ID of the Unit def being produced at this facility
		Private mlAIFacProdEnd() As Int32	'cycle when the production will be finished
		Private mlAIFacUB As Int32 = -1

		Private mlAIUnitIdx() As Int32		'indices into the global array of units
		Private mlAIUnitID() As Int32		'ID of the unit (for verification purposes)
		Private mlAIUnitProdID() As Int32	'ID OF the facility def being produced by this unit
		Private mlAIUnitProdEnd() As Int32	'cycle when the production will be finished
		Private mlAIUnitProdX() As Int32	'locx to produce
		Private mlAIUnitProdZ() As Int32	'locz to produce
		Private mlAIUnitUB As Int32 = -1

		Private mlTargetID As Int32 = -1		'current target Facility ID of our forces in this envir
		Private mlTargetIdx As Int32 = -1		'current target facility global index that our forces target in this envir
		Private mlTargetX As Int32 = Int32.MinValue			'current LocX of the facility being targeted
		Private mlTargetZ As Int32 = Int32.MinValue			'current LocZ of the facility being targeted
		Private mlLastSetTargetSent As Int32

		Private mlBaseCenterX As Int32 = Int32.MinValue
		Private mlBaseCenterZ As Int32 = Int32.MinValue

		Private mbPlacementRequestSent As Boolean = False

		Private mlColonizerIdx() As Int32		'indices into the global array of units
		Private mlColonizerID() As Int32		'ID of the unit (for verification purposes)
		Private mlColonizerUB As Int32 = -1
		Private mbColonizerMoveOrderSent As Boolean = False

		Private mlWaveNum As Int32 = 0

		Private mbReset As Boolean = False

		Public Sub AddAIEntity(ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16, ByVal lEntityIdx As Int32)
			If iEntityTypeID = ObjectType.eUnit Then
				If mlAIUnitIdx Is Nothing Then ReDim mlAIUnitIdx(-1)
				SyncLock mlAIUnitIdx
					Dim lIdx As Int32 = -1
					For X As Int32 = 0 To mlAIUnitUB
						If mlAIUnitID(X) = lEntityID Then
							'already here, so return
							Return
						ElseIf lIdx = -1 AndAlso mlAIUnitIdx(X) = -1 Then
							lIdx = X
						End If
					Next X
					If lIdx = -1 Then
						Dim lunitub As Int32 = mlAIUnitUB + 1
						ReDim Preserve mlAIUnitIdx(lunitub)
						ReDim Preserve mlAIUnitID(lunitub)
						ReDim Preserve mlAIUnitProdID(lunitub)
						ReDim Preserve mlAIUnitProdEnd(lunitub)
						ReDim Preserve mlAIUnitProdX(lunitub)
						ReDim Preserve mlAIUnitProdZ(lunitub)
						mlAIUnitIdx(lunitub) = -1
						mlAIUnitID(lunitub) = -1
						mlAIUnitProdID(lunitub) = -1
						mlAIUnitProdEnd(lunitub) = -1
						mlAIUnitProdX(lunitub) = 0
						mlAIUnitProdZ(lunitub) = 0
						lIdx = lunitub
						mlAIUnitUB = lunitub
					End If
					mlAIUnitIdx(lIdx) = lEntityIdx
					mlAIUnitID(lIdx) = lEntityID
				End SyncLock

			ElseIf iEntityTypeID = ObjectType.eFacility Then

				If mlAIFacIdx Is Nothing Then ReDim mlAIFacIdx(-1)
				SyncLock mlAIFacIdx
					Dim lIdx As Int32 = -1
					For X As Int32 = 0 To mlAIFacUB
						If mlAIFacID(X) = lEntityID Then
							'already here, so return
							Return
						ElseIf lIdx = -1 AndAlso mlAIFacIdx(X) = -1 Then
							lIdx = X
						End If
					Next X
					If lIdx = -1 Then
						Dim lfacub As Int32 = mlAIFacUB + 1
						ReDim Preserve mlAIFacIdx(lfacub)
						ReDim Preserve mlAIFacID(lfacub)
						ReDim Preserve mlAIFacProdID(lfacub)
						ReDim Preserve mlAIFacProdEnd(lfacub)
						mlAIFacIdx(lfacub) = -1
						mlAIFacID(lfacub) = -1
						mlAIFacProdID(lfacub) = -1
						mlAIFacProdEnd(lfacub) = -1
						lIdx = lfacub
						mlAIFacUB = lfacub
					End If
					mlAIFacIdx(lIdx) = lEntityIdx
					mlAIFacID(lIdx) = lEntityID
				End SyncLock
			End If
		End Sub

		Public Sub RemoveAIEntity(ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16)
			If iEntityTypeID = ObjectType.eUnit Then
				For X As Int32 = 0 To mlAIUnitUB
					If mlAIUnitID(X) = lEntityID Then
						mlAIUnitID(X) = -1
						mlAIUnitIdx(X) = -1
						Return
					End If
				Next X
			ElseIf iEntityTypeID = ObjectType.eFacility Then
				For X As Int32 = 0 To mlAIFacUB
					If mlAIFacID(X) = lEntityID Then
						mlAIFacID(X) = -1
						mlAIFacIdx(X) = -1
						Return
					End If
				Next X
			End If
		End Sub

		Public Function PlayerLocatedHere() As Boolean
			If iTypeID = ObjectType.ePlanet Then
				'ok, is our object nothing
				If Me.oObject Is Nothing = False Then
					'ok, ctype it
					Dim oPlanet As Planet = CType(oObject, Planet)
					Dim lCurUB As Int32 = -1
					If oPlanet.lColonysHereIdx Is Nothing = False Then lCurUB = Math.Min(oPlanet.lColonysHereUB, oPlanet.lColonysHereIdx.GetUpperBound(0))
					For X As Int32 = 0 To lCurUB
						Dim lColonyIdx As Int32 = oPlanet.lColonysHereIdx(X)
						If lColonyIdx > -1 AndAlso lColonyIdx <= glColonyUB Then
							'ok, a colony is here... 
							Dim oColony As Colony = goColony(lColonyIdx)
                            If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID <> AureliusPlayerID AndAlso oColony.Owner.yIronCurtainState <> eIronCurtainState.IronCurtainIsUpOnEverything AndAlso (oColony.Owner.yIronCurtainState <> eIronCurtainState.IronCurtainIsUpOnSelectedPlanet OrElse oColony.Owner.lIronCurtainPlanet <> Me.lEnvirID) Then
                                mbReset = False
                                Return True
                            End If
						End If
					Next X
				End If
			End If
			Return False
		End Function

		Public Function AIAssetsHere() As Boolean
			Dim lCurUB As Int32 = -1
			If mlAIUnitIdx Is Nothing = False Then lCurUB = Math.Min(mlAIUnitUB, mlAIUnitIdx.GetUpperBound(0))
			For X As Int32 = 0 To lCurUB
				If mlAIUnitIdx(X) > -1 Then
					If glUnitIdx(mlAIUnitIdx(X)) = mlAIUnitID(X) Then
						Dim oUnit As Epica_Entity = goUnit(mlAIUnitIdx(X))
						If oUnit Is Nothing = False Then
							If oUnit.yProductionType = 0 Then Return True
						End If
					End If
				End If
			Next X
			Return False
		End Function

		Public Sub SetAttackTarget()

			'Do we already have a target?
			If mlTargetIdx > -1 Then
				If glFacilityIdx(mlTargetIdx) = mlTargetID Then
					'Check if we need to resend order to attack...
					If glCurrentCycle - mlLastSetTargetSent > 1800 Then		'1 minute
						SendAttackOrder()
					End If

					Return
				End If
			End If

			'No, so find one
			Dim lCurrIdx As Int32 = -1
			Dim lCurrID As Int32 = -1
			Dim lCurrX As Int32 = 0
			Dim lCurrZ As Int32 = 0
			Dim lScore As Int32 = 0

			If iTypeID = ObjectType.ePlanet Then
				Dim oPlanet As Planet = CType(oObject, Planet)
				If oPlanet Is Nothing = False Then
					'Ok, go thru the colonies...
					Dim lCurUB As Int32 = -1
					If oPlanet.lColonysHereIdx Is Nothing = False Then lCurUB = Math.Min(oPlanet.lColonysHereUB, oPlanet.lColonysHereIdx.GetUpperBound(0))
					For X As Int32 = 0 To lCurUB
						Dim lColonyIdx As Int32 = oPlanet.lColonysHereIdx(X)
						If lColonyIdx > -1 Then
							Dim oColony As Colony = goColony(lColonyIdx)
                            If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID <> gl_HARDCODE_PIRATE_PLAYER_ID Then
                                If CType(oColony.ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet AndAlso CType(oColony.ParentObject, Epica_GUID).ObjectID = oPlanet.ObjectID Then
                                    'ok, valid colony
                                    Dim lFacUB As Int32 = -1
                                    If oColony.lChildrenIdx Is Nothing = False Then lFacUB = Math.Min(oColony.lChildrenIdx.GetUpperBound(0), oColony.ChildrenUB)
                                    For Y As Int32 = 0 To lFacUB
                                        If oColony.lChildrenIdx(Y) <> -1 Then
                                            Dim oFac As Facility = oColony.oChildren(Y)
                                            If oFac Is Nothing = False Then
                                                'ok, now, determine score... 
                                                Dim lTmpScore As Int32 = 0
                                                If mlTargetX <> Int32.MinValue AndAlso mlTargetZ <> Int32.MinValue Then
                                                    'ok, use distance
                                                    Dim lX As Int32 = Math.Abs(oFac.LocX - mlTargetX)
                                                    Dim lZ As Int32 = Math.Abs(oFac.LocZ - mlTargetZ)
                                                    lTmpScore = 500000 - (lX + lZ)
                                                Else
                                                    'ok, use importance
                                                    Select Case oFac.yProductionType
                                                        Case ProductionType.eCommandCenterSpecial
                                                            lTmpScore = 10000
                                                        Case ProductionType.eProduction, ProductionType.eAerialProduction
                                                            lTmpScore = 8000
                                                        Case ProductionType.ePowerCenter
                                                            lTmpScore = 6000
                                                        Case Else
                                                            lTmpScore = 5000
                                                    End Select
                                                End If

                                                If lTmpScore > lScore Then
                                                    lScore = lTmpScore
                                                    lCurrIdx = oFac.ServerIndex
                                                    lCurrID = oFac.ObjectID
                                                    lCurrX = oFac.LocX
                                                    lCurrZ = oFac.LocZ
                                                End If

                                            End If
                                        End If
                                    Next Y

                                End If
                            End If
						End If
					Next X

					'Ok, we are here... check our results
					mlTargetIdx = lCurrIdx
					mlTargetID = lCurrID
					mlTargetX = lCurrX
					mlTargetZ = lCurrZ
					If lCurrIdx <> -1 Then
						SendAttackOrder()
					End If
				End If
			End If
		End Sub

		Private Sub SendAttackOrder()
			'Ok, give our units an order
			Dim lCnt As Int32 = 0
			Dim lCurUB As Int32 = -1
			If mlAIUnitIdx Is Nothing = False Then lCurUB = Math.Min(mlAIUnitIdx.GetUpperBound(0), mlAIUnitUB)
			Dim lIDs(lCurUB) As Int32
			For X As Int32 = 0 To lCurUB
				lIDs(X) = -1
				If mlAIUnitIdx(X) > -1 AndAlso mlAIUnitID(X) = glUnitIdx(mlAIUnitIdx(X)) Then
					lIDs(X) = mlAIUnitID(X)
					lCnt += 1
				End If
			Next X
			If lCnt > 0 Then
				Dim yMsg(7 + (lCnt * 6)) As Byte
				Dim lPos As Int32 = 0
				'Msg Code (2), TargetID (4), TargetTypeID (2), GUIDList...
				System.BitConverter.GetBytes(GlobalMessageCode.eSetPrimaryTarget).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(mlTargetID).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(ObjectType.eFacility).CopyTo(yMsg, lPos) : lPos += 2
				For X As Int32 = 0 To lCurUB
					If lIDs(X) <> -1 Then
						System.BitConverter.GetBytes(lIDs(X)).CopyTo(yMsg, lPos) : lPos += 4
						System.BitConverter.GetBytes(ObjectType.eUnit).CopyTo(yMsg, lPos) : lPos += 2
					End If
				Next X

				Dim oPlanet As Planet = CType(Me.oObject, Planet)
				If oPlanet Is Nothing = False Then
					oPlanet.oDomain.DomainSocket.SendData(yMsg)
				End If

				mlLastSetTargetSent = glCurrentCycle
			End If
		End Sub

		Public Function GetFacilityProductionCapability() As Byte
			'ok, this requires a Factory or a spaceport
			Dim yResult As Byte = 0

			Dim lCurUB As Int32 = -1
			If mlAIFacIdx Is Nothing = False Then lCurUB = Math.Min(mlAIFacUB, mlAIFacIdx.GetUpperBound(0))
			For X As Int32 = 0 To lCurUB
				If mlAIFacIdx(X) > -1 Then
					If glFacilityIdx(mlAIFacIdx(X)) = mlAIFacID(X) Then
						Dim oFac As Facility = goFacility(mlAIFacIdx(X))
						If oFac Is Nothing = False Then
							If oFac.yProductionType = ProductionType.eAerialProduction Then
								yResult = yResult Or eyFacilityProductionCapable.eSpaceport
							ElseIf oFac.yProductionType = ProductionType.eCommandCenterSpecial Then
								yResult = yResult Or eyFacilityProductionCapable.eCommandCenter
							ElseIf oFac.yProductionType = ProductionType.eProduction Then
								yResult = yResult Or eyFacilityProductionCapable.eFactory
							End If
						End If
					Else
						mlAIFacIdx(X) = -1
					End If
				End If
			Next X

			Return yResult
		End Function

		Public Function EngineerOnSite() As Byte
			Dim lCurUB As Int32 = -1
			Dim yResult As Byte = 0
			If mlAIUnitIdx Is Nothing = False Then lCurUB = Math.Min(mlAIUnitUB, mlAIUnitIdx.GetUpperBound(0))
			For X As Int32 = 0 To lCurUB
				If mlAIUnitIdx(X) > -1 Then
					If glUnitIdx(mlAIUnitIdx(X)) = mlAIUnitID(X) Then
						Dim oUnit As Epica_Entity = goUnit(mlAIUnitIdx(X))
						If oUnit Is Nothing = False Then
							If oUnit.yProductionType = ProductionType.eFacility Then
								If oUnit.bProducing = False Then

									If mlAIUnitProdID(X) = -1 Then
										Return 2
									Else : yResult = 1
									End If

								End If
							End If
						End If
					Else
						mlAIUnitIdx(X) = -1
					End If
				End If
			Next X
			Return yResult
		End Function

		Public Function ColonizerForceArrived() As Boolean
			Dim bResult As Boolean = False

			Dim lCurUB As Int32 = -1
			If mlColonizerIdx Is Nothing = False Then lCurUB = Math.Min(mlColonizerUB, mlColonizerIdx.GetUpperBound(0))
			For X As Int32 = 0 To lCurUB
				If mlColonizerIdx(X) > -1 Then
					If mlColonizerID(X) = glUnitIdx(mlColonizerIdx(X)) Then
						Dim oUnit As Unit = goUnit(mlColonizerIdx(X))

						If oUnit Is Nothing = False Then
							Dim oParent As Epica_GUID = CType(oUnit.ParentObject, Epica_GUID)
							If oParent Is Nothing = False Then
								If oParent.ObjectID = Me.lEnvirID AndAlso oParent.ObjTypeID = Me.iTypeID Then
									If oUnit.yProductionType = ProductionType.eFacility Then
										'Ok, check the unit's parent
										bResult = True
									End If
									AddAIEntity(mlColonizerID(X), ObjectType.eUnit, mlColonizerIdx(X))
									mlColonizerIdx(X) = -1
								End If
							End If
						End If
					End If
				End If
			Next X
			Return bResult
		End Function

		Public Function ColonizerForceEnRoute() As Boolean
			Dim lCurUB As Int32 = -1

			If mbColonizerMoveOrderSent = False Then
				'Determine Dest Location
				If CheckBaseCenterLoc() = True Then
					lCurUB = -1
					If mlColonizerIdx Is Nothing = False Then lCurUB = Math.Min(mlColonizerIdx.GetUpperBound(0), mlColonizerUB)

					If lCurUB <> -1 Then
						'Send Colonizer Force to target... probably should do formation... can we guarantee it will always work?
						Dim lIDs(lCurUB) As Int32
						Dim lCnt As Int32 = 0
						For X As Int32 = 0 To lCurUB
							lIDs(X) = -1
							If mlColonizerIdx(X) > -1 AndAlso mlColonizerID(X) = glUnitIdx(mlColonizerIdx(X)) Then
								lCnt += 1
								lIDs(X) = mlColonizerID(X)
							End If
						Next X

						Dim yMsg(21 + (lCnt * 6)) As Byte
						Dim lPos As Int32 = 0
						System.BitConverter.GetBytes(GlobalMessageCode.eForcedMoveSpeedMove).CopyTo(yMsg, lPos) : lPos += 2
						System.BitConverter.GetBytes(mlBaseCenterX).CopyTo(yMsg, lPos) : lPos += 4
						System.BitConverter.GetBytes(mlBaseCenterZ).CopyTo(yMsg, lPos) : lPos += 4
						System.BitConverter.GetBytes(0S).CopyTo(yMsg, lPos) : lPos += 2
						System.BitConverter.GetBytes(Me.lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
						System.BitConverter.GetBytes(Me.iTypeID).CopyTo(yMsg, lPos) : lPos += 2
						System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4
						For X As Int32 = 0 To lCurUB
							If lIDs(X) <> -1 Then
								System.BitConverter.GetBytes(lIDs(X)).CopyTo(yMsg, lPos) : lPos += 4
								System.BitConverter.GetBytes(ObjectType.eUnit).CopyTo(yMsg, lPos) : lPos += 2
							End If
						Next X
						Dim oPlanet As Planet = CType(Me.oObject, Planet)
						If oPlanet Is Nothing Then Return True
						oPlanet.oDomain.DomainSocket.SendData(yMsg)

						'Mark State Colonizer Force En Route
						mbColonizerMoveOrderSent = True

						Return True
					End If
				Else : Return True		'return true for now
				End If
			End If

			lCurUB = -1
			If mlColonizerIdx Is Nothing = False Then lCurUB = Math.Min(mlColonizerUB, mlColonizerIdx.GetUpperBound(0))
			For X As Int32 = 0 To lCurUB
				If mlColonizerIdx(X) > -1 Then
					If mlColonizerID(X) = glUnitIdx(mlColonizerIdx(X)) Then
						Dim oUnit As Unit = goUnit(mlColonizerIdx(X))
						If oUnit Is Nothing = False Then
							If oUnit.yProductionType = ProductionType.eFacility Then
								'Ok, got an engineer it *should* be en route or will be soon
								Return True
							End If
						End If
					End If
				End If
			Next X

			mbColonizerMoveOrderSent = False
			Return False
		End Function

		Public Sub AddEngineerToProdQueueOfCC()
			Dim lCurUB As Int32 = -1
			If mlAIFacIdx Is Nothing = False Then lCurUB = Math.Min(mlAIFacUB, mlAIFacIdx.GetUpperBound(0))
			For X As Int32 = 0 To lCurUB
				If mlAIFacIdx(X) > -1 Then
					If glFacilityIdx(mlAIFacIdx(X)) = mlAIFacID(X) Then
						Dim oFac As Facility = goFacility(mlAIFacIdx(X))
						If oFac Is Nothing = False Then
							If oFac.yProductionType = ProductionType.eCommandCenterSpecial Then
								'Ok, do it
								If mlAIFacProdID(X) = -1 Then
									mlAIFacProdID(X) = 30
									mlAIFacProdEnd(X) = glCurrentCycle + 120
								End If
								Return
							End If
						End If
					Else
						mlAIFacIdx(X) = -1
					End If
				End If
			Next X
		End Sub

		Public Sub BuildNextFacility()
			Dim lCC As Int32 = 0
			Dim lPowerGen As Int32 = 0
			Dim lFactory As Int32 = 0
			Dim lSpacePort As Int32 = 0
			Dim lTurret As Int32 = 0

			If mbPlacementRequestSent = True Then Return

			If mlBaseCenterX = Int32.MinValue OrElse mlBaseCenterZ = Int32.MinValue Then
				mlBaseCenterX = Int32.MaxValue
				mlBaseCenterZ = Int32.MaxValue
				'Request the start loc... should have been done already...
				Dim yMsg(17) As Byte
				Dim lPos As Int32 = 0
				System.BitConverter.GetBytes(GlobalMessageCode.eGetPirateStartLoc).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(Me.lEnvirID).CopyTo(yMsg, 2) : lPos += 4
				System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 6) : lPos += 4
				System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 10) : lPos += 4
				System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 14) : lPos += 4
				Dim oPlanet As Planet = CType(Me.oObject, Planet)
				If oPlanet Is Nothing = False Then
					oPlanet.oDomain.DomainSocket.SendData(yMsg)
				End If
			ElseIf mlBaseCenterX = Int32.MaxValue OrElse mlBaseCenterZ = Int32.MaxValue Then
				'waiting on a response from region
				Return
			End If

			Dim lCurUB As Int32 = -1
			If mlAIFacIdx Is Nothing = False Then lCurUB = Math.Min(mlAIFacUB, mlAIFacIdx.GetUpperBound(0))
			For X As Int32 = 0 To lCurUB
				If mlAIFacIdx(X) > -1 Then
					If glFacilityIdx(mlAIFacIdx(X)) = mlAIFacID(X) Then
						Dim oFac As Facility = goFacility(mlAIFacIdx(X))
						If oFac Is Nothing = False Then
							If oFac.yProductionType = ProductionType.eCommandCenterSpecial Then
								lCC += 1
							ElseIf oFac.yProductionType = ProductionType.eProduction Then
								lFactory += 1
							ElseIf oFac.yProductionType = ProductionType.eAerialProduction Then
								lSpacePort += 1
                            ElseIf oFac.yProductionType = ProductionType.ePowerCenter OrElse oFac.EntityDef.ObjectID = 1558 Then
                                lPowerGen += 1
							Else
								lTurret += 1
							End If
						End If
					Else
						mlAIFacIdx(X) = -1
					End If
				End If
			Next X

			lCurUB = -1
			If mlAIUnitIdx Is Nothing = False Then lCurUB = Math.Min(mlAIUnitIdx.GetUpperBound(0), mlAIUnitUB)
			For X As Int32 = 0 To lCurUB
				If mlAIUnitIdx(X) > -1 Then
					If glUnitIdx(mlAIUnitIdx(X)) = mlAIUnitID(X) AndAlso mlAIUnitProdID(X) > 0 Then
						If mlAIUnitProdID(X) = 5 Then
							lCC += 1
						ElseIf mlAIUnitProdID(X) = 1558 Then
							lPowerGen += 1
						ElseIf mlAIUnitProdID(X) = 49 Then
							lFactory += 1
						ElseIf mlAIUnitProdID(X) = 48 Then
							lSpacePort += 1
						ElseIf mlAIUnitProdID(X) = 46 Then
							lTurret += 1
						End If
					End If
				End If
			Next X

			'Now, check our counts...
			Dim lBuildDefID As Int32 = -1
			Dim iBuildDefTypeID As Int16 = -1 
			If lCC = 0 Then
				'Build a command center
				lBuildDefID = 5
				iBuildDefTypeID = ObjectType.eFacilityDef
			ElseIf lPowerGen < 2 Then
				'Build a power generator
				lBuildDefID = 1558
				iBuildDefTypeID = ObjectType.eFacilityDef
			ElseIf lFactory < 1 Then
				'build factory
				lBuildDefID = 49
				iBuildDefTypeID = ObjectType.eFacilityDef
			ElseIf lPowerGen < 4 Then
				'build a power generator
				lBuildDefID = 1558
				iBuildDefTypeID = ObjectType.eFacilityDef
			ElseIf lSpacePort < 1 Then
				'build a spaceport
				lBuildDefID = 48
				iBuildDefTypeID = ObjectType.eFacilityDef
			ElseIf lTurret < 8 Then
				'build a turret
				lBuildDefID = 46
				iBuildDefTypeID = ObjectType.eFacilityDef
			End If

			'Ok, we are here... so let's do this
			If lBuildDefID <> -1 AndAlso iBuildDefTypeID <> -1 Then
				Dim yMsg(35) As Byte
				Dim lPos As Int32 = 0
				Dim oPlanet As Planet = CType(Me.oObject, Planet)
				If oPlanet Is Nothing = False Then
					System.BitConverter.GetBytes(GlobalMessageCode.ePlacePirateAssets).CopyTo(yMsg, lPos) : lPos += 2
					oPlanet.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
					System.BitConverter.GetBytes(mlBaseCenterX).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(mlBaseCenterZ).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(1I).CopyTo(yMsg, lPos) : lPos += 4

					System.BitConverter.GetBytes(lBuildDefID).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(iBuildDefTypeID).CopyTo(yMsg, lPos) : lPos += 2
					System.BitConverter.GetBytes(lBuildDefID).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(iBuildDefTypeID).CopyTo(yMsg, lPos) : lPos += 2

					mbPlacementRequestSent = True
					oPlanet.oDomain.DomainSocket.SendData(yMsg)
				End If
			End If

		End Sub

		Public Sub ProcessProduction(ByVal bUnitProdOnly As Boolean)
			Dim lCurUB As Int32 = -1

			If bUnitProdOnly = False Then
				If mlAIFacIdx Is Nothing = False Then lCurUB = Math.Min(mlAIFacIdx.GetUpperBound(0), mlAIFacUB)
				For X As Int32 = 0 To lCurUB
					If mlAIFacIdx(X) <> -1 Then
						If mlAIFacID(X) = glFacilityIdx(mlAIFacIdx(X)) Then
							Dim oFac As Facility = goFacility(mlAIFacIdx(X))
							If oFac Is Nothing = False Then
								If mlAIFacProdID(X) = -1 Then
									'Not producing... add something to produce
									If oFac.yProductionType = ProductionType.eProduction Then
										'factory
										Dim lRoll As Int32 = CInt(Rnd() * 100)
										If lRoll < 30 Then
											mlAIFacProdID(X) = l_JEEP_DEF_ID
											mlAIFacProdEnd(X) = glCurrentCycle + 450  '15 seconds
										ElseIf lRoll < 60 Then
											mlAIFacProdID(X) = l_SM_TANK_DEF_ID
											mlAIFacProdEnd(X) = glCurrentCycle + 750	'25 seconds
										Else
											mlAIFacProdID(X) = l_LG_TANK_DEF_ID
											mlAIFacProdEnd(X) = glCurrentCycle + 1800	'1 minute
										End If
									ElseIf oFac.yProductionType = ProductionType.eAerialProduction Then
										'spaceport
										Dim lRoll As Int32 = CInt(Rnd() * 100)
										'spaceport: fighter, plund, fury, smiter
                                        'If lRoll < 35 Then
                                        mlAIFacProdID(X) = l_BLACK_HORNET_DEF_ID
                                        mlAIFacProdEnd(X) = glCurrentCycle + 1350   '45 seconds
                                        'ElseIf lRoll < 65 Then
                                        '	mlAIFacProdID(X) = l_PLUNDERER_DEF_ID
                                        '	mlAIFacProdEnd(X) = glCurrentCycle + 3600	'2 minutes
                                        'ElseIf lRoll < 85 Then
                                        '	mlAIFacProdID(X) = l_FURY_DEF_ID
                                        '	mlAIFacProdEnd(X) = glCurrentCycle + 5400	'3 minutes
                                        'Else
                                        '	mlAIFacProdID(X) = l_SMITER_DEF_ID
                                        '	mlAIFacProdEnd(X) = glCurrentCycle + 7200	'4 minutes
                                        'End If
									End If
								ElseIf mlAIFacProdEnd(X) < glCurrentCycle Then
									'Ok, done producing... finish it...

									Dim iCombat As eiBehaviorPatterns
									Dim iTarget As eiTacticalAttrs

									If mlAIFacProdID(X) = l_BLACK_HORNET_DEF_ID Then
										iCombat = eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Maximize_Damage Or eiBehaviorPatterns.eTactics_Maneuver
										iTarget = eiTacticalAttrs.eFighterClass Or eiTacticalAttrs.eArmedUnit
									Else
										iCombat = eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Maximize_Damage
										iTarget = eiTacticalAttrs.eFacility
									End If

									Dim lIdx As Int32 = SpawnNewUnit(mlAIFacProdID(X), iCombat, iTarget, oFac.LocX, oFac.LocZ, Me.oObject)

									If lIdx > -1 Then
										Dim oUnit As Epica_Entity = goUnit(lIdx)
										If oUnit Is Nothing = False Then
											oFac.SendUndockFirstWaypoint(oUnit)
											AddAIEntity(oUnit.ObjectID, ObjectType.eUnit, lIdx)
										End If
									End If

									mlAIFacProdID(X) = -1
									mlAIFacProdEnd(X) = Int32.MaxValue
								End If
							End If
						Else : mlAIFacIdx(X) = -1
						End If
					End If
				Next X
			End If

			lCurUB = -1
			If mlAIUnitIdx Is Nothing = False Then lCurUB = Math.Min(mlAIUnitIdx.GetUpperBound(0), mlAIUnitUB)
			For X As Int32 = 0 To lCurUB
				If mlAIUnitIdx(X) <> -1 Then
					If mlAIUnitID(X) = glUnitIdx(mlAIUnitIdx(X)) Then
						Dim oUnit As Unit = goUnit(mlAIUnitIdx(X))
						If oUnit Is Nothing = False Then
							If mlAIUnitProdID(X) <> -1 AndAlso mlAIUnitProdEnd(X) < glCurrentCycle Then
								'ok, done producing... finish it...
                                Dim lIdx As Int32 = SpawnNewFacility(mlAIUnitProdID(X), eiBehaviorPatterns.eEngagement_Stand_Ground Or eiBehaviorPatterns.eTactics_LaunchChildren Or eiBehaviorPatterns.eTactics_Normal, eiTacticalAttrs.eArmedUnit, mlAIUnitProdX(X), mlAIUnitProdZ(X), Me.oObject, True)
								If lIdx > -1 Then
									Dim oFac As Epica_Entity = goFacility(lIdx)
									If oFac Is Nothing = False Then
										AddAIEntity(oFac.ObjectID, ObjectType.eFacility, lIdx)
									End If
								End If
								'AddAIEntity(oUnit.ObjectID, ObjectType.eUnit, lIdx)
								mlAIUnitProdID(X) = -1
								mlAIUnitProdEnd(X) = Int32.MaxValue
							End If
						End If
					End If
				End If
			Next X
		End Sub

		Public Sub AddColonizerForceEntity(ByVal lEntityID As Int32, ByVal lEntityIdx As Int32)
			'colonizer force is always units
			If mlColonizerIdx Is Nothing Then ReDim mlColonizerIdx(-1)
			SyncLock mlColonizerIdx
				Dim lIdx As Int32 = -1
				For X As Int32 = 0 To mlColonizerUB
					If mlColonizerID(X) = lEntityID Then
						'already here, so return
						Return
					ElseIf lIdx = -1 AndAlso mlColonizerIdx(X) = -1 Then
						lIdx = X
					End If
				Next X
				If lIdx = -1 Then
					Dim lunitub As Int32 = mlColonizerUB + 1
					ReDim Preserve mlColonizerIdx(lunitub)
					ReDim Preserve mlColonizerID(lunitub)
					mlColonizerIdx(lunitub) = -1
					mlColonizerID(lunitub) = -1
					lIdx = lunitub
					mlColonizerUB = lunitub
				End If
				mlColonizerIdx(lIdx) = lEntityIdx
				mlColonizerID(lIdx) = lEntityID
			End SyncLock
		End Sub

		Public Sub ResetEnvir()
			'Ok, need to send all units to somewhere and remove production queues...
			If mbReset = False Then
				'Ok, tell all of our forces to return to the command center... the pathfinding stop event should spread them out
				'Determine position of command center
				Dim lCCX As Int32 = Int32.MinValue
				Dim lCCZ As Int32 = Int32.MinValue

				Dim lCurUB As Int32 = -1
				If mlAIFacIdx Is Nothing = False Then lCurUB = Math.Min(mlAIFacUB, mlAIFacIdx.GetUpperBound(0))
				For X As Int32 = 0 To lcurub
					If mlAIFacIdx(X) > -1 Then
						If mlAIFacID(X) = glFacilityIdx(mlAIFacIdx(X)) Then
							Dim oFac As Facility = goFacility(mlAIFacIdx(X))
							If oFac Is Nothing = False Then
								If oFac.yProductionType = ProductionType.eCommandCenterSpecial Then
									lCCX = oFac.LocX
									lCCZ = oFac.LocZ
								Else
									'clear all production queues except the CC (as it may be building an engineer)
									mlAIFacProdID(X) = -1
									mlAIFacProdEnd(X) = -1
								End If
							End If
						Else : mlAIFacIdx(X) = -1
						End If
					End If
				Next X

				'Wait for the colony to be built first
				If lCCX = Int32.MinValue OrElse lCCZ = Int32.MinValue Then Return

				Dim oPlanet As Planet = CType(Me.oObject, Planet)
				If oPlanet Is Nothing Then Return

				'Ok, we are here so the colony exists... send our units there now, let's go ahead and set our reset flag
				mbReset = True

				'Now, send all of our units there
				lCurUB = -1
				If mlAIUnitIdx Is Nothing = False Then lCurUB = Math.Min(mlAIUnitUB, mlAIUnitIdx.GetUpperBound(0))
				For X As Int32 = 0 To lCurUB
					If mlAIUnitIdx(X) > -1 Then
						If mlAIUnitID(X) = glUnitIdx(mlAIUnitIdx(X)) Then
							Dim oUnit As Unit = goUnit(mlAIUnitIdx(X))
							If oUnit Is Nothing = False Then
								'Now, send a destination of the unit
								Dim yMove(23) As Byte
								System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequest).CopyTo(yMove, 0)
								System.BitConverter.GetBytes(lCCX).CopyTo(yMove, 2)
								System.BitConverter.GetBytes(lCCZ).CopyTo(yMove, 6)
								System.BitConverter.GetBytes(0S).CopyTo(yMove, 10)
								System.BitConverter.GetBytes(Me.lEnvirID).CopyTo(yMove, 12)
								System.BitConverter.GetBytes(ObjectType.ePlanet).CopyTo(yMove, 16)
								oUnit.GetGUIDAsString.CopyTo(yMove, 18)
								oPlanet.oDomain.DomainSocket.SendData(yMove)
							End If
						Else : mlAIUnitIdx(X) = -1
						End If
					End If
				Next X

			End If
		End Sub

		Public Sub HandlePlacePirateAsset(ByVal lItemID As Int32, ByVal iItemTypeID As Int16, ByVal lLocX As Int32, ByVal lLocZ As Int32)
			'Now, find our unit and set it to build the facility
			Dim lCurUB As Int32 = -1
			If mlAIUnitIdx Is Nothing = False Then lCurUB = Math.Min(mlAIUnitIdx.GetUpperBound(0), mlAIUnitUB)
			For X As Int32 = 0 To lCurUB
				If mlAIUnitIdx(X) <> -1 Then
					If mlAIUnitID(X) = glUnitIdx(mlAIUnitIdx(X)) Then
						Dim oUnit As Unit = goUnit(mlAIUnitIdx(X))
						If oUnit Is Nothing = False AndAlso oUnit.yProductionType = ProductionType.eFacility Then
							If mlAIUnitProdID(X) = -1 Then
								mlAIUnitProdEnd(X) = Int32.MaxValue
								mlAIUnitProdID(X) = lItemID
								mlAIUnitProdX(X) = lLocX
								mlAIUnitProdZ(X) = lLocZ

								'Now, send our move command
								Dim yMsg(23) As Byte
								Dim lPos As Int32 = 0
								System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequest).CopyTo(yMsg, lPos) : lPos += 2
								System.BitConverter.GetBytes(lLocX).CopyTo(yMsg, lPos) : lPos += 4
								System.BitConverter.GetBytes(lLocZ).CopyTo(yMsg, lPos) : lPos += 4
								System.BitConverter.GetBytes(0S).CopyTo(yMsg, lPos) : lPos += 2
								System.BitConverter.GetBytes(Me.lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
								System.BitConverter.GetBytes(Me.iTypeID).CopyTo(yMsg, lPos) : lPos += 2
								oUnit.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6

								Dim oPlanet As Planet = CType(Me.oObject, Planet)
								If oPlanet Is Nothing = False Then
									oPlanet.oDomain.DomainSocket.SendData(yMsg)
								End If
								mbPlacementRequestSent = False
							End If
						End If
					End If
				End If
			Next X
		End Sub

		Public Sub EngineerReachedDest(ByVal lID As Int32)
			Dim lCurUB As Int32 = -1
			If mlAIUnitIdx Is Nothing = False Then lCurUB = Math.Min(mlAIUnitIdx.GetUpperBound(0), mlAIUnitUB)
			For X As Int32 = 0 To lCurUB
				If mlAIUnitIdx(X) <> -1 Then
					If mlAIUnitID(X) = lID AndAlso glUnitIdx(mlAIUnitIdx(X)) = lID Then
						'ok, the unit has made it... begin producing
						Select Case mlAIUnitProdID(X)
							Case 5
								mlAIUnitProdEnd(X) = glCurrentCycle + 1800	'1 minute
							Case 1558
								mlAIUnitProdEnd(X) = glCurrentCycle + 1800	'1 minute
							Case 49
								mlAIUnitProdEnd(X) = glCurrentCycle + 3600	'2 minutes
							Case 48
								mlAIUnitProdEnd(X) = glCurrentCycle + 5400	'3 minutes
							Case 46
								mlAIUnitProdEnd(X) = glCurrentCycle + 450	'15 seconds
						End Select
						Return
					End If
				End If
			Next X

		End Sub

		Public Function CheckBaseCenterLoc() As Boolean
			If mlBaseCenterX = Int32.MinValue OrElse mlBaseCenterZ = Int32.MinValue Then
				mlBaseCenterX = Int32.MaxValue
				mlBaseCenterZ = Int32.MaxValue
				'Request the start loc... should have been done already...
				Dim yMsg(17) As Byte
				Dim lPos As Int32 = 0
				System.BitConverter.GetBytes(GlobalMessageCode.eGetPirateStartLoc).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(Me.lEnvirID).CopyTo(yMsg, 2) : lPos += 4
				System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 6) : lPos += 4
				System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 10) : lPos += 4
				System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 14) : lPos += 4
				Dim oPlanet As Planet = CType(Me.oObject, Planet)
				If oPlanet Is Nothing = False Then
					oPlanet.oDomain.DomainSocket.SendData(yMsg)
				End If
				Return False
			ElseIf mlBaseCenterX = Int32.MaxValue OrElse mlBaseCenterZ = Int32.MaxValue Then
				'waiting on a response from region
				Return False
			End If
			Return True
		End Function

		Public Sub SetBaseCenterLoc(ByVal lX As Int32, ByVal lY As Int32)
			mlBaseCenterX = lX
			mlBaseCenterZ = lY
		End Sub
	End Class

	Private moEnvirs() As AureliusAIEnvir
	Private mlEnvirUB As Int32 = -1

	Private mlLastProcessCycle As Int32 = 0

	Public Shared lSpawnSystemLastPlanetPlacement As Int32 = -1

	Public Sub AddSystemAIEnvir(ByRef oSystem As SolarSystem)
		Dim lCurUB As Int32 = mlEnvirUB + 1
		Dim lFinalUB As Int32 = mlEnvirUB + (2 + oSystem.mlPlanetUB)

		ReDim Preserve moEnvirs(lFinalUB)

		'add the system
		moEnvirs(lCurUB) = New AureliusAIEnvir
		moEnvirs(lCurUB).oObject = oSystem
		moEnvirs(lCurUB).lEnvirID = oSystem.ObjectID
		moEnvirs(lCurUB).iTypeID = ObjectType.eSolarSystem

		lCurUB += 1
		'now, the planets
		For Y As Int32 = 0 To oSystem.mlPlanetUB
			Dim lTmp As Int32 = oSystem.GetPlanetIdx(Y)
			If lTmp > -1 Then
				moEnvirs(lCurUB) = New AureliusAIEnvir
				moEnvirs(lCurUB).oObject = goPlanet(lTmp)
				moEnvirs(lCurUB).lEnvirID = goPlanet(lTmp).ObjectID
				moEnvirs(lCurUB).iTypeID = ObjectType.ePlanet
			End If
			lCurUB += 1
		Next Y

		mlEnvirUB = lFinalUB

	End Sub

	Public Sub ProcessAureliusAI()

		If glCurrentCycle - mlLastProcessCycle > ml_AI_PROCESS_INTERVAL Then
			mlLastProcessCycle = glCurrentCycle
            If gb_IS_TEST_SERVER = True Then Return
            'Cycle thru environments
            If True = True Then Return
			For X As Int32 = 0 To mlEnvirUB

				If moEnvirs(X).lEnvirID < 1 Then
					If moEnvirs(X).oObject Is Nothing = False Then
						moEnvirs(X).lEnvirID = CType(moEnvirs(X).oObject, Epica_GUID).ObjectID
					End If
					Continue For
				End If

				If moEnvirs(X).iTypeID = ObjectType.eSolarSystem Then Continue For

				Dim bPlayerHere As Boolean = moEnvirs(X).PlayerLocatedHere()

				'Do I have Assets Here?
				If bPlayerHere = True AndAlso moEnvirs(X).AIAssetsHere() = True Then
					'Yes: Select a target within the environment and send units to attack
					moEnvirs(X).SetAttackTarget()
				End If

				'Alway check production first
                moEnvirs(X).ProcessProduction(True)

				'   Can I Produce?
				Dim yProduceable As Byte = moEnvirs(X).GetFacilityProductionCapability()
				If (yProduceable And (AureliusAIEnvir.eyFacilityProductionCapable.eFactory Or AureliusAIEnvir.eyFacilityProductionCapable.eSpaceport)) <> 0 Then
					'Yes: If Player in environment - 
					If bPlayerHere = True Then
						'add units to queues as necessary
                        moEnvirs(X).ProcessProduction(False)
					Else
						moEnvirs(X).ResetEnvir()
					End If
				Else
					'No: Engineer On Site?
					Dim yEngineerStatus As Byte = moEnvirs(X).EngineerOnSite()		'returns 0 if no engineers, 1 if engineer present but busy, > 1 if engineer is not busy
					If yEngineerStatus <> 0 Then
						'Yes: Is Engineer Idle?
						If yEngineerStatus <> 1 Then
							'ok, idle, order it to build
							moEnvirs(X).BuildNextFacility()
						End If
					Else
						'No: Can I build Engineers?
						If (yProduceable And AureliusAIEnvir.eyFacilityProductionCapable.eCommandCenter) <> 0 Then
							'Yes: Add an Engineer to the build queue of the command center
							moEnvirs(X).AddEngineerToProdQueueOfCC()
						Else
							'No: Colonizer Force Arrived?
							If moEnvirs(X).ColonizerForceArrived() = True Then
								'Yes: Begin Producing Colony... there should be a space engineer included in the colonizer force
								moEnvirs(X).BuildNextFacility()
							Else
								'No: Colonizer Force En Route?
								If moEnvirs(X).ColonizerForceEnRoute() = False Then
									'Spawn Colonizer Force on Homeworld
									SpawnColonizerForce(moEnvirs(X))
								End If
							End If
						End If
					End If
				End If
			Next X

		End If

	End Sub

	Public Sub AddAIEntity(ByVal lParentID As Int32, ByVal iParentTypeID As Int16, ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16, ByVal lEntityIdx As Int32)
		For X As Int32 = 0 To mlEnvirUB
			If moEnvirs(X).lEnvirID = lParentID AndAlso moEnvirs(X).iTypeID = iParentTypeID Then
				moEnvirs(X).AddAIEntity(lEntityID, iEntityTypeID, lEntityIdx)
				Return
			End If
		Next X

		LogEvent(LogEventType.CriticalError, "AureliusAI.AddAIFacility could not find AI Envir Object. Envir: " & lParentID & ", " & iParentTypeID)
	End Sub

	Public Sub RemoveAIEntity(ByVal lParentID As Int32, ByVal iParentTypeID As Int16, ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16)
		For X As Int32 = 0 To mlEnvirUB
			If moEnvirs(X).lEnvirID = lParentID AndAlso moEnvirs(X).iTypeID = iParentTypeID Then
				moEnvirs(X).RemoveAIEntity(lEntityID, iEntityTypeID)
				Return
			End If
		Next X

	End Sub

	Private Sub SpawnColonizerForce(ByRef oEnvir As AureliusAIEnvir)
        Dim l_SPACE_ENG_DEF_ID As Int32 = 3776
        If gb_IS_TEST_SERVER = True Then l_SPACE_ENG_DEF_ID = 3064
		Const l_AURELIUM_PRIME As Int32 = 508
		
		'Spawn the force on the homeworld
		Dim oHomeworld As Object = GetEpicaPlanet(l_AURELIUM_PRIME)
		Dim lSpawnX As Int32 = 0
		Dim lSpawnZ As Int32 = 0

		' - 2 space engineers
		Dim lSpaceEng1 As Int32 = SpawnNewUnit(l_SPACE_ENG_DEF_ID, eiBehaviorPatterns.eEngagement_Stand_Ground Or eiBehaviorPatterns.eTactics_Maneuver Or eiBehaviorPatterns.eTactics_Normal, eiTacticalAttrs.eArmedUnit, lSpawnX, lSpawnZ, oHomeworld)
		Dim lSpaceEng2 As Int32 = SpawnNewUnit(l_SPACE_ENG_DEF_ID, eiBehaviorPatterns.eEngagement_Stand_Ground Or eiBehaviorPatterns.eTactics_Maneuver Or eiBehaviorPatterns.eTactics_Normal, eiTacticalAttrs.eArmedUnit, lSpawnX, lSpawnZ, oHomeworld)

		If lSpaceEng1 = -1 AndAlso lSpaceEng2 = -1 Then
			LogEvent(LogEventType.CriticalError, "SpawnColonizerForce could not spawn space engineer.")
			Return
		End If

		' - 10 black hornets
		Dim lFighters(9) As Int32
		For X As Int32 = 0 To lFighters.GetUpperBound(0)
			lFighters(X) = SpawnNewUnit(l_BLACK_HORNET_DEF_ID, eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Maneuver Or eiBehaviorPatterns.eTactics_Normal, eiTacticalAttrs.eFighterClass, lSpawnX, lSpawnZ, oHomeworld)
		Next X

		' - 2 plunderers
        'Dim lPlunderer1 As Int32 = SpawnNewUnit(l_PLUNDERER_DEF_ID, eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Maximize_Damage, eiTacticalAttrs.eArmedUnit Or eiTacticalAttrs.eFacility, lSpawnX, lSpawnZ, oHomeworld)
        'Dim lPlunderer2 As Int32 = SpawnNewUnit(l_PLUNDERER_DEF_ID, eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Maximize_Damage, eiTacticalAttrs.eArmedUnit Or eiTacticalAttrs.eFacility, lSpawnX, lSpawnZ, oHomeworld)

		' - 1 fury
        'Dim lFury As Int32 = SpawnNewUnit(l_FURY_DEF_ID, eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Normal, eiTacticalAttrs.eFacility, lSpawnX, lSpawnZ, oHomeworld)

		'Add the entities to our colonizer force
		If lSpaceEng1 <> -1 Then
			Dim oUnit As Unit = goUnit(lSpaceEng1)
			If oUnit Is Nothing = False Then
				oEnvir.AddColonizerForceEntity(oUnit.ObjectID, lSpaceEng1)
			End If
		End If
		If lSpaceEng2 <> -1 Then
			Dim oUnit As Unit = goUnit(lSpaceEng2)
			If oUnit Is Nothing = False Then
				oEnvir.AddColonizerForceEntity(oUnit.ObjectID, lSpaceEng2)
			End If
		End If
		For X As Int32 = 0 To lFighters.GetUpperBound(0)
			If lFighters(X) <> -1 Then
				Dim oUnit As Unit = goUnit(lFighters(X))
				If oUnit Is Nothing = False Then
					oEnvir.AddColonizerForceEntity(oUnit.ObjectID, lFighters(X))
				End If
			End If
		Next X
        'If lPlunderer1 <> -1 Then
        '	Dim oUnit As Unit = goUnit(lPlunderer1)
        '	If oUnit Is Nothing = False Then
        '		oEnvir.AddColonizerForceEntity(oUnit.ObjectID, lPlunderer1)
        '	End If
        'End If
        'If lPlunderer2 <> -1 Then
        '	Dim oUnit As Unit = goUnit(lPlunderer2)
        '	If oUnit Is Nothing = False Then
        '		oEnvir.AddColonizerForceEntity(oUnit.ObjectID, lPlunderer2)
        '	End If
        'End If
        'If lFury <> -1 Then
        '	Dim oUnit As Unit = goUnit(lFury)
        '	If oUnit Is Nothing = False Then
        '		oEnvir.AddColonizerForceEntity(oUnit.ObjectID, lFury)
        '	End If
        'End If

	End Sub

	Private Shared Function SpawnNewUnit(ByVal lDefID As Int32, ByVal iCombatTactics As eiBehaviorPatterns, ByVal iTargetingTactics As eiTacticalAttrs, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByRef oParent As Object) As Int32
		Dim oDef As Epica_Entity_Def = GetEpicaUnitDef(lDefID)
		If oDef Is Nothing = False Then
			'ok, let's spawn them
			Dim oUnit As New Unit()
			With oUnit
				.bProducing = False
				.CurrentProduction = Nothing
				.CurrentSpeed = 0
				.EntityDef = oDef
				.CurrentStatus = 0
				For Y As Int32 = 0 To oDef.lSideCrits.Length - 1
					.CurrentStatus = .CurrentStatus Or oDef.lSideCrits(Y)
				Next Y

				oDef.DefName.CopyTo(.EntityName, 0)
				.ExpLevel = 0
				.Fuel_Cap = oDef.Fuel_Cap

				.iCombatTactics = iCombatTactics
				.iTargetingTactics = iTargetingTactics

				.LocAngle = 0

				'Determine where to place the units
				.LocX = lLocX
				.LocZ = lLocZ

				.Owner = GetEpicaPlayer(gl_HARDCODE_PIRATE_PLAYER_ID)
				.ParentObject = oParent
				.ObjTypeID = ObjectType.eUnit

				.Q1_HP = oDef.Q1_MaxHP
				.Q2_HP = oDef.Q2_MaxHP
				.Q3_HP = oDef.Q3_MaxHP
				.Q4_HP = oDef.Q4_MaxHP
				.Shield_HP = oDef.Shield_MaxHP
				.Structure_HP = oDef.Structure_MaxHP
				.yProductionType = oDef.ProductionTypeID

				.DataChanged()

				Dim lIdx As Int32 = -1
				If .SaveObject() = True Then
					'Now, find a suitable place...
                    'SyncLock goUnit
                    '	For Y As Int32 = 0 To glUnitUB
                    '		If glUnitIdx(Y) = -1 Then
                    '			lIdx = Y
                    '			Exit For
                    '		End If
                    '	Next Y

                    '	If lIdx = -1 Then
                    '		ReDim Preserve glUnitIdx(glUnitUB + 1)
                    '		ReDim Preserve goUnit(glUnitUB + 1)
                    '		glUnitUB += 1
                    '		lIdx = glUnitUB
                    '	End If

                    '	goUnit(lIdx) = oUnit
                    '	glUnitIdx(lIdx) = oUnit.ObjectID
                    'End SyncLock
                    lIdx = AddUnitToGlobalArray(oUnit)

					.DataChanged()
					.SaveObject()
				Else
					LogEvent(LogEventType.CriticalError, "SpawnNextWave Add Unit failed.")
				End If

				Dim iTemp As Int16 = CType(oParent, Epica_GUID).ObjTypeID
				If iTemp = ObjectType.ePlanet Then
					Dim oPlanet As Planet = CType(oParent, Planet)
					oPlanet.oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oUnit, GlobalMessageCode.eAddObjectCommand))
				ElseIf iTemp = ObjectType.eSolarSystem Then
					Dim oSystem As SolarSystem = CType(oParent, SolarSystem)
					oSystem.oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oUnit, GlobalMessageCode.eAddObjectCommand))
				End If

				Return lIdx

				'Dim lMoveToX As Int32 = oFac.LocX + CInt(Rnd() * 600) - 300
				'Dim lMoveToY As Int32 = oFac.LocZ + CInt(Rnd() * 600) - 300

				''Now, send a destination of the unit
				'Dim yMove(23) As Byte
				'System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequest).CopyTo(yMove, 0)
				'System.BitConverter.GetBytes(lMoveToX).CopyTo(yMove, 2)
				'System.BitConverter.GetBytes(lMoveToY).CopyTo(yMove, 6)
				'System.BitConverter.GetBytes(0S).CopyTo(yMove, 10)
				'System.BitConverter.GetBytes(lPlanetID).CopyTo(yMove, 12)
				'System.BitConverter.GetBytes(ObjectType.ePlanet).CopyTo(yMove, 16)
				'.GetGUIDAsString.CopyTo(yMove, 18)
				'oPlanet.oDomain.DomainSocket.SendData(yMove)
			End With
		End If

		Return Nothing
	End Function

    Public Shared Function SpawnNewFacility(ByVal lDefID As Int32, ByVal iCombatTactics As eiBehaviorPatterns, ByVal iTargetingTactics As eiTacticalAttrs, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByRef oParent As Object, ByVal bAddToEnvir As Boolean) As Int32
        Dim oDef As FacilityDef = GetEpicaFacilityDef(lDefID)
        If oDef Is Nothing = False Then
            'ok, let's spawn them
            Dim oFac As New Facility()
            With oFac
                .bProducing = False
                .CurrentProduction = Nothing
                .CurrentSpeed = 0
                .EntityDef = oDef
                .CurrentStatus = 0
                For Y As Int32 = 0 To oDef.lSideCrits.Length - 1
                    .CurrentStatus = .CurrentStatus Or oDef.lSideCrits(Y)
                Next Y

                oDef.DefName.CopyTo(.EntityName, 0)
                .ExpLevel = 0
                .Fuel_Cap = oDef.Fuel_Cap

                .iCombatTactics = iCombatTactics
                .iTargetingTactics = iTargetingTactics

                .LocAngle = 0

                'Determine where to place the units
                .LocX = lLocX
                .LocZ = lLocZ

                .Owner = GetEpicaPlayer(gl_HARDCODE_PIRATE_PLAYER_ID)
                .ParentObject = oParent
                .ObjTypeID = ObjectType.eFacility

                .Q1_HP = oDef.Q1_MaxHP
                .Q2_HP = oDef.Q2_MaxHP
                .Q3_HP = oDef.Q3_MaxHP
                .Q4_HP = oDef.Q4_MaxHP
                .Shield_HP = oDef.Shield_MaxHP
                .Structure_HP = oDef.Structure_MaxHP
                .yProductionType = oDef.ProductionTypeID

                .DataChanged()

                Dim lIdx As Int32 = -1
                If .SaveObject() = True Then
                    'Now, find a suitable place...
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

                    .DataChanged()
                    .SaveObject()
                Else
                    LogEvent(LogEventType.CriticalError, "SpawnNextWave Add Facility failed.")
                End If

                If bAddToEnvir = False Then
                    Dim iTemp As Int16 = CType(oParent, Epica_GUID).ObjTypeID
                    If iTemp = ObjectType.ePlanet Then
                        Dim oPlanet As Planet = CType(oParent, Planet)
                        oPlanet.oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oFac, GlobalMessageCode.eAddObjectCommand))
                    ElseIf iTemp = ObjectType.eSolarSystem Then
                        Dim oSystem As SolarSystem = CType(oParent, SolarSystem)
                        oSystem.oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oFac, GlobalMessageCode.eAddObjectCommand))
                    End If
                End If

                Return lIdx

                'Dim lMoveToX As Int32 = oFac.LocX + CInt(Rnd() * 600) - 300
                'Dim lMoveToY As Int32 = oFac.LocZ + CInt(Rnd() * 600) - 300

                ''Now, send a destination of the unit
                'Dim yMove(23) As Byte
                'System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequest).CopyTo(yMove, 0)
                'System.BitConverter.GetBytes(lMoveToX).CopyTo(yMove, 2)
                'System.BitConverter.GetBytes(lMoveToY).CopyTo(yMove, 6)
                'System.BitConverter.GetBytes(0S).CopyTo(yMove, 10)
                'System.BitConverter.GetBytes(lPlanetID).CopyTo(yMove, 12)
                'System.BitConverter.GetBytes(ObjectType.ePlanet).CopyTo(yMove, 16)
                '.GetGUIDAsString.CopyTo(yMove, 18)
                'oPlanet.oDomain.DomainSocket.SendData(yMove)
            End With
        End If

        Return Nothing
    End Function

	Public Shared Sub SpawnNextWave(ByVal lPlayerID As Int32, ByVal lPlanetID As Int32)
		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing Then Return
		If oPlayer.oSocket Is Nothing Then Return
		If oPlayer.PlayerIsDead = True OrElse oPlayer.yPlayerPhase <> eyPlayerPhase.eInitialPhase Then
			'End our timer for player survival thru the waves
			Dim lSeconds As Int32 = oPlayer.TotalPlayTime
			If oPlayer.oSocket Is Nothing = False Then
				lSeconds += CInt(Now.Subtract(oPlayer.LastLogin).TotalSeconds)
			End If
			oPlayer.PlayedTimeAtEndOfWaves = lSeconds
			Return
		End If

		Dim oPlanet As Planet = GetEpicaPlanet(lPlanetID)
		If oPlanet Is Nothing Then Return

		Dim lWaveID As Int32 = 1

		If oPlayer.TutorialPhaseWaves < 0 Then oPlayer.TutorialPhaseWaves = 0
		oPlayer.TutorialPhaseWaves += 1
		lWaveID = Math.Max(oPlayer.TutorialPhaseWaves, 1)

		If oPlayer.PlayedTimeWhenFirstWave = Int32.MinValue Then
			Dim lSeconds As Int32 = oPlayer.TotalPlayTime
			If oPlayer.oSocket Is Nothing = False Then
				lSeconds += CInt(Now.Subtract(oPlayer.LastLogin).TotalSeconds)
			End If
			oPlayer.PlayedTimeWhenFirstWave = lSeconds
		End If

        'Ok, let's get the colony in question
        Dim lColonyIdx As Int32 = oPlayer.GetColonyFromParent(lPlanetID, ObjectType.ePlanet)
        If lColonyIdx = -1 Then Return
        If glColonyIdx(lColonyIdx) = -1 Then Return

        Dim oColony As Colony = goColony(lColonyIdx)

        If lWaveID > 10 Then
            If oColony Is Nothing = False Then
                oColony.Population = 0
                Return
            End If
        End If

		'Send an alert to the client indicating the wave number....
		Dim yMsg(5) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.ePirateWaveSpawn).CopyTo(yMsg, 0)
		System.BitConverter.GetBytes(lWaveID).CopyTo(yMsg, 2)
		oPlayer.SendPlayerMessage(yMsg, False, 0)

        Dim oFac As Facility = Nothing
		If oColony Is Nothing = False Then
			'ok, determine a facility... we prefer Command Centers... then Factories... then random
			Dim lFacCnt As Int32 = 0
			Dim lUB As Int32 = -1
			If oColony.oChildren Is Nothing = False Then lUB = Math.Min(oColony.ChildrenUB, oColony.oChildren.GetUpperBound(0))
			For X As Int32 = 0 To lUB
				If oColony.oChildren(X) Is Nothing = False Then
					lFacCnt += 1
					If oColony.oChildren(X).yProductionType = ProductionType.eCommandCenterSpecial Then
						oFac = oColony.oChildren(X)
						Exit For
					ElseIf (oColony.oChildren(X).yProductionType And ProductionType.eProduction) <> 0 Then
						If oFac Is Nothing = False Then oFac = oColony.oChildren(X)
					End If
				End If
			Next X
			If oFac Is Nothing Then
				'ok, get a random facility
				Dim lFacility As Int32 = CInt(Rnd() * lFacCnt)
				Dim lLastValidIdx As Int32 = -1
				For X As Int32 = 0 To lUB
					If oColony.oChildren(X) Is Nothing = False Then
						lFacility -= 1
						lLastValidIdx = X
						If lFacility < 1 Then
							oFac = oColony.oChildren(X)
							Exit For
						End If
					End If
				Next X
				If oFac Is Nothing Then
					If lLastValidIdx <> -1 Then
						oFac = oColony.oChildren(lLastValidIdx)
					Else
						'Not sure what to do here....
						Return
					End If
				End If
			End If
		End If
		If oFac Is Nothing Then Return

		Dim lObjDefID(6) As Int32
		Dim lObjCnts(6) As Int32
		For X As Int32 = 0 To 6
			lObjCnts(X) = 0
		Next X

		lObjDefID(0) = l_JEEP_DEF_ID : lObjDefID(1) = l_SM_TANK_DEF_ID : lObjDefID(2) = l_LG_TANK_DEF_ID
		lObjDefID(3) = l_BLACK_HORNET_DEF_ID : lObjDefID(4) = l_PLUNDERER_DEF_ID : lObjDefID(5) = l_FURY_DEF_ID
		lObjDefID(6) = l_SMITER_DEF_ID


		'ok, determine what to spawn
		Select Case lWaveID
			Case 1
				lObjCnts(0) = 2 : lObjCnts(1) = 2 : lObjCnts(2) = 1
			Case 2
				lObjCnts(0) = 5 : lObjCnts(1) = 5 : lObjCnts(2) = 5 : lObjCnts(3) = 2
			Case 3
				lObjCnts(0) = 5 : lObjCnts(1) = 5 : lObjCnts(2) = 15 : lObjCnts(3) = 5
			Case 4
				lObjCnts(0) = 0 : lObjCnts(1) = 10 : lObjCnts(2) = 15 : lObjCnts(3) = 10
			Case 5
				lObjCnts(2) = 25 : lObjCnts(3) = 25
			Case 6
				lObjCnts(0) = 10 : lObjCnts(3) = 15 : lObjCnts(4) = 5
			Case 7
				lObjCnts(1) = 5 : lObjCnts(2) = 10 : lObjCnts(3) = 15 : lObjCnts(4) = 5
			Case 8
				lObjCnts(2) = 15 : lObjCnts(3) = 20 : lObjCnts(4) = 10
			Case 9
				lObjCnts(2) = 15 : lObjCnts(3) = 20 : lObjCnts(4) = 10 : lObjCnts(5) = 2
			Case 10
				lObjCnts(2) = 10 : lObjCnts(3) = 20 : lObjCnts(4) = 5 : lObjCnts(5) = 5
			Case 11
				lObjCnts(2) = 10 : lObjCnts(3) = 20 : lObjCnts(4) = 10 : lObjCnts(5) = 10
			Case 12
				lObjCnts(4) = 15 : lObjCnts(5) = 10
			Case 13
				lObjCnts(4) = 15 : lObjCnts(5) = 10 : lObjCnts(6) = 2
			Case 14
				lObjCnts(3) = 10 : lObjCnts(4) = 15 : lObjCnts(5) = 10 : lObjCnts(6) = 5
			Case Else
				lObjCnts(3) = 10 + (5 * (lWaveID - 15))
				lObjCnts(4) = 20
				lObjCnts(5) = 10
				lObjCnts(6) = 5 + (5 * (lWaveID - 15))
		End Select

		Dim lRoll As Int32 = CInt(Rnd() * 100)
		Dim lLocX As Int32
		Dim lLocZ As Int32
		If lRoll < 25 Then
			lLocX = 10800 : lLocZ = -14400
		ElseIf lRoll < 50 Then
			lLocX = -16341 : lLocZ = -1550
		ElseIf lRoll < 75 Then
			lLocX = 12000 : lLocZ = 17900
		Else
			lLocX = 15800 : lLocZ = 2800
		End If

		'Now, go through our different types and spawn them
		For lType As Int32 = 0 To 6
			'ok, are there units to spawn of this type?
			If lObjCnts(lType) <> 0 Then

				Dim iCombatTactics As eiBehaviorPatterns
				Dim iTargetingTactics As eiTacticalAttrs

				If lType = 3 Then
                    iCombatTactics = eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Maximize_Damage Or eiBehaviorPatterns.eTactics_Maneuver
                    iTargetingTactics = eiTacticalAttrs.eFighterClass Or eiTacticalAttrs.eArmedUnit Or eiTacticalAttrs.eFighterTargetRadar
                Else
                    iCombatTactics = eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eTactics_Maximize_Damage
                    iTargetingTactics = eiTacticalAttrs.eFacility
				End If

				Dim lUnitIdx As Int32 = SpawnNewUnit(lObjDefID(lType), iCombatTactics, iTargetingTactics, lLocX + CInt(Rnd() * 400) - 200, lLocZ + CInt(Rnd() * 400) - 200, CObj(oPlanet))
				If lUnitIdx <> -1 Then
					Dim oUnit As Unit = goUnit(lUnitIdx)
					If oUnit Is Nothing = False Then
						Dim lMoveToX As Int32 = oFac.LocX + CInt(Rnd() * 600) - 300
						Dim lMoveToY As Int32 = oFac.LocZ + CInt(Rnd() * 600) - 300

						'Now, send a destination of the unit
						Dim yMove(23) As Byte
						System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequest).CopyTo(yMove, 0)
						System.BitConverter.GetBytes(lMoveToX).CopyTo(yMove, 2)
						System.BitConverter.GetBytes(lMoveToY).CopyTo(yMove, 6)
						System.BitConverter.GetBytes(0S).CopyTo(yMove, 10)
						System.BitConverter.GetBytes(lPlanetID).CopyTo(yMove, 12)
						System.BitConverter.GetBytes(ObjectType.ePlanet).CopyTo(yMove, 16)
						oUnit.GetGUIDAsString.CopyTo(yMove, 18)
						oPlanet.oDomain.DomainSocket.SendData(yMove)
					End If
				End If

			End If
		Next lType

		AddToQueue(glCurrentCycle + 1800, QueueItemType.eSpawnNextPirateWave, lPlayerID, lPlanetID, -1, -1, 0, 0, 0, 0)

	End Sub

	Public Sub HandlePlacePirateAsset(ByRef yData() As Byte)
		'MscCode, EnvirGUID, StartLoc, PlayerID, ItemCnt, Items (GUID, Loc)
		Dim lPos As Int32 = 2
		Dim lEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		lPos += 16	'for start loc and playerid and itemcnt (which should always be 1)

		Dim lItemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim iItemTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
		Dim lLocX As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lLocZ As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4


		For X As Int32 = 0 To mlEnvirUB
			If moEnvirs(X).lEnvirID = lEnvirID AndAlso moEnvirs(X).iTypeID = iEnvirTypeID Then
				moEnvirs(X).HandlePlacePirateAsset(lItemID, iItemTypeID, lLocX, lLocZ)
				Return
			End If
		Next X

	End Sub

	Public Sub HandleGetPirateStartLoc(ByVal lPlanetID As Int32, ByVal lLocX As Int32, ByVal lLocY As Int32)
		For X As Int32 = 0 To mlEnvirUB
			If moEnvirs(X).iTypeID = ObjectType.eSolarSystem Then Continue For
			If moEnvirs(X).lEnvirID = lPlanetID Then
				moEnvirs(X).SetBaseCenterLoc(lLocX, lLocY)
			End If
		Next X
	End Sub

	Public Sub HandlePirateStopped(ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lObjID As Int32, ByVal iObjTypeID As Int16)
		For X As Int32 = 0 To mlEnvirUB
			If moEnvirs(X).lEnvirID = lEnvirID AndAlso moEnvirs(X).iTypeID = iEnvirTypeID Then
				moEnvirs(X).EngineerReachedDest(lObjID)
				Return
			End If
		Next X
	End Sub
End Class