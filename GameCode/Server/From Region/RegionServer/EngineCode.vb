Option Strict On

'This module contains all of the engine code for the RegionServer
Module EngineCode
    Private Const ml_FUEL_CYCLE_INTERVAL As Int32 = 300   '10 seconds
	Private mlLastFuelCycle As Int32

    Private mbRunAntiLightCode As Boolean = True

#Region "  Movement Registers  "
	Private mlMovingUB As Int32 = -1
	Private mlMovingIdx(-1) As Int32		'serverindex
	Private mlMovingID(-1) As Int32		'ID of the entity moving

	Public Enum MovementCommand As Byte
		AddEntityMoving = 0
		RemoveEntityMoving = 1
		PackMovementRegisters = 2
	End Enum

	Public Sub ResetMovementRegisters()
		mlMovingUB = -1
		ReDim mlMovingIdx(-1)
        ReDim mlMovingID(-1)

        Dim lCurUB As Int32 = -1
        If glEntityIdx Is Nothing = False Then lCurUB = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        For X As Int32 = 0 To lCurUB
            If glEntityIdx(X) > 0 Then
                Dim oEntity As Epica_Entity = goEntity(X)
                If oEntity Is Nothing = False Then
                    oEntity.bInMovementRegister = False
                End If
            End If
        Next X
	End Sub
	'Public Sub AddEntityMoving(ByVal lServerIdx As Int32, ByVal lEntityID As Int32)
	'	SyncLockMovementRegisters(MovementCommand.AddEntityMoving, lServerIdx, lEntityID)
	'End Sub
	'Public Sub RemoveEntityMoving(ByVal lServerIdx As Int32, ByVal lEntityID As Int32)
	'	SyncLockMovementRegisters(MovementCommand.RemoveEntityMoving, lServerIdx, lEntityID)
	'End Sub

	Private Sub DoAddEntityMoving(ByVal lServerIdx As Int32, ByVal lEntityID As Int32)
		If lServerIdx = -1 OrElse lEntityID = -1 Then Return
		Dim oEntity As Epica_Entity = goEntity(lServerIdx)
		If oEntity Is Nothing Then Return

		'See if this works again, may solve a LOT of issues
        If oEntity.bInMovementRegister = True Then Return
        oEntity.bInMovementRegister = True

        'HACK: MSC - 7/24/08 - this difference equation ensures we do not get any weird warping occuring. This assumes that are LocX and LocZ are set correctly from the primary.
        If glCurrentCycle - oEntity.LastCycleMoved > 150 Then
            oEntity.LastCycleMoved = glCurrentCycle
        End If

        Dim lCurUB As Int32 = Math.Min(Math.Min(mlMovingUB, mlMovingIdx.GetUpperBound(0)), mlMovingID.GetUpperBound(0))
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To lCurUB
            If mlMovingIdx(X) = -1 AndAlso lIdx = -1 Then
                lIdx = X
                mlMovingIdx(X) = -2 'set us to -2 so we do not get overwritten elsewhere
            ElseIf mlMovingIdx(X) = lServerIdx AndAlso mlMovingID(X) = lEntityID Then
                'already in there
                If lIdx <> -1 Then mlMovingIdx(lIdx) = -1
                Return
            End If
        Next X

        'SyncLock mlMovingIdx
        Try
            If lIdx = -1 Then
                lIdx = mlMovingUB + 1
                If lIdx > mlMovingIdx.GetUpperBound(0) Then
                    ReDim Preserve mlMovingIdx(mlMovingUB + 10000)
                    ReDim Preserve mlMovingID(mlMovingUB + 10000)
                End If
                'ReDim Preserve mlMovingIdx(lIdx)
                mlMovingIdx(lIdx) = -2
                'ReDim Preserve mlMovingID(lIdx)
                mlMovingUB += 1
            End If
            mlMovingIdx(lIdx) = lServerIdx
            mlMovingID(lIdx) = lEntityID
        Catch
            'do nothing, they will need to send it moving again
        End Try
        'End SyncLock
	End Sub
	Private Sub DoRemoveEntityMoving(ByVal lServerIdx As Int32, ByVal lEntityID As Int32)
		If lServerIdx = -1 OrElse lEntityID = -1 Then Return
		Dim oEntity As Epica_Entity = goEntity(lServerIdx)
		If oEntity Is Nothing Then Return
		oEntity.bInMovementRegister = False

		Dim lCurUB As Int32 = Math.Min(Math.Min(mlMovingUB, mlMovingIdx.GetUpperBound(0)), mlMovingID.GetUpperBound(0))
		For X As Int32 = 0 To lCurUB
			If mlMovingIdx(X) = lServerIdx AndAlso mlMovingID(X) = lEntityID Then
				mlMovingIdx(X) = -1 : mlMovingID(X) = -1
			End If
		Next X
	End Sub
	Private Sub DoPackMovementRegisters()
		Dim lCurUB As Int32 = Math.Min(Math.Min(mlMovingUB, mlMovingIdx.GetUpperBound(0)), mlMovingID.GetUpperBound(0))
		Dim lGaps(lCurUB) As Int32

		Dim lNextGapPlacement As Int32 = 0

		Dim lLastE2BStart As Int32 = lCurUB
		For b2e As Int32 = 0 To lCurUB
			If mlMovingIdx(b2e) = -1 Then
				Dim bPlaced As Boolean = False
				For e2b As Int32 = lLastE2BStart To b2e Step -1
					If mlMovingIdx(e2b) <> -1 Then
						bPlaced = True
						mlMovingIdx(b2e) = mlMovingIdx(e2b)
						mlMovingID(b2e) = mlMovingID(e2b)
						mlMovingIdx(e2b) = -1
						mlMovingID(e2b) = -1
						lLastE2BStart = e2b
						Exit For
					End If
				Next e2b
				If bPlaced = False Then Exit For
			End If
		Next b2e

		For X As Int32 = lCurUB To 0 Step -1
			If mlMovingIdx(X) <> -1 Then
				mlMovingUB = X
				Exit For
			End If
		Next X
	End Sub

	Public Sub SyncLockMovementRegisters(ByVal yType As MovementCommand, ByVal lServerIdx As Int32, ByVal lEntityID As Int32)
		SyncLock mlMovingIdx
			Select Case yType
				Case MovementCommand.AddEntityMoving
					DoAddEntityMoving(lServerIdx, lEntityID)
				Case MovementCommand.PackMovementRegisters
					DoPackMovementRegisters()
				Case MovementCommand.RemoveEntityMoving
					DoRemoveEntityMoving(lServerIdx, lEntityID)
			End Select
		End SyncLock
	End Sub


#End Region

	'#Region "  Combat Registers  "
	'	Private mlCombatUB As Int32 = -1
	'	Private mlCombatIdx(-1) As Int32		'serverindex
	'	Private mlCombatID(-1) As Int32		'ID of the entity in combat

	'	Public Sub AddEntityCombat(ByVal lServerIdx As Int32, ByVal lEntityID As Int32)
	'		If lServerIdx = -1 OrElse lEntityID = -1 Then Return
	'		Dim oEntity As Epica_Entity = goEntity(lServerIdx)
	'		If oEntity Is Nothing Then Return
	'		If oEntity.bInMovementRegister = True Then Return
	'		oEntity.bInMovementRegister = True

	'		Dim lCurUB As Int32 = Math.Min(Math.Min(mlCombatUB, mlCombatIdx.GetUpperBound(0)), mlCombatID.GetUpperBound(0))
	'		Dim lIdx As Int32 = -1
	'		For X As Int32 = 0 To lCurUB
	'			If mlCombatIdx(X) = -1 AndAlso lIdx = -1 Then
	'				lIdx = X
	'				mlCombatIdx(X) = -2	'set us to -2 so we do not get overwritten elsewhere
	'			ElseIf mlCombatIdx(X) = lServerIdx AndAlso mlCombatID(X) = lEntityID Then
	'				'already in there
	'				Return
	'			End If
	'		Next X

	'		If lIdx = -1 Then
	'			ReDim mlCombatIdx(mlMovingUB + 1) : mlCombatIdx(mlMovingUB + 1) = -2
	'			lIdx = mlMovingUB + 1
	'			mlCombatIdx(lIdx) = -2
	'			ReDim mlCombatID(lIdx)
	'			mlMovingUB += 1
	'		End If
	'		mlCombatIdx(lIdx) = lServerIdx
	'		mlCombatID(lIdx) = lEntityID
	'	End Sub
	'	Public Sub RemoveEntityCombat(ByVal lServerIdx As Int32, ByVal lEntityID As Int32)
	'		If lServerIdx = -1 OrElse lEntityID = -1 Then Return
	'		Dim oEntity As Epica_Entity = goEntity(lServerIdx)
	'		If oEntity Is Nothing Then Return
	'		oEntity.bInMovementRegister = False
	'		oEntity.ClearJamModifiers()

	'		Dim lCurUB As Int32 = Math.Min(Math.Min(mlMovingUB, mlCombatIdx.GetUpperBound(0)), mlCombatID.GetUpperBound(0))
	'		For X As Int32 = 0 To lCurUB
	'			If mlCombatIdx(X) = lServerIdx AndAlso mlCombatID(X) = lEntityID Then
	'				mlCombatIdx(X) = -1 : mlCombatID(X) = -1
	'			End If
	'		Next X
	'	End Sub
	'#End Region

	'NOTE: You must make sure that this routine is very identical to the Client Movement routine in order to
	'  ensure that the client portal is accurately displaying real data. The server is always right, therefore
	'  this code HAS to be incredibly accurate
	Public Sub HandleMovement()
		Dim X As Int32
		Dim bTurned As Boolean
		Dim iTemp As Short
		Dim iTurnAmt As Short
		Dim lCyclesToStop As Int32
		Dim fDistToStop As Single
		Dim fTotDist As Single
		Dim lLargeSector As Int32
		Dim lSmallSector As Int32
		Dim lGrid As Int32
		Dim lSmallPerRow As Int32
		Dim Y As Int32
		Dim oTmpUnit As Epica_Entity
		Dim yRelID As Byte

		Dim oPrimaryTarget As Epica_Entity = Nothing
		Dim lPrimaryTargetScore As Int32
		Dim lPrimaryFacing As Int32

        Dim lFaceTargetScores(3) As Int32       '1 for each facing
        Dim lFighterTargetScores(3) As Int32
		Dim lFacing As Int32

		Dim lTemp As Int32
		Dim lTemp2 As Int32

		Dim lEnvirDist As Int32

		Dim yOutMsg() As Byte

		Dim fTempVelX As Single
		Dim fTempVelZ As Single
		Dim dVecAngleRads As Single

		'for dealing with Skipping...
		Dim lCyclesLapsed As Int32

		'to remove redundancy
		Dim fAbsDiffX As Single
		Dim fAbsDiffZ As Single

		Dim oTmpDef As Epica_Entity_Def
		Dim oTmpDef2 As Epica_Entity_Def = Nothing

		Dim lRelTinyX As Int32
		Dim lRelTinyZ As Int32

		Dim bSkipReverse As Boolean

		Dim bResetManeuver As Boolean
		Dim bSendDIMsg As Boolean

		Dim fMaxSpeed As Single

		'Dim bUpdateFuel As Boolean = ((glCurrentCycle - mlLastFuelCycle) >= ml_FUEL_CYCLE_INTERVAL)

		'Cycle through ALL global units. The unit is environment-specific... However, only units move and thus, they are
		'  the only things that need to be passed thru the movement engine... When aggression occurs, the test is made
		'  against Environment.oObjs(lIdx) which is an EObject. EObject can relate to a Unit or a Facility.
		Try
			'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
			Dim lCurUB As Int32 = Math.Min(Math.Min(mlMovingUB, mlMovingIdx.GetUpperBound(0)), mlMovingID.GetUpperBound(0))
			For X = 0 To lCurUB

				If mlMovingIdx(X) > -1 AndAlso mlMovingID(X) <> -1 Then

					If glEntityIdx(mlMovingIdx(X)) <> mlMovingID(X) Then
						mlMovingIdx(X) = -1
						mlMovingID(X) = -1
						Continue For
					End If

					bTurned = False
					bResetManeuver = False

					Dim oEntity As Epica_Entity = goEntity(mlMovingIdx(X))
                    If oEntity Is Nothing Then
                        'NOTE: This may cause issues setting glEntityIdx() to -1
                        glEntityIdx(mlMovingIdx(X)) = -1
                        mlMovingIdx(X) = -1
                        Continue For
                    End If

					With oEntity
						If .ParentEnvir Is Nothing Then
                            glEntityIdx(mlMovingIdx(X)) = -1
                            RemoveLookupEntity(.ObjectID, .ObjTypeID)
                            Continue For
						End If

						'Are we moving?'If .LocX <> .DestX Or .LocZ <> .DestZ Or .LocAngle <> .DestAngle Then
						lCyclesLapsed = glCurrentCycle - .LastCycleMoved
						'lCyclesLapsed = 1
						If lCyclesLapsed < 1 Then Continue For

						'Get the def object...
						oTmpDef = Nothing
						If .lEntityDefServerIndex <> -1 Then
							If glEntityDefIdx(.lEntityDefServerIndex) <> -1 Then
								oTmpDef = goEntityDefs(.lEntityDefServerIndex)
							End If
						End If

						If oTmpDef Is Nothing Then Continue For

						If (oTmpDef.BaseAcceleration = 0 OrElse oTmpDef.BaseMaxSpeed = 0) AndAlso (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
							'.CurrentStatus -= elUnitStatus.eUnitMoving
							.CurrentStatus = .CurrentStatus Xor elUnitStatus.eUnitMoving
                            'RemoveEntityMoving(.ServerIndex, .ObjectID)
                            If .bForceAggressionTest = False Then SyncLockMovementRegisters(MovementCommand.RemoveEntityMoving, .ServerIndex, .ObjectID)
						End If

						If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 AndAlso (lCyclesLapsed > 0) Then
							gblMovements += 1

                            Dim fInitialVel As Single = .TotalVelocity

							Dim fStartDiffX As Single = .DestX - .LocX
							Dim fStartDiffZ As Single = .DestZ - .LocZ

							Dim bBreakoutEntity As Boolean = False

							Dim lCyclesAccelerate As Int32 = 0
							Dim lCyclesDecelerate As Int32 = 0
							Dim iTAx100 As Int16

							'Deal with turning here if it needs to be done
							'NOTE: Angle is now measured in 10th's of a degree. therefore, a full circle has 3600 angle units
							If .LocAngle <> .DestAngle Then
								bTurned = True

								'Change our Angle
								iTemp = .DestAngle - .LocAngle
								If .yFormationTurnAmount <> 0 Then iTurnAmt = .yFormationTurnAmount Else iTurnAmt = .TurnAmount

								'Ok, we turn, while we turn, we decelerate... unless our turn amount x100 > Abs(LocAngle - DestAngle), 
								'at which point, we accelerate

								'we decelerate unless we are in the x100 area, at which point, we accelerate
								'so, we cannot turn more than 3600, the difference of dest and loc, or my turn amount * cycles
								Dim iMaxTurnPotential As Int16 = CShort(Math.Min(3600, Math.Min(iTurnAmt * lCyclesLapsed, Math.Abs(iTemp))))
								'calculate turn amount x100
								If .iFormationTurnAmount100 <> 0 Then iTAx100 = .iFormationTurnAmount100 Else iTAx100 = .TurnAmountTimes100
								Dim lTurnInDecelerate As Int32 = Math.Min(iMaxTurnPotential, Math.Abs(iTemp))
								Dim lTurnInAccelerate As Int32 = 0 '= Math.Min(lTurnInDecelerate, Math.Max(0, iTAx100 - Math.Abs(iTemp) + iMaxTurnPotential))
								If Math.Abs(iTemp) - lTurnInDecelerate < iTAx100 Then
									lTurnInAccelerate = Math.Min(iMaxTurnPotential, Math.Abs(iTAx100 - Math.Abs(iTemp) - lTurnInDecelerate))
								End If
								lTurnInDecelerate -= lTurnInAccelerate
								lCyclesDecelerate = lTurnInDecelerate \ iTurnAmt
								lCyclesAccelerate = lCyclesLapsed - lCyclesDecelerate  'lTurnInAccelerate \ iTurnAmt

								'Ok, apply our turn amount
								If Math.Abs(iTemp) = 1800S Then
									.LocAngle -= iMaxTurnPotential
								Else
									If iTemp < 0 Then
										If iTemp > -1800S Then
											'CCW
											.LocAngle -= iMaxTurnPotential
										Else
											'CW
											.LocAngle += iMaxTurnPotential
										End If
									Else
										If iTemp > 1800S Then
											'CCW
											.LocAngle -= iMaxTurnPotential
										Else
											'CW
											.LocAngle += iMaxTurnPotential
										End If
									End If	'if iTemp < 0
									If .LocAngle < 0 Then
										.LocAngle += 3600S
									ElseIf .LocAngle > 3599S Then
										.LocAngle -= 3600S
									End If
								End If

							Else
								bTurned = False
								lCyclesAccelerate = lCyclesLapsed
								lCyclesDecelerate = 0
							End If		'.LocAngle <> .DestAngle


							'Caculate our MaxSpeed
							If .yFormationMaxSpeed <> 0 Then fMaxSpeed = .yFormationMaxSpeed Else fMaxSpeed = .MaxSpeed
							If oTmpDef.yChassisType = ChassisType.eGroundBased Then
								fMaxSpeed *= .ParentEnvir.GetEnvirSpeedMod(.LocX, .LocZ, .LocAngle)
								If fMaxSpeed < 1 Then fMaxSpeed = 1
								'Else : fMaxSpeed = fMaxSpeed ' oTmpDef.BaseMaxSpeed
							End If

							'Determine our distance to begin slowing down both X and Z
							'lCyclesToStop = CInt(Math.Floor(System.Math.Abs(.TotalVelocity / .Acceleration))) ' oTmpDef.Acceleration)
							Dim fAcc As Single
							If .yFormationManeuver <> 0 Then fAcc = .fFormationAcceleration Else fAcc = .Acceleration
							If fAcc = 0 Then lCyclesToStop = 0 Else lCyclesToStop = CInt(System.Math.Abs(.TotalVelocity / fAcc))

							fDistToStop = System.Math.Abs(.TotalVelocity * lCyclesToStop) + (0.5F * fAcc * lCyclesToStop) ' oTmpDef.Acceleration * lCyclesToStop)
							fAbsDiffX = System.Math.Abs(.DestX - .MoveLocX)
							fAbsDiffZ = System.Math.Abs(.DestZ - .MoveLocZ)
							fTotDist = fAbsDiffX + fAbsDiffZ

							Dim fAdditionalTravel As Single = 0.0F
							Dim bDoTurnCalc As Boolean = False

							'Ok, we may need to burn some cycles decelerating (for example, if we are in a turn)
							If lCyclesDecelerate <> 0 Then
								'ok, decelerate because we are turning... no DI, maneuver change or stop event
								If .TotalVelocity <> 0 Then
									'ok, we are moving while decelerating, determine by how much we are moving
									Dim lActualCyclesDecelerate As Int32 = Math.Min(lCyclesDecelerate, lCyclesToStop)
									Dim fTmpAcc As Single = fAcc * 0.25F
									If lActualCyclesDecelerate > 43000 Then lActualCyclesDecelerate = 43000
									fAdditionalTravel += (.TotalVelocity * lActualCyclesDecelerate) + (-fTmpAcc * (lActualCyclesDecelerate * lActualCyclesDecelerate) * 0.5F) ' ((fTmpAcc * lActualCyclesDecelerate) * (fTmpAcc * lActualCyclesDecelerate)) * 0.5F
									.TotalVelocity -= (fTmpAcc * lActualCyclesDecelerate)
                                    If .TotalVelocity < 0 OrElse Single.IsNaN(.TotalVelocity) = True Then .TotalVelocity = 0
									bDoTurnCalc = True
								End If
							End If

							'Now, we are here, lCyclesAccelerate are the remaining movement cycles, while moving...
							'  we will attempt to accelerate as long as we will have room remaining to decelerate
							'  furthermore, we cannot exceed our maxspeed
							If lCyclesAccelerate <> 0 Then
								Dim lActualAcc As Int32 = 0
								'I can accelerate to a point where my maxspeed and distance travelled will exceed the ability to stop
								'm = (al - l) +/- sqrt ( 5l^2 - a^2l^2 + 4a^2d + 4ad )
								'    -------------------------------------------------
								'						2a+2
								'a = acceleration
								'l = vel start
								'd = dist to target

								Dim fM As Single
								If .TotalVelocity <> fMaxSpeed Then
									Dim fVelAtStart As Single = .TotalVelocity
									Dim fASquared As Single = fAcc * fAcc
									Dim fLSquared As Single = fVelAtStart * fVelAtStart
									fM = 5.0F * fLSquared
									fM -= fASquared * fLSquared
									fM += 4.0F * fASquared * fTotDist
									fM += 4.0F * fAcc * fTotDist
									fM = CSng(Math.Sqrt(fM))
									fM += (fAcc * fVelAtStart - fVelAtStart)
									fM /= (2 * fAcc + 2)
									fM = Math.Min(fM, fMaxSpeed)
								Else : fM = fMaxSpeed
								End If

								'fM now has our max speed before we need to decelerate
								If .TotalVelocity < fM Then lActualAcc = CInt(Math.Min(lCyclesAccelerate, (fM - .TotalVelocity) / fAcc))

								If lActualAcc <> 0 Then
									'ok, now reduce lActualAcc from lCyclesAccelerate
									lCyclesAccelerate -= lActualAcc
									'Next, we accelerate
									If lActualAcc > 43000 Then lActualAcc = 43000
									fAdditionalTravel += (.TotalVelocity * lActualAcc) + (fAcc * (lActualAcc * lActualAcc) * 0.5F)
									.TotalVelocity += (lActualAcc * fAcc)
								End If

								'Now, do we have remaining cycles? if so, we need to calculate the distance travelled at present speed (max)
								If lCyclesAccelerate <> 0 Then
									'ok, did we actually accelerate?
									.TotalVelocity = fM
									If lActualAcc <> 0 Then
										'recalculate our dist to stop
										fDistToStop = System.Math.Abs(.TotalVelocity * lCyclesToStop) + (0.5F * fAcc * lCyclesToStop)
									End If
									'Ok, now, determine if we need to decelerate...
									If fTotDist - fAdditionalTravel < fDistToStop Then
										'ok, we need to decelerate
										lCyclesDecelerate = lCyclesAccelerate
										lCyclesToStop = CInt(System.Math.Abs(.TotalVelocity / fAcc))
										Dim lActualCyclesDecelerate As Int32 = Math.Min(lCyclesDecelerate, lCyclesToStop)

										If lCyclesToStop - lActualCyclesDecelerate < 34 Then
											If .yResetMoveByPlayer <> 0 Then
												If (.CurrentStatus And elUnitStatus.eMovedByPlayer) <> 0 Then
													.CurrentStatus = .CurrentStatus Xor elUnitStatus.eMovedByPlayer
												End If
												.yResetMoveByPlayer = 0
											End If

											'first, are we in maneuvers? and not moved by the player?
											bSendDIMsg = True
											If (.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 AndAlso _
											   (.iCombatTactics And eiBehaviorPatterns.eTactics_Maneuver) <> 0 AndAlso _
											   (.CurrentStatus And elUnitStatus.eMovedByPlayer) <> 0 Then
												'Are we targeting a valid primary?
												If .lPrimaryTargetServerIdx <> -1 Then
                                                    If glEntityIdx(.lPrimaryTargetServerIdx) > 0 Then
                                                        bSendDIMsg = False
                                                        'yes, we are... do we know our min speed?
                                                        If .MinSpeed = Single.MinValue Then
                                                            'no, set a temp unit
                                                            oTmpUnit = goEntity(.lPrimaryTargetServerIdx)
                                                            'now, check that entity's speed
                                                            If oTmpUnit.TotalVelocity > 1 Then
                                                                .MinSpeed = ((fMaxSpeed - oTmpUnit.TotalVelocity) * (oTmpUnit.TotalVelocity / fMaxSpeed)) + oTmpUnit.TotalVelocity
                                                            End If
                                                            .MinSpeed = Math.Max(.MinSpeed, fMaxSpeed * 0.25F) ' / 4.0F)		'was 4.0f
                                                        End If

                                                        'Now, we know our min speed, are we moving slower than that?
                                                        If .TotalVelocity < .MinSpeed Then
                                                            bResetManeuver = True
                                                            'reset our min speed for the next time
                                                            .MinSpeed = Single.MinValue
                                                        End If
                                                    Else : .lPrimaryTargetServerIdx = -1
                                                    End If
												End If
											End If

                                            If bSendDIMsg = True Then
                                                If .bDecelerating = False AndAlso lCyclesToStop < 34 Then

                                                    'Ok, check our change environment
                                                    If .yChangeEnvironments = ChangeEnvironmentType.ePlanetToSystem OrElse .yChangeEnvironments = ChangeEnvironmentType.eSystemToPlanet Then
                                                        goMsgSys.SendToPrimary(.GetObjectUpdateMsg(GlobalMessageCode.eRemoveObject, RemovalType.eChangingEnvironments))
                                                        ReDim yOutMsg(7)
                                                        System.BitConverter.GetBytes(GlobalMessageCode.eEntityChangeEnvironment).CopyTo(yOutMsg, 0)
                                                        System.BitConverter.GetBytes(.ObjectID).CopyTo(yOutMsg, 2)
                                                        System.BitConverter.GetBytes(.ObjTypeID).CopyTo(yOutMsg, 6)

                                                        .bForceAggressionTest = False

                                                        .ParentEnvir.RemoveEntity(.ServerIndex, CType(.yChangeEnvironments, RemovalType), True, False, -1)

                                                        goMsgSys.SendToPathfinding(yOutMsg)
                                                        bBreakoutEntity = True
                                                        Exit For
                                                        'ElseIf .yChangeEnvironments = ChangeEnvironmentType.eDocking Then
                                                        '	goMsgSys.SendToPrimary(.GetObjectUpdateMsg(GlobalMessageCode.eRemoveObject, RemovalType.eDocking))
                                                        '	ReDim yOutMsg(7)
                                                        '	System.BitConverter.GetBytes(GlobalMessageCode.eDockCommand).CopyTo(yOutMsg, 0)
                                                        '	.GetGUIDAsString.CopyTo(yOutMsg, 2)
                                                        '	.bForceAggressionTest = False
                                                        '	.ParentEnvir.RemoveEntity(.ServerIndex, CType(.yChangeEnvironments, RemovalType), True, False)
                                                        '	goMsgSys.SendToPrimary(yOutMsg)
                                                        '	bBreakoutEntity = True
                                                        '	Exit For
                                                    Else
                                                        .bDecelerating = True
                                                        goMsgSys.SendDecelerationImminentMsg(oEntity)
                                                    End If

                                                End If
                                            End If
										End If

										If lActualCyclesDecelerate > 43000 Then lActualCyclesDecelerate = 43000
										fAdditionalTravel += (.TotalVelocity * lActualCyclesDecelerate) + (fAcc * (lActualCyclesDecelerate * lActualCyclesDecelerate) * 0.5F)
										.TotalVelocity -= (fAcc * lActualCyclesDecelerate)
                                        If .TotalVelocity < 0 OrElse Single.IsNaN(.TotalVelocity) = True Then .TotalVelocity = 0
									Else
										fAdditionalTravel += (.TotalVelocity * lCyclesAccelerate)
									End If
								End If
							End If

							If fAdditionalTravel > fTotDist Then fAdditionalTravel = fTotDist
							If fAdditionalTravel < .TotalVelocity Then fAdditionalTravel = .TotalVelocity

							If bDoTurnCalc = True Then
								If fTotDist <> 0 Then
									.VelX = .LocAngle Mod 8	' (.LocAngle \ 10) Mod 8
									.VelX *= 0.125F
									.VelX *= fAdditionalTravel
									.VelZ = fAdditionalTravel - .VelX
								End If
							Else
								If fTotDist <> 0 Then
									.VelX = fAbsDiffX
									.VelX = .VelX / fTotDist
									.VelX *= fAdditionalTravel
									.VelZ = fAdditionalTravel - .VelX
								End If
							End If

							'So, at this point we have a VelX and VelZ
							If .VelX = 0 Then
								'Vertical
								'If .VelZ < 0 Then
								'	dVecAngleRads = gdHalfPie
								'Else
								dVecAngleRads = gdPieAndAHalf
								'End If
							ElseIf .VelZ = 0 Then
								'horizontal
								'If .VelX >= 0 Then
								dVecAngleRads = 0
								'Else : dVecAngleRads = gdPi
								'End If
							Else
								'angled
								dVecAngleRads = CSng(Math.Atan(Math.Abs(.VelZ / .VelX)))
								'If .VelX >= 0 AndAlso .VelZ >= 0 Then
								dVecAngleRads = gdTwoPie - dVecAngleRads
								'ElseIf .VelX < 0 AndAlso .VelZ >= 0 Then
								'	dVecAngleRads += gdPi
								'ElseIf .VelX < 0 AndAlso .VelZ < 0 Then
								'	dVecAngleRads = gdPi - dVecAngleRads
								'End If
							End If 
							Const fPiOver180 As Single = gdPi / 180.0F
							dVecAngleRads = (.LocAngle * 0.1F) * fPiOver180 - dVecAngleRads
							Dim fTmpCos As Single = CSng(Math.Cos(dVecAngleRads))
							Dim fTmpSin As Single = CSng(Math.Sin(dVecAngleRads))
							fTempVelX = (.VelX * fTmpCos) + (.VelZ * fTmpSin)
							fTempVelZ = -((.VelX * fTmpSin) - (.VelZ * fTmpCos))
							.VelX = fTempVelX
							.VelZ = fTempVelZ

							'.bMovedThisFuelCycle = .bMovedThisFuelCycle OrElse .VelX <> 0.0F OrElse .VelZ <> 0.0F

                            If CInt(.MoveLocX) <> .DestX Then  '.MoveLocX <> .DestX Then
                                If fAbsDiffX > Math.Abs(.VelX) Then
                                    .MoveLocX += .VelX
                                Else
                                    .MoveLocX = .DestX
                                    .LocX = .DestX
                                End If
                            Else
                                'If .VelZ < 0 Then .VelZ = -fAdditionalTravel Else .VelZ = fAdditionalTravel
                                If .DestZ < .LocZ Then
                                    .MoveLocZ -= (fAdditionalTravel - Math.Abs(.VelZ))
                                Else : .MoveLocZ += (fAdditionalTravel - Math.Abs(.VelZ))
                                End If
                            End If
                            If CInt(.MoveLocZ) <> .DestZ Then  '.MoveLocZ <> .DestZ Then	
                                If fAbsDiffZ > Math.Abs(.VelZ) Then
                                    .MoveLocZ += .VelZ
                                Else
                                    .MoveLocZ = .DestZ
                                    .LocZ = .DestZ
                                End If
                            Else
                                'If .VelX < 0 Then .VelX = -.TotalVelocity Else .VelX = .TotalVelocity
                                'If .VelX < 0 Then
                                '    .MoveLocX -= (fAdditionalTravel - Math.Abs(.VelX))
                                'Else : .MoveLocX += (fAdditionalTravel - Math.Abs(.VelX))
                                'End If
                                'If .DestX < .LocX Then .MoveLocX -= (fAdditionalTravel - Math.Abs(.VelX)) Else .MoveLocX += (fAdditionalTravel - Math.Abs(.VelX))
                                If .DestX < .LocX Then
                                    .MoveLocX -= (fAdditionalTravel - Math.Abs(.VelX))
                                Else : .MoveLocX += (fAdditionalTravel - Math.Abs(.VelX))
                                End If
                            End If

							'This handles Map Wrap X (East West).
							' The pathfinding server will send us a coordinate outside of the bounds
							' of the environment to move to so the movement engine will work properly.
							' So, once we exceed the X extent of the environment (assuming we are on
							' a planet) we adjust our location and our destination to compensate for
							' the map wrap event.
							If .MoveLocX < .ParentEnvir.lMinPosX Then
                                'If .ParentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                                .MoveLocX = .ParentEnvir.lMinPosX
                                If .DestX < .ParentEnvir.lMinPosX Then .DestX = .ParentEnvir.lMinPosX
                                '.DestX = .ParentEnvir.lMinPosX
                                'Else
                                '	.MoveLocX += .ParentEnvir.lMapWrapAdjustX
                                '	.DestX += .ParentEnvir.lMapWrapAdjustX
                                'End If
							ElseIf .MoveLocX > .ParentEnvir.lMaxPosX Then
                                'If .ParentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                                .MoveLocX = .ParentEnvir.lMaxPosX
                                If .DestX > .ParentEnvir.lMaxPosX Then .DestX = .ParentEnvir.lMaxPosX
                                'Else
                                '	.MoveLocX -= .ParentEnvir.lMapWrapAdjustX
                                '	.DestX -= .ParentEnvir.lMapWrapAdjustX
                                'End If
							End If

							'  However, for now, in all cases, we just restrict movement beyond the extents
							If .MoveLocZ < .ParentEnvir.lMinPosZ Then
								.MoveLocZ = .ParentEnvir.lMinPosZ
                                If .DestZ < .ParentEnvir.lMinPosZ Then .DestZ = .ParentEnvir.lMinPosZ
							ElseIf .MoveLocZ > .ParentEnvir.lMaxPosZ Then
								.MoveLocZ = .ParentEnvir.lMaxPosZ
                                If .DestZ > .ParentEnvir.lMaxPosZ Then .DestZ = .ParentEnvir.lMaxPosZ
							End If

							'Now, check for total stop, if we are closer than the acceleration, then stop us
							'  my movelocx and z have changed so I can no longer use fAbsDiffX and Z
							'If bTurned = False OrElse ((iTAx100) > Math.Abs(.LocAngle - .DestAngle)) Then
							'	If fTotDist > fDistToStop Then
							Dim fLimiter As Single = Math.Max(0.5F, fAcc)
							If System.Math.Abs(.DestX - .MoveLocX) < fLimiter OrElse (.TotalVelocity < fAcc AndAlso .bDecelerating = True) Then	'OrElse lCyclesToStop < 2 Then ' fAcc Then ' oTmpDef.BaseAcceleration Then ' oTmpDef.Acceleration Then
								.MoveLocX = .DestX
								.LocX = .DestX
								.VelX = 0
							Else : .LocX = CInt(.MoveLocX)
							End If
							If System.Math.Abs(.DestZ - .MoveLocZ) < fLimiter OrElse (.TotalVelocity < fAcc AndAlso .bDecelerating = True) Then	'OrElse lCyclesToStop < 2 Then ' fAcc Then ' oTmpDef.BaseAcceleration Then ' oTmpDef.Acceleration Then
								.MoveLocZ = .DestZ
								.LocZ = .DestZ
								.VelZ = 0
							Else : .LocZ = CInt(.MoveLocZ)
							End If

							'Now, one final check, did our quadrant change?
							Dim fAfterDiffX As Single = .DestX - .LocX
							Dim fAfterDiffZ As Single = .DestZ - .LocZ
							If (fStartDiffX > 0 AndAlso fAfterDiffX < 0) OrElse (fStartDiffX < 0 AndAlso fAfterDiffX > 0) OrElse _
							   (fStartDiffZ > 0 AndAlso fAfterDiffZ < 0) OrElse (fStartDiffZ < 0 AndAlso fAfterDiffZ > 0) Then
								.LocX = .DestX
								.LocZ = .DestZ
								.MoveLocX = .LocX
								.MoveLocZ = .LocZ
								.VelX = 0
								.VelZ = 0
								.TotalVelocity = 0
							End If

                            If .bDecelerating = True AndAlso .TotalVelocity >= fInitialVel + (.Acceleration * 0.5F) Then
                                If Math.Abs(.MoveLocX - .DestX) < 150 AndAlso Math.Abs(.MoveLocZ - .DestZ) < 150 Then
                                    .LocX = .DestX
                                    .LocZ = .DestZ
                                    .MoveLocX = .LocX
                                    .MoveLocZ = .LocZ
                                    .TotalVelocity = 0
                                End If
                            End If

							'now, set our LocX and LocZ (the Int32 versions to our MoveLocX and MoveLocZ
							.LocY = 0	'we set Y to 0 here... if LOS needs it or something else, we calculate it then

							'Ok, now, check bResetManeuver
							If bResetManeuver = True Then
								If .lPrimaryTargetServerIdx <> -1 Then
                                    If glEntityIdx(.lPrimaryTargetServerIdx) > 0 Then
                                        SetNextManeuverPoint(oEntity, oTmpDef.lModelRangeOffset)
                                    Else : .lPrimaryTargetServerIdx = -1
                                    End If
								End If
							End If

							'Finally, get our destination angle  
							If bTurned = True Then
								'Was .MoveLocX <> .DestX Or .MoveLocZ <> .DestZ
								If .LocX <> .DestX OrElse .LocZ <> .DestZ Then
									'TODO: (Future) May want to make this inline...
									.DestAngle = CShort(LineAngleDegrees(.LocX, .LocZ, .DestX, .DestZ)) * 10S
								End If
							End If

							'We set this here (while moving) and also when the unit receives orders to move (to ensure no time-warping occurs)
							.LastCycleMoved = glCurrentCycle

							'Do a check now to see if we have reached our destination
							'If .DestX = .MoveLocX AndAlso .DestZ = .MoveLocZ AndAlso .DestAngle = .LocAngle Then
							If .DestX = .LocX AndAlso .DestZ = .LocZ AndAlso .DestAngle = .LocAngle Then
								.TotalVelocity = 0
								.VelX = 0
								.VelZ = 0
								.MoveLocX = .LocX
								.MoveLocZ = .LocZ
								If .LocAngle <> .TrueDestAngle AndAlso .TrueDestAngle <> -1 AndAlso .yChangeEnvironments = 0 Then
									.DestAngle = .TrueDestAngle
									.TrueDestAngle = -1
								Else
									'Ok, we have reached our destination, we need to generate a message to send
									'  to the pathfinding server and the environment's connected players
									'If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then .CurrentStatus -= elUnitStatus.eUnitMoving
									If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then .CurrentStatus = .CurrentStatus Xor elUnitStatus.eUnitMoving
									'If (.CurrentStatus And elUnitStatus.eMovedByPlayer) <> 0 Then .CurrentStatus -= elUnitStatus.eMovedByPlayer
									If (.CurrentStatus And elUnitStatus.eMovedByPlayer) <> 0 Then
										.CurrentStatus = .CurrentStatus Xor elUnitStatus.eMovedByPlayer
										'.lTetherPointX = .LocX
                                        '.lTetherPointZ = .LocZ
									End If
									'RemoveEntityMoving(.ServerIndex, .ObjectID)
                                    'SyncLockMovementRegisters(MovementCommand.RemoveEntityMoving, .ServerIndex, .ObjectID)
									.yResetMoveByPlayer = 0

									'gfrmDisplayForm.AddEventLine("Unit Stopped")

									'TODO: (Future) May want to create a QUEUE and call a FLUSH command at the end of process
                                    If .yChangeEnvironments <> 0 AndAlso .yChangeEnvironments <> ChangeEnvironmentType.eDocking Then
                                        SyncLockMovementRegisters(MovementCommand.RemoveEntityMoving, .ServerIndex, .ObjectID)
#If EXTENSIVELOGGING = 1 Then
                                gfrmDisplayForm.AddEventLine("ChangeEnvironments: " & .ObjectID & ", " & .ObjTypeID & " changing environments type: " & .yChangeEnvironments)
#End If
                                        '1) Send off our save message to Primary...
                                        If .yChangeEnvironments = ChangeEnvironmentType.eDocking Then
                                            goMsgSys.SendToPrimary(.GetObjectUpdateMsg(GlobalMessageCode.eRemoveObject, RemovalType.eDocking))
                                        Else
                                            goMsgSys.SendToPrimary(.GetObjectUpdateMsg(GlobalMessageCode.eRemoveObject, RemovalType.eChangingEnvironments))
                                        End If

                                        '3) Send Change Environment Complete for entity
                                        If .yChangeEnvironments = ChangeEnvironmentType.eDocking Then
                                            ReDim yOutMsg(7)
                                            System.BitConverter.GetBytes(GlobalMessageCode.eDockCommand).CopyTo(yOutMsg, 0)
                                            .GetGUIDAsString.CopyTo(yOutMsg, 2)
                                            'goMsgSys.SendToPrimary(yOutMsg)
                                            .bForceAggressionTest = False
                                        Else
                                            ReDim yOutMsg(7)
                                            System.BitConverter.GetBytes(GlobalMessageCode.eEntityChangeEnvironment).CopyTo(yOutMsg, 0)
                                            System.BitConverter.GetBytes(.ObjectID).CopyTo(yOutMsg, 2)
                                            System.BitConverter.GetBytes(.ObjTypeID).CopyTo(yOutMsg, 6)
                                            'goMsgSys.SendToPathfinding(yOutMsg)
                                            .bForceAggressionTest = False
                                        End If

                                        .ParentEnvir.RemoveEntity(.ServerIndex, CType(.yChangeEnvironments, RemovalType), True, False, -1)

                                        If System.BitConverter.ToInt16(yOutMsg, 0) = GlobalMessageCode.eDockCommand Then
                                            goMsgSys.SendToPrimary(yOutMsg)
                                        Else : goMsgSys.SendToPathfinding(yOutMsg)
                                        End If
                                        bBreakoutEntity = True
                                        Exit For
                                    Else

                                        yOutMsg = goMsgSys.CreateStopObjectCommand(oEntity)
                                        goMsgSys.SendToPathfinding(yOutMsg)
                                        If .ParentEnvir.lPlayersInEnvirCnt > 0 Then goMsgSys.BroadcastToEnvironmentClients(yOutMsg, .ParentEnvir)
                                        'RemoveEntityMoving(.ServerIndex, .ObjectID)
                                        .bForceAggressionTest = True

                                        'If .TrueDestAngle <> -1 AndAlso .LocAngle <> .TrueDestAngle Then
                                        '	.DestAngle = .TrueDestAngle
                                        '	AddEntityMoving(.ServerIndex, .ObjectID)

                                        '	If .ParentEnvir.lPlayersInEnvirCnt > 0 Then
                                        '		Dim yData(17) As Byte
                                        '		System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectCommand).CopyTo(yData, 0)
                                        '		System.BitConverter.GetBytes(.ObjectID).CopyTo(yData, 2)
                                        '		System.BitConverter.GetBytes(.ObjTypeID).CopyTo(yData, 6)
                                        '		System.BitConverter.GetBytes(.DestX).CopyTo(yData, 8)
                                        '		System.BitConverter.GetBytes(.DestZ).CopyTo(yData, 12)
                                        '		System.BitConverter.GetBytes(CShort(lTemp)).CopyTo(yData, 16)

                                        '		goMsgSys.BroadcastToEnvironmentClients(yData, .ParentEnvir)
                                        '	End If
                                        'End If

                                    End If
								End If
							End If
							If bBreakoutEntity = True Then Continue For

							'If .ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
							'    lSmallPerRow = gl_PLANET_SMALL_PER_ROW
							'Else
							'    lSmallPerRow = gl_SYSTEM_SMALL_PER_ROW
							'End If
							lSmallPerRow = gl_SMALL_PER_ROW

							lTemp = .LocZ + .ParentEnvir.lHalfEnvirSize
							'lLargeSector = CInt(Math.Floor(lTemp / .ParentEnvir.lGridSquareSize)) * .ParentEnvir.lGridsPerRow
							lLargeSector = (lTemp \ .ParentEnvir.lGridSquareSize) * .ParentEnvir.lGridsPerRow

							'lTemp -= CInt((lLargeSector / .ParentEnvir.lGridsPerRow) * .ParentEnvir.lGridSquareSize)
							lTemp -= ((lLargeSector \ .ParentEnvir.lGridsPerRow) * .ParentEnvir.lGridSquareSize)

							'lSmallSector = CInt(Math.Floor(lTemp / gl_SMALL_GRID_SQUARE_SIZE)) * lSmallPerRow
							lSmallSector = ((lTemp \ gl_SMALL_GRID_SQUARE_SIZE) * lSmallPerRow)

							'lTemp -= CInt((lSmallSector / lSmallPerRow) * gl_SMALL_GRID_SQUARE_SIZE)
							lTemp -= ((lSmallSector \ lSmallPerRow) * gl_SMALL_GRID_SQUARE_SIZE)

							'lRelTinyZ = CInt(Math.Floor(lTemp / gl_FINAL_GRID_SQUARE_SIZE))
							lRelTinyZ = lTemp \ gl_FINAL_GRID_SQUARE_SIZE

							lTemp2 = .LocX + .ParentEnvir.lHalfEnvirSize
							'lTemp = CInt(Math.Floor(lTemp2 / .ParentEnvir.lGridSquareSize))
							lTemp = lTemp2 \ .ParentEnvir.lGridSquareSize

							lLargeSector += lTemp
							lTemp2 -= (lTemp * .ParentEnvir.lGridSquareSize)

							'lTemp = CInt(Math.Floor(lTemp2 / gl_SMALL_GRID_SQUARE_SIZE))
							lTemp = lTemp2 \ gl_SMALL_GRID_SQUARE_SIZE

							lSmallSector += lTemp
							lTemp2 -= (lTemp * gl_SMALL_GRID_SQUARE_SIZE)

							'lRelTinyX = CInt(Math.Floor(lTemp2 / gl_FINAL_GRID_SQUARE_SIZE))
							lRelTinyX = lTemp2 \ gl_FINAL_GRID_SQUARE_SIZE

						Else
							'for the aggression test...
							'lSector = .lTinySectorID
							lRelTinyX = .lTinyX
							lRelTinyZ = .lTinyZ
							lLargeSector = .lGridIndex
							lSmallSector = .lSmallSectorID
                            'RemoveEntityMoving(.ServerIndex, .ObjectID)

                            If (.CurrentStatus And elUnitStatus.eUnitMoving) = 0 AndAlso .bForceAggressionTest = False Then
                                SyncLockMovementRegisters(MovementCommand.RemoveEntityMoving, .ServerIndex, .ObjectID)
                            End If
						End If	'if we are moving

						'And now our aggression test...
						'If (lLargeSector = .lGridIndex AndAlso lSmallSector = .lSmallSectorID AndAlso lSector = .lTinySectorID) = False Then 'Or .bForceAggressionTest = True Then
						'If lLargeSector <> .lGridIndex Or lSmallSector <> .lSmallSectorID Or lSector <> .lTinySectorID Then 'Or .bForceAggressionTest = True Then
                        If (.CurrentStatus And elUnitStatus.eUnitCloaked) = 0 AndAlso .yChangeEnvironments = 0 AndAlso (.ParentEnvir.bEnvirAtColonyLimit = True OrElse (.Owner.lIronCurtainPlanetID <> Int32.MinValue AndAlso (.ParentEnvir.ObjTypeID <> ObjectType.ePlanet OrElse .ParentEnvir.ObjectID <> .Owner.lIronCurtainPlanetID))) AndAlso ((.bForceAggressionTest = True AndAlso (glCurrentCycle - .lLastForceAggressionTest > gl_FORCE_AGGRESSION_THRESHOLD)) OrElse lLargeSector <> .lGridIndex OrElse lSmallSector <> .lSmallSectorID OrElse lRelTinyX <> .lTinyX OrElse lRelTinyZ <> .lTinyZ) Then

                            If .bForceAggressionTest = True Then
                                .lLastForceAggressionTest = glCurrentCycle
                            End If

                            'now, if we didn't change any sectors, but instead entered from a force aggress test, then
                            bSkipReverse = .bForceAggressionTest = True AndAlso .bNewAddedEntity = False AndAlso (lLargeSector = .lGridIndex) AndAlso _
                              (lSmallSector = .lSmallSectorID) AndAlso (lRelTinyX = .lTinyX) AndAlso (lRelTinyZ = .lTinyZ)

                            'are we changing large?
                            If lLargeSector <> .lGridIndex Then
                                'yes, remove me from current large
                                .ParentEnvir.oGrid(.lGridIndex).RemoveEntity(.lGridEntityIdx)
                                'add mew to new

                                'If lLargeSector > .ParentEnvir.oGrid.GetUpperBound(0) Then
                                If lLargeSector > .ParentEnvir.GetGridUB OrElse lLargeSector < 0 Then
                                    gfrmDisplayForm.AddEventLine("NOTE: Large Sector exceeded grid UB! ObjID = " & .ObjectID & ", Type = " & .ObjTypeID & ", ParentID = " & .ParentEnvir.ObjectID & ", ParentType = " & .ParentEnvir.ObjTypeID & ", Invalid GridID: " & lLargeSector & ", LocX: " & .LocX & ", LocZ: " & .LocZ)
                                Else
                                    .lGridEntityIdx = .ParentEnvir.oGrid(lLargeSector).AddEntity(.ServerIndex)
                                    .lGridIndex = lLargeSector
                                End If

                                'Now, if the parent environment is a system, run a wormhole proximity test
                                If .ParentEnvir.ObjTypeID = ObjectType.eSolarSystem AndAlso (.CurrentStatus And elUnitStatus.eRadarOperational) <> 0 Then
                                    CType(.ParentEnvir.oGeoObject, SolarSystem).HandleWormholeProximityTest(oEntity)
                                End If
                            End If
                            .lSmallSectorID = lSmallSector
                            '.lTinySectorID = lSector
                            .lTinyX = lRelTinyX
                            .lTinyZ = lRelTinyZ

                            '2) Check for Aggression and 3) Target Acquisition
                            'On Error Resume Next
                            'Potential Aggression is set when a unit enters or leaves the environment
                            '  the server scans thru the environment for the players and compares their rels
                            '  to each other. If the rel is war, then PotentialAggression is true, otherwise its false
                            If .ParentEnvir.PotentialAggression = True AndAlso (.ObjTypeID <> ObjectType.eFacility OrElse (.CurrentStatus And elUnitStatus.eFacilityPowered) <> 0) Then
                                'NOTE: Not sure if I like this extra loop
                                Dim bCPMajorOverage As Boolean = False
                                Dim bCPOverage As Boolean = False
                                If .Owner Is Nothing = False Then 'AndAlso .Owner.PlayerHasWar() = True Then
                                    Dim lTmpOwnerID As Int32 = .Owner.ObjectID
                                    If lTmpOwnerID <> gl_HARDCODE_PIRATE_PLAYER_ID AndAlso .ObjTypeID <> ObjectType.eFacility Then
                                        For lTmpIdx As Int32 = 0 To .ParentEnvir.lPlayersWhoHaveUnitsHereUB
                                            If .ParentEnvir.lPlayersWhoHaveUnitsHereIdx(lTmpIdx) = lTmpOwnerID Then
                                                'We use 1500 as the CP Overage here, because we want to know if it is 1.5x higher
                                                bCPMajorOverage = (.ParentEnvir.lPlayersWhoHaveUnitsHereCP(lTmpIdx) > (1.5F * .Owner.lCPLimit))
                                                bCPOverage = (.ParentEnvir.lPlayersWhoHaveUnitsHereCP(lTmpIdx) > .Owner.lCPLimit)
                                                Exit For
                                            End If
                                        Next lTmpIdx
                                    End If
                                End If

                                Dim bEngagedAtStart As Boolean = (.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0

                                If (.iCombatTactics And eiBehaviorPatterns.eEngagement_Hold_Fire) = 0 Then
                                    'Ensure that our targets are updated
                                    .VerifyTargets()
                                    If oTmpDef.JamEffect <> 0 AndAlso oTmpDef.JamStrength <> 0 Then .ClearJamTargets()

                                    'If the entity already has all of its targets, then we skip it...
                                    'If ((.CurrentStatus And elUnitStatus.eUnitEngaged) = 0) OrElse _
                                    '   ((.CurrentStatus And elUnitStatus.eSide1HasTarget) = 0) OrElse _
                                    '   ((.CurrentStatus And elUnitStatus.eSide2HasTarget) = 0) OrElse _
                                    '   ((.CurrentStatus And elUnitStatus.eSide3HasTarget) = 0) OrElse _
                                    '   ((.CurrentStatus And elUnitStatus.eSide4HasTarget) = 0) Then
                                    If True = True Then

                                        'If .ObjTypeID = ObjectType.eUnit AndAlso .ObjectID = 2703 Then Stop
                                        gblAggressions += 1
                                        'check if the environment currently has warring factions
                                        'Ok, now check the unit's ranges to other units in the environment...
                                        lPrimaryTargetScore = Int32.MinValue
                                        lPrimaryFacing = 0
                                        For Y = 0 To 3
                                            lFaceTargetScores(Y) = Int32.MinValue
                                            lFighterTargetScores(Y) = Int32.MinValue
                                        Next Y
                                        oPrimaryTarget = Nothing

                                        If (.CurrentStatus And elUnitStatus.eTargetingByPlayer) <> 0 Then
                                            If .lPrimaryTargetServerIdx = -1 OrElse glEntityIdx(.lPrimaryTargetServerIdx) < 1 Then
                                                'No longer targeting that target because it no longer exists
                                                '.CurrentStatus -= elUnitStatus.eTargetingByPlayer
                                                .CurrentStatus = .CurrentStatus Xor elUnitStatus.eTargetingByPlayer
                                                .lPrimaryTargetServerIdx = -1
                                            Else
                                                oPrimaryTarget = goEntity(.lPrimaryTargetServerIdx)
                                            End If
                                        End If

                                        'Determine what adjust array to use (for map wrapping of planets)
                                        Dim lAdjust() As Int32 = .ParentEnvir.lGridIdxAdjust
                                        If .ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
                                            Dim lModVal As Int32 = .lGridIndex Mod .ParentEnvir.lGridsPerRow
                                            If lModVal = 0 Then
                                                'I'm on the left edge, use the left edge adjust
                                                lAdjust = .ParentEnvir.lLeftEdgeGridIdxAdjust
                                            ElseIf lModVal = .ParentEnvir.lGridsPerRow - 1 Then
                                                'I'm on the right edge, use the right edge adjust
                                                lAdjust = .ParentEnvir.lRightEdgeGridIdxAdjust
                                            End If
                                        End If

                                        For lGrid = 0 To lAdjust.GetUpperBound(0)
                                            'gonna reuse the llargesector here...
                                            'lLargeSector = .lGridIndex + .ParentEnvir.lGridIdxAdjust(lGrid)
                                            lLargeSector = .lGridIndex + lAdjust(lGrid)

                                            If lLargeSector > -1 AndAlso lLargeSector < .ParentEnvir.lGridUB + 1 Then
                                                'Ok, now, cycle thru the smaller list in the Grid object
                                                Dim oTmpGrid As EnvirGrid = .ParentEnvir.oGrid(lLargeSector)
                                                If oTmpGrid Is Nothing Then Continue For
                                                For Y = 0 To oTmpGrid.lEntityUB
                                                    lTemp = oTmpGrid.lEntities(Y)
                                                    If lTemp <> -1 Then      'is the unit not nothing?
                                                        oTmpUnit = goEntity(lTemp)
                                                        If oTmpUnit Is Nothing Then
                                                            oTmpGrid.lEntities(Y) = -1
                                                            Continue For
                                                        End If

                                                        If oTmpUnit.ParentEnvir.ObjectID <> .ParentEnvir.ObjectID OrElse oTmpUnit.ParentEnvir.ObjTypeID <> .ParentEnvir.ObjTypeID Then
                                                            oTmpGrid.RemoveEntity(Y)
                                                            oTmpUnit.lGridEntityIdx = oTmpUnit.ParentEnvir.oGrid(oTmpUnit.lGridIndex).AddEntity(oTmpUnit.ServerIndex)
                                                            Continue For
                                                        End If

                                                        If oTmpUnit.ParentEnvir.bEnvirAtColonyLimit = False AndAlso (oTmpUnit.Owner.lIronCurtainPlanetID = Int32.MinValue OrElse (oTmpUnit.ParentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oTmpUnit.Owner.lIronCurtainPlanetID = oTmpUnit.ParentEnvir.ObjectID)) Then
                                                            'that player is in an iron curtain
                                                            Continue For
                                                        End If

                                                        'Exclude myself
                                                        If (oTmpUnit.ObjectID <> .ObjectID OrElse oTmpUnit.ObjTypeID <> .ObjTypeID) AndAlso (oTmpUnit.CurrentStatus And elUnitStatus.eUnitCloaked) = 0 Then
                                                            'check the range to target
                                                            'TODO: Need to take into account DetectionResist and DisruptionResist
                                                            'Perhaps, use a different mechanism? A UnitSignature if you will...
                                                            '  the signature determines how easy radar can detect the unit...
                                                            '  this signature is a final number represented by a mathematical
                                                            '  equation based on the unit's Size and DetectionResist
                                                            'Kinda like saying a unit that is 10000 hull and 1000 units of distance
                                                            '  away is just as easy to 'detect' as a unit that is 100000 hull and
                                                            '  10000 units of distance away... have to think about this one
                                                            lEnvirDist = Int32.MaxValue
                                                            lTemp = giRelativeSmall(lGrid, .lSmallSectorID, oTmpUnit.lSmallSectorID)
                                                            If lTemp <> -1 Then
                                                                lRelTinyX = glBaseRelTinyX(lTemp) + oTmpUnit.lTinyX + (gl_HALF_FINAL_PER_ROW - .lTinyX)
                                                                If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                                                    lRelTinyZ = glBaseRelTinyZ(lTemp) + oTmpUnit.lTinyZ + (gl_HALF_FINAL_PER_ROW - .lTinyZ)
                                                                    If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                                                        lEnvirDist = gyDistances(lRelTinyX, lRelTinyZ) - oTmpUnit.lModelRangeOffset
                                                                    End If
                                                                End If
                                                            End If

                                                            If (.CurrentStatus And elUnitStatus.eRadarOperational) <> 0 Then
                                                                If lEnvirDist <= oTmpDef.lOptRadarRange + .lJamRangeMod AndAlso ((oTmpUnit.CurrentStatus And elUnitStatus.eUnitCloaked) = 0) Then
                                                                    'Ok, the unit is visible... NOW we do a test for aggression *grin*
                                                                    If bCPMajorOverage = True AndAlso oTmpUnit.ObjTypeID = ObjectType.eUnit Then
                                                                        yRelID = elRelTypes.eWar
                                                                    Else : yRelID = .Owner.GetPlayerRelScore(oTmpUnit.lOwnerID, True, oTmpUnit.ServerIndex)
                                                                    End If
                                                                    'yRelID should never equal 0... we should limit it to 1
                                                                    '  now, if yRelID is 0, then this player has no relationship
                                                                    '  but we don't really care about the relationship, we care about war
                                                                    If yRelID <= elRelTypes.eWar AndAlso yRelID > 0 Then
                                                                        'If .lOwnerID = oTmpUnit.lOwnerID Then Stop

                                                                        'Ok, the player is at war with this unit... aggress...
                                                                        ' to Aggress, we also determine our target...
                                                                        'This means we should continue through the list of units... but store
                                                                        ' this value if we need to...
                                                                        lFacing = WhatSideCanFire(oEntity, lRelTinyX, lRelTinyZ)

                                                                        oTmpDef2 = Nothing
                                                                        If oTmpUnit.lEntityDefServerIndex <> -1 Then
                                                                            If glEntityDefIdx(oTmpUnit.lEntityDefServerIndex) <> -1 Then
                                                                                oTmpDef2 = goEntityDefs(oTmpUnit.lEntityDefServerIndex)
                                                                            End If
                                                                        End If
                                                                        If oTmpDef2 Is Nothing Then Continue For

                                                                        If .ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
                                                                            lTemp = GetPrimaryTargetScore(oEntity, oTmpUnit, lEnvirDist, oTmpDef2.PlanetTacticalAnalysis)
                                                                        Else : lTemp = GetPrimaryTargetScore(oEntity, oTmpUnit, lEnvirDist, oTmpDef2.SpaceTacticalAnalysis)
                                                                        End If

                                                                        If oTmpUnit.ServerIndex = .lPrimaryTargetServerIdx AndAlso (.CurrentStatus And elUnitStatus.eTargetingByPlayer) <> 0 Then
                                                                            lTemp = Int32.MaxValue
                                                                        End If

                                                                        If lTemp > lFaceTargetScores(lFacing) Then
                                                                            .lTargetsServerIdx(lFacing) = oTmpUnit.ServerIndex
                                                                            'If oTmpUnit.lOwnerID = .lOwnerID Then
                                                                            '	gfrmDisplayForm.AddEventLine("Shooting at myself: HandleMovement.AggressionTest.TargetServerIdx")
                                                                            'End If
                                                                            lFaceTargetScores(lFacing) = lTemp
                                                                            .Ranges(lFacing) = lEnvirDist
                                                                        End If
                                                                        If (oTmpDef2.SpaceTacticalAnalysis And eiTacticalAttrs.eFighterClass) <> 0 AndAlso lTemp > lFighterTargetScores(lFacing) Then
                                                                            .lFighterTargetServerIdx(lFacing) = oTmpUnit.ServerIndex
                                                                            lFighterTargetScores(lFacing) = lTemp
                                                                            .lFighterTargetRange(lFacing) = lEnvirDist
                                                                        End If
                                                                        'need to check if the player assigned the target...
                                                                        If ((.CurrentStatus And elUnitStatus.eTargetingByPlayer) = 0) Then
                                                                            If lTemp > lPrimaryTargetScore Then
                                                                                .lPrimaryTargetServerIdx = oTmpUnit.ServerIndex
                                                                                'If oTmpUnit.lOwnerID = .lOwnerID Then
                                                                                '	gfrmDisplayForm.AddEventLine("Shooting at myself: HandleMovement.AggressionTest.PrimaryTargetServerIdx")
                                                                                'End If
                                                                                lPrimaryTargetScore = lTemp
                                                                                oPrimaryTarget = oTmpUnit
                                                                                lPrimaryFacing = lFacing
                                                                                .CurrentStatus = .CurrentStatus Or elUnitStatus.eUnitEngaged
                                                                                AddEntityInCombat(.ObjectID, .ObjTypeID, .ServerIndex)
                                                                                .lNextFireWpnEvent = 0

                                                                                'MSC - 12/17/07 - Put this in to split up weapons fire a bit better
                                                                                If .lWpnNextFireCycle Is Nothing = False Then
                                                                                    For Z As Int32 = 0 To .lWpnNextFireCycle.GetUpperBound(0)
                                                                                        If .lWpnNextFireCycle(Z) < glCurrentCycle Then
                                                                                            '.lWpnNextFireCycle(Z) = glCurrentCycle + CInt(Rnd() * 10)
                                                                                            .lWpnNextFireCycle(Z) = glCurrentCycle + CInt(Rnd() * 10) '+ oTmpDef.WeaponDefs(Z).ROF_Delay
                                                                                        End If
                                                                                    Next Z
                                                                                End If

                                                                                'AddEntityCombat(.ServerIndex, .ObjectID)
                                                                            End If
                                                                        ElseIf .lPrimaryTargetServerIdx = oTmpUnit.ServerIndex Then
                                                                            lPrimaryFacing = lFacing
                                                                            .CurrentStatus = .CurrentStatus Or elUnitStatus.eUnitEngaged
                                                                            AddEntityInCombat(.ObjectID, .ObjTypeID, .ServerIndex)
                                                                            If .lWpnNextFireCycle Is Nothing = False Then
                                                                                For Z As Int32 = 0 To .lWpnNextFireCycle.GetUpperBound(0)
                                                                                    If .lWpnNextFireCycle(Z) < glCurrentCycle Then
                                                                                        .lWpnNextFireCycle(Z) = glCurrentCycle + CInt(Rnd() * 10) '+ oTmpDef.WeaponDefs(Z).ROF_Delay
                                                                                    End If
                                                                                Next Z
                                                                            End If
                                                                        End If
                                                                        If oTmpDef.JamEffect <> 0 AndAlso oTmpDef.JamStrength <> 0 Then .AddJamTarget(oTmpUnit.ServerIndex, lTemp)
                                                                    ElseIf oTmpUnit.lOwnerID = .lOwnerID AndAlso bEngagedAtStart = True Then
                                                                        If (oTmpUnit.CurrentStatus And (elUnitStatus.eUnitMoving Or elUnitStatus.eUnitEngaged Or elUnitStatus.eUnitCloaked Or elUnitStatus.eMoveLock Or elUnitStatus.eTargetingByPlayer)) = 0 AndAlso _
                                                                          (oTmpUnit.CurrentStatus And elUnitStatus.eRadarOperational) <> 0 AndAlso _
                                                                          (oTmpUnit.iCombatTactics And (eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eEngagement_Pursue)) <> 0 Then
                                                                            bEngagedAtStart = False
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If

                                                            ''Now, check in reverse... the distance is still the same, just need to check 
                                                            ''  other unit's radar range
                                                            If bSkipReverse = False Then
                                                                If lEnvirDist <> Int32.MaxValue Then
                                                                    If ((oTmpUnit.iCombatTactics And eiBehaviorPatterns.eEngagement_Hold_Fire) = 0) Then
                                                                        'If (oTmpUnit.CurrentStatus And elUnitStatus.eUnitEngaged) = 0 Then
                                                                        'TODO: Once again, take into account disruptionresist, etc...
                                                                        If (oTmpUnit.CurrentStatus And elUnitStatus.eRadarOperational) <> 0 Then

                                                                            'Check for iron curtain
                                                                            If oTmpUnit.ParentEnvir.bEnvirAtColonyLimit = True OrElse (oTmpUnit.Owner.lIronCurtainPlanetID <> Int32.MinValue AndAlso (oTmpUnit.ParentEnvir.ObjTypeID <> ObjectType.ePlanet OrElse oTmpUnit.Owner.lIronCurtainPlanetID <> oTmpUnit.ParentEnvir.ObjectID)) Then
                                                                                If oTmpUnit.lEntityDefServerIndex <> -1 Then
                                                                                    If glEntityDefIdx(oTmpUnit.lEntityDefServerIndex) <> -1 Then
                                                                                        oTmpDef2 = goEntityDefs(oTmpUnit.lEntityDefServerIndex)
                                                                                    End If
                                                                                End If

                                                                                If oTmpDef2 Is Nothing Then Continue For

                                                                                If lEnvirDist <= oTmpDef2.lOptRadarRange + oTmpUnit.lJamRangeMod Then
                                                                                    'Ok, the moving unit is visible... Now, we need to test for aggression
                                                                                    yRelID = oTmpUnit.Owner.GetPlayerRelScore(.lOwnerID, True, .ServerIndex)
                                                                                    If yRelID <= elRelTypes.eWar AndAlso yRelID > 0 Then
                                                                                        'Ok, aggression time, same reason... if a unit is able
                                                                                        '  to reverse aggress, we simply force it to aggress next
                                                                                        '  cycle... so it picks the best target based on its script
                                                                                        oTmpUnit.bForceAggressionTest = True
                                                                                        SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oTmpUnit.ServerIndex, oTmpUnit.ObjectID) 'AddEntityMoving(oTmpUnit.ServerIndex, oTmpUnit.ObjectID)
                                                                                    End If
                                                                                End If
                                                                                oTmpDef2 = Nothing
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If      'if unit is nothing
                                                Next Y
                                            End If
                                        Next lGrid

                                        'By this point, I know whether or not I engaged
                                        If (.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then

#If EXTENSIVELOGGING = 1 Then
                                    .DisplayTargetMsg()
#End If

                                            'This block of code is Flocking... However, it has not been tested and should be tested before being released
                                            'MSC - 11/16/08 - added this but has not been released (will make the next region release) - if the unit was not engaged at start but is now, then flock
                                            'If bEngagedAtStart = True AndAlso bCPOverage = False Then
                                            If bEngagedAtStart = False AndAlso bCPOverage = False Then
                                                'ok, I've engaged, see if any units flock with me...
                                                For lGrid = 0 To lAdjust.GetUpperBound(0)
                                                    'gonna reuse the llargesector here...
                                                    'lLargeSector = .lGridIndex + .ParentEnvir.lGridIdxAdjust(lGrid)
                                                    lLargeSector = .lGridIndex + lAdjust(lGrid)

                                                    If lLargeSector > -1 AndAlso lLargeSector < .ParentEnvir.lGridUB + 1 Then
                                                        'Ok, now, cycle thru the smaller list in the Grid object
                                                        Dim oTmpGrid As EnvirGrid = .ParentEnvir.oGrid(lLargeSector)
                                                        If oTmpGrid Is Nothing Then Continue For
                                                        For Y = 0 To oTmpGrid.lEntityUB
                                                            lTemp = oTmpGrid.lEntities(Y)
                                                            If lTemp <> -1 Then      'is the unit not nothing?
                                                                oTmpUnit = goEntity(lTemp)
                                                                If oTmpUnit Is Nothing Then
                                                                    oTmpGrid.lEntities(Y) = -1
                                                                    Continue For
                                                                End If

                                                                If oTmpUnit.ParentEnvir.ObjectID <> .ParentEnvir.ObjectID OrElse oTmpUnit.ParentEnvir.ObjTypeID <> .ParentEnvir.ObjTypeID Then
                                                                    oTmpGrid.RemoveEntity(Y)
                                                                    oTmpUnit.lGridEntityIdx = oTmpUnit.ParentEnvir.oGrid(oTmpUnit.lGridIndex).AddEntity(oTmpUnit.ServerIndex)
                                                                    Continue For
                                                                End If

                                                                If (oTmpUnit.CurrentStatus And (elUnitStatus.eUnitMoving Or elUnitStatus.eUnitEngaged Or elUnitStatus.eUnitCloaked Or elUnitStatus.eMoveLock Or elUnitStatus.eTargetingByPlayer)) <> 0 Then Continue For
                                                                If (oTmpUnit.CurrentStatus And elUnitStatus.eRadarOperational) = 0 Then Continue For
                                                                If oTmpUnit.lOwnerID <> .lOwnerID AndAlso (oTmpUnit.Owner.lGuildID < 1 OrElse oTmpUnit.Owner.lGuildID <> .Owner.lGuildID) Then Continue For
                                                                If oTmpUnit.ObjectID = .ObjectID AndAlso oTmpUnit.ObjTypeID = .ObjTypeID Then Continue For
                                                                If (oTmpUnit.iCombatTactics And (eiBehaviorPatterns.eEngagement_Engage Or eiBehaviorPatterns.eEngagement_Pursue)) = 0 Then Continue For


                                                                lEnvirDist = Int32.MaxValue
                                                                lTemp = giRelativeSmall(lGrid, .lSmallSectorID, oTmpUnit.lSmallSectorID)
                                                                If lTemp <> -1 Then
                                                                    lRelTinyX = glBaseRelTinyX(lTemp) + oTmpUnit.lTinyX + (gl_HALF_FINAL_PER_ROW - .lTinyX)
                                                                    If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                                                        lRelTinyZ = glBaseRelTinyZ(lTemp) + oTmpUnit.lTinyZ + (gl_HALF_FINAL_PER_ROW - .lTinyZ)
                                                                        If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                                                            lEnvirDist = gyDistances(lRelTinyX, lRelTinyZ) - oTmpUnit.lModelRangeOffset
                                                                        End If
                                                                    End If
                                                                End If


                                                                If lEnvirDist <= oTmpDef.lOptRadarRange + .lJamRangeMod + oTmpUnit.lJamRangeMod Then
                                                                    goMsgSys.SendAIMoveRequestToPathfinding(oTmpUnit, .LocX, .LocZ, .LocAngle)
                                                                End If
                                                            End If
                                                        Next Y
                                                    End If
                                                Next lGrid
                                            End If


                                            'Ok, I found an engagement... therefore, I should also know a primary target
                                            'Ok, now, set the primary target

                                            If oPrimaryTarget Is Nothing = False Then
                                                Dim lTmpPTargetIdx As Int32 = .lPrimaryTargetServerIdx
                                                If (.CurrentStatus And elUnitStatus.eTargetingByPlayer) = 0 OrElse lTmpPTargetIdx = -1 OrElse glEntityIdx(lTmpPTargetIdx) < 1 Then
                                                    .lPrimaryTargetServerIdx = oPrimaryTarget.ServerIndex
                                                Else
                                                    If lTmpPTargetIdx <> oPrimaryTarget.ServerIndex Then
                                                        oPrimaryTarget = goEntity(lTmpPTargetIdx)
                                                    End If
                                                End If

                                                'If oPrimaryTarget.lOwnerID = .lOwnerID Then
                                                '	gfrmDisplayForm.AddEventLine("Shooting at my own guy Primary Target.")
                                                'End If
                                                If oPrimaryTarget.ParentEnvir.ObjectID <> .ParentEnvir.ObjectID OrElse oPrimaryTarget.ParentEnvir.ObjTypeID <> .ParentEnvir.ObjTypeID Then
                                                    gfrmDisplayForm.AddEventLine("Shooting outside my environment (primary)")
                                                End If

                                                lPrimaryFacing = 0
                                                For lTmpIdx As Int32 = 0 To 3
                                                    If .lTargetsServerIdx(lTmpIdx) = .lPrimaryTargetServerIdx Then
                                                        .lPrimaryTargetRange = .Ranges(lTmpIdx)
                                                        lPrimaryFacing = lTmpIdx
                                                        Exit For
                                                    End If

                                                    'If .lTargetsServerIdx(lTmpIdx) <> -1 AndAlso goEntity(.lTargetsServerIdx(lTmpIdx)).lOwnerID = .lOwnerID Then
                                                    '	gfrmDisplayForm.AddEventLine("Shooting at my own guy target server idx.")
                                                    'End If
                                                    If .lTargetsServerIdx(lTmpIdx) <> -1 AndAlso (goEntity(.lTargetsServerIdx(lTmpIdx)).ParentEnvir.ObjectID <> .ParentEnvir.ObjectID OrElse goEntity(.lTargetsServerIdx(lTmpIdx)).ParentEnvir.ObjTypeID <> .ParentEnvir.ObjTypeID) Then
                                                        gfrmDisplayForm.AddEventLine("Shooting outside my environment (target)")
                                                    End If
                                                Next
                                            End If

                                            'Now, check for extended movement options... if the unit was moved
                                            '  by a player OR the unit has the option Stand Ground, then we do
                                            '  not move the unit...

                                            If ((.CurrentStatus And elUnitStatus.eMovedByPlayer) = 0) Then 'AndAlso ((.iCombatTactics And eiBehaviorPatterns.eEngagement_Stand_Ground) = 0)
                                                'If .ObjectID = 5521 Then Stop
                                                'If oTmpDef.MaxSpeed > 0 AndAlso ((.CurrentStatus And elUnitStatus.eEngineOperational) <> 0) Then
                                                If .MaxSpeed > 0 AndAlso ((.CurrentStatus And elUnitStatus.eEngineOperational) <> 0) Then
                                                    'Ok, the unit has engines and they are operational... check our engagement pattern
                                                    If (.iCombatTactics And eiBehaviorPatterns.eEngagement_Engage) <> 0 Then
                                                        If (.CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
                                                            'we are not moving, by our orders, we need to get a preferred
                                                            '  range, and a preferred facing and then move us
                                                            lTemp = GetPreferredRange(oEntity)
                                                            'Ok, ltemp tells us preferred range, check that
                                                            If .lPrimaryTargetRange > lTemp Then
                                                                'Ok, we are out of our preferred range...
                                                                SetPreferredRangeDest(oEntity, oPrimaryTarget, lTemp)
                                                            Else
                                                                'Ok, we are in our preferred range... so let's get our preferred facing
                                                                lTemp = GetPreferredSideFacing(oEntity)

                                                                'Basically copied from stand ground
                                                                If lPrimaryFacing <> lTemp Then
                                                                    'Ok, need to change our facing
                                                                    lTemp = (lPrimaryFacing - lTemp) * 900
                                                                    lTemp += .LocAngle

                                                                    If lTemp < 0 Then lTemp += 3600
                                                                    If lTemp > 3599 Then lTemp -= 3600

                                                                    'MSC - remarked out because units would sometimes not turn because lTemp = .LocAngle from these lines... however, once removed, everything worked fine
                                                                    'lTemp -= 1800
                                                                    'If lTemp < 0 Then lTemp += 3600
                                                                    'If lTemp > 3599 Then lTemp -= 3600

                                                                    If lTemp <> .LocAngle Then

                                                                        .SetDest(.LocX, .LocZ, CShort(lTemp))
                                                                        'because of how the engine is, set destangle to truedest
                                                                        .DestAngle = .TrueDestAngle

                                                                        SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)

                                                                        'gfrmDisplayForm.AddEventLine("Entity Changing Angles")

                                                                        'If (.CurrentStatus And elUnitStatus.eMovedByPlayer) <> 0 Then .CurrentStatus -= elUnitStatus.eMovedByPlayer
                                                                        If (.CurrentStatus And elUnitStatus.eMovedByPlayer) <> 0 Then .CurrentStatus = .CurrentStatus Xor elUnitStatus.eMovedByPlayer

                                                                        If .ParentEnvir.lPlayersInEnvirCnt > 0 Then
                                                                            Dim yData(17) As Byte
                                                                            System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectCommand).CopyTo(yData, 0)
                                                                            System.BitConverter.GetBytes(.ObjectID).CopyTo(yData, 2)
                                                                            System.BitConverter.GetBytes(.ObjTypeID).CopyTo(yData, 6)
                                                                            System.BitConverter.GetBytes(.DestX).CopyTo(yData, 8)
                                                                            System.BitConverter.GetBytes(.DestZ).CopyTo(yData, 12)
                                                                            System.BitConverter.GetBytes(CShort(lTemp)).CopyTo(yData, 16)

                                                                            goMsgSys.BroadcastToEnvironmentClients(yData, .ParentEnvir)
                                                                        End If
                                                                    End If

                                                                End If
                                                            End If
                                                        End If

                                                    ElseIf (.iCombatTactics And eiBehaviorPatterns.eEngagement_Stand_Ground) <> 0 OrElse _
                                                      (.iCombatTactics And eiBehaviorPatterns.eEngagement_Pursue) <> 0 Then
                                                        'Stand ground and pursue units do sit-and-spin moves :)
                                                        If ((.CurrentStatus And elUnitStatus.eUnitMoving) = 0) Then
                                                            'lTemp is preferred side facing target
                                                            lTemp = GetPreferredSideFacing(oEntity)

                                                            If lPrimaryFacing <> lTemp Then
                                                                'Ok, need to change our facing
                                                                lTemp = (lPrimaryFacing - lTemp) * 900
                                                                lTemp += .LocAngle

                                                                If lTemp < 0 Then lTemp += 3600
                                                                If lTemp > 3599 Then lTemp -= 3600

                                                                'MSC - remarked out because units would sometimes not turn because lTemp = .LocAngle from these lines... however, once removed, everything worked fine
                                                                'lTemp += 1800
                                                                'If lTemp < 0 Then lTemp += 3600
                                                                'If lTemp > 3599 Then lTemp -= 3600

                                                                If lTemp <> .LocAngle Then

                                                                    .SetDest(.LocX, .LocZ, CShort(lTemp))
                                                                    'because of how the engine is, set destangle to truedest
                                                                    .DestAngle = .TrueDestAngle

                                                                    'gfrmDisplayForm.AddEventLine("Entity Changing Angles")

                                                                    'If (.CurrentStatus And elUnitStatus.eMovedByPlayer) <> 0 Then .CurrentStatus -= elUnitStatus.eMovedByPlayer
                                                                    If (.CurrentStatus And elUnitStatus.eMovedByPlayer) <> 0 Then .CurrentStatus = .CurrentStatus Xor elUnitStatus.eMovedByPlayer

                                                                    If .ParentEnvir.lPlayersInEnvirCnt > 0 Then
                                                                        Dim yData(17) As Byte
                                                                        System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectCommand).CopyTo(yData, 0)
                                                                        System.BitConverter.GetBytes(.ObjectID).CopyTo(yData, 2)
                                                                        System.BitConverter.GetBytes(.ObjTypeID).CopyTo(yData, 6)
                                                                        System.BitConverter.GetBytes(.DestX).CopyTo(yData, 8)
                                                                        System.BitConverter.GetBytes(.DestZ).CopyTo(yData, 12)
                                                                        System.BitConverter.GetBytes(CShort(lTemp)).CopyTo(yData, 16)

                                                                        goMsgSys.BroadcastToEnvironmentClients(yData, .ParentEnvir)
                                                                    End If
                                                                End If

                                                            End If
                                                        End If
                                                    ElseIf (.iCombatTactics And eiBehaviorPatterns.eEngagement_Evade) <> 0 Then
                                                        'The evade pattern is a stored 'Move to' command...
                                                        If (.CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
                                                            If (.LocX <> .lExtendedCT_1 OrElse .LocZ <> .lExtendedCT_2) AndAlso (.CurrentStatus And elUnitStatus.eEngineOperational) <> 0 AndAlso _
                                                             .MaxSpeed > 0 AndAlso .Maneuver > 0 Then
                                                                goMsgSys.SendAIMoveRequestToPathfinding(oEntity, .lExtendedCT_1, .lExtendedCT_2, 0)
                                                            ElseIf .ObjTypeID <> ObjectType.eFacility Then
                                                                'ok, we have evaded to the location, let's set our tactiuc to stand ground
                                                                .iCombatTactics -= eiBehaviorPatterns.eEngagement_Evade
                                                                .iCombatTactics = .iCombatTactics Or eiBehaviorPatterns.eEngagement_Stand_Ground
                                                            End If
                                                        End If
                                                    ElseIf (.iCombatTactics And eiBehaviorPatterns.eEngagement_Dock_With_Target) <> 0 Then
                                                        If (.CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then goMsgSys.SendAIDockRequestToPathfinding(oEntity, .lExtendedCT_1, CShort(.lExtendedCT_2))
                                                    End If

                                                    If (.iCombatTactics And eiBehaviorPatterns.eTactics_Maneuver) <> 0 AndAlso _
                                                       (.iCombatTactics And eiBehaviorPatterns.eEngagement_Stand_Ground) = 0 Then
                                                        'Ok, determine what side I will hit
                                                        If .lPrimaryTargetServerIdx <> -1 Then
                                                            If glEntityIdx(.lPrimaryTargetServerIdx) > 0 Then
                                                                lTemp = WhatSideDoIHit(oEntity, goEntity(.lPrimaryTargetServerIdx))

                                                                'Kinda weird, but here we go...
                                                                'Ok, 0 = 1800, 1 = 900, 2 = 0, 3 = 2700
                                                                Select Case lTemp
                                                                    Case 0 : .iManApproach = 1800
                                                                    Case 1 : .iManApproach = 900
                                                                    Case 2 : .iManApproach = 0
                                                                    Case 3 : .iManApproach = 2700
                                                                End Select

                                                                SetNextManeuverPoint(oEntity, oTmpDef.lModelRangeOffset)
                                                            Else : .lPrimaryTargetServerIdx = -1
                                                            End If
                                                        End If


                                                    End If
                                                End If
                                            End If

                                            If (.iCombatTactics And eiBehaviorPatterns.eTactics_LaunchChildren) <> 0 Then
                                                'Ok, the unit will also launch any children that it has
                                                If bCPOverage = False AndAlso .bInAILaunchAll = False Then
                                                    .bInAILaunchAll = True
                                                    goMsgSys.SendAILaunchAll(oEntity)
                                                    'Now, that we have, we remove that combat tactic
                                                    '.iCombatTactics = .iCombatTactics Xor eiBehaviorPatterns.eTactics_LaunchChildren
                                                End If
                                            End If

                                            'Finally, update all objects that I have them targeted...
                                            If .lPrimaryTargetServerIdx <> -1 Then
                                                If glEntityIdx(.lPrimaryTargetServerIdx) > 0 Then
                                                    goEntity(.lPrimaryTargetServerIdx).AddTargetedBy(.ServerIndex, .ObjectID)
                                                Else : .lPrimaryTargetServerIdx = -1
                                                End If
                                            End If
                                            For Y = 0 To 3
                                                If .lTargetsServerIdx(Y) <> -1 Then
                                                    If glEntityIdx(.lTargetsServerIdx(Y)) > 0 Then
                                                        goEntity(.lTargetsServerIdx(Y)).AddTargetedBy(.ServerIndex, .ObjectID)
                                                    Else : .lTargetsServerIdx(Y) = -1
                                                    End If
                                                End If
                                            Next Y
                                        Else
                                            'Ok, I am not engaged... so, let's check to see if I need to return to a tether point
                                            If .lTetherPointX <> Int32.MinValue AndAlso .lTetherPointZ <> Int32.MinValue Then
                                                If .lTetherPointX <> .LocX OrElse .lTetherPointZ <> .LocZ Then
                                                    If (.iCombatTactics And (eiBehaviorPatterns.eEngagement_Pursue Or eiBehaviorPatterns.eEngagement_Engage)) <> 0 Then
                                                        If (.CurrentStatus And (elUnitStatus.eMovedByPlayer Or elUnitStatus.eUnitMoving)) = 0 Then
                                                            If .LocX <> .lTetherPointX OrElse .LocZ <> .lTetherPointZ Then
                                                                If glCurrentCycle - .lLastTetherCheck > 300 Then
                                                                    .lLastTetherCheck = glCurrentCycle
                                                                    goMsgSys.SendAIMoveRequestToPathfinding(oEntity, .lTetherPointX, .lTetherPointZ, 0)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Else
                                        'Ok, check if finalize stop event is true
                                        If (.CurrentStatus And elUnitStatus.eUnitMoving) = 0 AndAlso .bFinalizeStopEvent = True Then
                                            'Ok, send finalize stop event to Pathfinding server
                                            .bFinalizeStopEvent = False
                                            Dim yFinalStopEvent(7) As Byte
                                            System.BitConverter.GetBytes(GlobalMessageCode.eFinalizeStopEvent).CopyTo(yFinalStopEvent, 0)
                                            .GetGUIDAsString.CopyTo(yFinalStopEvent, 2)
                                            goMsgSys.SendToPathfinding(yFinalStopEvent)
                                        End If
                                    End If
                                ElseIf bSkipReverse = False Then
                                    'Determine what adjust array to use (for map wrapping of planets)
                                    Dim lAdjust() As Int32 = .ParentEnvir.lGridIdxAdjust
                                    If .ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
                                        Dim lModVal As Int32 = .lGridIndex Mod .ParentEnvir.lGridsPerRow
                                        If lModVal = 0 Then
                                            'I'm on the left edge, use the left edge adjust
                                            lAdjust = .ParentEnvir.lLeftEdgeGridIdxAdjust
                                        ElseIf lModVal = .ParentEnvir.lGridsPerRow - 1 Then
                                            'I'm on the right edge, use the right edge adjust
                                            lAdjust = .ParentEnvir.lRightEdgeGridIdxAdjust
                                        End If
                                    End If

                                    'Unit is on Hold Fire, but other units might not be since there is a 
                                    'potential aggression in the environment... have to check reversed
                                    For lGrid = 0 To lAdjust.GetUpperBound(0)
                                        'gonna reuse the llargesector here..
                                        'lLargeSector = .lGridIndex + .ParentEnvir.lGridIdxAdjust(lGrid)
                                        lLargeSector = .lGridIndex + lAdjust(lGrid)
                                        If lLargeSector > -1 AndAlso lLargeSector < .ParentEnvir.lGridUB + 1 Then
                                            'Ok, now, cycle thru the smaller list in the Grid object
                                            Dim oTmpGrid As EnvirGrid = .ParentEnvir.oGrid(lLargeSector)
                                            If oTmpGrid Is Nothing Then Continue For
                                            For Y = 0 To oTmpGrid.lEntityUB
                                                lTemp = oTmpGrid.lEntities(Y)
                                                If lTemp <> -1 Then      'is the unit not nothing?
                                                    oTmpUnit = goEntity(lTemp)
                                                    If oTmpUnit Is Nothing Then
                                                        oTmpGrid.lEntities(Y) = -1
                                                        Continue For
                                                    End If

                                                    If oTmpUnit.ParentEnvir.ObjectID <> .ParentEnvir.ObjectID OrElse oTmpUnit.ParentEnvir.ObjTypeID <> .ParentEnvir.ObjTypeID Then
                                                        oTmpGrid.RemoveEntity(Y)
                                                        oTmpUnit.lGridEntityIdx = oTmpUnit.ParentEnvir.oGrid(oTmpUnit.lGridIndex).AddEntity(oTmpUnit.ServerIndex)
                                                        Continue For
                                                    End If

                                                    If oTmpUnit.ParentEnvir.bEnvirAtColonyLimit = False AndAlso (oTmpUnit.Owner.lIronCurtainPlanetID = Int32.MinValue OrElse (oTmpUnit.ParentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oTmpUnit.Owner.lIronCurtainPlanetID = oTmpUnit.ParentEnvir.ObjectID)) Then
                                                        'that player is in an iron curtain
                                                        Continue For
                                                    End If

                                                    Dim lFlagList As Int32 = elUnitStatus.eUnitEngaged Or elUnitStatus.eSide1HasTarget Or elUnitStatus.eSide2HasTarget Or elUnitStatus.eSide3HasTarget Or elUnitStatus.eSide4HasTarget
                                                    If oTmpUnit.lOwnerID <> .lOwnerID AndAlso (.CurrentStatus And lFlagList) <> lFlagList AndAlso (.CurrentStatus And elUnitStatus.eUnitCloaked) = 0 Then

                                                        lEnvirDist = Int32.MaxValue
                                                        lTemp = giRelativeSmall(lGrid, .lSmallSectorID, oTmpUnit.lSmallSectorID)
                                                        If lTemp <> -1 Then
                                                            lRelTinyX = glBaseRelTinyX(lTemp) + oTmpUnit.lTinyX + (gl_HALF_FINAL_PER_ROW - .lTinyX)
                                                            If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                                                lRelTinyZ = glBaseRelTinyZ(lTemp) + oTmpUnit.lTinyZ + (gl_HALF_FINAL_PER_ROW - .lTinyZ)
                                                                If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                                                    lEnvirDist = gyDistances(lRelTinyX, lRelTinyZ) - oTmpUnit.lModelRangeOffset

                                                                    lFacing = WhatSideCanFire(oTmpUnit, lRelTinyX, lRelTinyZ)
                                                                    If .lTargetsServerIdx(lFacing) = -1 AndAlso ((oTmpUnit.CurrentStatus And elUnitStatus.eRadarOperational) <> 0) Then
                                                                        If oTmpUnit.lEntityDefServerIndex <> -1 Then
                                                                            If glEntityDefIdx(oTmpUnit.lEntityDefServerIndex) <> -1 Then
                                                                                oTmpDef2 = goEntityDefs(oTmpUnit.lEntityDefServerIndex)
                                                                            End If
                                                                        End If
                                                                        If lEnvirDist < oTmpDef2.lOptRadarRange + oTmpUnit.lJamRangeMod Then
                                                                            'Ok, the moving unit is visible to me... now, test for aggression
                                                                            yRelID = oTmpUnit.Owner.GetPlayerRelScore(.lOwnerID, True, .ServerIndex)
                                                                            If yRelID <= elRelTypes.eWar AndAlso yRelID > 0 Then
                                                                                'Use the bForceAggressionTest so that the unit enacts its own scripts
                                                                                oTmpUnit.bForceAggressionTest = True
                                                                                SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oTmpUnit.ServerIndex, oTmpUnit.ObjectID) 'AddEntityMoving(oTmpUnit.ServerIndex, oTmpUnit.ObjectID)
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            Next Y
                                        End If
                                    Next lGrid

                                End If
                            ElseIf .ParentEnvir.PotentialFirstContact = True Then 'AndAlso .ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
                                If (.CurrentStatus And elUnitStatus.eRadarOperational) <> 0 Then
                                    'Essentially the same as an aggression test but with different results...
                                    'Determine what adjust array to use (for map wrapping of planets)
                                    Dim lAdjust() As Int32 = .ParentEnvir.lGridIdxAdjust
                                    If .ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
                                        Dim lModVal As Int32 = .lGridIndex Mod .ParentEnvir.lGridsPerRow
                                        If lModVal = 0 Then
                                            'I'm on the left edge, use the left edge adjust
                                            lAdjust = .ParentEnvir.lLeftEdgeGridIdxAdjust
                                        ElseIf lModVal = .ParentEnvir.lGridsPerRow - 1 Then
                                            'I'm on the right edge, use the right edge adjust
                                            lAdjust = .ParentEnvir.lRightEdgeGridIdxAdjust
                                        End If
                                    End If

                                    Dim bBreakOut As Boolean = False
                                    For lGrid = 0 To lAdjust.GetUpperBound(0)
                                        'gonna reuse the llargesector here... 
                                        lLargeSector = .lGridIndex + lAdjust(lGrid)

                                        If lLargeSector > -1 AndAlso lLargeSector < .ParentEnvir.lGridUB + 1 Then
                                            'Ok, now, cycle thru the smaller list in the Grid object
                                            Dim oTmpGrid As EnvirGrid = .ParentEnvir.oGrid(lLargeSector)
                                            If oTmpGrid Is Nothing Then Continue For
                                            For Y = 0 To oTmpGrid.lEntityUB
                                                lTemp = oTmpGrid.lEntities(Y)
                                                If lTemp <> -1 Then      'is the unit not nothing?
                                                    oTmpUnit = goEntity(lTemp)

                                                    If oTmpUnit.ParentEnvir.ObjectID <> .ParentEnvir.ObjectID OrElse oTmpUnit.ParentEnvir.ObjTypeID <> .ParentEnvir.ObjTypeID Then
                                                        oTmpGrid.RemoveEntity(Y)
                                                        oTmpUnit.lGridEntityIdx = oTmpUnit.ParentEnvir.oGrid(oTmpUnit.lGridIndex).AddEntity(oTmpUnit.ServerIndex)
                                                        Continue For
                                                    End If

                                                    'Ok, first, check if the unit belongs to me
                                                    If oTmpUnit.lOwnerID <> .lOwnerID Then
                                                        'Ok, now, check if the player already has relations
                                                        If .Owner.HasPlayerRelationship(oTmpUnit.lOwnerID) = False Then
                                                            'Ok, let's see if they are in range...
                                                            lEnvirDist = Int32.MaxValue
                                                            lTemp = giRelativeSmall(lGrid, .lSmallSectorID, oTmpUnit.lSmallSectorID)
                                                            If lTemp <> -1 Then
                                                                lRelTinyX = glBaseRelTinyX(lTemp) + oTmpUnit.lTinyX + (gl_HALF_FINAL_PER_ROW - .lTinyX)
                                                                If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                                                    lRelTinyZ = glBaseRelTinyZ(lTemp) + oTmpUnit.lTinyZ + (gl_HALF_FINAL_PER_ROW - .lTinyZ)
                                                                    If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                                                        lEnvirDist = gyDistances(lRelTinyX, lRelTinyZ) - oTmpUnit.lModelRangeOffset
                                                                    End If
                                                                End If
                                                            End If


                                                            If lEnvirDist <= oTmpDef.lOptRadarRange + .lJamRangeMod AndAlso ((oTmpUnit.CurrentStatus And elUnitStatus.eUnitCloaked) = 0) Then
                                                                'Yes, they are in radar range... so we'll do this easy like
                                                                .Owner.GetPlayerRelScore(oTmpUnit.lOwnerID, True, oTmpUnit.ServerIndex)
                                                                bBreakOut = True
                                                                Exit For
                                                            End If
                                                        End If
                                                    End If

                                                End If      'if unit is nothing
                                            Next Y
                                            If bBreakOut = True Then Exit For
                                        End If
                                    Next lGrid
                                End If
                            End If  'Potential Aggression in environment

                            .bForceAggressionTest = False   'ensure that our force aggression test is false
                            .bNewAddedEntity = False
                        ElseIf bTurned = True Then      'Didnt change sector, but did i turn???
                            'Yes, I turned, didn't change sectors, but turned...
                            If ((.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0) Then .VerifyTargets()
                        End If  'SectorID changed

					End With	'with current entity 
				End If	'if glUnitIdx(X) <> -1
			Next X		'all globals...
		Catch ex As Exception
			gfrmDisplayForm.AddEventLine("HandleMovement: " & ex.Message)
		End Try
	End Sub

    Private mcolInCombat As New Collection
    Private Sub DoCombatCollectionAction(ByVal yType As Byte, ByVal lID As Int32, ByVal iTypeID As Int16, ByVal lIdx As Int32)
        SyncLock mcolInCombat
            Try
                Dim sKey As String = lID & ";" & iTypeID
                If yType = 0 Then   'add
                    If mcolInCombat.Contains(sKey) = True Then Return
                    mcolInCombat.Add(lIdx, sKey)
                Else                'remove
                    If mcolInCombat.Contains(sKey) = True Then mcolInCombat.Remove(sKey)
                End If
            Catch
            End Try
        End SyncLock
    End Sub

    Public Sub AddEntityInCombat(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal lIdx As Int32)
        'SyncLock mcolInCombat
        '    Try
        '        Dim sKey As String = lID & ";" & iTypeID
        '        If mcolInCombat.Contains(sKey) = True Then Return
        '        mcolInCombat.Add(lIdx, sKey)
        '    Catch
        '    End Try
        'End SyncLock
        DoCombatCollectionAction(0, lID, iTypeID, lIdx)
    End Sub
    Public Sub RemoveEntityInCombat(ByVal lID As Int32, ByVal iTypeID As Int16)
        'SyncLock mcolInCombat
        '    Try
        '        Dim sKey As String = lID & ";" & iTypeID
        '        If mcolInCombat.Contains(sKey) = True Then mcolInCombat.Remove(sKey)
        '    Catch
        '    End Try
        'End SyncLock
        DoCombatCollectionAction(1, lID, iTypeID, -1)
    End Sub

    Private mbDoTheNewCode As Boolean = False
	Private mlPreviousJamCycle As Int32
	Private mlJamState As Int32 = 0		'0 = no jamming, 1 = clear jam modifiers, 2 = disres jamming, 3 = jammer jamming, 4 = remaining jamming
	Public Sub HandleCombat()
        'Dim X As Int32
		Dim Y As Int32

		Dim oTarget As Epica_Entity
		Dim oWeapon As WeaponDef
		Dim lDmg As Int32

		Dim lTmp As Int32

		Dim yTemp As Byte

		Dim ySideHit As Byte

		Dim lBaseToHit As Int32
		Dim lRoll As Int32

        'Dim oTmpUnit As Epica_Entity		'for resetting targets

		Dim oCurDef As Epica_Entity_Def = Nothing
		Dim oTargetDef As Epica_Entity_Def = Nothing

		Dim bSendHPUpdateMsg As Boolean			'indicates if we are to send a UnitHP update message

		Dim bUnitDestroyed As Boolean
		Dim lKilledByID As Int32 = -1
		Dim oEvent As WeaponEvent

		Dim lRange As Int32

		If mlJamState = 0 Then
			If glCurrentCycle - mlPreviousJamCycle > 150 Then
				mlJamState = 1
				mlPreviousJamCycle = glCurrentCycle
			End If
		Else
			mlJamState += 1
			If mlJamState > 4 Then mlJamState = 0
		End If



		'Go thru each entity to get shots fired at the entity
        'Dim lCurUB As Int32 = Math.Min(glEntityUB, glEntityIdx.GetUpperBound(0))
        'lCurUB = Math.Min(lCurUB, goEntity.GetUpperBound(0))
        '		For X = 0 To lCurUB
        For Each X As Int32 In mcolInCombat
            If X > -1 AndAlso X <= glEntityUB AndAlso glEntityIdx(X) > 0 Then
                Dim oEntity As Epica_Entity = goEntity(X)
                If oEntity Is Nothing Then
                    glEntityIdx(X) = -1
                    Continue For
                End If
                lKilledByID = -1
                With oEntity
                    'get our definition
                    If .lEntityDefServerIndex <> -1 Then
                        If glEntityDefIdx(.lEntityDefServerIndex) <> -1 Then
                            oCurDef = goEntityDefs(.lEntityDefServerIndex)
                        End If
                    End If

                    'reset our update checker...
                    bSendHPUpdateMsg = False
                    bUnitDestroyed = False

                    If oCurDef Is Nothing Then Continue For

                    If .lTargetsServerIdx Is Nothing = False Then
                        For Y = 0 To .lTargetsServerIdx.GetUpperBound(0)
                            Dim lTempTargetServerIdx As Int32 = .lTargetsServerIdx(Y)
                            If lTempTargetServerIdx <> -1 Then
                                If glEntityIdx(lTempTargetServerIdx) < 1 Then
                                    .lTargetsServerIdx(Y) = -1
                                    .bForceAggressionTest = True
                                    SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)
                                Else
                                    Dim oChkUnit As Epica_Entity = goEntity(lTempTargetServerIdx)
                                    If oChkUnit Is Nothing = True Then
                                        .lTargetsServerIdx(Y) = -1
                                        .bForceAggressionTest = True
                                        SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)
                                    Else
                                        'ok, check the target's owner for iron curtain
                                        If .ParentEnvir.bEnvirAtColonyLimit = False AndAlso (oChkUnit.Owner.lIronCurtainPlanetID = Int32.MinValue OrElse (oChkUnit.Owner.lIronCurtainPlanetID = .ParentEnvir.ObjectID AndAlso .ParentEnvir.ObjTypeID = ObjectType.ePlanet)) Then
                                            .lTargetsServerIdx(Y) = -1
                                            .bForceAggressionTest = True
                                            SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)
                                        End If
                                    End If
                                End If
                            End If
                        Next Y
                    End If

                    If .lPrimaryTargetServerIdx <> -1 Then
                        Dim lTempTargetServerIdx As Int32 = .lPrimaryTargetServerIdx
                        If lTempTargetServerIdx <> -1 Then
                            If glEntityIdx(lTempTargetServerIdx) < 1 Then
                                .lPrimaryTargetServerIdx = -1
                                .bForceAggressionTest = True
                                SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)
                            Else
                                Dim oChkUnit As Epica_Entity = goEntity(lTempTargetServerIdx)
                                If oChkUnit Is Nothing = True Then
                                    .lPrimaryTargetServerIdx = -1
                                    .bForceAggressionTest = True
                                    SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)
                                Else
                                    'ok, check the target's owner for iron curtain
                                    Try
                                        If .ParentEnvir.bEnvirAtColonyLimit = False AndAlso (oChkUnit.Owner Is Nothing OrElse oChkUnit.Owner.lIronCurtainPlanetID = Int32.MinValue OrElse (.ParentEnvir Is Nothing = False AndAlso oChkUnit.Owner.lIronCurtainPlanetID = .ParentEnvir.ObjectID AndAlso .ParentEnvir.ObjTypeID = ObjectType.ePlanet)) Then
                                            .lPrimaryTargetServerIdx = -1
                                            .bForceAggressionTest = True
                                            SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)
                                        End If
                                    Catch
                                    End Try
                                End If
                            End If
                        End If
                    End If

                    'check for Shield HP Regen....
                    'If (.CurrentStatus And elUnitStatus.eShieldOperational) <> 0 Then
                    '	If .ShieldHP < oCurDef.Shield_MaxHP AndAlso .lNextShieldRechargeCycle < glCurrentCycle Then
                    '		.ShieldHP += oCurDef.ShieldRecharge
                    '		If .ShieldHP > oCurDef.Shield_MaxHP Then .ShieldHP = oCurDef.Shield_MaxHP
                    '		bSendHPUpdateMsg = True
                    '		.lNextShieldRechargeCycle = glCurrentCycle + oCurDef.ShieldRechargeFreq
                    '	End If
                    'Else
                    '	.ShieldHP = 0
                    'End If

                    If mlJamState = 1 Then .ClearJamModifiers()

                    'now, check for weapon events
                    If .lNextWpnEvent < glCurrentCycle Then
                        .lLastEngagement = glCurrentCycle
                        .lNextWpnEvent = Int32.MaxValue
                        For Y = 0 To .lWpnEventUB
                            If .yWpnEventUsed(Y) <> 0 Then
                                If .oWpnEvents(Y).lEventCycle < glCurrentCycle Then
                                    'Ok, need to process this event
                                    oEvent = .oWpnEvents(Y)
                                    .yWpnEventUsed(Y) = 0

                                    'If oEvent.fDegradation <> 0 Then oEvent.HandlePulseDegradation(.LocX, .LocZ)

                                    bSendHPUpdateMsg = True

                                    If oEvent.lFromIdx <> -1 Then
                                        If glEntityIdx(oEvent.lFromIdx) > 0 Then
                                            Dim oTmpFromEntity As Epica_Entity = goEntity(oEvent.lFromIdx)
                                            If oTmpFromEntity Is Nothing = False Then
                                                ySideHit = WhatSideDoIHit(oTmpFromEntity, oEntity)
                                            Else
                                                ySideHit = oEvent.ySideHit
                                            End If
                                        Else
                                            ySideHit = oEvent.ySideHit
                                        End If
                                    Else
                                        ySideHit = oEvent.ySideHit
                                    End If

                                    If oEvent.yCritical <> eyCriticalHitType.NoCriticalHit AndAlso oEvent.yCritical <> eyCriticalHitType.NormalCriticalHit Then
                                        'ok, fighter sub system critical hit
                                        bSendHPUpdateMsg = .FighterSubsystemHit(ySideHit, oEvent.yCritical)
                                        If (oCurDef.yChassisType And (ChassisType.eNavalBased Or ChassisType.eGroundBased)) = 0 AndAlso (.CurrentStatus And elUnitStatus.eEngineOperational) = 0 AndAlso .ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
                                            .StructuralHP = 0       'set to 0, the unit is dead
                                        End If
                                        oEvent.yCritical = eyCriticalHitType.NormalCriticalHit
                                    End If

                                    If mbRunAntiLightCode = True Then
                                        If .lOwnerID = 20944 OrElse .lOwnerID = 29216 OrElse .lOwnerID = 20900 OrElse .lOwnerID = 20842 OrElse .lOwnerID = 28522 OrElse .lOwnerID = 24396 Then
                                            oEvent.lPierceDmg *= 100
                                            oEvent.lImpactDmg *= 100
                                            oEvent.lBeamDmg *= 100
                                            oEvent.lECMDmg *= 100
                                            oEvent.lFlameDmg *= 100
                                            oEvent.lChemicalDmg *= 100
                                        End If
                                    End If

                                    If .ShieldHP > 0 Then
                                        'Ok, we're hitting shields, no resistances... all damage is damage
                                        lDmg = oEvent.lPierceDmg + oEvent.lImpactDmg + oEvent.lBeamDmg + oEvent.lECMDmg + oEvent.lFlameDmg + oEvent.lChemicalDmg
                                        'Dim lCritBonus As Int32 = 1

                                        'when critically hitting shields, damage is quadrupled
                                        If oEvent.yCritical <> eyCriticalHitType.NoCriticalHit Then lDmg *= 4

                                        'Enveloping hits on shields do double damage...
                                        If ySideHit = 255 Then lDmg *= 2

                                        'Ok, reduce our shields
                                        Dim lNewShieldHP As Int32 = .ShieldHP - lDmg
                                        .ShieldHP = lNewShieldHP
                                        If lNewShieldHP < 0 Then
                                            'lDmg = lDmg + .ShieldHP
                                            'lDmg = Math.Abs(lNewShieldHP)
                                            'lDmg = lDmg / 2        'only 50% will break thru the shields
                                            'however, this is unaffected by armor resistances...

                                            '2010-06-23 EpicFail: It now does an armor test.  We ditch the crit and envelop as they were against the shield alone.
                                            'What percent of the damage got thru
                                            Dim fMult As Single = CSng(Math.Abs(lNewShieldHP) / lDmg)
                                            lDmg = 0
                                            Dim lIntegrityTest As Int32 = CInt(Rnd() * 100) 'gaurenteed 1% chance to pass ignore armor resists.
                                            If lIntegrityTest > oCurDef.ArmorIntegrity Then
                                                'here, we use the Fast Resist values which are already precalculated
                                                If .Owner.yDoctrineOfLeadership = eDoctrineOfLeadershipSetting.eDefense Then
                                                    Dim fDoctrineAdd As Single = 0.0F
                                                    fDoctrineAdd = -0.02F
                                                    With oCurDef
                                                        lDmg += CInt((.FastBeamResist + fDoctrineAdd) * oEvent.lBeamDmg * fMult)
                                                        lDmg += CInt((.FastChemicalResist + fDoctrineAdd) * oEvent.lChemicalDmg * fMult)
                                                        lDmg += CInt((.FastECMResist + fDoctrineAdd) * oEvent.lECMDmg * fMult)
                                                        lDmg += CInt((.FastFlameResist + fDoctrineAdd) * oEvent.lFlameDmg * fMult)
                                                        lDmg += CInt((.FastImpactResist + fDoctrineAdd) * oEvent.lImpactDmg * fMult)
                                                        lDmg += CInt((.FastPiercingResist + fDoctrineAdd) * oEvent.lPierceDmg * fMult)
                                                    End With
                                                Else
                                                    With oCurDef
                                                        lDmg += CInt(.FastBeamResist * oEvent.lBeamDmg * fMult)
                                                        lDmg += CInt(.FastChemicalResist * oEvent.lChemicalDmg * fMult)
                                                        lDmg += CInt(.FastECMResist * oEvent.lECMDmg * fMult)
                                                        lDmg += CInt(.FastFlameResist * oEvent.lFlameDmg * fMult)
                                                        lDmg += CInt(.FastImpactResist * oEvent.lImpactDmg * fMult)
                                                        lDmg += CInt(.FastPiercingResist * oEvent.lPierceDmg * fMult)
                                                    End With
                                                End If
                                            Else
                                                lDmg += CInt(oEvent.lBeamDmg * fMult)
                                                lDmg += CInt(oEvent.lChemicalDmg * fMult)
                                                lDmg += CInt(oEvent.lECMDmg * fMult)
                                                lDmg += CInt(oEvent.lFlameDmg * fMult)
                                                lDmg += CInt(oEvent.lImpactDmg * fMult)
                                                lDmg += CInt(oEvent.lPierceDmg * fMult)
                                            End If

                                            'If .ObjTypeID = ObjectType.eUnit Then .AddPlayerDmg(oEvent.FromPlayerID, lDmg, oEvent.yAttackerType, oCurDef.yHullType) Else .FacAddPlayerDmg(oEvent.FromPlayerID, lDmg)

                                            If ySideHit = 255 Then
                                                For yTmpSideHit As Byte = 0 To 3
                                                    If .ArmorHP(yTmpSideHit) > 0 Then
                                                        .ArmorHP(yTmpSideHit) -= lDmg
                                                        If .ArmorHP(yTmpSideHit) < 0 Then
                                                            Dim lTmpDmg As Int32 = Math.Abs(.ArmorHP(yTmpSideHit))
                                                            .ArmorHP(yTmpSideHit) = 0
                                                            .StructuralHP -= lTmpDmg
                                                            If .StructuralHP < 0 Then .StructuralHP = 0
                                                        End If
                                                    Else
                                                        .StructuralHP -= lDmg
                                                        If .StructuralHP < 0 Then .StructuralHP = 0
                                                    End If
                                                Next yTmpSideHit
                                            Else
                                                'Code that is there already
                                                If .ArmorHP(ySideHit) > 0 Then
                                                    .ArmorHP(ySideHit) -= lDmg
                                                    If .ArmorHP(ySideHit) < 0 Then
                                                        Dim lTmpDmg As Int32 = Math.Abs(.ArmorHP(ySideHit))
                                                        .ArmorHP(ySideHit) = 0
                                                        .StructuralHP -= lTmpDmg
                                                        If .StructuralHP < 0 Then .StructuralHP = 0
                                                    End If
                                                Else
                                                    'ok, the hit was to structural... a hit in this manner
                                                    '  will NEVER produce a critical hit
                                                    .StructuralHP -= lDmg
                                                    If .StructuralHP < 0 Then .StructuralHP = 0
                                                End If
                                            End If
                                            .ShieldHP = 0
                                        End If
                                    Else
                                        If ySideHit = 255 Then
                                            For ySideHit = 0 To 3
                                                'ok, hitting armor/structure: the shields are down, test armor integrity
                                                If .ArmorHP(ySideHit) > 0 AndAlso oEvent.yCritical = eyCriticalHitType.NoCriticalHit Then
                                                    'Ok, now apply our damage to the resistance values
                                                    lDmg = 0
                                                    Dim lIntegrityTest As Int32 = CInt(Rnd() * 100)
                                                    If lIntegrityTest > oCurDef.ArmorIntegrity Then
                                                        'here, we use the Fast Resist values which are already precalculated
                                                        If .Owner.yDoctrineOfLeadership = eDoctrineOfLeadershipSetting.eDefense Then
                                                            Dim fDoctrineAdd As Single = 0.0F
                                                            fDoctrineAdd = -0.02F
                                                            With oCurDef
                                                                lDmg += CInt((.FastBeamResist + fDoctrineAdd) * oEvent.lBeamDmg)
                                                                lDmg += CInt((.FastChemicalResist + fDoctrineAdd) * oEvent.lChemicalDmg)
                                                                lDmg += CInt((.FastECMResist + fDoctrineAdd) * oEvent.lECMDmg)
                                                                lDmg += CInt((.FastFlameResist + fDoctrineAdd) * oEvent.lFlameDmg)
                                                                lDmg += CInt((.FastImpactResist + fDoctrineAdd) * oEvent.lImpactDmg)
                                                                lDmg += CInt((.FastPiercingResist + fDoctrineAdd) * oEvent.lPierceDmg)
                                                            End With
                                                        Else
                                                            With oCurDef
                                                                lDmg += CInt(.FastBeamResist * oEvent.lBeamDmg)
                                                                lDmg += CInt(.FastChemicalResist * oEvent.lChemicalDmg)
                                                                lDmg += CInt(.FastECMResist * oEvent.lECMDmg)
                                                                lDmg += CInt(.FastFlameResist * oEvent.lFlameDmg)
                                                                lDmg += CInt(.FastImpactResist * oEvent.lImpactDmg)
                                                                lDmg += CInt(.FastPiercingResist * oEvent.lPierceDmg)
                                                            End With
                                                        End If
                                                    Else
                                                        'With oCurDef
                                                        lDmg += oEvent.lBeamDmg
                                                        lDmg += oEvent.lChemicalDmg
                                                        lDmg += oEvent.lECMDmg
                                                        lDmg += oEvent.lFlameDmg
                                                        lDmg += oEvent.lImpactDmg
                                                        lDmg += oEvent.lPierceDmg
                                                        'End With
                                                    End If

                                                    'If .ObjTypeID = ObjectType.eUnit Then .AddPlayerDmg(oEvent.FromPlayerID, lDmg, oEvent.yAttackerType, oCurDef.yHullType) Else .FacAddPlayerDmg(oEvent.FromPlayerID, lDmg)

                                                    .ArmorHP(ySideHit) -= lDmg
                                                    If .ArmorHP(ySideHit) < 0 Then
                                                        Dim lTmpDmg As Int32 = Math.Abs(.ArmorHP(ySideHit))
                                                        .ArmorHP(ySideHit) = 0
                                                        .StructuralHP -= lTmpDmg
                                                        If .StructuralHP < 0 Then .StructuralHP = 0
                                                    End If
                                                Else
                                                    'Ok, unit has no more armor on that side or the integrity test failed, start eating away at 
                                                    ' the structure which always takes the brunt of the force
                                                    lDmg = oEvent.lPierceDmg + oEvent.lImpactDmg + oEvent.lBeamDmg + oEvent.lECMDmg + oEvent.lFlameDmg + oEvent.lChemicalDmg

                                                    'if the structure is hit critically
                                                    If oEvent.yCritical <> eyCriticalHitType.NoCriticalHit Then
                                                        'and armor still exists, then damage is quadrupled
                                                        'If .ArmorHP(ySideHit) > 0 Then
                                                        '	lDmg *= 4
                                                        'Else
                                                        '	'otherwise, it is times 10
                                                        '	lDmg *= 10
                                                        'End If
                                                        If .ArmorHP(ySideHit) < 1 Then
                                                            lDmg *= 4
                                                        End If
                                                    End If

                                                    If .ObjTypeID = ObjectType.eUnit AndAlso .ArmorHP(ySideHit) = 0 Then
                                                        If oCurDef.HullSize > 670 OrElse (oCurDef.yChassisType And ChassisType.eGroundBased) = 0 Then
                                                            If (lDmg > .StructuralHP \ 10) OrElse oEvent.yCritical <> eyCriticalHitType.NoCriticalHit Then
                                                                'Ok, damage was enough to test for critical hit...
                                                                bUnitDestroyed = .CriticalHit(ySideHit)
                                                            End If
                                                        End If
                                                    End If

                                                    .StructuralHP -= lDmg

                                                    'If .ObjTypeID = ObjectType.eUnit Then .AddPlayerDmg(oEvent.FromPlayerID, lDmg, oEvent.yAttackerType, oCurDef.yHullType) Else .FacAddPlayerDmg(oEvent.FromPlayerID, lDmg)

                                                    If .StructuralHP < 1 OrElse bUnitDestroyed Then
                                                        bUnitDestroyed = True
                                                        .StructuralHP = 0

                                                        'TODO: Add an AOE effect here based on the hull size of the entity dying

                                                        ''Update the entities that had this entity targeted
                                                        'For lTmp = 0 To .lTargetedByUB
                                                        '	If .lTargetedByIdx(lTmp) <> -1 Then
                                                        '		If glEntityIdx(.lTargetedByIdx(lTmp)) <> -1 Then
                                                        '			oTmpUnit = goEntity(.lTargetedByIdx(lTmp))
                                                        '			If oTmpUnit Is Nothing = False Then
                                                        '				oTmpUnit.ApplyExpGain(oCurDef.CombatRating, oEvent.lFromIdx = .lTargetedByIdx(lTmp))

                                                        '				If oTmpUnit.lPrimaryTargetServerIdx = .ServerIndex Then
                                                        '					If (oTmpUnit.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then oTmpUnit.CurrentStatus -= elUnitStatus.eUnitEngaged
                                                        '					oTmpUnit.lPrimaryTargetServerIdx = -1
                                                        '					oTmpUnit.bForceAggressionTest = True
                                                        '				End If
                                                        '				If oTmpUnit.lTargetsServerIdx(0) = .ServerIndex Then
                                                        '					oTmpUnit.lTargetsServerIdx(0) = -1
                                                        '					If (oTmpUnit.CurrentStatus And elUnitStatus.eSide1HasTarget) <> 0 Then
                                                        '						oTmpUnit.CurrentStatus -= elUnitStatus.eSide1HasTarget
                                                        '					End If
                                                        '					oTmpUnit.bForceAggressionTest = True
                                                        '				End If
                                                        '				If oTmpUnit.lTargetsServerIdx(1) = .ServerIndex Then
                                                        '					oTmpUnit.lTargetsServerIdx(1) = -1
                                                        '					If (oTmpUnit.CurrentStatus And elUnitStatus.eSide2HasTarget) <> 0 Then
                                                        '						oTmpUnit.CurrentStatus -= elUnitStatus.eSide2HasTarget
                                                        '					End If
                                                        '					oTmpUnit.bForceAggressionTest = True
                                                        '				End If
                                                        '				If oTmpUnit.lTargetsServerIdx(2) = .ServerIndex Then
                                                        '					oTmpUnit.lTargetsServerIdx(2) = -1
                                                        '					If (oTmpUnit.CurrentStatus And elUnitStatus.eSide3HasTarget) <> 0 Then
                                                        '						oTmpUnit.CurrentStatus -= elUnitStatus.eSide3HasTarget
                                                        '					End If
                                                        '					oTmpUnit.bForceAggressionTest = True
                                                        '				End If
                                                        '				If oTmpUnit.lTargetsServerIdx(3) = .ServerIndex Then
                                                        '					oTmpUnit.lTargetsServerIdx(3) = -1
                                                        '					If (oTmpUnit.CurrentStatus And elUnitStatus.eSide4HasTarget) <> 0 Then
                                                        '						oTmpUnit.CurrentStatus -= elUnitStatus.eSide4HasTarget
                                                        '					End If
                                                        '					oTmpUnit.bForceAggressionTest = True
                                                        '				End If
                                                        '			End If
                                                        '		End If
                                                        '	End If
                                                        'Next lTmp

                                                        Exit For

                                                    End If
                                                End If
                                            Next ySideHit
                                        Else
                                            'ok, hitting armor/structure: the shields are down, test armor integrity
                                            If .ArmorHP(ySideHit) > 0 AndAlso oEvent.yCritical = eyCriticalHitType.NoCriticalHit Then
                                                'Ok, now apply our damage to the resistance values
                                                lDmg = 0

                                                Dim lIntegrityTest As Int32 = CInt(Rnd() * 100)
                                                If lIntegrityTest > oCurDef.ArmorIntegrity Then
                                                    'here, we use the Fast Resist values which are already precalculated
                                                    If .Owner.yDoctrineOfLeadership = eDoctrineOfLeadershipSetting.eDefense Then
                                                        Dim fDoctrineAdd As Single = 0.0F
                                                        fDoctrineAdd = -0.02F
                                                        With oCurDef
                                                            lDmg += CInt((.FastBeamResist + fDoctrineAdd) * oEvent.lBeamDmg)
                                                            lDmg += CInt((.FastChemicalResist + fDoctrineAdd) * oEvent.lChemicalDmg)
                                                            lDmg += CInt((.FastECMResist + fDoctrineAdd) * oEvent.lECMDmg)
                                                            lDmg += CInt((.FastFlameResist + fDoctrineAdd) * oEvent.lFlameDmg)
                                                            lDmg += CInt((.FastImpactResist + fDoctrineAdd) * oEvent.lImpactDmg)
                                                            lDmg += CInt((.FastPiercingResist + fDoctrineAdd) * oEvent.lPierceDmg)
                                                        End With
                                                    Else
                                                        With oCurDef
                                                            lDmg += CInt(.FastBeamResist * oEvent.lBeamDmg)
                                                            lDmg += CInt(.FastChemicalResist * oEvent.lChemicalDmg)
                                                            lDmg += CInt(.FastECMResist * oEvent.lECMDmg)
                                                            lDmg += CInt(.FastFlameResist * oEvent.lFlameDmg)
                                                            lDmg += CInt(.FastImpactResist * oEvent.lImpactDmg)
                                                            lDmg += CInt(.FastPiercingResist * oEvent.lPierceDmg)
                                                        End With
                                                    End If
                                                Else
                                                    'here, we use the Fast Resist values which are already precalculated
                                                    'With oCurDef
                                                    lDmg += oEvent.lBeamDmg
                                                    lDmg += oEvent.lChemicalDmg
                                                    lDmg += oEvent.lECMDmg
                                                    lDmg += oEvent.lFlameDmg
                                                    lDmg += oEvent.lImpactDmg
                                                    lDmg += oEvent.lPierceDmg
                                                    'End With
                                                End If

                                                'Damage to an armor plate NEVER passes through to the structure
                                                .ArmorHP(ySideHit) -= lDmg

                                                'If .ObjTypeID = ObjectType.eUnit Then .AddPlayerDmg(oEvent.FromPlayerID, lDmg, oEvent.yAttackerType, oCurDef.yHullType) Else .FacAddPlayerDmg(oEvent.FromPlayerID, lDmg)

                                                If .ArmorHP(ySideHit) < 0 Then
                                                    Dim lTmpDmg As Int32 = Math.Abs(.ArmorHP(ySideHit))
                                                    .ArmorHP(ySideHit) = 0

                                                    .StructuralHP -= lTmpDmg
                                                    If .StructuralHP < 1 Then .StructuralHP = 0
                                                End If
                                            Else
                                                'Ok, unit has no more armor on that side, start eating away at 
                                                ' the structure which always takes the brunt of the force
                                                lDmg = oEvent.lPierceDmg + oEvent.lImpactDmg + oEvent.lBeamDmg + oEvent.lECMDmg + oEvent.lFlameDmg + oEvent.lChemicalDmg

                                                'if the structure is hit critically
                                                If oEvent.yCritical <> eyCriticalHitType.NoCriticalHit Then
                                                    'and armor still exists, then damage is quadrupled
                                                    'If .ArmorHP(ySideHit) > 0 Then
                                                    '	lDmg *= 4
                                                    'Else
                                                    '	'otherwise, it is times 10
                                                    '	lDmg *= 10
                                                    'End If
                                                    If .ArmorHP(ySideHit) < 1 Then
                                                        lDmg *= 4
                                                    End If
                                                End If

                                                'If .ObjTypeID = ObjectType.eUnit Then .AddPlayerDmg(oEvent.FromPlayerID, lDmg, oEvent.yAttackerType, oCurDef.yHullType) Else .FacAddPlayerDmg(oEvent.FromPlayerID, lDmg)

                                                If .ArmorHP(ySideHit) = 0 AndAlso .ObjTypeID = ObjectType.eUnit Then
                                                    If oCurDef.HullSize > 670 OrElse (oCurDef.yChassisType And ChassisType.eGroundBased) = 0 Then
                                                        If (lDmg > .StructuralHP \ 10) OrElse oEvent.yCritical <> eyCriticalHitType.NoCriticalHit Then
                                                            'Ok, damage was enough to test for critical hit...
                                                            bUnitDestroyed = .CriticalHit(ySideHit)
                                                        End If
                                                    End If
                                                End If

                                                .StructuralHP -= lDmg
                                                If .StructuralHP < 0 OrElse bUnitDestroyed = True Then .StructuralHP = 0
                                            End If

                                        End If
                                    End If

                                    If .StructuralHP < 1 OrElse bUnitDestroyed = True Then
                                        bUnitDestroyed = True
                                        .StructuralHP = 0

                                        If oEvent.lFromIdx > -1 AndAlso glEntityIdx(oEvent.lFromIdx) > 0 Then
                                            Dim oTmpFromEntity As Epica_Entity = goEntity(oEvent.lFromIdx)
                                            If oTmpFromEntity Is Nothing = False Then
                                                lKilledByID = oTmpFromEntity.lOwnerID
                                            End If
                                        End If

                                        'TODO: Add an AOE effect here based on the hull size of the entity dying

                                        .ClearTargetLists(True, lKilledByID)

                                    End If

                                    'Now, check for an AOE effect
                                    If oEvent.AOERange > 0 Then ApplyAOEDamage(oEvent, oEntity)

                                    'Now that we are done, remove the event
                                    oEvent = Nothing
                                    .oWpnEvents(Y) = Nothing
                                    .yWpnEventUsed(Y) = 0
                                Else
                                    .lNextWpnEvent = Math.Min(.lNextWpnEvent, .oWpnEvents(Y).lEventCycle)
                                End If
                            End If
                        Next Y
                    End If

                    'Did we take damage and are there people who care???
                    If bSendHPUpdateMsg = True Then
                        'for asynchronous persistence, set it to now because we don't want to spam the server...
                        'goEntity(X).lUpdatePrimaryWithHPRequest = glCurrentCycle

                        'Now, is someone in the environment?
                        If .ParentEnvir.lPlayersInEnvirCnt > 0 Then
                            'Yup, update everyone, if necessary
                            goMsgSys.CheckForBroadcastEntityHP(oEntity, oCurDef)
                            'goMsgSys.BroadcastToEnvironmentClients(goMsgSys.CreateEntityHPMsg(goEntity(X), oCurDef), goEntity(X).ParentEnvir)
                        End If
                    End If

                    'Check for Micro explosion
                    If (oCurDef.lExtendedFlags And Epica_Entity_Def.eExtendedFlagValues.ContainsMicroTech) <> 0 Then
                        If (.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 OrElse (.iCombatTactics And eiBehaviorPatterns.eEngagement_Hold_Fire) <> 0 Then
                            If .lNextFireWpnEvent <= glCurrentCycle OrElse (.iCombatTactics And eiBehaviorPatterns.eEngagement_Hold_Fire) <> 0 Then
                                If CInt(Rnd() * oCurDef.lMicroReliabilityTest) = 1 Then
                                    'Destroy the unit
                                    .StructuralHP = 0
                                    bUnitDestroyed = True

                                    Dim yMicroExp(25) As Byte
                                    System.BitConverter.GetBytes(GlobalMessageCode.eMicroExplosion).CopyTo(yMicroExp, 0)
                                    .GetGUIDAsString.CopyTo(yMicroExp, 2)
                                    .ParentEnvir.GetGUIDAsString.CopyTo(yMicroExp, 8)
                                    System.BitConverter.GetBytes(.LocX).CopyTo(yMicroExp, 14)
                                    System.BitConverter.GetBytes(.LocZ).CopyTo(yMicroExp, 18)
                                    System.BitConverter.GetBytes(.lOwnerID).CopyTo(yMicroExp, 22)
                                    goMsgSys.SendToPrimary(yMicroExp)
                                End If
                            End If
                        End If
                    End If

                    'Now, if the unit is still alive. :)
                    If bUnitDestroyed = False Then
                        'ok, time to fire back
                        If (.CurrentStatus And elUnitStatus.eUnitEngaged) <> 0 Then
                            .lLastEngagement = glCurrentCycle

                            'Ok, check our targets
                            'MSC - 05/13/08 - added for new fire weapon optimization
                            Dim bSendTargetSwitch As Boolean = False
                            For Y = 0 To 3
                                If .lTargetsServerIdx(Y) <> .lPreviousTargetServerIdx(Y) Then
                                    bSendTargetSwitch = True
                                    .lPreviousTargetServerIdx(Y) = .lTargetsServerIdx(Y)
                                End If
                            Next Y
                            If bSendTargetSwitch = True Then
                                'ok, send our target switch
                                goMsgSys.BroadcastToEnvironmentClients(goMsgSys.CreateSetTargetMessage(oEntity), oEntity.ParentEnvir)
                            End If
                            '//05/13/08

                            If oCurDef.JamStrength <> 0 Then
                                If mlJamState = 2 Then
                                    'anti-disruption resistance and scan res jamming only - effect 2
                                    If oCurDef.JamEffect = 2 Then .ProcessJamming()
                                ElseIf mlJamState = 3 Then
                                    'anti-jamming jamming only - effects 4 and 5
                                    If oCurDef.JamEffect = 4 OrElse oCurDef.JamEffect = 5 Then .ProcessJamming()
                                ElseIf mlJamState = 4 Then
                                    'all remaining jamming
                                    If oCurDef.JamEffect <> 2 AndAlso oCurDef.JamEffect <> 4 AndAlso oCurDef.JamEffect <> 5 Then .ProcessJamming()
                                End If
                            End If

                            If .lNextFireWpnEvent <= glCurrentCycle Then
                                .lNextFireWpnEvent = Int32.MaxValue

                                If (.CurrentStatus And elUnitStatus.eUnitCloaked) <> 0 Then
                                    .ClearTargetLists(False, -1)
                                    If .lNextWpnEvent = Int32.MaxValue Then RemoveEntityInCombat(.ObjectID, .ObjTypeID)
                                    Continue For
                                End If

                                If .lWpnNextFireCycle Is Nothing Then ReDim .lWpnNextFireCycle(oCurDef.WeaponDefUB)
                                If .lWpnNextFireCycle.GetUpperBound(0) <> oCurDef.WeaponDefUB Then ReDim Preserve .lWpnNextFireCycle(oCurDef.WeaponDefUB)

                                For Y = 0 To oCurDef.WeaponDefUB
                                    If .lWpnNextFireCycle(Y) <= glCurrentCycle Then
                                        If (.CurrentStatus And oCurDef.WeaponDefs(Y).lEntityStatusGroup) <> 0 Then
                                            oWeapon = oCurDef.WeaponDefs(Y)

                                            Select Case oWeapon.WpnGroup
                                                Case WeaponGroupType.BombWeaponGroup
                                                    Continue For
                                                Case WeaponGroupType.PointDefenseWeaponGroup
                                                    If (.iCombatTactics And eiBehaviorPatterns.eHoldPointDefenseWpnGrp) <> 0 Then
                                                        .lNextFireWpnEvent = Math.Min(.lNextFireWpnEvent, .lWpnNextFireCycle(Y))
                                                        Continue For
                                                    End If
                                                Case WeaponGroupType.PrimaryWeaponGroup
                                                    If (.iCombatTactics And eiBehaviorPatterns.eHoldPrimaryWpnGrp) <> 0 Then
                                                        .lNextFireWpnEvent = Math.Min(.lNextFireWpnEvent, .lWpnNextFireCycle(Y))
                                                        Continue For
                                                    End If
                                                Case WeaponGroupType.SecondaryWeaponGroup
                                                    If (.iCombatTactics And eiBehaviorPatterns.eHoldSecondaryWpnGrp) <> 0 Then
                                                        .lNextFireWpnEvent = Math.Min(.lNextFireWpnEvent, .lWpnNextFireCycle(Y))
                                                        Continue For
                                                    End If
                                            End Select

                                            'MSC 05/13/08 - added for new fire weapon optimization
                                            Dim ySideFiring As Byte = oWeapon.ArcID

                                            'Clear our target so we can get it...
                                            oTarget = Nothing

                                            'POINT DEFENSE STUFF!!!
                                            Dim bHaveTarget As Boolean = False
                                            If oWeapon.WpnGroup = WeaponGroupType.PointDefenseWeaponGroup Then
                                                If (.CurrentStatus And elUnitStatus.eRadarOperational) = 0 Then Continue For
                                                'ok, handle point defense... 
                                                'point defense weapons will attempt to target in this order:
                                                'missiles
                                                Dim lMissileTargetRange As Int32 = 0
                                                Dim oMissileTarget As MissileMgr.Missile = .GetMissileTarget(ySideFiring, oWeapon.lRange, lMissileTargetRange, oCurDef)
                                                If oMissileTarget Is Nothing = False Then
                                                    ProcessPDSToMissileShot(oEntity, oCurDef, oMissileTarget, oWeapon, lMissileTargetRange)

                                                    'regardless of hit or miss, update our fire cycle...
                                                    .lWpnNextFireCycle(Y) = glCurrentCycle + oWeapon.ROF_Delay
                                                    .lNextFireWpnEvent = Math.Min(.lNextFireWpnEvent, .lWpnNextFireCycle(Y))
                                                    Continue For
                                                End If

                                                'mines
                                                'TODO: Implement mine point defense...

                                                'fighters
                                                Dim lFighterIdx As Int32 = -1
                                                If oWeapon.ArcID = UnitArcs.eAllArcs Then
                                                    'ok, find the closest
                                                    Dim lLowestRange As Int32 = Int32.MaxValue
                                                    For lTmpIdx As Int32 = 0 To 3
                                                        If .lFighterTargetServerIdx(lTmpIdx) > -1 Then
                                                            If .lFighterTargetRange(lTmpIdx) < lLowestRange Then
                                                                lLowestRange = .lFighterTargetRange(lTmpIdx)
                                                                lFighterIdx = lTmpIdx
                                                            End If
                                                        End If
                                                    Next lTmpIdx
                                                Else
                                                    lFighterIdx = oWeapon.ArcID
                                                End If
                                                If lFighterIdx > -1 Then
                                                    'ok, fire at the fighter...
                                                    lRange = .lFighterTargetRange(lFighterIdx)
                                                    lFighterIdx = .lFighterTargetServerIdx(lFighterIdx)

                                                    If lFighterIdx > -1 Then
                                                        oTarget = goEntity(lFighterIdx)
                                                        If oTarget Is Nothing = False Then
                                                            bHaveTarget = True
                                                        End If
                                                    End If
                                                End If

                                                'everything else... run like normal...
                                            ElseIf oWeapon.WpnGroup = WeaponGroupType.PrimaryWeaponGroup Then
                                                'Ok, primary weapon groups shoot ONLY at primary targets. if the weapon cannot fire at the primary target, it follows the rules
                                                '  if the unit is not moving, the weapon acts as a secondary weapon
                                                '  if the unit is moving by the player, the weapon acts as a secondary weapon
                                                '  if the unit is moving but not by the player, the weapon holds fire
                                                '  the assumption is that secondary weapons fire at primary targets when capable

                                                'So, if the unit is moving....
                                                If ySideFiring <> UnitArcs.eAllArcs AndAlso (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                                                    'Is the unit moved by the player?
                                                    'If (.CurrentStatus And elUnitStatus.eMovedByPlayer) = 0 Then
                                                    'it was not, so now we check... if the weapon's arc is pointing towards the primary
                                                    Dim lTempTargetServerIdx As Int32 = .lPrimaryTargetServerIdx
                                                    If lTempTargetServerIdx <> -1 AndAlso .lTargetsServerIdx(ySideFiring) <> lTempTargetServerIdx Then
                                                        'Ok, so we set our target to nothing
                                                        oTarget = Nothing
                                                        'and set our bHaveTarget, this avoids the combat engine from assigning the target
                                                        bHaveTarget = True
                                                    End If
                                                    'End If
                                                End If
                                            End If

                                            If bHaveTarget = False Then
                                                If oWeapon.ArcID = UnitArcs.eAllArcs Then
                                                    lRange = .lPrimaryTargetRange

                                                    'MSC 05/13/08 - added for new fire weapon optimization
                                                    For Z As Int32 = 0 To 3
                                                        If .lTargetsServerIdx(Z) = .lPrimaryTargetServerIdx Then
                                                            ySideFiring = CByte(Z)
                                                            Exit For
                                                        End If
                                                    Next Z
                                                    If ySideFiring = UnitArcs.eAllArcs Then ySideFiring = UnitArcs.eForwardArc
                                                Else : lRange = .Ranges(oWeapon.ArcID)
                                                End If

                                                'MSC - 08/02/07 - REMARKED OUT PER JAMIE'S REQUEST
                                                'If (oWeapon.lRange >= lRange) AndAlso (oWeapon.lAmmoCap = -1 OrElse .lWpnAmmoCnt(Y) > 0) Then
                                                'If (oWeapon.lRange >= lRange) Then

                                                If oWeapon.ArcID = UnitArcs.eAllArcs Then
                                                    'use the Primary Target right now... but eventually, point defense needs to 
                                                    '  be factored in as well
                                                    Dim lTempTargetServerIdx As Int32 = .lPrimaryTargetServerIdx
                                                    If lTempTargetServerIdx <> -1 Then
                                                        If glEntityIdx(lTempTargetServerIdx) > 0 Then
                                                            oTarget = goEntity(lTempTargetServerIdx)
                                                        Else : .lPrimaryTargetServerIdx = -1
                                                        End If
                                                    End If
                                                Else
                                                    Dim lTempTargetServerIdx As Int32 = .lTargetsServerIdx(oWeapon.ArcID)
                                                    If lTempTargetServerIdx <> -1 Then
                                                        If glEntityIdx(lTempTargetServerIdx) > 0 Then
                                                            oTarget = goEntity(lTempTargetServerIdx)
                                                        Else : .lTargetsServerIdx(oWeapon.ArcID) = -1
                                                        End If
                                                    End If
                                                End If

                                            End If

                                            If oTarget Is Nothing = False Then

                                                If .ParentEnvir.ObjectID <> oTarget.ParentEnvir.ObjectID OrElse .ParentEnvir.ObjTypeID <> oTarget.ParentEnvir.ObjTypeID Then
                                                    gfrmDisplayForm.AddEventLine("Shooting outside my environment!")
                                                    oTarget = Nothing
                                                    .bForceAggressionTest = True
                                                    SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)
                                                    Continue For
                                                End If

                                                If mbDoTheNewCode = True Then
                                                    Try
                                                        If (oTarget.CurrentStatus And elUnitStatus.eUnitEngaged) = 0 AndAlso oTarget.bForceAggressionTest = False Then
                                                            oTarget.bForceAggressionTest = True
                                                            SyncLockMovementRegisters(MovementCommand.AddEntityMoving, oTarget.ServerIndex, oTarget.ObjectID)
                                                        End If
                                                    Catch
                                                    End Try
                                                End If

                                                If (oTarget.CurrentStatus And elUnitStatus.eUnitCloaked) <> 0 Then
                                                    'Ok, we're shooting at a cloaked target, clear it
                                                    .bForceAggressionTest = True
                                                    SyncLockMovementRegisters(MovementCommand.AddEntityMoving, .ServerIndex, .ObjectID) 'AddEntityMoving(.ServerIndex, .ObjectID)
                                                    Continue For
                                                End If
                                                'If .lOwnerID = oTarget.lOwnerID Then
                                                '	gfrmDisplayForm.AddEventLine("Shooting at my own guy HandleCombat.")
                                                'End If

                                                'If oTarget.bInCombatRegister = False Then
                                                '	AddEntityCombat(oTarget.ServerIndex, oTarget.ObjectID)
                                                'End If

                                                'set our oTmpDef
                                                If oTarget.lEntityDefServerIndex <> -1 Then
                                                    If glEntityDefIdx(oTarget.lEntityDefServerIndex) <> -1 Then
                                                        oTargetDef = goEntityDefs(oTarget.lEntityDefServerIndex)
                                                    End If
                                                End If

                                                'Ok, if we are a missile
                                                If oWeapon.yWeaponType >= WeaponType.eMissile_Color_1 AndAlso oWeapon.yWeaponType <= WeaponType.eMissile_Color_9 Then
                                                    'Then simply add the missile weapon shot...
                                                    Dim lMissileID As Int32 = goMissileMgr.AddMissileShot(oTarget, oEntity, oCurDef, oWeapon, oTargetDef)
                                                    'Update our fired weapon
                                                    .lWpnNextFireCycle(Y) = glCurrentCycle + oWeapon.ROF_Delay
                                                    .lNextFireWpnEvent = Math.Min(.lNextFireWpnEvent, .lWpnNextFireCycle(Y))
                                                    'Continue For   'MSC - 09/20/08
                                                Else
                                                    'ok, determine if we hit...
                                                    If oWeapon.WpnGroup = WeaponGroupType.PointDefenseWeaponGroup Then
                                                        If (oTarget.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                                                            lBaseToHit = GetBaseToHit(oEntity, oCurDef, oTarget.VelX, oTarget.VelZ, oTargetDef.HullSizeTargetMod, oWeapon, lRange, 1, oTargetDef.HullSize)
                                                        Else
                                                            lBaseToHit = GetBaseToHit(oEntity, oCurDef, 0, 0, oTargetDef.HullSizeTargetMod, oWeapon, lRange, 1, oTargetDef.HullSize)
                                                        End If
                                                    Else
                                                        If (oTarget.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                                                            lBaseToHit = GetBaseToHit(oEntity, oCurDef, oTarget.VelX, oTarget.VelZ, oTargetDef.HullSizeTargetMod, oWeapon, lRange, 0, oTargetDef.HullSize)
                                                        Else
                                                            lBaseToHit = GetBaseToHit(oEntity, oCurDef, 0, 0, oTargetDef.HullSizeTargetMod, oWeapon, lRange, 0, oTargetDef.HullSize)
                                                        End If
                                                    End If

                                                    lRoll = CInt(Rnd() * 100) + 1

                                                    If lRoll < lBaseToHit Then

                                                        'add our event, we hit
                                                        oEvent = New WeaponEvent()

                                                        'what side will i hit?
                                                        oEvent.ySideHit = WhatSideDoIHit(oEntity, oTarget)

                                                        If lRoll <= CInt(.Owner.CriticalHitMax) + CInt(oCurDef.CriticalHitModifier) Then
                                                            oEvent.yCritical = eyCriticalHitType.NormalCriticalHit
                                                            If oCurDef.HullSize <= gl_MAX_FIGHTER_HULL AndAlso (oCurDef.yChassisType And (ChassisType.eAtmospheric Or ChassisType.eSpaceBased)) <> 0 Then
                                                                If (.iTargetingTactics And (eiTacticalAttrs.eFighterTargetEngines Or eiTacticalAttrs.eFighterTargetHangar Or eiTacticalAttrs.eFighterTargetRadar Or eiTacticalAttrs.eFighterTargetShields Or eiTacticalAttrs.eFighterTargetWeapons)) <> 0 Then
                                                                    If glCurrentCycle - oTarget.lLastFighterCritCycle > oTargetDef.lFighterCritDelay Then
                                                                        oTarget.lLastFighterCritCycle = glCurrentCycle
                                                                        'ok, roll...
                                                                        Dim lCritRoll As Int32 = CInt(Rnd() * 4)
                                                                        Select Case lCritRoll
                                                                            Case 0
                                                                                If (.iTargetingTactics And eiTacticalAttrs.eFighterTargetEngines) <> 0 Then oEvent.yCritical = eyCriticalHitType.FighterSubsystemEngine
                                                                            Case 1
                                                                                If (.iTargetingTactics And eiTacticalAttrs.eFighterTargetHangar) <> 0 Then oEvent.yCritical = eyCriticalHitType.FighterSubsystemHangar
                                                                            Case 2
                                                                                If (.iTargetingTactics And eiTacticalAttrs.eFighterTargetRadar) <> 0 Then oEvent.yCritical = eyCriticalHitType.FighterSubsystemRadar
                                                                            Case 3
                                                                                If (.iTargetingTactics And eiTacticalAttrs.eFighterTargetShields) <> 0 Then oEvent.yCritical = eyCriticalHitType.FighterSubsystemShields
                                                                            Case Else
                                                                                If (.iTargetingTactics And eiTacticalAttrs.eFighterTargetWeapons) <> 0 Then oEvent.yCritical = eyCriticalHitType.FighterSubsystemWeapons
                                                                        End Select
                                                                    End If
                                                                End If
                                                            End If
                                                            If oTargetDef.HullSize < 670 AndAlso (oTargetDef.yChassisType And ChassisType.eGroundBased) <> 0 Then
                                                                oEvent.yCritical = eyCriticalHitType.NoCriticalHit
                                                            End If
                                                        Else : oEvent.yCritical = eyCriticalHitType.NoCriticalHit
                                                        End If

                                                        'Now, our damages
                                                        If .DamageMod <> 0 Then
                                                            Dim fMult As Single = 1.0F + (.DamageMod / 100.0F)
                                                            If oWeapon.PiercingMaxDmg > 0 Then oEvent.lPierceDmg = CInt(((Rnd() * oWeapon.PiercingMaxMinRange) + oWeapon.PiercingMinDmg) * fMult)
                                                            If oWeapon.ImpactMaxDmg > 0 Then oEvent.lImpactDmg = CInt(((Rnd() * oWeapon.ImpactMaxMinRange) + oWeapon.ImpactMinDmg) * fMult)
                                                            If oWeapon.BeamMaxDmg > 0 Then oEvent.lBeamDmg = CInt(((Rnd() * oWeapon.BeamMaxMinRange) + oWeapon.BeamMinDmg) * fMult)
                                                            If oWeapon.ECMMaxDmg > 0 Then oEvent.lECMDmg = CInt(((Rnd() * oWeapon.ECMMaxMinRange) + oWeapon.ECMMinDmg) * fMult)
                                                            If oWeapon.FlameMaxDmg > 0 Then oEvent.lFlameDmg = CInt(((Rnd() * oWeapon.FlameMaxMinRange) + oWeapon.FlameMinDmg) * fMult)
                                                            If oWeapon.ChemicalMaxDmg > 0 Then oEvent.lChemicalDmg = CInt(((Rnd() * oWeapon.ChemicalMaxMinRange) + oWeapon.ChemicalMinDmg) * fMult)
                                                        Else
                                                            If oWeapon.PiercingMaxDmg > 0 Then oEvent.lPierceDmg = CInt(Rnd() * oWeapon.PiercingMaxMinRange) + oWeapon.PiercingMinDmg
                                                            If oWeapon.ImpactMaxDmg > 0 Then oEvent.lImpactDmg = CInt(Rnd() * oWeapon.ImpactMaxMinRange) + oWeapon.ImpactMinDmg
                                                            If oWeapon.BeamMaxDmg > 0 Then oEvent.lBeamDmg = CInt(Rnd() * oWeapon.BeamMaxMinRange) + oWeapon.BeamMinDmg
                                                            If oWeapon.ECMMaxDmg > 0 Then oEvent.lECMDmg = CInt(Rnd() * oWeapon.ECMMaxMinRange) + oWeapon.ECMMinDmg
                                                            If oWeapon.FlameMaxDmg > 0 Then oEvent.lFlameDmg = CInt(Rnd() * oWeapon.FlameMaxMinRange) + oWeapon.FlameMinDmg
                                                            If oWeapon.ChemicalMaxDmg > 0 Then oEvent.lChemicalDmg = CInt(Rnd() * oWeapon.ChemicalMaxMinRange) + oWeapon.ChemicalMinDmg
                                                        End If


                                                        'from... and to
                                                        oEvent.lFromIdx = .ServerIndex
                                                        oEvent.lToIdx = oTarget.ServerIndex
                                                        oEvent.FromPlayerID = .lOwnerID

                                                        If oWeapon.fPulseDegradation <> 0 AndAlso oWeapon.yWeaponType >= WeaponType.eShortGreenPulse AndAlso oWeapon.yWeaponType <= WeaponType.eShortPurplePulse Then
                                                            oEvent.lFromX = .LocX
                                                            oEvent.lFromZ = .LocZ
                                                            oEvent.fDegradation = oWeapon.fPulseDegradation
                                                        End If

                                                        'Now, when will it hit??? guesstimate
                                                        'oEvent.lEventCycle = glCurrentCycle + CInt(lRange / oWeapon.WeaponSpeed)
                                                        oEvent.lEventCycle = glCurrentCycle + CInt(lRange * oWeapon.fOneOverWeaponSpeed)
                                                        oEvent.AOERange = oWeapon.AOERange
                                                        oEvent.yAOEDmgType = oWeapon.yAOEDmgType
                                                        oEvent.yAttackerType = oCurDef.yHullType

                                                        'Add the event
                                                        yTemp = 0
                                                        For lTmp = 0 To oTarget.lWpnEventUB
                                                            If oTarget.yWpnEventUsed(lTmp) = 0 Then
                                                                'add it
                                                                yTemp = 1
                                                                oTarget.oWpnEvents(lTmp) = oEvent
                                                                oTarget.yWpnEventUsed(lTmp) = 255
                                                                oTarget.lNextWpnEvent = Math.Min(oTarget.lNextWpnEvent, oEvent.lEventCycle)
                                                                Exit For
                                                            End If
                                                        Next lTmp

                                                        If yTemp = 0 Then
                                                            oTarget.lWpnEventUB += 1
                                                            ReDim Preserve oTarget.oWpnEvents(oTarget.lWpnEventUB)
                                                            ReDim Preserve oTarget.yWpnEventUsed(oTarget.lWpnEventUB)
                                                            oTarget.oWpnEvents(oTarget.lWpnEventUB) = oEvent
                                                            oTarget.lNextWpnEvent = Math.Min(oTarget.lNextWpnEvent, oEvent.lEventCycle)
                                                            oTarget.yWpnEventUsed(oTarget.lWpnEventUB) = 255
                                                        End If

                                                        'Update our fired weapon
                                                        .lWpnNextFireCycle(Y) = glCurrentCycle + oWeapon.ROF_Delay '+ CInt(Rnd() * 2 - 1)

                                                        'Ok, a weapon was fired, so set up our message to send...
                                                        If .ParentEnvir.lPlayersInEnvirCnt > 0 Then
                                                            Dim yHitType As Byte = 1
                                                            If oTarget.ShieldHP > 0 AndAlso (oTarget.CurrentStatus And elUnitStatus.eShieldOperational) <> 0 Then yHitType = 2
                                                            If .ParentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                                                                '40000 comes from max entity clip plane on client,the 125 comes from
                                                                '  10,000,000 / 40000 = 250 / 2 = 125 to offset to 0
                                                                Dim lCameraPosX As Int32 = (.LocX \ 40000) + 125
                                                                Dim lCameraPosZ As Int32 = (.LocZ \ 40000) + 125
                                                                'goMsgSys.BroadcastToEnvironmentClients_Filter(goMsgSys.CreateFireWeaponMessage(.ObjectID, _
                                                                '  .ObjTypeID, oTarget.ObjectID, oTarget.ObjTypeID, oWeapon.WeaponType, yHitType), .ParentEnvir, lCameraPosX, lCameraPosZ)
                                                                If oWeapon.WpnGroup = WeaponGroupType.PointDefenseWeaponGroup Then
                                                                    goMsgSys.BroadcastToEnvironmentClients_Filter(goMsgSys.CreatePointDefenseMessage(GlobalMessageCode.ePDS_AtUnitHit, oEntity, oTarget.ObjectID, oTarget.ObjTypeID, oWeapon.yWeaponType, yHitType), _
                                                                        .ParentEnvir, lCameraPosX, lCameraPosZ)
                                                                Else
                                                                    goMsgSys.BroadcastToEnvironmentClients_Filter(goMsgSys.CreateFireWeaponMessage(.ObjectID, .ObjTypeID, ySideFiring, oWeapon.yWeaponType, yHitType, oWeapon.AOERange), _
                                                                        .ParentEnvir, lCameraPosX, lCameraPosZ)
                                                                End If
                                                            Else
                                                                'goMsgSys.BroadcastToEnvironmentClients(goMsgSys.CreateFireWeaponMessage(.ObjectID, _
                                                                '  .ObjTypeID, oTarget.ObjectID, oTarget.ObjTypeID, oWeapon.WeaponType, yHitType), .ParentEnvir)
                                                                If oWeapon.WpnGroup = WeaponGroupType.PointDefenseWeaponGroup Then
                                                                    goMsgSys.BroadcastToEnvironmentClients(goMsgSys.CreatePointDefenseMessage(GlobalMessageCode.ePDS_AtUnitHit, oEntity, oTarget.ObjectID, oTarget.ObjTypeID, oWeapon.yWeaponType, yHitType), .ParentEnvir)
                                                                Else
                                                                    goMsgSys.BroadcastToEnvironmentClients(goMsgSys.CreateFireWeaponMessage(.ObjectID, .ObjTypeID, ySideFiring, oWeapon.yWeaponType, yHitType, oWeapon.AOERange), .ParentEnvir)
                                                                End If
                                                            End If
                                                        End If

                                                    Else
                                                        'we missed, still need to increment the fire cycle
                                                        .lWpnNextFireCycle(Y) = glCurrentCycle + oWeapon.ROF_Delay '+ CInt(Rnd() * 2 - 1)

                                                        'Is the weapon AOE?
                                                        If oWeapon.AOERange > 0 AndAlso oWeapon.yAOEDmgType > 0 Then
                                                            oEvent = New WeaponEvent()
                                                            'Now, our damages
                                                            oEvent.lToIdx = oTarget.ServerIndex
                                                            oEvent.AOERange = oWeapon.AOERange
                                                            oEvent.yAOEDmgType = oWeapon.yAOEDmgType
                                                            If .DamageMod <> 0 Then
                                                                Dim fMult As Single = 1.0F + (.DamageMod / 100.0F)

                                                                Select Case oWeapon.yAOEDmgType
                                                                    Case eyAOEDmgType.eBeamDmg
                                                                        If oWeapon.BeamMaxDmg > 0 Then oEvent.lBeamDmg = CInt(((Rnd() * oWeapon.BeamMaxMinRange) + oWeapon.BeamMinDmg) * fMult)
                                                                    Case eyAOEDmgType.eChemicalDmg
                                                                        If oWeapon.ChemicalMaxDmg > 0 Then oEvent.lChemicalDmg = CInt(((Rnd() * oWeapon.ChemicalMaxMinRange) + oWeapon.ChemicalMinDmg) * fMult)
                                                                    Case eyAOEDmgType.eECMDmg
                                                                        If oWeapon.ECMMaxDmg > 0 Then oEvent.lECMDmg = CInt(((Rnd() * oWeapon.ECMMaxMinRange) + oWeapon.ECMMinDmg) * fMult)
                                                                    Case eyAOEDmgType.eFlameDmg
                                                                        If oWeapon.FlameMaxDmg > 0 Then oEvent.lFlameDmg = CInt(((Rnd() * oWeapon.FlameMaxMinRange) + oWeapon.FlameMinDmg) * fMult)
                                                                    Case eyAOEDmgType.eImpactDmg
                                                                        If oWeapon.ImpactMaxDmg > 0 Then oEvent.lImpactDmg = CInt(((Rnd() * oWeapon.ImpactMaxMinRange) + oWeapon.ImpactMinDmg) * fMult)
                                                                    Case eyAOEDmgType.ePierceDmg
                                                                        If oWeapon.PiercingMaxDmg > 0 Then oEvent.lPierceDmg = CInt(((Rnd() * oWeapon.PiercingMaxMinRange) + oWeapon.PiercingMinDmg) * fMult)
                                                                    Case Else
                                                                        If oWeapon.PiercingMaxDmg > 0 Then oEvent.lPierceDmg = CInt(((Rnd() * oWeapon.PiercingMaxMinRange) + oWeapon.PiercingMinDmg) * fMult)
                                                                        If oWeapon.ImpactMaxDmg > 0 Then oEvent.lImpactDmg = CInt(((Rnd() * oWeapon.ImpactMaxMinRange) + oWeapon.ImpactMinDmg) * fMult)
                                                                        If oWeapon.BeamMaxDmg > 0 Then oEvent.lBeamDmg = CInt(((Rnd() * oWeapon.BeamMaxMinRange) + oWeapon.BeamMinDmg) * fMult)
                                                                        If oWeapon.ECMMaxDmg > 0 Then oEvent.lECMDmg = CInt(((Rnd() * oWeapon.ECMMaxMinRange) + oWeapon.ECMMinDmg) * fMult)
                                                                        If oWeapon.FlameMaxDmg > 0 Then oEvent.lFlameDmg = CInt(((Rnd() * oWeapon.FlameMaxMinRange) + oWeapon.FlameMinDmg) * fMult)
                                                                        If oWeapon.ChemicalMaxDmg > 0 Then oEvent.lChemicalDmg = CInt(((Rnd() * oWeapon.ChemicalMaxMinRange) + oWeapon.ChemicalMinDmg) * fMult)
                                                                End Select

                                                            Else
                                                                Select Case oWeapon.yAOEDmgType
                                                                    Case eyAOEDmgType.eBeamDmg
                                                                        If oWeapon.BeamMaxDmg > 0 Then oEvent.lBeamDmg = (CInt(Rnd() * oWeapon.BeamMaxMinRange) + oWeapon.BeamMinDmg) '+ .DamageMod
                                                                    Case eyAOEDmgType.eChemicalDmg
                                                                        If oWeapon.ChemicalMaxDmg > 0 Then oEvent.lChemicalDmg = (CInt(Rnd() * oWeapon.ChemicalMaxMinRange) + oWeapon.ChemicalMinDmg) '+ .DamageMod
                                                                    Case eyAOEDmgType.eECMDmg
                                                                        If oWeapon.ECMMaxDmg > 0 Then oEvent.lECMDmg = (CInt(Rnd() * oWeapon.ECMMaxMinRange) + oWeapon.ECMMinDmg) '+ .DamageMod
                                                                    Case eyAOEDmgType.eFlameDmg
                                                                        If oWeapon.FlameMaxDmg > 0 Then oEvent.lFlameDmg = (CInt(Rnd() * oWeapon.FlameMaxMinRange) + oWeapon.FlameMinDmg) '+ .DamageMod
                                                                    Case eyAOEDmgType.eImpactDmg
                                                                        If oWeapon.ImpactMaxDmg > 0 Then oEvent.lImpactDmg = (CInt(Rnd() * oWeapon.ImpactMaxMinRange) + oWeapon.ImpactMinDmg) '+ .DamageMod
                                                                    Case eyAOEDmgType.ePierceDmg
                                                                        If oWeapon.PiercingMaxDmg > 0 Then oEvent.lPierceDmg = (CInt(Rnd() * oWeapon.PiercingMaxMinRange) + oWeapon.PiercingMinDmg) '+ .DamageMod
                                                                    Case Else
                                                                        If oWeapon.PiercingMaxDmg > 0 Then oEvent.lPierceDmg = (CInt(Rnd() * oWeapon.PiercingMaxMinRange) + oWeapon.PiercingMinDmg) '+ .DamageMod
                                                                        If oWeapon.ImpactMaxDmg > 0 Then oEvent.lImpactDmg = (CInt(Rnd() * oWeapon.ImpactMaxMinRange) + oWeapon.ImpactMinDmg) '+ .DamageMod
                                                                        If oWeapon.BeamMaxDmg > 0 Then oEvent.lBeamDmg = (CInt(Rnd() * oWeapon.BeamMaxMinRange) + oWeapon.BeamMinDmg) '+ .DamageMod
                                                                        If oWeapon.ECMMaxDmg > 0 Then oEvent.lECMDmg = (CInt(Rnd() * oWeapon.ECMMaxMinRange) + oWeapon.ECMMinDmg) '+ .DamageMod
                                                                        If oWeapon.FlameMaxDmg > 0 Then oEvent.lFlameDmg = (CInt(Rnd() * oWeapon.FlameMaxMinRange) + oWeapon.FlameMinDmg) '+ .DamageMod
                                                                        If oWeapon.ChemicalMaxDmg > 0 Then oEvent.lChemicalDmg = (CInt(Rnd() * oWeapon.ChemicalMaxMinRange) + oWeapon.ChemicalMinDmg) '+ .DamageMod
                                                                End Select
                                                            End If
                                                            oEvent.FromPlayerID = .lOwnerID
                                                            oEvent.yAttackerType = oCurDef.yHullType
                                                            Dim lRndX As Int32 = oTarget.LocX + CInt((Rnd() * 100) - 50)
                                                            Dim lRndZ As Int32 = oTarget.LocZ + CInt((Rnd() * 100) - 50)
                                                            ApplyAOEDamage(oEvent, lRndX, lRndZ, .ParentEnvir, .lOwnerID)
                                                        End If

                                                        'and send our missed fire weapon message
                                                        If .ParentEnvir.lPlayersInEnvirCnt > 0 Then
                                                            If .ParentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                                                                '40000 comes from max entity clip plane on client,the 125 comes from
                                                                '  10,000,000 / 40000 = 250 / 2 = 125 to offset to 0
                                                                Dim lCameraPosX As Int32 = (.LocX \ 40000) + 125
                                                                Dim lCameraPosZ As Int32 = (.LocZ \ 40000) + 125
                                                                'goMsgSys.BroadcastToEnvironmentClients_Filter(goMsgSys.CreateFireWeaponMessage(.ObjectID, _
                                                                '  .ObjTypeID, oTarget.ObjectID, oTarget.ObjTypeID, oWeapon.WeaponType, 0), .ParentEnvir, lCameraPosX, lCameraPosZ)
                                                                If oWeapon.WpnGroup = WeaponGroupType.PointDefenseWeaponGroup Then
                                                                    goMsgSys.BroadcastToEnvironmentClients_Filter(goMsgSys.CreatePointDefenseMessage(GlobalMessageCode.ePDS_AtUnitMiss, oEntity, oTarget.ObjectID, oTarget.ObjTypeID, oWeapon.yWeaponType, 0), .ParentEnvir, lCameraPosX, lCameraPosZ)
                                                                Else
                                                                    goMsgSys.BroadcastToEnvironmentClients_Filter(goMsgSys.CreateFireWeaponMessage(.ObjectID, .ObjTypeID, ySideFiring, oWeapon.yWeaponType, 0, oWeapon.AOERange), .ParentEnvir, lCameraPosX, lCameraPosZ)
                                                                End If
                                                            Else
                                                                'goMsgSys.BroadcastToEnvironmentClients(goMsgSys.CreateFireWeaponMessage(.ObjectID, .ObjTypeID, oTarget.ObjectID, oTarget.ObjTypeID, oWeapon.WeaponType, 0), .ParentEnvir)
                                                                If oWeapon.WpnGroup = WeaponGroupType.PointDefenseWeaponGroup Then
                                                                    goMsgSys.BroadcastToEnvironmentClients(goMsgSys.CreatePointDefenseMessage(GlobalMessageCode.ePDS_AtUnitMiss, oEntity, oTarget.ObjectID, oTarget.ObjTypeID, oWeapon.yWeaponType, 0), .ParentEnvir)
                                                                Else
                                                                    goMsgSys.BroadcastToEnvironmentClients(goMsgSys.CreateFireWeaponMessage(.ObjectID, .ObjTypeID, ySideFiring, oWeapon.yWeaponType, 0, oWeapon.AOERange), .ParentEnvir)
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If

                                                'MSC - 08/02/2007 - REMARKED OUT PER JAMIE'S REQUEST
                                                ''Regardless of hit or miss, if the weapon is ammo based
                                                'If oWeapon.lAmmoCap <> -1 Then
                                                '    'Decrease the ammo count
                                                '    .lWpnAmmoCnt(Y) -= 1

                                                '    If .lWpnAmmoCnt(Y) = 0 Then
                                                '        'Ok, send off our message for reload now to the primary server
                                                '        goMsgSys.SendReloadRequestMsg(goEntity(X), oWeapon)
                                                '    End If
                                                'End If

                                                'At this point, a conflict is present... So... two updates:
                                                'EngageEnemy for the attacker
                                                .Owner.CheckForAlert(PlayerAlertType.eEngagedEnemy, .ObjectID, .ObjTypeID, .ParentEnvir.ObjectID, .ParentEnvir.ObjTypeID, oTarget.Owner.ObjectID, .LocX, .LocZ)
                                                'UnderAttack for the target
                                                oTarget.Owner.CheckForAlert(PlayerAlertType.eUnderAttack, oTarget.ObjectID, oTarget.ObjTypeID, .ParentEnvir.ObjectID, .ParentEnvir.ObjTypeID, .Owner.ObjectID, .LocX, .LocZ)

                                                'ok, regardless of hit or miss, if the unit was shot at and it is not
                                                '  engaged, then the unit needs to be smart about it...
                                                If (oTarget.CurrentStatus And elUnitStatus.eUnitEngaged) = 0 Then
                                                    AddEntityInCombat(oTarget.ObjectID, oTarget.ObjTypeID, oTarget.ServerIndex)
                                                    If (oTarget.iCombatTactics And eiBehaviorPatterns.eEngagement_Evade) <> 0 Then
                                                        'Ok, unit needs to move if it isn't
                                                        If (oTarget.CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
                                                            If (oTarget.CurrentStatus And elUnitStatus.eEngineOperational) <> 0 AndAlso _
                                                               oTarget.MaxSpeed > 0 AndAlso oTarget.Maneuver > 0 Then
                                                                goMsgSys.SendAIMoveRequestToPathfinding(oTarget, oTarget.lExtendedCT_1, oTarget.lExtendedCT_2, 0)
                                                            ElseIf oTarget.ObjTypeID <> ObjectType.eFacility Then
                                                                oTarget.iCombatTactics -= eiBehaviorPatterns.eEngagement_Evade
                                                                oTarget.iCombatTactics = oTarget.iCombatTactics Or eiBehaviorPatterns.eEngagement_Stand_Ground
                                                            End If
                                                        End If
                                                    ElseIf (oTarget.iCombatTactics And eiBehaviorPatterns.eEngagement_Dock_With_Target) <> 0 Then
                                                        'Dock with the Primary Target... function call?
                                                        If (oTarget.CurrentStatus And elUnitStatus.eUnitMoving) = 0 Then
                                                            goMsgSys.SendAIDockRequestToPathfinding(oTarget, oTarget.lExtendedCT_1, CShort(oTarget.lExtendedCT_2))
                                                            oTarget.iCombatTactics = oTarget.iCombatTactics Xor eiBehaviorPatterns.eEngagement_Dock_With_Target
                                                            oTarget.iCombatTactics = oTarget.iCombatTactics Or eiBehaviorPatterns.eEngagement_Stand_Ground
                                                        End If
                                                    ElseIf (oTarget.iCombatTactics And eiBehaviorPatterns.eEngagement_Hold_Fire) = 0 Then
                                                        'If oTargetDef.MaxSpeed <> 0 Then
                                                        If oTarget.MaxSpeed <> 0 Then
                                                            'Ok, it needs to engage, it will engage the guy shooting at it...
                                                            If (oTarget.CurrentStatus And elUnitStatus.eTargetingByPlayer) = 0 AndAlso oTarget.bAIMoveRequestPending = False Then
                                                                oTarget.lPrimaryTargetServerIdx = .ServerIndex
                                                                oTarget.CurrentStatus = oTarget.CurrentStatus Or elUnitStatus.eUnitEngaged
                                                                AddEntityInCombat(oTarget.ObjectID, oTarget.ObjTypeID, oTarget.ServerIndex)
                                                                oTarget.lNextFireWpnEvent = 0

                                                                'MSC - 12/17/07 - Put this in to split up weapons fire a bit better
                                                                If oTarget.lWpnNextFireCycle Is Nothing = False Then
                                                                    For Z As Int32 = 0 To oTarget.lWpnNextFireCycle.GetUpperBound(0)
                                                                        If oTarget.lWpnNextFireCycle(Z) < glCurrentCycle Then
                                                                            oTarget.lWpnNextFireCycle(Z) = glCurrentCycle + CInt(Rnd() * 10) '+ otargetdef.WeaponDefs(Z).ROF_Delay 
                                                                        End If
                                                                    Next Z
                                                                End If

                                                                'Move the unit into range if its combat tactics permit it
                                                                If (oTarget.iCombatTactics And eiBehaviorPatterns.eEngagement_Stand_Ground) = 0 AndAlso oTargetDef Is Nothing = False Then
                                                                    Dim lPrefRng As Int32 = Math.Max(GetPreferredRange(oTarget), oTargetDef.lOptRadarRange)
                                                                    If lPrefRng < lRange Then
                                                                        SetPreferredRangeDest(oTarget, oEntity, lPrefRng)
                                                                        'Dim fAngle As Single = LineAngleDegrees(.LocX, .LocZ, oTarget.LocX, oTarget.LocZ)
                                                                        'Dim lTX As Int32 = .LocX + (lPrefRng * gl_FINAL_GRID_SQUARE_SIZE)
                                                                        'Dim lTZ As Int32 = .LocZ
                                                                        'RotatePoint(.LocX, .LocZ, lTX, lTZ, fAngle)
                                                                        'goMsgSys.SendAIMoveRequestToPathfinding(oTarget, lTX, lTZ, 1S)
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If

                                            End If
                                        End If
                                        'Else
                                        '	.lNextFireWpnEvent = Math.Min(.lNextFireWpnEvent, .lWpnNextFireCycle(Y))
                                    End If
                                    .lNextFireWpnEvent = Math.Min(.lNextFireWpnEvent, .lWpnNextFireCycle(Y))
                                Next Y
                            End If
                        ElseIf .lIncMissileCnt > 0 AndAlso (.CurrentStatus And elUnitStatus.eRadarOperational) <> 0 Then
                            'ok, I have an incoming missile...
                            .lLastEngagement = glCurrentCycle

                            'We need to check for point defense weapons on board...
                            If .lNextFireWpnEvent <= glCurrentCycle Then
                                .lNextFireWpnEvent = Int32.MaxValue
                                For Y = 0 To oCurDef.WeaponDefUB
                                    If .lWpnNextFireCycle(Y) <= glCurrentCycle Then
                                        If (.CurrentStatus And oCurDef.WeaponDefs(Y).lEntityStatusGroup) <> 0 Then
                                            oWeapon = oCurDef.WeaponDefs(Y)
                                            If oWeapon.WpnGroup <> WeaponGroupType.PointDefenseWeaponGroup Then Continue For

                                            'Determine what side we are firing with...
                                            Dim ySideFiring As Byte = oWeapon.ArcID

                                            'We only care about missiles and mines in this case... all other targets *should* not exist at this time because we are unengaged
                                            Dim lMissileTargetRange As Int32 = 0
                                            Dim oMissileTarget As MissileMgr.Missile = .GetMissileTarget(ySideFiring, oWeapon.lRange, lMissileTargetRange, oCurDef)
                                            If oMissileTarget Is Nothing = False Then
                                                ProcessPDSToMissileShot(oEntity, oCurDef, oMissileTarget, oWeapon, lMissileTargetRange)

                                                'regardless of hit or miss, update our fire cycle...
                                                .lWpnNextFireCycle(Y) = glCurrentCycle + oWeapon.ROF_Delay
                                                .lNextFireWpnEvent = Math.Min(.lNextFireWpnEvent, .lWpnNextFireCycle(Y))

                                                Continue For
                                            End If

                                            '    'mines
                                            '    'TODO: Implement mine point defense...

                                        End If

                                    End If
                                    .lNextFireWpnEvent = Math.Min(.lNextFireWpnEvent, .lWpnNextFireCycle(Y))
                                Next Y
                            End If


                        End If

                    Else
                        If .lOwnerID <> gl_HARDCODE_PIRATE_PLAYER_ID AndAlso lKilledByID <> gl_HARDCODE_PIRATE_PLAYER_ID AndAlso lKilledByID <> -1 Then
                            Dim oWT As WarTracker = .ParentEnvir.GetOrAddWarTracker(.lOwnerID)
                            If oWT Is Nothing = False Then
                                oWT.AddEntityKilled(oCurDef, lKilledByID)
                            End If
                            oWT = Nothing
                            oWT = .ParentEnvir.GetOrAddWarTracker(lKilledByID)
                            If oWT Is Nothing = False Then
                                oWT.AddKill(.lOwnerID, .ObjTypeID)
                            End If
                            oWT = Nothing
                        End If
                        .ParentEnvir.RemoveEntity(X, RemovalType.eObjectDestroyed, True, True, lKilledByID)

                        ''the unit is destroyed...
                        'Dim yRemoveMsg() As Byte = goMsgSys.CreateRemoveObjectMsg(.ObjectID, .ObjTypeID, RemovalType.eObjectDestroyed, .LocX, .LocZ)
                        ''broadcast to clients
                        'If .ParentEnvir.lPlayersInEnvirCnt > 0 Then goMsgSys.BroadcastToEnvironmentClients(yRemoveMsg, .ParentEnvir)
                        ''send to primary
                        'goMsgSys.SendToPrimary(yRemoveMsg)

                        ''Now, remove the object from the server...
                        'glEntityIdx(X) = -1
                        ''remove it from the environment
                        '.ParentEnvir.oGrid(.lGridIndex).RemoveEntity(.lGridEntityIdx)

                    End If

                End With

                If bUnitDestroyed = True Then goEntity(X) = Nothing
            End If
        Next X

    End Sub

    Private Sub ProcessPDSToMissileShot(ByRef oEntity As Epica_Entity, ByRef oCurDef As Epica_Entity_Def, ByRef oMissileTarget As MissileMgr.Missile, ByVal oWeapon As WeaponDef, ByVal lMissileTargetRange As Int32)
        With oEntity
            Dim lBaseToHit As Int32 = GetBaseToHit(oEntity, oCurDef, oMissileTarget.VelX, oMissileTarget.VelZ, 0, oWeapon, lMissileTargetRange, 1, 3)
            Dim lRoll As Int32 = CInt(Rnd() * 100) + 1

            Dim bHitMissile As Boolean = lRoll < lBaseToHit
            'and send our missed fire weapon message
            If .ParentEnvir.lPlayersInEnvirCnt > 0 Then
                If .ParentEnvir.ObjTypeID = ObjectType.eSolarSystem Then
                    Dim lCameraPosX As Int32 = (.LocX \ 40000) + 125
                    Dim lCameraPosZ As Int32 = (.LocZ \ 40000) + 125
                    If bHitMissile = True Then
                        goMsgSys.BroadcastToEnvironmentClients_Filter(goMsgSys.CreatePointDefenseMessage(GlobalMessageCode.ePDS_AtMissileHit, oEntity, oMissileTarget.lMissileID, 0, oWeapon.yWeaponType, 0), .ParentEnvir, lCameraPosX, lCameraPosZ)
                    Else
                        goMsgSys.BroadcastToEnvironmentClients_Filter(goMsgSys.CreatePointDefenseMessage(GlobalMessageCode.ePDS_AtMissileMiss, oEntity, oMissileTarget.lMissileID, 0, oWeapon.yWeaponType, 0), .ParentEnvir, lCameraPosX, lCameraPosZ)
                    End If
                Else
                    If bHitMissile = True Then
                        goMsgSys.BroadcastToEnvironmentClients(goMsgSys.CreatePointDefenseMessage(GlobalMessageCode.ePDS_AtMissileHit, oEntity, oMissileTarget.lMissileID, 0, oWeapon.yWeaponType, 0), .ParentEnvir)
                    Else
                        goMsgSys.BroadcastToEnvironmentClients(goMsgSys.CreatePointDefenseMessage(GlobalMessageCode.ePDS_AtMissileMiss, oEntity, oMissileTarget.lMissileID, 0, oWeapon.yWeaponType, 0), .ParentEnvir)
                    End If
                End If
            End If

            If bHitMissile = True Then
                'we hit the missile... determine damage... on a critical hit, the missile is automatically destroyed
                If lRoll <= CInt(.Owner.CriticalHitMax) + CInt(oCurDef.CriticalHitModifier) Then
                    'destroy the missile
                    goMissileMgr.DetonateMissile(oMissileTarget, oMissileTarget.lMissileIdx, MissileMgr.eyDetonateType.BreakApart)
                Else
                    Dim lTotalDmg As Int32 = 0
                    With oWeapon
                        If .PiercingMaxDmg > 0 Then lTotalDmg += (CInt(Rnd() * .PiercingMaxMinRange) + .PiercingMinDmg) '+ .DamageMod
                        If .ImpactMaxDmg > 0 Then lTotalDmg += (CInt(Rnd() * .ImpactMaxMinRange) + .ImpactMinDmg) '+ .DamageMod
                        If .BeamMaxDmg > 0 Then lTotalDmg += (CInt(Rnd() * .BeamMaxMinRange) + .BeamMinDmg) '+ .DamageMod
                        If .ECMMaxDmg > 0 Then lTotalDmg += (CInt(Rnd() * .ECMMaxMinRange) + .ECMMinDmg) '+ .DamageMod
                        If .FlameMaxDmg > 0 Then lTotalDmg += (CInt(Rnd() * .FlameMaxMinRange) + .FlameMinDmg) '+ .DamageMod
                        If .ChemicalMaxDmg > 0 Then lTotalDmg += (CInt(Rnd() * .ChemicalMaxMinRange) + .ChemicalMinDmg) '+ .DamageMod
                    End With
                    If .DamageMod <> 0 Then
                        Dim fMult As Single = 1.0F + (.DamageMod / 100.0F)
                        lTotalDmg = CInt(lTotalDmg * fMult)
                    End If
                    oMissileTarget.lHitpoints -= lTotalDmg
                    If oMissileTarget.lHitpoints < 1 Then
                        goMissileMgr.DetonateMissile(oMissileTarget, oMissileTarget.lMissileIdx, MissileMgr.eyDetonateType.BreakApart)
                    End If
                End If
            End If
        End With
    End Sub

    Private mbUseWpnCalcedAcc As Boolean = True
    Private Function GetBaseToHit(ByVal oAttacker As Epica_Entity, ByVal oCurDef As Epica_Entity_Def, ByVal fTargetVelX As Single, ByVal fTargetVelZ As Single, ByVal lTargetHullSizeTargetMod As Int32, ByVal oWeapon As WeaponDef, ByVal lRange As Int32, ByVal yPointDefenseType As Byte, ByVal lTargetHullSize As Int32) As Int32
        Dim lBaseToHit As Int32 = 0
        With oAttacker
            If (.CurrentStatus And elUnitStatus.eRadarOperational) <> 0 Then

                Dim lAcc As Int32
                If mbUseWpnCalcedAcc = True Then
                    If yPointDefenseType <> 0 Then
                        lAcc = oWeapon.lPD_AccVal 'oCurDef.ScanResolution
                    Else : lAcc = oWeapon.lNrm_AccVal 'oCurDef.Weapon_Acc
                    End If
                Else
                    If yPointDefenseType <> 0 Then
                        lAcc = oCurDef.ScanResolution
                    Else : lAcc = oCurDef.Weapon_Acc
                    End If
                End If

                Dim lTemp As Int32 = oCurDef.lOptRadarRange
                If oCurDef.ModelID = 16 OrElse oCurDef.ModelID = 138 Then
                    lTemp = Math.Min(255, lTemp * 2)
                End If

                If lTemp + .lJamRangeMod >= lRange Then
                    lBaseToHit = lAcc
                ElseIf (lTemp + .lJamRangeMod) <> 0 Then
                    'lBaseToHit = oCurDef.Weapon_Acc - CInt(((lRange * oCurDef.fOneOverOptRadarRange) - 1) * oCurDef.Weapon_Acc)
                    lBaseToHit = lAcc - CInt(((lRange * oCurDef.fOneOverOptRadarRange) - 1) * lAcc)
                Else : lBaseToHit = 5
                End If

                Dim fTargetVel As Single = Math.Abs(fTargetVelX) + Math.Abs(fTargetVelZ)
                'lBaseToHit += CInt((fTargetVel * -2.0F) + (.TotalVelocity * -0.2F))    '(CInt(oTarget.TotalVelocity * -2) + CInt(.TotalVelocity / -5))
                If (.CurrentStatus And elUnitStatus.eUnitMoving) <> 0 Then
                    'abs(theirs - mine)
                    Dim lVelMod As Int32 = CInt(Math.Abs(fTargetVelX - .VelX) + Math.Abs(fTargetVelZ - .VelZ))
                    lBaseToHit += -lVelMod
                Else
                    Dim lVelMod As Int32 = CInt(Math.Abs(fTargetVelX) + Math.Abs(fTargetVelZ))
                    lBaseToHit += -lVelMod
                End If

                'lBaseToHit += oTargetDef.HullSizeTargetMod
                lBaseToHit += Math.Min(lTargetHullSizeTargetMod, lAcc)

                'Now, add to it the unit's to hit mod
                lBaseToHit += .ToHitMod

                lBaseToHit += .lJamWpnAccMod

                If yPointDefenseType = 0 Then
                    If oWeapon.blOverallMaxDmg > lTargetHullSize * 3 Then
                        Dim fTemp As Single = -(oWeapon.blOverallMaxDmg / (3.0F * lTargetHullSize))
                        If fTemp < -50000000 Then
                            lBaseToHit = -50000000
                        Else
                            lBaseToHit += CInt(fTemp)
                        End If
                    End If
                End If

                'lBaseToHit += CInt((1 - (lRange / oWeapon.MaxRangeAccuracy)) * 50) 

                Dim fRangeMod As Single = 0
                Dim lWpnRng As Int32 = oWeapon.lRange
                If yPointDefenseType = 2 Then lWpnRng *= 2

                If (lWpnRng >= lRange) Then
                    'fRangeMod = (1.0F - CSng(lRange * oWeapon.fOneOverRange)) * 50.0F
                ElseIf yPointDefenseType = 0 OrElse (lRange - fTargetVel) > 0 Then
                    'lBaseToHit += CInt(lBaseToHit * (-((lRange - oWeapon.lRange) / (oCurDef.lOptRadarRange - oWeapon.lRange))))
                    If oCurDef.lOptRadarRange = 0 Then
                        fRangeMod = -lBaseToHit * 10
                    Else
                        'fRangeMod = (lBaseToHit * CSng(((lRange - oWeapon.lRange) / (oCurDef.lOptRadarRange - oWeapon.lRange))))
                        fRangeMod = (lBaseToHit * CSng(((lRange - oWeapon.lRange) * oWeapon.fOneOverOptMinusRange)))
                        fRangeMod *= fRangeMod
                        fRangeMod = -Math.Abs(fRangeMod)
                    End If
                End If
                lBaseToHit += CInt(fRangeMod)

                If lBaseToHit < 1 Then lBaseToHit = 1
                If lBaseToHit > 98 Then lBaseToHit = 98

            Else
                'its called blind fire... and because of the extreme circumstance...
                '  they only receive a 5% chance to hit the target
                lBaseToHit = 5
            End If
        End With
        Return lBaseToHit

    End Function

    Private Function GetPrimaryTargetScore(ByVal oUnit As Epica_Entity, ByVal oTarget As Epica_Entity, ByVal lDist As Int32, ByVal iTargetTacticalAnalysis As Int16) As Int32
        Dim lBitMatch As Int32

        'TODO: Move this code into the area where GetPrimaryTargetScore is called, only if your score is higher do you want to call this
        If oUnit.ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
            'since both objects share the same environment...

            If oUnit.lOwnerID = gl_HARDCODE_PIRATE_PLAYER_ID Then
                If oTarget.yProductionType = ProductionType.eCommandCenterSpecial Then Return Int32.MaxValue
            End If

            With CType(oUnit.ParentEnvir.oGeoObject, Planet)
                If oUnit.LocY = 0 Then oUnit.LocY = .GetHeightAtPoint(oUnit.LocX, oUnit.LocZ, True)
                If oTarget.LocY = 0 Then oTarget.LocY = .GetHeightAtPoint(oTarget.LocX, oTarget.LocZ, True)

                'TODO: 16 here is hardcoded, I will want to have a lookup array for the Model ID to get this value
                'TODO: a bit of a hack.... but what do we do?
                If oTarget.ObjTypeID = ObjectType.eUnit AndAlso oUnit.ObjTypeID = ObjectType.eUnit Then

                    'TODO: Unremark this if LOS is not what is causing the no-shoots

                    'If .HasLineOfSight(oUnit.LocX, oUnit.LocY, oUnit.LocZ, oTarget.LocX, oTarget.LocY, oTarget.LocZ, 16) = False Then
                    '    'Ok, doesn't have LOS... what now? Return 0 for now but we may want to make the AI smarter...
                    '    Return 0
                    'End If
                End If
            End With
        End If

        'lBitMatch = GetMatchingBitCount(oUnit.iTargetingTactics, oTarget.TacticalAnalysis) * 1600
        lBitMatch = GetMatchingBitCount(oUnit.iTargetingTactics, iTargetTacticalAnalysis) * 1600
        lBitMatch += CInt((256 - lDist)) ' * 6.25F)

        Return lBitMatch
    End Function

    Public Function GetPreferredRange(ByVal oEntity As Epica_Entity) As Int32
        Dim oTmpDef As Epica_Entity_Def = Nothing
        Dim X As Int32

        Dim lRange As Int32 = Int32.MaxValue

        Dim bIncludeSecondary As Boolean = False

        'Get our entity definition and set a bool for whether to include secondary weapons for range calcs
        With oEntity
            If .lEntityDefServerIndex <> -1 Then
                If glEntityDefIdx(.lEntityDefServerIndex) <> -1 Then
                    oTmpDef = goEntityDefs(.lEntityDefServerIndex)
                End If
            End If

            bIncludeSecondary = (.iCombatTactics And eiBehaviorPatterns.eTactics_Maximize_Damage) <> 0
        End With

        If oTmpDef Is Nothing = False Then
            With oTmpDef
                lRange = .lOptRadarRange + oEntity.lJamRangeMod - 2
                For X = 0 To .WeaponDefUB
                    If .WeaponDefs(X).WpnGroup = WeaponGroupType.PrimaryWeaponGroup Then
                        If (oEntity.CurrentStatus And .WeaponDefs(X).lEntityStatusGroup) <> 0 Then
                            If .WeaponDefs(X).lRange < lRange Then lRange = .WeaponDefs(X).lRange
                        End If
                    ElseIf bIncludeSecondary = True AndAlso .WeaponDefs(X).WpnGroup = WeaponGroupType.SecondaryWeaponGroup Then
                        If (oEntity.CurrentStatus And .WeaponDefs(X).lEntityStatusGroup) <> 0 Then
                            If .WeaponDefs(X).lRange < lRange Then lRange = .WeaponDefs(X).lRange
                        End If
                    End If
                Next X
            End With
        End If

        Return Math.Max(1, lRange)

    End Function

    Public Function GetPreferredSideFacing(ByVal oEntity As Epica_Entity) As Byte
        'returns the facing number indicating which side the unit prefers to face a target...
        Dim X As Int32
        Dim lTotalHP As Int32
        Dim lSide As Int32
        Dim lScore As Int32
        Dim lTemp As Int32

        Dim oTmpDef As Epica_Entity_Def

        'If oEntity.ObjectID = 5056 Then Stop 

        With oEntity
            'ok, acquire our def
            If .lEntityDefServerIndex <> -1 Then
                If glEntityDefIdx(.lEntityDefServerIndex) <> -1 Then
                    oTmpDef = goEntityDefs(.lEntityDefServerIndex)

                    If (.iCombatTactics And eiBehaviorPatterns.eTactics_Maximize_Damage) <> 0 Then
                        'Maximize damage to target
                        'Side Facing is equal to the side with the OPERATIONAL weaponry that can deal the most damage
                        lScore = Int32.MinValue
                        Dim lSideFirePower(3) As Int32
                        For X = 0 To oTmpDef.WeaponDefUB
                            If oTmpDef.WeaponDefs(X).WpnGroup = WeaponGroupType.PrimaryWeaponGroup Then
                                If (oTmpDef.WeaponDefs(X).lEntityStatusGroup And .CurrentStatus) <> 0 Then
                                    If oTmpDef.WeaponDefs(X).ArcID <> UnitArcs.eAllArcs Then
                                        lSideFirePower(oTmpDef.WeaponDefs(X).ArcID) += oTmpDef.WeaponDefs(X).lFirePowerRating
                                    End If
                                End If
                            End If
                        Next X

                        'Now, what's our highest
                        For X = 0 To 3
                            If lSideFirePower(X) > lScore Then
                                lScore = lSideFirePower(X)
                                lSide = X
                            End If
                        Next X
                    ElseIf (.iCombatTactics And eiBehaviorPatterns.eTactics_Minimize_Damage_To_Self) <> 0 Then
                        'Minimize damage to self
                        'Side Facing is equal to the side with the most amount of hitpoints
                        lScore = Int32.MinValue
                        If .ArmorHP(0) > lScore Then lScore = .ArmorHP(0) : lSide = 0
                        If .ArmorHP(1) > lScore Then lScore = .ArmorHP(1) : lSide = 1
                        If .ArmorHP(2) > lScore Then lScore = .ArmorHP(2) : lSide = 2
                        If .ArmorHP(3) > lScore Then lScore = .ArmorHP(3) : lSide = 3
                    Else    ' (oUnit.iCombatTactics And eiBehaviorPatterns.eTactics_Normal) Then
                        'Normal and everything else
                        'Side facing is special...
                        'lTotalHP = oUnit.Q1_HP + oUnit.Q2_HP + oUnit.Q3_HP + oUnit.Q4_HP
                        lTotalHP = .ArmorHP(0) + .ArmorHP(1) + .ArmorHP(2) + .ArmorHP(3)

                        'If lTotalHP = 0 Then Return 0

                        Dim lSideFirePower(3) As Int32
                        For X = 0 To oTmpDef.WeaponDefUB
                            If oTmpDef.WeaponDefs(X).WpnGroup = WeaponGroupType.PrimaryWeaponGroup Then
                                If (oTmpDef.WeaponDefs(X).lEntityStatusGroup And .CurrentStatus) <> 0 Then
                                    If oTmpDef.WeaponDefs(X).ArcID <> UnitArcs.eAllArcs Then
                                        lSideFirePower(oTmpDef.WeaponDefs(X).ArcID) += oTmpDef.WeaponDefs(X).lFirePowerRating
                                    End If
                                End If
                            End If
                        Next X
                        If lTotalHP <> 0 Then
                            lTemp = CInt(lSideFirePower(0) * (.ArmorHP(0) / lTotalHP))
                            If lTemp > lScore Then lScore = lTemp : lSide = 0
                            lTemp = CInt(lSideFirePower(1) * (.ArmorHP(1) / lTotalHP))
                            If lTemp > lScore Then lScore = lTemp : lSide = 1
                            lTemp = CInt(lSideFirePower(2) * (.ArmorHP(2) / lTotalHP))
                            If lTemp > lScore Then lScore = lTemp : lSide = 2
                            lTemp = CInt(lSideFirePower(3) * (.ArmorHP(3) / lTotalHP))
                            If lTemp > lScore Then lScore = lTemp : lSide = 3
                        Else
                            lTemp = lSideFirePower(0)
                            If lTemp > lScore Then lScore = lTemp : lSide = 0
                            lTemp = lSideFirePower(1)
                            If lTemp > lScore Then lScore = lTemp : lSide = 1
                            lTemp = lSideFirePower(2)
                            If lTemp > lScore Then lScore = lTemp : lSide = 2
                            lTemp = lSideFirePower(3)
                            If lTemp > lScore Then lScore = lTemp : lSide = 3
                        End If

                    End If
                End If
            End If
        End With

        'Maneuver and Launch Children are handled elsewhere

        Return CByte(lSide)

    End Function

    Private Function GetMatchingBitCount(ByVal lVal1 As Int32, ByVal lVal2 As Int32) As Int32
        Dim lValTest As Int32
        Dim lRes As Int32

        lValTest = lVal1 And lVal2

        lRes = -1 * (CInt((lValTest And eiTacticalAttrs.eFighterClass) <> 0) + _
            CInt((lValTest And eiTacticalAttrs.eCargoShipClass) <> 0) + _
            CInt((lValTest And eiTacticalAttrs.eTroopClass) <> 0) + _
            CInt((lValTest And eiTacticalAttrs.eEscortClass) <> 0) + _
            CInt((lValTest And eiTacticalAttrs.eCapitalClass) <> 0) + _
            CInt((lValTest And eiTacticalAttrs.eArmedUnit) <> 0) + _
            CInt((lValTest And eiTacticalAttrs.eUnarmedUnit) <> 0) + _
            CInt((lValTest And eiTacticalAttrs.eFacility) <> 0))

        Return lRes
    End Function

    Public Sub ApplyAOEDamage(ByVal oEvent As WeaponEvent, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByRef oEnvir As Envir, ByVal lAttackerID As Int32)
        Dim lGrid As Int32
        Dim lLargeSector As Int32
        Dim lRelTinyX As Int32
        Dim lRelTinyZ As Int32
        Dim X As Int32
        Dim oTmpUnit As Epica_Entity
        Dim lTemp As Int32
        Dim lTemp2 As Int32

        Dim Y As Int32
        Dim lIdx As Int32

        Dim lBaseFacing As Int32

        'Determine our aggression engine data...
        Dim lGridIndex As Int32
        Dim lSmallSectorID As Int32
        Dim lTinyX As Int32
        Dim lTinyZ As Int32

        'Any changes to the aggression engine calculations need to be made here too
        lTemp = lLocZ + oEnvir.lHalfEnvirSize
        'lLargeSector = CInt(Math.Floor(lTemp / oEnvir.lGridSquareSize)) * oEnvir.lGridsPerRow
        lLargeSector = (lTemp \ oEnvir.lGridSquareSize) * oEnvir.lGridsPerRow
        'lTemp -= CInt((lLargeSector / oEnvir.lGridsPerRow) * oEnvir.lGridSquareSize)
        lTemp -= ((lLargeSector \ oEnvir.lGridsPerRow) * oEnvir.lGridSquareSize)
        'lSmallSectorID = CInt(Math.Floor(lTemp / gl_SMALL_GRID_SQUARE_SIZE)) * gl_SMALL_PER_ROW
        lSmallSectorID = ((lTemp \ gl_SMALL_GRID_SQUARE_SIZE) * gl_SMALL_PER_ROW)
        'lTemp -= CInt((lSmallSectorID / gl_SMALL_PER_ROW) * gl_SMALL_GRID_SQUARE_SIZE)
        lTemp -= ((lSmallSectorID \ gl_SMALL_PER_ROW) * gl_SMALL_GRID_SQUARE_SIZE)
        'lTinyZ = CInt(Math.Floor(lTemp / gl_FINAL_GRID_SQUARE_SIZE))
        lTinyZ = lTemp \ gl_FINAL_GRID_SQUARE_SIZE
        'lTemp2 = lLocX + oEnvir.lHalfEnvirSize
        lTemp2 = lLocX + oEnvir.lHalfEnvirSize
        'lTemp = CInt(Math.Floor(lTemp2 / oEnvir.lGridSquareSize))
        lTemp = lTemp2 \ oEnvir.lGridSquareSize
        lLargeSector += lTemp
        lTemp2 -= (lTemp * oEnvir.lGridSquareSize)
        lTemp = lTemp2 \ gl_SMALL_GRID_SQUARE_SIZE
        lSmallSectorID += lTemp
        lTemp2 -= (lTemp * gl_SMALL_GRID_SQUARE_SIZE)
        lTinyX = lTemp2 \ gl_FINAL_GRID_SQUARE_SIZE

        If lSmallSectorID < 0 OrElse lSmallSectorID > giRelativeSmall.GetUpperBound(1) Then Return

        lGridIndex = lLargeSector

        'Determine what adjust array to use (for map wrapping of planets)
        Dim lAdjust() As Int32 = oEnvir.lGridIdxAdjust
        If oEnvir.ObjTypeID = ObjectType.ePlanet Then
            Dim lModVal As Int32 = lGridIndex Mod oEnvir.lGridsPerRow
            If lModVal = 0 Then
                'I'm on the left edge, use the left edge adjust
                lAdjust = oEnvir.lLeftEdgeGridIdxAdjust
            ElseIf lModVal = oEnvir.lGridsPerRow - 1 Then
                'I'm on the right edge, use the right edge adjust
                lAdjust = oEnvir.lRightEdgeGridIdxAdjust
            End If
        End If

        If oEvent.AOERange = 0 Then Return
        Dim f1OverRange As Single = 1.0F / Math.Max(1, oEvent.AOERange)

        'Ok, go through an aggression test on this event
        For lGrid = 0 To lAdjust.GetUpperBound(0)
            'gonna reuse the llargesector here...
            'lLargeSector = lGridIndex + oEnvir.lGridIdxAdjust(lGrid)
            lLargeSector = lGridIndex + lAdjust(lGrid)

            If lLargeSector > -1 AndAlso lLargeSector < oEnvir.lGridUB + 1 Then
                'Ok, now, cycle thru the smaller list in the Grid object
                Dim oTmpGrid As EnvirGrid = oEnvir.oGrid(lLargeSector)
                If oTmpGrid Is Nothing Then Continue For
                For X = 0 To oTmpGrid.lEntityUB
                    lTemp = oTmpGrid.lEntities(X)
                    If lTemp <> -1 Then      'is the unit not nothing?
                        oTmpUnit = goEntity(lTemp)

                        If oTmpUnit Is Nothing Then
                            oTmpGrid.lEntities(X) = -1
                            Continue For
                        End If

                        If oTmpUnit.Owner Is Nothing = False AndAlso oEnvir.bEnvirAtColonyLimit = False AndAlso (oTmpUnit.Owner.lIronCurtainPlanetID = Int32.MinValue OrElse (oEnvir.ObjTypeID = ObjectType.ePlanet AndAlso oEnvir.ObjectID = oTmpUnit.Owner.lIronCurtainPlanetID)) Then Continue For
                        If oTmpUnit.lOwnerID <> lAttackerID AndAlso oTmpUnit.Owner.GetPlayerRelScore(lAttackerID, False, -1) > elRelTypes.eWar Then Continue For

                        'get our relative small
                        lTemp = giRelativeSmall(lGrid, lSmallSectorID, oTmpUnit.lSmallSectorID)
                        If lTemp <> -1 Then
                            lRelTinyX = glBaseRelTinyX(lTemp) + oTmpUnit.lTinyX + (gl_HALF_FINAL_PER_ROW - lTinyX)
                            If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                lRelTinyZ = glBaseRelTinyZ(lTemp) + oTmpUnit.lTinyZ + (gl_HALF_FINAL_PER_ROW - lTinyZ)
                                If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                    Dim lRangeToEpicenter As Int32 = Math.Max(0, (gyDistances(lRelTinyX, lRelTinyZ) - oTmpUnit.lModelRangeOffset))
                                    If lRangeToEpicenter <= oEvent.AOERange Then
                                        'Set our base facing value
                                        lBaseFacing = glFacing(lRelTinyX, lRelTinyZ)
                                        lBaseFacing -= 1800
                                        If lBaseFacing < 0 Then lBaseFacing += 3600

                                        'ok, find a place to add this event to
                                        lIdx = -1
                                        For Y = 0 To oTmpUnit.lWpnEventUB
                                            If oTmpUnit.yWpnEventUsed(Y) = 0 Then
                                                lIdx = Y
                                                'ensure that the event is cleared
                                                oTmpUnit.oWpnEvents(Y) = Nothing
                                                Exit For
                                            End If
                                        Next Y

                                        If lIdx = -1 Then
                                            oTmpUnit.lWpnEventUB += 1
                                            ReDim Preserve oTmpUnit.oWpnEvents(oTmpUnit.lWpnEventUB)
                                            ReDim Preserve oTmpUnit.yWpnEventUsed(oTmpUnit.lWpnEventUB)
                                            lIdx = oTmpUnit.lWpnEventUB
                                        End If

                                        oTmpUnit.yWpnEventUsed(lIdx) = 255
                                        oTmpUnit.oWpnEvents(lIdx) = New WeaponEvent()
                                        With oTmpUnit.oWpnEvents(lIdx)
                                            Dim fMult As Single = (CInt(oEvent.AOERange) - lRangeToEpicenter) * f1OverRange

                                            .AOERange = 0
                                            .yCritical = eyCriticalHitType.NoCriticalHit
                                            .lBeamDmg = CInt(oEvent.lBeamDmg * fMult)
                                            .lChemicalDmg = CInt(oEvent.lChemicalDmg * fMult)
                                            .lECMDmg = CInt(oEvent.lECMDmg * fMult)
                                            .lFlameDmg = CInt(oEvent.lFlameDmg * fMult)
                                            .lImpactDmg = CInt(oEvent.lImpactDmg * fMult)
                                            .lPierceDmg = CInt(oEvent.lPierceDmg * fMult)
                                            .lEventCycle = glCurrentCycle
                                            .lFromIdx = oEvent.lToIdx
                                            .lToIdx = oTmpUnit.ServerIndex
                                            .ySideHit = WhatSideDoIHit(lGridIndex, lSmallSectorID, lTinyX, lTinyZ, oTmpUnit)
                                            .FromPlayerID = oEvent.FromPlayerID
                                            .yAttackerType = oEvent.yAttackerType

                                            oTmpUnit.lNextWpnEvent = Math.Min(oTmpUnit.lNextWpnEvent, .lEventCycle)
                                        End With

                                    End If
                                End If
                            End If
                        End If
                    End If
                Next X
            End If
        Next lGrid

    End Sub

    Public Sub ApplyAOEDamage(ByVal oEvent As WeaponEvent, ByVal oSourceEntity As Epica_Entity)
        Dim lGrid As Int32
        Dim lLargeSector As Int32
        Dim lRelTinyX As Int32
        Dim lRelTinyZ As Int32
        Dim X As Int32
        Dim oTmpUnit As Epica_Entity
        Dim lTemp As Int32

        If oEvent.yAOEDmgType = 0 Then Return

        If mbRunAntiLightCode = True Then
            If oSourceEntity.lOwnerID = 20944 OrElse oSourceEntity.lOwnerID = 29216 OrElse oSourceEntity.lOwnerID = 20900 OrElse oSourceEntity.lOwnerID = 20842 OrElse oSourceEntity.lOwnerID = 28522 OrElse oSourceEntity.lOwnerID = 24396 Then
                oEvent.lPierceDmg \= 100
                oEvent.lImpactDmg \= 100
                oEvent.lBeamDmg \= 100
                oEvent.lECMDmg \= 100
                oEvent.lFlameDmg \= 100
                oEvent.lChemicalDmg \= 100
            End If
        End If
        oEvent.SetForAOEDmg()

        Dim Y As Int32
        Dim lIdx As Int32

        With oSourceEntity
            'Determine what adjust array to use (for map wrapping of planets)
            Dim lAdjust() As Int32 = .ParentEnvir.lGridIdxAdjust
            If .ParentEnvir.ObjTypeID = ObjectType.ePlanet Then
                Dim lModVal As Int32 = .lGridIndex Mod .ParentEnvir.lGridsPerRow
                If lModVal = 0 Then
                    'I'm on the left edge, use the left edge adjust
                    lAdjust = .ParentEnvir.lLeftEdgeGridIdxAdjust
                ElseIf lModVal = .ParentEnvir.lGridsPerRow - 1 Then
                    'I'm on the right edge, use the right edge adjust
                    lAdjust = .ParentEnvir.lRightEdgeGridIdxAdjust
                End If
            End If

            If oEvent.AOERange = 0 Then Return
            Dim f1OverRange As Single = 1.0F / Math.Max(oEvent.AOERange, 1)

            For lGrid = 0 To lAdjust.GetUpperBound(0)
                'gonna reuse the llargesector here...
                'lLargeSector = .lGridIndex + .ParentEnvir.lGridIdxAdjust(lGrid)
                lLargeSector = .lGridIndex + lAdjust(lGrid)

                If lLargeSector > -1 AndAlso lLargeSector < .ParentEnvir.lGridUB + 1 Then
                    'Ok, now, cycle thru the smaller list in the Grid object
                    Dim oTmpGrid As EnvirGrid = .ParentEnvir.oGrid(lLargeSector)
                    For X = 0 To oTmpGrid.lEntityUB
                        lTemp = oTmpGrid.lEntities(X)
                        If lTemp <> -1 Then      'is the unit not nothing?
                            oTmpUnit = goEntity(lTemp)

                            If oTmpUnit Is Nothing Then
                                oTmpGrid.lEntities(X) = -1
                                Continue For
                            End If

                            If oTmpUnit.Owner Is Nothing = False AndAlso .ParentEnvir.bEnvirAtColonyLimit = False AndAlso (oTmpUnit.Owner.lIronCurtainPlanetID = Int32.MinValue OrElse (.ParentEnvir.ObjTypeID = ObjectType.ePlanet AndAlso .ParentEnvir.ObjectID = oTmpUnit.Owner.lIronCurtainPlanetID)) Then Continue For
                            If oTmpUnit.lOwnerID <> oSourceEntity.lOwnerID AndAlso oTmpUnit.Owner.GetPlayerRelScore(oSourceEntity.lOwnerID, False, -1) > elRelTypes.eWar Then Continue For

                            'Exclude the source object 
                            If oTmpUnit.ObjectID <> .ObjectID OrElse oTmpUnit.ObjTypeID <> .ObjTypeID Then
                                lTemp = giRelativeSmall(lGrid, .lSmallSectorID, oTmpUnit.lSmallSectorID)
                                If lTemp <> -1 Then
                                    lRelTinyX = glBaseRelTinyX(lTemp) + oTmpUnit.lTinyX + (gl_HALF_FINAL_PER_ROW - .lTinyX)
                                    If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                        lRelTinyZ = glBaseRelTinyZ(lTemp) + oTmpUnit.lTinyZ + (gl_HALF_FINAL_PER_ROW - .lTinyZ)
                                        If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                            Dim lRangeToEpicenter As Int32 = Math.Max(0, (gyDistances(lRelTinyX, lRelTinyZ) - oTmpUnit.lModelRangeOffset))
                                            If lRangeToEpicenter <= oEvent.AOERange Then
                                                'ok, find a place to add this event to
                                                lIdx = -1
                                                For Y = 0 To oTmpUnit.lWpnEventUB
                                                    If oTmpUnit.yWpnEventUsed(Y) = 0 Then
                                                        lIdx = Y
                                                        'ensure that the event is cleared
                                                        oTmpUnit.oWpnEvents(Y) = Nothing
                                                        Exit For
                                                    End If
                                                Next Y

                                                If lIdx = -1 Then
                                                    oTmpUnit.lWpnEventUB += 1
                                                    ReDim Preserve oTmpUnit.oWpnEvents(oTmpUnit.lWpnEventUB)
                                                    ReDim Preserve oTmpUnit.yWpnEventUsed(oTmpUnit.lWpnEventUB)
                                                    lIdx = oTmpUnit.lWpnEventUB
                                                End If

                                                oTmpUnit.yWpnEventUsed(lIdx) = 255
                                                oTmpUnit.oWpnEvents(lIdx) = New WeaponEvent()
                                                With oTmpUnit.oWpnEvents(lIdx)

                                                    Dim fMult As Single = (CInt(oEvent.AOERange) - lRangeToEpicenter) * f1OverRange

                                                    .AOERange = 0
                                                    .yCritical = eyCriticalHitType.NoCriticalHit
                                                    .lBeamDmg = CInt(oEvent.lBeamDmg * fMult)
                                                    .lChemicalDmg = CInt(oEvent.lChemicalDmg * fMult)
                                                    .lECMDmg = CInt(oEvent.lECMDmg * fMult)
                                                    .lFlameDmg = CInt(oEvent.lFlameDmg * fMult)
                                                    .lImpactDmg = CInt(oEvent.lImpactDmg * fMult)
                                                    .lPierceDmg = CInt(oEvent.lPierceDmg * fMult)
                                                    .lEventCycle = glCurrentCycle
                                                    .lFromIdx = oEvent.lToIdx
                                                    .lToIdx = oTmpUnit.ServerIndex
                                                    .ySideHit = WhatSideDoIHit(oSourceEntity, oTmpUnit)
                                                    .FromPlayerID = oEvent.FromPlayerID
                                                    .yAttackerType = oEvent.yAttackerType

                                                    oTmpUnit.lNextWpnEvent = Math.Min(oTmpUnit.lNextWpnEvent, .lEventCycle)
                                                End With

                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If      'if unit is nothing
                    Next X
                End If
            Next lGrid
        End With

    End Sub

    Public Sub HandleBombardment()
        Dim X As Int32
        Try
            Dim lUB As Int32 = -1
            If gyBombRequestUsed Is Nothing = False AndAlso goBombRequests Is Nothing = False Then lUB = Math.Min(Math.Min(glBombRequestUB, gyBombRequestUsed.GetUpperBound(0)), goBombRequests.GetUpperBound(0))
            For X = 0 To lUB
                If gyBombRequestUsed(X) > 0 Then
                    goBombRequests(X).HandleBombardmentRequest()
                End If
            Next X
        Catch
        End Try
    End Sub

    Private Sub SetNextManeuverPoint(ByVal oEntity As Epica_Entity, ByVal lModelOffset As Int32)
        Dim fTempVelX As Single
        Dim fTempVelZ As Single

        Const ml_MANEUVER_CYCLE_LIMIT As Int32 = 120        '4 seconds

        'set up our target x and target z
        With oEntity

            'now, just for safety
            If (.CurrentStatus And elUnitStatus.eEngineOperational) = 0 Then 'OrElse (.CurrentStatus And elUnitStatus.eFuelBayOperational) = 0 Then
                If (.iCombatTactics And eiBehaviorPatterns.eTactics_Maneuver) <> 0 Then
                    'The unit can no longer perform maneuvers
                    .iCombatTactics -= eiBehaviorPatterns.eTactics_Maneuver
                    Return
                End If
            End If

            'If .lLastManCalcCycle <> Int32.MinValue Then
            '	'we need to check the cycle difference
            '	If glCurrentCycle - .lLastManCalcCycle < ml_MANEUVER_CYCLE_LIMIT Then
            '		.yCalcCycleStrikes += CByte(1)
            '		If .yCalcCycleStrikes >= 200 Then
            '			.yCalcCycleStrikes = 0
            '			If (.iCombatTactics And eiBehaviorPatterns.eTactics_Maneuver) <> 0 Then
            '				'The unit can no longer perform maneuvers
            '				.iCombatTactics -= eiBehaviorPatterns.eTactics_Maneuver
            '				Return
            '			End If
            '		End If
            '		Return
            '	End If
            'End If
            If .lLastManCalcCycle <> Int32.MinValue AndAlso glCurrentCycle - .lLastManCalcCycle < ml_MANEUVER_CYCLE_LIMIT AndAlso .bDecelerating = False Then Return
            .lLastManCalcCycle = glCurrentCycle

            Dim TargetX As Single = goEntity(.lPrimaryTargetServerIdx).LocX
            Dim TargetZ As Single = goEntity(.lPrimaryTargetServerIdx).LocZ

            '=============== NEW METHOD ================
            'ok, determine if we are a broadsider (left/right strong) or a jouster (front/back strong)
            Dim yPrefFace As Byte = GetPreferredSideFacing(oEntity)
            Dim lFinalDestX As Int32
            Dim lFinalDestZ As Int32
            Dim lMaxRange As Int32 = GetPreferredRange(oEntity) * gl_FINAL_GRID_SQUARE_SIZE * 2     '7/24/08 - added this for a possible manuever fix
            lMaxRange += (lModelOffset * gl_FINAL_GRID_SQUARE_SIZE)
            If yPrefFace = UnitArcs.eForwardArc OrElse yPrefFace = UnitArcs.eBackArc Then
                'jouster - a jouster toggles between close and far, close and far... therefore, their myManPointNum can be 0 or 1
                If .myManPointNum = 0 Then
                    .myManPointNum = 1
                    'Joust the target, move to the target's location
                    lFinalDestX = 0 'CInt(TargetX)
                    lFinalDestZ = 0 'CInt(TargetZ)
                Else
                    .myManPointNum = 0
                    'Unjoust the target, move away from the target, we move away in the appropriate direction and distance
                    lFinalDestX = lMaxRange
                    lFinalDestZ = 0

                    Dim fAngleTemp As Single = ((goEntity(.lPrimaryTargetServerIdx).LocAngle + .iManApproach) - 1800)
                    fAngleTemp += (Rnd() * 900) - 450
                    fAngleTemp *= 0.1F '/= 10.0F
                    RotatePoint(0, 0, lFinalDestX, lFinalDestZ, fAngleTemp)
                End If
            Else
                'assume broadsider - a broadsider moves in a circular motion around the target keeping one side (left or right) always facing the target
                'therefore, their myManPointNum cycles from 0 to 1 to 2 to 3 or from 0 to 3 to 2 to 1 depending on their preference
                If yPrefFace = UnitArcs.eLeftArc Then
                    '0 1 2 3
                    Dim lTemp As Int32 = CInt(.myManPointNum) + 1
                    If lTemp > 3 Then lTemp = 0
                    .myManPointNum = CByte(lTemp)
                Else
                    '0 3 2 1
                    Dim lTemp As Int32 = CInt(.myManPointNum) - 1
                    If lTemp < 0 Then lTemp = 3
                    .myManPointNum = CByte(lTemp)
                End If

                'Now, based on our point determines where we are...
                If .myManPointNum = 0 Then
                    lFinalDestX = lMaxRange
                    lFinalDestZ = 0
                ElseIf .myManPointNum = 1 Then
                    lFinalDestX = 0
                    lFinalDestZ = lMaxRange
                ElseIf .myManPointNum = 2 Then
                    lFinalDestX = -lMaxRange
                    lFinalDestZ = 0
                Else
                    lFinalDestX = 0
                    lFinalDestZ = -lMaxRange
                End If
                Dim fAngleTemp As Single = ((goEntity(.lPrimaryTargetServerIdx).LocAngle + .iManApproach) - 1800) * 0.1F '/ 10.0F
                RotatePoint(0, 0, lFinalDestX, lFinalDestZ, fAngleTemp)
            End If

            'Now, we have to account for the target being in motion

            'cycledist = cycles at max speed to reach target
            Dim fCycleDist As Single = Math.Abs(goEntity(.lPrimaryTargetServerIdx).LocX - .LocX) + Math.Abs(goEntity(.lPrimaryTargetServerIdx).LocZ - .LocZ)
            fCycleDist /= (CSng(.MaxSpeed) * 0.5F) ' / 2.0F)

            fTempVelX = lFinalDestX
            fTempVelZ = lFinalDestZ


            fTempVelX += goEntity(.lPrimaryTargetServerIdx).LocX + (goEntity(.lPrimaryTargetServerIdx).VelX * fCycleDist)
            fTempVelZ += goEntity(.lPrimaryTargetServerIdx).LocZ + (goEntity(.lPrimaryTargetServerIdx).VelZ * fCycleDist)

            If fTempVelX > Int32.MinValue AndAlso fTempVelX < Int32.MaxValue Then
                If fTempVelZ > Int32.MinValue AndAlso fTempVelZ < Int32.MaxValue Then
                    .DestX = CInt(fTempVelX)
                    .DestZ = CInt(fTempVelZ)
                    .DestAngle = CShort(LineAngleDegrees(.LocX, .LocZ, .DestX, .DestZ) * 10)
                Else
                    Return
                End If
            Else
                Return
            End If


            '=================== END OF NEW METHOD ================

            '=================== OLD METHOD ==================
            'Dim lCyclesToStop As Int32
            'If .myManPointNum = 0 Then
            '	.myManPointNum = 1
            '	fTempVelX = goEntityDefs(goEntity(.lPrimaryTargetServerIdx).lEntityDefServerIndex).lManeuverOffsetX_Min
            '	fTempVelZ = goEntityDefs(goEntity(.lPrimaryTargetServerIdx).lEntityDefServerIndex).lManeuverOffsetZ_Min
            'ElseIf .myManPointNum = 1 Then
            '	.myManPointNum = 2
            '	fTempVelX = goEntityDefs(goEntity(.lPrimaryTargetServerIdx).lEntityDefServerIndex).lManeuverOffsetX_Max
            '	fTempVelZ = -(goEntityDefs(goEntity(.lPrimaryTargetServerIdx).lEntityDefServerIndex).lManeuverOffsetZ_Max)
            'ElseIf .myManPointNum = 2 Then
            '	.myManPointNum = 3
            '	fTempVelX = goEntityDefs(goEntity(.lPrimaryTargetServerIdx).lEntityDefServerIndex).lManeuverOffsetX_Min
            '	fTempVelZ = goEntityDefs(goEntity(.lPrimaryTargetServerIdx).lEntityDefServerIndex).lManeuverOffsetZ_Min
            'Else
            '	.myManPointNum = 0
            '	fTempVelX = goEntityDefs(goEntity(.lPrimaryTargetServerIdx).lEntityDefServerIndex).lManeuverOffsetX_Max
            '	fTempVelZ = goEntityDefs(goEntity(.lPrimaryTargetServerIdx).lEntityDefServerIndex).lManeuverOffsetZ_Max
            'End If

            ''Offset the value by my Radaroffset for the model
            'fTempVelX += (lModelOffset * gl_FINAL_GRID_SQUARE_SIZE)
            'fTempVelZ += (lModelOffset * gl_FINAL_GRID_SQUARE_SIZE)

            ''Dim fOptRadar As Single = CSng((goEntityDefs(.lEntityDefServerIndex).lOptRadarRange + (goEntityDefs(goEntity(.lPrimaryTargetServerIdx).lEntityDefServerIndex).lModelRangeOffset)) * gl_FINAL_GRID_SQUARE_SIZE)
            'Dim fOptRadar As Single = CSng(goEntityDefs(.lEntityDefServerIndex).lOptRadarRange * gl_FINAL_GRID_SQUARE_SIZE)

            'If fTempVelX < 0 Then
            '	fTempVelX = Math.Max(fTempVelX, -fOptRadar)
            'Else : fTempVelX = Math.Min(fTempVelX, fOptRadar)
            'End If
            'If fTempVelZ < 0 Then
            '	fTempVelZ = Math.Max(fTempVelZ, -fOptRadar)
            'Else : fTempVelZ = Math.Min(fTempVelZ, fOptRadar)
            'End If

            'RotatePoint(0, 0, CInt(fTempVelX), CInt(fTempVelZ), ((goEntity(.lPrimaryTargetServerIdx).LocAngle + .iManApproach) - 1800) / 10.0F)

            'lCyclesToStop = CInt((Math.Abs((fTempVelX + TargetX) - .LocX) + Math.Abs((fTempVelZ + TargetZ) - .LocZ)) / (.TotalVelocity + 1))
            '.DestX = CInt(fTempVelX + TargetX + (goEntity(.lPrimaryTargetServerIdx).VelX * lCyclesToStop))
            '.DestZ = CInt(fTempVelZ + TargetZ + (goEntity(.lPrimaryTargetServerIdx).VelZ * lCyclesToStop))
            '====================== END OF OLD METHOD =====================

            .yChangeEnvironments = 0
            Dim yMsg(17) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectCommand).CopyTo(yMsg, 0)
            .GetGUIDAsString.CopyTo(yMsg, 2)
            System.BitConverter.GetBytes(.DestX).CopyTo(yMsg, 8)
            System.BitConverter.GetBytes(.DestZ).CopyTo(yMsg, 12)
            System.BitConverter.GetBytes(.DestAngle).CopyTo(yMsg, 16)
            .SetDest(.DestX, .DestZ, .DestAngle)
            If .ParentEnvir.lPlayersInEnvirCnt > 0 Then goMsgSys.BroadcastToEnvironmentClients(yMsg, .ParentEnvir)
        End With
    End Sub

    Private Sub SetPreferredRangeDest(ByRef oAttacker As Epica_Entity, ByRef oTarget As Epica_Entity, ByVal lRange As Int32)
        If oTarget Is Nothing = False AndAlso oAttacker Is Nothing = False Then
            'adjust our range to actual coordinates
            lRange *= gl_FINAL_GRID_SQUARE_SIZE

            Dim lDestX As Int32 = lRange
            Dim lDestZ As Int32 = 0

            'Get our angle to target (probably faster way then this
            Dim fAngle As Single = LineAngleDegrees(oAttacker.LocX, oAttacker.LocZ, oTarget.LocX, oTarget.LocZ) * 0.1F '/ 10.0F
            'now, rotate our points
            RotatePoint(0, 0, lDestX, lDestZ, fAngle)

            'Now, add our points
            lDestX += oTarget.LocX
            lDestZ += oTarget.LocZ

            goMsgSys.SendAIMoveRequestToPathfinding(oAttacker, lDestX, lDestZ, 0)
        End If
    End Sub

End Module

Public Class BombardmentRequest
    Public lPlayerID As Int32       'player ID that is requesting the bombardment

    Private moPlanet As Planet        'REFERENCE ONLY!
    Public ReadOnly PlanetID As Int32

    'Location of the planet in the system
    Private mlPlanetLocX As Int32
    Private mlPlanetLocY As Int32
    Private mlPlanetLocZ As Int32
    Private mlParentEnvirIndex As Int32 = -1    'index of the environment for the parent system
    Private mlPlanetEnvirIndex As Int32 = -1        'index of the environment for the planet

    'Environment Grid data
    Private mlTinyX As Int32
    Private mlTinyZ As Int32
    Private mlGridIndex As Int32
    Private mlSmallSectorID As Int32

    'Target Location data
    Public lTargetX As Int32
    Public lTargetZ As Int32

    'Now, bombardment ROF Multiplier
    Private lROFMultiplier As Int32

    'Accuracy Range
    Private lAccuracyRange As Int32

    Public BombRequestValid As Boolean = False

    Public Sub New(ByVal lPlanetID As Int32, ByVal yBombardType As Byte)
        Dim X As Int32
        Dim Y As Int32
        Dim bFound As Boolean = False

        PlanetID = lPlanetID

        For X = 0 To goGalaxy.mlSystemUB
            For Y = 0 To goGalaxy.moSystems(X).PlanetUB
                If goGalaxy.moSystems(X).moPlanets(Y).ObjectID = lPlanetID Then
                    bFound = True
                    moPlanet = goGalaxy.moSystems(X).moPlanets(Y)
                    mlPlanetLocX = moPlanet.LocX
                    mlPlanetLocY = moPlanet.LocY
                    mlPlanetLocZ = moPlanet.LocZ

                    'Now our environment indices
                    mlPlanetEnvirIndex = moPlanet.EnvirIdx
                    mlParentEnvirIndex = goGalaxy.moSystems(X).EnvirIdx
                    Exit For
                End If
            Next Y

            If bFound = True Then Exit For
        Next X

        If bFound = False Then
            BombRequestValid = False
            Exit Sub
        End If

        Select Case yBombardType
            Case BombardType.eHighYield_BT
                lAccuracyRange = 20000
                lROFMultiplier = 1
            Case BombardType.eNormal_BT
                lAccuracyRange = 10000
                lROFMultiplier = 2
            Case BombardType.eTargeted_BT
                lAccuracyRange = 5000
                lROFMultiplier = 4
            Case BombardType.ePrecision_BT
                lAccuracyRange = 1000
                lROFMultiplier = 10
        End Select

        BombRequestValid = ((mlPlanetEnvirIndex <> -1) AndAlso (moPlanet Is Nothing = False) AndAlso (mlParentEnvirIndex <> -1))

        If BombRequestValid = True Then SetGridLocs()
    End Sub

    Protected Overrides Sub Finalize()
        moPlanet = Nothing
        MyBase.Finalize()
    End Sub

    Private Sub SetGridLocs()
        Dim lTemp As Int32
        Dim lTemp2 As Int32

        With goEnvirs(mlParentEnvirIndex)

            lTemp = mlPlanetLocZ + .lHalfEnvirSize
            mlGridIndex = (lTemp \ .lGridSquareSize) * .lGridsPerRow
            lTemp -= (mlGridIndex \ .lGridsPerRow) * .lGridSquareSize
            mlSmallSectorID = (lTemp \ gl_SMALL_GRID_SQUARE_SIZE) * gl_SMALL_PER_ROW
            lTemp -= ((mlSmallSectorID \ gl_SMALL_PER_ROW) * gl_SMALL_GRID_SQUARE_SIZE)
            mlTinyZ = (lTemp \ gl_FINAL_GRID_SQUARE_SIZE)

            lTemp2 = mlPlanetLocX + .lHalfEnvirSize
            lTemp = (lTemp2 \ .lGridSquareSize)
            mlGridIndex += lTemp
            lTemp2 -= (lTemp * .lGridSquareSize)
            lTemp = (lTemp2 \ gl_SMALL_GRID_SQUARE_SIZE)
            mlSmallSectorID += lTemp
            lTemp2 -= (lTemp * gl_SMALL_GRID_SQUARE_SIZE)
            mlTinyX = (lTemp2 \ gl_FINAL_GRID_SQUARE_SIZE)

            If mlSmallSectorID < 0 Then
                mlSmallSectorID = 0
            ElseIf mlSmallSectorID > giRelativeSmall.GetUpperBound(1) Then
                mlSmallSectorID = giRelativeSmall.GetUpperBound(1)
            End If
        End With

    End Sub

    Public Sub HandleBombardmentRequest()
        Dim lGrid As Int32
        Dim lLargeSector As Int32
        Dim X As Int32
        Dim lTemp As Int32
        Dim oTmpUnit As Epica_Entity
        Dim lEnvirDist As Int32
        Dim lRelTinyX As Int32
        Dim lRelTinyZ As Int32
        Dim Y As Int32

        Dim oCurDef As Epica_Entity_Def
        Dim oWeapon As WeaponDef
        Dim lFacing As Int32

        'Where shots hit
        Dim lLocX As Int32
        Dim lLocZ As Int32
        Dim yMsg() As Byte

        Dim lCPLimit As Int32 = 300
        For X = 0 To glPlayerUB
            If glPlayerIdx(X) = lPlayerID Then
                lCPLimit = goPlayers(X).lCPLimit
                Exit For
            End If
        Next X

        'If either environment is overloaded, ignore the bombardment request
        With goEnvirs(mlParentEnvirIndex)
            For X = 0 To .lPlayersWhoHaveUnitsHereUB
                If .lPlayersWhoHaveUnitsHereIdx(X) = lPlayerID Then
                    If .lPlayersWhoHaveUnitsHereCP(X) > lCPLimit Then
                        Return
                    End If
                    Exit For
                End If
            Next X
        End With
        With goEnvirs(mlPlanetEnvirIndex)
            For X = 0 To .lPlayersWhoHaveUnitsHereUB
                If .lPlayersWhoHaveUnitsHereIdx(X) = lPlayerID Then
                    If .lPlayersWhoHaveUnitsHereCP(X) > lCPLimit Then
                        Return
                    End If
                    Exit For
                End If
            Next X
        End With

        'Determine what adjust array to use (for map wrapping of planets)
        Dim lAdjust() As Int32 = goEnvirs(mlParentEnvirIndex).lGridIdxAdjust
        If goEnvirs(mlParentEnvirIndex).ObjTypeID = ObjectType.ePlanet Then
            Dim lModVal As Int32 = mlGridIndex Mod goEnvirs(mlParentEnvirIndex).lGridsPerRow
            If lModVal = 0 Then
                'I'm on the left edge, use the left edge adjust
                lAdjust = goEnvirs(mlParentEnvirIndex).lLeftEdgeGridIdxAdjust
            ElseIf lModVal = goEnvirs(mlParentEnvirIndex).lGridsPerRow - 1 Then
                'I'm on the right edge, use the right edge adjust
                lAdjust = goEnvirs(mlParentEnvirIndex).lRightEdgeGridIdxAdjust
            End If
        End If

        For lGrid = 0 To lAdjust.GetUpperBound(0)
            'gonna reuse the llargesector here...
            'lLargeSector = mlGridIndex + goEnvirs(mlParentEnvirIndex).lGridIdxAdjust(lGrid)
            lLargeSector = mlGridIndex + lAdjust(lGrid)

            If lLargeSector > -1 AndAlso lLargeSector < goEnvirs(mlParentEnvirIndex).lGridUB + 1 Then
                'Ok, now, cycle thru the smaller list in the Grid object
                Dim oTmpGrid As EnvirGrid = goEnvirs(mlParentEnvirIndex).oGrid(lLargeSector)
                If oTmpGrid Is Nothing Then Continue For
                For X = 0 To oTmpGrid.lEntityUB
                    lTemp = oTmpGrid.lEntities(X)
                    If lTemp <> -1 Then      'is the unit not nothing?
                        oTmpUnit = goEntity(lTemp)
                        If oTmpUnit Is Nothing Then
                            oTmpGrid.lEntities(X) = -1
                            Continue For
                        End If
                        'We only care about units that belong to the bombardment requesting player
                        If oTmpUnit.Owner.ObjectID = lPlayerID Then
                            lEnvirDist = Int32.MaxValue

                            lTemp = giRelativeSmall(lGrid, mlSmallSectorID, oTmpUnit.lSmallSectorID)
                            If lTemp <> -1 Then
                                lRelTinyX = glBaseRelTinyX(lTemp) + oTmpUnit.lTinyX + (gl_HALF_FINAL_PER_ROW - mlTinyX)
                                If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                    lRelTinyZ = glBaseRelTinyZ(lTemp) + oTmpUnit.lTinyZ + (gl_HALF_FINAL_PER_ROW - mlTinyZ)
                                    If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                        lEnvirDist = gyDistances(lRelTinyX, lRelTinyZ) - oTmpUnit.lModelRangeOffset
                                        lFacing = WhatSideCanFire(oTmpUnit, lRelTinyX, lRelTinyZ)
                                    End If
                                End If
                            End If

                            'TODO: Determine how range affects bombardment here...

                            'Now, get our entity def
                            oCurDef = goEntityDefs(oTmpUnit.lEntityDefServerIndex)
                            'With our entity
                            With oTmpUnit
                                'Loop thru the weapons on the entity
                                For Y = 0 To oCurDef.WeaponDefUB
                                    'If the weapon can fire, and is facing the planet or a bomb wpn, and is operational
                                    'oCurDef.WeaponDefs(Y).ArcID = lFacing AndAlso _
                                    If .lWpnNextFireCycle(Y) <= glCurrentCycle AndAlso _
                                       oCurDef.WeaponDefs(Y).WpnGroup = WeaponGroupType.BombWeaponGroup AndAlso _
                                      (.CurrentStatus And oCurDef.WeaponDefs(Y).lEntityStatusGroup) <> 0 Then

                                        'Ok, set our weapon def
                                        oWeapon = oCurDef.WeaponDefs(Y)

                                        'Is the wpn in range?
                                        If (oWeapon.lRange * 2) >= lEnvirDist Then
                                            'Ok, weapon is in range... do our shot
                                            'because of it being bombardment, we can do instantaneous (not event)

                                            If oWeapon.yWeaponType >= WeaponType.eBomb_Green AndAlso oWeapon.yWeaponType <= WeaponType.eBomb_Purple Then
                                                If goEnvirs(mlPlanetEnvirIndex).lOBShotCycle <> glCurrentCycle Then
                                                    goEnvirs(mlPlanetEnvirIndex).lOBShotCycle = glCurrentCycle
                                                    goEnvirs(mlPlanetEnvirIndex).lOBShots = 0
                                                End If
                                                If goEnvirs(mlPlanetEnvirIndex).lOBShots > 10 Then
                                                    Continue For
                                                End If
                                                goEnvirs(mlPlanetEnvirIndex).lOBShots += 1
                                            End If

                                            'determine where we hit
                                            lLocX = CInt(Rnd() * (lAccuracyRange * 2)) - lAccuracyRange
                                            lLocZ = CInt(Rnd() * (lAccuracyRange * 2)) - lAccuracyRange

                                            'Now apply it to the center of the bombardment request
                                            lLocX += Me.lTargetX
                                            lLocZ += Me.lTargetZ

                                            'Check our boundaries
                                            'If lLocX > goEnvirs(Me.mlPlanetEnvirIndex).lMaxPosX Then lLocX -= goEnvirs(Me.mlPlanetEnvirIndex).lMapWrapAdjustX
                                            'If lLocX < goEnvirs(Me.mlPlanetEnvirIndex).lMinPosX Then lLocX += goEnvirs(Me.mlPlanetEnvirIndex).lMapWrapAdjustX
                                            If lLocZ > goEnvirs(Me.mlPlanetEnvirIndex).lMaxPosZ Then lLocZ = goEnvirs(Me.mlPlanetEnvirIndex).lMaxPosZ - (lLocZ - goEnvirs(Me.mlPlanetEnvirIndex).lMaxPosZ)
                                            If lLocZ < goEnvirs(Me.mlPlanetEnvirIndex).lMinPosZ Then lLocZ = goEnvirs(Me.mlPlanetEnvirIndex).lMinPosZ - (lLocZ - goEnvirs(Me.mlPlanetEnvirIndex).lMinPosZ) 'Math.Abs(lLocZ)

                                            ApplyBombardDamage(oWeapon, lLocX, lLocZ, .Owner, oCurDef.yHullType)

                                            'Update our fire cycle
                                            .lWpnNextFireCycle(Y) = glCurrentCycle + (oWeapon.ROF_Delay * lROFMultiplier)

                                            'Ok, now, weapon has fired, send the message to two environments...
                                            yMsg = goMsgSys.CreateBombardFireMessage(oTmpUnit, moPlanet, oWeapon.yWeaponType, lLocX, lLocZ, oWeapon.AOERange)
                                            If goEnvirs(mlPlanetEnvirIndex).lPlayersInEnvirCnt > 0 Then
                                                'send to this environment
                                                goMsgSys.BroadcastToEnvironmentClients(yMsg, goEnvirs(mlPlanetEnvirIndex))
                                            End If
                                            If goEnvirs(mlParentEnvirIndex).lPlayersInEnvirCnt > 0 Then
                                                'send to this environment
                                                goMsgSys.BroadcastToEnvironmentClients(yMsg, goEnvirs(mlParentEnvirIndex))
                                            End If

                                        End If      'Weapon in Range?
                                    End If          'Weapon Can Fire, Facing Planet or bomb, is operational
                                Next Y              'all weapons of the entity

                            End With

                        End If      'if owner = bombrequest.player
                    End If          'idx not -1

                Next X              'loop thru units in grid

            End If              'valid grid?
        Next lGrid

    End Sub

    Private Shared mbDoNewBombCode As Boolean = True
    Private Sub ApplyBombardDamage(ByVal oWeaponDef As WeaponDef, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByRef oShootingPlayer As Player, ByVal yAttackingHullType As Byte)
        Dim lGrid As Int32
        Dim lLargeSector As Int32
        Dim lRelTinyX As Int32
        Dim lRelTinyZ As Int32
        Dim X As Int32
        Dim oTmpUnit As Epica_Entity
        Dim lTemp As Int32
        Dim lTemp2 As Int32

        Dim Y As Int32
        Dim lIdx As Int32

        Dim lBaseFacing As Int32

        'Determine our aggression engine data... for the shot
        Dim lGridIndex As Int32
        Dim lSmallSectorID As Int32
        Dim lTinyX As Int32
        Dim lTinyZ As Int32

        Dim lAOE As Int32 = Math.Min(255, (oWeaponDef.AOERange + 1) * 20I)
        Dim oEnvir As Envir = goEnvirs(mlPlanetEnvirIndex)

        Dim lBeamDmg As Int32
        Dim lChemicalDmg As Int32
        Dim lECMDmg As Int32
        Dim lFlameDmg As Int32
        Dim lImpactDmg As Int32
        Dim lPierceDmg As Int32

        With oWeaponDef
            If .BeamMaxDmg > 0 Then lBeamDmg = CInt(Rnd() * .BeamMaxMinRange) + .BeamMinDmg
            If .ChemicalMaxDmg > 0 Then lChemicalDmg = CInt(Rnd() * .ChemicalMaxMinRange) + .ChemicalMinDmg
            If .ECMMaxDmg > 0 Then lECMDmg = CInt(Rnd() * .ECMMaxMinRange) + .ECMMinDmg
            If .FlameMaxDmg > 0 Then lFlameDmg = CInt(Rnd() * .FlameMaxMinRange) + .FlameMinDmg
            If .ImpactMaxDmg > 0 Then lImpactDmg = CInt(Rnd() * .ImpactMaxMinRange) + .ImpactMinDmg
            If .PiercingMaxDmg > 0 Then lPierceDmg = CInt(Rnd() * .PiercingMaxMinRange) + .PiercingMinDmg
        End With

        'NOTE: Any changes to the aggression engine calculations need to be made here too
        lTemp = lLocZ + oEnvir.lHalfEnvirSize
        'lLargeSector = CInt(Math.Floor(lTemp / oEnvir.lGridSquareSize)) * oEnvir.lGridsPerRow
        lLargeSector = (lTemp \ oEnvir.lGridSquareSize) * oEnvir.lGridsPerRow
        'lTemp -= CInt((lLargeSector / oEnvir.lGridsPerRow) * oEnvir.lGridSquareSize)
        lTemp -= (lLargeSector \ oEnvir.lGridsPerRow) * oEnvir.lGridSquareSize
        lSmallSectorID = (lTemp \ gl_SMALL_GRID_SQUARE_SIZE) * gl_SMALL_PER_ROW
        lTemp -= (lSmallSectorID \ gl_SMALL_PER_ROW) * gl_SMALL_GRID_SQUARE_SIZE
        lTinyZ = (lTemp \ gl_FINAL_GRID_SQUARE_SIZE)
        lTemp2 = lLocX + oEnvir.lHalfEnvirSize
        lTemp = (lTemp2 \ oEnvir.lGridSquareSize)
        lLargeSector += lTemp
        lTemp2 -= (lTemp * oEnvir.lGridSquareSize)
        lTemp = (lTemp2 \ gl_SMALL_GRID_SQUARE_SIZE)
        lSmallSectorID += lTemp
        lTemp2 -= (lTemp * gl_SMALL_GRID_SQUARE_SIZE)
        lTinyX = (lTemp2 \ gl_FINAL_GRID_SQUARE_SIZE)

        If lSmallSectorID < 0 OrElse lSmallSectorID > giRelativeSmall.GetUpperBound(1) Then Return

        lGridIndex = lLargeSector

        'Determine what adjust array to use (for map wrapping of planets)
        Dim lAdjust() As Int32 = oEnvir.lGridIdxAdjust
        If oEnvir.ObjTypeID = ObjectType.ePlanet Then
            Dim lModVal As Int32 = lGridIndex Mod oEnvir.lGridsPerRow
            If lModVal = 0 Then
                'I'm on the left edge, use the left edge adjust
                lAdjust = oEnvir.lLeftEdgeGridIdxAdjust
            ElseIf lModVal = oEnvir.lGridsPerRow - 1 Then
                'I'm on the right edge, use the right edge adjust
                lAdjust = oEnvir.lRightEdgeGridIdxAdjust
            End If
        End If

        'Ok, go through an aggression test on this event
        For lGrid = 0 To lAdjust.GetUpperBound(0)
            'gonna reuse the llargesector here...
            'lLargeSector = lGridIndex + oEnvir.lGridIdxAdjust(lGrid)
            lLargeSector = lGridIndex + lAdjust(lGrid)
            If lLargeSector > -1 AndAlso lLargeSector < oEnvir.lGridUB + 1 Then
                'Ok, now, cycle thru the smaller list in the Grid object
                Dim oTmpGrid As EnvirGrid = oEnvir.oGrid(lLargeSector)
                If oTmpGrid Is Nothing Then Continue For
                For X = 0 To oTmpGrid.lEntityUB
                    lTemp = oTmpGrid.lEntities(X)
                    If lTemp <> -1 Then      'is the unit not nothing?
                        oTmpUnit = goEntity(lTemp)

                        If oTmpUnit Is Nothing Then Continue For
                        If oEnvir.bEnvirAtColonyLimit = False AndAlso (oTmpUnit.Owner.lIronCurtainPlanetID = Int32.MinValue OrElse (oTmpUnit.Owner.lIronCurtainPlanetID = oEnvir.ObjectID AndAlso oEnvir.ObjTypeID = ObjectType.ePlanet)) Then Continue For

                        If oTmpUnit.Owner.ObjectID <> oShootingPlayer.ObjectID Then
                            If oTmpUnit.Owner.GetPlayerRelScore(oShootingPlayer.ObjectID, True, oTmpUnit.ServerIndex) > elRelTypes.eWar AndAlso oShootingPlayer.GetPlayerRelScore(oTmpUnit.Owner.ObjectID, False, -1) > elRelTypes.eWar Then
                                Continue For
                                ''ok, need to declare war
                                'Dim oRel As PlayerRel = New PlayerRel()
                                'oRel.oPlayerRegards = oShootingPlayer
                                'oRel.oThisPlayer = oTmpUnit.Owner
                                'oRel.WithThisScore = elRelTypes.eWar - 10
                                'oShootingPlayer.SetPlayerRel(oTmpUnit.lOwnerID, oRel)

                                'oRel = New PlayerRel
                                'oRel.oPlayerRegards = oTmpUnit.Owner
                                'oRel.oThisPlayer = oShootingPlayer
                                'oRel.WithThisScore = elRelTypes.eWar - 10
                                'oTmpUnit.Owner.SetPlayerRel(oShootingPlayer.ObjectID, oRel)
                                'oRel = Nothing

                                ''Now, send a message to the primary server that the new relationship is established, it will handle the rest
                                'Dim yMsg(30) As Byte
                                'System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerRel).CopyTo(yMsg, 0)
                                'System.BitConverter.GetBytes(oShootingPlayer.ObjectID).CopyTo(yMsg, 2)
                                'System.BitConverter.GetBytes(oTmpUnit.Owner.ObjectID).CopyTo(yMsg, 6)
                                'yMsg(10) = elRelTypes.eWar - 10


                                'Try
                                '    With oTmpUnit
                                '        .ParentEnvir.GetGUIDAsString.CopyTo(yMsg, 11)
                                '        .GetGUIDAsString.CopyTo(yMsg, 17)
                                '        System.BitConverter.GetBytes(.LocX).CopyTo(yMsg, 23)
                                '        System.BitConverter.GetBytes(.LocZ).CopyTo(yMsg, 27)
                                '    End With
                                'Catch
                                '    'do nothing
                                'End Try

                                'goMsgSys.SendToPrimary(yMsg)

                            End If
                        End If

                        'get our relative small
                        lTemp = giRelativeSmall(lGrid, lSmallSectorID, oTmpUnit.lSmallSectorID)
                        If lTemp <> -1 Then
                            lRelTinyX = glBaseRelTinyX(lTemp) + oTmpUnit.lTinyX + (gl_HALF_FINAL_PER_ROW - lTinyX)
                            If lRelTinyX > -1 AndAlso lRelTinyX < gl_REL_TINY_MAX Then
                                lRelTinyZ = glBaseRelTinyZ(lTemp) + oTmpUnit.lTinyZ + (gl_HALF_FINAL_PER_ROW - lTinyZ)
                                If lRelTinyZ > -1 AndAlso lRelTinyZ < gl_REL_TINY_MAX Then
                                    If (gyDistances(lRelTinyX, lRelTinyZ) - oTmpUnit.lModelRangeOffset) <= lAOE Then
                                        'Set our base facing value
                                        lBaseFacing = glFacing(lRelTinyX, lRelTinyZ)
                                        lBaseFacing -= 1800
                                        If lBaseFacing < 0 Then lBaseFacing += 3600

                                        'ok, find a place to add this event to
                                        lIdx = -1
                                        For Y = 0 To oTmpUnit.lWpnEventUB
                                            If oTmpUnit.yWpnEventUsed(Y) = 0 Then
                                                lIdx = Y
                                                'ensure that the event is cleared
                                                oTmpUnit.oWpnEvents(Y) = Nothing
                                                oTmpUnit.oWpnEvents(Y) = New WeaponEvent
                                                Exit For
                                            End If
                                        Next Y

                                        If lIdx = -1 Then
                                            oTmpUnit.lWpnEventUB += 1
                                            ReDim Preserve oTmpUnit.oWpnEvents(oTmpUnit.lWpnEventUB)
                                            ReDim Preserve oTmpUnit.yWpnEventUsed(oTmpUnit.lWpnEventUB)
                                            oTmpUnit.oWpnEvents(oTmpUnit.lWpnEventUB) = New WeaponEvent()
                                            lIdx = oTmpUnit.lWpnEventUB
                                        End If

                                        oTmpUnit.yWpnEventUsed(lIdx) = 255
                                        With oTmpUnit.oWpnEvents(lIdx)
                                            .AOERange = 0
                                            .yCritical = eyCriticalHitType.NoCriticalHit
                                            .lBeamDmg = lBeamDmg
                                            .lChemicalDmg = lChemicalDmg
                                            .lECMDmg = lECMDmg
                                            .lFlameDmg = lFlameDmg
                                            .lImpactDmg = lImpactDmg
                                            .lPierceDmg = lPierceDmg
                                            .lEventCycle = glCurrentCycle
                                            .lFromIdx = -1
                                            .lToIdx = oTmpUnit.ServerIndex
                                            .ySideHit = WhatSideDoIHit(lGridIndex, lSmallSectorID, lTinyX, lTinyZ, oTmpUnit)
                                            .FromPlayerID = oShootingPlayer.ObjectID
                                            .yAttackerType = yAttackingHullType

                                            oTmpUnit.lNextWpnEvent = Math.Min(oTmpUnit.lNextWpnEvent, .lEventCycle)
                                        End With

                                        If mbDoNewBombCode = True Then
                                            Try
                                                If (oTmpUnit.CurrentStatus And elUnitStatus.eUnitEngaged) = 0 AndAlso oTmpUnit.bForceAggressionTest = False Then
                                                    oTmpUnit.CurrentStatus = oTmpUnit.CurrentStatus Or elUnitStatus.eUnitEngaged
                                                    AddEntityInCombat(oTmpUnit.ObjectID, oTmpUnit.ObjTypeID, oTmpUnit.ServerIndex)
                                                End If
                                            Catch
                                            End Try
                                        End If


                                    End If
                                End If
                            End If
                        End If
                    End If
                Next X
            End If
        Next lGrid

        oEnvir = Nothing

    End Sub

End Class