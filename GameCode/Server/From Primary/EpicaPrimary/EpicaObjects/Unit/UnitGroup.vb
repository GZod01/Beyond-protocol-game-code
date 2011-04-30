Public Class UnitGroup  'represents a fleet, army, etc... a group of units
    Inherits Epica_GUID

    Public UnitGroupName(19) As Byte

    'This array is a reference to the ID of the unit
    Private GroupUnitsID() As Int32
    'This Array points to the global goUnit Array... the value of GroupUnitsIdx is the index of goUnit
    Private GroupUnitsIdx() As Int32
    Private mbUnitRegrouped() As Boolean
    Public UnitUB As Int32 = -1

    Public oOwner As Player

    'This is the environment the fleet currently exists
    Public oParentObject As Object

    Private mlInterSystemMoveCyclesRemaining As Int32     'number of cycles remaining for an inter-system movement
    Public lLastInterSystemCycleUpdate As Int32         'last cycle that the intersystem move cycles remaining was calculated
    Public lInterSystemTargetID As Int32 = -1           'target system ID of the intersystem movement
    Public lInterSystemOriginID As Int32 = -1           'origin system ID of the intersystem movement
    Public yInterSystemSpeed As Byte

    Public lOriginX As Int32
    Public lOriginY As Int32
    Public lOriginZ As Int32

    Public lQueueItemIdx As Int32 = -1

    Private mySendString() As Byte

    Private mlReinforcerIdx() As Int32      'server index of the goFacility array for a facility ordered to reinforce this unitgroup
    Private mlReinforcerID() As Int32       'FacilityID of the facility ordered to reinforce this unitgroup
    Private mlReinforcerUB As Int32 = -1

    Public lDefaultFormationID As Int32 = -1

    Private mlInterSystemTotalCycles As Int32 = -1
    Public ReadOnly Property InterSystemTotalCycles() As Int32
        Get
            If mlInterSystemTotalCycles = -1 Then
                If lInterSystemOriginID = -1 OrElse lInterSystemTargetID = -1 Then Return 0

                Dim lOX As Int32 = lOriginX
                Dim lOY As Int32 = lOriginY
                Dim lOZ As Int32 = lOriginZ
                Dim lTX As Int32
                Dim lTY As Int32
                Dim lTZ As Int32
                For X As Int32 = 0 To glSystemUB
                    If glSystemIdx(X) = Me.lInterSystemTargetID Then
                        With goSystem(X)
                            lTX = .LocX : lTY = .LocY : lTZ = .LocZ
                        End With
                        Exit For
                    End If
                Next X

                Dim fDX As Single = lOX - lTX : fDX *= fDX
                Dim fDY As Single = lOY - lTY : fDY *= fDY
                Dim fDZ As Single = lOZ - lTZ : fDZ *= fDZ
                Dim fDist As Single = CSng(Math.Sqrt(fDX + fDY + fDZ))
                fDist *= 10000000

                'TODO: Determine the player's Speed Multiplier here
                Dim fPlayerMult As Single = 1.0F
                fDist /= (yInterSystemSpeed * fPlayerMult)

                If fDist > Int32.MaxValue Then mlInterSystemTotalCycles = 0 Else mlInterSystemTotalCycles = CInt(fDist)
            End If
            Return mlInterSystemTotalCycles
        End Get
    End Property

    Public Property InterSystemMoveCyclesRemaining() As Int32
        Get
            'if the target is not -1 andalso we are in transit (our parent is the galaxy object)
            If lInterSystemTargetID <> -1 AndAlso CType(oParentObject, Epica_GUID).ObjTypeID = ObjectType.eGalaxy Then
                ' we are moving, so update our value
                Dim lDiff As Int32 = glCurrentCycle - lLastInterSystemCycleUpdate
                mlInterSystemMoveCyclesRemaining -= lDiff
                lLastInterSystemCycleUpdate = glCurrentCycle
            Else : lLastInterSystemCycleUpdate = glCurrentCycle
            End If

            Return mlInterSystemMoveCyclesRemaining
        End Get
        Set(ByVal value As Int32)
            mlInterSystemMoveCyclesRemaining = value
            lLastInterSystemCycleUpdate = glCurrentCycle
        End Set
    End Property

    Public Function GetObjAsString() As Byte()
        Dim lCnt As Int32 = 0

        For X As Int32 = 0 To UnitUB
            If GroupUnitsIdx(X) <> -1 AndAlso GroupUnitsID(X) = glUnitIdx(GroupUnitsIdx(X)) Then
                lCnt += 1
            End If
        Next X

        Dim lReinforceCnt As Int32 = 0
        For X As Int32 = 0 To mlReinforcerUB
            If mlReinforcerIdx(X) <> -1 AndAlso mlReinforcerID(X) = glFacilityIdx(mlReinforcerIdx(X)) AndAlso mlReinforcerID(X) <> -1 Then lReinforceCnt += 1
        Next X

        'here we will return the entire object as a string
        ReDim mySendString(72 + (lCnt * 10) + (lReinforceCnt * 10)) '76

        Dim lPos As Int32 = 0

        'Dim lWPUpkeep As Int32 = 0
        'Dim lWPUpkeepPos As Int32

        GetGUIDAsString.CopyTo(mySendString, lPos) : lPos += 6
        System.BitConverter.GetBytes(oOwner.ObjectID).CopyTo(mySendString, lPos) : lPos += 4
        UnitGroupName.CopyTo(mySendString, lPos) : lPos += 20
        CType(oParentObject, Epica_GUID).GetGUIDAsString.CopyTo(mySendString, lPos) : lPos += 6

        System.BitConverter.GetBytes(InterSystemMoveCyclesRemaining).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lInterSystemTargetID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lInterSystemOriginID).CopyTo(mySendString, lPos) : lPos += 4
        mySendString(lPos) = yInterSystemSpeed : lPos += 1

        System.BitConverter.GetBytes(lOriginX).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lOriginY).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lOriginZ).CopyTo(mySendString, lPos) : lPos += 4

        System.BitConverter.GetBytes(lDefaultFormationID).CopyTo(mySendString, lPos) : lPos += 4
        'lWPUpkeepPos = lPos : lPos += 4

        System.BitConverter.GetBytes(lCnt).CopyTo(mySendString, lPos) : lPos += 4
        For X As Int32 = 0 To UnitUB
            If GroupUnitsIdx(X) <> -1 AndAlso GroupUnitsID(X) = glUnitIdx(GroupUnitsIdx(X)) Then
				System.BitConverter.GetBytes(glUnitIdx(GroupUnitsIdx(X))).CopyTo(mySendString, lPos) : lPos += 4
				Dim oUnit As Unit = goUnit(GroupUnitsIdx(X))
				If oUnit Is Nothing = False Then
                    CType(oUnit.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(mySendString, lPos)
                    'lWPUpkeep += oUnit.EntityDef.WarpointUpkeep
				End If
				lPos += 6
			End If
        Next X
        System.BitConverter.GetBytes(lReinforceCnt).CopyTo(mySendString, lPos) : lPos += 4
        For X As Int32 = 0 To mlReinforcerUB
            If mlReinforcerIdx(X) <> -1 AndAlso mlReinforcerID(X) = glFacilityIdx(mlReinforcerIdx(X)) AndAlso mlReinforcerID(X) <> -1 Then
				System.BitConverter.GetBytes(mlReinforcerID(X)).CopyTo(mySendString, lPos) : lPos += 4
				Dim oFac As Facility = goFacility(mlReinforcerIdx(X))
				If oFac Is Nothing = False Then
					CType(oFac.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(mySendString, lPos)
				End If
				lPos += 6
            End If
        Next X

        'System.BitConverter.GetBytes(lWPUpkeep).CopyTo(mySendString, lWPUpkeepPos)

        Return mySendString
    End Function

#Region "  Launch to Reinforce logic  "
    Public Sub AddReinforcer(ByVal lFacilityID As Int32)
        For X As Int32 = 0 To mlReinforcerUB
            If mlReinforcerID(X) = lFacilityID Then
                If mlReinforcerIdx(X) = -1 Then
                    For Y As Int32 = 0 To glFacilityUB
                        If glFacilityIdx(Y) = lFacilityID Then
                            mlReinforcerIdx(X) = Y
                            Exit For
                        End If
                    Next Y
                End If
                Return
            End If
        Next X

        Dim lFacIdx As Int32 = -1
        For X As Int32 = 0 To glFacilityUB
            If glFacilityIdx(X) = lFacilityID Then
                lFacIdx = X
                Exit For
            End If
        Next X
        If lFacIdx = -1 Then Return

        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To mlReinforcerUB
            If mlReinforcerIdx(X) = -1 Then
                lIdx = X
                Exit For
            End If
        Next X
        If lIdx = -1 Then
            mlReinforcerUB += 1
            ReDim Preserve mlReinforcerIdx(mlReinforcerUB)
            ReDim Preserve mlReinforcerID(mlReinforcerUB)
            lIdx = mlReinforcerUB
        End If
        mlReinforcerIdx(lIdx) = lFacIdx
        mlReinforcerID(lIdx) = lFacilityID
    End Sub
    Public Sub RemoveReinforcer(ByVal lFacilityID As Int32)
        For X As Int32 = 0 To mlReinforcerUB
            If mlReinforcerIdx(X) <> -1 AndAlso mlReinforcerID(X) = lFacilityID Then
                mlReinforcerID(X) = -1
                mlReinforcerIdx(X) = -1
            End If
        Next X
    End Sub
    Private Function HasReinforcers() As Boolean
        For X As Int32 = 0 To mlReinforcerUB
            If mlReinforcerIdx(X) <> -1 Then
                If glFacilityIdx(mlReinforcerIdx(X)) = mlReinforcerID(X) AndAlso mlReinforcerID(X) <> -1 Then
                    Return True
                Else
                    mlReinforcerIdx(X) = -1
                    mlReinforcerID(X) = -1
                End If
            End If
        Next X
        Return False
    End Function
    Private Sub ProcessReinforcers(ByVal lDefID As Int32, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16)
        Dim oDef As Epica_Entity_Def = GetEpicaUnitDef(lDefID)
        If oDef Is Nothing = True Then Return
		Dim oProducer As Facility = Nothing

		'if odef.yChassisType and chassistype.eGroundBased)
		If iEnvirTypeID = ObjectType.eSolarSystem Then
			If (oDef.yChassisType And ChassisType.eSpaceBased) = 0 Then Return
		ElseIf iEnvirTypeID = ObjectType.ePlanet Then
            If (oDef.yChassisType And ChassisType.eAtmospheric) = 0 AndAlso (oDef.yChassisType And ChassisType.eGroundBased) = 0 AndAlso (oDef.yChassisType And ChassisType.eNavalBased) = 0 Then Return
		End If

        Dim oNextBestFac As Facility = Nothing
        Dim oNextBestUnit As Unit = Nothing
        Dim lNextBestCombatRating As Int32 = 0

        For X As Int32 = 0 To mlReinforcerUB
            'check if reinforcer is still valid
            If mlReinforcerIdx(X) <> -1 Then
                If mlReinforcerID(X) <> -1 AndAlso glFacilityIdx(mlReinforcerIdx(X)) = mlReinforcerID(X) Then
                    'reinforcer valid, set facility
                    Dim oFac As Facility = goFacility(mlReinforcerIdx(X))
                    If oFac Is Nothing = False Then
                        'Is facility powered and is the hangar operational
						If oFac Is Nothing = True OrElse _
						   oFac.Active = False OrElse _
						   (oFac.CurrentStatus And elUnitStatus.eHangarOperational) = 0 Then Continue For

                        If (oDef.yChassisType And (ChassisType.eGroundBased Or ChassisType.eNavalBased)) <> 0 Then
                            If oFac.ParentObject Is Nothing Then Continue For
                            With CType(oFac.ParentObject, Epica_GUID)
                                If .ObjectID <> lEnvirID OrElse .ObjTypeID <> iEnvirTypeID Then Continue For
                            End With
                        End If

						'Check contents of facility's hangar, 
						'   is there a unit that matches the def of the unit recently lost that is NOT in a battlegroup already?
						For lHIdx As Int32 = 0 To oFac.lHangarUB
							If oFac.lHangarIdx(lHIdx) <> -1 AndAlso oFac.oHangarContents(lHIdx).ObjTypeID = ObjectType.eUnit Then
								With CType(oFac.oHangarContents(lHIdx), Unit)
									If .EntityDef.ObjectID = oDef.ObjectID AndAlso .lFleetID < 1 Then
										'   Yes:
										'        Place unit in queue with Undock_To_Reinforce
										'        TargetID is unit's parent
										'        ObjectGUID is Unit
										'        Set Unit's BattlegroupID to the UnitGroup and add the unit to the battlegroup
										'TODO: Remove this lookup loop with a serverindex lookup or something
										Dim bFound As Boolean = False
										For Z As Int32 = 0 To glUnitUB
											If glUnitIdx(Z) = .ObjectID Then
                                                AddUnit(Z, True)
												bFound = True
												Exit For
											End If
										Next Z
										If bFound = True Then
											'AddToQueueExtended(glCurrentCycle, QueueItemType.eUndock_To_Reinforce, .ObjectID, .ObjTypeID, oFac.ObjectID, oFac.ObjTypeID, lEnvirID, iEnvirTypeID, lLocX, lLocZ)
											AddToQueue(glCurrentCycle, QueueItemType.eUndock_To_Reinforce, .ObjectID, .ObjTypeID, oFac.ObjectID, oFac.ObjTypeID, lEnvirID, iEnvirTypeID, lLocX, lLocZ)
											'        DONE()
											Return
										End If
									Else
										'Ok, check combat rating
										If .EntityDef.yChassisType = oDef.yChassisType AndAlso .EntityDef.CombatRating > lNextBestCombatRating Then
											lNextBestCombatRating = .EntityDef.CombatRating
											oNextBestFac = oFac
											oNextBestUnit = CType(oFac.oHangarContents(lHIdx), Unit)
										End If
									End If
								End With
							End If
						Next lHIdx

						'   No: not found
						'        If facility is not producing, save facility for producer use
						If oFac.bProducing = False AndAlso oProducer Is Nothing Then
							If oFac.yProductionType = oDef.RequiredProductionTypeID Then
								If oFac.ParentColony Is Nothing = False Then
									Dim lParentID As Int32 = CType(oFac.ParentObject, Epica_GUID).ObjectID
									Dim iParentType As Int16 = CType(oFac.ParentObject, Epica_GUID).ObjTypeID

									Dim bCanBuildReinforcement As Boolean = False
                                    If (oDef.yChassisType And (ChassisType.eGroundBased Or ChassisType.eNavalBased)) <> 0 Then
                                        If lParentID = lEnvirID AndAlso iParentType = iEnvirTypeID Then
                                            bCanBuildReinforcement = True
                                        End If
                                    Else : bCanBuildReinforcement = True
                                    End If

                                    If bCanBuildReinforcement = True AndAlso oFac.ParentColony.HasRequiredResources(oDef.ProductionRequirements, oFac, eyHasRequiredResourcesFlags.NoReduction Or eyHasRequiredResourcesFlags.DoNotWait) = eResourcesResult.Sufficient Then
                                        oProducer = oFac
                                    End If
								End If
							End If
						End If
					Else
						'reinforcer no longer valid
						mlReinforcerID(X) = -1
						mlReinforcerIdx(X) = -1
					End If
				Else
					'reinforcer no longer valid
					mlReinforcerID(X) = -1
					mlReinforcerIdx(X) = -1
				End If
			End If
		Next X

        'If none were found then we're here, check oNextBestUnit and oNextBestFac
        If oNextBestUnit Is Nothing = False AndAlso oNextBestFac Is Nothing = False Then
            'Ok, let's use it

            'TODO: Remove this unnecessary llookup
            Dim bFound As Boolean = False
            For Z As Int32 = 0 To glUnitUB
                If glUnitIdx(Z) = oNextBestUnit.ObjectID Then
                    AddUnit(Z, True)
                    bFound = True
                    Exit For
                End If
            Next Z
            If bFound = True Then
				'AddToQueueExtended(glCurrentCycle, QueueItemType.eUndock_To_Reinforce, oNextBestUnit.ObjectID, oNextBestUnit.ObjTypeID, oNextBestFac.ObjectID, oNextBestFac.ObjTypeID, lEnvirID, iEnvirTypeID, lLocX, lLocZ)
				AddToQueue(glCurrentCycle, QueueItemType.eUndock_To_Reinforce, oNextBestUnit.ObjectID, oNextBestUnit.ObjTypeID, oNextBestFac.ObjectID, oNextBestFac.ObjTypeID, lEnvirID, iEnvirTypeID, lLocX, lLocZ)
                Return
            End If
        End If
        'Ok, if we're here, check our producer, perhaps we can build what we need
        If oProducer Is Nothing = False Then
            'order producer to build a replacement
            'TODO: Remove unnecessary search
            Dim lFacIdx As Int32 = -1
            For X As Int32 = 0 To glFacilityUB
                If glFacilityIdx(X) = oProducer.ObjectID Then
                    lFacIdx = X
                    Exit For
                End If
            Next X
            If lFacIdx <> -1 Then
                Dim blCycles As Int64 = oProducer.mlProdPoints
                Dim lCycle As Int32 = glCurrentCycle
                If blCycles <> 0 Then blCycles = oDef.ProductionRequirements.PointsRequired \ oProducer.mlProdPoints
                If blCycles + glCurrentCycle < Int32.MaxValue Then lCycle = CInt(blCycles + glCurrentCycle)

                If oProducer.AddProduction(oDef.ObjectID, oDef.ObjTypeID, 254, 1, 0) = True Then
                    'Ok, we got it... respond with success
					AddEntityProducing(lFacIdx, ObjectType.eFacility, oProducer.ObjectID)
                    '       Place in QUEUE: Reinforcements_Built
                    '         ObjectGUID is facility
                    '         TargetID1 is unitgroupID
                    '         TargetID2 is UnitDefID being built
                    '       Set queue time for when production is completed
					'AddToQueueExtended(lCycle, QueueItemType.eReinforcements_Built, oProducer.ObjectID, oProducer.ObjTypeID, Me.ObjectID, oDef.ObjectID, lEnvirID, iEnvirTypeID, lLocX, lLocZ)
					AddToQueue(lCycle, QueueItemType.eReinforcements_Built, oProducer.ObjectID, oProducer.ObjTypeID, Me.ObjectID, oDef.ObjectID, lEnvirID, iEnvirTypeID, lLocX, lLocZ)
                End If
            End If
        End If
    End Sub
    Public Sub FinalizeLaunchToReinforce(ByRef oEntity As Epica_Entity, ByVal lEnvirID As Int32, ByVal lEnvirTypeID As Int32, ByVal lLocX As Int32, ByVal lLocZ As Int32)

        'first, verify we can get there
        If oEntity.ObjTypeID = ObjectType.eUnit Then
            If (CType(oEntity, Unit).EntityDef.yChassisType And (ChassisType.eGroundBased Or ChassisType.eNavalBased)) <> 0 Then
                If lEnvirTypeID <> ObjectType.ePlanet Then Return
                With CType(oEntity.ParentObject, Epica_GUID)
                    If .ObjectID <> lEnvirID OrElse .ObjTypeID <> lEnvirTypeID Then Return
                End With
            End If
        Else : Return
        End If

        If lEnvirID > 0 AndAlso lEnvirTypeID > 0 AndAlso (lEnvirTypeID = ObjectType.ePlanet OrElse lEnvirTypeID = ObjectType.eSolarSystem) Then

            Dim lListX(UnitUB) As Int32
            Dim lListZ(UnitUB) As Int32
            Dim lIdx As Int32 = -1

            For X As Int32 = 0 To UnitUB
                If GroupUnitsIdx(X) <> -1 Then
					If glUnitIdx(GroupUnitsIdx(X)) = GroupUnitsID(X) AndAlso GroupUnitsID(X) <> -1 Then
						If goUnit(GroupUnitsIdx(X)).ObjectID <> oEntity.ObjectID Then
							With CType(goUnit(GroupUnitsIdx(X)).ParentObject, Epica_GUID)
								If .ObjectID = lEnvirID AndAlso .ObjTypeID = lEnvirTypeID Then
									lIdx += 1
									lListX(lIdx) = goUnit(GroupUnitsIdx(X)).LocX
									lListZ(lIdx) = goUnit(GroupUnitsIdx(X)).LocZ
								End If
							End With
						End If
					Else
						GroupUnitsIdx(X) = -1
						GroupUnitsID(X) = -1
					End If
				End If
            Next X

            'Ok, now, llist is populated and lIdx is our ubound
            If lIdx <> -1 Then
                'calculate our loc...
                Dim lCloseMatches(lIdx) As Int32
                For X As Int32 = 0 To lIdx
                    lCloseMatches(X) = 0
                    For Y As Int32 = 0 To lIdx
                        If X <> Y Then
                            Dim fX As Single = lListX(X) - lListX(Y)
                            Dim fZ As Single = lListZ(X) - lListZ(Y)
                            fX *= fX
                            fZ *= fZ
                            If fX + fZ < 655360000 Then     'this number is (1024 * 25)^2
                                lCloseMatches(X) += 1
                            End If
                        End If
                    Next Y
                Next X

                'Now, who has the most matches
                Dim lHighest As Int32 = 0
                Dim lHighestLocIdx As Int32 = -1
                For X As Int32 = 0 To lIdx
                    If lCloseMatches(X) > lHighest Then
                        lHighest = lCloseMatches(X)
                        lHighestLocIdx = X
                    End If
                Next X

                If lHighestLocIdx <> -1 Then
                    lLocX = lListX(lHighestLocIdx)
                    lLocZ = lListZ(lHighestLocIdx)
                End If
            End If

            'Now, we have a loc, send it to the region
            Dim yMsg(23) As Byte
            Dim lPos As Int32 = 0
            With oEntity
				System.BitConverter.GetBytes(GlobalMessageCode.eMoveObjectRequest).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(lLocX).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(lLocZ).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(0).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(CShort(lEnvirTypeID)).CopyTo(yMsg, lPos) : lPos += 2
				.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6

				If lEnvirTypeID = ObjectType.ePlanet Then
					Dim oPlanet As Planet = GetEpicaPlanet(lEnvirID)
					If oPlanet Is Nothing = False Then oPlanet.oDomain.DomainSocket.SendData(yMsg)
				ElseIf lEnvirTypeID = ObjectType.eSolarSystem Then
					Dim oSystem As SolarSystem = GetEpicaSystem(lEnvirID)
					If oSystem Is Nothing = False Then oSystem.oDomain.DomainSocket.SendData(yMsg)
				End If
			End With
		End If

	End Sub
#End Region

    Public Sub AddUnit(ByVal lUnitServerIdx As Int32, ByVal bSave As Boolean)
        Dim lIdx As Int32 = -1

        For X As Int32 = 0 To UnitUB
            If GroupUnitsIdx(X) = lUnitServerIdx Then Return
            If GroupUnitsIdx(X) = -1 AndAlso lIdx = -1 Then lIdx = X
        Next X

        If lIdx = -1 Then
            UnitUB += 1
            ReDim Preserve GroupUnitsIdx(UnitUB)
            ReDim Preserve GroupUnitsID(UnitUB)
            lIdx = UnitUB
        End If

        GroupUnitsIdx(lIdx) = lUnitServerIdx
        GroupUnitsID(lIdx) = glUnitIdx(lUnitServerIdx)

        'Ok, check if the unit already has a fleet ID
        Dim oUnit As Unit = goUnit(lUnitServerIdx)
        If oUnit Is Nothing = False Then
            If oUnit.lFleetID > 0 AndAlso oUnit.lFleetID <> Me.ObjectID Then
                Dim lFleetID As Int32 = oUnit.lFleetID
                Dim oUnitGroup As UnitGroup = GetEpicaUnitGroup(lFleetID)
                If oUnitGroup Is Nothing = False Then oUnitGroup.RemoveUnit(oUnit.ObjectID, True, False)
            End If
            'Now, set the Unit's fleet ID to me
            oUnit.lFleetID = Me.ObjectID
            If bSave = True Then oUnit.SaveObject()
        End If

        CalculateActualSystemSpeed()
    End Sub

	Public Sub RemoveUnit(ByVal lID As Int32, ByVal bAlertPlayer As Boolean, ByVal bUnitDestroyed As Boolean)
		Dim X As Int32
		Dim bNeedToRecalc As Boolean = False
		Dim bUnitGroupExists As Boolean = False
		Dim bHasReinforcers As Boolean = HasReinforcers()
		Dim lDefID As Int32 = -1

		Dim lLocX As Int32 = 0
		Dim lLocZ As Int32 = 0
		Dim lEnvirID As Int32 = -1
		Dim iEnvirTypeID As Int16 = -1
		For X = 0 To UnitUB
			If GroupUnitsIdx(X) <> -1 AndAlso glUnitIdx(GroupUnitsIdx(X)) = lID Then
				'found it... check if the unit's speed is the yInterSystemSpeed value
				Dim oUnit As Unit = goUnit(GroupUnitsIdx(X))
				If oUnit Is Nothing = False Then
					With oUnit
						lLocX = .LocX
						lLocZ = .LocZ
					End With
					If oUnit.ParentObject Is Nothing = False Then
						Try
							With CType(oUnit.ParentObject, Epica_GUID)
								lEnvirID = .ObjectID
								iEnvirTypeID = .ObjTypeID
							End With
							If iEnvirTypeID = ObjectType.eFacility Then
								With CType(CType(oUnit.ParentObject, Facility).ParentObject, Epica_GUID)
									lEnvirID = .ObjectID
									iEnvirTypeID = .ObjTypeID
								End With
							ElseIf iEnvirTypeID = ObjectType.eUnit Then
								With CType(CType(oUnit.ParentObject, Unit).ParentObject, Epica_GUID)
									lEnvirID = .ObjectID
									iEnvirTypeID = .ObjTypeID
								End With
							End If
						Catch
							lEnvirID = -1
							iEnvirTypeID = -1
						End Try
					End If

					lDefID = oUnit.EntityDef.ObjectID
					bNeedToRecalc = (oUnit.EntityDef.MaxSpeed = yInterSystemSpeed)
                    oUnit.lFleetID = -1
                    If bUnitDestroyed = False Then oUnit.SaveObject()
					GroupUnitsIdx(X) = -1
					GroupUnitsID(X) = -1
				End If

				Exit For
			End If
		Next X

		For X = 0 To UnitUB
            If GroupUnitsIdx(X) <> -1 AndAlso glUnitIdx(GroupUnitsIdx(X)) <> -1 Then
                If GroupUnitsID(X) = glUnitIdx(GroupUnitsIdx(X)) Then
                    bUnitGroupExists = True
                    Exit For
                End If
            End If
        Next X

		If bUnitGroupExists = False Then
			For X = 0 To glUnitGroupUB
                If glUnitGroupIdx(X) = Me.ObjectID Then
                    goUnitGroup(X).DeleteMe()
                    glUnitGroupIdx(X) = -1
                    goUnitGroup(X) = Nothing
                    Exit For
                End If
			Next X
		Else
			If bUnitDestroyed = True AndAlso bHasReinforcers = True AndAlso lDefID <> -1 Then
				'Ok, let's do some reinforcement
				ProcessReinforcers(lDefID, lLocX, lLocZ, lEnvirID, iEnvirTypeID)
			End If

            If bAlertPlayer = True AndAlso (oOwner.lConnectedPrimaryID > -1 OrElse oOwner.HasOnlineAliases(AliasingRights.eModifyBattleGroups) = True) Then
                Dim yMsg(9) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eRemoveFromFleet).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 2)
                System.BitConverter.GetBytes(lID).CopyTo(yMsg, 6)
                oOwner.SendPlayerMessage(yMsg, False, AliasingRights.eViewBattleGroups)
            End If

			If bNeedToRecalc = True Then CalculateActualSystemSpeed()
		End If

	End Sub

	Public Sub CalculateActualSystemSpeed()
		Dim lMaxSpeed As Int32 = Int32.MaxValue

		Dim yInitialSpeed As Byte = yInterSystemSpeed

		For X As Int32 = 0 To Me.UnitUB
			If GroupUnitsIdx(X) <> -1 AndAlso glUnitIdx(GroupUnitsIdx(X)) = GroupUnitsID(X) Then
				Dim oUnit As Unit = goUnit(GroupUnitsIdx(X))
				If oUnit Is Nothing = False Then
					With oUnit
						If (.EntityDef.yChassisType And ChassisType.eSpaceBased) <> 0 Then
							'Ok, it is a space-based unit, so check if its parent is a unit or facility
							Dim iTypeID As Int16 = CType(.ParentObject, Epica_GUID).ObjTypeID
							If iTypeID <> ObjectType.eFacility AndAlso iTypeID <> ObjectType.eUnit Then
								'Ok, it can fly and it is flying
								lMaxSpeed = Math.Min(.EntityDef.MaxSpeed, lMaxSpeed)
							End If
						End If
					End With
				End If
			End If
		Next X

		If lMaxSpeed < 0 Then lMaxSpeed = 0
		If lMaxSpeed > 255 Then lMaxSpeed = 255
		yInterSystemSpeed = CByte(lMaxSpeed)

        If yInitialSpeed <> yInterSystemSpeed AndAlso (oOwner.lConnectedPrimaryID > -1 OrElse oOwner.HasOnlineAliases(AliasingRights.eViewBattleGroups) = True) Then
            Dim yMsg(6) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eUpdateFleetSpeed).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(ObjectID).CopyTo(yMsg, 2)
            yMsg(6) = yInterSystemSpeed

            oOwner.SendPlayerMessage(yMsg, False, AliasingRights.eViewBattleGroups)
        End If
	End Sub

	Public Function SaveObject() As Boolean
		Dim bResult As Boolean = False
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

		Try
			If ObjectID = -1 Then
				'INSERT
                sSQL = "INSERT INTO tblUnitGroup (UnitGroupName, OwnerID, ParentID, ParentTypeID, InterSystemCycles, " & _
                  "InterSystemTargetID, InterSystemOrigin, InterSystemSpeed, OriginX, OriginY, OriginZ, DefaultFormationID) VALUES ('" & _
                  MakeDBStr(BytesToString(UnitGroupName)) & "', " & oOwner.ObjectID & ", " & CType(oParentObject, Epica_GUID).ObjectID & ", " & _
                  CType(oParentObject, Epica_GUID).ObjTypeID & ", " & InterSystemMoveCyclesRemaining & ", " & _
                  lInterSystemTargetID & ", " & lInterSystemOriginID & ", " & yInterSystemSpeed & ", " & lOriginX & _
                  ", " & lOriginY & ", " & lOriginZ & ", " & lDefaultFormationID & ")"
			Else
                sSQL = "UPDATE tblUnitGroup SET UnitGroupName = '" & MakeDBStr(BytesToString(UnitGroupName)) & "', OwnerID = " & _
                  oOwner.ObjectID & ", ParentID = " & CType(oParentObject, Epica_GUID).ObjectID & ", ParentTypeID = " & _
                  CType(oParentObject, Epica_GUID).ObjTypeID & ", InterSystemCycles = " & InterSystemMoveCyclesRemaining & _
                  ", InterSystemTargetID = " & lInterSystemTargetID & ", InterSystemOrigin = " & lInterSystemOriginID & _
                  ", InterSystemSpeed = " & yInterSystemSpeed & ", OriginX = " & lOriginX & ", OriginY = " & lOriginY & _
                  ", OriginZ = " & lOriginZ & ", DefaultFormationID = " & lDefaultFormationID & " WHERE UnitGroupID = " & Me.ObjectID
			End If
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			If oComm.ExecuteNonQuery() = 0 Then
				Err.Raise(-1, "SaveObject", "No records affected when saving object!")
			End If
			If ObjectID = -1 Then
				Dim oData As OleDb.OleDbDataReader
				oComm = Nothing
                sSQL = "SELECT MAX(UnitGroupID) FROM tblUnitGroup WHERE UnitGroupName = '" & MakeDBStr(BytesToString(UnitGroupName)) & "'"
				oComm = New OleDb.OleDbCommand(sSQL, goCN)
				oData = oComm.ExecuteReader(CommandBehavior.Default)
				If oData.Read Then
					ObjectID = CInt(oData(0))
				End If
				oData.Close()
				oData = Nothing
			End If
			bResult = True
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
		Finally
			oComm = Nothing
		End Try
		Return bResult
	End Function

	Public Function GetSaveObjectText() As String
		Dim sSQL As String

		Try
			If ObjectID = -1 Then
				'INSERT
                sSQL = "INSERT INTO tblUnitGroup (UnitGroupName, OwnerID, ParentID, ParentTypeID, InterSystemCycles, " & _
                  "InterSystemTargetID, InterSystemOrigin, InterSystemSpeed, OriginX, OriginY, OriginZ, DefaultFormationID) VALUES ('" & _
                  MakeDBStr(BytesToString(UnitGroupName)) & "', " & oOwner.ObjectID & ", " & CType(oParentObject, Epica_GUID).ObjectID & ", " & _
                  CType(oParentObject, Epica_GUID).ObjTypeID & ", " & InterSystemMoveCyclesRemaining & ", " & _
                  lInterSystemTargetID & ", " & lInterSystemOriginID & ", " & yInterSystemSpeed & ", " & lOriginX & _
                  ", " & lOriginY & ", " & lOriginZ & ", " & lDefaultFormationID & ")"
			Else
                sSQL = "UPDATE tblUnitGroup SET UnitGroupName = '" & MakeDBStr(BytesToString(UnitGroupName)) & "', OwnerID = " & _
                  oOwner.ObjectID & ", ParentID = " & CType(oParentObject, Epica_GUID).ObjectID & ", ParentTypeID = " & _
                  CType(oParentObject, Epica_GUID).ObjTypeID & ", InterSystemCycles = " & InterSystemMoveCyclesRemaining & _
                  ", InterSystemTargetID = " & lInterSystemTargetID & ", InterSystemOrigin = " & lInterSystemOriginID & _
                  ", InterSystemSpeed = " & yInterSystemSpeed & ", OriginX = " & lOriginX & ", OriginY = " & lOriginY & _
                  ", OriginZ = " & lOriginZ & ", DefaultFormationID = " & lDefaultFormationID & " WHERE UnitGroupID = " & Me.ObjectID
			End If
			Return sSQL
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
		End Try
		Return ""
	End Function

	Public Sub DeleteMe()
        Try
            For X As Int32 = 0 To UnitUB
                If GroupUnitsIdx(X) <> -1 AndAlso glUnitIdx(GroupUnitsIdx(X)) = GroupUnitsID(X) Then
                    Dim oUnit As Unit = goUnit(X)
                    If oUnit Is Nothing = False Then oUnit.lFleetID = -1
                End If
            Next X
        Catch
        End Try

        Try
            Dim sSQL As String = "DELETE FROM tblUnitGroup WHERE UnitGroupID = " & Me.ObjectID
            Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
        Catch
        End Try

        Dim yMsg(5) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eDeleteFleet).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, 2)
        Me.oOwner.SendPlayerMessage(yMsg, False, AliasingRights.eViewBattleGroups)
    End Sub

    Public Function SetDestSystem(ByVal lSystemID As Int32, ByVal bForceArrival As Boolean) As Boolean

        Try
            Dim bAlreadyMoving As Boolean = False

            'If we don't have a parent object, then we return false
            If Me.oParentObject Is Nothing Then Return False

            'Find the dest system
            Dim oDestSys As SolarSystem = Nothing
            For X As Int32 = 0 To glSystemUB
                If glSystemIdx(X) = lSystemID Then
                    oDestSys = goSystem(X)
                    Exit For
                End If
            Next X

            'Did we find the system? if not, return false
            If oDestSys Is Nothing Then Return False

            If oDestSys.SystemType = 0 Then
                Dim bGood As Boolean = False
                For X As Int32 = 0 To oDestSys.mlWormholeUB
                    If oDestSys.moWormholes(X) Is Nothing = False Then
                        If (oDestSys.moWormholes(X).WormholeFlags And (elWormholeFlag.eSystem1Jumpable Or elWormholeFlag.eSystem2Jumpable Or elWormholeFlag.eSystem2Detectable)) = (elWormholeFlag.eSystem1Jumpable Or elWormholeFlag.eSystem2Jumpable Or elWormholeFlag.eSystem2Detectable) Then
                            bGood = True
                            Exit For
                        End If
                    End If
                Next X
                If bGood = False Then Return False
            ElseIf oDestSys.SystemType = SolarSystem.elSystemType.RespawnSystem Then
                Return False
            End If

            'Our origin location
            Dim lOX As Int32 = Int32.MinValue
            Dim lOY As Int32 = Int32.MinValue
            Dim lOZ As Int32 = Int32.MinValue

            Dim lOriginID As Int32

            If CType(oParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
                With CType(oParentObject, Planet).ParentSystem
                    lOX = .LocX : lOY = .LocY : lOZ = .LocZ
                    lOriginID = .ObjectID
                End With
            ElseIf CType(oParentObject, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
                With CType(oParentObject, SolarSystem)
                    lOX = .LocX : lOY = .LocY : lOZ = .LocZ
                    lOriginID = .ObjectID
                End With
            ElseIf CType(oParentObject, Epica_GUID).ObjTypeID = ObjectType.eGalaxy Then
                'Ok, a special situation, the fleet is already en route somewhere... so figure out the locs of origin and dest
                Dim lTX As Int32 = Int32.MinValue
                Dim lTY As Int32 = Int32.MinValue
                Dim lTZ As Int32 = Int32.MinValue

                lOriginID = Me.lInterSystemOriginID

                bAlreadyMoving = True

                For X As Int32 = 0 To glSystemUB
                    If glSystemIdx(X) = Me.lInterSystemTargetID Then
                        With goSystem(X)
                            lTX = .LocX : lTY = .LocY : lTZ = .LocZ
                            Exit For
                        End With
                    End If
                Next X

                lOX = lOriginX
                lOY = lOriginY
                lOZ = lOriginZ

                'Now, determine our multiplier
                Dim fMult As Single = 1.0F - CSng(InterSystemMoveCyclesRemaining / InterSystemTotalCycles)
                'Now, determine where the fleet is
                lOX += CInt((lTX - lOX) * fMult)
                lOY += CInt((lTY - lOY) * fMult)
                lOZ += CInt((lTZ - lOZ) * fMult)
            Else
                'Where are we?
                Return False
            End If

            'Ok, check that our locs were set
            If lOX = Int32.MinValue OrElse lOY = Int32.MinValue OrElse lOZ = Int32.MinValue Then Return False

            'Ok, now... make sure it is within reasonable range
            Dim fX As Single = (lOX - oDestSys.LocX)
            Dim fY As Single = (lOY - oDestSys.LocY)
            Dim fZ As Single = (lOZ - oDestSys.LocZ)
            fX *= fX
            fY *= fY
            fZ *= fZ

            Dim fDist As Single = CSng(Math.Sqrt(fX + fY + fZ))
            fDist *= 10000000
            'TODO: Determine the player's non gravity well multiplier here
            Dim fPlayerMult As Single = 1.0F
            fDist /= (Me.yInterSystemSpeed * fPlayerMult)
            If fDist = 0 AndAlso bAlreadyMoving = True Then fDist = InterSystemTotalCycles - InterSystemMoveCyclesRemaining
            If fDist > Int32.MaxValue Then Return False

            mlInterSystemTotalCycles = CInt(fDist)
            mlInterSystemMoveCyclesRemaining = mlInterSystemTotalCycles

            'Ok, if we are here... then we can continue
            'Set the new Origin location as this is important for time calculations
            lOriginX = lOX
            lOriginY = lOY
            lOriginZ = lOZ
            Me.lInterSystemOriginID = lOriginID
            Me.lInterSystemTargetID = lSystemID
            Me.lLastInterSystemCycleUpdate = glCurrentCycle

            'Now what to do with the data
            If bAlreadyMoving = True Then
                'Ok, just change the queue event as the fleet is already moving
                If Me.lQueueItemIdx = -1 Then
                    'Or just add us
                    Me.lQueueItemIdx = AddToQueueResult(glCurrentCycle + InterSystemMoveCyclesRemaining, QueueItemType.eSystemToSystemMove, ObjectID, ObjTypeID, lInterSystemTargetID, ObjectType.eSolarSystem, 0, 0, 0, 0)
                Else
                    'change the queue item
                    AlterQueueItem(Me.lQueueItemIdx, glCurrentCycle + InterSystemMoveCyclesRemaining, QueueItemType.eSystemToSystemMove, ObjectID, ObjTypeID, lInterSystemTargetID, ObjectType.eSolarSystem)
                End If
            Else
                'Ok, determine the Regroup point
                'Then, tell all units in this unit group to move to the Regroup point in a special message to the region srv
                SendRegroupMsg()
            End If

            Return True
        Catch

            If bForceArrival = True Then
                'ok, we errored out, most likely, we just need to have them arrive
                If Me.lQueueItemIdx = -1 Then
                    'Or just add us
                    Me.lQueueItemIdx = AddToQueueResult(glCurrentCycle + 9000, QueueItemType.eSystemToSystemMove, ObjectID, ObjTypeID, lInterSystemTargetID, ObjectType.eSolarSystem, 0, 0, 0, 0)
                Else
                    'change the queue item
                    AlterQueueItem(Me.lQueueItemIdx, glCurrentCycle + 9000, QueueItemType.eSystemToSystemMove, ObjectID, ObjTypeID, lInterSystemTargetID, ObjectType.eSolarSystem)
                End If
            End If

            Return False
        End Try

    End Function

	Private Sub SendRegroupMsg()
		'Now, prepare our message
		Dim yMsg() As Byte
		Dim lPos As Int32 = 0

		Dim lEnvirID() As Int32 = Nothing
		Dim iEnvirTypeID() As Int16 = Nothing
		Dim lEnvirUB As Int32 = -1
		Dim lEnvirCnt() As Int32 = Nothing
		Dim oEnvir() As Object = Nothing

		Erase mbUnitRegrouped
		ReDim mbUnitRegrouped(UnitUB)

		For X As Int32 = 0 To UnitUB
			If GroupUnitsIdx(X) <> -1 AndAlso glUnitIdx(GroupUnitsIdx(X)) = GroupUnitsID(X) Then
				Dim oUnit As Unit = goUnit(GroupUnitsIdx(X))
				If oUnit Is Nothing = False AndAlso (oUnit.EntityDef.yChassisType And ChassisType.eSpaceBased) <> 0 Then
					With CType(oUnit.ParentObject, Epica_GUID)
						If .ObjTypeID = ObjectType.ePlanet OrElse .ObjTypeID = ObjectType.eSolarSystem Then
							Dim lIdx As Int32 = -1
							For Y As Int32 = 0 To lEnvirUB
								If lEnvirID(Y) = .ObjectID AndAlso iEnvirTypeID(Y) = .ObjTypeID Then
									lIdx = Y
									Exit For
								End If
							Next Y

							If lIdx = -1 Then
								lEnvirUB += 1
								ReDim Preserve lEnvirID(lEnvirUB)
								ReDim Preserve iEnvirTypeID(lEnvirUB)
								ReDim Preserve lEnvirCnt(lEnvirUB)
								ReDim Preserve oEnvir(lEnvirUB)
								lEnvirID(lEnvirUB) = .ObjectID
								iEnvirTypeID(lEnvirUB) = .ObjTypeID
								lEnvirCnt(lEnvirUB) = 0
								lIdx = lEnvirUB
								oEnvir(lEnvirUB) = oUnit.ParentObject
							End If
							lEnvirCnt(lIdx) += 1
						End If
					End With
				End If
			End If
		Next X

		For lIdx As Int32 = 0 To lEnvirUB
            ReDim yMsg(13 + (lEnvirCnt(lIdx) * 4))

			System.BitConverter.GetBytes(GlobalMessageCode.eSetFleetDest).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(lInterSystemOriginID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(lInterSystemTargetID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(lEnvirCnt(lIdx)).CopyTo(yMsg, lPos) : lPos += 4

			For X As Int32 = 0 To UnitUB
				If GroupUnitsIdx(X) <> -1 AndAlso glUnitIdx(GroupUnitsIdx(X)) = GroupUnitsID(X) Then
					Dim oUnit As Unit = goUnit(GroupUnitsIdx(X))
					If oUnit Is Nothing = False AndAlso (oUnit.EntityDef.yChassisType And ChassisType.eSpaceBased) <> 0 Then
						With CType(oUnit.ParentObject, Epica_GUID)
							If .ObjectID = lEnvirID(lIdx) AndAlso .ObjTypeID = iEnvirTypeID(lIdx) Then
								System.BitConverter.GetBytes(oUnit.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
							End If
						End With
					End If
				End If
			Next X

			'Now, get the envir type
			If CType(oEnvir(lIdx), Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
				CType(oEnvir(lIdx), Planet).oDomain.DomainSocket.SendData(yMsg)
			Else
				'must be a system
				CType(oEnvir(lIdx), SolarSystem).oDomain.DomainSocket.SendData(yMsg)
			End If
		Next lIdx

	End Sub

	Public Sub UnitReady(ByVal lUnitID As Int32)
		For X As Int32 = 0 To UnitUB
			If GroupUnitsIdx(X) <> -1 AndAlso glUnitIdx(GroupUnitsIdx(X)) = lUnitID Then
				mbUnitRegrouped(X) = True
				Exit For
			End If
		Next X

		'Now, go thru and check everything
		Dim bGood As Boolean = True
		For X As Int32 = 0 To UnitUB
			If GroupUnitsIdx(X) <> -1 AndAlso glUnitIdx(GroupUnitsIdx(X)) = GroupUnitsID(X) Then
				If mbUnitRegrouped(X) = False Then
					bGood = False
					Exit For
				End If
			End If
		Next

        If bGood = True Then

            LogEvent(LogEventType.ExtensiveLogging, "Battlegroup entering deep space mvmt: " & Me.ObjectID)

            'Ok, we are good, commence the movement process
            Dim oDomain As DomainServer = Nothing
            Dim lOldParentID As Int32 = CType(oParentObject, Epica_GUID).ObjectID
            Dim iOldParentTypeID As Int16 = CType(oParentObject, Epica_GUID).ObjTypeID
            If iOldParentTypeID = ObjectType.eSolarSystem Then
                oDomain = CType(oParentObject, SolarSystem).oDomain
            ElseIf iOldParentTypeID = ObjectType.ePlanet Then
                oDomain = CType(oParentObject, Planet).oDomain
            End If

            Me.oParentObject = goGalaxy(0)      'TODO: When multiple galaxys are in place this will need work

            'Ok, get total cycles
            InterSystemMoveCyclesRemaining = InterSystemTotalCycles
            Me.lQueueItemIdx = AddToQueueResult(glCurrentCycle + InterSystemMoveCyclesRemaining, QueueItemType.eSystemToSystemMove, ObjectID, ObjTypeID, lInterSystemTargetID, ObjectType.eSolarSystem, 0, 0, 0, 0)

            'Send the update to the player
            SendPlayerSetFleetDestMsg()

            'Now... send a special message to the Region Server for unit removal
            If oDomain Is Nothing = False Then
                Dim yMsg() As Byte
                Dim lPos As Int32 = 0
                Dim lCnt As Int32 = 0
                For X As Int32 = 0 To Me.UnitUB
                    If GroupUnitsIdx(X) <> -1 AndAlso glUnitIdx(GroupUnitsIdx(X)) = GroupUnitsID(X) Then lCnt += 1
                Next X

                ReDim yMsg(15 + (lCnt * 4))

                lPos = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eFleetInterSystemMoving).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lOldParentID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(iOldParentTypeID).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4

                For X As Int32 = 0 To Me.UnitUB
                    If GroupUnitsIdx(X) <> -1 AndAlso glUnitIdx(GroupUnitsIdx(X)) = GroupUnitsID(X) Then
                        LogEvent(LogEventType.ExtensiveLogging, "  BG Mvmt UnitID: " & GroupUnitsID(X) & " for " & Me.ObjectID)
                        System.BitConverter.GetBytes(GroupUnitsID(X)).CopyTo(yMsg, lPos) : lPos += 4
                    End If
                Next X

                oDomain.DomainSocket.SendData(yMsg)
            End If

        End If
	End Sub

	Public Sub SendPlayerSetFleetDestMsg()
        If Me.oOwner.lConnectedPrimaryID > -1 OrElse Me.oOwner.HasOnlineAliases(AliasingRights.eViewBattleGroups) = True Then
            Dim yMsg(35) As Byte
            Dim lPos As Int32 = 0

            'Update the player with the new fleet destination data
            lPos = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetFleetDest).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(Me.lInterSystemTargetID).CopyTo(yMsg, lPos) : lPos += 4

            System.BitConverter.GetBytes(lOriginX).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lOriginY).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lOriginZ).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lInterSystemOriginID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(InterSystemMoveCyclesRemaining).CopyTo(yMsg, lPos) : lPos += 4
            CType(oParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6

            Me.oOwner.SendPlayerMessage(yMsg, False, AliasingRights.eViewBattleGroups)
        End If
	End Sub

    Public Sub UpdateUnitGroupLocation()

        Dim oObj() As Epica_GUID = Nothing
        Dim lUB As Int32 = -1
        Dim lCnt() As Int32 = Nothing

        For X As Int32 = 0 To UnitUB
            If GroupUnitsIdx(X) > -1 Then
                Dim lIdx As Int32 = GroupUnitsIdx(X)
                If lIdx > -1 Then
                    Dim lID As Int32 = GroupUnitsID(X)
                    If lID = glUnitIdx(lIdx) Then
                        Dim oUnit As Unit = goUnit(lIdx)
                        If oUnit Is Nothing = False Then
                            If oUnit.ParentObject Is Nothing = False Then

                                Try
                                    Dim oTmp As Epica_GUID = CType(oUnit.ParentObject, Epica_GUID)
                                    If oTmp Is Nothing = False Then
                                        Dim lTmpIdx As Int32 = -1
                                        For Y As Int32 = 0 To lUB
                                            If oObj(Y).ObjectID = oTmp.ObjectID AndAlso oObj(Y).ObjTypeID = oTmp.ObjTypeID Then
                                                lTmpIdx = Y
                                                Exit For
                                            End If
                                        Next Y
                                        If lTmpIdx = -1 Then
                                            lUB += 1
                                            ReDim Preserve oObj(lUB)
                                            ReDim Preserve lCnt(lUB)
                                            oObj(lUB) = oTmp
                                            lCnt(lUB) = 0
                                            lTmpIdx = lUB
                                        End If

                                        lCnt(lTmpIdx) += 1
                                    End If
                                Catch
                                End Try
                            End If
                        End If
                    End If
                End If
            End If
        Next X

        Dim lMax As Int32 = 0
        Dim oMaxObj As Epica_GUID = Nothing

        For X As Int32 = 0 To lUB
            If lCnt(X) > lMax Then
                lMax = lCnt(X)
                oMaxObj = oObj(X)
            End If
        Next X

        If oMaxObj Is Nothing Then
            'group is dead
            Me.DeleteMe()
            For X As Int32 = 0 To glUnitGroupUB
                If glUnitGroupIdx(X) = Me.ObjectID Then
                    glUnitGroupIdx(X) = -1
                    goUnitGroup(X) = Nothing
                    Exit For
                End If
            Next X
        Else

            Dim bSendMsg As Boolean = False
            If Me.oParentObject Is Nothing = False Then
                Dim oTmp As Epica_GUID = CType(Me.oParentObject, Epica_GUID)
                If oTmp.ObjTypeID = ObjectType.eGalaxy Then
                    If Me.lInterSystemTargetID <> -1 AndAlso Me.InterSystemMoveCyclesRemaining > 0 Then
                        Return
                    End If
                End If
                If oTmp.ObjectID <> oMaxObj.ObjectID OrElse oTmp.ObjTypeID <> oMaxObj.ObjTypeID Then
                    Me.oParentObject = oMaxObj
                    bSendMsg = True
                End If
            Else
                Me.oParentObject = oMaxObj
                bSendMsg = True
            End If

            If bSendMsg = True Then
                Dim yMsg() As Byte = goMsgSys.GetAddObjectMessage(Me, GlobalMessageCode.eAddObjectCommand)
                If yMsg Is Nothing = False Then Me.oOwner.SendPlayerMessage(yMsg, False, AliasingRights.eViewBattleGroups)
            End If

        End If

    End Sub

	Public Function TestUnitGroupArrival(ByVal lTargetSystemID As Int32) As Int32
		If Me.lInterSystemTargetID <> lTargetSystemID Then Return -1 '-1 ensures the event is removed

		'Any positive number means to delay the arrival longer by the result
        If Me.InterSystemMoveCyclesRemaining > 0 Then Return Me.InterSystemMoveCyclesRemaining

		'Ok, if we are here, then all is well... let's determine where the fleet will enter... get our dest system
		Dim oDestSys As SolarSystem = Nothing
		For X As Int32 = 0 To glSystemUB
			If glSystemIdx(X) = lTargetSystemID Then
				oDestSys = goSystem(X)
				Exit For
			End If
		Next X
		If oDestSys Is Nothing Then
			LogEvent(LogEventType.CriticalError, "Unit Group " & Me.ObjectID & " arrived but could not find target system!")
			Return -1 'ensures the removal
		End If

		'Set my parent to the new location
		Me.oParentObject = oDestSys
		'set my transit values
		Me.lInterSystemTargetID = -1
		Me.lInterSystemOriginID = -1


		Dim fAngle As Single
		Dim fX1 As Single = oDestSys.LocX
		Dim fY1 As Single = oDestSys.LocZ
		Dim fX2 As Single = lOriginX
		Dim fY2 As Single = lOriginZ

		'NOTE: This is essentially LineAngleDegrees but this app doesn't have it so I didn't add it, if we use it again, just copy this out of here
		Dim fDX As Single = fX2 - fX1
		Dim fDY As Single = fY2 - fY1
		If fDX = 0 Then	'vertical
			If fDY < 0 Then
				fAngle = Math.PI / 2.0F
			Else : fAngle = Math.PI * 1.5F
			End If
		ElseIf fDY = 0 Then	'horizontal
			If fDX < 0 Then
				fAngle = CSng(Math.PI)
			Else : fAngle = 0.0F
			End If
		Else
			fAngle = CSng(Math.Atan(Math.Abs(fDY / fDX)))
			If fDX > -1 AndAlso fDY > -1 Then
				fAngle = CSng((Math.PI) * 2.0F) - fAngle
			ElseIf fDX < 0 AndAlso fDY > -1 Then
				fAngle = CSng(Math.PI) + fAngle
			ElseIf fDX < 0 AndAlso fDY < 0 Then
				fAngle = CSng(Math.PI) - fAngle
			End If
		End If
		'Adjust the angle to degrees
		fAngle *= 57.2957764F			'gdDegreePerRad
		'Adjust for the weird stuff up above
		fAngle = 360.0F - fAngle
		'End of Line Angle Degrees

		'Now, we put the values in from Dest to Origin, so the angle is already good... let's make our values
		'Reuse some older values
		fX1 = 5000000
		fY1 = 0

		'NOTE: This is RotatePoint, but this app doesn't have it so I didn't add it, if we do, just copy this code
		Dim fRads As Single = fAngle * (CSng(Math.PI) / 180.0F)
		fDX = fX1	' was originally lEndX - lAxisX
		fDY = fY1	' was originally lEndY - lAxisY
		'The following two formulae should be lAxisX or lAxisY + the value
		fX1 = ((fDX * CSng(Math.Cos(fRads))) + (fDY * CSng(Math.Sin(fRads))))
		fY1 = -((fDX * CSng(Math.Sin(fRads))) + (fDY * CSng(Math.Cos(fRads))))
		'End of RotatePoint

		'fX1 and fY1 have our entry point X,Z in a circle... we now need to move that point to the edge
		'Gonna reuse X2 and Y2 as the ABSOLUTE values
		fX2 = Math.Abs(fX1) : fY2 = Math.Abs(fY1)
		If fX2 <> 5000000 AndAlso fY2 <> 5000000 Then
			If fX2 > fY2 Then
				'make X1 5000000 and adjust y
				fY1 *= (5000000 / fX2)		'use the absolute value of X here
				If fX1 < 0 Then fX1 = -5000000 Else fX1 = 5000000
			Else
				'Make Y1 5000000 and adjust x
				fX1 *= (5000000 / fY2)		'use the absolute value of y here
				If fY1 < 0 Then fY1 = -5000000 Else fY1 = 5000000
			End If
		End If

        fX1 = oDestSys.FleetJumpPointX
        fY1 = oDestSys.FleetJumpPointZ

		'Ok, fX1 and fY1 have our final coordinates
		Dim lLocX As Int32 = CInt(fX1)
		Dim lLocZ As Int32 = CInt(fY1)
		Dim iLocA As Int16 = CShort((fAngle * 10) + 1800S)
		If iLocA > 3600S Then iLocA -= 3600S

        LogEvent(LogEventType.ExtensiveLogging, "Battlegroup arrived. ID: " & Me.ObjectID & ". OwnerID: " & Me.oOwner.ObjectID)

        'Now, add all of the units to that system...
        For X As Int32 = 0 To UnitUB
            If GroupUnitsIdx(X) <> -1 AndAlso glUnitIdx(GroupUnitsIdx(X)) <> -1 AndAlso glUnitIdx(GroupUnitsIdx(X)) = GroupUnitsID(X) Then
                Dim oUnit As Unit = goUnit(GroupUnitsIdx(X))
                If oUnit Is Nothing = False Then

                    LogEvent(LogEventType.ExtensiveLogging, "  Unit Arrival: " & oUnit.ObjectID & " for BG: " & Me.ObjectID)

                    With oUnit
                        .LocX = lLocX
                        .LocZ = lLocZ
                        .LocAngle = iLocA
                        .ParentObject = oDestSys
                        .DataChanged()
                        .SaveObject()
                    End With

                    If oDestSys.InMyDomain = False Then
                        'goMsgSys.SendPassThruMsg(goMsgSys.GetAddObjectMessage(oUnit, GlobalMessageCode.eAddObjectCommand_CE), -1, oDestSys.ObjectID, oDestSys.ObjTypeID)
                        'goMsgSys.SendMsgToOperator(goMsgSys.GetAddObjectMessage(oUnit, GlobalMessageCode.eEntityChangingPrimary))
                        oUnit.ProcessChangePrimaryServers(oDestSys)
                    Else
                        oDestSys.oDomain.DomainSocket.SendData(goMsgSys.GetAddObjectMessage(oUnit, GlobalMessageCode.eAddObjectCommand_CE))
                    End If


                End If
            End If
        Next X

        If Me.lInterSystemOriginID > -1 Then
            Dim oOriginSys As SolarSystem = GetEpicaSystem(Me.lInterSystemOriginID)
            If oOriginSys Is Nothing = False AndAlso oOriginSys.SystemType = SolarSystem.elSystemType.SpawnSystem Then
                For X As Int32 = 0 To oOriginSys.mlWormholeUB
                    Dim oWH As Wormhole = oOriginSys.moWormholes(X)
                    If oWH Is Nothing = False Then
                        If oWH.System1 Is Nothing = False AndAlso oWH.System1.ObjectID = oOriginSys.ObjectID Then
                            oWH.WormholeFlags = oWH.WormholeFlags Or elWormholeFlag.eSystem1Detectable Or elWormholeFlag.eSystem2Detectable
                            oWH.QueueMeToSave()
                        End If
                    End If
                Next X
            End If
        End If

        'Now, update the player with this update to this fleet
        If Me.oOwner.lConnectedPrimaryID > -1 OrElse Me.oOwner.HasOnlineAliases(AliasingRights.eViewBattleGroups) = True Then
            Me.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(Me, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewBattleGroups)
        End If

        Return -1       'ensure event removal
	End Function

    Public Function CanAlterComposition() As Boolean
        'If the unit group is in the galaxy object, its contents cannot be modified
        If CType(Me.oParentObject, Epica_GUID).ObjTypeID = ObjectType.eGalaxy Then Return False
        Return True
    End Function

    Public Function GetElementCount() As Int32
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To UnitUB
            If GroupUnitsIdx(X) <> -1 Then
                If glUnitIdx(GroupUnitsIdx(X)) = GroupUnitsID(X) AndAlso GroupUnitsID(X) <> -1 Then
                    lCnt += 1
                Else
                    GroupUnitsIdx(X) = -1
                    GroupUnitsID(X) = -1
                End If
            End If
        Next X
        Return lCnt
    End Function

    Public Sub PopulateAvailablePersonnel(ByRef lColAvail As Int32, ByRef lEnlAvail As Int32, ByRef lOffAvail As Int32, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lCallerUnitID As Int32)
        For X As Int32 = 0 To UnitUB
            Try
				If GroupUnitsIdx(X) <> -1 AndAlso glUnitIdx(GroupUnitsIdx(X)) = GroupUnitsID(X) Then
					Dim oUnit As Unit = goUnit(GroupUnitsIdx(X))
					If oUnit Is Nothing Then Continue For
					With oUnit
						If .ObjectID <> lCallerUnitID Then
							If CType(.ParentObject, Epica_GUID).ObjectID = lEnvirID AndAlso CType(.ParentObject, Epica_GUID).ObjTypeID = iEnvirTypeID Then
								If (.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
									For Y As Int32 = 0 To .lCargoUB
										If .lCargoIdx(Y) <> -1 Then
											Select Case .oCargoContents(Y).ObjTypeID
												Case ObjectType.eColonists
                                                    lColAvail += .oCargoContents(Y).ObjectID + 1
												Case ObjectType.eEnlisted
                                                    lEnlAvail += .oCargoContents(Y).ObjectID + 1
												Case ObjectType.eOfficers
                                                    lOffAvail += .oCargoContents(Y).ObjectID + 1
											End Select
										End If
									Next Y
								End If
							End If
						End If
					End With
				End If
            Catch
                'do nothing
            End Try
        Next X
    End Sub

    Public Sub ReduceUnitGroupsPersonnel(ByVal lCol As Int32, ByVal lEnl As Int32, ByVal lOff As Int32, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16, ByVal lCallerUnitID As Int32)
        For X As Int32 = 0 To UnitUB
            Try
				If GroupUnitsIdx(X) <> -1 AndAlso glUnitIdx(GroupUnitsIdx(X)) = GroupUnitsID(X) Then
					Dim oUnit As Unit = goUnit(GroupUnitsIdx(X))
					If oUnit Is Nothing Then Continue For
					With oUnit
						If .ObjectID <> lCallerUnitID Then
							If CType(.ParentObject, Epica_GUID).ObjectID = lEnvirID AndAlso CType(.ParentObject, Epica_GUID).ObjTypeID = iEnvirTypeID Then
								If (.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
									For Y As Int32 = 0 To .lCargoUB
										If .lCargoIdx(Y) <> -1 Then
											Select Case .oCargoContents(Y).ObjTypeID
												Case ObjectType.eColonists
													Dim lTempVal As Int32 = Math.Min(lCol, .oCargoContents(Y).ObjectID)
													.oCargoContents(Y).ObjectID -= lTempVal
													If .oCargoContents(Y).ObjectID = 0 Then .lCargoIdx(Y) = -1
													lCol -= lTempVal
												Case ObjectType.eEnlisted
													Dim lTempVal As Int32 = Math.Min(lEnl, .oCargoContents(Y).ObjectID)
													.oCargoContents(Y).ObjectID -= lTempVal
													If .oCargoContents(Y).ObjectID = 0 Then .lCargoIdx(Y) = -1
													lEnl -= lTempVal
												Case ObjectType.eOfficers
													Dim lTempVal As Int32 = Math.Min(lOff, .oCargoContents(Y).ObjectID)
													.oCargoContents(Y).ObjectID -= lTempVal
													If .oCargoContents(Y).ObjectID = 0 Then .lCargoIdx(Y) = -1
													lOff -= lTempVal
											End Select
										End If
									Next Y

									If lCol = 0 AndAlso lEnl = 0 AndAlso lOff = 0 Then Return
								End If
							End If
						End If
					End With
				End If
            Catch
                'do nothing
            End Try
        Next X
    End Sub

    Public Function GetUnitIdx(ByVal lIndex As Int32) As Int32
        If lIndex > -1 AndAlso lIndex < UnitUB + 1 Then
            If GroupUnitsIdx(lIndex) <> -1 AndAlso glUnitIdx(GroupUnitsIdx(lIndex)) = GroupUnitsID(lIndex) Then Return GroupUnitsIdx(lIndex)
            Return -1
        End If
    End Function

    Public Function AdjustComponentCacheForUnitGroup(ByVal lComponentID As Int32, ByVal iComponentTypeID As Int16, ByVal lComponentOwnerID As Int32, ByVal lCacheQty As Int32, ByVal lEnvirID As Int32, ByVal iEnvirTypeID As Int16) As Int32
        For X As Int32 = 0 To UnitUB
            If GroupUnitsIdx(X) > -1 AndAlso GroupUnitsIdx(X) <= glUnitUB Then
                If glUnitIdx(GroupUnitsIdx(X)) = GroupUnitsID(X) Then
                    Dim oUnit As Unit = goUnit(GroupUnitsIdx(X))
                    If oUnit Is Nothing = False Then
                        If (oUnit.CurrentStatus And elUnitStatus.eCargoBayOperational) <> 0 AndAlso oUnit.Cargo_Cap > 0 Then
                            If oUnit.ParentObject Is Nothing = False Then
                                With CType(oUnit.ParentObject, Epica_GUID)
                                    If .ObjectID <> lEnvirID OrElse .ObjTypeID <> iEnvirTypeID Then Continue For
                                End With
                            End If
                            Dim lQty As Int32 = Math.Min(oUnit.Cargo_Cap, lCacheQty)
                            oUnit.AddComponentCacheToCargo(lComponentID, iComponentTypeID, lQty, lComponentOwnerID)
                            If lQty <> lCacheQty Then lCacheQty -= lQty Else Return 0
                        End If
                    End If
                End If
            End If
        Next X
        Return lCacheQty
    End Function
End Class
