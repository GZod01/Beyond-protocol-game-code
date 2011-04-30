Public Class Facility
    Inherits Epica_Entity

    Private Const ml_MAX_FACILITY_SIZE As Int32 = 25

    Public ParentColony As Colony           'parent colony supplies this facility with resources (power, minerals, etc...)

    Public ReadOnly Property MaxWorkers() As Int32
        Get
            Return Me.EntityDef.WorkerFactor
        End Get
    End Property

    Public Production() As EntityProduction
    Public ProductionIdx() As Int32
    Public ProductionUB As Int32 = -1
    Public ServerIndex As Int32 = -1    'index in the global array
    Public EntityDef As FacilityDef

    Public ColonyArrayIndex As Int32 = -1

    Public AutoLaunch As Boolean = False

    Public lPirate_Counter As Int32 = 0         'number of items produced this production group
    Public lPirate_Counter_Max As Int32 = 0     'max number of items produced for this group
    Public lPirate_For_PlayerID As Int32 = 0
    Public lPirate_Items() As Int32             'IDs for the items built by this facility for this group
    Public lPirate_ItemUB As Int32 = -1

    Public bExcludeFromColonyQueue As Boolean = False

    Private mlTradePostSellSlotsUsed As Int32 = 0
    Public Property lTradePostSellSlotsUsed() As Int32
        Get
            Return mlTradePostSellSlotsUsed
        End Get
        Set(ByVal value As Int32)
            mlTradePostSellSlotsUsed = value
            If mlTradePostSellSlotsUsed < 0 Then
                mlTradePostSellSlotsUsed = 0
            End If
        End Set
    End Property
    Private mlTradePostBuySlotsUsed As Int32 = 0
    Public Property lTradePostBuySlotsUsed() As Int32
        Get
            Return mlTradePostBuySlotsUsed
        End Get
        Set(ByVal value As Int32)
            mlTradePostBuySlotsUsed = value
            If mlTradePostBuySlotsUsed < 0 Then
                mlTradePostBuySlotsUsed = 0
            End If
        End Set
    End Property


    Public lReinforcingUnitGroupID As Int32 = -1        'who I am reinforcing

	Private mySendString() As Byte

    Public PreviousActive As Boolean

    Public oMiningBid As MineBuyOrderManager = Nothing

    Public Function SetActive(ByVal bNewVal As Boolean) As Boolean
        'Ok, if the new Val is True and the current value is False, we need to check for power
        'If the new val is false and the current value is true, we need to check for power
        Dim bActive As Boolean = Active
        Dim lTemp As Int32
        Dim bResult As Boolean = True

        PreviousActive = bActive

        If bActive <> bNewVal Then
            If bActive = True Then
                'Ok, this facility is already on... if it is, then determine if it is a power producer
                If PowerConsumption < 0 Then
                    'Yes... if the colony didn't have this power... would the need overcome the created?
                    lTemp = ParentColony.PowerGeneration + PowerConsumption      'powergeneration is +, powerconsumption is -
                    If ParentColony.PowerConsumption > lTemp Then
                        'don't allow the switch off as it would cause other facilities to become inoperable
                        bResult = False
                    Else
                        'Yeah, you can switch it off
						CurrentStatus = CurrentStatus Xor elUnitStatus.eFacilityPowered 'CurrentStatus -= elUnitStatus.eFacilityPowered

                        If Me.ParentColony Is Nothing = False Then Me.ParentColony.UpdateAllValues(Me.ColonyArrayIndex)

                        'ParentColony.AddNonWorkers(Me.CurrentWorkers)
                        'Me.CurrentWorkers = 0
                        goMsgSys.SendDomainEntityStatus(Me, -elUnitStatus.eFacilityPowered)
                        bResult = True
                    End If
                Else
					If (CurrentStatus And elUnitStatus.eFacilityPowered) <> 0 Then CurrentStatus = CurrentStatus Xor elUnitStatus.eFacilityPowered 'CurrentStatus -= elUnitStatus.eFacilityPowered
                    If Me.ParentColony Is Nothing = False Then Me.ParentColony.UpdateAllValues(Me.ColonyArrayIndex)
                    'ParentColony.AddNonWorkers(Me.CurrentWorkers)
                    'Me.CurrentWorkers = 0       'this will update our power consumption
                    goMsgSys.SendDomainEntityStatus(Me, -elUnitStatus.eFacilityPowered)
                    bResult = True

                    If Me.CurrentProduction Is Nothing = False Then
                        Dim oTech As Epica_Tech = Me.Owner.GetTech(Me.CurrentProduction.ProductionID, Me.CurrentProduction.ProductionTypeID)
                        If oTech Is Nothing = False Then
                            oTech.RemoveResearcher(Me.ObjectID)
                        End If
                        oTech = Nothing
                    End If
                    Me.CurrentProduction = Nothing
                End If
            Else
                'Ok, the facility is not on but is being switched on... so, set our flag to on, for now... but without event
                CurrentStatus = CurrentStatus Or elUnitStatus.eFacilityPowered

                If Me.ParentColony.PowerConsumption + EntityDef.PowerFactor <= Me.ParentColony.PowerGeneration Then
                    'Yes, we can turn on the facility
                    If Me.ParentColony Is Nothing = False Then Me.ParentColony.UpdateAllValues(Me.ColonyArrayIndex)
                    goMsgSys.SendDomainEntityStatus(Me, elUnitStatus.eFacilityPowered)
                    bResult = True
                Else
                    'Not enough power to turn on facility
					CurrentStatus = CurrentStatus Xor elUnitStatus.eFacilityPowered 'CurrentStatus -= elUnitStatus.eFacilityPowered
                    bResult = False
                End If

            End If
        End If

        If Me.Owner Is Nothing = False Then
            If Me.Owner.lConnectedPrimaryID > -1 OrElse Me.Owner.HasOnlineAliases(AliasingRights.eViewColonyStats Or AliasingRights.eAlterAutoLaunchPower) = True Then
                Dim yResp(11) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityStatus).CopyTo(yResp, 0)
                GetGUIDAsString.CopyTo(yResp, 2)
                If (CurrentStatus And elUnitStatus.eFacilityPowered) = 0 Then
                    System.BitConverter.GetBytes(-elUnitStatus.eFacilityPowered).CopyTo(yResp, 8)
                Else
                    System.BitConverter.GetBytes(elUnitStatus.eFacilityPowered).CopyTo(yResp, 8)
                End If

                Me.Owner.SendPlayerMessage(yResp, False, AliasingRights.eViewColonyStats Or AliasingRights.eAlterAutoLaunchPower)
            End If
		End If

		Return bResult

	End Function

	Public ReadOnly Property Active() As Boolean
		Get
			Return ((CurrentStatus And elUnitStatus.eFacilityPowered) <> 0)
		End Get
	End Property

	Public Function GetObjAsString() As Byte()
		If mbStringReady = False Then
			'NOTE: because the object we are returning is associated to an EntityDefID... it is assumed that the request
			'  sender has the entity def object. 
			Dim lAmmoLen As Int32
			If lCurrentAmmo Is Nothing Then lAmmoLen = 0 Else lAmmoLen = 8 * lCurrentAmmo.Length

            ReDim mySendString(107 + lAmmoLen)  '101
			Dim lPos As Int32
			Dim X As Int32

			GetGUIDAsString.CopyTo(mySendString, 0)

			'NOTE: this could be very dangerous... the IDX array stores -1 for objects that no longer exist...
			If ParentObject Is Nothing Then
				System.BitConverter.GetBytes(CInt(-1)).CopyTo(mySendString, 6)
				System.BitConverter.GetBytes(CShort(-1)).CopyTo(mySendString, 10)
			Else
				CType(ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(mySendString, 6)
			End If
			System.BitConverter.GetBytes(EntityDef.ObjectID).CopyTo(mySendString, 12)
			EntityName.CopyTo(mySendString, 16)
			If Owner Is Nothing Then
				System.BitConverter.GetBytes(CInt(-1)).CopyTo(mySendString, 36)
			Else
				System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(mySendString, 36)
			End If
			System.BitConverter.GetBytes(LocX).CopyTo(mySendString, 40)
			System.BitConverter.GetBytes(LocZ).CopyTo(mySendString, 44)

			System.BitConverter.GetBytes(LocAngle).CopyTo(mySendString, 48)

			System.BitConverter.GetBytes(Structure_HP).CopyTo(mySendString, 50)
			System.BitConverter.GetBytes(Fuel_Cap).CopyTo(mySendString, 54)
			System.BitConverter.GetBytes(Shield_HP).CopyTo(mySendString, 58)
			mySendString(62) = ExpLevel

			System.BitConverter.GetBytes(CInt(-1)).CopyTo(mySendString, 63)

			System.BitConverter.GetBytes(CurrentStatus).CopyTo(mySendString, 67)

			System.BitConverter.GetBytes(Q1_HP).CopyTo(mySendString, 71)
			System.BitConverter.GetBytes(Q2_HP).CopyTo(mySendString, 75)
			System.BitConverter.GetBytes(Q3_HP).CopyTo(mySendString, 79)
			System.BitConverter.GetBytes(Q4_HP).CopyTo(mySendString, 83)

			System.BitConverter.GetBytes(CInt(0)).CopyTo(mySendString, 87)		'TODO: No longer used
			If ParentColony Is Nothing = False Then System.BitConverter.GetBytes(ParentColony.ObjectID).CopyTo(mySendString, 91) Else System.BitConverter.GetBytes(-1I).CopyTo(mySendString, 91)
			System.BitConverter.GetBytes(CShort(0)).CopyTo(mySendString, 95)	'TODO: No longer used, width
			System.BitConverter.GetBytes(CShort(0)).CopyTo(mySendString, 97)	'TODO: No longer used, height
			mySendString(99) = yProductionType

			System.BitConverter.GetBytes(iCombatTactics).CopyTo(mySendString, 100)
            System.BitConverter.GetBytes(iTargetingTactics).CopyTo(mySendString, 104)

			If lCurrentAmmo Is Nothing = False Then
                System.BitConverter.GetBytes(CShort(lCurrentAmmo.Length)).CopyTo(mySendString, 106)
                lPos = 108

				For X = 0 To lCurrentAmmo.Length - 1
					System.BitConverter.GetBytes(EntityDef.WeaponDefs(X).ObjectID).CopyTo(mySendString, lPos) : lPos += 4
					System.BitConverter.GetBytes(lCurrentAmmo(X)).CopyTo(mySendString, lPos) : lPos += 4
				Next X
            Else : System.BitConverter.GetBytes(CShort(0)).CopyTo(mySendString, 106)
			End If

			mbStringReady = True
		End If
		Return mySendString
	End Function

	Public Sub New()
		ReDim lCurrentAmmo(-1)
	End Sub

	Public Function AddProduction(ByVal lProdID As Int32, ByVal iProdTypeID As Int16, ByVal yOrderNumber As Byte, ByVal lQuantity As Int32, ByVal blPointsProduced As Int64) As Boolean
		Dim X As Int32
		Dim lIdx As Int32
		Dim bFound As Boolean = False
		Dim lLowestIdx As Int32 = -1

		'Ok, first, let's make sure we are doing the right thing... 
		If lProdID < 0 Then				'negative prodid indicates clearing queue
			ClearQueue()
			If Me.CurrentProduction Is Nothing = False Then
				If Me.yProductionType = ProductionType.eResearch Then
					Dim oTech As Epica_Tech = Me.Owner.GetTech(Me.CurrentProduction.ProductionID, Me.CurrentProduction.ProductionTypeID)
					If oTech Is Nothing = False Then
						Me.CurrentProduction = Nothing
						Me.bProducing = False
						oTech.RemoveResearcher(Me.ObjectID)

						If Me.ParentColony Is Nothing = False Then
							AddToQueue(glCurrentCycle + 30, QueueItemType.eCheckColonyResearchQueue, Me.ParentColony.ObjectID, 0, 0, 0, 0, 0, 0, 0)
						End If
					End If
				End If
			End If
			Return True
		ElseIf iProdTypeID < 0 Then		'negative prodtypeid indicatings removing queue item
			'in this case, the Quantity is the listcount which should indicate our production order...
			lIdx = lQuantity
			'the prod type id is the prod type id but it is negative
			iProdTypeID = Math.Abs(iProdTypeID)

			For X = 0 To ProductionUB
				If ProductionIdx(X) <> -1 Then lIdx -= 1

				If ProductionIdx(X) = lProdID AndAlso Production(X).ProductionTypeID = iProdTypeID Then
					'see if we are near our quantity...
					If Math.Abs(lIdx) < 2 Then
						'Ok, set this one
						ProductionIdx(X) = -1
						Return True
					End If
				End If
			Next X
			Return False
		ElseIf lQuantity < 0 Then
			'Ok, we are removing queue items...
			lQuantity = Math.Abs(lQuantity)
			If CurrentProduction Is Nothing = False Then
				If CurrentProduction.ProductionID = lProdID AndAlso CurrentProduction.ProductionTypeID = iProdTypeID Then
					If CurrentProduction.lProdCount > 1 AndAlso CurrentProduction.lProdCount > lQuantity Then
						lIdx = Math.Min(lQuantity, CurrentProduction.lProdCount - 1)

						lQuantity -= lIdx
						CurrentProduction.lProdCount -= lIdx
					Else
						'Refund credits only
						If CurrentProduction.ProdCost Is Nothing = False Then Owner.blCredits += CurrentProduction.ProdCost.CreditCost
						CurrentProduction.lProdCount = 0
                        If bProducing = True Then GetNextProduction(False)
						Return True
					End If
				End If
			End If

			For X = 0 To ProductionUB
				If ProductionIdx(X) = lProdID AndAlso Production(X).ProductionTypeID = iProdTypeID Then
					lIdx = Math.Min(lQuantity, Production(X).lProdCount)
					lQuantity -= lIdx
					Production(X).lProdCount -= lIdx
					If Production(X).lProdCount = 0 Then ProductionIdx(X) = -1
					If lQuantity = 0 Then Exit For
				End If
			Next X

			Return True
		End If


		lIdx = -1

		'If CurrentProduction Is Nothing = False Then
		'    If CurrentProduction.ProductionID = lProdID Then
		'        If CurrentProduction.ProductionTypeID = iProdTypeID Then
		'            'currently building it
		'            If yOrderNumber = 254 Then
		'                bFound = True
		'                CurrentProduction.lProdCount += lQuantity
		'            Else
		'                'Oh, not good...
		'                'TODO: We need to tell the current production to move elsewhere
		'            End If
		'        End If
		'    End If
		'End If

		If bFound = False Then

			For X = ProductionUB To 0 Step -1
				If ProductionIdx(X) <> -1 Then
					If ProductionIdx(X) = lProdID AndAlso Production(X).ProductionTypeID = iProdTypeID Then
						'Already there... increment the quantity
						bFound = True
						Production(X).lProdCount += lQuantity
					End If

					Exit For
				End If
			Next

            'If bFound = False Then
            '	For X = ProductionUB To 0 Step -1
            '		If ProductionIdx(X) = -1 Then
            '			lIdx = X
            '			Exit For
            '		End If
            '	Next X
            'End If
            If bFound = False Then
                Dim lLastIdx As Int32 = -1
                For X = ProductionUB To 0 Step -1
                    If ProductionIdx(X) <> -1 AndAlso lLastIdx <> -1 Then
                        lIdx = lLastIdx
                        Exit For
                    ElseIf ProductionIdx(X) = -1 Then ' AndAlso lLastIdx = -1 Then
                        lLastIdx = X
                    End If
                    'If ProductionIdx(X) = -1 Then
                    '    lIdx = X
                    '    Exit For
                    'End If
                Next X
                If lIdx = -1 Then lIdx = lLastIdx
            End If

			'For X = 0 To ProductionUB
			'    If ProductionIdx(X) = lProdID Then
			'        If Production(X).ProductionTypeID = iProdTypeID Then
			'            'already there... check the order number
			'            If (Production(X).OrderNumber = yOrderNumber) Or yOrderNumber = 254 Then
			'                'nothing else to do
			'                bFound = True
			'                Production(X).lProdCount += lQuantity
			'                Exit For
			'            Else
			'                For Y = 0 To ProductionUB
			'                    If ProductionIdx(Y) <> -1 AndAlso (ProductionIdx(Y) <> lProdID OrElse Production(Y).ProductionTypeID <> iProdTypeID) Then
			'                        If Production(Y).OrderNumber >= yOrderNumber Then
			'                            Production(Y).OrderNumber += CByte(1)
			'                        End If
			'                    End If
			'                Next Y
			'                Production(X).OrderNumber = yOrderNumber
			'                bFound = True
			'                Production(X).lProdCount += lQuantity
			'                Exit For
			'            End If
			'        End If
			'    ElseIf lIdx = -1 AndAlso ProductionIdx(X) = -1 Then
			'        lIdx = X
			'    End If
			'Next X
		End If

		'Ok, if we're here, then we have a new production, if lIDX = -1 then we need to make space
		If bFound = False Then
			If lIdx = -1 Then
				ProductionUB += 1
				ReDim Preserve Production(ProductionUB)
				ReDim Preserve ProductionIdx(ProductionUB)
				Production(ProductionUB) = New EntityProduction()
				lIdx = ProductionUB
			Else
				Production(lIdx) = Nothing
				Production(lIdx) = New EntityProduction()
			End If

			ProductionIdx(lIdx) = -1
			If FillProductionItemValues(Production(lIdx), lProdID, iProdTypeID, yOrderNumber, lQuantity) = True Then
				ProductionIdx(lIdx) = lProdID
			Else : Return False
			End If
			Production(lIdx).PointsProduced = blPointsProduced

			If blPointsProduced <> 0 Then
				Production(lIdx).bPaidFor = True
			End If

		End If

		lLowestIdx = ReSortProduction()

		'regardless, we are now producing
		If bProducing = False Then
            If GetNextProduction(False) = False Then Return False
            'MSC - 06/20/08 - added the MAX here to see if that perma fixes the tech research resetting to 0
            If Me.CurrentProduction Is Nothing = False Then Me.CurrentProduction.PointsProduced = Math.Max(Me.CurrentProduction.PointsProduced, blPointsProduced)
			RecalcProduction()
		End If

		Return True
	End Function

	Private Function ReSortProduction() As Int32
		Dim lCnt As Int32
		Dim yTemp1 As Byte
		Dim yTemp2 As Byte
		Dim lLowestIdx As Int32 = -1
		Dim lIdx As Int32
		Dim Y As Int32
		Dim X As Int32

		'Do our sort... first, get our count
		lCnt = 0
		For X = 0 To ProductionUB
			If ProductionIdx(X) <> -1 Then lCnt += 1
		Next X
		'set our current minimum value... NOTE: a ordernumber should never be allowed to be set to 0
		yTemp1 = 0
		'Now, loop through our count
		For X = 0 To lCnt - 1
			'clear our Index for this loop
			lIdx = -1
			'set our MAX value
			yTemp2 = Byte.MaxValue
			'Now, loop through the production array
			For Y = 0 To ProductionUB
				If ProductionIdx(Y) <> -1 Then
					If Production(Y).OrderNumber < yTemp2 AndAlso ((Production(Y).OrderNumber < yTemp1) Or yTemp1 = 0) Then
						'store this value away
						lIdx = Y
						'but is there one smaller??
						yTemp2 = Production(Y).OrderNumber
					End If
				End If
			Next Y

			'Now, did we find one?
			If lIdx <> -1 Then
				yTemp1 = yTemp2
				Production(lIdx).OrderNumber = CByte(X + 1)

				If X = 0 Then lLowestIdx = lIdx
			Else
				Exit For
			End If
		Next X
	End Function

	Public Sub RemoveProduction(ByVal lObjID As Int32, ByVal yOrderNumber As Byte)
		Dim X As Int32

		For X = 0 To ProductionUB
			If ProductionIdx(X) = lObjID AndAlso Production(X).OrderNumber = yOrderNumber Then
				Production(X) = Nothing
				ProductionIdx(X) = -1
				Exit For
			End If
		Next X

        GetNextProduction(False)

        If Me.Owner.lConnectedPrimaryID > -1 OrElse Me.Owner.HasOnlineAliases(AliasingRights.eAddProduction Or AliasingRights.eCancelProduction) = True Then
            Dim yMsg(7) As Byte
            Me.GetGUIDAsString.CopyTo(yMsg, 2)
            Dim yResp() As Byte = goMsgSys.HandleGetEntityProduction(yMsg)
            If yResp Is Nothing = False Then Me.Owner.CrossPrimarySafeSendMsg(yResp) ' goMsgSys.HandleGetEntityProduction(Me.Owner.oSocket, yMsg)
        End If
	End Sub

    Public Function GetNextProduction(ByVal bUseCurrentProd As Boolean) As Boolean
        Dim X As Int32
        Dim yLowestNum As Byte = 255
        Dim lIndex As Int32 = -1

        Dim yMsg() As Byte
        Try
            'Now, determine next production
            If bUseCurrentProd = False OrElse CurrentProduction Is Nothing Then
                For X = 0 To ProductionUB
                    If ProductionIdx(X) <> -1 Then
                        If Production(X).OrderNumber < yLowestNum Then
                            lIndex = X
                            yLowestNum = Production(X).OrderNumber
                        End If
                    End If
                Next X
            End If

            If lIndex <> -1 OrElse (bUseCurrentProd = True AndAlso CurrentProduction Is Nothing = False) Then
                If ParentColony Is Nothing = False OrElse Me.Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID Then
                    'Ok, now, check if we can produce it or not.
                    If (bUseCurrentProd = False OrElse CurrentProduction Is Nothing = True) AndAlso Production Is Nothing = False AndAlso lIndex <= Production.GetUpperBound(0) AndAlso Production(lIndex) Is Nothing = False Then
                        CurrentProduction = Production(lIndex)
                        Production(lIndex) = Nothing
                        ProductionIdx(lIndex) = -1
                    End If

                    Dim lStartProdID As Int32 = -1
                    Dim iStartProdTypeID As Int16 = -1

                    If CurrentProduction Is Nothing = False Then
                        lStartProdID = CurrentProduction.ProductionID
                        iStartProdTypeID = CurrentProduction.ProductionTypeID
                        If CurrentProduction.ProdCost Is Nothing = False AndAlso CurrentProduction.ProdCost.ProductionCostType <> 1 AndAlso gb_Main_Loop_Running = True Then
                            CurrentProduction.bPaidFor = False
                            CurrentProduction.PointsProduced = 0
                        End If
                    End If

                    Dim yResResult As eResourcesResult = eResourcesResult.Insufficient_Clear
                    If CurrentProduction Is Nothing = False Then
                        If Me.Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID OrElse CurrentProduction.bPaidFor = True Then
                            yResResult = eResourcesResult.Sufficient
                        Else
                            yResResult = ParentColony.HasRequiredResources(CurrentProduction.ProdCost, Me, eyHasRequiredResourcesFlags.NoFlags)
                        End If
                    End If

                    If yResResult = eResourcesResult.Sufficient Then

                        If lStartProdID <> CurrentProduction.ProductionID OrElse iStartProdTypeID <> CurrentProduction.ProductionTypeID Then
                            Return True
                        End If

                        CurrentProduction.bPaidFor = True

                        Dim lSetFinishCycle As Int32 = 0

                        Select Case CurrentProduction.ProductionTypeID
                            'Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHangarTech, ObjectType.eHullTech, ObjectType.eMineralTech, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.ePrototype
                            Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.eMineralTech, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.ePrototype, ObjectType.eSpecialTech
                                'If Me.Owner.ObjectID = 1 Or Me.Owner.ObjectID = 6 Then CurrentProduction.PointsRequired = 1 Else 
                                CurrentProduction.PointsRequired = CurrentProduction.ProdCost.PointsRequired

                                'Ok, are we researching?
                                If Me.yProductionType = ProductionType.eResearch AndAlso CurrentProduction.ProductionTypeID <> ObjectType.eMineralTech Then
                                    'Ok, add us to the tech

                                    Dim oTech As Epica_Tech = Me.Owner.GetTech(CurrentProduction.ProductionID, CurrentProduction.ProductionTypeID)
                                    If oTech Is Nothing = False AndAlso oTech.ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eResearched AndAlso oTech.ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
                                        oTech.AssignResearcher(Me.ObjectID)
                                        If oTech.IsPrimaryResearcher(Me.ObjectID) = False Then
                                            lSetFinishCycle = Int32.MaxValue
                                        Else
                                            CurrentProduction.PointsProduced = 0
                                            CurrentProduction.SaveObject(True)
                                        End If
                                    Else : Return False
                                    End If
                                End If
                            Case ObjectType.eRepairItem
                                If CurrentProduction.ProdCost Is Nothing = True Then CurrentProduction.PointsRequired = 1 Else CurrentProduction.PointsRequired = CurrentProduction.ProdCost.PointsRequired
                                'TODO: Check if the repair target still requires this item to be repaired
                                'TODO: If the item being repaired is the Structure, need to set a flag that the unit will not launch

                            Case Else
                                CurrentProduction.PointsRequired = CurrentProduction.ProdCost.PointsRequired
                        End Select

                        If Owner.DeathBudgetEndTime > glCurrentCycle Then
                            'If CurrentProduction.ProductionTypeID <> ObjectType.eSpecialTech Then CurrentProduction.PointsRequired = 1
                            If Me.yProductionType <> ProductionType.eResearch Then
                                CurrentProduction.PointsRequired = 1
                            End If
                        End If

                        If Owner.yPlayerPhase = eyPlayerPhase.eInitialPhase AndAlso Owner.lTutorialStep = 205 Then CurrentProduction.PointsRequired = 1

                        If lIndex <> -1 Then ProductionIdx(lIndex) = -1
                        bProducing = True
                        CurrentProduction.lLastUpdateCycle = glCurrentCycle
                        CurrentProduction.lFinishCycle = lSetFinishCycle
                    ElseIf yResResult = eResourcesResult.Insufficient_Clear Then
                        'Notify the player that they don't have the resources (if that player is connected)
                        If Me.Owner.lConnectedPrimaryID > -1 OrElse Me.Owner.HasOnlineAliases(AliasingRights.eAddProduction Or AliasingRights.eViewTreasury) = True Then
                            ReDim yMsg(13)
                            System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProdFailed).CopyTo(yMsg, 0)
                            GetGUIDAsString.CopyTo(yMsg, 2)
                            System.BitConverter.GetBytes(CurrentProduction.ProductionID).CopyTo(yMsg, 8)
                            System.BitConverter.GetBytes(CurrentProduction.ProductionTypeID).CopyTo(yMsg, 12)
                            Me.Owner.SendPlayerMessage(yMsg, False, AliasingRights.eAddProduction Or AliasingRights.eViewTreasury)
                        End If

                        ProductionUB = -1
                        ReDim Production(ProductionUB)
                        ReDim ProductionIdx(ProductionUB)
                        bProducing = False
                        CurrentProduction = Nothing

                        Return False
                    End If
                End If
            Else
                ProductionUB = -1
                ReDim Production(ProductionUB)
                ReDim ProductionIdx(ProductionUB)
                bProducing = False
                CurrentProduction = Nothing

                'TODO: Send a message to the player that the facility is now dormant
            End If
        Catch ex As Exception
            'Just Clearing the queue here causes things to just stop building... and not clearing it causes things to build instantly
            LogEvent(LogEventType.CriticalError, "GetNextProduction: " & ex.Message)
            ProductionUB = -1
            ReDim Production(ProductionUB)
            ReDim ProductionIdx(ProductionUB)
            bProducing = False
            CurrentProduction = Nothing
        End Try
        Return True
    End Function

    'NOTE: This assumes the following statements have been executed already:
    '  DELETE FROM tblStructureProduction
    '  DELETE FROM tblAgentEffects WHERE EffectedItemTypeID = " & objecttype.efacility
    '  DELETE FROM tblMineBuyOrder
    '  DELETE FROM tblMineBuyOrderBid
    Public Function GetSaveObjectText() As String
        Dim sSQL As String

        If ObjectID < 1 Then
            SaveObject()
            Return ""
        End If

        Try
            Dim oSB As New System.Text.StringBuilder


            'UPDATE
            sSQL = "UPDATE tblStructure SET ParentID=" & CType(ParentObject, Epica_GUID).ObjectID & ", ParentTypeID=" & _
              CType(ParentObject, Epica_GUID).ObjTypeID & ", CurrentWorkers=" & Me.lPirate_For_PlayerID & ", OwnerID="
            If Owner Is Nothing = False Then
                sSQL = sSQL & Owner.ObjectID & ", "
            Else : sSQL = sSQL & "-1, "
            End If
            If ParentColony Is Nothing = False Then
                sSQL = sSQL & "ColonyID = " & ParentColony.ObjectID & ", "
            Else : sSQL = sSQL & "ColonyID = -1, "
            End If
            sSQL = sSQL & "FacilityName = '" & MakeDBStr(BytesToString(EntityName)) & "', Height = 0, Width = 0" & _
              ", LocX = " & LocX & ", LocY = " & LocZ & ", FacilityDefID = " & EntityDef.ObjectID & _
              ", LocAngle = " & LocAngle & ", Q1_HP = " & Q1_HP & ", Q2_HP = " & Q2_HP & ", Q3_HP = " & Q3_HP & _
              ", Q4_HP = " & Q4_HP & ", CurrentSpeed = " & CurrentSpeed & ", Structure_HP = " & Structure_HP & _
              ", Hangar_Cap = " & Hangar_Cap & ", Cargo_Cap = " & Cargo_Cap & _
              ", Fuel_Cap = " & Fuel_Cap & ", Shield_HP = " & Shield_HP & ", ExpLevel = " & ExpLevel & _
              ", CurrentStatus = " & CurrentStatus & ", ProductionTypeID = " & yProductionType & _
              ", CombatTactics = " & iCombatTactics & ", TargetingTactics = " & iTargetingTactics & _
              ", CurrentColonists = 0, CurrentEnlisted = 0, CurrentOfficers = 0, AutoLaunch = "
            If Me.AutoLaunch = True Then sSQL &= "1 " Else sSQL &= "0 "
            sSQL &= ", ReinforcingUnitGroupID = " & lReinforcingUnitGroupID & ", ExcludeFromColonyQueue = "
            If Me.bExcludeFromColonyQueue = True Then sSQL &= "1" Else sSQL &= "0"
            sSQL &= " WHERE StructureID = " & ObjectID

            oSB.AppendLine(sSQL)

            'sSQL = "DELETE FROM tblStructureProduction WHERE StructureID = " & ObjectID
            'oSB.AppendLine(sSQL)

            'Now, we need to save the Production items
            If CurrentProduction Is Nothing = False Then
                Dim bRecalc As Boolean = False
                If Me.yProductionType = ProductionType.eResearch Then
                    Select Case CType(CurrentProduction.ProductionTypeID, ObjectType)
                        Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eSpecialTech, ObjectType.eWeaponTech
                            Dim oTech As Epica_Tech = Me.Owner.GetTech(CurrentProduction.ProductionID, CurrentProduction.ProductionTypeID)
                            If oTech Is Nothing Then oTech = goInitialPlayer.GetTech(CurrentProduction.ProductionID, CurrentProduction.ProductionTypeID)
                            If oTech Is Nothing = False Then
                                If oTech.IsPrimaryResearcher(Me.ObjectID) = True Then
                                    bRecalc = True
                                Else
                                    oTech.RecalcPrimarysProdFactor()
                                    CurrentProduction.PointsProduced = oTech.GetPrimarysPointsProduced()
                                End If
                            End If
                        Case Else
                            bRecalc = True
                    End Select
                Else
                    If Me.bProducing = False Then
                        CurrentProduction = Nothing
                    Else
                        bRecalc = True
                    End If
                End If

                If bRecalc = True Then RecalcProduction()
                If CurrentProduction Is Nothing = False Then oSB.AppendLine(CurrentProduction.GetSaveObjectText())

            End If
            For X As Int32 = 0 To Math.Min(25, ProductionUB)
                If ProductionIdx(X) <> -1 Then
                    oSB.AppendLine(Production(X).GetSaveObjectText())
                End If
            Next X

            oSB.AppendLine(MyBase.GetSaveAgentEffectsText())

            If oMiningBid Is Nothing = False Then oSB.AppendLine(oMiningBid.GetSaveObjectText()) ' oMiningBid.SaveObject()

            Return oSB.ToString()

        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)

        End Try
        Return ""
    End Function

	Public Function SaveObject() As Boolean
		Dim bResult As Boolean = False
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand
		'Dim oOrder As ObjectOrder

		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

		Try
			If ObjectID = -1 Then
				'INSERT
                sSQL = "INSERT INTO tblStructure (ParentID, ParentTypeID, CurrentWorkers, OwnerID, ColonyID, " & _
                  "FacilityName, Height, Width, LocX, LocY, FacilityDefID, LocAngle, Q1_HP, " & _
                  "Q2_HP, Q3_HP, Q4_HP, CurrentSpeed, Structure_HP, Hangar_Cap, Cargo_Cap, Fuel_Cap, " & _
                  "Shield_HP, ExpLevel, CurrentStatus, ProductionTypeID, CombatTactics, TargetingTactics, CurrentColonists, " & _
                  "CurrentEnlisted, CurrentOfficers, AutoLaunch, ReinforcingUnitGroupID, ExcludeFromColonyQueue, RallyPointX, RallyPointZ, RallyPointA, RallyPointEnvirID, RallyPointEnvirTypeID) VALUES (" & _
                  CType(ParentObject, Epica_GUID).ObjectID & ", " & CType(ParentObject, Epica_GUID).ObjTypeID & ", " & Me.lPirate_For_PlayerID & ", "
				If Owner Is Nothing = False Then
					sSQL = sSQL & Owner.ObjectID & ", "
				Else : sSQL = sSQL & "-1, "
				End If
				If ParentColony Is Nothing = False Then
					sSQL = sSQL & ParentColony.ObjectID & ", "
				Else : sSQL = sSQL & "-1, "
				End If
				sSQL = sSQL & "'" & MakeDBStr(BytesToString(EntityName)) & "', 0, 0, " & LocX & ", " & LocZ & ", " & _
				  EntityDef.ObjectID & ", " & LocAngle & ", " & Q1_HP & ", " & Q2_HP & ", " & Q3_HP & ", " & _
				  Q4_HP & "," & CurrentSpeed & ", " & Structure_HP & ", " & Hangar_Cap & _
				  ", " & Cargo_Cap & ", " & Fuel_Cap & ", " & Shield_HP & ", " & ExpLevel & ", " & CurrentStatus & _
				  ", " & yProductionType & ", " & iCombatTactics & ", " & iTargetingTactics & ", 0, 0, 0, "
				If Me.AutoLaunch = True Then sSQL &= "1" Else sSQL &= "0"
                sSQL &= ", " & lReinforcingUnitGroupID & ", "
                If Me.bExcludeFromColonyQueue = True Then sSQL &= "1" Else sSQL &= "0"
                sSQL &= ", " & RallyPointX.ToString & ", " & RallyPointZ.ToString & ", " & RallyPointA.ToString & ", " & RallyPointEnvirID.ToString & ", " & RallyPointEnvirTypeID.ToString
                sSQL &= ")"
			Else
				'UPDATE
				sSQL = "UPDATE tblStructure SET ParentID=" & CType(ParentObject, Epica_GUID).ObjectID & ", ParentTypeID=" & _
				  CType(ParentObject, Epica_GUID).ObjTypeID & ", CurrentWorkers=" & Me.lPirate_For_PlayerID & ", OwnerID="
				If Owner Is Nothing = False Then
					sSQL = sSQL & Owner.ObjectID & ", "
				Else : sSQL = sSQL & "-1, "
				End If
				If ParentColony Is Nothing = False Then
					sSQL = sSQL & "ColonyID = " & ParentColony.ObjectID & ", "
				Else : sSQL = sSQL & "ColonyID = -1, "
				End If
				sSQL = sSQL & "FacilityName = '" & MakeDBStr(BytesToString(EntityName)) & "', Height = 0, Width = 0" & _
				  ", LocX = " & LocX & ", LocY = " & LocZ & ", FacilityDefID = " & EntityDef.ObjectID & _
				  ", LocAngle = " & LocAngle & ", Q1_HP = " & Q1_HP & ", Q2_HP = " & Q2_HP & ", Q3_HP = " & Q3_HP & _
				  ", Q4_HP = " & Q4_HP & ", CurrentSpeed = " & CurrentSpeed & ", Structure_HP = " & Structure_HP & _
				  ", Hangar_Cap = " & Hangar_Cap & ", Cargo_Cap = " & Cargo_Cap & _
				  ", Fuel_Cap = " & Fuel_Cap & ", Shield_HP = " & Shield_HP & ", ExpLevel = " & ExpLevel & _
				  ", CurrentStatus = " & CurrentStatus & ", ProductionTypeID = " & yProductionType & _
				  ", CombatTactics = " & iCombatTactics & ", TargetingTactics = " & iTargetingTactics & _
				  ", CurrentColonists = 0, CurrentEnlisted = 0, CurrentOfficers = 0, AutoLaunch = "
				If Me.AutoLaunch = True Then sSQL &= "1 " Else sSQL &= "0 "
                sSQL &= ", ReinforcingUnitGroupID = " & lReinforcingUnitGroupID & ", ExcludeFromColonyQueue = "
                If Me.bExcludeFromColonyQueue = True Then sSQL &= "1" Else sSQL &= "0"
                sSQL &= ", RallyPointX=" & RallyPointX.ToString & ", RallyPointZ=" & RallyPointZ.ToString & ", RallyPointA=" & RallyPointA.ToString & ", RallyPointEnvirID=" & RallyPointEnvirID.ToString & ", RallyPointEnvirTypeID=" & RallyPointEnvirTypeID.ToString
                sSQL &= " WHERE StructureID = " & ObjectID
			End If
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			If oComm.ExecuteNonQuery() = 0 Then
				Err.Raise(-1, "SaveObject", "No records affected when saving object!")
			End If
			If ObjectID = -1 Then
				Dim oData As OleDb.OleDbDataReader
				oComm = Nothing
                sSQL = "SELECT MAX(StructureID) FROM tblStructure WHERE FacilityName = '" & MakeDBStr(BytesToString(EntityName)) & "'"
				oComm = New OleDb.OleDbCommand(sSQL, goCN)
				oData = oComm.ExecuteReader(CommandBehavior.Default)
				If oData.Read Then
					ObjectID = CInt(oData(0))
				End If
				oData.Close()
				oData = Nothing
            End If

            Try
                sSQL = "DELETE FROM tblStructureProduction WHERE StructureID = " & ObjectID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oComm.ExecuteNonQuery()
                oComm = Nothing
            Catch
            End Try

            'Now, we need to save the Production items
            If CurrentProduction Is Nothing = False Then
                Dim bRecalc As Boolean = False
                If Me.yProductionType = ProductionType.eResearch Then
                    Select Case CType(CurrentProduction.ProductionTypeID, ObjectType)
                        Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eSpecialTech, ObjectType.eWeaponTech
                            Dim oTech As Epica_Tech = Me.Owner.GetTech(CurrentProduction.ProductionID, CurrentProduction.ProductionTypeID)
                            If oTech Is Nothing Then oTech = goInitialPlayer.GetTech(CurrentProduction.ProductionID, CurrentProduction.ProductionTypeID)
                            If oTech Is Nothing = False Then
                                If oTech.IsPrimaryResearcher(Me.ObjectID) = True Then
                                    bRecalc = True
                                Else
                                    oTech.RecalcPrimarysProdFactor()
                                    CurrentProduction.PointsProduced = oTech.GetPrimarysPointsProduced()
                                End If
                            End If
                        Case Else
                            bRecalc = True
                    End Select
                Else
                    If Me.bProducing = False Then
                        CurrentProduction = Nothing
                    Else
                        bRecalc = True
                    End If
                End If

                If bRecalc = True Then RecalcProduction()
                If CurrentProduction Is Nothing = False AndAlso CurrentProduction.SaveObject(False) = False Then
                    'Err.Raise(-1, "Facility.Production.SaveObject()", "Unable to save production item (current).")
                End If
            End If
            For X As Int32 = 0 To Math.Min(25, ProductionUB)
                If ProductionIdx(X) <> -1 Then
                    If Production(X).SaveObject(False) = False Then
                        'Err.Raise(-1, "Facility.Production.SaveObject()", "Unable to save production item.")
                    End If
                End If
            Next X

            If oMiningBid Is Nothing = False Then oMiningBid.SaveObject()

            MyBase.SaveAgentEffects()

            bNeedsSaved = False
            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
            Return bResult
    End Function

	Public Function VerifyCompletion() As Boolean
		Dim blProduced As Int64

		If bProducing = True Then
			Dim blRequired As Int64 = CurrentProduction.PointsRequired
			If yProductionType = ProductionType.eResearch Then
				blRequired = CLng(CurrentProduction.PointsRequired * Me.Owner.fFactionResearchTimeMultiplier)
			End If

			blProduced = CurrentProduction.PointsProduced
			Dim blCyclesPassed As Int64 = (glCurrentCycle - CurrentProduction.lLastUpdateCycle)
			blProduced += (blCyclesPassed * CLng(mlProdPoints))
			Return blProduced >= blRequired
		Else : Return False
		End If
	End Function

	Public Sub RecalcProduction()
		If ParentColony Is Nothing Then
			mlProdPoints = 100
			Return
		End If

		If bProducing = True Then
			If CurrentProduction Is Nothing = False Then
				'Ok, first, store off are points accumulated thus far
				CurrentProduction.PointsProduced += (CLng(glCurrentCycle - CurrentProduction.lLastUpdateCycle) * CLng(mlProdPoints))
			Else
				bProducing = False
				Return
			End If
		End If
		'mlProdPoints = CInt(Math.Ceiling(EntityDef.ProdFactor * (mlCurrentWorkers / MaxWorkers)))
        If Active = True OrElse Me.yProductionType = ProductionType.eMining Then
            If ParentColony.MoraleMultiplier < 0.0F Then
                mlProdPoints = CInt(Math.Ceiling(EntityDef.ProdFactor * 0.05))
            Else : mlProdPoints = CInt(Math.Ceiling(EntityDef.ProdFactor * ParentColony.WorkforceEfficiency * ParentColony.MoraleMultiplier))
            End If
        Else : mlProdPoints = 0
        End If
        If mlProdPoints < CInt(Math.Ceiling(EntityDef.ProdFactor * 0.05F)) Then mlProdPoints = CInt(Math.Ceiling(EntityDef.ProdFactor * 0.05F))

        If CurrentProduction Is Nothing = False Then
            If Me.yProductionType <> ProductionType.eRefining Then
                If CurrentProduction.ProductionTypeID = ObjectType.eMineral Then
                    mlProdPoints \= 2
                    If mlProdPoints < 1 Then mlProdPoints = 1
                End If
            Else
                If CurrentProduction.ProductionTypeID = ObjectType.eArmorTech Then
                    mlProdPoints *= 10
                End If
            End If
        End If

        If Me.Owner Is Nothing = False AndAlso Me.Owner.bInFullLockDown = True Then
            If CurrentProduction Is Nothing = True OrElse CurrentProduction.ProductionTypeID <> ObjectType.eSpecialTech Then mlProdPoints = 0
        End If

        If yProductionType = ProductionType.eResearch AndAlso bProducing = True AndAlso CurrentProduction Is Nothing = False AndAlso CurrentProduction.ProductionTypeID <> ObjectType.eMineralTech Then
            Dim oTech As Epica_Tech = Owner.GetTech(CurrentProduction.ProductionID, CurrentProduction.ProductionTypeID)
            If oTech Is Nothing = False Then
                If oTech.IsPrimaryResearcher(Me.ObjectID) = True Then
                    mlProdPoints = oTech.GetPrimaryResearcherProdFactor()
                Else
                    'Ok, I'm in recalc production and i am not the primary researcher...
                    CurrentProduction.PointsProduced = 0
                    oTech.RecalcPrimarysProdFactor()
                    Return
                End If
            Else : Return
            End If
        End If

        'Ok, determine our prod factor for agent effects
        Dim fProdFactorMult As Single = 1.0F
        For X As Int32 = 0 To mlAgentEffectUB
            If EffectValid(X) = True Then
                If moAgentEffects(X).yType = AgentEffectType.eProduction Then
                    If moAgentEffects(X).bAmountAsPerc = True Then
                        fProdFactorMult *= (moAgentEffects(X).lAmount / 100.0F)
                    Else : mlProdPoints -= moAgentEffects(X).lAmount
                    End If
                End If
            End If
        Next X
        mlProdPoints = CInt(mlProdPoints * fProdFactorMult)
        If mlProdPoints < 1 Then mlProdPoints = 1

        'Corruption - if planet is corrupted, no production
        If Me.ParentObject Is Nothing = False AndAlso CType(ParentObject, Epica_GUID).ObjTypeID = ObjectType.ePlanet Then
            If CType(Me.ParentObject, Planet).PlanetInCorruption(Me.Owner.lStartedEnvirID, Me.Owner.iStartedEnvirTypeID) = True Then
                mlProdPoints = 0
                If Me.yProductionType = ProductionType.eResearch AndAlso CurrentProduction Is Nothing = False Then
                    CurrentProduction.lFinishCycle = Int32.MaxValue
                    CurrentProduction.lLastUpdateCycle = glCurrentCycle
                    Return
                End If
            End If
        End If

        'Now, how much production time are we going to need?
        If Active AndAlso bProducing = True Then
            If mlProdPoints <> 0 Then
                Dim blRequired As Int64 = CurrentProduction.PointsRequired
                If yProductionType = ProductionType.eResearch Then
                    blRequired = CLng(CurrentProduction.PointsRequired * Me.Owner.fFactionResearchTimeMultiplier)
                End If

                Dim blTemp As Int64 = CLng(blRequired - CurrentProduction.PointsProduced) \ mlProdPoints
                If blTemp + glCurrentCycle > Int32.MaxValue Then
                    CurrentProduction.lFinishCycle = Int32.MaxValue
                Else
                    CurrentProduction.lFinishCycle = CInt(blTemp + glCurrentCycle)
                End If
                CurrentProduction.lLastUpdateCycle = glCurrentCycle
            Else
                bProducing = False
            End If
            'ElseIf (EntityDef.ProductionTypeID = ProductionType.eColonists) OrElse (EntityDef.ProductionTypeID = ProductionType.eCommandCenterSpecial) Then
            '    OptColonists = mlProdPoints
            '    MaxColonists = CInt(Math.Ceiling(mlProdPoints * 1.1F))

            '    'TODO: If the facility is ever lost or the mlProdPoints drops below the current colonists, then what happens?
        Else
            bProducing = False
        End If
	End Sub

	Public ReadOnly Property PowerConsumption() As Int32
		Get
			If (CurrentStatus And elUnitStatus.eFacilityPowered) <> 0 Then
				Dim lResult As Int32 = EntityDef.PowerFactor
				If lResult < 0 Then
					'power generator... test for agents
					Dim fPowerMult As Single = 1.0F
					For X As Int32 = 0 To mlAgentEffectUB
						If EffectValid(X) = True Then
							If moAgentEffects(X).yType = AgentEffectType.eProduction Then
								If moAgentEffects(X).bAmountAsPerc = True Then
									fPowerMult *= (moAgentEffects(X).lAmount / 100.0F)
								Else : lResult += Math.Abs(moAgentEffects(X).lAmount)		'we absolute because negative values would increase production
								End If
							End If
						End If
					Next X
					lResult = CInt(lResult * fPowerMult)
				End If
				Return lResult
			Else
				Return 0I
			End If
		End Get
	End Property

	Public Sub RemoveMe()
		'ParentColony.AddNonWorkers(Me.CurrentWorkers)
		CurrentStatus = CurrentStatus Xor elUnitStatus.eFacilityPowered 'CurrentStatus -= elUnitStatus.eFacilityPowered
		'Me.CurrentWorkers = 0
		'Me.PowerConsumption = 0
        If ParentColony Is Nothing = False Then ParentColony.UpdateAllValues(-1)

        If Me.EntityDef Is Nothing = False AndAlso (Me.EntityDef.ModelID And 255) = 148 Then
            For X As Int32 = 0 To glPlanetUB
                If glPlanetIdx(X) > -1 Then
                    Dim oPlanet As Planet = goPlanet(X)
                    If oPlanet Is Nothing = False AndAlso oPlanet.RingMinerID = Me.ObjectID Then oPlanet.RingMinerID = -1
                End If
            Next X
        End If
	End Sub

	Private Function FillProductionItemValues(ByRef oProd As EntityProduction, ByVal lProdID As Int32, ByVal iProdTypeID As Int16, ByVal yOrderNumber As Byte, ByVal lQuantity As Int32) As Boolean
		With oProd
			.oParent = Me
			.ProductionID = lProdID
			.ProductionTypeID = iProdTypeID
			.OrderNumber = yOrderNumber
			.PointsProduced = 0
			.lProdCount = lQuantity

			'Now, get the Production Cost associated to the item...
			Select Case iProdTypeID
				Case ObjectType.eUnitDef
					'creating a new unit
					Dim oTmpdef As Epica_Entity_Def = GetEpicaUnitDef(lProdID)
					.ProdCost = oTmpdef.ProductionRequirements
					oTmpdef = Nothing
				Case ObjectType.eFacilityDef
					'creating a new facility
					Dim oTmpDef As FacilityDef = GetEpicaFacilityDef(lProdID)
					.ProdCost = oTmpDef.ProductionRequirements
					oTmpDef = Nothing
				Case ObjectType.eCredits
					'creating credits
				Case ObjectType.eMorale
					'creating morale
					'Case ObjectType.ePower
					'    'shouldn't be here...
					'Case ObjectType.eColonists
					'    'shouldn't be here...
					'Case ObjectType.eFood
					'    'shouldn't be here...
					'Case ObjectType.eMiningPower
					'    'shouldn't be here...
				Case ObjectType.eEnlisted
					'TODO: Un-hardcode this
					.ProdCost = New ProductionCost()
					.ProdCost.ColonistCost = Owner.oSpecials.yEnlistedColonistCost
					.ProdCost.CreditCost = 200
					.ProdCost.EnlistedCost = 0
					.ProdCost.ItemCostUB = -1
					ReDim .ProdCost.ItemCosts(-1)
					.ProdCost.PointsRequired = 150000
				Case ObjectType.eOfficers
					'TODO: Un-hardcode this
					.ProdCost = New ProductionCost
					.ProdCost.ColonistCost = 0
					.ProdCost.CreditCost = 700
					.ProdCost.EnlistedCost = Owner.oSpecials.yOfficersEnlistedCost
					.ProdCost.ItemCostUB = -1
					ReDim .ProdCost.ItemCosts(-1)
					.ProdCost.PointsRequired = 300000
				Case ObjectType.eMineralTech
					'researching a mineral
					For X As Int32 = 0 To Me.Owner.lPlayerMineralUB
						If Me.Owner.oPlayerMinerals(X).lMineralID = lProdID Then
							.ProdCost = Me.Owner.oPlayerMinerals(X).GetProductionCost()
							Exit For
						End If
					Next X
				Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, _
				  ObjectType.eHullTech, ObjectType.eRadarTech, _
				  ObjectType.ePrototype, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.eSpecialTech
					Dim oTmpTech As Epica_Tech = Me.Owner.GetTech(lProdID, iProdTypeID)

                    Dim bResult As Boolean = False
                    If oTmpTech Is Nothing = False AndAlso oTmpTech.ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
                        If (Me.yProductionType And ProductionType.eProduction) <> 0 Then
                            'Ok a producer trying to produce this requires that the tech be researched
                            If oTmpTech.ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eResearched OrElse oTmpTech.ObjTypeID = ObjectType.eSpecialTech Then Return False
                            .ProdCost = oTmpTech.GetCurrentProductionCost
                            bResult = True
                        ElseIf Me.yProductionType = ProductionType.eRefining Then
                            If oTmpTech.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                                If iProdTypeID = ObjectType.eArmorTech OrElse iProdTypeID = ObjectType.eAlloyTech Then
                                    bResult = True
                                    .ProdCost = oTmpTech.GetCurrentProductionCost
                                End If
                            End If
                        ElseIf (Me.yProductionType And ProductionType.eResearch) <> 0 Then
                            If oTmpTech.ComponentDevelopmentPhase <> Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                                .ProdCost = oTmpTech.GetCurrentProductionCost
                                bResult = True
                            End If
                        End If
                    End If
                    If bResult = False Then Return False
				Case ObjectType.eMineral
					'Refinery is building a mineral
					Dim oTmpMin As Mineral = GetEpicaMineral(lProdID)
					If oTmpMin Is Nothing Then Return False
					If oTmpMin.lAlloyTechID < 1 Then Return False
					Dim oTmpTech As Epica_Tech = Me.Owner.GetTech(oTmpMin.lAlloyTechID, ObjectType.eAlloyTech)
					If oTmpTech Is Nothing Then Return False
					.ProdCost = oTmpTech.GetProductionCost
			End Select
		End With

		Return True
	End Function

	''' <summary>
	''' This routine pushes the current production list back 2, puts the current production in position 1
	'''   and then puts the production passed in in position 0
	''' </summary>
	''' <remarks></remarks>
	Public Function InsertProduction(ByVal lProdID As Int32, ByVal iProdTypeID As Int16, ByVal lQuantity As Int32) As Boolean

		Dim lPushBackCnt As Int32 = 1

		'Ok, if the current production is not nothing...
		If CurrentProduction Is Nothing = False Then
			'Ok, need to recalc production
			Me.RecalcProduction()
			lPushBackCnt += 1
		End If

		'First, let's compact our list if it isn't already
		Dim lProdCnt As Int32 = 0
		For X As Int32 = 0 To Me.ProductionUB
			If Me.ProductionIdx(X) <> -1 Then lProdCnt += 1
		Next X

		'Now, go back through and do our deal
		Dim lCurrIdx As Int32 = 0
		For X As Int32 = 0 To Me.ProductionUB
			If Me.ProductionIdx(X) <> -1 Then
				Me.ProductionIdx(lCurrIdx) = Me.ProductionIdx(X)
				Me.Production(lCurrIdx) = Me.Production(X)
				lCurrIdx += 1
			End If
		Next X

		'Now... our new count
		lProdCnt += lPushBackCnt
		Me.ProductionUB = lProdCnt - 1
		ReDim Preserve Me.ProductionIdx(Me.ProductionUB)
		ReDim Preserve Me.Production(Me.ProductionUB)

		'Now, go through, shift our current contents up two
		For X As Int32 = Me.ProductionUB - lPushBackCnt To 0 Step -1
			Me.ProductionIdx(X + lPushBackCnt) = Me.ProductionIdx(X)
			Me.Production(X + lPushBackCnt) = Me.Production(X)
		Next X

		'Now, put our current production item in....
		If lPushBackCnt = 2 Then
			Me.ProductionIdx(1) = CurrentProduction.ProductionID
			Me.Production(1) = CurrentProduction
		End If

		'Now, put our new item in
		Me.Production(0) = New EntityProduction()

		CurrentProduction = Nothing
		If Me.FillProductionItemValues(Me.Production(0), lProdID, iProdTypeID, 0, lQuantity) = True Then
			Me.ProductionIdx(0) = lProdID
			Return True
		Else : Return False
		End If
	End Function

	Public Sub ClearQueue()
		'Ok, for now, clearing the queue will remove all Production objects in the Production array AND set the CurrentProduction's ProdCnt = 1
		ProductionUB = -1
		ReDim Production(ProductionUB)
		ReDim ProductionIdx(ProductionUB)

		If CurrentProduction Is Nothing = False AndAlso CurrentProduction.lProdCount > 1 Then
			CurrentProduction.lProdCount = 1
		End If
	End Sub

    'Public Sub HandleRepairItem(ByRef oEntity As Epica_Entity, ByRef oDef As Epica_Entity_Def, ByVal lRepairItem As Int32)
    '	'Now, what are we repairing?
    '	Select Case lRepairItem
    '		Case -2, -3, -4, -5
    '			'Armor side
    '			Dim lCurrent As Int32
    '			Dim lMax As Int32

    '			If lRepairItem = -2 Then
    '				lCurrent = oEntity.Q1_HP : lMax = oDef.Q1_MaxHP
    '			ElseIf lRepairItem = -3 Then
    '				lCurrent = oEntity.Q2_HP : lMax = oDef.Q2_MaxHP
    '			ElseIf lRepairItem = -4 Then
    '				lCurrent = oEntity.Q3_HP : lMax = oDef.Q3_MaxHP
    '			Else : lCurrent = oEntity.Q4_HP : lMax = oDef.Q4_MaxHP
    '			End If

    '			If lCurrent <> lMax AndAlso lMax <> 0 Then
    '				Dim lRepNeeded As Int32 = 0
    '				If oDef.oPrototype.oArmorTech Is Nothing Then lRepNeeded = lMax - lCurrent Else lRepNeeded = CInt(Math.Ceiling((lMax - lCurrent) / oDef.oPrototype.oArmorTech.lHPPerPlate))
    '				AddRepairItem(lRepairItem, lRepNeeded, oDef.oPrototype.oArmorTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '			End If

    '		Case -6
    '			'structure
    '			Dim lCurrent As Int32 = oEntity.Structure_HP
    '			Dim lMax As Int32 = oDef.Structure_MaxHP

    '			If lCurrent <> lMax AndAlso lMax <> 0 Then
    '				Dim lRepNeeded As Int32 = lMax - lCurrent
    '				AddRepairItem(lRepairItem, lRepNeeded, oDef.oPrototype.oHullTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '			End If
    '		Case -7
    '			'all defensive values
    '			Dim lCurrent As Int32 = oEntity.Structure_HP
    '			Dim lMax As Int32 = oDef.Structure_MaxHP
    '			If lCurrent <> lMax AndAlso lMax <> 0 Then
    '				Dim lRepNeeded As Int32 = lMax - lCurrent
    '				AddRepairItem(-6, lRepNeeded, oDef.oPrototype.oHullTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '			End If
    '			lCurrent = oEntity.Q1_HP : lMax = oDef.Q1_MaxHP
    '			If lCurrent <> lMax AndAlso lMax <> 0 Then
    '				Dim lRepNeeded As Int32 = CInt(Math.Ceiling((lMax - lCurrent) / oDef.oPrototype.oArmorTech.lHPPerPlate))
    '				AddRepairItem(-2, lRepNeeded, oDef.oPrototype.oArmorTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '			End If
    '			lCurrent = oEntity.Q2_HP : lMax = oDef.Q2_MaxHP
    '			If lCurrent <> lMax AndAlso lMax <> 0 Then
    '				Dim lRepNeeded As Int32 = CInt(Math.Ceiling((lMax - lCurrent) / oDef.oPrototype.oArmorTech.lHPPerPlate))
    '				AddRepairItem(-3, lRepNeeded, oDef.oPrototype.oArmorTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '			End If
    '			lCurrent = oEntity.Q3_HP : lMax = oDef.Q3_MaxHP
    '			If lCurrent <> lMax AndAlso lMax <> 0 Then
    '				Dim lRepNeeded As Int32 = CInt(Math.Ceiling((lMax - lCurrent) / oDef.oPrototype.oArmorTech.lHPPerPlate))
    '				AddRepairItem(-4, lRepNeeded, oDef.oPrototype.oArmorTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '			End If
    '			lCurrent = oEntity.Q4_HP : lMax = oDef.Q4_MaxHP
    '			If lCurrent <> lMax AndAlso lMax <> 0 Then
    '				Dim lRepNeeded As Int32 = CInt(Math.Ceiling((lMax - lCurrent) / oDef.oPrototype.oArmorTech.lHPPerPlate))
    '				AddRepairItem(-5, lRepNeeded, oDef.oPrototype.oArmorTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '			End If
    '		Case Int32.MaxValue
    '			'repair all components
    '			Dim lStatus As Int32 = oEntity.CurrentStatus
    '			Dim lDefStatus As Int32 = 0
    '			For X As Int32 = 0 To oDef.lSideCrits.GetUpperBound(0)
    '				lDefStatus = lDefStatus Or oDef.lSideCrits(X)
    '			Next X

    '			If (lDefStatus And elUnitStatus.eEngineOperational) <> 0 Then
    '				If (lStatus And elUnitStatus.eEngineOperational) = 0 Then
    '					AddRepairItem(elUnitStatus.eEngineOperational, 1, oDef.oPrototype.oEngineTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '				End If
    '			End If
    '			If (lDefStatus And elUnitStatus.eFuelBayOperational) <> 0 Then
    '				If (lStatus And elUnitStatus.eFuelBayOperational) = 0 Then
    '					AddRepairItem(elUnitStatus.eFuelBayOperational, 1, oDef.oPrototype.oHullTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '				End If
    '			End If
    '			If (lDefStatus And elUnitStatus.eRadarOperational) <> 0 Then
    '				If (lStatus And elUnitStatus.eRadarOperational) = 0 Then
    '					AddRepairItem(elUnitStatus.eRadarOperational, 1, oDef.oPrototype.oRadarTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '				End If
    '			End If
    '			If (lDefStatus And elUnitStatus.eShieldOperational) <> 0 Then
    '				If (lStatus And elUnitStatus.eShieldOperational) = 0 Then
    '					AddRepairItem(elUnitStatus.eShieldOperational, 1, oDef.oPrototype.oShieldTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '				End If
    '			End If
    '			If (lDefStatus And elUnitStatus.eCargoBayOperational) <> 0 Then
    '				If (lStatus And elUnitStatus.eCargoBayOperational) = 0 Then
    '					AddRepairItem(elUnitStatus.eCargoBayOperational, 1, oDef.oPrototype.oHullTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '				End If
    '			End If
    '			If (lDefStatus And elUnitStatus.eHangarOperational) <> 0 Then
    '				If (lStatus And elUnitStatus.eHangarOperational) = 0 Then
    '					AddRepairItem(elUnitStatus.eHangarOperational, 1, oDef.oPrototype.oHullTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '				End If
    '			End If

    '			'Now for weapons
    '			If (lDefStatus And elUnitStatus.eAftWeapon1) <> 0 Then
    '				If (lStatus And elUnitStatus.eAftWeapon1) = 0 Then
    '					For X As Int32 = 0 To oDef.WeaponDefUB
    '						If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eAftWeapon1 Then
    '							AddRepairItem(elUnitStatus.eAftWeapon1, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '						End If
    '					Next X
    '				End If
    '			End If
    '			If (lDefStatus And elUnitStatus.eAftWeapon2) <> 0 Then
    '				If (lStatus And elUnitStatus.eAftWeapon2) = 0 Then
    '					For X As Int32 = 0 To oDef.WeaponDefUB
    '						If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eAftWeapon2 Then
    '							AddRepairItem(elUnitStatus.eAftWeapon2, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '						End If
    '					Next X
    '				End If
    '			End If
    '			If (lDefStatus And elUnitStatus.eLeftWeapon1) <> 0 Then
    '				If (lStatus And elUnitStatus.eLeftWeapon1) = 0 Then
    '					For X As Int32 = 0 To oDef.WeaponDefUB
    '						If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eLeftWeapon1 Then
    '							AddRepairItem(elUnitStatus.eLeftWeapon1, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '						End If
    '					Next X
    '				End If
    '			End If
    '			If (lDefStatus And elUnitStatus.eLeftWeapon2) <> 0 Then
    '				If (lStatus And elUnitStatus.eLeftWeapon2) = 0 Then
    '					For X As Int32 = 0 To oDef.WeaponDefUB
    '						If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eLeftWeapon2 Then
    '							AddRepairItem(elUnitStatus.eLeftWeapon2, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '						End If
    '					Next X
    '				End If
    '			End If
    '			If (lDefStatus And elUnitStatus.eForwardWeapon1) <> 0 Then
    '				If (lStatus And elUnitStatus.eForwardWeapon1) = 0 Then
    '					For X As Int32 = 0 To oDef.WeaponDefUB
    '						If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eForwardWeapon1 Then
    '							AddRepairItem(elUnitStatus.eForwardWeapon1, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '						End If
    '					Next X
    '				End If
    '			End If
    '			If (lDefStatus And elUnitStatus.eForwardWeapon2) <> 0 Then
    '				If (lStatus And elUnitStatus.eForwardWeapon2) = 0 Then
    '					For X As Int32 = 0 To oDef.WeaponDefUB
    '						If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eForwardWeapon2 Then
    '							AddRepairItem(elUnitStatus.eForwardWeapon2, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '						End If
    '					Next X
    '				End If
    '			End If
    '			If (lDefStatus And elUnitStatus.eRightWeapon1) <> 0 Then
    '				If (lStatus And elUnitStatus.eRightWeapon1) = 0 Then
    '					For X As Int32 = 0 To oDef.WeaponDefUB
    '						If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eRightWeapon1 Then
    '							AddRepairItem(elUnitStatus.eRightWeapon1, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '						End If
    '					Next X
    '				End If
    '			End If
    '			If (lDefStatus And elUnitStatus.eRightWeapon2) <> 0 Then
    '				If (lStatus And elUnitStatus.eRightWeapon2) = 0 Then
    '					For X As Int32 = 0 To oDef.WeaponDefUB
    '						If oDef.WeaponDefs(X).lEntityStatusGroup = elUnitStatus.eRightWeapon2 Then
    '							AddRepairItem(elUnitStatus.eRightWeapon2, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '						End If
    '					Next X
    '				End If
    '			End If

    '		Case Is > 0
    '			'repair a component
    '			Dim lDefStatus As Int32 = 0
    '			For X As Int32 = 0 To oDef.lSideCrits.GetUpperBound(0)
    '				lDefStatus = lDefStatus Or oDef.lSideCrits(X)
    '			Next X
    '			If (lDefStatus And lRepairItem) <> 0 Then
    '				If (oEntity.CurrentStatus And lRepairItem) = 0 Then
    '					Select Case lRepairItem
    '						Case elUnitStatus.eEngineOperational
    '							AddRepairItem(lRepairItem, 1, oDef.oPrototype.oEngineTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '						Case elUnitStatus.eRadarOperational
    '							AddRepairItem(lRepairItem, 1, oDef.oPrototype.oRadarTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '						Case elUnitStatus.eShieldOperational
    '							AddRepairItem(lRepairItem, 1, oDef.oPrototype.oShieldTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '						Case elUnitStatus.eCargoBayOperational, elUnitStatus.eHangarOperational, elUnitStatus.eFuelBayOperational
    '							AddRepairItem(lRepairItem, 1, oDef.oPrototype.oHullTech, oDef.ProductionRequirements.PointsRequired, oEntity)
    '						Case Else
    '							Dim lTemp As Int32

    '							lTemp = elUnitStatus.eAftWeapon1
    '							For X As Int32 = 0 To oDef.WeaponDefUB
    '								If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
    '									AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '								End If
    '							Next X
    '							lTemp = elUnitStatus.eAftWeapon2
    '							For X As Int32 = 0 To oDef.WeaponDefUB
    '								If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
    '									AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '								End If
    '							Next X
    '							lTemp = elUnitStatus.eForwardWeapon1
    '							For X As Int32 = 0 To oDef.WeaponDefUB
    '								If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
    '									AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '								End If
    '							Next X
    '							lTemp = elUnitStatus.eForwardWeapon2
    '							For X As Int32 = 0 To oDef.WeaponDefUB
    '								If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
    '									AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '								End If
    '							Next X
    '							lTemp = elUnitStatus.eRightWeapon1
    '							For X As Int32 = 0 To oDef.WeaponDefUB
    '								If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
    '									AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '								End If
    '							Next X
    '							lTemp = elUnitStatus.eRightWeapon2
    '							For X As Int32 = 0 To oDef.WeaponDefUB
    '								If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
    '									AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '								End If
    '							Next X
    '							lTemp = elUnitStatus.eLeftWeapon1
    '							For X As Int32 = 0 To oDef.WeaponDefUB
    '								If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
    '									AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '								End If
    '							Next X
    '							lTemp = elUnitStatus.eLeftWeapon2
    '							For X As Int32 = 0 To oDef.WeaponDefUB
    '								If oDef.WeaponDefs(X).lEntityStatusGroup = lTemp Then
    '									AddRepairItem(lTemp, 1, CType(oDef.WeaponDefs(X).oWeaponDef.RelatedWeapon, Epica_Tech), oDef.ProductionRequirements.PointsRequired, oEntity)
    '								End If
    '							Next X
    '					End Select
    '				End If
    '			End If

    '		Case Else
    '			'tell the parent to cancel repair for this entity 
    '			For X As Int32 = 0 To Me.ProductionUB
    '				If Me.ProductionIdx(X) <> -1 AndAlso Me.Production(X).ProductionTypeID = ObjectType.eRepairItem Then
    '					If Me.Production(X).lExtendedID = oEntity.ObjectID AndAlso Me.Production(X).iExtendedTypeID = oEntity.ObjTypeID Then
    '						Me.ProductionIdx(X) = -1
    '					End If
    '				End If
    '			Next X
    '	End Select
    'End Sub

    '   Public Sub RepairCompleted(ByRef oProd As EntityProduction)
    '       Dim oEntity As Epica_Entity = CType(GetEpicaObject(oProd.lExtendedID, oProd.iExtendedTypeID), Epica_Entity)
    '       'Verify that it is still in the hangar (for facility repairs)
    '       Dim bFound As Boolean = False
    '       Dim bRepairMyself As Boolean = False

    '       If oEntity.ObjectID = Me.ObjectID AndAlso oEntity.ObjTypeID = Me.ObjTypeID Then
    '           bFound = True
    '           bRepairMyself = True
    '       Else
    '           For X As Int32 = 0 To Me.lHangarUB
    '               If Me.lHangarIdx(X) = oProd.lExtendedID AndAlso Me.oHangarContents(X).ObjTypeID = oProd.iExtendedTypeID Then
    '                   bFound = True
    '                   Exit For
    '               End If
    '           Next X
    '       End If

    '       If oEntity Is Nothing OrElse bFound = False Then
    '           'Ok, that entity no longer exists... or no longer is in my hangar
    '           ClearProductionByExtended(oProd.lExtendedID, oProd.iExtendedTypeID)
    '           Return
    '       End If

    '       Dim oDef As Epica_Entity_Def
    '       If oEntity.ObjTypeID = ObjectType.eUnit Then
    '           oDef = CType(oEntity, Unit).EntityDef
    '       Else : oDef = CType(oEntity, Facility).EntityDef
    '       End If
    '       If oDef Is Nothing Then Return
    '       If oDef.oPrototype Is Nothing Then Return

    '       Dim lQty As Int32 = 1

    '       Select Case oProd.ProductionID
    '           Case -2 'q1
    '               lQty = oEntity.Q1_HP
    '               If oDef.oPrototype.oArmorTech Is Nothing Then oEntity.Q1_HP += 10 Else oEntity.Q1_HP += oDef.oPrototype.oArmorTech.lHPPerPlate * 10
    '               If oEntity.Q1_HP > oDef.Q1_MaxHP Then oEntity.Q1_HP = oDef.Q1_MaxHP
    '               lQty = oEntity.Q1_HP - lQty
    '           Case -3 'q2
    '               lQty = oEntity.Q2_HP
    '               If oDef.oPrototype.oArmorTech Is Nothing Then oEntity.Q2_HP += 10 Else oEntity.Q2_HP += oDef.oPrototype.oArmorTech.lHPPerPlate * 10
    '               If oEntity.Q2_HP > oDef.Q2_MaxHP Then oEntity.Q2_HP = oDef.Q2_MaxHP
    '               lQty = oEntity.Q2_HP - lQty
    '           Case -4 'q3
    '               lQty = oEntity.Q3_HP
    '               If oDef.oPrototype.oArmorTech Is Nothing Then oEntity.Q3_HP += 10 Else oEntity.Q3_HP += oDef.oPrototype.oArmorTech.lHPPerPlate * 10
    '               If oEntity.Q3_HP > oDef.Q3_MaxHP Then oEntity.Q3_HP = oDef.Q3_MaxHP
    '               lQty = oEntity.Q3_HP - lQty
    '           Case -5 'q4
    '               lQty = oEntity.Q4_HP
    '               If oDef.oPrototype.oArmorTech Is Nothing Then oEntity.Q4_HP += 10 Else oEntity.Q4_HP += oDef.oPrototype.oArmorTech.lHPPerPlate * 10
    '               If oEntity.Q4_HP > oDef.Q4_MaxHP Then oEntity.Q4_HP = oDef.Q4_MaxHP
    '               lQty = oEntity.Q4_HP - lQty
    '           Case -6
    '               'structure
    '               oEntity.Structure_HP = oDef.Structure_MaxHP
    '           Case Is > 0
    '               'repair a component
    '               oEntity.CurrentStatus = oEntity.CurrentStatus Or oProd.ProductionID
    '       End Select

    '       If bRepairMyself = True Then
    '           'Ok, how do we do this?
    '           Dim iTemp As Int16 = CType(oEntity.ParentObject, Epica_GUID).ObjTypeID
    '           If iTemp = ObjectType.ePlanet OrElse iTemp = ObjectType.eSolarSystem Then
    '               Dim yMsg(15) As Byte
    '               System.BitConverter.GetBytes(GlobalMessageCode.eRepairCompleted).CopyTo(yMsg, 0)
    '               oEntity.GetGUIDAsString.CopyTo(yMsg, 2)

    '               System.BitConverter.GetBytes(oProd.ProductionID).CopyTo(yMsg, 8)
    '               System.BitConverter.GetBytes(lQty).CopyTo(yMsg, 12)

    '               If iTemp = ObjectType.ePlanet Then
    '                   CType(oEntity.ParentObject, Planet).oDomain.DomainSocket.SendData(yMsg)
    '               Else : CType(oEntity.ParentObject, SolarSystem).oDomain.DomainSocket.SendData(yMsg)
    '               End If
    '           End If
    '       End If



    '       'If Me.AutoLaunch = True Then
    '       '    Dim bLaunch As Boolean = True
    '       '    For X As Int32 = 0 To Me.ProductionUB
    '       '        If Me.ProductionIdx(X) <> -1 Then
    '       '            If Me.Production(X).ProductionTypeID = ObjectType.eRepairItem Then
    '       '                If Me.Production(X).lExtendedID = oProd.lExtendedID AndAlso Me.Production(X).iExtendedTypeID = oProd.iExtendedTypeID Then
    '       '                    bLaunch = False
    '       '                    Exit For
    '       '                End If
    '       '            End If
    '       '        End If
    '       '    Next X

    '       '    If bLaunch = True Then
    '       '        AddToQueue(glCurrentCycle, QueueItemType.eHandleUndockRequest_QIT, oEntity.ObjectID, oEntity.ObjTypeID, Me.ObjectID, Me.ObjTypeID)
    '       '    End If
    '       'End If

    '   End Sub

    'Private Sub ClearProductionByExtended(ByVal lExtendedID As Int32, ByVal iExtendedTypeID As Int16)
    '    If CurrentProduction Is Nothing = False Then
    '        If CurrentProduction.lExtendedID = lExtendedID AndAlso CurrentProduction.iExtendedTypeID = iExtendedTypeID Then
    '            CurrentProduction.lProdCount = 0
    '        End If
    '    End If

    '    For X As Int32 = 0 To ProductionUB
    '        If ProductionIdx(X) <> -1 AndAlso Production(X).lExtendedID = lExtendedID AndAlso Production(X).iExtendedTypeID = iExtendedTypeID Then
    '            ProductionIdx(X) = -1
    '        End If
    '    Next X
    'End Sub

	Public Sub AddAmmoProdItem(ByVal lNewProdID As Int32, ByRef oProdCost As ProductionCost, ByVal lQuantity As Int32, ByVal lExtID As Int32, ByVal iExtTypeID As Int16, ByVal yLoadAmmoType As Byte, ByVal fAmmoSize As Single, ByVal lWpnTechID As Int32)
		Dim X As Int32
		Dim lIdx As Int32 = -1

		'Always add these to the end of the queue regardless
		For X = ProductionUB To 0 Step -1
			If ProductionIdx(X) = -1 Then
				lIdx = X
				Exit For
			End If
		Next X

		'Ok, if we're here, then we have a new production, if lIDX = -1 then we need to make space 
		If lIdx = -1 Then
			ProductionUB += 1
			ReDim Preserve Production(ProductionUB)
			ReDim Preserve ProductionIdx(ProductionUB)
			Production(ProductionUB) = New EntityProduction()
			lIdx = ProductionUB
		Else
			Production(lIdx) = Nothing
			Production(lIdx) = New EntityProduction()
		End If

		ProductionIdx(lIdx) = -1

		With Production(lIdx)
			.oParent = Me
			.ProductionID = lNewProdID
			.ProductionTypeID = ObjectType.eAmmunition
			.OrderNumber = 254
			.PointsProduced = 0
			.lProdCount = lQuantity

			.lExtendedID = lExtID
			.iExtendedTypeID = iExtTypeID
			.yExtendedType = yLoadAmmoType
			.fAmmoSize = fAmmoSize
			.lAmmoWpnTechID = lWpnTechID
			.lAmmoQty = lQuantity

			.ProdCost = oProdCost

			ProductionIdx(lIdx) = -lNewProdID			 '???
		End With

		'regardless, we are now producing
		If bProducing = False Then
            If GetNextProduction(False) = False Then Return
			For X = 0 To glFacilityUB
				If glFacilityIdx(X) = Me.ObjectID Then
					AddEntityProducing(X, Me.ObjTypeID, Me.ObjectID)
					Exit For
				End If
			Next X
			RecalcProduction()
		End If
	End Sub

    Public Sub LoadFromDataReader(ByRef oData As OleDb.OleDbDataReader)
        With Me
            .ObjectID = CInt(oData("StructureID"))
            .ObjTypeID = ObjectType.eFacility
            .EntityDef = GetEpicaFacilityDef(CInt(oData("FacilityDefID")))
            .ParentObject = GetEpicaObject(CInt(oData("ParentID")), CShort(oData("ParentTypeID")))
            Debug.Assert(.ParentObject Is Nothing = False)
            .ParentColony = GetEpicaColony(CInt(oData("ColonyID")))
            If .ParentColony Is Nothing = False Then
                .ParentColony.AddChildFacility(Me)
            End If

            .CurrentSpeed = CByte(oData("CurrentSpeed"))
            .Fuel_Cap = CInt(oData("Fuel_Cap"))

            .Q1_HP = CInt(oData("Q1_HP"))
            .Q2_HP = CInt(oData("Q2_HP"))
            .Q3_HP = CInt(oData("Q3_HP"))
            .Q4_HP = CInt(oData("Q4_HP"))
            .Shield_HP = CInt(oData("Shield_HP"))
            .Structure_HP = CInt(oData("Structure_HP"))
            .ExpLevel = CByte(oData("ExpLevel"))
            .CurrentStatus = CInt(oData("CurrentStatus"))

            .EntityName = StringToBytes(CStr(oData("FacilityName")))
            .LocX = CInt(oData("LocX"))
            .LocZ = CInt(oData("LocY"))
            .LocAngle = CShort(oData("LocAngle"))

            .iCombatTactics = CInt(oData("CombatTactics"))
            .iTargetingTactics = CShort(oData("TargetingTactics"))

            .Owner = GetEpicaPlayer(CInt(oData("OwnerID")))

            .yProductionType = CByte(oData("ProductionTypeID"))
            .lReinforcingUnitGroupID = CInt(oData("ReinforcingUnitGroupID"))
            '.CurrentWorkers = CInt(oData("CurrentWorkers"))
            '.CurrentColonists = CInt(oData("CurrentColonists"))
            '.AvailableEnlisted = CInt(oData("CurrentEnlisted"))
            '.AvailableOfficers = CInt(oData("CurrentOfficers"))

            '.UniversalFacilityID = oData("UniversalFacilityID")

            .AutoLaunch = CByte(oData("AutoLaunch")) <> 0
            .bExcludeFromColonyQueue = CInt(oData("ExcludeFromColonyQueue")) <> 0

            'If (.CurrentStatus And elUnitStatus.eFacilityPowered) <> 0 Then
            '    .CurrentStatus -= elUnitStatus.eFacilityPowered
            '    .SetActive(True)
            'End If

            If .EntityDef Is Nothing = False Then
                ReDim .lCurrentAmmo(.EntityDef.WeaponDefUB)
                For lTemp As Int32 = 0 To .lCurrentAmmo.Length - 1
                    .lCurrentAmmo(lTemp) = -1
                Next lTemp
            End If

            If .Owner Is Nothing = False AndAlso .Owner.ObjectID = gl_HARDCODE_PIRATE_PLAYER_ID Then
                If .yProductionType = ProductionType.eProduction Then
                    .lPirate_For_PlayerID = CInt(oData("CurrentWorkers"))
                    AddPirateProductionItem(Me)
                ElseIf .yProductionType = ProductionType.eTradePost Then
                    If .ParentObject Is Nothing = False AndAlso CType(.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eSolarSystem Then
                        CType(.ParentObject, SolarSystem).lNPCTradepostID = .ObjectID
                    End If
                End If
            End If

            .PreviousActive = .Active

            .RallyPointX = CInt(oData("RallyPointX"))
            .RallyPointZ = CInt(oData("RallyPointZ"))
            .RallyPointA = CShort(oData("RallyPointA"))
            .RallyPointEnvirID = CInt(oData("RallyPointEnvirID"))
            .RallyPointEnvirTypeID = CShort(oData("RallyPointEnvirTypeID"))

            .bNeedsSaved = False

        End With
    End Sub

    Public Sub fac_DONOTCALL_AddRepairItem(ByVal lRepairItemID As Int32, ByVal lQuantity As Int32, ByRef oComponent As Epica_Tech, ByVal blEntityDefPointsRequired As Int64, ByRef oEntity As Epica_Entity)
        Dim X As Int32
        Dim bFound As Boolean = False
        Dim lIdx As Int32 = -1

        If bFound = False Then
            For X = ProductionUB To 0 Step -1
                If ProductionIdx(X) <> -1 Then
                    If ProductionIdx(X) = lRepairItemID AndAlso Production(X).ProductionTypeID = ObjectType.eRepairItem Then
                        'Already there... increment the quantity
                        bFound = True
                        Production(X).lProdCount += lQuantity
                    End If

                    Exit For
                End If
            Next

            If bFound = False Then
                For X = ProductionUB To 0 Step -1
                    If ProductionIdx(X) = -1 Then
                        lIdx = X
                        Exit For
                    End If
                Next X
            End If
        End If

        'Ok, if we're here, then we have a new production, if lIDX = -1 then we need to make space
        If bFound = False Then
            If lIdx = -1 Then
                ProductionUB += 1
                ReDim Preserve Production(ProductionUB)
                ReDim Preserve ProductionIdx(ProductionUB)
                Production(ProductionUB) = New EntityProduction()
                lIdx = ProductionUB
            Else
                Production(lIdx) = Nothing
                Production(lIdx) = New EntityProduction()
            End If

            ProductionIdx(lIdx) = -1
            Production(lIdx) = CreateRepairProduction(lRepairItemID, lQuantity, oComponent, blEntityDefPointsRequired, oEntity)
            ProductionIdx(lIdx) = lRepairItemID
        End If

        'regardless, we are now producing
        If bProducing = False Then
            If GetNextProduction(False) = False Then
                Return
            End If
            For X = 0 To glFacilityUB
                If glFacilityIdx(X) = Me.ObjectID Then
                    AddEntityProducing(X, Me.ObjTypeID, Me.ObjectID)
                    Exit For
                End If
            Next X
            RecalcProduction()
        End If
    End Sub

    Public Sub DismantleChildFacility(ByRef oFac As Facility)
        'dismantling a station module
        If Me.ParentColony Is Nothing OrElse oFac.ParentColony Is Nothing Then Return
        If Me.ParentColony.ObjectID <> oFac.ParentColony.ObjectID Then Return

        'Ok, let's do this... get our recycled numbers
        Dim col As Collection = oFac.RecycleMe(oFac.EntityDef)
        If col.Count > 0 Then

            For Each oItem As RecycledPartsList In col
                Dim lCargoSpaceReq As Int32 = oItem.lQuantity
                If ParentColony.Cargo_Cap > 0 Then
                    Dim lQty As Int32 = Math.Min(ParentColony.Cargo_Cap, lCargoSpaceReq)
                    ParentColony.AdjustColonyComponentCache(oItem.lObjID, oItem.iObjTypeID, oItem.lExtended, lQty)
                    If lQty <> lCargoSpaceReq Then Exit For
                End If
            Next
        End If

        'Now, remove the facility
        With oFac
            .RemoveMe()
            .DeleteEntity(.ServerIndex)
            If .ServerIndex > -1 Then
                If glFacilityIdx(.ServerIndex) <> -1 AndAlso glFacilityIdx(.ServerIndex) = .ObjectID Then
                    glFacilityIdx(.ServerIndex) = -1
                    RemoveLookupFacility(.ObjectID, .ObjTypeID)
                End If
            End If
        End With
    End Sub
    Public Sub DismantleChildUnit(ByRef oUnit As Unit)
        If Me.ParentColony Is Nothing Then Return

        Dim col As Collection = oUnit.RecycleMe(oUnit.EntityDef)
        If col.Count > 0 Then
            For Each oItem As RecycledPartsList In col
                Dim lCargoSpaceReq As Int32 = oItem.lQuantity
                If ParentColony.Cargo_Cap > 0 Then
                    Dim lQty As Int32 = Math.Min(ParentColony.Cargo_Cap, lCargoSpaceReq)
                    ParentColony.AdjustColonyComponentCache(oItem.lObjID, oItem.iObjTypeID, oItem.lExtended, lQty)
                    If lQty <> lCargoSpaceReq Then Exit For
                End If
            Next
        End If
    End Sub
End Class
