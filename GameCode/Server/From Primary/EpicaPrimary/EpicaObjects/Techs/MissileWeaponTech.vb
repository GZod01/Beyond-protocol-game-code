Option Strict On

Public Class MissileWeaponTech
    Inherits BaseWeaponTech

    Private Const mfHullSizeToMaxDamageFactor As Single = 10.0F

    Public ROF As Int16
    Public MaxDmg As Int32
    Public MissileHullSize As Int32
    Public MaxSpeed As Byte
    Public Maneuver As Byte
    Public FlightTime As Int16
    Public HomingAccuracy As Byte
    Public PayloadType As Byte
    Public ExplosionRadius As Byte
    Public StructureHP As Int32

    Public lBodyMineralID As Int32
    Public lNoseMineralID As Int32
    Public lFlapsMineralID As Int32
    Public lFuelMineralID As Int32
    Public lPayloadMineralID As Int32

    'Public blAmmoCostCredits As Int64 = 0
    'Public blAmmoCostPoints As Int64 = 0
    'Public fAmmoMin1Cost As Single = 0.0F
    'Public fAmmoMin2Cost As Single = 0.0F
    'Public fAmmoMin3Cost As Single = 0.0F
    'Public fAmmoMin4Cost As Single = 0.0F
    'Public fAmmoMin5Cost As Single = 0.0F

#Region "  Helpers  "
    Private moBodyMineral As Mineral = Nothing
    Public ReadOnly Property BodyMineral() As Mineral
        Get
            If moBodyMineral Is Nothing OrElse moBodyMineral.ObjectID <> lBodyMineralID Then moBodyMineral = GetEpicaMineral(lBodyMineralID)
            Return moBodyMineral
        End Get
    End Property
    Private moNoseMineral As Mineral = Nothing
    Public ReadOnly Property NoseMineral() As Mineral
        Get
            If moNoseMineral Is Nothing OrElse moNoseMineral.ObjectID <> lNoseMineralID Then moNoseMineral = GetEpicaMineral(lNoseMineralID)
            Return moNoseMineral
        End Get
    End Property
    Private moFlapsMineral As Mineral = Nothing
    Public ReadOnly Property FlapsMineral() As Mineral
        Get
            If moFlapsMineral Is Nothing OrElse moFlapsMineral.ObjectID <> lFlapsMineralID Then moFlapsMineral = GetEpicaMineral(lFlapsMineralID)
            Return moFlapsMineral
        End Get
    End Property
    Private moFuelMineral As Mineral = Nothing
    Public ReadOnly Property FuelMineral() As Mineral
        Get
            If moFuelMineral Is Nothing OrElse moFuelMineral.ObjectID <> lFuelMineralID Then moFuelMineral = GetEpicaMineral(lFuelMineralID)
            Return moFuelMineral
        End Get
    End Property
    Private moPayloadMineral As Mineral = Nothing
    Public ReadOnly Property PayloadMineral() As Mineral
        Get
            If moPayloadMineral Is Nothing OrElse moPayloadMineral.ObjectID <> lPayloadMineralID Then moPayloadMineral = GetEpicaMineral(lPayloadMineralID)
            Return moPayloadMineral
        End Get
    End Property
#End Region

    Public Function GetObjAsString() As Byte()
        If moResearchCost Is Nothing OrElse moProductionCost Is Nothing Then
            CalculateBothProdCosts()
            mbStringReady = False
        End If

        'here we will return the entire object as a string
        'If mbStringReady = False Then
        Dim yDefMsg() As Byte = Nothing
        If Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eResearched Then
            Dim oDef As WeaponDef = Me.GetWeaponDefResult()
            If oDef Is Nothing = False Then yDefMsg = oDef.GetObjAsString()
        End If
        If yDefMsg Is Nothing = False Then
            ReDim mySendString(BASE_WEAPON_MSG_HEADER_LENGTH + yDefMsg.Length + 40) '76)
        Else : ReDim mySendString(BASE_WEAPON_MSG_HEADER_LENGTH + 40) '76)
        End If

        Dim lPos As Int32 = MyBase.FillBaseWeaponMsgHdr()

        System.BitConverter.GetBytes(ROF).CopyTo(mySendString, lPos) : lPos += 2
        System.BitConverter.GetBytes(MaxDmg).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(MissileHullSize).CopyTo(mySendString, lPos) : lPos += 4
        mySendString(lPos) = MaxSpeed : lPos += 1
        mySendString(lPos) = Maneuver : lPos += 1
        System.BitConverter.GetBytes(FlightTime).CopyTo(mySendString, lPos) : lPos += 2
        mySendString(lPos) = HomingAccuracy : lPos += 1
        mySendString(lPos) = PayloadType : lPos += 1
        mySendString(lPos) = ExplosionRadius : lPos += 1
        System.BitConverter.GetBytes(StructureHP).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lBodyMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lNoseMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lFlapsMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lFuelMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lPayloadMineralID).CopyTo(mySendString, lPos) : lPos += 4

        'System.BitConverter.GetBytes(blAmmoCostCredits).CopyTo(mySendString, lPos) : lPos += 8
        'System.BitConverter.GetBytes(blAmmoCostPoints).CopyTo(mySendString, lPos) : lPos += 8
        'System.BitConverter.GetBytes(fAmmoMin1Cost).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(fAmmoMin2Cost).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(fAmmoMin3Cost).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(fAmmoMin4Cost).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(fAmmoMin5Cost).CopyTo(mySendString, lPos) : lPos += 4 

        If yDefMsg Is Nothing = False Then
            yDefMsg.CopyTo(mySendString, lPos) : lPos += yDefMsg.Length
        End If

        'mbStringReady = True
        'End If
        Return mySendString
    End Function

    Protected Overrides Sub CalculateBothProdCosts()
        If moProductionCost Is Nothing OrElse moResearchCost Is Nothing Then ComponentDesigned()
    End Sub

    'Public Overrides Sub ComponentDesigned()
    '	If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
    '	With moResearchCost
    '		.ObjectID = Me.ObjectID
    '		.ObjTypeID = Me.ObjTypeID
    '		Erase .ItemCosts
    '		.ItemCostUB = -1
    '	End With
    '	If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
    '	With moProductionCost
    '		.ObjectID = Me.ObjectID
    '		.ObjTypeID = Me.ObjTypeID
    '		Erase .ItemCosts
    '		.ItemCostUB = -1
    '	End With

    '	Try
    '		Dim fHullSizeDamage As Single = CSng(MaxDmg / MissileHullSize) / mfHullSizeToMaxDamageFactor
    '		Dim fAll4 As Single = GetLowLookup((MaxSpeed * MissileHullSize \ 100000) + 1, 1)
    '		fAll4 += GetLowLookup((Maneuver * MissileHullSize \ 10000) + 1, 1)
    '		Dim fDPS As Single
    '		Dim fROFSeconds As Single = 0
    '		Dim fFltTimeSeconds As Single = FlightTime / 30.0F

    '		Dim bNotStudyFlaw As Boolean = False

    '		If ROF < 1 Then
    '			fDPS = MaxDmg / 100.0F
    '			fROFSeconds = -1
    '		Else
    '			fROFSeconds = ROF / 30.0F
    '			fDPS = MaxDmg / fROFSeconds
    '		End If

    '		Dim fResearchCost As Single = 0.0F
    '		Dim fProductionCost As Single = 0.0F
    '		Dim fSuccess As Single = 0.0F
    '		Dim lSuccessCnt As Int32 = 0

    '		Dim fImpactDamage As Single = Math.Min(MissileHullSize * MaxSpeed, (MaxDmg * 0.45F))
    '		Dim fPayloadDamage As Single = Math.Max(MaxDmg - fImpactDamage, (MaxDmg * 0.1F))

    '		Dim lMaxCost As Int32 = 0
    '		Dim lMaxSuccess As Int32 = 0

    '		'BODY MATERIAL BEGIN!!
    '		Dim lBodyComplexity As Int32 = BodyMineral.GetPropertyValue(eMinPropID.Complexity)
    '		'Body.Hardness
    '		Dim uBodyHardness As MaterialPropertyItem
    '		With uBodyHardness
    '			.lMineralID = lBodyMineralID
    '			.lPropertyID = eMinPropID.Hardness

    '			.lActualValue = BodyMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)

    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If

    '			.fMinAmt = ((.lActualValue - .lKnownValue + 1) / 100.0F) * fHullSizeDamage * GetHighLookup((.lActualValue \ 10) - 1, 0) / 100.0F * (MaxDmg / 100.0F)
    '			.fSuccess = ((100.0F / fHullSizeDamage) * Math.Abs(1 - CSng(.lKnownValue / (.lActualValue + 1))) + 1)
    '			.fPower = 0 : .fHull = 0 : .lEnlisted = 0 : .lOfficer = 0
    '			.fResearch = GetLowLookup(CInt(Math.Ceiling((.AKScore - 0.99F) * 10)), 3) * lBodyComplexity * fHullSizeDamage * MaxDmg * 10.0F
    '			.fProduction = CSng(MaxDmg / .lActualValue) * 100.0F * fHullSizeDamage

    '			fResearchCost += .fResearch
    '			fProductionCost += .fProduction
    '			fSuccess += .fSuccess
    '			lSuccessCnt += 1

    '			If bNotStudyFlaw = False AndAlso .fSuccess < lMaxSuccess OrElse (.fSuccess = lMaxSuccess AndAlso .fResearch > lMaxCost) Then
    '				lMaxSuccess = CInt(.fSuccess)
    '				lMaxCost = CInt(.fResearch)
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop1
    '				'If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '			End If
    '		End With
    '		'Body.Density
    '		Dim uBodyDensity As MaterialPropertyItem
    '		With uBodyDensity
    '			.lMineralID = lBodyMineralID
    '			.lPropertyID = eMinPropID.Density

    '			.lActualValue = BodyMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)

    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop2 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If

    '			.fMinAmt = (Math.Abs(1 - CSng((.lKnownValue + 1) / (.lActualValue + 1))) + 1) * fAll4 * GetLowLookup((.lActualValue \ 10) + 1, 0) / 200.0F
    '			.fSuccess = ((fAll4 / 100.0F) * Math.Abs(1 - CSng((.lKnownValue + 1) / (.lActualValue + 1)) + 1) + 1) - .fMinAmt
    '			.fPower = 0 : .fHull = 0 : .lEnlisted = 0 : .lOfficer = 0
    '			.fResearch = GetLowLookup(CInt(.AKScore), 3) * lBodyComplexity * 1000.0F
    '			.fProduction = .lActualValue * fAll4

    '			fResearchCost += .fResearch
    '			fProductionCost += .fProduction
    '			fSuccess += .fSuccess
    '			lSuccessCnt += 1

    '			If bNotStudyFlaw = False AndAlso .fSuccess < lMaxSuccess OrElse (.fSuccess = lMaxSuccess AndAlso .fResearch > lMaxCost) Then
    '				lMaxSuccess = CInt(.fSuccess)
    '				lMaxCost = CInt(.fResearch)
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop2
    '				'If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '			End If
    '		End With
    '		'Body.Malleable
    '		Dim uBodyMalleable As MaterialPropertyItem
    '		With uBodyMalleable
    '			.lMineralID = lBodyMineralID
    '			.lPropertyID = eMinPropID.Malleable

    '			.lActualValue = BodyMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)

    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop3 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If

    '			.fMinAmt = (Math.Abs(1 - CSng(.lKnownValue / .lActualValue)) + 1) * fHullSizeDamage * GetHighLookup((.lActualValue \ 10) - 1, 0) / 1000.0F
    '			.fSuccess = ((fAll4 / 100.0F) * Math.Abs(1 - CSng(.lKnownValue / .lActualValue) + 1) + 1) - (.fMinAmt / 10.0F)
    '			.fPower = 0 : .fHull = 0 : .lEnlisted = 0 : .lOfficer = 0
    '			.fResearch = GetLowLookup(CInt(Math.Ceiling((.AKScore - 0.99F) * 10)), 3) * 1000.0F * lBodyComplexity
    '			.fProduction = CSng(MaxDmg / .lActualValue) * 10000.0F

    '			fResearchCost += .fResearch
    '			fProductionCost += .fProduction
    '			fSuccess += .fSuccess
    '			lSuccessCnt += 1

    '			If bNotStudyFlaw = False AndAlso .fSuccess < lMaxSuccess OrElse (.fSuccess = lMaxSuccess AndAlso .fResearch > lMaxCost) Then
    '				lMaxSuccess = CInt(.fSuccess)
    '				lMaxCost = CInt(.fResearch)
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop3
    '				'If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '			End If
    '		End With

    '		'NOSE MATERIAL BEGIN
    '		Dim lNoseComplexity As Int32 = NoseMineral.GetPropertyValue(eMinPropID.Complexity)
    '		'Nose.Hardness
    '		Dim uNoseHardness As MaterialPropertyItem
    '		With uNoseHardness
    '			.lMineralID = lNoseMineralID
    '			.lPropertyID = eMinPropID.Hardness

    '			.lActualValue = NoseMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)

    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If

    '			.fMinAmt = (Math.Abs(1 - CSng(.lKnownValue / .lActualValue)) + 1) * CSng(fImpactDamage / .lActualValue) * GetHighLookup((.lActualValue \ 10) - 1, 0) / 100.0F
    '			.fSuccess = ((fHullSizeDamage / 100.0F) * Math.Abs(1 - CSng(.lKnownValue / .lActualValue) + 1) + 1) - .fMinAmt
    '			.fPower = 0 : .fHull = 0 : .lEnlisted = 0 : .lOfficer = 0
    '			.fResearch = GetLowLookup(CInt(.AKScore), 3) * 1000.0F * lNoseComplexity
    '			.fProduction = .lActualValue * fAll4

    '			fResearchCost += .fResearch
    '			fProductionCost += .fProduction
    '			fSuccess += .fSuccess
    '			lSuccessCnt += 1

    '			If bNotStudyFlaw = False AndAlso .fSuccess < lMaxSuccess OrElse (.fSuccess = lMaxSuccess AndAlso .fResearch > lMaxCost) Then
    '				lMaxSuccess = CInt(.fSuccess)
    '				lMaxCost = CInt(.fResearch)
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop1
    '				'If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '			End If
    '		End With
    '		'Nose.Compress
    '		Dim uNoseCompress As MaterialPropertyItem
    '		With uNoseCompress
    '			.lMineralID = lNoseMineralID
    '			.lPropertyID = eMinPropID.Compressibility

    '			.lActualValue = NoseMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)

    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop2 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If

    '			.fMinAmt = (Math.Abs(1 - CSng((.lKnownValue + 1) / (.lActualValue + 1))) + 1) * fAll4 * GetLowLookup((.lActualValue \ 10) + 1, 0) / 100.0F
    '			.fSuccess = ((fAll4 / 100.0F) * Math.Abs(1 - CSng((.lKnownValue + 1) / (.lActualValue + 1)) + 1) + 1) - .fMinAmt
    '			.fPower = 0 : .fHull = 0 : .lEnlisted = 0 : .lOfficer = 0
    '			.fResearch = GetLowLookup(CInt(Math.Ceiling((.AKScore - 0.99F) * 10)), 3) * 1000.0F * lNoseComplexity * fAll4
    '			.fProduction = .lActualValue * fAll4

    '			fResearchCost += .fResearch
    '			fProductionCost += .fProduction
    '			fSuccess += .fSuccess
    '			lSuccessCnt += 1

    '			If bNotStudyFlaw = False AndAlso .fSuccess < lMaxSuccess OrElse (.fSuccess = lMaxSuccess AndAlso .fResearch > lMaxCost) Then
    '				lMaxSuccess = CInt(.fSuccess)
    '				lMaxCost = CInt(.fResearch)
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop2
    '				'If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '			End If
    '		End With

    '		'FLAPS MATERIAL BEGIN!!
    '		Dim lFlapsComplexity As Int32 = FlapsMineral.GetPropertyValue(eMinPropID.Complexity)
    '		'Flaps.Density
    '		Dim uFlapsDensity As MaterialPropertyItem
    '		With uFlapsDensity
    '			.lMineralID = lFlapsMineralID
    '			.lPropertyID = eMinPropID.Density

    '			.lActualValue = FlapsMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)

    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If

    '			.fMinAmt = (Math.Abs(1 - CSng(.lKnownValue / .lActualValue)) + 1) * fAll4 * GetHighLookup((.lActualValue \ 10) - 1, 0) / 100.0F
    '			.fSuccess = ((fAll4 / 100.0F) * Math.Abs(1 - CSng(.lKnownValue / .lActualValue) + 1) + 1) - .fMinAmt
    '			.fPower = 0 : .fHull = 0 : .lEnlisted = 0 : .lOfficer = 0
    '			.fResearch = GetLowLookup(CInt(.AKScore), 3) * 1000.0F * lFlapsComplexity
    '			.fProduction = GetLowLookup(CInt(Math.Ceiling(Maneuver / 3.0F)), 3) * MaxSpeed * (255 - .lActualValue)

    '			fResearchCost += .fResearch
    '			fProductionCost += .fProduction
    '			fSuccess += .fSuccess
    '			lSuccessCnt += 1

    '			If bNotStudyFlaw = False AndAlso .fSuccess < lMaxSuccess OrElse (.fSuccess = lMaxSuccess AndAlso .fResearch > lMaxCost) Then
    '				lMaxSuccess = CInt(.fSuccess)
    '				lMaxCost = CInt(.fResearch)
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop1
    '				'If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '			End If
    '		End With
    '		'Flaps.ThermalExpansion
    '		Dim uFlapsThermExp As MaterialPropertyItem
    '		With uFlapsThermExp
    '			.lMineralID = lFlapsMineralID
    '			.lPropertyID = eMinPropID.ThermalExpansion

    '			.lActualValue = FlapsMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)

    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop2 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If

    '			.fMinAmt = (Math.Abs(1 - CSng((.lKnownValue + 1) / (.lActualValue + 1))) + 1) * fAll4 * CSng(Maneuver) * GetLowLookup((.lActualValue \ 10) + 1, 0) / 1000.0F
    '			.fSuccess = ((fAll4 / 100.0F) * Math.Abs(1 - CSng((.lKnownValue + 1) / (.lActualValue + 1)) + 1) + 1) - .fMinAmt
    '			.fPower = 0 : .fHull = 0 : .lEnlisted = 0 : .lOfficer = 0
    '			.fResearch = GetLowLookup(CInt(Math.Ceiling((.AKScore - 0.99F) * 10)), 3) * 1000.0F * lFlapsComplexity * CInt(Math.Ceiling(Maneuver / 10.0F))
    '			.fProduction = .lActualValue * fAll4

    '			fResearchCost += .fResearch
    '			fProductionCost += .fProduction
    '			fSuccess += .fSuccess
    '			lSuccessCnt += 1

    '			If bNotStudyFlaw = False AndAlso .fSuccess < lMaxSuccess OrElse (.fSuccess = lMaxSuccess AndAlso .fResearch > lMaxCost) Then
    '				lMaxSuccess = CInt(.fSuccess)
    '				lMaxCost = CInt(.fResearch)
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop2
    '				'If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '			End If
    '		End With
    '		'Flaps.TempSens
    '		Dim uFlapsTempSens As MaterialPropertyItem
    '		With uFlapsTempSens
    '			.lMineralID = lFlapsMineralID
    '			.lPropertyID = eMinPropID.TemperatureSensitivity

    '			.lActualValue = FlapsMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)

    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop3 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If

    '			.fMinAmt = (Math.Abs(1 - CSng((.lKnownValue + 1) / (.lActualValue + 1))) + 1) * fAll4 * GetLowLookup((.lActualValue \ 10) + 1, 0) / 20.0F
    '			.fSuccess = ((fAll4 / 100.0F) * Math.Abs(1 - CSng((.lKnownValue + 1) / (.lActualValue + 1)) + 1) + 1) - .fMinAmt
    '			.fPower = 0 : .fHull = 0 : .lEnlisted = 0 : .lOfficer = 0
    '			.fResearch = GetLowLookup(CInt(.AKScore), 3) * 1000.0F * lFlapsComplexity
    '			.fProduction = .lActualValue * fAll4

    '			fResearchCost += .fResearch
    '			fProductionCost += .fProduction
    '			fSuccess += .fSuccess
    '			lSuccessCnt += 1

    '			If bNotStudyFlaw = False AndAlso .fSuccess < lMaxSuccess OrElse (.fSuccess = lMaxSuccess AndAlso .fResearch > lMaxCost) Then
    '				lMaxSuccess = CInt(.fSuccess)
    '				lMaxCost = CInt(.fResearch)
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop3
    '				'If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '			End If
    '		End With

    '		'BEGIN FUEL MATERIAL
    '		Dim lFuelComplexity As Int32 = FuelMineral.GetPropertyValue(eMinPropID.Complexity)
    '		'Fuel.ChemReact
    '		Dim uFuelChemReact As MaterialPropertyItem
    '		With uFuelChemReact
    '			.lMineralID = lFuelMineralID
    '			.lPropertyID = eMinPropID.ChemicalReactance

    '			.lActualValue = FuelMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)

    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If

    '			.fMinAmt = (Math.Abs(1 - CSng(.lKnownValue / .lActualValue)) + 1) * fAll4 * GetHighLookup((.lActualValue \ 10) - 1, 0) / 100.0F
    '			.fSuccess = ((fAll4 / 100.0F) * Math.Abs(1 - CSng(.lKnownValue / .lActualValue) + 1) + 1)
    '			.fPower = 0 : .fHull = 0 : .lOfficer = 0

    '			If fROFSeconds < 3 AndAlso fROFSeconds > 0 Then
    '				.lEnlisted = CInt(5 - fROFSeconds)
    '			Else : .lEnlisted = 0
    '			End If
    '			If fROFSeconds < 0 Then
    '				.fResearch = 0
    '			Else : .fResearch = GetHighLookup(CInt(fDPS / 10.0F), 3)
    '			End If
    '			.fProduction = CSng(MaxDmg / .lActualValue) * 10000.0F

    '			fResearchCost += .fResearch
    '			fProductionCost += .fProduction
    '			fSuccess += .fSuccess
    '			lSuccessCnt += 1

    '			If bNotStudyFlaw = False AndAlso .fSuccess < lMaxSuccess OrElse (.fSuccess = lMaxSuccess AndAlso .fResearch > lMaxCost) Then
    '				lMaxSuccess = CInt(.fSuccess)
    '				lMaxCost = CInt(.fResearch)
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop1
    '				'If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '			End If
    '		End With
    '		'Fuel.BoilingPt
    '		Dim uFuelBoilingPt As MaterialPropertyItem
    '		With uFuelBoilingPt
    '			.lMineralID = lFuelMineralID
    '			.lPropertyID = eMinPropID.BoilingPoint

    '			.lActualValue = FuelMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)

    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop2 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If

    '			.fMinAmt = (Math.Abs(1 - CSng((.lKnownValue + 1) / (.lActualValue + 1))) + 1) * fAll4 * CSng(Maneuver) * GetLowLookup((.lActualValue \ 10) + 1, 0) / 1000.0F
    '			.fSuccess = ((fAll4 / 100.0F) * Math.Abs(1 - CSng((.lKnownValue + 1) / (.lActualValue + 1)) + 1) + 1) - .fMinAmt
    '			.fPower = 0 : .fHull = 0 : .lOfficer = 0

    '			If MaxDmg > 300 AndAlso fROFSeconds > 0 Then
    '				.lEnlisted = uFuelChemReact.lKnownValue \ .lKnownValue
    '			Else : .lEnlisted = 0
    '			End If
    '			.fResearch = GetLowLookup(CInt(.AKScore), 3) * 1000.0F * lFuelComplexity * (.lEnlisted + 1)
    '			.fProduction = .lActualValue * fAll4

    '			fResearchCost += .fResearch
    '			fProductionCost += .fProduction
    '			fSuccess += .fSuccess
    '			lSuccessCnt += 1

    '			If bNotStudyFlaw = False AndAlso .fSuccess < lMaxSuccess OrElse (.fSuccess = lMaxSuccess AndAlso .fResearch > lMaxCost) Then
    '				lMaxSuccess = CInt(.fSuccess)
    '				lMaxCost = CInt(.fResearch)
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop2
    '				'If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '			End If
    '		End With
    '		'Fuel.Density
    '		Dim uFuelDensity As MaterialPropertyItem
    '		With uFuelDensity
    '			.lMineralID = lFuelMineralID
    '			.lPropertyID = eMinPropID.Density

    '			.lActualValue = FuelMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)

    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop3 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If

    '			.fMinAmt = 0 : .fSuccess = 0 : .fPower = 0 : .fHull = 0 : .lEnlisted = 0 : .lOfficer = 0
    '			.fResearch = GetLowLookup(CInt(Math.Ceiling((.AKScore - 0.99F) * 10)), 3) * 1000.0F * lFuelComplexity * CInt(Math.Ceiling(MaxSpeed / 10.0F))
    '			.fProduction = GetLowLookup(CInt(Math.Ceiling(.lKnownValue / 8.0F)), 3) * .lActualValue
    '			fResearchCost += .fResearch
    '			fProductionCost += .fProduction

    '			If bNotStudyFlaw = False AndAlso .fSuccess < lMaxSuccess OrElse (.fSuccess = lMaxSuccess AndAlso .fResearch > lMaxCost) Then
    '				lMaxSuccess = CInt(.fSuccess)
    '				lMaxCost = CInt(.fResearch)
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop3
    '				'If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '			End If
    '		End With

    '		'BEGIN PAYLOAD MATERIAL
    '		Dim lPayloadComplexity As Int32 = PayloadMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim uPayloadProp As MaterialPropertyItem
    '		With uPayloadProp
    '			.lMineralID = lPayloadMineralID
    '			If PayloadType = 0 Then
    '				.lPropertyID = eMinPropID.Combustiveness
    '			Else : .lPropertyID = eMinPropID.ChemicalReactance
    '			End If

    '			.lActualValue = PayloadMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)

    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat5_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If

    '			.fMinAmt = (Math.Abs(1 - CSng(.lKnownValue / .lActualValue)) + 1) * ((fPayloadDamage * CSng(ExplosionRadius + 1)) / .lActualValue) * GetHighLookup((.lActualValue \ 10) - 1, 0) / 100
    '			.fSuccess = ((fHullSizeDamage / 100.0F) * Math.Abs(1 - CSng(.lKnownValue / .lActualValue) + 1) + 1) - .fMinAmt
    '			.fPower = 0 : .fHull = 0 : .lEnlisted = 0 : .lOfficer = 0
    '			.fResearch = GetLowLookup(CInt(.AKScore), 3) * lPayloadComplexity * fHullSizeDamage * MaxDmg
    '			.fProduction = CSng(MaxDmg / .lActualValue) * 100.0F * fHullSizeDamage

    '			fResearchCost += .fResearch
    '			fProductionCost += .fProduction
    '			fSuccess += .fSuccess
    '			lSuccessCnt += 1

    '			If bNotStudyFlaw = False AndAlso .fSuccess < lMaxSuccess OrElse (.fSuccess = lMaxSuccess AndAlso .fResearch > lMaxCost) Then
    '				lMaxSuccess = CInt(.fSuccess)
    '				lMaxCost = CInt(.fResearch)
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat5_Prop1
    '				'If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '			End If
    '		End With

    '		'Now, for our other values...
    '		Dim fModifiedRandomSeed As Single = (RandomSeed * 0.7F) + 0.8F			'between .8 and 1.5

    '		Me.CurrentSuccessChance = CInt(fSuccess / lSuccessCnt)
    '		Me.SuccessChanceIncrement = 1
    '		Me.PowerRequired = CInt(CSng(MaxDmg / 100.0F) * fModifiedRandomSeed)
    '		If fROFSeconds < 0 Then
    '			Me.HullRequired = Me.MissileHullSize
    '		Else : Me.HullRequired = CInt(MissileHullSize * 3 * fModifiedRandomSeed)
    '		End If

    '		Dim lEnlistedCost As Int32 = uFuelChemReact.lEnlisted + uFuelBoilingPt.lEnlisted
    '		Dim lOfficerCost As Int32 = lEnlistedCost \ 10

    '		'Ok, determine our difficulty modifier... to this, we determine estimated hull usage percentages
    '		Dim fFuelHullUsage As Single = (fFltTimeSeconds * CSng(MissileHullSize) * CSng(MaxSpeed))
    '		fFuelHullUsage /= ((uFuelChemReact.lActualValue + uFuelChemReact.lKnownValue) * 2.0F)
    '		fFuelHullUsage = CSng(Math.Sqrt(fFuelHullUsage) / MissileHullSize)
    '		Dim fBodyHullUsage As Single = 0.2F + (StructureHP / (3.0F * ((uBodyHardness.lActualValue + uBodyHardness.lKnownValue) / 2.0F)))
    '		Dim fNavHullUsage As Single = 0.1F + ((HomingAccuracy / 255.0F) * 0.35F)
    '		Dim fRemainder As Single = 1.0F - (fFuelHullUsage + fBodyHullUsage + fNavHullUsage)

    '		Dim fDifficultyModifier As Single
    '		Dim fAvailableSize As Single = MissileHullSize * fRemainder
    '		Dim fRequiredSize As Single = uPayloadProp.fMinAmt

    '		'Now, adjust the available and required in case on is below 0
    '		If fAvailableSize < 0 OrElse fRequiredSize < 0 Then
    '			fAvailableSize += Math.Abs(Math.Min(fAvailableSize, fRequiredSize)) + 1
    '			fRequiredSize += Math.Abs(Math.Min(fAvailableSize, fRequiredSize)) + 1
    '		End If
    '		'Now, use those values to get our difficulty modifier
    '		If fPayloadDamage > 1 Then fDifficultyModifier = fRequiredSize / fAvailableSize Else fDifficultyModifier = 1
    '		fDifficultyModifier = Math.Max(fDifficultyModifier, 0.6F)

    '		'Now, get our ItoAC value
    '		Dim fItoAC As Single = (lBodyComplexity + lNoseComplexity + lFlapsComplexity + lFuelComplexity + lPayloadComplexity) / 5.0F
    '		fItoAC = Math.Max(1, fItoAC - (PopIntel))

    '		'Ok, we can now get the remaining values
    '		Dim blResearchCredits As Int64 = CLng(fResearchCost * fDifficultyModifier * fModifiedRandomSeed * Math.Max(GetLowLookup(CInt(Math.Ceiling(HomingAccuracy / 10.0F)), 1), 1))
    '		If blResearchCredits < 100 Then blResearchCredits = 100

    '		Dim blResearchPoints As Int64 = CLng(blResearchCredits * fItoAC / 10)
    '		blResearchPoints = GetResearchPointsCost(blResearchPoints)
    '		blResearchCredits *= Me.ResearchAttempts

    '		Dim blProdCredits As Int64 = CLng(((fProductionCost * (fDifficultyModifier / 3.0F)) + 10000) * fModifiedRandomSeed)
    '		Dim blProdPoints As Int64 = CLng((blProdCredits * CInt(Math.Ceiling(fItoAC)) \ 10) * GetLowLookup(CInt(Math.Ceiling(StructureHP / 7.0F)), 3))
    '		blProdPoints += CInt((MaxDmg + MissileHullSize + CInt(MaxSpeed) + CInt(Maneuver) + fROFSeconds + fFltTimeSeconds + CInt(HomingAccuracy) + CInt(PayloadType) + CInt(ExplosionRadius) + StructureHP) * 1000)

    '		Dim lBodyCost As Int32 = CInt(Math.Ceiling((uBodyHardness.fMinAmt + uBodyDensity.fMinAmt + uBodyMalleable.fMinAmt) * (fBodyHullUsage * MissileHullSize) * fModifiedRandomSeed))
    '		Dim lNoseCost As Int32 = CInt(Math.Ceiling((uNoseHardness.fMinAmt + uNoseCompress.fMinAmt) * fModifiedRandomSeed))
    '		Dim lFlapsCost As Int32 = CInt(Math.Ceiling((uFlapsDensity.fMinAmt + uFlapsThermExp.fMinAmt + uFlapsTempSens.fMinAmt) * (fNavHullUsage * MissileHullSize) * fModifiedRandomSeed))
    '		Dim lFuelCost As Int32 = CInt(Math.Ceiling((uFuelChemReact.fMinAmt + uFuelBoilingPt.fMinAmt) * (fFuelHullUsage * MissileHullSize) * fModifiedRandomSeed))
    '		Dim lPayloadCost As Int32 = CInt(Math.Ceiling(uPayloadProp.fMinAmt * fModifiedRandomSeed))

    '		Dim blLauncherProdCredits As Int64 = CLng(blProdCredits * 3 * ((0.7F - (fModifiedRandomSeed - 0.8F)) + 0.8F))
    '		Dim blLauncherProdPoints As Int64 = CLng((((blLauncherProdCredits * Math.Ceiling(fItoAC) / 10.0F)) * (GetLowLookup(CInt(Math.Floor(fDPS / 25.0F)), 0) + 1)) / 100.0F)

    '		lBodyCost = Math.Max(lBodyCost \ 10, 1)
    '		lNoseCost = Math.Max(lNoseCost \ 10, 1)
    '		lFlapsCost = Math.Max(lFlapsCost \ 10, 1)
    '		lFuelCost = Math.Max(lFuelCost \ 10, 1)
    '		lPayloadCost = Math.Max(lPayloadCost \ 10, 1)

    '		If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
    '		With moResearchCost
    '			.ObjectID = Me.ObjectID
    '			.ObjTypeID = Me.ObjTypeID
    '			.ColonistCost = 0 : .EnlistedCost = 0 : .OfficerCost = 0
    '			.CreditCost = blResearchCredits
    '			Erase .ItemCosts
    '			.ItemCostUB = -1
    '			.PointsRequired = blResearchPoints

    '			.AddProductionCostItem(lBodyMineralID, ObjectType.eMineral, lBodyCost)
    '			.AddProductionCostItem(lNoseMineralID, ObjectType.eMineral, lNoseCost)
    '			.AddProductionCostItem(lFlapsMineralID, ObjectType.eMineral, lFlapsCost)
    '			.AddProductionCostItem(lFuelMineralID, ObjectType.eMineral, lFuelCost)
    '			.AddProductionCostItem(lPayloadMineralID, ObjectType.eMineral, lPayloadCost)
    '		End With
    '		If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
    '		With moProductionCost
    '			.ObjectID = Me.ObjectID
    '			.ObjTypeID = Me.ObjTypeID
    '			.ColonistCost = 0 : .EnlistedCost = 0 : .OfficerCost = 0
    '               .CreditCost = blLauncherProdCredits + blProdCredits
    '			Erase .ItemCosts
    '			.ItemCostUB = -1
    '               .PointsRequired = blLauncherProdPoints + blProdPoints

    '			.AddProductionCostItem(lBodyMineralID, ObjectType.eMineral, lBodyCost)
    '			.AddProductionCostItem(lNoseMineralID, ObjectType.eMineral, lNoseCost)
    '			.AddProductionCostItem(lFlapsMineralID, ObjectType.eMineral, lFlapsCost)
    '			.AddProductionCostItem(lFuelMineralID, ObjectType.eMineral, lFuelCost)
    '			.AddProductionCostItem(lPayloadMineralID, ObjectType.eMineral, lPayloadCost)
    '		End With

    '           'Me.blAmmoCostCredits = blProdCredits
    '           'Me.blAmmoCostPoints = blProdPoints
    '           'Me.fAmmoMin1Cost = lBodyCost
    '           'Me.fAmmoMin2Cost = lNoseCost
    '           'Me.fAmmoMin3Cost = lFlapsCost
    '           'Me.fAmmoMin4Cost = lFuelCost
    '           'Me.fAmmoMin5Cost = lPayloadCost

    '	Catch ex As Exception
    '		Me.ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
    '		Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
    '	End Try

    'End Sub

    Public Overrides Sub ComponentDesigned()
        Try

            If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
            With moResearchCost
                .ObjectID = Me.ObjectID
                .ObjTypeID = Me.ObjTypeID
                Erase .ItemCosts
                .ItemCostUB = -1
            End With
            If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
            With moProductionCost
                .ObjectID = Me.ObjectID
                .ObjTypeID = Me.ObjTypeID
                Erase .ItemCosts
                .ItemCostUB = -1
            End With

            Dim oComputer As New MissileTechComputer
            With oComputer
                Dim fROF As Single = Math.Max(1, ROF) / 30.0F
                .decDPS = CDec(MaxDmg) / CDec(fROF)
                .decDPS = Math.Max(.decDPS, (.decDPS + CDec(MaxDmg)) / 2D)

                Dim decNormalizer As Decimal = 1D
                Dim lMaxGuns As Int32 = 1
                Dim lMaxDPS As Int32 = 1
                Dim lMaxHullSize As Int32 = 1
                Dim lHullAvail As Int32 = 1
                Dim lMaxPower As Int32 = 1
                TechBuilderComputer.GetTypeValues(yHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)

                .lHullTypeID = Me.yHullTypeID
                .blExplosionRadius = ExplosionRadius
                .blHomingAccuracy = HomingAccuracy
                .blHullSize = MissileHullSize
                .blManeuver = Maneuver
                .blMaxDamage = MaxDmg
                .blMaxSpeed = MaxSpeed
                .blRange = FlightTime
                .blROF = ROF
                .blStructHP = StructureHP
                .yPayloadType = Me.PayloadType

                .blLockedProdCost = blSpecifiedProdCost
                .blLockedProdTime = blSpecifiedProdTime
                .blLockedResCost = blSpecifiedResCost
                .blLockedResTime = blSpecifiedResTime
                .lLockedColonists = lSpecifiedColonists
                .lLockedEnlisted = lSpecifiedEnlisted
                .lLockedHull = lSpecifiedHull
                .lLockedMin1 = lSpecifiedMin1
                .lLockedMin2 = lSpecifiedMin2
                .lLockedMin3 = lSpecifiedMin3
                .lLockedMin4 = lSpecifiedMin4
                .lLockedMin5 = lSpecifiedMin5
                .lLockedMin6 = lSpecifiedMin6
                .lLockedOfficers = lSpecifiedOfficers
                .lLockedPower = lSpecifiedPower

                .lMineral1ID = lBodyMineralID
                .lMineral2ID = lNoseMineralID
                .lMineral3ID = lFlapsMineralID
                .lMineral4ID = lFuelMineralID
                .lMineral5ID = lPayloadMineralID
                .lMineral6ID = -1

                If .BuilderCostValueChange(Me.ObjTypeID, lMaxDPS * 10, PopIntel, Me.Owner.oSpecials.HullPerResidence) = False Then
                    ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
                    Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
                    Return
                End If

                If .bIsMicroTech = True Then
                    If MajorDesignFlaw = 0 OrElse (MajorDesignFlaw And eComponentDesignFlaw.eGoodDesign) <> 0 Then MajorDesignFlaw = eComponentDesignFlaw.eMicroTech
                End If
            End With

            Me.CurrentSuccessChance = 100 'CInt(lAKScoreValue - fAvgResist - GetLowLookup(CInt(Math.Floor(fHPPerHullMult)), 3))
            Me.SuccessChanceIncrement = 1
            Me.HullRequired = oComputer.lResultHull
            Me.PowerRequired = oComputer.lResultPower

            If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
            Dim bValid As Boolean = True
            Try

                moResearchCost.ObjectID = Me.ObjectID
                moResearchCost.ObjTypeID = Me.ObjTypeID
                moResearchCost.ColonistCost = oComputer.lResultColonists
                moResearchCost.EnlistedCost = oComputer.lResultEnlisted
                moResearchCost.OfficerCost = oComputer.lResultOfficers
                moResearchCost.CreditCost = oComputer.blResultResCost
                moResearchCost.PointsRequired = oComputer.blResultResTime

                With moResearchCost
                    .ItemCostUB = -1
                    Erase .ItemCosts

                    .AddProductionCostItem(oComputer.lMineral1ID, ObjectType.eMineral, oComputer.lResultMin1)
                    .AddProductionCostItem(oComputer.lMineral2ID, ObjectType.eMineral, oComputer.lResultMin2)
                    .AddProductionCostItem(oComputer.lMineral3ID, ObjectType.eMineral, oComputer.lResultMin3)
                    .AddProductionCostItem(oComputer.lMineral4ID, ObjectType.eMineral, oComputer.lResultMin4)
                    .AddProductionCostItem(oComputer.lMineral5ID, ObjectType.eMineral, oComputer.lResultMin5)
                End With
            Catch
                ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
                bValid = False
            End Try

            If moProductionCost Is Nothing Then moProductionCost = New ProductionCost

            Try
                moProductionCost.ObjectID = Me.ObjectID
                moProductionCost.ObjTypeID = Me.ObjTypeID
                moProductionCost.ColonistCost = oComputer.lResultColonists
                moProductionCost.EnlistedCost = oComputer.lResultEnlisted
                moProductionCost.OfficerCost = oComputer.lResultOfficers
                moProductionCost.CreditCost = oComputer.blResultProdCost
                moProductionCost.PointsRequired = oComputer.blResultProdTime

                With moProductionCost
                    .ItemCostUB = -1
                    Erase .ItemCosts

                    .AddProductionCostItem(oComputer.lMineral1ID, ObjectType.eMineral, oComputer.lResultMin1)
                    .AddProductionCostItem(oComputer.lMineral2ID, ObjectType.eMineral, oComputer.lResultMin2)
                    .AddProductionCostItem(oComputer.lMineral3ID, ObjectType.eMineral, oComputer.lResultMin3)
                    .AddProductionCostItem(oComputer.lMineral4ID, ObjectType.eMineral, oComputer.lResultMin4)
                    .AddProductionCostItem(oComputer.lMineral5ID, ObjectType.eMineral, oComputer.lResultMin5)
                End With
            Catch
                ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
                bValid = False
            End Try
            If bValid = False Then Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
        Catch
            ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
            Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
        End Try
    End Sub


    Protected Overrides Sub FinalizeResearch()
        'Apply Bonuses
        'Dim fMult As Single = 1.0F + (Owner.oSpecials.yMissileExpRadBonus / 100.0F)
        Dim lValue As Int32 = CInt(ExplosionRadius) + CInt(Owner.oSpecials.yMissileExpRadBonus) ' CInt(Math.Ceiling(ExplosionRadius * fMult))
        If lValue > 255 Then lValue = 255
        If lValue < 0 Then lValue = 0
        ExplosionRadius = CByte(lValue)

        'fMult = 1.0F + (Owner.oSpecials.yMissileFlightTimeBonus / 100.0F)
        lValue = CInt(FlightTime) + CInt(Owner.oSpecials.yMissileFlightTimeBonus) 'CInt(Math.Ceiling(FlightTime * fMult))
        If lValue > Int16.MaxValue Then lValue = Int16.MaxValue
        If lValue < 0 Then lValue = 0
        FlightTime = CShort(lValue)

        Dim fMult As Single = 1.0F - (Owner.oSpecials.yMissileHullSizeBonus / 100.0F)
        lValue = CInt(Math.Ceiling(MissileHullSize * fMult))
        If lValue < 1 Then lValue = 1
        If lValue > Int32.MaxValue Then lValue = Int32.MaxValue
        MissileHullSize = lValue

        'fMult = 1.0F + (Owner.oSpecials.yMissileManeuverBonus / 100.0F)
        lValue = CInt(Maneuver) + CInt(Owner.oSpecials.yMissileManeuverBonus) 'CInt(Math.Ceiling(Maneuver * fMult))
        If lValue > 255 Then lValue = 255
        If lValue < 1 Then lValue = 1
        Maneuver = CByte(lValue)

        'MSC - 05/01/09 - the special tech says 5% not just plus 5
        'MaxDmg = CInt(MaxDmg) + CInt(Owner.oSpecials.yMissileMaxDmgBonus)
        If Owner.oSpecials.yMissileMaxDmgBonus > 0 Then
            fMult = 1.0F + (Owner.oSpecials.yMissileMaxDmgBonus / 100.0F)
            MaxDmg = CInt(Math.Ceiling(MaxDmg * fMult))
        End If

        'fMult = 1.0F + (Owner.oSpecials.yMissileMaxSpeedBonus / 100.0F)
        lValue = CInt(MaxSpeed) + CInt(Owner.oSpecials.yMissileMaxSpeedBonus) ' CInt(Math.Ceiling(MaxSpeed * fMult))
        If lValue > 255 Then lValue = 255
        If lValue < 1 Then lValue = 1
        MaxSpeed = CByte(lValue)
    End Sub

    Public Overrides Function GetNonOwnerMsg() As Byte()
        Dim yResult(58) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yResult, lPos) : lPos += 2
        Me.GetGUIDAsString.CopyTo(yResult, lPos) : lPos += 6

        WeaponName.CopyTo(yResult, lPos) : lPos += 20
        yResult(lPos) = CByte(WeaponClassTypeID) : lPos += 1
        yResult(lPos) = CByte(WeaponTypeID) : lPos += 1
        System.BitConverter.GetBytes(PowerRequired).CopyTo(yResult, lPos) : lPos += 4
        System.BitConverter.GetBytes(HullRequired).CopyTo(yResult, lPos) : lPos += 4
        yResult(lPos) = yHullTypeID : lPos += 1
        System.BitConverter.GetBytes(ROF).CopyTo(yResult, lPos) : lPos += 2
        System.BitConverter.GetBytes(MissileHullSize).CopyTo(yResult, lPos) : lPos += 4
        System.BitConverter.GetBytes(MaxDmg).CopyTo(yResult, lPos) : lPos += 4
        yResult(lPos) = HomingAccuracy : lPos += 1
        yResult(lPos) = MaxSpeed : lPos += 1
        yResult(lPos) = Maneuver : lPos += 1
        yResult(lPos) = ExplosionRadius : lPos += 1
        System.BitConverter.GetBytes(FlightTime).CopyTo(yResult, lPos) : lPos += 2
        System.BitConverter.GetBytes(StructureHP).CopyTo(yResult, lPos) : lPos += 4

        Return yResult
    End Function

    Public Overrides Function GetWeaponDefResult() As WeaponDef
        Dim oWpnDef As WeaponDef = New WeaponDef
        Dim fTemp As Single = 0.0F

        Dim fMaxDmgMult As Single = 1.0F
        Dim fMinDmgMult As Single = 1.0F

        Dim oTech As Epica_Tech = Owner.GetTech(137, ObjectType.eSpecialTech)
        If oTech Is Nothing = False AndAlso oTech.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eResearched Then
            fMaxDmgMult += 0.05F
        End If
        oTech = Nothing

        With oWpnDef
            .Accuracy = HomingAccuracy
            .AmmoReloadDelay = 0
            .AmmoSize = 0 'MissileHullSize
            .AOERange = ExplosionRadius

            Dim lTempImpact As Int32 = MaxDmg \ 10 'Math.Min((MissileHullSize * MaxSpeed), CInt(MaxDmg * 0.45F))
            Dim lTempPayload As Int32 = CInt(Math.Max(MaxDmg - lTempImpact, MaxDmg \ 10))

            .BeamMaxDmg = 0
            .BeamMinDmg = Me.StructureHP
            .ECMMaxDmg = 0
            .ECMMinDmg = 0
            .ImpactMaxDmg = lTempImpact 'CInt(MissileHullSize * MaxSpeed)
            .ImpactMinDmg = lTempImpact \ 10 'CInt(MissileHullSize * MaxSpeed)
            .PiercingMaxDmg = 0
            .PiercingMinDmg = 0

            If PayloadType = 0 Then
                .ChemicalMaxDmg = 0
                .ChemicalMinDmg = 0
                .FlameMaxDmg = lTempPayload
                .FlameMinDmg = .FlameMaxDmg \ 10
            Else
                .ChemicalMaxDmg = lTempPayload
                .ChemicalMinDmg = .ChemicalMaxDmg \ 10
                .FlameMaxDmg = 0
                .FlameMinDmg = 0
            End If


            .ObjectID = -1
            .ObjTypeID = ObjectType.eWeaponDef

            .Range = Me.FlightTime
            .MaxSpeed = Me.MaxSpeed
            .Maneuver = Me.Maneuver

            .RelatedWeapon = Me
            .ROF = Me.ROF
            .WeaponName = Me.WeaponName
            .WeaponType = Me.WeaponTypeID

            'Now, calculate the FirePowerRating...
            Dim lMinDmg As Int64 = .BeamMinDmg + .ChemicalMinDmg + .ECMMinDmg + .FlameMinDmg + .ImpactMinDmg + .PiercingMinDmg
            Dim lMaxDmg As Int64 = .BeamMaxDmg + .ChemicalMaxDmg + .ECMMaxDmg + .FlameMaxDmg + .ImpactMaxDmg + .PiercingMaxDmg

            .lFirePowerRating = CInt(CSng((lMaxDmg * 4) + (lMinDmg * 8)) / (.ROF / 30.0F))
        End With

        Return oWpnDef
    End Function

    Public Overrides Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblWeapon(WeaponName, WeaponClassType, WeaponTypeID, PowerRequired, HullRequired, " & _
                   "CurrentSuccessChance, SuccessChanceIncrement, ResearchAttempts, RandomSeed, ResPhase, OwnerID, ROF, " & _
                   "MaxDmg, ShotHullSize, MaxSpeed, Maneuver, FlightTime, Accuracy, PayloadType, ExplosionRadius, StructureHP, " & _
                   "Mineral1ID, Mineral2ID, Mineral3ID, Mineral4ID, Mineral5ID, ErrorReasonCode, MajorDesignFlaw, PopIntel, bArchived, HullTypeID) VALUES ('" & _
                   MakeDBStr(BytesToString(WeaponName)) & "', " & CByte(WeaponClassTypeID) & ", " & CByte(WeaponTypeID) & ", " & _
                   PowerRequired & ", " & HullRequired & ", " & CurrentSuccessChance & ", " & SuccessChanceIncrement & ", " & _
                   ResearchAttempts & ", " & RandomSeed & ", " & CInt(ComponentDevelopmentPhase) & ", "
                If Me.Owner Is Nothing = False Then
                    sSQL &= Owner.ObjectID & ", "
                Else : sSQL &= "-1, "
                End If
                sSQL &= ROF & ", " & MaxDmg & ", " & MissileHullSize & ", " & MaxSpeed & ", " & Maneuver & ", " & _
                  FlightTime & ", " & HomingAccuracy & ", " & PayloadType & ", " & ExplosionRadius & ", " & StructureHP & ", " & _
                  lBodyMineralID & ", " & lNoseMineralID & ", " & lFlapsMineralID & ", " & lFuelMineralID & ", " & _
                  lPayloadMineralID & ", " & ErrorReasonCode & ", " & MajorDesignFlaw & ", " & PopIntel & ", " & yArchived & ", " & yHullTypeID & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblWeapon SET WeaponName = '" & MakeDBStr(BytesToString(WeaponName)) & "', WeaponClassType = " & _
                  CByte(WeaponClassTypeID) & ", WeaponTypeID = " & CByte(WeaponTypeID) & ", PowerRequired = " & _
                  PowerRequired & ", HullRequired = " & HullRequired & ", CurrentSuccessChance = " & _
                  CurrentSuccessChance & ", SuccessChanceIncrement = " & SuccessChanceIncrement & ", ResearchAttempts = " & _
                  ResearchAttempts & ", RandomSeed = " & RandomSeed & ", ResPhase = " & CInt(ComponentDevelopmentPhase) & _
                  ", OwnerID = "
                If Me.Owner Is Nothing = False Then
                    sSQL &= Owner.ObjectID & ", ROF = "
                Else : sSQL &= "-1, ROF = "
                End If
                sSQL &= ROF & ", MaxDmg = " & MaxDmg & ", ShotHullSize = " & MissileHullSize & ", MaxSpeed = " & MaxSpeed & _
                  ", Maneuver = " & Maneuver & ", FlightTime = " & FlightTime & ", Accuracy = " & HomingAccuracy & _
                  ", PayloadType = " & PayloadType & ", ExplosionRadius = " & ExplosionRadius & ", StructureHP = " & StructureHP & _
                  ", Mineral1ID = " & lBodyMineralID & ", Mineral2ID = " & lNoseMineralID & ", Mineral3ID = " & lFlapsMineralID & _
                  ", Mineral4ID = " & lFuelMineralID & ", Mineral5ID = " & lPayloadMineralID & ", ErrorReasonCode = " & _
                  Me.ErrorReasonCode & ", MajorDesignFlaw = " & MajorDesignFlaw & ", PopIntel = " & PopIntel & _
                  ", bArchived = " & yArchived & ", HullTypeID = " & yHullTypeID & " WHERE WeaponID = " & Me.ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(WeaponID) FROM tblWeapon WHERE WeaponName = '" & MakeDBStr(BytesToString(WeaponName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            MyBase.FinalizeSave()

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        'bNeedsSaved = False
        Return bResult
    End Function

    Public Overrides Function SetFromDesignMsg(ByRef yData() As Byte) As Boolean
        Dim lPos As Int32 = 14  '2 for msg code, 2 for objtypeid, 4 for id, 6 for researcher guid
        Dim bRes As Boolean = False
        Try
            With Me
                .ObjTypeID = ObjectType.eWeaponTech
                .WeaponClassTypeID = CType(yData(lPos), WeaponClassType) : lPos += 1
                .WeaponTypeID = CType(yData(lPos), WeaponType) : lPos += 1

                .lSpecifiedHull = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lSpecifiedPower = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .blSpecifiedResCost = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                .blSpecifiedResTime = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                .blSpecifiedProdCost = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                .blSpecifiedProdTime = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
                .lSpecifiedColonists = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lSpecifiedEnlisted = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lSpecifiedOfficers = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lSpecifiedMin1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lSpecifiedMin2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lSpecifiedMin3 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lSpecifiedMin4 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lSpecifiedMin5 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lSpecifiedMin6 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .yHullTypeID = yData(lPos) : lPos += 1

                .ROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .MaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .MissileHullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .MaxSpeed = yData(lPos) : lPos += 1
                .Maneuver = yData(lPos) : lPos += 1
                .FlightTime = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .HomingAccuracy = yData(lPos) : lPos += 1
                .PayloadType = yData(lPos) : lPos += 1
                .ExplosionRadius = yData(lPos) : lPos += 1
                .StructureHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                .lBodyMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lNoseMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lFlapsMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lFuelMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lPayloadMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                ReDim .WeaponName(19)
                Array.Copy(yData, lPos, .WeaponName, 0, 20)
                lPos += 20
            End With

            bRes = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "MissileWeaponTech.SetFromDesignMsg: " & ex.Message)
        End Try

        Return bRes
    End Function

    Public Overrides Function ValidateDesign() As Boolean
        Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidMaterialsSelected
        If lBodyMineralID < 1 OrElse Owner.IsMineralDiscovered(lBodyMineralID) = False Then
            LogEvent(LogEventType.PossibleCheat, "MissileWeaponTech.ValidateDesign invalid minerals: " & Me.Owner.ObjectID)
            Return False
        End If
        If lNoseMineralID < 1 OrElse Owner.IsMineralDiscovered(lNoseMineralID) = False Then
            LogEvent(LogEventType.PossibleCheat, "MissileWeaponTech.ValidateDesign invalid minerals: " & Me.Owner.ObjectID)
            Return False
        End If
        If lFlapsMineralID < 1 OrElse Owner.IsMineralDiscovered(lFlapsMineralID) = False Then
            LogEvent(LogEventType.PossibleCheat, "MissileWeaponTech.ValidateDesign invalid minerals: " & Me.Owner.ObjectID)
            Return False
        End If
        If lFuelMineralID < 1 OrElse Owner.IsMineralDiscovered(lFuelMineralID) = False Then
            LogEvent(LogEventType.PossibleCheat, "MissileWeaponTech.ValidateDesign invalid minerals: " & Me.Owner.ObjectID)
            Return False
        End If
        If lPayloadMineralID < 1 OrElse Owner.IsMineralDiscovered(lPayloadMineralID) = False Then
            LogEvent(LogEventType.PossibleCheat, "MissileWeaponTech.ValidateDesign invalid minerals: " & Me.Owner.ObjectID)
            Return False
        End If

        Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
        If MaxDmg < 1 Then
            LogEvent(LogEventType.PossibleCheat, "MissileWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If
        If MissileHullSize < 1 Then
            LogEvent(LogEventType.PossibleCheat, "MissileWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If
        If FlightTime < 1 Then
            LogEvent(LogEventType.PossibleCheat, "MissileWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If
        If StructureHP < 0 OrElse StructureHP > 30 Then
            LogEvent(LogEventType.PossibleCheat, "MissileWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If
        If PayloadType > Owner.oSpecials.yPayloadTypeAvailable Then
            LogEvent(LogEventType.PossibleCheat, "MissileWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If
        If MissileHullSize > 10 Then Return False


        If CheckDesignImpossibility() = True Then
            LogEvent(LogEventType.PossibleCheat, "MissileWeaponTech.ValidateDesign CheckDesignPossiblity Failed: " & Me.Owner.ObjectID)
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
            Return False
        End If

        Me.ErrorReasonCode = TechBuilderErrorReason.eNoError
        Return True
    End Function

    'Private moAmmoCost As ProductionCost = Nothing
    'Public Function GetAmmoCost() As ProductionCost
    '	If moAmmoCost Is Nothing Then
    '		moAmmoCost = New ProductionCost
    '		With moAmmoCost
    '			.ColonistCost = 0 : .EnlistedCost = 0 : .OfficerCost = 0
    '			.CreditCost = blAmmoCostCredits
    '			.ObjectID = Me.ObjectID
    '			.ObjTypeID = ObjectType.eAmmunition
    '			'.PC_ID = -1
    '			.PointsRequired = blAmmoCostPoints
    '			.ProductionCostType = 0

    '			Dim lTemp As Int32 = CInt(fAmmoMin1Cost)
    '			If lTemp <> 0 Then .AddProductionCostItem(lBodyMineralID, ObjectType.eMineral, lTemp)
    '			lTemp = CInt(fAmmoMin2Cost)
    '			If lTemp <> 0 Then .AddProductionCostItem(lNoseMineralID, ObjectType.eMineral, lTemp)
    '			lTemp = CInt(fAmmoMin3Cost)
    '			If lTemp <> 0 Then .AddProductionCostItem(lFlapsMineralID, ObjectType.eMineral, lTemp)
    '			lTemp = CInt(fAmmoMin4Cost)
    '			If lTemp <> 0 Then .AddProductionCostItem(lFuelMineralID, ObjectType.eMineral, lTemp)
    '			lTemp = CInt(fAmmoMin5Cost)
    '			If lTemp <> 0 Then .AddProductionCostItem(lPayloadMineralID, ObjectType.eMineral, lTemp)
    '		End With
    '	End If
    '	Return moAmmoCost
    'End Function

    Public Overrides Function TechnologyScore() As Integer
        If mlStoredTechScore = Int32.MinValue Then
            Try
                Dim fROF As Single = Me.ROF / 30.0F
                Dim fTemp As Single = CInt(Me.MaxDmg) * CInt(Me.ExplosionRadius + 1)
                fTemp *= fTemp
                fTemp /= (fROF / 10.0F)

                Dim fFlightTime As Single = Me.FlightTime / 30.0F
                fTemp *= (CInt(MaxSpeed) + CInt(Maneuver) + fFlightTime + CInt(HomingAccuracy) + StructureHP)
                fTemp /= Me.MissileHullSize
                fTemp /= (Me.HullRequired + Me.PowerRequired)
                fTemp /= 1000.0F

                mlStoredTechScore = CInt(fTemp)
            Catch ex As Exception
                mlStoredTechScore = 1000I
            End Try
        End If
        Return mlStoredTechScore
    End Function

    Public Overrides Function GetPlayerTechKnowledgeMsg(ByVal yTechLvl As PlayerTechKnowledge.KnowledgeType) As Byte()
        Dim yMsg() As Byte = Nothing
        Dim lPos As Int32 = 0

        'set up our msg size
        lPos = 31
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eNameOnly Then lPos += 11
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then lPos += 33
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then lPos += 96 '20
        ReDim yMsg(lPos - 1)
        lPos = 0

        'Default Attributes
        GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        Me.WeaponName.CopyTo(yMsg, lPos) : lPos += 20
        yMsg(lPos) = CByte(WeaponClassTypeID) : lPos += 1

        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
            'at least settings 1 
            yMsg(lPos) = CByte(WeaponTypeID) : lPos += 1
            System.BitConverter.GetBytes(PowerRequired).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(HullRequired).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = Me.yHullTypeID : lPos += 1
        End If
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
            'at least settings 2  
            System.BitConverter.GetBytes(ROF).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(MaxDmg).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(MissileHullSize).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = MaxSpeed : lPos += 1
            yMsg(lPos) = Maneuver : lPos += 1
            System.BitConverter.GetBytes(FlightTime).CopyTo(yMsg, lPos) : lPos += 2
            yMsg(lPos) = HomingAccuracy : lPos += 1
            yMsg(lPos) = PayloadType : lPos += 1
            yMsg(lPos) = ExplosionRadius : lPos += 1
            System.BitConverter.GetBytes(StructureHP).CopyTo(yMsg, lPos) : lPos += 4
            If Me.GetProductionCost Is Nothing = False Then
                With Me.GetProductionCost
                    System.BitConverter.GetBytes(.ColonistCost).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.EnlistedCost).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(.OfficerCost).CopyTo(yMsg, lPos) : lPos += 4
                End With
            End If
        End If
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then
            'Full Knowledge 
            System.BitConverter.GetBytes(lBodyMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lNoseMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lFlapsMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lFuelMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lPayloadMineralID).CopyTo(yMsg, lPos) : lPos += 4

            System.BitConverter.GetBytes(lSpecifiedHull).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedPower).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(blSpecifiedResCost).CopyTo(yMsg, lPos) : lPos += 8
            System.BitConverter.GetBytes(blSpecifiedResTime).CopyTo(yMsg, lPos) : lPos += 8
            System.BitConverter.GetBytes(blSpecifiedProdCost).CopyTo(yMsg, lPos) : lPos += 8
            System.BitConverter.GetBytes(blSpecifiedProdTime).CopyTo(yMsg, lPos) : lPos += 8
            System.BitConverter.GetBytes(lSpecifiedColonists).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedEnlisted).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedOfficers).CopyTo(yMsg, lPos) : lPos += 4

            System.BitConverter.GetBytes(lSpecifiedMin1).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedMin2).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedMin3).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedMin4).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedMin5).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedMin6).CopyTo(yMsg, lPos) : lPos += 4
        End If

        Return yMsg
    End Function

    Protected Overrides Function GetSaveWeaponText() As String
        Dim sSQL As String

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!
        If ObjectID < 1 Then
            SaveObject()
            Return ""
        End If

        Try

            'UPDATE
            sSQL = "UPDATE tblWeapon SET WeaponName = '" & MakeDBStr(BytesToString(WeaponName)) & "', WeaponClassType = " & _
              CByte(WeaponClassTypeID) & ", WeaponTypeID = " & CByte(WeaponTypeID) & ", PowerRequired = " & _
              PowerRequired & ", HullRequired = " & HullRequired & ", CurrentSuccessChance = " & _
              CurrentSuccessChance & ", SuccessChanceIncrement = " & SuccessChanceIncrement & ", ResearchAttempts = " & _
              ResearchAttempts & ", RandomSeed = " & RandomSeed & ", ResPhase = " & CInt(ComponentDevelopmentPhase) & _
              ", OwnerID = "
            If Me.Owner Is Nothing = False Then
                sSQL &= Owner.ObjectID & ", ROF = "
            Else : sSQL &= "-1, ROF = "
            End If
            sSQL &= ROF & ", MaxDmg = " & MaxDmg & ", ShotHullSize = " & MissileHullSize & ", MaxSpeed = " & MaxSpeed & _
              ", Maneuver = " & Maneuver & ", FlightTime = " & FlightTime & ", Accuracy = " & HomingAccuracy & _
              ", PayloadType = " & PayloadType & ", ExplosionRadius = " & ExplosionRadius & ", StructureHP = " & StructureHP & _
              ", Mineral1ID = " & lBodyMineralID & ", Mineral2ID = " & lNoseMineralID & ", Mineral3ID = " & lFlapsMineralID & _
              ", Mineral4ID = " & lFuelMineralID & ", Mineral5ID = " & lPayloadMineralID & ", ErrorReasonCode = " & _
              Me.ErrorReasonCode & ", MajorDesignFlaw = " & MajorDesignFlaw & ", PopIntel = " & PopIntel & _
              ", bArchived = " & yArchived & ", HullTypeID = " & yHullTypeID & " WHERE WeaponID = " & Me.ObjectID
            Return sSQL
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        End Try
        Return ""
    End Function

    Public Overrides Sub FillFromPrimaryAddMsg(ByVal yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        With Me
            .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ObjTypeID = ObjectType.eAlloyTech : lPos += 2
            .Owner = GetEpicaPlayer(System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
            .ComponentDevelopmentPhase = CType(System.BitConverter.ToInt32(yData, lPos), Epica_Tech.eComponentDevelopmentPhase) : lPos += 4
            .ErrorReasonCode = yData(lPos) : lPos += 1
            lPos += 1   'researchercnt
            .MajorDesignFlaw = yData(lPos) : lPos += 1

            .lSpecifiedHull = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedPower = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .blSpecifiedResCost = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            .blSpecifiedResTime = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            .blSpecifiedProdCost = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            .blSpecifiedProdTime = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
            .lSpecifiedColonists = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedEnlisted = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedOfficers = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin2 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin3 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin4 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin5 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lSpecifiedMin6 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            ReDim .WeaponName(19)
            Array.Copy(yData, lPos, .WeaponName, 0, 20) : lPos += 20
            .WeaponClassTypeID = CType(yData(lPos), WeaponClassType) : lPos += 1
            .WeaponTypeID = CType(yData(lPos), WeaponType) : lPos += 1
            .PowerRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .HullRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .yHullTypeID = yData(lPos) : lPos += 1

            '======= end of header =========

            .ROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .MaxDmg = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .MissileHullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .MaxSpeed = yData(lPos) : lPos += 1
            .Maneuver = yData(lPos) : lPos += 1
            .FlightTime = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .HomingAccuracy = yData(lPos) : lPos += 1
            .PayloadType = yData(lPos) : lPos += 1
            .ExplosionRadius = yData(lPos) : lPos += 1
            .StructureHP = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lBodyMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lNoseMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lFlapsMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lFuelMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lPayloadMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            If Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eResearched Then
                Dim oDef As New WeaponDef
                lPos = oDef.FillFromPrimaryMsg(yData, lPos)
            End If

            If .Owner Is Nothing = False Then
                Dim lFirstIdx As Int32 = -1
                For X As Int32 = 0 To .Owner.mlWeaponUB
                    If .Owner.mlWeaponIdx(X) = .ObjectID Then
                        .Owner.moWeapon(X) = Me
                        Return
                    ElseIf lFirstIdx = -1 AndAlso .Owner.mlWeaponIdx(X) = -1 Then
                        lFirstIdx = X
                    End If
                Next X

                If lFirstIdx = -1 Then
                    lFirstIdx = .Owner.mlWeaponUB + 1
                    ReDim Preserve .Owner.mlWeaponIdx(lFirstIdx)
                    ReDim Preserve .Owner.moWeapon(lFirstIdx)
                    .Owner.mlWeaponUB = lFirstIdx
                End If
                .Owner.moWeapon(lFirstIdx) = Me
                .Owner.mlWeaponIdx(lFirstIdx) = Me.ObjectID
            End If
        End With

        If Me.moResearchCost Is Nothing Then
            Me.moResearchCost = New ProductionCost
            lPos = Me.moResearchCost.FillFromPrimaryAddMsg(yData, lPos)
        Else
            Dim oTmp As New ProductionCost
            lPos = oTmp.FillFromPrimaryAddMsg(yData, lPos)
        End If

        If Me.moProductionCost Is Nothing Then
            Me.moProductionCost = New ProductionCost
            lPos = Me.moProductionCost.FillFromPrimaryAddMsg(yData, lPos)
        Else
            Dim oTmp As New ProductionCost
            lPos = oTmp.FillFromPrimaryAddMsg(yData, lPos)
        End If

    End Sub

    Public Sub CheckForMicroTech()
        Dim oComputer As New MissileTechComputer
        With oComputer
            Dim fROF As Single = Math.Max(1, ROF) / 30.0F
            .decDPS = CDec(MaxDmg) / CDec(fROF)
            .decDPS = Math.Max(.decDPS, (.decDPS + CDec(MaxDmg)) / 2D)

            Dim decNormalizer As Decimal = 1D
            Dim lMaxGuns As Int32 = 1
            Dim lMaxDPS As Int32 = 1
            Dim lMaxHullSize As Int32 = 1
            Dim lHullAvail As Int32 = 1
            Dim lMaxPower As Int32 = 1
            TechBuilderComputer.GetTypeValues(yHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)

            .lHullTypeID = Me.yHullTypeID
            .blExplosionRadius = ExplosionRadius
            .blHomingAccuracy = HomingAccuracy
            .blHullSize = MissileHullSize
            .blManeuver = Maneuver
            .blMaxDamage = MaxDmg
            .blMaxSpeed = MaxSpeed
            .blRange = FlightTime
            .blROF = ROF
            .blStructHP = StructureHP
            .yPayloadType = Me.PayloadType

            .blLockedProdCost = blSpecifiedProdCost
            .blLockedProdTime = blSpecifiedProdTime
            .blLockedResCost = blSpecifiedResCost
            .blLockedResTime = blSpecifiedResTime
            .lLockedColonists = lSpecifiedColonists
            .lLockedEnlisted = lSpecifiedEnlisted
            .lLockedHull = lSpecifiedHull
            .lLockedMin1 = lSpecifiedMin1
            .lLockedMin2 = lSpecifiedMin2
            .lLockedMin3 = lSpecifiedMin3
            .lLockedMin4 = lSpecifiedMin4
            .lLockedMin5 = lSpecifiedMin5
            .lLockedMin6 = lSpecifiedMin6
            .lLockedOfficers = lSpecifiedOfficers
            .lLockedPower = lSpecifiedPower

            .lMineral1ID = lBodyMineralID
            .lMineral2ID = lNoseMineralID
            .lMineral3ID = lFlapsMineralID
            .lMineral4ID = lFuelMineralID
            .lMineral5ID = lPayloadMineralID
            .lMineral6ID = -1

            If .BuilderCostValueChange(Me.ObjTypeID, lMaxDPS * 10, PopIntel, Me.Owner.oSpecials.HullPerResidence) = False Then
                ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
                Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
                Return
            End If

            If .bIsMicroTech = True Then
                Dim yOrig As Byte = MajorDesignFlaw
                If MajorDesignFlaw = 0 OrElse (MajorDesignFlaw And eComponentDesignFlaw.eGoodDesign) <> 0 Then MajorDesignFlaw = eComponentDesignFlaw.eMicroTech
                If yOrig <> MajorDesignFlaw Then Me.SaveObject()
            End If
        End With
    End Sub
End Class
