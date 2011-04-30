Option Strict On

Public Class BeamWeaponTech
    Inherits BaseWeaponTech

    Public MaxDamage As Int32
    Public MaxRange As Int16
    Public ROF As Int16
    Public Accuracy As Byte
    Public yDmgType As Byte

    Public CoilID As Int32
    Public CouplerID As Int32
    Public CasingID As Int32
    Public FocuserID As Int32
    Public MediumID As Int32

#Region "  Helpers  "
    Private moCoilMineral As Mineral = Nothing
    Public ReadOnly Property CoilMineral() As Mineral
        Get
            If moCoilMineral Is Nothing OrElse moCoilMineral.ObjectID <> CoilID Then moCoilMineral = GetEpicaMineral(CoilID)
            Return moCoilMineral
        End Get
    End Property
    Private moCouplerMineral As Mineral = Nothing
    Public ReadOnly Property CouplerMineral() As Mineral
        Get
            If moCouplerMineral Is Nothing OrElse moCouplerMineral.ObjectID <> CouplerID Then moCouplerMineral = GetEpicaMineral(CouplerID)
            Return moCouplerMineral
        End Get
    End Property
    Private moCasingMineral As Mineral = Nothing
    Public ReadOnly Property CasingMineral() As Mineral
        Get
            If moCasingMineral Is Nothing OrElse moCasingMineral.ObjectID <> CasingID Then moCasingMineral = GetEpicaMineral(CasingID)
            Return moCasingMineral
        End Get
    End Property
    Private moFocuserMineral As Mineral = Nothing
    Public ReadOnly Property FocuserMineral() As Mineral
        Get
            If moFocuserMineral Is Nothing OrElse moFocuserMineral.ObjectID <> FocuserID Then moFocuserMineral = GetEpicaMineral(FocuserID)
            Return moFocuserMineral
        End Get
    End Property
    Private moMediumMineral As Mineral = Nothing
    Public ReadOnly Property MediumMineral() As Mineral
        Get
            If moMediumMineral Is Nothing OrElse moMediumMineral.ObjectID <> MediumID Then moMediumMineral = GetEpicaMineral(MediumID)
            Return moMediumMineral
        End Get
    End Property
#End Region

    Public ReadOnly Property ActualAccuracy() As Int32
        Get
            '=ROUND((TargetEff / 255)*100,0)
            Return CInt((Accuracy / 255.0F) * 100.0F)
        End Get
    End Property

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
            ReDim mySendString(BASE_WEAPON_MSG_HEADER_LENGTH + 29 + yDefMsg.Length)
        Else : ReDim mySendString(BASE_WEAPON_MSG_HEADER_LENGTH + 29)
        End If

        Dim lPos As Int32 = MyBase.FillBaseWeaponMsgHdr()

        System.BitConverter.GetBytes(MaxDamage).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(MaxRange).CopyTo(mySendString, lPos) : lPos += 2
        System.BitConverter.GetBytes(ROF).CopyTo(mySendString, lPos) : lPos += 2
        mySendString(lPos) = Accuracy : lPos += 1
        mySendString(lPos) = yDmgType : lPos += 1

        System.BitConverter.GetBytes(CoilID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(CouplerID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(CasingID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(FocuserID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(MediumID).CopyTo(mySendString, lPos) : lPos += 4

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

    '	'Check our values... for auto-errors
    '	'There are no auto-errors for the solid beam weapon... that's just cool

    '	Try
    '		Dim fROFSeconds As Single = ROF / 30.0F
    '		Dim bNotStudyFlaw As Boolean = False
    '		Dim fHighestScore As Single = 0.0F

    '		'======== BEGIN COIL MATERIAL ========
    '		Dim uCoil_Density As MaterialPropertyItem2
    '		With uCoil_Density
    '			.lMineralID = CoilID
    '			.lPropertyID = eMinPropID.Density

    '			.lActualValue = CoilMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 156 '232
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop1
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uCoil_SuperC As MaterialPropertyItem2
    '		With uCoil_SuperC
    '			.lMineralID = CoilID
    '			.lPropertyID = eMinPropID.SuperconductivePoint

    '			.lActualValue = CoilMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.6F
    '			.lGoalValue = 171 ' 187
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop2 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop2
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uCoil_MagProd As MaterialPropertyItem2
    '		With uCoil_MagProd
    '			.lMineralID = CoilID
    '			.lPropertyID = eMinPropID.MagneticProduction

    '			.lActualValue = CoilMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 5
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop3 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop3
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uCoil_MagReact As MaterialPropertyItem2
    '		With uCoil_MagReact
    '			.lMineralID = CoilID
    '			.lPropertyID = eMinPropID.MagneticReaction

    '			.lActualValue = CoilMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.2F
    '			.lGoalValue = 167 '194
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop4 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop4
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim lCoilComplexity As Int32 = CoilMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fCoilInc As Single = CSng(lCoilComplexity / PopIntel)
    '		fCoilInc *= (uCoil_Density.FinalScore + uCoil_MagProd.FinalScore + uCoil_MagReact.FinalScore + uCoil_SuperC.FinalScore)
    '		'============ END OF COIL ===========

    '		'============ BEGIN OF COUPLER ======
    '		Dim uCoupler_Reflect As MaterialPropertyItem2
    '		With uCoupler_Reflect
    '			.lMineralID = CouplerID
    '			.lPropertyID = eMinPropID.Reflection

    '			.lActualValue = CouplerMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.2F
    '			.lGoalValue = 128 '180
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop1
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uCoupler_Quantum As MaterialPropertyItem2
    '		With uCoupler_Quantum
    '			.lMineralID = CouplerID
    '			.lPropertyID = eMinPropID.Quantum

    '			.lActualValue = CouplerMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 102 '170
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop2 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop2
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uCoupler_ThermExp As MaterialPropertyItem2
    '		With uCoupler_ThermExp
    '			.lMineralID = CouplerID
    '			.lPropertyID = eMinPropID.ThermalExpansion

    '			.lActualValue = CouplerMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.6F
    '			.lGoalValue = 5
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop3 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop3
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uCoupler_TempSens As MaterialPropertyItem2
    '		With uCoupler_TempSens
    '			.lMineralID = CouplerID
    '			.lPropertyID = eMinPropID.TemperatureSensitivity

    '			.lActualValue = CouplerMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 13
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop4 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop4
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim lCouplerComplexity As Int32 = CouplerMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fCouplerInc As Single = CSng(lCouplerComplexity / PopIntel)
    '		fCouplerInc *= (uCoupler_Quantum.FinalScore + uCoupler_Reflect.FinalScore + uCoupler_TempSens.FinalScore + uCoupler_ThermExp.FinalScore)
    '		'============ END OF COUPLER ========

    '		'============ CASING BEGIN ==========
    '		Dim uCasing_Density As MaterialPropertyItem2
    '		With uCasing_Density
    '			.lMineralID = CasingID
    '			.lPropertyID = eMinPropID.Density

    '			.lActualValue = CasingMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.2F
    '			.lGoalValue = 142 '200
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop1
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uCasing_ThermCond As MaterialPropertyItem2
    '		With uCasing_ThermCond
    '			.lMineralID = CasingID
    '			.lPropertyID = eMinPropID.ThermalConductance

    '			.lActualValue = CasingMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 122 ' 180
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop2 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop2
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uCasing_TempSens As MaterialPropertyItem2
    '		With uCasing_TempSens
    '			.lMineralID = CasingID
    '			.lPropertyID = eMinPropID.TemperatureSensitivity

    '			.lActualValue = CasingMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.6F
    '			.lGoalValue = 4
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop3 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop3
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uCasing_Malleable As MaterialPropertyItem2
    '		With uCasing_Malleable
    '			.lMineralID = CasingID
    '			.lPropertyID = eMinPropID.Malleable

    '			.lActualValue = CasingMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 2
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop4 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop4
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim lCasingComplexity As Int32 = CasingMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fCasingInc As Single = CSng(lCasingComplexity / PopIntel)
    '		fCasingInc *= (uCasing_Density.FinalScore + uCasing_Malleable.FinalScore + uCasing_TempSens.FinalScore + uCasing_ThermCond.FinalScore)
    '		'============ END OF CASING =========

    '		'============ FOCUSER MINERAL =======
    '		Dim uFocus_Refract As MaterialPropertyItem2
    '		With uFocus_Refract
    '			.lMineralID = FocuserID
    '			.lPropertyID = eMinPropID.Refraction

    '			.lActualValue = FocuserMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.4F
    '			.lGoalValue = 163 ' 230
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop1
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uFocus_Quantum As MaterialPropertyItem2
    '		With uFocus_Quantum
    '			.lMineralID = FocuserID
    '			.lPropertyID = eMinPropID.Quantum

    '			.lActualValue = FocuserMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 95 ' 160
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop2 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop2
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uFocus_TempSens As MaterialPropertyItem2
    '		With uFocus_TempSens
    '			.lMineralID = FocuserID
    '			.lPropertyID = eMinPropID.TemperatureSensitivity

    '			.lActualValue = FocuserMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.2F
    '			.lGoalValue = 10
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop3 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop3
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uFocus_ThermExp As MaterialPropertyItem2
    '		With uFocus_ThermExp
    '			.lMineralID = FocuserID
    '			.lPropertyID = eMinPropID.ThermalExpansion

    '			.lActualValue = FocuserMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.3F
    '			.lGoalValue = 10
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop4 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop4
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim lFocusComplexity As Int32 = FocuserMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fFocusInc As Single = CSng(lFocusComplexity / PopIntel)
    '		fFocusInc *= (uFocus_Quantum.FinalScore + uFocus_Refract.FinalScore + uFocus_TempSens.FinalScore + uFocus_ThermExp.FinalScore)
    '		'============ END OF FOCUSER ========

    '		'============ BEGIN MEDIUM ==========
    '		Dim uMedium_Quantum As MaterialPropertyItem2
    '		With uMedium_Quantum
    '			.lMineralID = MediumID
    '			.lPropertyID = eMinPropID.Quantum

    '			.lActualValue = MediumMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.5F
    '			.lGoalValue = 108 '200
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat5_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat5_Prop1
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uMedium_Reflect As MaterialPropertyItem2
    '		With uMedium_Reflect
    '			.lMineralID = MediumID
    '			.lPropertyID = eMinPropID.Reflection

    '			.lActualValue = MediumMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.3F
    '			.lGoalValue = 10
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat5_Prop2 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat5_Prop2
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uMedium_BoilingPt As MaterialPropertyItem2
    '		With uMedium_BoilingPt
    '			.lMineralID = MediumID
    '			.lPropertyID = eMinPropID.BoilingPoint

    '			.lActualValue = MediumMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 142 '200
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat5_Prop3 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat5_Prop3
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uMedium_Refract As MaterialPropertyItem2
    '		With uMedium_Refract
    '			.lMineralID = MediumID
    '			.lPropertyID = eMinPropID.Refraction

    '			.lActualValue = MediumMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 10
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat5_Prop4 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat5_Prop4
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim lMediumComplexity As Int32 = MediumMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fMediumInc As Single = CSng(lMediumComplexity / PopIntel)
    '		fMediumInc *= (uMedium_BoilingPt.FinalScore + uMedium_Quantum.FinalScore + uMedium_Reflect.FinalScore + uMedium_Refract.FinalScore)
    '		'============ END OF MEDIUM =========

    '		If fHighestScore < 10.0F AndAlso bNotStudyFlaw = False Then
    '			MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eGoodDesign
    '		End If

    '		Dim fAverageAK As Single = uCoil_Density.AKScore + uCoil_MagProd.AKScore + uCoil_MagReact.AKScore + uCoil_SuperC.AKScore
    '		fAverageAK += uCoupler_Quantum.AKScore + uCoupler_Reflect.AKScore + uCoupler_TempSens.AKScore + uCoupler_ThermExp.AKScore
    '		fAverageAK += uCasing_Density.AKScore + uCasing_Malleable.AKScore + uCasing_TempSens.AKScore + uCasing_ThermCond.AKScore
    '		fAverageAK += uFocus_Quantum.AKScore + uFocus_Refract.AKScore + uFocus_TempSens.AKScore + uFocus_ThermExp.AKScore
    '		fAverageAK += uMedium_BoilingPt.AKScore + uMedium_Quantum.AKScore + uMedium_Reflect.AKScore + uMedium_Refract.AKScore
    '		fAverageAK /= 20.0F

    '		Dim fAverageComplexity As Single = lCasingComplexity + lCoilComplexity + lCouplerComplexity + lMediumComplexity + lFocusComplexity
    '		fAverageComplexity /= 5.0F

    '		Dim fSumInc As Single = fCasingInc + fCoilInc + fCouplerInc + fMediumInc + fFocusInc
    '		Dim fTemp As Single = (fSumInc * fAverageComplexity / PopIntel) * -1
    '		Me.CurrentSuccessChance = CInt(Math.Floor(fTemp))

    '		Dim fCurrentDPS As Single = Me.MaxDamage / fROFSeconds

    '		'=ROUNDUP(VLOOKUP(B47,Sheet2!J2:K15,2,TRUE)*(B4+1)/1000*randomseed*(1+B1/1000),0)
    '		Dim fModifiedRandomSeed As Single = (0.7F * RandomSeed) + 0.8F
    '		fTemp = JtoKLookup(CInt(fCurrentDPS)) * (CSng(Me.Accuracy) + 1) / 1000.0F * fModifiedRandomSeed * (1 + MaxDamage / 1000.0F)
    '		Me.HullRequired = CInt(Math.Ceiling(fTemp))

    '		Dim fPowerOverage As Single = 1.0F
    '		If fROFSeconds < 4 Then
    '			fPowerOverage = CSng(Math.Pow(5 - fROFSeconds, 2))
    '		End If

    '		'=ROUNDUP(VLOOKUP(B47,Sheet2!J2:L15,3,TRUE)*B48*B2/100*randomseed+B47,0)
    '		fTemp = JtoLLookup(CInt(fCurrentDPS)) * fPowerOverage * CSng(MaxRange) / 100.0F * fModifiedRandomSeed + fCurrentDPS
    '		Me.PowerRequired = CInt(Math.Ceiling(fTemp))

    '		Dim lEnlisted As Int32 = (Math.Abs(Me.CurrentSuccessChance) + MaxDamage) \ 100
    '		Dim lOfficers As Int32 = lEnlisted \ 5

    '		'=ROUNDUP(MAX(B43/Intelligence*B45*B49,randomseed*1000000),0)
    '		fTemp = CSng(fAverageComplexity / PopIntel * Me.HullRequired * Me.PowerRequired)
    '		fTemp = Math.Max(fTemp, fModifiedRandomSeed * 1000000)
    '		fTemp *= (fSumInc / 5.0F) * 10.0F
    '		Dim blResearchCredits As Int64 = CLng(Math.Ceiling(fTemp))
    '		Dim blResearchPoints As Int64 = CLng(Math.Ceiling(blResearchCredits * fModifiedRandomSeed / 100.0F * fAverageAK))
    '		blResearchPoints = Math.Max(blResearchPoints, CLng(10800000 * fModifiedRandomSeed))

    '		Dim blProdCredits As Int64 = CLng(Math.Ceiling(Math.Max(blResearchPoints, 10000 * Me.HullRequired * fModifiedRandomSeed)))
    '		blProdCredits = Math.Max(blProdCredits, CLng(fModifiedRandomSeed * 1000000))
    '		blProdCredits = CLng(blProdCredits * (fSumInc / 5.0F))

    '		Dim lCoilCost As Int32 = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fCoilInc))
    '		Dim lCouplerCost As Int32 = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fCouplerInc))
    '		Dim lCasingCost As Int32 = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fCasingInc))
    '		Dim lFocusCost As Int32 = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fFocusInc))
    '		Dim lMediumCost As Int32 = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fMediumInc))

    '		Dim blProdPoints As Int64 = CLng((lCoilCost + lCouplerCost + lCasingCost + lFocusCost + lMediumCost) * fAverageComplexity * fModifiedRandomSeed)
    '		blProdPoints = Math.Max(blProdPoints, CLng(1800000 * fModifiedRandomSeed))

    '		If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
    '		With moResearchCost
    '			.ObjectID = Me.ObjectID
    '			.ObjTypeID = Me.ObjTypeID
    '			.ColonistCost = 0 : .EnlistedCost = 0 : .OfficerCost = 0
    '			.CreditCost = blResearchCredits
    '			Erase .ItemCosts
    '			.ItemCostUB = -1
    '			.PointsRequired = blResearchPoints

    '			.AddProductionCostItem(CoilID, ObjectType.eMineral, CInt(lCoilCost * fAverageComplexity / PopIntel * fModifiedRandomSeed / 10))
    '			.AddProductionCostItem(CouplerID, ObjectType.eMineral, CInt(lCouplerCost * fAverageComplexity / PopIntel * fModifiedRandomSeed / 10))
    '			.AddProductionCostItem(CasingID, ObjectType.eMineral, CInt(lCasingCost * fAverageComplexity / PopIntel * fModifiedRandomSeed / 10))
    '			.AddProductionCostItem(FocuserID, ObjectType.eMineral, CInt(lFocusCost * fAverageComplexity / PopIntel * fModifiedRandomSeed / 10))
    '			.AddProductionCostItem(MediumID, ObjectType.eMineral, CInt(lMediumCost * fAverageComplexity / PopIntel * fModifiedRandomSeed / 10))
    '		End With
    '		If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
    '		With moProductionCost
    '			.ObjectID = Me.ObjectID
    '			.ObjTypeID = Me.ObjTypeID
    '			.ColonistCost = 0 : .EnlistedCost = lEnlisted : .OfficerCost = lOfficers
    '			.CreditCost = blProdCredits
    '			Erase .ItemCosts
    '			.ItemCostUB = -1
    '			.PointsRequired = blProdPoints

    '			.AddProductionCostItem(CoilID, ObjectType.eMineral, Math.Max(1, lCoilCost \ 10))
    '			.AddProductionCostItem(CouplerID, ObjectType.eMineral, Math.Max(1, lCouplerCost \ 10))
    '			.AddProductionCostItem(CasingID, ObjectType.eMineral, Math.Max(1, lCasingCost \ 10))
    '			.AddProductionCostItem(FocuserID, ObjectType.eMineral, Math.Max(1, lFocusCost \ 10))
    '			.AddProductionCostItem(MediumID, ObjectType.eMineral, Math.Max(1, lMediumCost \ 10))
    '		End With

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

            Dim oComputer As New SolidTechComputer
            With oComputer

                Dim fROF As Single = Math.Max(1, Me.ROF) / 30.0F

                Dim lMinPierce As Int32 = 0
                Dim lMaxPierce As Int32 = 0
                Dim lMinBurn As Int32 = 0
                Dim lMaxBurn As Int32 = 0
                Dim lMinBeam As Int32 = 0
                Dim lMaxBeam As Int32 = 0

                If yDmgType = 0 Then
                    'Piercing
                    Dim lVal As Int32 = CInt(MaxDamage * 0.3F)
                    lMaxPierce = lVal
                    lVal = MaxDamage - lVal
                    lMaxBeam = lVal
                Else
                    'thermal
                    Dim lVal As Int32 = CInt(MaxDamage * 0.3F)
                    lMaxBurn = lVal
                    lVal = MaxDamage - lVal
                    lMaxBeam = lVal
                End If
                lMinBeam = lMaxBeam \ 10
                .decDPS = MaxDamage / CDec(fROF)
                .decDPS = Math.Max(.decDPS, (.decDPS + MaxDamage) / 2D)


                Dim decNormalizer As Decimal = 1D
                Dim lMaxGuns As Int32 = 1
                Dim lMaxDPS As Int32 = 1
                Dim lMaxHullSize As Int32 = 1
                Dim lHullAvail As Int32 = 1
                Dim lMaxPower As Int32 = 1

                TechBuilderComputer.GetTypeValues(yHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)
                .lHullTypeID = yHullTypeID

                .blAccuracy = Accuracy
                .blMaxDmg = MaxDamage
                .blMaxRng = MaxRange

                .lMineral1ID = CoilID
                .lMineral2ID = CouplerID
                .lMineral3ID = CasingID
                .lMineral4ID = FocuserID
                .lMineral5ID = MediumID
                .lMineral6ID = -1

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

                If .BuilderCostValueChange(Me.ObjTypeID, lMaxDPS, PopIntel, Me.Owner.oSpecials.HullPerResidence) = False Then
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

    Private Function JtoKLookup(ByVal lValue As Int32) As Int32
        If lValue < 10 Then
            Return 72
        ElseIf lValue < 12 Then
            Return 136
        ElseIf lValue < 14 Then
            Return 240
        ElseIf lValue < 21 Then
            Return 400
        ElseIf lValue < 40 Then
            Return 6400
        ElseIf lValue < 55 Then
            Return 1600
        ElseIf lValue < 60 Then
            Return 16000
        ElseIf lValue < 125 Then
            Return 32000
        ElseIf lValue < 250 Then
            Return 64000
        ElseIf lValue < 500 Then
            Return 200000
        ElseIf lValue < 1900 Then
            Return 800000
        ElseIf lValue < 15000 Then
            Return 3200000
        Else : Return 3200000
        End If
    End Function
    Private Function JtoLLookup(ByVal lValue As Int32) As Int32
        If lValue < 10 Then
            Return 250
        ElseIf lValue < 12 Then
            Return 303
        ElseIf lValue < 14 Then
            Return 100
        ElseIf lValue < 21 Then
            Return 347
        ElseIf lValue < 40 Then
            Return 536
        ElseIf lValue < 55 Then
            Return 1042
        ElseIf lValue < 60 Then
            Return 1389
        ElseIf lValue < 125 Then
            Return 1515
        ElseIf lValue < 250 Then
            Return 3125
        ElseIf lValue < 500 Then
            Return 6250
        ElseIf lValue < 1900 Then
            Return 12500
        ElseIf lValue < 15000 Then
            Return 46900
        Else : Return 375000
        End If
    End Function


    Protected Overrides Sub FinalizeResearch()
        'Apply bonuses
        'fMult = 1.0F + (Owner.oSpecials.ySolidBeamOptRngBonus / 100.0F)
        Dim lValue As Int32 = CInt(MaxRange) + CInt(Owner.oSpecials.ySolidBeamOptRngBonus) 'CInt(Math.Ceiling(MaxRange * fMult))
        If lValue > Int16.MaxValue Then lValue = Int16.MaxValue
        MaxRange = CShort(lValue)

        'fMult = 1.0F - (Owner.oSpecials.ySolidBeamROFBonus / 100.0F)
        lValue = CInt(ROF) - CInt(Owner.oSpecials.ySolidBeamROFBonus) ' CInt(Math.Ceiling(ROF * fMult))
        If lValue < 1 Then lValue = 1
        If lValue > Int16.MaxValue Then lValue = Int16.MaxValue
        ROF = CShort(lValue)

        Dim yTemp As Byte = Owner.oSpecials.ySolidBeamPowerBonus
        If yTemp > 0 Then
            Dim fMult As Single = 1.0F - (yTemp / 100.0F)
            lValue = CInt(Math.Ceiling(Me.PowerRequired * fMult))
            If lValue > Int32.MaxValue Then lValue = Int32.MaxValue
            If lValue < 1 Then lValue = 1
            PowerRequired = lValue
        End If

        'MSC - 05/01/09 - fixes the percentage based dmg bonus
        yTemp = Owner.oSpecials.ySolidBeamMaxDmgBonus
        If yTemp > 0 Then
            Dim fMult As Single = 1.0F + (yTemp / 100.0F)
            lValue = CInt(Math.Ceiling(MaxDamage * fMult))
            If lValue > Int32.MaxValue Then lValue = Int32.MaxValue
            If lValue < 1 Then lValue = 1
            MaxDamage = lValue
        End If
        'MaxDamage += CInt(Owner.oSpecials.ySolidBeamMaxDmgBonus) 'CInt(Math.Ceiling(MaxDamage * fMult))

    End Sub

    Public Overrides Function GetNonOwnerMsg() As Byte()
        Dim yResult(47) As Byte
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
        System.BitConverter.GetBytes(MaxDamage).CopyTo(yResult, lPos) : lPos += 4
        System.BitConverter.GetBytes(MaxRange).CopyTo(yResult, lPos) : lPos += 2
        yResult(lPos) = Accuracy : lPos += 1

        Return yResult
    End Function

    Public Overrides Function GetWeaponDefResult() As WeaponDef
        Dim oWpnDef As WeaponDef = New WeaponDef
        Dim fTemp As Single = 0.0F

        With oWpnDef
            .Accuracy = CByte(ActualAccuracy)
            .AmmoReloadDelay = 0
            .AmmoSize = 0
            .AOERange = 0

            If yDmgType = 0 Then
                'Piercing
                Dim lVal As Int32 = CInt(MaxDamage * 0.3F)
                .PiercingMaxDmg = lVal
                lVal = MaxDamage - lVal
                .BeamMaxDmg = lVal
            Else
                'thermal
                Dim lVal As Int32 = CInt(MaxDamage * 0.3F)
                .FlameMaxDmg = lVal
                lVal = MaxDamage - lVal
                .BeamMaxDmg = lVal
            End If
            .BeamMinDmg = .BeamMaxDmg \ 10

            If Owner.oSpecials.ySolidBeamMinDmgBonus <> 0 Then
                'Dim fMult As Single = 1.0F + (Owner.oSpecials.ySolidBeamMinDmgBonus / 100.0F)
                'Dim lValue As Int32 = CInt(Math.Ceiling(.BeamMinDmg * fMult))
                .BeamMinDmg = .BeamMinDmg + CInt(Owner.oSpecials.ySolidBeamMinDmgBonus)
            End If


            .ObjectID = -1
            .ObjTypeID = ObjectType.eWeaponDef

            .Range = Me.MaxRange
            .MaxSpeed = 100
            .Maneuver = 0

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
                sSQL = "INSERT INTO tblWeapon (WeaponName, WeaponClassType, WeaponTypeID, PowerRequired, HullRequired, CurrentSuccessChance, " & _
                  "SuccessChanceIncrement, ResearchAttempts, RandomSeed, ResPhase, OwnerID, ROF, Mineral1ID, Mineral2ID, Mineral3ID, " & _
                  "Mineral4ID, Mineral5ID, ErrorReasonCode, MajorDesignFlaw, MaxDmg, OptimumRange, Accuracy, PayloadType, PopIntel, bArchived, HullTypeID) VALUES ('" & _
                  MakeDBStr(BytesToString(WeaponName)) & "', " & CInt(WeaponClassTypeID) & ", " & CInt(WeaponTypeID) & ", " & _
                  PowerRequired & ", " & HullRequired & ", " & CurrentSuccessChance & ", " & SuccessChanceIncrement & ", " & _
                  ResearchAttempts & ", " & RandomSeed & ", " & CInt(ComponentDevelopmentPhase) & ", "
                If Owner Is Nothing = False Then
                    sSQL &= Owner.ObjectID & ", "
                Else : sSQL &= "-1, "
                End If
                sSQL &= ROF & ", " & CoilID & ", " & CouplerID & ", " & CasingID & ", " & FocuserID & ", " & MediumID & ", " & _
                  ErrorReasonCode & ", " & MajorDesignFlaw & ", " & MaxDamage & ", " & MaxRange & ", " & Accuracy & ", " & yDmgType & _
                  ", " & PopIntel & ", " & yArchived & ", " & yHullTypeID & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblWeapon SET WeaponName = '" & MakeDBStr(BytesToString(WeaponName)) & "', WeaponClassType = " & _
                  CInt(WeaponClassTypeID) & ", WeaponTypeID = " & CInt(WeaponTypeID) & ", PowerRequired = " & PowerRequired & _
                  ", HullRequired = " & HullRequired & ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
                  SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", Randomseed = " & RandomSeed & ", ResPhase = " & _
                  CInt(ComponentDevelopmentPhase) & ", OwnerID = "
                If Owner Is Nothing = False Then
                    sSQL &= Owner.ObjectID & ", ROF = "
                Else : sSQL &= "-1, ROF = "
                End If
                sSQL &= ROF & ", Mineral1ID = " & CoilID & ", Mineral2ID = " & CouplerID & ", Mineral3ID = " & CasingID & _
                  ", Mineral4ID = " & FocuserID & ", Mineral5ID = " & MediumID & ", ErrorReasonCode = " & ErrorReasonCode & _
                  ", MajorDesignFlaw = " & MajorDesignFlaw & ", MaxDmg = " & MaxDamage & ", OptimumRange = " & MaxRange & _
                  ", Accuracy = " & Accuracy & ", PayloadType = " & yDmgType & ", PopIntel = " & PopIntel & ", bArchived = " & _
                  yArchived & ", HullTypeID = " & yHullTypeID & " WHERE WeaponID = " & Me.ObjectID
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
                .MaxDamage = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .MaxRange = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .Accuracy = yData(lPos) : lPos += 1
                .yDmgType = yData(lPos) : lPos += 1

                .MediumID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .FocuserID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .CasingID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .CouplerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .CoilID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                ReDim .WeaponName(19)
                Array.Copy(yData, lPos, .WeaponName, 0, 20)
                lPos += 20
            End With

            bRes = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "SolidBeamWeaponTech.SetFromDesignMsg: " & ex.Message)
        End Try

        Return bRes
    End Function

    Public Overrides Function ValidateDesign() As Boolean
        Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidMaterialsSelected
        If CoilID < 1 OrElse Owner.IsMineralDiscovered(CoilID) = False Then
            LogEvent(LogEventType.PossibleCheat, "BeamWeaponTech.ValidateDesign invalid mineral selection: " & Me.Owner.ObjectID)
            Return False
        End If
        If CouplerID < 1 OrElse Owner.IsMineralDiscovered(CouplerID) = False Then
            LogEvent(LogEventType.PossibleCheat, "BeamWeaponTech.ValidateDesign invalid mineral selection: " & Me.Owner.ObjectID)
            Return False
        End If
        If CasingID < 1 OrElse Owner.IsMineralDiscovered(CasingID) = False Then
            LogEvent(LogEventType.PossibleCheat, "BeamWeaponTech.ValidateDesign invalid mineral selection: " & Me.Owner.ObjectID)
            Return False
        End If
        If FocuserID < 1 OrElse Owner.IsMineralDiscovered(FocuserID) = False Then
            LogEvent(LogEventType.PossibleCheat, "BeamWeaponTech.ValidateDesign invalid mineral selection: " & Me.Owner.ObjectID)
            Return False
        End If
        If MediumID < 1 OrElse Owner.IsMineralDiscovered(MediumID) = False Then
            LogEvent(LogEventType.PossibleCheat, "BeamWeaponTech.ValidateDesign invalid mineral selection: " & Me.Owner.ObjectID)
            Return False
        End If

        Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
        If MaxDamage < 1 Then
            LogEvent(LogEventType.PossibleCheat, "BeamWeaponTech.ValidateDesign invalid max damage: " & Me.Owner.ObjectID)
            Return False
        End If
        If MaxRange = 0 Then
            LogEvent(LogEventType.PossibleCheat, "BeamWeaponTech.ValidateDesign invalid max range: " & Me.Owner.ObjectID)
            Return False
        End If
        If ROF < 1 Then
            LogEvent(LogEventType.PossibleCheat, "BeamWeaponTech.ValidateDesign invalid rof: " & Me.Owner.ObjectID)
            Return False
        End If
        If yDmgType <> 0 AndAlso yDmgType <> 1 Then
            LogEvent(LogEventType.PossibleCheat, "BeamWeaponTech.ValidateDesign invalid dmg type: " & Me.Owner.ObjectID)
            Return False
        End If

        If MaxRange > Owner.oSpecials.iSolidBeamOptimumRange Then
            LogEvent(LogEventType.PossibleCheat, "BeamWeaponTech.ValidateDesign invalid max range: " & Me.Owner.ObjectID)
            Return False
        End If
        If Accuracy > Owner.oSpecials.ySolidBeamMaxAccuracy Then
            LogEvent(LogEventType.PossibleCheat, "BeamWeaponTech.ValidateDesign invalid accuracy: " & Me.Owner.ObjectID)
            Return False
        End If
        If MaxDamage > Owner.oSpecials.lSolidBeamMaxDmg Then
            LogEvent(LogEventType.PossibleCheat, "BeamWeaponTech.ValidateDesign invalid max damage: " & Me.Owner.ObjectID)
            Return False
        End If
        If ROF < Owner.oSpecials.iSolidBeamROF Then
            LogEvent(LogEventType.PossibleCheat, "BeamWeaponTech.ValidateDesign invalid ROF: " & Me.Owner.ObjectID)
            Return False
        End If

        If CheckDesignImpossibility() = True Then
            LogEvent(LogEventType.PossibleCheat, "BeamWeaponTech.ValidateDesign CheckDesignPossiblity Failed: " & Me.Owner.ObjectID)
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
            Return False
        End If

        Me.ErrorReasonCode = TechBuilderErrorReason.eNoError
        Return True

    End Function

    Public Overrides Function TechnologyScore() As Integer
        If mlStoredTechScore = Int32.MinValue Then
            Try
                Dim fROF As Single = Me.ROF / 30.0F
                mlStoredTechScore = CInt(((Me.MaxDamage * CInt(MaxRange) * CInt(Accuracy)) / fROF) - (Me.PowerRequired + Me.HullRequired)) \ 1000I
            Catch ex As Exception
                mlStoredTechScore = 1000
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
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then lPos += 22
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
            System.BitConverter.GetBytes(MaxDamage).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(MaxRange).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(ROF).CopyTo(yMsg, lPos) : lPos += 2
            yMsg(lPos) = Accuracy : lPos += 1
            yMsg(lPos) = yDmgType : lPos += 1
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
            System.BitConverter.GetBytes(CoilID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(CouplerID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(CasingID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(FocuserID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(MediumID).CopyTo(yMsg, lPos) : lPos += 4

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
        If ObjectID = -1 Then
            SaveObject()
            Return ""
        End If

        Try
            'UPDATE
            sSQL = "UPDATE tblWeapon SET WeaponName = '" & MakeDBStr(BytesToString(WeaponName)) & "', WeaponClassType = " & _
              CInt(WeaponClassTypeID) & ", WeaponTypeID = " & CInt(WeaponTypeID) & ", PowerRequired = " & PowerRequired & _
              ", HullRequired = " & HullRequired & ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
              SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", Randomseed = " & RandomSeed & ", ResPhase = " & _
              CInt(ComponentDevelopmentPhase) & ", OwnerID = "
            If Owner Is Nothing = False Then
                sSQL &= Owner.ObjectID & ", ROF = "
            Else : sSQL &= "-1, ROF = "
            End If
            sSQL &= ROF & ", Mineral1ID = " & CoilID & ", Mineral2ID = " & CouplerID & ", Mineral3ID = " & CasingID & _
              ", Mineral4ID = " & FocuserID & ", Mineral5ID = " & MediumID & ", ErrorReasonCode = " & ErrorReasonCode & _
              ", MajorDesignFlaw = " & MajorDesignFlaw & ", MaxDmg = " & MaxDamage & ", OptimumRange = " & MaxRange & _
              ", Accuracy = " & Accuracy & ", PayloadType = " & yDmgType & ", PopIntel = " & PopIntel & ", bArchived = " & _
              yArchived & ", HullTypeID = " & yHullTypeID & " WHERE WeaponID = " & Me.ObjectID

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

            .MaxDamage = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .MaxRange = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .ROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .Accuracy = yData(lPos) : lPos += 1
            .yDmgType = yData(lPos) : lPos += 1

            .CoilID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CouplerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CasingID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .FocuserID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .MediumID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

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

        Dim oComputer As New SolidTechComputer
        With oComputer

            Dim fROF As Single = Math.Max(1, Me.ROF) / 30.0F

            Dim lMinPierce As Int32 = 0
            Dim lMaxPierce As Int32 = 0
            Dim lMinBurn As Int32 = 0
            Dim lMaxBurn As Int32 = 0
            Dim lMinBeam As Int32 = 0
            Dim lMaxBeam As Int32 = 0

            If yDmgType = 0 Then
                'Piercing
                Dim lVal As Int32 = CInt(MaxDamage * 0.3F)
                lMaxPierce = lVal
                lVal = MaxDamage - lVal
                lMaxBeam = lVal
            Else
                'thermal
                Dim lVal As Int32 = CInt(MaxDamage * 0.3F)
                lMaxBurn = lVal
                lVal = MaxDamage - lVal
                lMaxBeam = lVal
            End If
            lMinBeam = lMaxBeam \ 10
            .decDPS = MaxDamage / CDec(fROF)
            .decDPS = Math.Max(.decDPS, (.decDPS + MaxDamage) / 2D)


            Dim decNormalizer As Decimal = 1D
            Dim lMaxGuns As Int32 = 1
            Dim lMaxDPS As Int32 = 1
            Dim lMaxHullSize As Int32 = 1
            Dim lHullAvail As Int32 = 1
            Dim lMaxPower As Int32 = 1

            TechBuilderComputer.GetTypeValues(yHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)
            .lHullTypeID = yHullTypeID

            .blAccuracy = Accuracy
            .blMaxDmg = MaxDamage
            .blMaxRng = MaxRange

            .lMineral1ID = CoilID
            .lMineral2ID = CouplerID
            .lMineral3ID = CasingID
            .lMineral4ID = FocuserID
            .lMineral5ID = MediumID
            .lMineral6ID = -1

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

            .bRunCheckForMicro = True
            If .BuilderCostValueChange(Me.ObjTypeID, lMaxDPS, PopIntel, Me.Owner.oSpecials.HullPerResidence) = False Then
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
 