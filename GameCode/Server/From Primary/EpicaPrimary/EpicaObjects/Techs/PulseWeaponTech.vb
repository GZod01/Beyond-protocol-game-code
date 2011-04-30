Option Strict On

Public Class PulseWeaponTech
    Inherits BaseWeaponTech

    Public InputEnergy As Int32
    Public CompressFactor As Single
    Public MaxRange As Byte
    Public ROF As Int16
    Public ScatterRadius As Byte

    Public CoilMatID As Int32
    Public AcceleratorMatID As Int32
    Public CasingMatID As Int32
    Public FocuserMatID As Int32
    Public CompressChmbrMatID As Int32

    Public mfPulseDegradation As Single

#Region "  Helpers  "
    Private moCoilMineral As Mineral = Nothing
    Public ReadOnly Property CoilMineral() As Mineral
        Get
            If moCoilMineral Is Nothing OrElse moCoilMineral.ObjectID <> CoilMatID Then moCoilMineral = GetEpicaMineral(CoilMatID)
            Return moCoilMineral
        End Get
    End Property
    Private moAcceleratorMineral As Mineral = Nothing
    Public ReadOnly Property AcceleratorMineral() As Mineral
        Get
            If moAcceleratorMineral Is Nothing OrElse moAcceleratorMineral.ObjectID <> AcceleratorMatID Then moAcceleratorMineral = GetEpicaMineral(AcceleratorMatID)
            Return moAcceleratorMineral
        End Get
    End Property
    Private moCasingMineral As Mineral = Nothing
    Public ReadOnly Property CasingMineral() As Mineral
        Get
            If moCasingMineral Is Nothing OrElse moCasingMineral.ObjectID <> CasingMatID Then moCasingMineral = GetEpicaMineral(CasingMatID)
            Return moCasingMineral
        End Get
    End Property
    Private moFocuserMineral As Mineral = Nothing
    Public ReadOnly Property FocuserMineral() As Mineral
        Get
            If moFocuserMineral Is Nothing OrElse moFocuserMineral.ObjectID <> FocuserMatID Then moFocuserMineral = GetEpicaMineral(FocuserMatID)
            Return moFocuserMineral
        End Get
    End Property
    Private moCompressChmbrMineral As Mineral = Nothing
    Public ReadOnly Property CompressChmbrMineral() As Mineral
        Get
            If moCompressChmbrMineral Is Nothing OrElse moCompressChmbrMineral.ObjectID <> CompressChmbrMatID Then moCompressChmbrMineral = GetEpicaMineral(CompressChmbrMatID)
            Return moCompressChmbrMineral
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
            ReDim mySendString(BASE_WEAPON_MSG_HEADER_LENGTH + yDefMsg.Length + 31)
        Else : ReDim mySendString(BASE_WEAPON_MSG_HEADER_LENGTH + 31)
        End If

        Dim lPos As Int32 = MyBase.FillBaseWeaponMsgHdr()

        System.BitConverter.GetBytes(InputEnergy).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(CompressFactor).CopyTo(mySendString, lPos) : lPos += 4
        mySendString(lPos) = MaxRange : lPos += 1
        System.BitConverter.GetBytes(ROF).CopyTo(mySendString, lPos) : lPos += 2
        mySendString(lPos) = ScatterRadius : lPos += 1

        System.BitConverter.GetBytes(CoilMatID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(AcceleratorMatID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(CasingMatID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(FocuserMatID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(CompressChmbrMatID).CopyTo(mySendString, lPos) : lPos += 4

        If yDefMsg Is Nothing = False Then
            yDefMsg.CopyTo(mySendString, lPos) : lPos += yDefMsg.Length
        End If

        ' mbStringReady = True
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
    '		Dim bNotStudyFlaw As Boolean = False
    '		Dim fHighestScore As Single = 0.0F
    '		Dim fROFInSecs As Single = ROF / 30.0F


    '		'=========== COIL ===========
    '		Dim uCo_Density As MaterialPropertyItem2
    '		With uCo_Density
    '			.lMineralID = CoilMineral.ObjectID
    '			.lPropertyID = eMinPropID.Density

    '			.lActualValue = CoilMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 135 '190
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
    '		Dim uCo_SuperC As MaterialPropertyItem2
    '		With uCo_SuperC
    '			.lMineralID = CoilMineral.ObjectID
    '			.lPropertyID = eMinPropID.SuperconductivePoint

    '			.lActualValue = CoilMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.6F
    '			.lGoalValue = 162 ' 187
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
    '		Dim uCo_MagProd As MaterialPropertyItem2
    '		With uCo_MagProd
    '			.lMineralID = CoilMineral.ObjectID
    '			.lPropertyID = eMinPropID.MagneticProduction

    '			.lActualValue = CoilMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.2F
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
    '		Dim uCo_MagReact As MaterialPropertyItem2
    '		With uCo_MagReact
    '			.lMineralID = CoilMineral.ObjectID
    '			.lPropertyID = eMinPropID.MagneticReaction

    '			.lActualValue = CoilMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 21
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
    '		Dim lCoComplexity As Int32 = CoilMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fCoilInc As Single = (uCo_Density.FinalScore + uCo_MagProd.FinalScore + uCo_MagReact.FinalScore + uCo_SuperC.FinalScore)
    '		fCoilInc *= CSng(lCoComplexity / PopIntel)
    '		'================ END OF COIL ===================

    '		'================ ACCELERATOR ===================
    '		Dim uAc_Quantum As MaterialPropertyItem2
    '		With uAc_Quantum
    '			.lMineralID = AcceleratorMineral.ObjectID
    '			.lPropertyID = eMinPropID.Quantum

    '			.lActualValue = AcceleratorMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.3F
    '			.lGoalValue = 109 '190
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
    '		Dim uAc_SuperC As MaterialPropertyItem2
    '		With uAc_SuperC
    '			.lMineralID = AcceleratorMineral.ObjectID
    '			.lPropertyID = eMinPropID.SuperconductivePoint

    '			.lActualValue = AcceleratorMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.5F
    '			.lGoalValue = 154 '210
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
    '		Dim uAc_MagReact As MaterialPropertyItem2
    '		With uAc_MagReact
    '			.lMineralID = AcceleratorMineral.ObjectID
    '			.lPropertyID = eMinPropID.MagneticReaction

    '			.lActualValue = AcceleratorMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 143 '183
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
    '		Dim uAc_MagProd As MaterialPropertyItem2
    '		With uAc_MagProd
    '			.lMineralID = AcceleratorMineral.ObjectID
    '			.lPropertyID = eMinPropID.MagneticProduction

    '			.lActualValue = AcceleratorMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 7
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
    '		Dim lAcComplexity As Int32 = AcceleratorMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fAcInc As Single = (uAc_MagProd.FinalScore + uAc_MagReact.FinalScore + uAc_Quantum.FinalScore + uAc_SuperC.FinalScore)
    '		fAcInc *= CSng(lAcComplexity / PopIntel)
    '		'================ END OF ACCELERATOR ============

    '		'================ CASING ========================
    '		Dim uCa_Density As MaterialPropertyItem2
    '		With uCa_Density
    '			.lMineralID = CasingMineral.ObjectID
    '			.lPropertyID = eMinPropID.Density

    '			.lActualValue = CasingMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.6F
    '			.lGoalValue = 132 ' 206
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
    '		Dim uCa_ThermCond As MaterialPropertyItem2
    '		With uCa_ThermCond
    '			.lMineralID = CasingMineral.ObjectID
    '			.lPropertyID = eMinPropID.ThermalConductance

    '			.lActualValue = CasingMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 112 '167
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
    '		Dim uCa_TempSens As MaterialPropertyItem2
    '		With uCa_TempSens
    '			.lMineralID = CasingMineral.ObjectID
    '			.lPropertyID = eMinPropID.TemperatureSensitivity

    '			.lActualValue = CasingMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.2F
    '			.lGoalValue = 6
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
    '		Dim uCa_Malleable As MaterialPropertyItem2
    '		With uCa_Malleable
    '			.lMineralID = CasingMineral.ObjectID
    '			.lPropertyID = eMinPropID.Malleable

    '			.lActualValue = CasingMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 3
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
    '		Dim lCaComplexity As Int32 = CasingMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fCaInc As Single = (uCa_Density.FinalScore + uCa_Malleable.FinalScore + uCa_TempSens.FinalScore + uCa_ThermCond.FinalScore)
    '		fCaInc *= CSng(lCaComplexity / PopIntel)
    '		'================ END OF CASING =================

    '		'================ FOCUSER =======================
    '		Dim uFo_Refract As MaterialPropertyItem2
    '		With uFo_Refract
    '			.lMineralID = FocuserMineral.ObjectID
    '			.lPropertyID = eMinPropID.Refraction

    '			.lActualValue = FocuserMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.4F
    '			.lGoalValue = 153 '220
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
    '		Dim uFo_Quantum As MaterialPropertyItem2
    '		With uFo_Quantum
    '			.lMineralID = FocuserMineral.ObjectID
    '			.lPropertyID = eMinPropID.Quantum

    '			.lActualValue = FocuserMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.4F
    '			.lGoalValue = 94 '160
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
    '		Dim uFo_TempSens As MaterialPropertyItem2
    '		With uFo_TempSens
    '			.lMineralID = FocuserMineral.ObjectID
    '			.lPropertyID = eMinPropID.TemperatureSensitivity

    '			.lActualValue = FocuserMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 5
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
    '		Dim uFo_ThermExp As MaterialPropertyItem2
    '		With uFo_ThermExp
    '			.lMineralID = FocuserMineral.ObjectID
    '			.lPropertyID = eMinPropID.ThermalExpansion

    '			.lActualValue = FocuserMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 9
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
    '		Dim lFoComplexity As Int32 = FocuserMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fFoInc As Single = (uFo_Quantum.FinalScore + uFo_Refract.FinalScore + uFo_TempSens.FinalScore + uFo_ThermExp.FinalScore)
    '		fFoInc *= CSng(lFoComplexity / PopIntel)
    '		'================ END OF FOCUSER ================

    '		'================ COMPRESS CHAMBER ==============
    '		Dim uCC_Density As MaterialPropertyItem2
    '		With uCC_Density
    '			.lMineralID = CompressChmbrMineral.ObjectID
    '			.lPropertyID = eMinPropID.Density

    '			.lActualValue = CompressChmbrMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.2F
    '			.lGoalValue = 110 '200
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
    '		Dim uCC_Compress As MaterialPropertyItem2
    '		With uCC_Compress
    '			.lMineralID = CompressChmbrMineral.ObjectID
    '			.lPropertyID = eMinPropID.Compressibility

    '			.lActualValue = CompressChmbrMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.2F
    '			.lGoalValue = 5
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
    '		Dim uCC_Malleable As MaterialPropertyItem2
    '		With uCC_Malleable
    '			.lMineralID = CompressChmbrMineral.ObjectID
    '			.lPropertyID = eMinPropID.Malleable

    '			.lActualValue = CompressChmbrMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.2F
    '			.lGoalValue = 1
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
    '		Dim uCC_ElectRes As MaterialPropertyItem2
    '		With uCC_ElectRes
    '			.lMineralID = CompressChmbrMineral.ObjectID
    '			.lPropertyID = eMinPropID.ElectricalResist

    '			.lActualValue = CompressChmbrMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.4F
    '			.lGoalValue = 121 '186
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
    '		Dim lCCComplexity As Int32 = CompressChmbrMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fCCInc As Single = (uCC_Compress.FinalScore + uCC_Density.FinalScore + uCC_ElectRes.FinalScore + uCC_Malleable.FinalScore)
    '		fCCInc *= CSng(lCCComplexity / PopIntel)
    '		'================ END OF COMPRESS CHAMBER =======

    '		If fHighestScore < 10.0F AndAlso bNotStudyFlaw = False Then
    '			MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eGoodDesign
    '		End If

    '		Dim fModifiedSeed As Single = 0.8F + (RandomSeed * 0.7F)

    '		Dim fAverageAK As Single = uCo_Density.AKScore + uCo_MagProd.AKScore + uCo_MagReact.AKScore + uCo_SuperC.AKScore
    '		fAverageAK += uAc_MagProd.AKScore + uAc_MagReact.AKScore + uAc_Quantum.AKScore + uAc_SuperC.AKScore
    '		fAverageAK += uCa_Density.AKScore + uCa_Malleable.AKScore + uCa_TempSens.AKScore + uCa_ThermCond.AKScore
    '		fAverageAK += uFo_Quantum.AKScore + uFo_Refract.AKScore + uFo_TempSens.AKScore + uFo_ThermExp.AKScore
    '		fAverageAK += uCC_Compress.AKScore + uCC_Density.AKScore + uCC_ElectRes.AKScore + uCC_Malleable.AKScore
    '		fAverageAK /= 20.0F

    '		Dim fAverageComplexity As Single = lCoComplexity + lAcComplexity + lCaComplexity + lFoComplexity + lCCComplexity
    '		fAverageComplexity /= 5.0F

    '		Dim fSumInc As Single = fCoilInc + fAcInc + fCaInc + fFoInc + fCCInc

    '		Me.CurrentSuccessChance = CInt(Math.Floor(-1 * fSumInc))

    '		Dim dblTemp As Double = (((InputEnergy / 10.0F) * CompressFactor) / fROFInSecs)
    '		If ScatterRadius > 0 Then dblTemp += CSng(Math.Pow(ScatterRadius, 1.3))
    '		Dim lCurrentDPS As Int32 = CInt(Math.Ceiling(dblTemp))

    '		Me.HullRequired = CInt(Math.Ceiling(CSng(GtoI2Lookup(lCurrentDPS)) * fModifiedSeed * fAverageComplexity / CSng(PopIntel)))

    '		'=ROUNDDOWN(randomseed*B43/Intelligence*B47/C47*C48,0)
    '		Dim lC48 As Int32 = GtoI3Lookup(lCurrentDPS)
    '		Me.PowerRequired = CInt(Math.Floor(fModifiedSeed * fAverageComplexity / CSng(PopIntel) * CSng(lCurrentDPS) / CSng(BtoGLookup(lC48)) * CSng(lC48)))

    '		Dim lEnlisted As Int32 = CInt(Math.Floor(CompressFactor / 20.0F))
    '		Dim lOfficers As Int32 = lEnlisted \ 5

    '		'Research Credits=ROUNDUP(B42*B43/Intelligence*randomseed*B45*B48+SUM(B1:B3)*B5,0)
    '		dblTemp = fAverageAK * fAverageComplexity / CSng(PopIntel) * fModifiedSeed * CSng(HullRequired) * CSng(PowerRequired) + (CSng(InputEnergy) + CompressFactor + CSng(MaxRange)) * CSng(ScatterRadius)
    '		Dim blResearchCredits As Int64 = CLng(Math.Ceiling(dblTemp))
    '		blResearchCredits = Math.Max(blResearchCredits, CLng(2000000 * fModifiedSeed))

    '		'Research Points
    '		Dim blResearchPoints As Int64 = CLng(Math.Ceiling(dblTemp * fAverageAK + (CSng(InputEnergy) + CompressFactor + CSng(MaxRange) + CSng(ScatterRadius))))
    '		blResearchPoints = Math.Max(blResearchPoints, CLng(10800000 * fModifiedSeed))

    '		Dim lCoilCost As Int32 = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fCoilInc))
    '		Dim lAccCost As Int32 = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fAcInc))
    '		Dim lCaseCost As Int32 = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fCaInc))
    '		Dim lFocusCost As Int32 = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fFoInc))
    '		Dim lCompChmbrCost As Int32 = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fCCInc))

    '		Dim blProdCredits As Int64 = CLng(Math.Ceiling((lCoilCost + lAccCost + lCaseCost + lFocusCost + lCompChmbrCost) * fModifiedSeed * fAverageComplexity / CSng(PopIntel) * CSng(lCurrentDPS) / 100.0F * CSng(MaxRange)))
    '		blProdCredits = Math.Max(blProdCredits, CLng(2000000 * fModifiedSeed))
    '		Dim blProdPoints As Int64 = CLng(Math.Ceiling(blProdCredits * CLng(MaxRange) * fModifiedSeed * fAverageComplexity / CSng(PopIntel)))
    '		blProdPoints = Math.Max(blProdPoints, CLng(900000 * fModifiedSeed))

    '		lCoilCost = Math.Max(1, lCoilCost \ 10)
    '		lAccCost = Math.Max(1, lAccCost \ 10)
    '		lCaseCost = Math.Max(1, lCaseCost \ 10)
    '		lFocusCost = Math.Max(1, lFocusCost \ 10)
    '		lCompChmbrCost = Math.Max(1, lCompChmbrCost \ 10)

    '		If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
    '		With moResearchCost
    '			.ObjectID = Me.ObjectID
    '			.ObjTypeID = Me.ObjTypeID
    '			.ColonistCost = 0 : .EnlistedCost = 0 : .OfficerCost = 0
    '			.CreditCost = blResearchCredits
    '			Erase .ItemCosts
    '			.ItemCostUB = -1
    '			.PointsRequired = blResearchPoints

    '			.AddProductionCostItem(CoilMatID, ObjectType.eMineral, lCoilCost * 5)
    '			.AddProductionCostItem(AcceleratorMatID, ObjectType.eMineral, lAccCost * 5)
    '			.AddProductionCostItem(CasingMatID, ObjectType.eMineral, lCaseCost * 5)
    '			.AddProductionCostItem(FocuserMatID, ObjectType.eMineral, lFocusCost * 5)
    '			.AddProductionCostItem(CompressChmbrMatID, ObjectType.eMineral, lCompChmbrCost * 5)
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

    '			.AddProductionCostItem(CoilMatID, ObjectType.eMineral, lCoilCost)
    '			.AddProductionCostItem(AcceleratorMatID, ObjectType.eMineral, lAccCost)
    '			.AddProductionCostItem(CasingMatID, ObjectType.eMineral, lCaseCost)
    '			.AddProductionCostItem(FocuserMatID, ObjectType.eMineral, lFocusCost)
    '			.AddProductionCostItem(CompressChmbrMatID, ObjectType.eMineral, lCompChmbrCost)
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

            Dim oComputer As New PulseTechComputer
            With oComputer
                Dim fROF As Single = Math.Max(1, ROF) / 30.0F
                Dim lMaxBeam As Int32 = GetMaxBeamDamage()
                Dim lMinBeam As Int32 = GetMinBeamDamage()
                Dim lMaxImpact As Int32 = GetMaxImpactDamage()
                Dim lMinImpact As Int32 = GetMinImpactDamage()

                Dim decNormalizer As Decimal = 1D
                Dim lMaxGuns As Int32 = 1
                Dim lMaxDPS As Int32 = 1
                Dim lMaxHullSize As Int32 = 1
                Dim lHullAvail As Int32 = 1
                Dim lMaxPower As Int32 = 1

                TechBuilderComputer.GetTypeValues(yHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)
                .decDPS = (CDec(lMaxBeam) + CDec(lMaxImpact)) / CDec(fROF)
                .decDPS = Math.Max(.decDPS, (.decDPS + CDec(lMaxBeam) + CDec(lMaxImpact)) / 2D)

                .lHullTypeID = yHullTypeID

                .blCompress = CLng(CompressFactor * 10)
                .blInputEnergy = InputEnergy
                .blMaxRange = MaxRange
                .blROF = ROF
                .blScatterRadius = ScatterRadius

                .lMineral1ID = CoilMatID
                .lMineral2ID = AcceleratorMatID
                .lMineral3ID = CasingMatID
                .lMineral4ID = FocuserMatID
                .lMineral5ID = CompressChmbrMatID
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


    Private Function GtoI3Lookup(ByVal lValue As Int32) As Int32
        If lValue < 11 Then
            Return 250
        ElseIf lValue < 12 Then
            Return 300
        ElseIf lValue < 14 Then
            Return 100
        ElseIf lValue < 21 Then
            Return 350
        ElseIf lValue < 40 Then
            Return 1000
        ElseIf lValue < 55 Then
            Return 525
        ElseIf lValue < 60 Then
            Return 1375
        ElseIf lValue < 125 Then
            Return 1500
        ElseIf lValue < 250 Then
            Return 3125
        ElseIf lValue < 500 Then
            Return 6250
        ElseIf lValue < 1900 Then
            Return 12500
        ElseIf lValue < 15000 Then
            Return 47500
        Else : Return 375000
        End If
    End Function
    Private Function GtoI2Lookup(ByVal lValue As Int32) As Int32
        If lValue < 10 Then
            Return 60
        ElseIf lValue < 11 Then
            Return 90
        ElseIf lValue < 12 Then
            Return 120
        ElseIf lValue < 14 Then
            Return 170
        ElseIf lValue < 21 Then
            Return 200
        ElseIf lValue < 40 Then
            Return 300
        ElseIf lValue < 55 Then
            Return 500
        ElseIf lValue < 60 Then
            Return 2000
        ElseIf lValue < 125 Then
            Return 8000
        ElseIf lValue < 250 Then
            Return 20000
        ElseIf lValue < 500 Then
            Return 40000
        ElseIf lValue < 1900 Then
            Return 80000
        ElseIf lValue < 15000 Then
            Return 250000
        Else : Return 1000000
        End If
    End Function
    Private Function BtoGLookup(ByVal lValue As Int32) As Int32
        If lValue < 250 Then
            Return 0
        ElseIf lValue < 300 Then
            Return 10
        ElseIf lValue < 350 Then
            Return 12
        ElseIf lValue < 525 Then
            Return 14
        ElseIf lValue < 1375 Then
            Return 40
        ElseIf lValue < 1500 Then
            Return 55
        ElseIf lValue < 3125 Then
            Return 60
        ElseIf lValue < 6250 Then
            Return 125
        ElseIf lValue < 12500 Then
            Return 250
        ElseIf lValue < 47500 Then
            Return 500
        ElseIf lValue < 375000 Then
            Return 1900
        Else : Return 15000
        End If
    End Function


    Protected Overrides Sub FinalizeResearch()
        'Apply bonuses
        'Dim fMult As Single = 1.0F + (Owner.oSpecials.yPulseOptRngBonus / 100.0F)
        'Dim lValue As Int32 = CInt(Math.Ceiling(MaxRange * fMult))
        Dim lValue As Int32 = CInt(MaxRange) + CInt(Owner.oSpecials.yPulseOptRngBonus)
        If lValue > 255 Then lValue = 255
        If lValue < 0 Then lValue = 0
        MaxRange = CByte(lValue)

        'fMult = 1.0F - (Owner.oSpecials.yPulseROFBonus / 100.0F)
        'lValue = CInt(Math.Ceiling(ROF * fMult))
        lValue = CInt(ROF) - CInt(Owner.oSpecials.yPulseROFBonus)
        If lValue > Int16.MaxValue Then lValue = Int16.MaxValue
        If lValue < 1 Then lValue = 1
        ROF = CShort(lValue)

        Dim yTemp As Byte = Owner.oSpecials.yPulsePowerReduced
        If yTemp > 0 Then
            Dim fMult As Single = 1.0F - (yTemp / 100.0F)
            lValue = CInt(Math.Ceiling(Me.PowerRequired * fMult))
            If lValue > Int32.MaxValue Then lValue = Int32.MaxValue
            If lValue < 1 Then lValue = 1
            PowerRequired = lValue
        End If
    End Sub

    Public Overrides Function GetNonOwnerMsg() As Byte()
        Dim yResult(42) As Byte
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
        yResult(lPos) = ScatterRadius : lPos += 1
        yResult(lPos) = MaxRange : lPos += 1

        Return yResult
    End Function

    Public Function GetAccuracy() As Byte
        Return 80   'TODO: What should this be?
    End Function

    Public Function GetMaxBeamDamage() As Int32
        'Dim fMult As Single = 1.0F + (Owner.oSpecials.yPulseMaxDmgBonus / 100.0F)
        'Return CInt(Math.Ceiling((InputEnergy / 10.0F) * CompressFactor * fMult))
        Return CInt(Math.Ceiling(GetMaxDamage() * 0.7F))
    End Function
    Public Function GetMinBeamDamage() As Int32
        'Dim fMult As Single = 1.0F + (Owner.oSpecials.yPulseMinDmgBonus / 100.0F)
        Return 0 'CInt(Math.Ceiling((lMaxBeamDmg \ 10) * fMult))
    End Function
    Public Function GetMaxImpactDamage() As Int32
        Return CInt(Math.Floor(GetMaxDamage() * 0.3F))
        'Return lMinBeamDmg
    End Function
    Public Function GetMinImpactDamage() As Int32
        Return 0 'lMaxImpactDmg \ 10
    End Function
    Public Function GetMaxDamage() As Int32
        'Dim lMaxBeam As Int32 = GetMaxBeamDamage()
        'Dim lMinBeam As Int32 = GetMinBeamDamage(lMaxBeam)
        'Return lMaxBeam + GetMaxImpactDamage(lMinBeam)
        Dim dblTemp As Double = (InputEnergy / 10.0F) * CompressFactor
        If dblTemp > Int32.MaxValue Then Return Int32.MaxValue Else Return CInt(Math.Ceiling(dblTemp))
    End Function
    Public Function GetMinDamage() As Int32
        'Dim lMaxBeam As Int32 = GetMaxBeamDamage()
        'Dim lMinBeam As Int32 = GetMinBeamDamage(lMaxBeam)
        'Dim lMaxImpact As Int32 = GetMaxImpactDamage(lMinBeam)
        'Return lMinBeam + GetMinImpactDamage(lMaxImpact)
        Return 0
    End Function

    Public Overrides Function GetWeaponDefResult() As WeaponDef
        Dim oWpnDef As WeaponDef = New WeaponDef
        Dim fTemp As Single = 0.0F

        With oWpnDef
            .Accuracy = GetAccuracy()
            .AmmoReloadDelay = 0
            .AmmoSize = 0
            .AOERange = ScatterRadius

            Dim lMaxDmgImprove As Int32 = Me.Owner.oSpecials.yPulseMaxDmgBonus
            Dim lMinDmgImprove As Int32 = Me.Owner.oSpecials.yPulseMinDmgBonus

            .BeamMaxDmg = GetMaxBeamDamage() + lMaxDmgImprove
            .BeamMinDmg = GetMinBeamDamage() + lMinDmgImprove

            .ImpactMaxDmg = GetMaxImpactDamage() + lMaxDmgImprove
            .ImpactMinDmg = GetMinImpactDamage() + lMinDmgImprove

            ''Our bonuses...
            'Dim yMaxDmgImprove As Byte = Me.Owner.oSpecials.yPulseMaxDmgBonus
            'Dim yMinDmgImprove As Byte = Me.Owner.oSpecials.yPulseMinDmgBonus 

            'If yMaxDmgImprove <> 0 Then
            '    .BeamMaxDmg += CInt(.BeamMaxDmg * (CSng(yMaxDmgImprove) / 100.0F))
            '    .ImpactMaxDmg += CInt(.ImpactMaxDmg * (CSng(yMaxDmgImprove) / 100.0F))
            'End If
            'If yMinDmgImprove <> 0 Then
            '    .BeamMinDmg += CInt(.BeamMinDmg * (CSng(yMinDmgImprove) / 100.0F))
            '    .ImpactMinDmg += CInt(.ImpactMinDmg * (CSng(yMinDmgImprove) / 100.0F))
            'End If

            .ObjectID = -1
            .ObjTypeID = ObjectType.eWeaponDef

            Dim iValue As Int16
            If mfPulseDegradation * 100 > Int16.MaxValue Then iValue = Int16.MaxValue Else iValue = CShort(mfPulseDegradation * 100)
            Dim yVals() As Byte = System.BitConverter.GetBytes(iValue)
            .MaxSpeed = yVals(0)
            .Maneuver = yVals(1)

            .ROF = Me.ROF
            .Range = Me.MaxRange

            .RelatedWeapon = Me

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
                  "Mineral4ID, Mineral5ID, ExplosionRadius, ErrorReasonCode, MajorDesignFlaw, ShotHullSize, Payload2Potential, OptimumRange, PopIntel, ProjectileHullSize, bArchived, HullTypeID) VALUES ('" & _
                  MakeDBStr(BytesToString(WeaponName)) & _
                  "', " & CInt(WeaponClassTypeID) & ", " & CInt(WeaponTypeID) & ", " & PowerRequired & ", " & HullRequired & ", " & CurrentSuccessChance & _
                  ", " & SuccessChanceIncrement & ", " & ResearchAttempts & ", " & RandomSeed & ", " & CInt(ComponentDevelopmentPhase) & ", "
                If Owner Is Nothing = False Then
                    sSQL &= Owner.ObjectID & ", "
                Else : sSQL &= "-1, "
                End If
                sSQL &= ROF & ", " & CoilMatID & ", " & AcceleratorMatID & ", " & CasingMatID & ", " & FocuserMatID & ", " & _
                  CompressChmbrMatID & ", " & ScatterRadius & ", " & Me.ErrorReasonCode & ", " & Me.MajorDesignFlaw & ", " & _
                  InputEnergy & ", " & CompressFactor & ", " & MaxRange & ", " & PopIntel & ", " & mfPulseDegradation & ", " & yArchived & ", " & yHullTypeID & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblWeapon SET WeaponName = '" & MakeDBStr(BytesToString(WeaponName)) & "', WeaponClassType = " & _
                  CInt(WeaponClassTypeID) & ", WeaponTypeID = " & CInt(WeaponTypeID) & ", PowerRequired = " & PowerRequired & ", HullRequired = " & _
                  HullRequired & ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
                  SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & RandomSeed & ", ResPhase = " & _
                  CInt(ComponentDevelopmentPhase) & ", OwnerID = "
                If Owner Is Nothing = False Then
                    sSQL &= Owner.ObjectID
                Else : sSQL &= "-1"
                End If
                sSQL &= ", ROF = " & ROF & ", Mineral1ID = " & CoilMatID & ", Mineral2ID = " & AcceleratorMatID & ", Mineral3ID = " & _
                  CasingMatID & ", Mineral4ID = " & FocuserMatID & ", Mineral5ID = " & CompressChmbrMatID & ", ExplosionRadius = " & _
                  ScatterRadius & ", ErrorReasonCode = " & ErrorReasonCode & ", MajorDesignFlaw = " & MajorDesignFlaw & _
                  ", ShotHullSize = " & InputEnergy & ", Payload2Potential = " & CompressFactor & ", OptimumRange = " & MaxRange & _
                  ", PopIntel = " & PopIntel & ", ProjectileHullSize = " & mfPulseDegradation & ", bArchived = " & yArchived & _
                  ", HullTypeID = " & yHullTypeID & " WHERE WeaponID = " & Me.ObjectID
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
                .InputEnergy = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .CompressFactor = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
                .MaxRange = yData(lPos) : lPos += 1
                .ScatterRadius = yData(lPos) : lPos += 1

                .CoilMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .AcceleratorMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .CasingMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .FocuserMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .CompressChmbrMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                ReDim .WeaponName(19)
                Array.Copy(yData, lPos, .WeaponName, 0, 20)
                lPos += 20

                .PowerRequired = .InputEnergy
            End With

            bRes = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "PulseWeaponTech.SetFromDesignMsg: " & ex.Message)
        End Try

        Return bRes
    End Function

    Public Overrides Function ValidateDesign() As Boolean
        Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidMaterialsSelected
        If CoilMatID < 1 OrElse Owner.IsMineralDiscovered(CoilMatID) = False Then
            LogEvent(LogEventType.PossibleCheat, "PulseWeaponTech.ValidateDesign invalid materials: " & Me.Owner.ObjectID)
            Return False
        End If
        If AcceleratorMatID < 1 OrElse Owner.IsMineralDiscovered(AcceleratorMatID) = False Then
            LogEvent(LogEventType.PossibleCheat, "PulseWeaponTech.ValidateDesign invalid materials: " & Me.Owner.ObjectID)
            Return False
        End If
        If CasingMatID < 1 OrElse Owner.IsMineralDiscovered(CasingMatID) = False Then
            LogEvent(LogEventType.PossibleCheat, "PulseWeaponTech.ValidateDesign invalid materials: " & Me.Owner.ObjectID)
            Return False
        End If
        If FocuserMatID < 1 OrElse Owner.IsMineralDiscovered(FocuserMatID) = False Then
            LogEvent(LogEventType.PossibleCheat, "PulseWeaponTech.ValidateDesign invalid materials: " & Me.Owner.ObjectID)
            Return False
        End If
        If CompressChmbrMatID < 1 OrElse Owner.IsMineralDiscovered(CompressChmbrMatID) = False Then
            LogEvent(LogEventType.PossibleCheat, "PulseWeaponTech.ValidateDesign invalid materials: " & Me.Owner.ObjectID)
            Return False
        End If

        Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
        If InputEnergy < 1 Then
            LogEvent(LogEventType.PossibleCheat, "PulseWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If
        If CompressFactor <= 0 Then
            LogEvent(LogEventType.PossibleCheat, "PulseWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If
        If MaxRange = 0 Then
            LogEvent(LogEventType.PossibleCheat, "PulseWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If
        If ROF < 1 Then
            LogEvent(LogEventType.PossibleCheat, "PulseWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If

        If MaxRange > Owner.oSpecials.iPulseOptRng Then
            LogEvent(LogEventType.PossibleCheat, "PulseWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If
        If CompressFactor > Owner.oSpecials.lPulseMaxCompressFactor Then
            LogEvent(LogEventType.PossibleCheat, "PulseWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If
        If InputEnergy > Owner.oSpecials.lPulseMaxInputEnergy Then
            LogEvent(LogEventType.PossibleCheat, "PulseWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If
        If ROF < Owner.oSpecials.iPulseWpnROF Then
            LogEvent(LogEventType.PossibleCheat, "PulseWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If

        If CheckDesignImpossibility() = True Then
            LogEvent(LogEventType.PossibleCheat, "PulseWeaponTech.ValidateDesign CheckDesignPossiblity Failed: " & Me.Owner.ObjectID)
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
                Dim fTemp As Single = ((InputEnergy * CompressFactor * CInt(MaxRange) * (ScatterRadius + 1)) / fROF)
                fTemp /= (Me.PowerRequired + Me.HullRequired)
                mlStoredTechScore = CInt(fTemp / 10.0F)
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
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then lPos += 24
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
            System.BitConverter.GetBytes(InputEnergy).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(CompressFactor).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = MaxRange : lPos += 1
            System.BitConverter.GetBytes(ROF).CopyTo(yMsg, lPos) : lPos += 2
            yMsg(lPos) = ScatterRadius : lPos += 1
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
            System.BitConverter.GetBytes(CoilMatID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(AcceleratorMatID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(CasingMatID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(FocuserMatID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(CompressChmbrMatID).CopyTo(yMsg, lPos) : lPos += 4

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
              CInt(WeaponClassTypeID) & ", WeaponTypeID = " & CInt(WeaponTypeID) & ", PowerRequired = " & PowerRequired & ", HullRequired = " & _
              HullRequired & ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
              SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & RandomSeed & ", ResPhase = " & _
              CInt(ComponentDevelopmentPhase) & ", OwnerID = "
            If Owner Is Nothing = False Then
                sSQL &= Owner.ObjectID
            Else : sSQL &= "-1"
            End If
            sSQL &= ", ROF = " & ROF & ", Mineral1ID = " & CoilMatID & ", Mineral2ID = " & AcceleratorMatID & ", Mineral3ID = " & _
              CasingMatID & ", Mineral4ID = " & FocuserMatID & ", Mineral5ID = " & CompressChmbrMatID & ", ExplosionRadius = " & _
              ScatterRadius & ", ErrorReasonCode = " & ErrorReasonCode & ", MajorDesignFlaw = " & MajorDesignFlaw & _
              ", ShotHullSize = " & InputEnergy & ", Payload2Potential = " & CompressFactor & ", OptimumRange = " & MaxRange & _
              ", PopIntel = " & PopIntel & ", ProjectileHullSize = " & mfPulseDegradation & ", bArchived = " & yArchived & _
              ", HullTypeID = " & yHullTypeID & " WHERE WeaponID = " & Me.ObjectID

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

            .InputEnergy = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CompressFactor = System.BitConverter.ToSingle(yData, lPos) : lPos += 4
            .MaxRange = yData(lPos) : lPos += 1
            .ROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .ScatterRadius = yData(lPos) : lPos += 1

            .CoilMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .AcceleratorMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CasingMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .FocuserMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CompressChmbrMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

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

        Dim oComputer As New PulseTechComputer
        With oComputer
            Dim fROF As Single = Math.Max(1, ROF) / 30.0F
            Dim lMaxBeam As Int32 = GetMaxBeamDamage()
            Dim lMinBeam As Int32 = GetMinBeamDamage()
            Dim lMaxImpact As Int32 = GetMaxImpactDamage()
            Dim lMinImpact As Int32 = GetMinImpactDamage()

            Dim decNormalizer As Decimal = 1D
            Dim lMaxGuns As Int32 = 1
            Dim lMaxDPS As Int32 = 1
            Dim lMaxHullSize As Int32 = 1
            Dim lHullAvail As Int32 = 1
            Dim lMaxPower As Int32 = 1

            TechBuilderComputer.GetTypeValues(yHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)
            .decDPS = (CDec(lMaxBeam) + CDec(lMaxImpact)) / CDec(fROF)
            .decDPS = Math.Max(.decDPS, (.decDPS + CDec(lMaxBeam) + CDec(lMaxImpact)) / 2D)

            .lHullTypeID = yHullTypeID

            .blCompress = CLng(CompressFactor * 10)
            .blInputEnergy = InputEnergy
            .blMaxRange = MaxRange
            .blROF = ROF
            .blScatterRadius = ScatterRadius

            .lMineral1ID = CoilMatID
            .lMineral2ID = AcceleratorMatID
            .lMineral3ID = CasingMatID
            .lMineral4ID = FocuserMatID
            .lMineral5ID = CompressChmbrMatID
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
