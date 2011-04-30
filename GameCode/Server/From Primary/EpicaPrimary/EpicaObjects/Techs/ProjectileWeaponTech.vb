Option Strict On

Public Class ProjectileWeaponTech
    Inherits BaseWeaponTech

    Public ProjectionType As Byte
    Public CartridgeSize As Int32
    Public PierceRatio As Byte
    Public ROF As Int16
    Public MaxRange As Int16
    Public PayloadType As Byte
    Public ExplosionRadius As Byte

    Private mlMinPierce As Int32
    Private mlMaxPierce As Int32
    Private mlMinImpact As Int32
    Private mlMaxImpact As Int32
    Private mlMinPayload As Int32
    Private mlMaxPayload As Int32

    'Public mfAmmoSize As Single
    Public mfPayload1PotentialPierce As Single
    Public mfPayload1PotentialImpact As Single
    Public mfPayload2Potential As Single

    Public BarrelMineralID As Int32
    Public CasingMineralID As Int32
    Public Payload1MineralID As Int32
    Public Payload2MineralID As Int32
    Public ProjectionMineralID As Int32

    'Public blAmmoCostCredits As Int64 = 0
    'Public blAmmoCostPoints As Int64 = 0
    'Public fAmmoMin1Cost As Single = 0.0F
    'Public fAmmoMin2Cost As Single = 0.0F
    'Public fAmmoMin3Cost As Single = 0.0F
    'Public fAmmoMin4Cost As Single = 0.0F
    'Public fAmmoMin5Cost As Single = 0.0F

#Region "  Helpers  "
    Private moBarrelMineral As Mineral = Nothing
    Public ReadOnly Property BarrelMineral() As Mineral
        Get
            If moBarrelMineral Is Nothing OrElse moBarrelMineral.ObjectID <> BarrelMineralID Then moBarrelMineral = GetEpicaMineral(BarrelMineralID)
            Return moBarrelMineral
        End Get
    End Property
    Private moCasingMineral As Mineral = Nothing
    Public ReadOnly Property CasingMineral() As Mineral
        Get
            If moCasingMineral Is Nothing OrElse moCasingMineral.ObjectID <> CasingMineralID Then moCasingMineral = GetEpicaMineral(CasingMineralID)
            Return moCasingMineral
        End Get
    End Property
    Private moPayload1Mineral As Mineral = Nothing
    Public ReadOnly Property Payload1Mineral() As Mineral
        Get
            If moPayload1Mineral Is Nothing OrElse moPayload1Mineral.ObjectID <> Payload1MineralID Then moPayload1Mineral = GetEpicaMineral(Payload1MineralID)
            Return moPayload1Mineral
        End Get
    End Property
    Private moPayload2Mineral As Mineral = Nothing
    Public ReadOnly Property Payload2Mineral() As Mineral
        Get
            If moPayload2Mineral Is Nothing OrElse moPayload2Mineral.ObjectID <> Payload2MineralID Then moPayload2Mineral = GetEpicaMineral(Payload2MineralID)
            Return moPayload2Mineral
        End Get
    End Property
    Private moProjectionMineral As Mineral = Nothing
    Public ReadOnly Property ProjectionMineral() As Mineral
        Get
            If moProjectionMineral Is Nothing OrElse moProjectionMineral.ObjectID <> ProjectionMineralID Then moProjectionMineral = GetEpicaMineral(ProjectionMineralID)
            Return moProjectionMineral
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
            ReDim mySendString(BASE_WEAPON_MSG_HEADER_LENGTH + yDefMsg.Length + 55) '95)
        Else : ReDim mySendString(BASE_WEAPON_MSG_HEADER_LENGTH + 55) '95)
        End If

        Dim lPos As Int32 = MyBase.FillBaseWeaponMsgHdr()

        mySendString(lPos) = ProjectionType : lPos += 1
        System.BitConverter.GetBytes(CartridgeSize).CopyTo(mySendString, lPos) : lPos += 4
        mySendString(lPos) = PierceRatio : lPos += 1
        System.BitConverter.GetBytes(ROF).CopyTo(mySendString, lPos) : lPos += 2
        System.BitConverter.GetBytes(MaxRange).CopyTo(mySendString, lPos) : lPos += 2
        mySendString(lPos) = PayloadType : lPos += 1
        mySendString(lPos) = ExplosionRadius : lPos += 1

        System.BitConverter.GetBytes(BarrelMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(CasingMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(Payload1MineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(Payload2MineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(ProjectionMineralID).CopyTo(mySendString, lPos) : lPos += 4

        System.BitConverter.GetBytes(mlMinPierce).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(mlMaxPierce).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(mlMinImpact).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(mlMaxImpact).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(mlMinPayload).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(mlMaxPayload).CopyTo(mySendString, lPos) : lPos += 4

        'System.BitConverter.GetBytes(mfAmmoSize).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(blAmmoCostCredits).CopyTo(mySendString, lPos) : lPos += 8
        'System.BitConverter.GetBytes(blAmmoCostPoints).CopyTo(mySendString, lPos) : lPos += 8
        'System.BitConverter.GetBytes(fAmmoMin1Cost).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(fAmmoMin2Cost).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(fAmmoMin3Cost).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(fAmmoMin4Cost).CopyTo(mySendString, lPos) : lPos += 4
        '      System.BitConverter.GetBytes(fAmmoMin5Cost).CopyTo(mySendString, lPos) : lPos += 4


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

    '    If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
    '    With moResearchCost
    '        .ObjectID = Me.ObjectID
    '        .ObjTypeID = Me.ObjTypeID
    '        Erase .ItemCosts
    '        .ItemCostUB = -1
    '    End With
    '    If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
    '    With moProductionCost
    '        .ObjectID = Me.ObjectID
    '        .ObjTypeID = Me.ObjTypeID
    '        Erase .ItemCosts
    '        .ItemCostUB = -1
    '    End With

    '    Try
    '        Dim fROFSeconds As Single = ROF / 30.0F
    '        Dim bNotStudyFlaw As Boolean = False
    '        Dim fHighestScore As Single = 0.0F

    '        '======== BARREL MATERIALS =========
    '        Dim uBarrel_TempSens As MaterialPropertyItem2
    '        With uBarrel_TempSens
    '            .lMineralID = BarrelMineralID
    '            .lPropertyID = eMinPropID.TemperatureSensitivity

    '            .lActualValue = BarrelMineral.GetPropertyValue(.lPropertyID)
    '            .lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '            .fNormalize = 0.7F
    '            .lGoalValue = 20
    '            If .lActualValue <> .lKnownValue Then
    '                bNotStudyFlaw = True
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '            End If
    '            If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop1
    '                If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '                fHighestScore = .FinalScore
    '            End If
    '        End With
    '        Dim uBarrel_ThermExp As MaterialPropertyItem2
    '        With uBarrel_ThermExp
    '            .lMineralID = BarrelMineralID
    '            .lPropertyID = eMinPropID.ThermalExpansion

    '            .lActualValue = BarrelMineral.GetPropertyValue(.lPropertyID)
    '            .lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '            .fNormalize = 0.2F
    '            .lGoalValue = 10
    '            If .lActualValue <> .lKnownValue Then
    '                bNotStudyFlaw = True
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop2 Or eComponentDesignFlaw.eShift_Not_Study
    '            End If
    '            If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop2
    '                If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '                fHighestScore = .FinalScore
    '            End If
    '        End With
    '        Dim uBarrel_ThermCond As MaterialPropertyItem2
    '        With uBarrel_ThermCond
    '            .lMineralID = BarrelMineralID
    '            .lPropertyID = eMinPropID.ThermalConductance

    '            .lActualValue = BarrelMineral.GetPropertyValue(.lPropertyID)
    '            .lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '            .fNormalize = 0.1F
    '            .lGoalValue = 102 ' 150
    '            If .lActualValue <> .lKnownValue Then
    '                bNotStudyFlaw = True
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop3 Or eComponentDesignFlaw.eShift_Not_Study
    '            End If
    '            If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop3
    '                If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '                fHighestScore = .FinalScore
    '            End If
    '        End With
    '        Dim lBarrelComplexity As Int32 = BarrelMineral.GetPropertyValue(eMinPropID.Complexity)
    '        Dim fBarrelInc As Single = (uBarrel_TempSens.FinalScore + uBarrel_ThermCond.FinalScore + uBarrel_ThermExp.FinalScore)
    '        '========= End of Barrel =============

    '        '========= CASING MATERIALS ==========
    '        Dim uCasing_TempSens As MaterialPropertyItem2
    '        With uCasing_TempSens
    '            .lMineralID = CasingMineralID
    '            .lPropertyID = eMinPropID.TemperatureSensitivity

    '            .lActualValue = CasingMineral.GetPropertyValue(.lPropertyID)
    '            .lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '            .fNormalize = 0.3F
    '            .lGoalValue = 135 ' 100
    '            If .lActualValue <> .lKnownValue Then
    '                bNotStudyFlaw = True
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '            End If
    '            If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop1
    '                If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '                fHighestScore = .FinalScore
    '            End If
    '        End With
    '        Dim uCasing_Density As MaterialPropertyItem2
    '        With uCasing_Density
    '            .lMineralID = CasingMineralID
    '            .lPropertyID = eMinPropID.Density

    '            .lActualValue = CasingMineral.GetPropertyValue(.lPropertyID)
    '            .lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '            .fNormalize = 0.3F
    '            .lGoalValue = 30
    '            If .lActualValue <> .lKnownValue Then
    '                bNotStudyFlaw = True
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop2 Or eComponentDesignFlaw.eShift_Not_Study
    '            End If
    '            If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop2
    '                If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '                fHighestScore = .FinalScore
    '            End If
    '        End With
    '        Dim uCasing_ThermCond As MaterialPropertyItem2
    '        With uCasing_ThermCond
    '            .lMineralID = CasingMineralID
    '            .lPropertyID = eMinPropID.ThermalConductance

    '            .lActualValue = CasingMineral.GetPropertyValue(.lPropertyID)
    '            .lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '            .fNormalize = 0.4F
    '            .lGoalValue = 10
    '            If .lActualValue <> .lKnownValue Then
    '                bNotStudyFlaw = True
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop3 Or eComponentDesignFlaw.eShift_Not_Study
    '            End If
    '            If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop3
    '                If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '                fHighestScore = .FinalScore
    '            End If
    '        End With
    '        Dim lCasingComplexity As Int32 = CasingMineral.GetPropertyValue(eMinPropID.Complexity)
    '        Dim fCasingInc As Single = (uCasing_Density.FinalScore + uCasing_TempSens.FinalScore + uCasing_ThermCond.FinalScore)
    '        '========= End of Casing =============

    '        '========= Payload1 Mineral ==========
    '        Dim uPayload1_Malleable As MaterialPropertyItem2
    '        With uPayload1_Malleable
    '            .lMineralID = Payload1MineralID
    '            .lPropertyID = eMinPropID.Malleable

    '            .lActualValue = Payload1Mineral.GetPropertyValue(.lPropertyID)
    '            .lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '            .fNormalize = 0.2F
    '            .lGoalValue = 10
    '            If .lActualValue <> .lKnownValue Then
    '                bNotStudyFlaw = True
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '            End If
    '            If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop1
    '                If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '                fHighestScore = .FinalScore
    '            End If
    '        End With
    '        Dim uPayload1_Hardness As MaterialPropertyItem2
    '        With uPayload1_Hardness
    '            .lMineralID = Payload1MineralID
    '            .lPropertyID = eMinPropID.Hardness

    '            .lActualValue = Payload1Mineral.GetPropertyValue(.lPropertyID)
    '            .lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '            .fNormalize = 0.4F
    '            .lGoalValue = 103 '100
    '            If .lActualValue <> .lKnownValue Then
    '                bNotStudyFlaw = True
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop2 Or eComponentDesignFlaw.eShift_Not_Study
    '            End If
    '            If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop2
    '                If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '                fHighestScore = .FinalScore
    '            End If
    '        End With
    '        Dim uPayload1_Density As MaterialPropertyItem2
    '        With uPayload1_Density
    '            .lMineralID = Payload1MineralID
    '            .lPropertyID = eMinPropID.Density

    '            .lActualValue = Payload1Mineral.GetPropertyValue(.lPropertyID)
    '            .lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '            .fNormalize = 0.4F
    '            .lGoalValue = 143 '100
    '            If .lActualValue <> .lKnownValue Then
    '                bNotStudyFlaw = True
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop3 Or eComponentDesignFlaw.eShift_Not_Study
    '            End If
    '            If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop3
    '                If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '                fHighestScore = .FinalScore
    '            End If
    '        End With
    '        Dim lP1Complexity As Int32 = Payload1Mineral.GetPropertyValue(eMinPropID.Complexity)
    '        Dim fP1Inc As Single = (uPayload1_Density.FinalScore + uPayload1_Hardness.FinalScore + uPayload1_Malleable.FinalScore)
    '        '========= End of Payload1 ===========

    '        '========= BEGIN Payload2 ============
    '        Dim uPayload2_Property As MaterialPropertyItem2
    '        With uPayload2_Property
    '            .lMineralID = Payload2MineralID

    '            .lGoalValue = 140

    '            Select Case Me.PayloadType
    '                Case 1  'explosive
    '                    .lPropertyID = eMinPropID.Combustiveness
    '                    .lGoalValue = 149
    '                Case 2  'chemical
    '                    .lPropertyID = eMinPropID.ChemicalReactance
    '                    .lGoalValue = 132
    '                Case 3  'magnetic
    '                    .lPropertyID = eMinPropID.MagneticProduction
    '                    .lGoalValue = 152
    '                Case Else
    '                    .lPropertyID = eMinPropID.Density
    '                    .lGoalValue = 144
    '            End Select

    '            .lActualValue = Payload2Mineral.GetPropertyValue(.lPropertyID)
    '            .lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '            .fNormalize = 1.0F
    '            '.lGoalValue = 140
    '            If .lActualValue <> .lKnownValue Then
    '                bNotStudyFlaw = True
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '            End If
    '            If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop1
    '                If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '                fHighestScore = .FinalScore
    '            End If
    '        End With
    '        Dim lP2Complexity As Int32 = Payload2Mineral.GetPropertyValue(eMinPropID.Complexity)
    '        Dim fP2Inc As Single = uPayload2_Property.FinalScore
    '        '============ End of Payload2 ===========

    '        '============ Begin Projection ==========
    '        Dim uProj_Prop As MaterialPropertyItem2
    '        With uProj_Prop
    '            .lMineralID = ProjectionMineralID

    '            If Me.ProjectionType = 0 Then
    '                .lPropertyID = eMinPropID.Combustiveness
    '                .lGoalValue = 132
    '            Else
    '                .lPropertyID = eMinPropID.MagneticProduction
    '                .lGoalValue = 142
    '            End If

    '            .lActualValue = ProjectionMineral.GetPropertyValue(.lPropertyID)
    '            .lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '            .fNormalize = 1.0F
    '            '.lGoalValue = 135
    '            If .lActualValue <> .lKnownValue Then
    '                bNotStudyFlaw = True
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat5_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '            End If
    '            If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '                MajorDesignFlaw = eComponentDesignFlaw.eMat5_Prop1
    '                If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '                fHighestScore = .FinalScore
    '            End If
    '        End With
    '        Dim lProjComplexity As Int32 = ProjectionMineral.GetPropertyValue(eMinPropID.Complexity)
    '        Dim fProjInc As Single = uProj_Prop.FinalScore
    '        '============ End of Production =========

    '        If fHighestScore < 10.0F AndAlso bNotStudyFlaw = False Then
    '            MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eGoodDesign
    '        End If

    '        Dim fModifiedRandomSeed As Single = (0.7F * RandomSeed) + 0.8F

    '        'average ak
    '        Dim fAverageAK As Single = uBarrel_TempSens.AKScore + uBarrel_ThermCond.AKScore + uBarrel_ThermExp.AKScore
    '        fAverageAK += uCasing_Density.AKScore + uCasing_TempSens.AKScore + uCasing_ThermCond.AKScore
    '        fAverageAK += uPayload1_Density.AKScore + uPayload1_Hardness.AKScore + uPayload1_Malleable.AKScore
    '        fAverageAK += uPayload2_Property.AKScore + uProj_Prop.AKScore
    '        fAverageAK /= 11.0F

    '        'average complexity
    '        Dim fAvgComplexity As Single = lBarrelComplexity + lCasingComplexity + lP1Complexity + lP2Complexity + lProjComplexity
    '        fAvgComplexity /= 5.0F

    '        'starting success chance
    '        Dim fTemp As Single = fBarrelInc + fCasingInc + fP1Inc + fP2Inc + fProjInc
    '        fTemp *= -CSng(Me.MaxRange)
    '        Me.CurrentSuccessChance = CInt(fTemp)

    '        'CurrentDPS=ROUNDUP(B2/B4,0)*IF(C8>0,C8+1,1)+B7
    '        Dim fCurrentDPS As Single = CSng(Math.Ceiling(Me.CartridgeSize / fROFSeconds))
    '        If Math.Ceiling(Math.Sqrt(ExplosionRadius)) > 0 Then
    '            fCurrentDPS *= CSng(Math.Ceiling(Math.Sqrt(ExplosionRadius)) + 1)
    '        End If
    '        fCurrentDPS += CSng(ExplosionRadius)

    '        'hull required... requires 2 cells... C41:
    '        '=IF(VLOOKUP(B43,Sheet2!I1:K14,2,TRUE)+1=1,70+(B2*2),VLOOKUP(B43,Sheet2!I1:K14,2,TRUE)+1)
    '        Dim lC41 As Int32 = IToJLookup(CInt(fCurrentDPS)) + 1
    '        If lC41 = 1 Then lC41 = 70 + (CartridgeSize * 2)
    '        'and d41=ROUNDUP(ROUNDUP(randomseed*AVERAGE(I36,I32,I27,I22,I17),0)/10*C41+B7+B2,0)
    '        Dim lD41 As Int32 = CInt(Math.Ceiling(fModifiedRandomSeed * ((fBarrelInc + fCasingInc + fP1Inc + fP2Inc + fProjInc) / 5.0F)))
    '        lD41 = CInt(Math.Ceiling(lD41 / 10.0F * lC41 + CSng(ExplosionRadius) + CSng(CartridgeSize)))
    '        If lC41 > lD41 Then
    '            Me.HullRequired = CInt(Math.Ceiling(lC41 / lD41) * lD41)
    '        Else : Me.HullRequired = lD41
    '        End If
    '        Me.HullRequired \= 2


    '        'Nominal DPS
    '        Dim lNominalDPS As Int32 = Hull2DPS(Me.HullRequired)

    '        'Power=MAX(step3 + IF(B42<11,-140,0),25*B43)
    '        'step1 = ROUNDUP(SUM(I36,I32,I27,I22,I17)*randomseed,0)
    '        fTemp = CSng(Math.Ceiling((fBarrelInc + fCasingInc + fP1Inc + fP2Inc + fProjInc) * fModifiedRandomSeed))
    '        'step2 = ((step1)/19)
    '        fTemp /= 19.0F
    '        'step3 = ROUNDDOWN(step2 * VLOOKUP(B43,Sheet2!C1:E14,3,TRUE) *0.75 + B5 + B7 + B2,0)
    '        fTemp = CSng(Math.Floor(fTemp * CtoELookup(CInt(fCurrentDPS)) * 0.75F + CSng(MaxRange) + CSng(ExplosionRadius) + CSng(CartridgeSize)))
    '        If lNominalDPS < 11 Then fTemp += -140
    '        fTemp = Math.Max(fTemp, 25.0F * fCurrentDPS)
    '        Me.PowerRequired = CInt(fTemp / 2.0F)

    '        'enlisted=ROUNDUP(B4*I22/10,0)+ROUNDDOWN(B2/1000,0)
    '        Dim lEnlisted As Int32 = CInt(Math.Ceiling(fROFSeconds * fCasingInc / 10.0F)) + (CartridgeSize \ 1000)
    '        Dim lOfficers As Int32 = lEnlisted \ 5

    '        'Research Credits=B38*IF(B43>B42,2.5,1)*B44*(B41/1000)*B39/intelligence*(1+B5)*IF(B6>0,B7+1,1)*randomseed*IF(B1=0,1,1.5)/B45
    '        '*B39/intelligence*(1+B5)*IF(B6>0,B7+1,1)*randomseed*IF(B1=0,1,1.5)/B45
    '        Dim dblTemp As Double = fAverageAK
    '        If fCurrentDPS > lNominalDPS Then dblTemp *= 2.5
    '        dblTemp *= Me.PowerRequired
    '        dblTemp *= (Me.HullRequired / 1000.0F)
    '        dblTemp *= fAvgComplexity
    '        dblTemp /= PopIntel
    '        dblTemp *= (1 + CSng(MaxRange))
    '        If PayloadType > 0 Then
    '            dblTemp *= (CSng(ExplosionRadius) + 1)
    '        End If
    '        dblTemp *= fModifiedRandomSeed
    '        If ProjectionType <> 0 Then dblTemp *= 1.5
    '        dblTemp /= lEnlisted

    '        Dim dblBaseResCred As Double = dblTemp

    '        '=IF(B46<1000000,1000000*IF(B43>B42,(1+B43-B42),1),B46)*B38
    '        If dblTemp < 1000000 Then
    '            '1000000*IF(B43>B42,(1+B43-B42),1)
    '            dblTemp = 1000000
    '            If fCurrentDPS > lNominalDPS Then dblTemp *= (1 + fCurrentDPS - lNominalDPS)
    '        End If
    '        Dim blResearchCredits As Int64 = CLng(dblTemp * fAverageAK)

    '        '=B46*B39/intelligence*randomseed
    '        dblTemp = dblBaseResCred * fAvgComplexity
    '        dblTemp /= PopIntel
    '        dblTemp *= fModifiedRandomSeed
    '        '=IF(B47<70000,70000*IF(B43>B42,(1+B43-B42),1),B47/10)
    '        If dblTemp < 70000 Then
    '            dblTemp = 70000
    '            If fCurrentDPS > lNominalDPS Then dblTemp *= (1 + fCurrentDPS - lNominalDPS)
    '        Else : dblTemp /= 10
    '        End If
    '        Dim blResearchPoints As Int64 = CLng(dblTemp)

    '        'Production Credits=B39*randomseed*B41*SUM(I17,I22,I27,I32,I36)
    '        dblTemp = fAvgComplexity * fModifiedRandomSeed * Me.HullRequired * (fBarrelInc + fCasingInc + fP1Inc + fP2Inc + fProjInc)
    '        '=IF(B48<100000,100000*IF(B43>B42,(1+B43-B42),1),B48)
    '        If dblTemp < 100000 Then
    '            dblTemp = 100000
    '            If fCurrentDPS > lNominalDPS Then
    '                dblTemp *= (1 + fCurrentDPS - lNominalDPS)
    '            End If
    '        End If
    '        Dim blProductionCredits As Int64 = CLng(dblTemp)

    '        'ProdPoints =B41*randomseed*B39/intelligence*SUM(I17,I22,I27,I32,I36)*5000
    '        dblTemp = Me.HullRequired * fModifiedRandomSeed * fAvgComplexity / PopIntel * (fBarrelInc + fCasingInc + fP1Inc + fP2Inc + fProjInc) * 5000
    '        '=IF(B49<800000,800000*IF(B43>B42,(1+B43-B42),1),B49)
    '        If dblTemp < 800000 Then
    '            dblTemp = 800000
    '            If fCurrentDPS > lNominalDPS Then
    '                dblTemp *= (1 + fCurrentDPS - lNominalDPS)
    '            End If
    '        End If
    '        Dim blProductionPoints As Int64 = CLng(dblTemp)

    '        Dim lBarrelCost As Int32 = CInt(PopIntel / fAvgComplexity * Me.HullRequired * fBarrelInc)
    '        Dim lCasingCost As Int32 = CInt(PopIntel / fAvgComplexity * Me.HullRequired * fCasingInc)
    '        Dim lP1Cost As Int32 = CInt(fCurrentDPS * PopIntel / fAvgComplexity * fP1Inc)
    '        Dim lP2Cost As Int32 = CInt(fCurrentDPS * 2 * PopIntel / fAvgComplexity * fP2Inc)
    '        Dim lProjectionCost As Int32 = CInt(Me.HullRequired * fProjInc)

    '        lBarrelCost = Math.Max(1, lBarrelCost \ 10)
    '        lCasingCost = Math.Max(1, lCasingCost \ 10)
    '        lP1Cost = Math.Max(1, lP1Cost \ 10)
    '        lP2Cost = Math.Max(1, lP2Cost \ 10)
    '        lProjectionCost = Math.Max(1, lProjectionCost \ 10)

    '        If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
    '        With moResearchCost
    '            .ObjectID = Me.ObjectID
    '            .ObjTypeID = Me.ObjTypeID
    '            .ColonistCost = 0 : .EnlistedCost = 0 : .OfficerCost = 0
    '            .CreditCost = blResearchCredits
    '            Erase .ItemCosts
    '            .ItemCostUB = -1
    '            .PointsRequired = blResearchPoints

    '            .AddProductionCostItem(BarrelMineralID, ObjectType.eMineral, lBarrelCost)
    '            .AddProductionCostItem(CasingMineralID, ObjectType.eMineral, lCasingCost)
    '            .AddProductionCostItem(Payload1MineralID, ObjectType.eMineral, lP1Cost)
    '            .AddProductionCostItem(Payload2MineralID, ObjectType.eMineral, lP2Cost)
    '            .AddProductionCostItem(ProjectionMineralID, ObjectType.eMineral, lProjectionCost)
    '        End With

    '        lBarrelCost = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fBarrelInc))
    '        lCasingCost = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fCasingInc))
    '        lP1Cost = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fP1Inc))
    '        lP2Cost = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fP2Inc))
    '        lProjectionCost = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fProjInc))

    '        lBarrelCost = Math.Max(1, lBarrelCost \ 10)
    '        lCasingCost = Math.Max(1, lCasingCost \ 10)
    '        lP1Cost = Math.Max(1, lP1Cost \ 10)
    '        lP2Cost = Math.Max(1, lP2Cost \ 10)
    '        lProjectionCost = Math.Max(1, lProjectionCost \ 10)

    '        If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
    '        With moProductionCost
    '            .ObjectID = Me.ObjectID
    '            .ObjTypeID = Me.ObjTypeID
    '            .ColonistCost = 0 : .EnlistedCost = lEnlisted : .OfficerCost = lOfficers
    '            .CreditCost = blProductionCredits
    '            Erase .ItemCosts
    '            .ItemCostUB = -1
    '            .PointsRequired = blProductionPoints

    '            .AddProductionCostItem(BarrelMineralID, ObjectType.eMineral, lBarrelCost)
    '            .AddProductionCostItem(CasingMineralID, ObjectType.eMineral, lCasingCost)
    '            .AddProductionCostItem(Payload1MineralID, ObjectType.eMineral, lP1Cost)
    '            .AddProductionCostItem(Payload2MineralID, ObjectType.eMineral, lP2Cost)
    '            .AddProductionCostItem(ProjectionMineralID, ObjectType.eMineral, lProjectionCost)
    '        End With

    '    Catch ex As Exception
    '        Me.ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
    '        Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
    '    End Try
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

            Dim oComputer As New ProjectileTechComputer
            With oComputer
                .lHullTypeID = Me.yHullTypeID

                Dim fROF As Single = Math.Max(1, ROF) / 30.0F
                Dim lHalfCart As Int32 = CartridgeSize \ 2

                Dim lMinPierce As Int32 = 0
                Dim lMinImpact As Int32 = 0
                Dim lMaxPierce As Int32 = CInt(lHalfCart * (CSng(PierceRatio) / 100.0F))
                Dim lMaxImpact As Int32 = CInt(lHalfCart * ((100 - CSng(PierceRatio)) / 100.0F))
                Dim lMinPayload As Int32 = 0
                Dim lMaxPayload As Int32 = lHalfCart

                Dim lMaxFlame As Int32 = 0
                Dim lMinFlame As Int32 = 0
                Dim lMaxChemical As Int32 = 0
                Dim lMinChemical As Int32 = 0
                Dim lMaxECM As Int32 = 0
                Dim lMinECM As Int32 = 0
                Select Case PayloadType
                    Case 1
                        lMaxFlame = lMaxPayload : lMinFlame = lMinPayload
                    Case 2
                        lMaxChemical = lMaxPayload : lMinChemical = lMinPayload
                    Case 3
                        lMaxECM = lMaxPayload : lMinECM = lMinPayload
                    Case Else
                        lMaxImpact += lMaxPayload
                        lMinImpact += lMinPayload
                End Select
                .decDPS = (CDec(lMaxPierce) + CDec(lMaxImpact) + CDec(lMaxFlame) + CDec(lMaxChemical) + CDec(lMaxECM)) / CDec(fROF)
                .decDPS = Math.Max(.decDPS, (.decDPS + CDec(lMaxPierce) + CDec(lMaxImpact) + CDec(lMaxFlame) + CDec(lMaxChemical) + CDec(lMaxECM)) / 2D)

                Dim decNormalizer As Decimal = 1D
                Dim lMaxGuns As Int32 = 1
                Dim lMaxDPS As Int32 = 1
                Dim lMaxHullSize As Int32 = 1
                Dim lHullAvail As Int32 = 1
                Dim lMaxPower As Int32 = 1
                TechBuilderComputer.GetTypeValues(yHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)

                .lMineral1ID = BarrelMineralID
                .lMineral2ID = CasingMineralID
                .lMineral3ID = Payload1MineralID
                .lMineral4ID = Payload2MineralID
                .lMineral5ID = ProjectionMineralID
                .lMineral6ID = -1

                .blCartridgeSize = CartridgeSize
                .blExplosionRadius = ExplosionRadius
                .blMaxRange = MaxRange
                .blPierceRatio = PierceRatio
                .blROF = ROF
                .yPayloadType = PayloadType
                .yProjectionType = ProjectionType

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


    Private Function IToJLookup(ByVal lValue As Int32) As Int32
        If lValue < 10 Then
            Return 0
        ElseIf lValue < 12 Then
            Return 90
        ElseIf lValue < 14 Then
            Return 120
        ElseIf lValue < 21 Then
            Return 170
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
    Private Function Hull2DPS(ByVal lValue As Int32) As Int32
        If lValue < 120 Then
            Return 10
        ElseIf lValue < 170 Then
            Return 12
        ElseIf lValue < 300 Then
            Return 14
        ElseIf lValue < 500 Then
            Return 40
        ElseIf lValue < 2000 Then
            Return 21
        ElseIf lValue < 8000 Then
            Return 55
        ElseIf lValue < 20000 Then
            Return 60
        ElseIf lValue < 40000 Then
            Return 125
        ElseIf lValue < 80000 Then
            Return 250
        ElseIf lValue < 250000 Then
            Return 500
        ElseIf lValue < 1000000 Then
            Return 1900
        Else : Return 15000
        End If
    End Function
    Private Function CtoELookup(ByVal lValue As Int32) As Int32
        If lValue < 10 Then
            Return 250
        ElseIf lValue < 12 Then
            Return 303
        ElseIf lValue < 14 Then
            Return 100
        ElseIf lValue < 21 Then
            Return 347 '1042
        ElseIf lValue < 40 Then
            Return 536 '347
        ElseIf lValue < 55 Then
            Return 1042 '536
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
        'Do Nothing
    End Sub

    Public Overrides Function GetNonOwnerMsg() As Byte()
        Dim yResult(48) As Byte
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
        System.BitConverter.GetBytes(CartridgeSize).CopyTo(yResult, lPos) : lPos += 4
        System.BitConverter.GetBytes(MaxRange).CopyTo(yResult, lPos) : lPos += 2
        yResult(lPos) = ExplosionRadius : lPos += 1
        yResult(lPos) = PierceRatio : lPos += 1

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
            .Accuracy = 80                  'TODO: What should this be?
            .AmmoReloadDelay = 0
            .AmmoSize = 0
            .AOERange = ExplosionRadius

            '         Dim fCartridgeSizePower As Single = CSng(Math.Sqrt(Math.Pow(CartridgeSize, 2.25F)))

            '         mlMinPierce = CInt(Math.Floor((CSng(PierceRatio) / 100.0F) * fCartridgeSizePower * mfPayload1PotentialPierce))
            '         mlMaxPierce = CInt(Math.Floor((CSng(PierceRatio) / 100.0F) * fCartridgeSizePower * 3 * mfPayload1PotentialPierce))
            '         mlMinImpact = CInt(Math.Floor(((100.0F - CSng(PierceRatio)) / 100.0F) * (fCartridgeSizePower * 2) * mfPayload1PotentialImpact))
            '         mlMaxImpact = CInt(Math.Floor(((100.0F - CSng(PierceRatio)) / 100.0F) * (fCartridgeSizePower * 4) * mfPayload1PotentialImpact))
            '         mlMinPayload = CInt(Math.Floor(mfPayload2Potential / 100.0F))
            'mlMaxPayload = CInt(Math.Floor(mfPayload2Potential / 80.0F))

            'ok, 1/2 of the total damage is Payload1 (pierce/impact)
            'the other half is payload2
            Dim lPayloadDmg As Int32 = CartridgeSize \ 2
            Dim lPI_Dmg As Int32 = lPayloadDmg '\ 2
            Dim lMagDmg As Int32 = 0

            If ProjectionType = 1 Then
                lMagDmg = lPI_Dmg \ 2
                lPI_Dmg = lMagDmg
            End If


            mlMinPierce = 0
            mlMinImpact = 0
            mlMaxPierce = CInt(lPI_Dmg * (PierceRatio / 100.0F))
            mlMaxImpact = CInt(lPI_Dmg * ((100 - PierceRatio) / 100.0F))
            mlMinPayload = 0
            mlMaxPayload = lPayloadDmg

            .BeamMaxDmg = 0 : .BeamMinDmg = 0
            .PiercingMaxDmg = mlMaxPierce : .PiercingMinDmg = mlMinPierce
            .ImpactMaxDmg = mlMaxImpact : .ImpactMinDmg = mlMinImpact

            .ChemicalMaxDmg = 0 : .ChemicalMinDmg = 0
            .ECMMaxDmg = lMagDmg : .ECMMinDmg = 0
            .FlameMaxDmg = 0 : .FlameMinDmg = 0

            Select Case PayloadType
                Case 1
                    .FlameMaxDmg = mlMaxPayload : .FlameMinDmg = mlMinPayload
                Case 2
                    .ChemicalMaxDmg = mlMaxPayload : .ChemicalMinDmg = mlMinPayload
                Case 3
                    .ECMMaxDmg += mlMaxPayload : .ECMMinDmg = mlMinPayload
                Case Else
                    .ImpactMaxDmg += mlMaxPayload
                    .ImpactMinDmg += mlMinPayload
            End Select

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
            Dim lMinDmg As Int32 = .BeamMinDmg + .ChemicalMinDmg + .ECMMinDmg + .FlameMinDmg + .ImpactMinDmg + .PiercingMinDmg
            Dim lMaxDmg As Int32 = .BeamMaxDmg + .ChemicalMaxDmg + .ECMMaxDmg + .FlameMaxDmg + .ImpactMaxDmg + .PiercingMaxDmg

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
                   "MaxSpeed, ShotHullSize, PierceRating, OptimumRange, PayloadType, ExplosionRadius, " & _
                   "Mineral1ID, Mineral2ID, Mineral3ID, Mineral4ID, Mineral5ID, ErrorReasonCode, Payload1PotentialPierce, " & _
                   "Payload1PotentialImpact, Payload2Potential, ProjectileHullSize, MajorDesignFlaw, PopIntel, bArchived, HullTypeID) VALUES ('" & _
                   MakeDBStr(BytesToString(WeaponName)) & "', " & CByte(WeaponClassTypeID) & ", " & CByte(WeaponTypeID) & ", " & _
                   PowerRequired & ", " & HullRequired & ", " & CurrentSuccessChance & ", " & SuccessChanceIncrement & ", " & _
                   ResearchAttempts & ", " & RandomSeed & ", " & CInt(ComponentDevelopmentPhase) & ", "
                If Me.Owner Is Nothing = False Then
                    sSQL &= Owner.ObjectID & ", "
                Else : sSQL &= "-1, "
                End If
                sSQL &= ROF & ", " & ProjectionType & ", " & CartridgeSize & ", " & PierceRatio & ", " & MaxRange & ", " & _
                  PayloadType & ", " & ExplosionRadius & ", " & BarrelMineralID & ", " & CasingMineralID & ", " & _
                  Payload1MineralID & ", " & Payload2MineralID & ", " & ProjectionMineralID & ", " & ErrorReasonCode & _
                  ", " & mfPayload1PotentialPierce & ", " & mfPayload1PotentialImpact & ", " & mfPayload2Potential & _
                  ", 0, " & MajorDesignFlaw & ", " & PopIntel & ", " & yArchived & ", " & yHullTypeID & ")"
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
                sSQL &= ROF & ", MaxSpeed = " & ProjectionType & ", ShotHullSize = " & CartridgeSize & ", PierceRating = " & _
                  PierceRatio & ", OptimumRange = " & MaxRange & ", PayloadType = " & PayloadType & ", ExplosionRadius = " & _
                  ExplosionRadius & ", Mineral1ID = " & BarrelMineralID & ", Mineral2ID = " & CasingMineralID & ", Mineral3ID = " & _
                  Payload1MineralID & ", Mineral4ID = " & Payload2MineralID & ", Mineral5ID = " & ProjectionMineralID & _
                  ", ErrorReasonCode = " & ErrorReasonCode & ", Payload1PotentialPierce = " & mfPayload1PotentialPierce & _
                  ", Payload1PotentialImpact = " & mfPayload1PotentialImpact & ", Payload2Potential = " & mfPayload2Potential & _
                  ", ProjectileHullSize = 0, MajorDesignFlaw = " & MajorDesignFlaw & ", PopIntel = " & PopIntel & _
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

                .ProjectionType = yData(lPos) : lPos += 1
                .CartridgeSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .PierceRatio = yData(lPos) : lPos += 1
                .ROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .MaxRange = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .PayloadType = yData(lPos) : lPos += 1
                .ExplosionRadius = yData(lPos) : lPos += 1


                .BarrelMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .CasingMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .Payload1MineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .Payload2MineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .ProjectionMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                ReDim .WeaponName(19)
                Array.Copy(yData, lPos, .WeaponName, 0, 20)
                lPos += 20
            End With

            bRes = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "ProjectileWeaponTech.SetFromDesignMsg: " & ex.Message)
        End Try

        Return bRes
    End Function

    Public Overrides Function ValidateDesign() As Boolean
        Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidMaterialsSelected
        If BarrelMineralID < 1 OrElse Owner.IsMineralDiscovered(BarrelMineralID) = False Then
            LogEvent(LogEventType.PossibleCheat, "ProjectileWeaponTech.ValidateDesign invalid materials: " & Me.Owner.ObjectID)
            Return False
        End If
        If CasingMineralID < 1 OrElse Owner.IsMineralDiscovered(CasingMineralID) = False Then
            LogEvent(LogEventType.PossibleCheat, "ProjectileWeaponTech.ValidateDesign invalid materials: " & Me.Owner.ObjectID)
            Return False
        End If
        If Payload1MineralID < 1 OrElse Owner.IsMineralDiscovered(Payload1MineralID) = False Then
            LogEvent(LogEventType.PossibleCheat, "ProjectileWeaponTech.ValidateDesign invalid materials: " & Me.Owner.ObjectID)
            Return False
        End If
        If Payload2MineralID < 1 OrElse Owner.IsMineralDiscovered(Payload2MineralID) = False Then
            LogEvent(LogEventType.PossibleCheat, "ProjectileWeaponTech.ValidateDesign invalid materials: " & Me.Owner.ObjectID)
            Return False
        End If
        If ProjectionMineralID < 1 OrElse Owner.IsMineralDiscovered(ProjectionMineralID) = False Then
            LogEvent(LogEventType.PossibleCheat, "ProjectileWeaponTech.ValidateDesign invalid materials: " & Me.Owner.ObjectID)
            Return False
        End If

        Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
        If CartridgeSize < 1 Then
            LogEvent(LogEventType.PossibleCheat, "ProjectileWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If
        If ROF < 30 Then
            LogEvent(LogEventType.PossibleCheat, "ProjectileWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If
        If MaxRange < 1 Then
            LogEvent(LogEventType.PossibleCheat, "ProjectileWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If

        If ExplosionRadius > Owner.oSpecials.yProjExpRadius Then
            LogEvent(LogEventType.PossibleCheat, "ProjectileWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If
        If MaxRange > Owner.oSpecials.iProjMaxRange Then
            LogEvent(LogEventType.PossibleCheat, "ProjectileWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If
        If PayloadType > Owner.oSpecials.yPayloadTypeAvailable Then
            LogEvent(LogEventType.PossibleCheat, "ProjectileWeaponTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If

        If CheckDesignImpossibility() = True Then
            LogEvent(LogEventType.PossibleCheat, "ProjectileWeaponTech.ValidateDesign CheckDesignPossiblity Failed: " & Me.Owner.ObjectID)
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

    '			Dim lTemp As Int32 = CInt(fAmmoMin2Cost)
    '			If lTemp <> 0 Then .AddProductionCostItem(CasingMineralID, ObjectType.eMineral, lTemp)
    '			lTemp = CInt(fAmmoMin3Cost)
    '			If lTemp <> 0 Then .AddProductionCostItem(Payload1MineralID, ObjectType.eMineral, lTemp)
    '			lTemp = CInt(fAmmoMin4Cost)
    '			If lTemp <> 0 Then .AddProductionCostItem(Payload2MineralID, ObjectType.eMineral, lTemp)
    '			lTemp = CInt(fAmmoMin5Cost)
    '			If lTemp <> 0 Then .AddProductionCostItem(ProjectionMineralID, ObjectType.eMineral, lTemp)
    '		End With
    '	End If
    '	Return moAmmoCost
    'End Function

    Public Overrides Function TechnologyScore() As Integer
        If mlStoredTechScore = Int32.MinValue Then
            Try
                Dim fROF As Single = Me.ROF / 30.0F
                mlStoredTechScore = CInt(((((Me.CartridgeSize + Me.ExplosionRadius) / fROF) * Me.MaxRange * (1.0F + (CInt(PayloadType) / 10.0F))) / (Me.PowerRequired + Me.HullRequired)) * 10)
            Catch
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
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then lPos += 24
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then lPos += 120 '48
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
            yMsg(lPos) = ProjectionType : lPos += 1
            System.BitConverter.GetBytes(CartridgeSize).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = PierceRatio : lPos += 1
            System.BitConverter.GetBytes(ROF).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(MaxRange).CopyTo(yMsg, lPos) : lPos += 2
            yMsg(lPos) = PayloadType : lPos += 1
            yMsg(lPos) = ExplosionRadius : lPos += 1
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
            System.BitConverter.GetBytes(BarrelMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(CasingMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(Payload1MineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(Payload2MineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(ProjectionMineralID).CopyTo(yMsg, lPos) : lPos += 4

            System.BitConverter.GetBytes(mlMinPierce).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(mlMaxPierce).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(mlMinImpact).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(mlMaxImpact).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(mlMinPayload).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(mlMaxPayload).CopyTo(yMsg, lPos) : lPos += 4
            'System.BitConverter.GetBytes(mfAmmoSize).CopyTo(yMsg, lPos) : lPos += 4

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
              CByte(WeaponClassTypeID) & ", WeaponTypeID = " & CByte(WeaponTypeID) & ", PowerRequired = " & _
              PowerRequired & ", HullRequired = " & HullRequired & ", CurrentSuccessChance = " & _
              CurrentSuccessChance & ", SuccessChanceIncrement = " & SuccessChanceIncrement & ", ResearchAttempts = " & _
              ResearchAttempts & ", RandomSeed = " & RandomSeed & ", ResPhase = " & CInt(ComponentDevelopmentPhase) & _
              ", OwnerID = "
            If Me.Owner Is Nothing = False Then
                sSQL &= Owner.ObjectID & ", ROF = "
            Else : sSQL &= "-1, ROF = "
            End If
            sSQL &= ROF & ", MaxSpeed = " & ProjectionType & ", ShotHullSize = " & CartridgeSize & ", PierceRating = " & _
              PierceRatio & ", OptimumRange = " & MaxRange & ", PayloadType = " & PayloadType & ", ExplosionRadius = " & _
              ExplosionRadius & ", Mineral1ID = " & BarrelMineralID & ", Mineral2ID = " & CasingMineralID & ", Mineral3ID = " & _
              Payload1MineralID & ", Mineral4ID = " & Payload2MineralID & ", Mineral5ID = " & ProjectionMineralID & _
              ", ErrorReasonCode = " & ErrorReasonCode & ", Payload1PotentialPierce = " & mfPayload1PotentialPierce & _
              ", Payload1PotentialImpact = " & mfPayload1PotentialImpact & ", Payload2Potential = " & mfPayload2Potential & _
              ", ProjectileHullSize = 0, MajorDesignFlaw = " & MajorDesignFlaw & ", PopIntel = " & PopIntel & _
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

            .ProjectionType = yData(lPos) : lPos += 1
            .CartridgeSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .PierceRatio = yData(lPos) : lPos += 1
            .ROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .MaxRange = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .PayloadType = yData(lPos) : lPos += 1
            .ExplosionRadius = yData(lPos) : lPos += 1

            .BarrelMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .CasingMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Payload1MineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Payload2MineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ProjectionMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            .mlMinPierce = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .mlMaxPierce = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .mlMinImpact = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .mlMaxImpact = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .mlMinPayload = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .mlMaxPayload = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

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

        Dim oComputer As New ProjectileTechComputer
        With oComputer
            .lHullTypeID = Me.yHullTypeID

            Dim fROF As Single = Math.Max(1, ROF) / 30.0F
            Dim lHalfCart As Int32 = CartridgeSize \ 2

            Dim lMinPierce As Int32 = 0
            Dim lMinImpact As Int32 = 0
            Dim lMaxPierce As Int32 = CInt(lHalfCart * (CSng(PierceRatio) / 100.0F))
            Dim lMaxImpact As Int32 = CInt(lHalfCart * ((100 - CSng(PierceRatio)) / 100.0F))
            Dim lMinPayload As Int32 = 0
            Dim lMaxPayload As Int32 = lHalfCart

            Dim lMaxFlame As Int32 = 0
            Dim lMinFlame As Int32 = 0
            Dim lMaxChemical As Int32 = 0
            Dim lMinChemical As Int32 = 0
            Dim lMaxECM As Int32 = 0
            Dim lMinECM As Int32 = 0
            Select Case PayloadType
                Case 1
                    lMaxFlame = lMaxPayload : lMinFlame = lMinPayload
                Case 2
                    lMaxChemical = lMaxPayload : lMinChemical = lMinPayload
                Case 3
                    lMaxECM = lMaxPayload : lMinECM = lMinPayload
                Case Else
                    lMaxImpact += lMaxPayload
                    lMinImpact += lMinPayload
            End Select
            .decDPS = (CDec(lMaxPierce) + CDec(lMaxImpact) + CDec(lMaxFlame) + CDec(lMaxChemical) + CDec(lMaxECM)) / CDec(fROF)
            .decDPS = Math.Max(.decDPS, (.decDPS + CDec(lMaxPierce) + CDec(lMaxImpact) + CDec(lMaxFlame) + CDec(lMaxChemical) + CDec(lMaxECM)) / 2D)

            Dim decNormalizer As Decimal = 1D
            Dim lMaxGuns As Int32 = 1
            Dim lMaxDPS As Int32 = 1
            Dim lMaxHullSize As Int32 = 1
            Dim lHullAvail As Int32 = 1
            Dim lMaxPower As Int32 = 1
            TechBuilderComputer.GetTypeValues(yHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)

            .lMineral1ID = BarrelMineralID
            .lMineral2ID = CasingMineralID
            .lMineral3ID = Payload1MineralID
            .lMineral4ID = Payload2MineralID
            .lMineral5ID = ProjectionMineralID
            .lMineral6ID = -1

            .blCartridgeSize = CartridgeSize
            .blExplosionRadius = ExplosionRadius
            .blMaxRange = MaxRange
            .blPierceRatio = PierceRatio
            .blROF = ROF
            .yPayloadType = PayloadType
            .yProjectionType = ProjectionType

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
                Dim yOrig As Byte = MajorDesignFlaw
                If MajorDesignFlaw = 0 OrElse (MajorDesignFlaw And eComponentDesignFlaw.eGoodDesign) <> 0 Then MajorDesignFlaw = eComponentDesignFlaw.eMicroTech
                If yOrig <> MajorDesignFlaw Then Me.SaveObject()
            End If
        End With
    End Sub
End Class
