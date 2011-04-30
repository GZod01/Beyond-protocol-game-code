Public Class ShieldTech
    Inherits Epica_Tech

    Public ShieldName(19) As Byte
    Public MaxHitPoints As Int32
    Public RechargeRate As Int32
    Public RechargeFreq As Int32

    Public lProjectionHullSize As Int32
    Public lCoilMineralID As Int32
    Public lAcceleratorMineralID As Int32
    Public lCasingMineralID As Int32

    Public PowerRequired As Int32
    Public HullRequired As Int32

    Public ColorValue As Byte

    Public HullTypeID As Byte

    Private mySendString() As Byte

    Public Function GetObjAsString() As Byte()
        If moResearchCost Is Nothing OrElse moProductionCost Is Nothing Then
            CalculateBothProdCosts()
            mbStringReady = False
        End If

        'here we will return the entire object as a string
        'If mbStringReady = False Then
        ReDim mySendString(Epica_Tech.BASE_OBJ_STRING_SIZE + 57)

        Dim lPos As Int32
        MyBase.GetBaseObjAsString(False).CopyTo(mySendString, 0)
        lPos = Epica_Tech.BASE_OBJ_STRING_SIZE

        ShieldName.CopyTo(mySendString, lPos) : lPos += 20
        System.BitConverter.GetBytes(MaxHitPoints).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(RechargeRate).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(RechargeFreq).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lProjectionHullSize).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lCasingMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lCoilMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lAcceleratorMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(PowerRequired).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(HullRequired).CopyTo(mySendString, lPos) : lPos += 4
        mySendString(lPos) = ColorValue : lPos += 1
        mySendString(lPos) = HullTypeID : lPos += 1

        'mbStringReady = True
        'End If
        Return mySendString
    End Function

    Public Overrides Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblShield (ShieldName, OwnerID, MaxHitPoints, RechargeRate, ProjectionHullSize, " & _
                  "CasingMineralID, CoilMineralID, AcceleratorMineralID, PowerRequired, HullRequired, RechargeFreq" & _
                  ", CurrentSuccessChance, SuccessChanceIncrement, ResearchAttempts, RandomSeed, ResPhase, ColorValue, " & _
                  "ErrorReasonCode, MajorDesignFlaw, PopIntel, bArchived, HullTypeID) VALUES ('" & MakeDBStr(BytesToString(ShieldName)) & "', " & Owner.ObjectID & ", " & _
                  MaxHitPoints & ", " & RechargeRate & ", " & lProjectionHullSize & ", " & lCasingMineralID & _
                  ", " & lCoilMineralID & ", " & lAcceleratorMineralID & ", " & PowerRequired & ", " & _
                  HullRequired & ", " & RechargeFreq & _
                  ", " & CurrentSuccessChance & ", " & SuccessChanceIncrement & ", " & ResearchAttempts & _
                  ", " & RandomSeed & ", " & ComponentDevelopmentPhase & ", " & ColorValue & ", " & ErrorReasonCode & _
                  ", " & MajorDesignFlaw & ", " & PopIntel & ", " & yArchived & ", " & HullTypeID & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblShield SET ShieldName = '" & MakeDBStr(BytesToString(ShieldName)) & "', OwnerID = " & Owner.ObjectID & _
                  ", MaxHitPoints = " & MaxHitPoints & ", RechargeRate = " & RechargeRate & ", ProjectionHullSize=" & _
                  lProjectionHullSize & ", CasingMineralID = " & lCasingMineralID & ", CoilMineralID = " & _
                  lCoilMineralID & ", AcceleratorMineralID = " & lAcceleratorMineralID & ", PowerRequired = " & _
                  PowerRequired & ", HullRequired = " & HullRequired & ", RechargeFreq = " & _
                  RechargeFreq & _
                  ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
                  SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & _
                  RandomSeed & ", ResPhase = " & ComponentDevelopmentPhase & ", ColorValue = " & ColorValue & _
                  ", ErrorReasonCode = " & Me.ErrorReasonCode & ", MajorDesignFlaw = " & MajorDesignFlaw & _
                  ", PopIntel = " & PopIntel & ", bArchived = " & yArchived & ", HullTypeID = " & HullTypeID & " WHERE ShielDID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(ShieldID) FROM tblShield WHERE ShieldName = '" & MakeDBStr(BytesToString(ShieldName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing

                ''Ok, set our object in the master array
                'Dim bAdded As Boolean = False
                'For X As Int32 = 0 To glShieldUB
                '    If glShieldIdx(X) = -1 Then
                '        glShieldIdx(X) = ObjectID
                '        goShield(X) = Me
                '        bAdded = True
                '        Exit For
                '    End If
                'Next X
                'If bAdded = True Then
                '    glShieldUB += 1
                '    ReDim Preserve glShieldIdx(glShieldUB)
                '    ReDim Preserve goShield(glShieldUB)
                '    glShieldIdx(glShieldUB) = ObjectID
                '    goShield(glShieldUB) = Me
                'End If
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

    Public Function GetSaveObjectText() As String
        Dim sSQL As String

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!
        If ObjectID = -1 Then
            SaveObject()
            Return ""
        End If

        Try
            Dim oSB As New System.Text.StringBuilder

            sSQL = "UPDATE tblShield SET ShieldName = '" & MakeDBStr(BytesToString(ShieldName)) & "', OwnerID = " & Owner.ObjectID & _
              ", MaxHitPoints = " & MaxHitPoints & ", RechargeRate = " & RechargeRate & ", ProjectionHullSize=" & _
              lProjectionHullSize & ", CasingMineralID = " & lCasingMineralID & ", CoilMineralID = " & _
              lCoilMineralID & ", AcceleratorMineralID = " & lAcceleratorMineralID & ", PowerRequired = " & _
              PowerRequired & ", HullRequired = " & HullRequired & ", RechargeFreq = " & _
              RechargeFreq & _
              ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
              SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & _
              RandomSeed & ", ResPhase = " & ComponentDevelopmentPhase & ", ColorValue = " & ColorValue & _
              ", ErrorReasonCode = " & Me.ErrorReasonCode & ", MajorDesignFlaw = " & MajorDesignFlaw & _
              ", PopIntel = " & PopIntel & ", bArchived = " & yArchived & ", HullTypeID = " & HullTypeID & " WHERE ShielDID = " & ObjectID
            oSB.AppendLine(sSQL)
            oSB.AppendLine(MyBase.GetFinalizeSaveText())
            Return oSB.ToString
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        End Try
        Return ""
    End Function

    Private moCoilMineral As Mineral
    Public ReadOnly Property CoilMineral() As Mineral
        Get
            If moCoilMineral Is Nothing OrElse moCoilMineral.ObjectID <> lCoilMineralID Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = lCoilMineralID Then
                        moCoilMineral = goMineral(X)
                        Exit For
                    End If
                Next X
            End If
            Return moCoilMineral
        End Get
    End Property

    Private moAcceleratorMineral As Mineral
    Public ReadOnly Property AcceleratorMineral() As Mineral
        Get
            If moAcceleratorMineral Is Nothing OrElse moAcceleratorMineral.ObjectID <> lAcceleratorMineralID Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = lAcceleratorMineralID Then
                        moAcceleratorMineral = goMineral(X)
                        Exit For
                    End If
                Next X
            End If
            Return moAcceleratorMineral
        End Get
    End Property

    Private moCasingMineral As Mineral
    Public ReadOnly Property CasingMineral() As Mineral
        Get
            If moCasingMineral Is Nothing OrElse moCasingMineral.ObjectID <> lCasingMineralID Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = lCasingMineralID Then
                        moCasingMineral = goMineral(X)
                        Exit For
                    End If
                Next X
            End If
            Return moCasingMineral
        End Get
    End Property

    Public Overrides Function ValidateDesign() As Boolean
        Dim bMinGood As Boolean = _
          (Me.lAcceleratorMineralID > 0 AndAlso Owner.IsMineralDiscovered(Me.lAcceleratorMineralID)) AndAlso _
          (Me.lCasingMineralID > 0 AndAlso Owner.IsMineralDiscovered(Me.lCasingMineralID)) AndAlso _
          (Me.lCoilMineralID > 0 AndAlso Owner.IsMineralDiscovered(Me.lCoilMineralID))

        Dim bValues As Boolean = Me.lProjectionHullSize > 0 AndAlso Me.RechargeRate > 0 AndAlso _
          Me.MaxHitPoints >= Me.RechargeRate AndAlso Me.RechargeFreq > 1

        If bMinGood = False Then
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidMaterialsSelected
            LogEvent(LogEventType.PossibleCheat, "ShieldTech.ValidateDesign invalid minerals: " & Me.Owner.ObjectID)
            Return False
        ElseIf bValues = False Then
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
            LogEvent(LogEventType.PossibleCheat, "ShieldTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        Else
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
            If MaxHitPoints > Owner.oSpecials.lShieldMaxHP Then
                LogEvent(LogEventType.PossibleCheat, "ShieldTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
            If lProjectionHullSize > Owner.oSpecials.lShieldProjHullSize Then
                LogEvent(LogEventType.PossibleCheat, "ShieldTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
            If RechargeFreq < Owner.oSpecials.iShieldRechargeFreq Then
                LogEvent(LogEventType.PossibleCheat, "ShieldTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
            If RechargeRate > Owner.oSpecials.lShieldRechargeRate Then
                LogEvent(LogEventType.PossibleCheat, "ShieldTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
            Me.ErrorReasonCode = TechBuilderErrorReason.eNoError
        End If

        If CheckDesignImpossibility() = True Then
            LogEvent(LogEventType.PossibleCheat, "ShieldTech.ValidateDesign CheckDesignPossiblity Failed: " & Me.Owner.ObjectID)
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
            Return False
        End If

        Return bMinGood AndAlso bValues
    End Function


    Private Function CheckDesignImpossibility() As Boolean
        Dim oTB As New ShieldTechComputer
        With Me
            oTB.blLockedProdCost = .blSpecifiedProdCost
            oTB.blLockedProdTime = .blSpecifiedProdTime
            oTB.blLockedResCost = .blSpecifiedResCost
            oTB.blLockedResTime = .blSpecifiedResTime

            oTB.lHullTypeID = .HullTypeID
            oTB.lLockedColonists = .lSpecifiedColonists
            oTB.lLockedEnlisted = .lSpecifiedEnlisted
            oTB.lLockedHull = .lSpecifiedHull
            oTB.lLockedMin1 = .lSpecifiedMin1
            oTB.lLockedMin2 = .lSpecifiedMin2
            oTB.lLockedMin3 = .lSpecifiedMin3
            oTB.lLockedMin4 = .lSpecifiedMin4
            oTB.lLockedMin5 = .lSpecifiedMin5
            oTB.lLockedMin6 = .lSpecifiedMin6
            oTB.lLockedOfficers = .lSpecifiedOfficers
            oTB.lLockedPower = .lSpecifiedPower
            oTB.lMineral1ID = .lCoilMineralID 'lAcceleratorMineralID
            oTB.lMineral2ID = .lAcceleratorMineralID 'lCasingMineralID
            oTB.lMineral3ID = .lCasingMineralID 'lCoilMineralID
            oTB.lMineral4ID = -1
            oTB.lMineral5ID = -1
            oTB.lMineral6ID = -1

            'Bogus: Do not back out specials.  These bonus's are applied after the design is finalized.
            'Dim fMult As Single = 1.0F + (Owner.oSpecials.yShieldMaxHPBonus / 100.0F)
            'fMult = 1.0F / fMult
            'oTB.blMaxHP = CLng(Math.Ceiling(.MaxHitPoints * fMult))

            'fMult = 1.0F + (Owner.oSpecials.yShieldProjHullSizeBonus / 100.0F)
            'fMult = 1.0F / fMult
            'oTB.blProjHull = CLng(Math.Ceiling(.lProjectionHullSize * fMult))

            'fMult = 1.0F - (Owner.oSpecials.yShieldRechargeFreqBonus / 100.0F)
            'fMult = 1.0F / fMult

            'Dim iRechFreq As Int16 = CShort(.RechargeFreq * fMult)

            'fMult = 1.0F + (Owner.oSpecials.yShieldRechargeRateBonus / 100.0F)
            'fMult = 1.0F / fMult
            'Dim lRechRate As Int32 = CInt(.RechargeRate * fMult)
            oTB.blMaxHP = .MaxHitPoints
            oTB.blProjHull = lProjectionHullSize
            Dim iRechFreq As Int16 = CShort(.RechargeFreq)
            Dim lRechRate As Int32 = CInt(.RechargeRate)

            Dim decRInt As Decimal = Math.Max(1, iRechFreq) / 30D
            oTB.decHPS = lRechRate / decRInt
        End With
        Return oTB.IsDesignImpossible(Me.ObjTypeID, Owner)
    End Function

    'Private Function CalculateMedian(ByVal ParamArray Values() As Int32) As Int32
    '    Dim lTmp() As Int32

    '    ReDim lTmp(Values.GetUpperBound(0))

    '    Dim lLastLowest As Int32 = Int32.MinValue
    '    Dim lLowest As Int32 = Int32.MaxValue

    '    Dim lValsAtLowest As Int32 = 0

    '    For Y As Int32 = 0 To lTmp.GetUpperBound(0)
    '        If lValsAtLowest = 0 Then
    '            lLowest = Int32.MaxValue
    '            For X As Int32 = 0 To Values.GetUpperBound(0)
    '                If Values(X) < lLowest AndAlso Values(X) > lLastLowest Then
    '                    lLowest = Values(X)
    '                    lValsAtLowest = 1
    '                ElseIf Values(X) = lLowest Then
    '                    lValsAtLowest += 1
    '                End If
    '            Next X
    '        End If

    '        lTmp(Y) = lLowest
    '        lValsAtLowest -= 1
    '        lLastLowest = lLowest
    '    Next Y

    '    Dim lHalf As Int32 = CInt(Math.Floor(Values.Length / 2))
    '    Dim bNeedsAvg As Boolean = (Values.Length Mod 2) <> 0

    '    If bNeedsAvg = True Then
    '        Return (Values(lHalf) + Values(lHalf + 1)) \ 2
    '    Else : Return Values(lHalf)
    '    End If
    'End Function

    'Private Function GetKnownValsMedian() As Int32
    '    Dim H8 As Int32 = Owner.GetMineralPropertyKnowledge(lCoilMineralID, eMinPropID.SuperconductivePoint)
    '    Dim I8 As Int32 = Owner.GetMineralPropertyKnowledge(lCoilMineralID, eMinPropID.Density)
    '    Dim J8 As Int32 = Owner.GetMineralPropertyKnowledge(lCoilMineralID, eMinPropID.MagneticProduction)

    '    Dim H9 As Int32 = Owner.GetMineralPropertyKnowledge(lAcceleratorMineralID, eMinPropID.Quantum)
    '    Dim I9 As Int32 = Owner.GetMineralPropertyKnowledge(lAcceleratorMineralID, eMinPropID.SuperconductivePoint)
    '    Dim J9 As Int32 = Owner.GetMineralPropertyKnowledge(lAcceleratorMineralID, eMinPropID.MagneticReaction)

    '    Dim H10 As Int32 = Owner.GetMineralPropertyKnowledge(lCasingMineralID, eMinPropID.Density)
    '    Dim I10 As Int32 = Owner.GetMineralPropertyKnowledge(lCasingMineralID, eMinPropID.ThermalConductance)
    '    Dim J10 As Int32 = Owner.GetMineralPropertyKnowledge(lCasingMineralID, eMinPropID.TemperatureSensitivity)

    '    Return CalculateMedian(H8, I8, J8, H9, I9, J9, H10, I10, J10)
    'End Function

    Public Overrides Function SetFromDesignMsg(ByRef yData() As Byte) As Boolean
        Dim lPos As Int32 = 14  '2 for msg code, 2 for objtypeid, 4 for id, 6 for researcher guid
        Dim bRes As Boolean = False
        Try
            With Me
                .ObjTypeID = ObjectType.eShieldTech
                .MaxHitPoints = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .RechargeRate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .RechargeFreq = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lProjectionHullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lCoilMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lAcceleratorMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lCasingMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .ColorValue = yData(lPos) : lPos += 1

                ReDim .ShieldName(19)
                Array.Copy(yData, lPos, .ShieldName, 0, 20)
                lPos += 20

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
                .HullTypeID = yData(lPos) : lPos += 1
            End With

            bRes = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "EngineTech.SetFromDesignMsg: " & ex.Message)
        End Try

        Return bRes

    End Function

    Protected Overrides Sub CalculateBothProdCosts()
        If moResearchCost Is Nothing OrElse moProductionCost Is Nothing Then ComponentDesigned()

        'If moResearchCost Is Nothing OrElse moProductionCost Is Nothing Then
        '    If moResearchCost Is Nothing Then moResearchCost = New ProductionCost

        '    Dim f_E16 As Single = 0.0F
        '    Dim f_F16 As Single = 0.0F
        '    Dim f_G16 As Single = 0.0F
        '    Dim l_B17 As Int32 = 0I

        '    Dim bValid As Boolean = True

        '    With moResearchCost
        '        .ColonistCost = 0 : .EnlistedCost = 0 : .OfficerCost = 0
        '        .ObjectID = Me.ObjectID
        '        .ObjTypeID = Me.ObjTypeID

        '        Try

        '            l_B17 = GetKnownValsMedian() \ 2I

        '            Me.CurrentSuccessChance = l_B17

        '            'E14=IF(shldmaxHP/shldprojhullsize<1,1,POWER(shldmaxHP/shldprojhullsize,3))
        '            Dim f_E14 As Single
        '            If (Me.MaxHitPoints / Me.lProjectionHullSize) < 1 Then f_E14 = 1.0F Else f_E14 = CSng(Math.Pow((Me.MaxHitPoints / Me.lProjectionHullSize), 3))

        '            'E16=ABS((POWER(E8,2)-POWER(H8,1.5))*((10-recharge)-(I8/40)+(J8/60)))
        '            f_E16 = CSng(Math.Abs((Math.Pow(CoilMineral.GetPropertyValue(eMinPropID.SuperconductivePoint), 2) - _
        '              Math.Pow(Owner.GetMineralPropertyKnowledge(lCoilMineralID, eMinPropID.SuperconductivePoint), 1.4F)) * _
        '              ((10 - (Me.RechargeFreq / 30)) - (Owner.GetMineralPropertyKnowledge(lCoilMineralID, eMinPropID.Density) / 40) + _
        '              (Owner.GetMineralPropertyKnowledge(lCoilMineralID, eMinPropID.MagneticProduction) / 60))))

        '            'F16=POWER(ROUNDUP(160/(H9+1),0),2)*(1+H9)/((257-I9) - F9) * (J9/G9+1)
        '            Dim lAccelQuantumKnow As Int32 = Owner.GetMineralPropertyKnowledge(lAcceleratorMineralID, eMinPropID.Quantum)
        '            f_F16 = CSng(Math.Pow(Math.Ceiling(160 / (lAccelQuantumKnow + 1)), 2) * _
        '              (1 + lAccelQuantumKnow) / ((257 - Owner.GetMineralPropertyKnowledge(lAcceleratorMineralID, eMinPropID.SuperconductivePoint)) - _
        '                 AcceleratorMineral.GetPropertyValue(eMinPropID.SuperconductivePoint)) * _
        '                 (Owner.GetMineralPropertyKnowledge(lAcceleratorMineralID, eMinPropID.MagneticReaction) / _
        '                 (AcceleratorMineral.GetPropertyValue(eMinPropID.MagneticReaction) + 1)))

        '            'G16=((shldprojhullsize*shldmaxHP)*(ABS((shldprojhullsize/shldmaxHP*100) - MAX(E10,H10)) + 1) * _
        '            '   (shldprojhullsize/MAX(I10,F10))+((G10-J10) * ((shldmaxHP/10) -rechargerate)))/C16
        '            f_G16 = CSng(((lProjectionHullSize * MaxHitPoints) * (Math.Abs((lProjectionHullSize / MaxHitPoints * 100) - _
        '                Math.Max(CasingMineral.GetPropertyValue(eMinPropID.Density), _
        '                Owner.GetMineralPropertyKnowledge(lCasingMineralID, eMinPropID.Density))) + 1) * _
        '                (lProjectionHullSize / Math.Max(Math.Max(Owner.GetMineralPropertyKnowledge(lCasingMineralID, eMinPropID.ThermalConductance), _
        '                CasingMineral.GetPropertyValue(eMinPropID.ThermalConductance)), 1)) + ((CasingMineral.GetPropertyValue(eMinPropID.TemperatureSensitivity) - _
        '                Owner.GetMineralPropertyKnowledge(lCasingMineralID, eMinPropID.TemperatureSensitivity)) * _
        '                ((MaxHitPoints / 10.0F) - RechargeRate))) / Math.Pow(MaxHitPoints, 2))

        '            'PowerReq=MAX(E8,F9)*(E14+randomseed*10)*POWER(J5,2)+shldmaxHP/100
        '            Me.PowerRequired = CInt(Math.Ceiling((Math.Max(CoilMineral.GetPropertyValue(eMinPropID.SuperconductivePoint), _
        '                Math.Max(CoilMineral.GetPropertyValue(eMinPropID.Density), _
        '                Math.Max(AcceleratorMineral.GetPropertyValue(eMinPropID.Quantum), _
        '                AcceleratorMineral.GetPropertyValue(eMinPropID.SuperconductivePoint)))) * (f_E14 + RandomSeed * 10.0F) * _
        '                Math.Pow(((60 / (Me.RechargeFreq / 30) * Me.RechargeRate) / lProjectionHullSize), 2) + MaxHitPoints / 100.0F)))

        '            '=shldpowerreq*POWER(E14,2)+ E16 + F16 + G16
        '            Dim dTmp As Double = Math.Ceiling(Me.PowerRequired * Math.Pow(f_E14, 2) + f_E16 + f_F16 + f_G16)
        '            If dTmp > Int32.MaxValue Then
        '                .CreditCost = Int32.MaxValue
        '            Else : .CreditCost = CInt(dTmp)
        '            End If

        '            .PointsRequired = GetResearchPointsCost(.CreditCost * 10)
        '            .CreditCost *= Me.ResearchAttempts
        '            'TODO: Reset ResearchAttempts?

        '            'HullRequired=ROUNDDOWN(shldprojhullsize*(randomseed/10),0)
        '            Me.HullRequired = CInt(Math.Floor(lProjectionHullSize * (RandomSeed / 10)))

        '            .ItemCostUB = -1
        '            .AddProductionCostItem(lCoilMineralID, ObjectType.eMineral, CInt(Math.Ceiling(Me.HullRequired * 0.3F) + 1))
        '            .AddProductionCostItem(lAcceleratorMineralID, ObjectType.eMineral, CInt(Math.Ceiling(Me.HullRequired * 0.25F) + 1))
        '            .AddProductionCostItem(lCasingMineralID, ObjectType.eMineral, Me.HullRequired)
        '        Catch
        '            bValid = False
        '            Me.ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
        '        End Try
        '    End With

        '    '==== BEGIN PRODUCTION COST ====
        '    If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
        '    With moProductionCost
        '        .ObjectID = Me.ObjectID
        '        .ObjTypeID = Me.ObjTypeID
        '        .ColonistCost = 0
        '        .EnlistedCost = 0
        '        .OfficerCost = 0

        '        If bValid = True Then
        '            Try
        '                Dim l_C8 As Int32 = CInt(Math.Ceiling(Me.HullRequired * 0.3)) + 1I
        '                Dim l_C9 As Int32 = CInt(Math.Ceiling(Me.HullRequired * 0.25)) + 1I
        '                Dim f_C19 As Single = CSng(Math.Pow((60 / (RechargeFreq / 30) * RechargeRate) / lProjectionHullSize, 2))

        '                .CreditCost = CInt((f_G16 / 5.0F + f_F16 / l_C9 + f_E16 / l_C8) * (RandomSeed + 1) * f_C19)
        '                .PointsRequired = (100 - Me.CurrentSuccessChance) * 9000

        '                .ItemCostUB = -1
        '                .AddProductionCostItem(lCoilMineralID, ObjectType.eMineral, l_C8)
        '                .AddProductionCostItem(lAcceleratorMineralID, ObjectType.eMineral, l_C9)
        '                .AddProductionCostItem(lCasingMineralID, ObjectType.eMineral, Me.HullRequired)
        '            Catch
        '                bValid = False
        '                Me.ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
        '            End Try
        '        End If
        '    End With

        '    If bValid = False Then Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
        'End If
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
    '		Dim fRechFreqTime As Single = (RechargeFreq / 30.0F)

    '		'to keep track of the major design flaw
    '		Dim fHighestScore As Single = 0
    '		Dim bNotStudyFlaw As Boolean = False

    '		'======== Coil ===========
    '		Dim uCoil_Density As MaterialPropertyItem2
    '		With uCoil_Density
    '			.lMineralID = CoilMineral.ObjectID
    '			.lPropertyID = eMinPropID.Density

    '			.lActualValue = CoilMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.3F
    '			.lGoalValue = 136 '200
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
    '			.lMineralID = CoilMineral.ObjectID
    '			.lPropertyID = eMinPropID.SuperconductivePoint

    '			.lActualValue = CoilMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.6F
    '			.lGoalValue = 20
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
    '			.lMineralID = CoilMineral.ObjectID
    '			.lPropertyID = eMinPropID.MagneticProduction

    '			.lActualValue = CoilMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 158 '200
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
    '		Dim lCoilComplexity As Int32 = CoilMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fCoilInc As Single = (uCoil_Density.FinalScore + uCoil_MagProd.FinalScore + uCoil_SuperC.FinalScore)
    '		fCoilInc *= CSng(lCoilComplexity / PopIntel)
    '		'================== END OF COIL ======================

    '		'================== ACCELERATOR ======================
    '		Dim uAcc_Quantum As MaterialPropertyItem2
    '		With uAcc_Quantum
    '			.lMineralID = AcceleratorMineral.ObjectID
    '			.lPropertyID = eMinPropID.Quantum

    '			.lActualValue = AcceleratorMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.3F
    '			.lGoalValue = 121 '200
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
    '		Dim uAcc_SuperC As MaterialPropertyItem2
    '		With uAcc_SuperC
    '			.lMineralID = AcceleratorMineral.ObjectID
    '			.lPropertyID = eMinPropID.SuperconductivePoint

    '			.lActualValue = AcceleratorMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.6F
    '			.lGoalValue = 20
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
    '		Dim uAcc_MagReact As MaterialPropertyItem2
    '		With uAcc_MagReact
    '			.lMineralID = AcceleratorMineral.ObjectID
    '			.lPropertyID = eMinPropID.MagneticReaction

    '			.lActualValue = AcceleratorMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 156 '200
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
    '		Dim lAccComplexity As Int32 = AcceleratorMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fAccInc As Single = uAcc_MagReact.FinalScore + uAcc_Quantum.FinalScore + uAcc_SuperC.FinalScore
    '		fAccInc *= CSng(lAccComplexity / PopIntel)
    '		'================== END OF ACCELERATOR ==============

    '		'================== CASING ==========================
    '		Dim uCase_Density As MaterialPropertyItem2
    '		With uCase_Density
    '			.lMineralID = CasingMineral.ObjectID
    '			.lPropertyID = eMinPropID.Density

    '			.lActualValue = CasingMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.3F
    '			.lGoalValue = 143 '200
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
    '		Dim uCase_ThermCond As MaterialPropertyItem2
    '		With uCase_ThermCond
    '			.lMineralID = CasingMineral.ObjectID
    '			.lPropertyID = eMinPropID.ThermalConductance

    '			.lActualValue = CasingMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.6F
    '			.lGoalValue = 112 '200
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
    '		Dim uCase_TempSens As MaterialPropertyItem2
    '		With uCase_TempSens
    '			.lMineralID = CasingMineral.ObjectID
    '			.lPropertyID = eMinPropID.TemperatureSensitivity

    '			.lActualValue = CasingMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 20
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
    '		Dim lCasingComplexity As Int32 = CasingMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fCasingInc As Single = uCase_Density.FinalScore + uCase_TempSens.FinalScore + uCase_ThermCond.FinalScore
    '		fCasingInc *= CSng(lCasingComplexity / PopIntel)
    '		'================== END OF CASING ===================

    '		If fHighestScore < 10.0F AndAlso bNotStudyFlaw = False Then
    '			MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eGoodDesign
    '		End If

    '		'ok, here we go, these values are at the top of the spreadsheet...
    '		Dim fHPPerSec As Single = Me.RechargeRate / fRechFreqTime
    '		Dim fFullRecharge As Single = Me.MaxHitPoints / fHPPerSec
    '		'=IF(maxHP>0.2*projhullsize,maxHP-(projhullsize*0.2),maxHP/projhullsize)*100
    '		Dim fOverage As Single
    '		If MaxHitPoints > 0.2F * Me.lProjectionHullSize Then
    '			fOverage = MaxHitPoints - (lProjectionHullSize * 0.2F)
    '		Else : fOverage = CSng(MaxHitPoints / lProjectionHullSize)
    '		End If
    '		fOverage *= 100.0F
    '		Dim fDPSForHull As Single = 1.0F + (fHPPerSec / HullToDPS(Me.lProjectionHullSize))
    '		'end of the top stuff

    '		'bottom stuff...
    '		Dim fModifiedSeed As Single = 0.8F + (RandomSeed * 0.7F)
    '		Dim lMaxPowerBasedOnHull As Int32 = MaxPowerLU(Me.lProjectionHullSize)
    '		Dim fHPToHullPercent As Single = CSng(Me.MaxHitPoints) / CSng(Me.lProjectionHullSize)
    '		Dim fResearchDifficulty As Single = fAccInc * fCasingInc * fCoilInc
    '		Dim fAverageComplexity As Single = (lCasingComplexity + lCoilComplexity + lAccComplexity) / 3.0F
    '		Dim fAverageAK As Single = uCoil_Density.AKScore + uCoil_MagProd.AKScore + uCoil_SuperC.AKScore
    '		fAverageAK += uAcc_MagReact.AKScore + uAcc_Quantum.AKScore + uAcc_SuperC.AKScore
    '		fAverageAK += uCase_Density.AKScore + uCase_TempSens.AKScore + uCase_ThermCond.AKScore
    '		fAverageAK /= 9.0F

    '		Me.CurrentSuccessChance = CInt(fAccInc + fCasingInc + fCoilInc) * -1

    '		'=(projhullsize/C32*randomseed*B25)/fullcharge+POWER(G7,2)
    '		Dim dblPower As Double = (lProjectionHullSize / CheckHullSize(lMaxPowerBasedOnHull) * fModifiedSeed * lMaxPowerBasedOnHull) / fFullRecharge + Math.Pow(fOverage, 2)
    '		'=IF(H8>1,POWER(D32,H8),D32)
    '		If fDPSForHull > 1 Then
    '			dblPower = Math.Pow(dblPower, fDPSForHull)
    '		End If
    '		dblPower += fOverage
    '		dblPower /= 20.0F
    '		If dblPower > Int32.MaxValue Then Me.PowerRequired = Int32.MaxValue Else Me.PowerRequired = CInt(dblPower)

    '		'Hull required...
    '		Me.HullRequired = CInt(Math.Ceiling((fAccInc + fCasingInc + fCoilInc) * (fAverageComplexity / 50.0F) * fOverage)) \ 10
    '		If Me.lProjectionHullSize <= 670 Then Me.HullRequired \= 3
    '		Me.HullRequired = Math.Max(Me.lProjectionHullSize \ 10, Me.HullRequired)

    '		Dim lEnlisted As Int32 = CInt(Math.Ceiling((fAccInc + fCasingInc + fCoilInc) / 30.0F))
    '		Dim lOfficers As Int32 = lEnlisted \ 5

    '		'Now, for research credits =ROUNDUP((B30*B26*B29*G7/B32*(1+(G7/20))),0)*D7*1000
    '		Dim dblTemp As Double = fAverageComplexity * fHPToHullPercent * fResearchDifficulty * fOverage / dblPower * (1 + (fOverage / 20.0F))
    '		Dim blResearchCost As Int64 = CLng(Math.Ceiling(dblTemp) * fHPPerSec * 1000.0F)

    '		'ResearchPoints  =B35*randomseed*100*E24
    '		Dim blResearchPoints As Int64 = CLng(blResearchCost * fModifiedSeed * 100 * fAverageAK)
    '		blResearchPoints = Math.Max(blResearchPoints, CLng(360000 * fModifiedSeed))

    '		'Production Credits =ROUNDUP(projhullsize*B25/B32*B33,0)
    '		Dim blProductionCost As Int64 = CLng(Math.Ceiling(CDbl(lProjectionHullSize) * CDbl(lMaxPowerBasedOnHull) / dblPower * CDbl(HullRequired)))

    '		'Production Points =(B33+1000000)+B37*randomseed*D7/maxHP
    '		Dim blProductionPoints As Int64 = CLng((HullRequired + 1000000L) + blProductionCost * fModifiedSeed * fHPPerSec / CDbl(MaxHitPoints))
    '		blProductionPoints = Math.Max(blProductionPoints, CLng(1800000 * fModifiedSeed))

    '		Dim lCoilCost As Int32 = CInt(Math.Abs(Math.Ceiling(fCoilInc * HullRequired) * CLng(CurrentSuccessChance)))
    '		Dim lAcceleratorCost As Int32 = CInt(Math.Abs(Math.Ceiling(fAccInc * HullRequired) * 0.2F))
    '		Dim lCasingCost As Int32 = CInt(Math.Abs(Math.Ceiling(fCasingInc * HullRequired) * CLng(CurrentSuccessChance)))

    '		lCoilCost = Math.Max(1, lCoilCost \ 10)
    '		lAcceleratorCost = Math.Max(1, lAcceleratorCost \ 10)
    '		lCasingCost = Math.Max(1, lCasingCost \ 10)

    '		If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
    '		With moResearchCost
    '			.ObjectID = Me.ObjectID
    '			.ObjTypeID = Me.ObjTypeID
    '			.ColonistCost = 0 : .EnlistedCost = 0 : .OfficerCost = 0
    '			.CreditCost = blResearchCost
    '			Erase .ItemCosts
    '			.ItemCostUB = -1
    '			.PointsRequired = GetResearchPointsCost(blResearchPoints)
    '			.CreditCost *= Me.ResearchAttempts

    '			.AddProductionCostItem(lCoilMineralID, ObjectType.eMineral, lCoilCost)
    '			.AddProductionCostItem(lAcceleratorMineralID, ObjectType.eMineral, lAcceleratorCost)
    '			.AddProductionCostItem(lCasingMineralID, ObjectType.eMineral, lCasingCost)
    '		End With
    '		If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
    '		With moProductionCost
    '			.ObjectID = Me.ObjectID
    '			.ObjTypeID = Me.ObjTypeID
    '			.ColonistCost = 0 : .EnlistedCost = lEnlisted : .OfficerCost = lOfficers
    '			.CreditCost = blProductionCost
    '			Erase .ItemCosts
    '			.ItemCostUB = -1
    '			.PointsRequired = blProductionPoints

    '			.AddProductionCostItem(lCoilMineralID, ObjectType.eMineral, CInt(Math.Ceiling(Me.HullRequired * 0.03F)))
    '			.AddProductionCostItem(lAcceleratorMineralID, ObjectType.eMineral, CInt(Math.Ceiling(Me.HullRequired * 0.07F)))
    '			.AddProductionCostItem(lCasingMineralID, ObjectType.eMineral, Math.Max(1, Me.HullRequired \ 10))
    '		End With

    '	Catch
    '		Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
    '		Me.ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
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

            Dim oComputer As New ShieldTechComputer
            With oComputer
                Dim decRInt As Decimal = Math.Max(1, RechargeFreq) / 30D
                .decHPS = RechargeRate / decRInt
                .blMaxHP = MaxHitPoints
                .blProjHull = lProjectionHullSize

                .lHullTypeID = Me.HullTypeID

                .lMineral1ID = lCoilMineralID 'lAcceleratorMineralID
                .lMineral2ID = lAcceleratorMineralID 'lCasingMineralID
                .lMineral3ID = lCasingMineralID 'lCoilMineralID
                .lMineral4ID = -1
                .lMineral5ID = -1
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

                If .BuilderCostValueChange(Me.ObjTypeID, 0D, PopIntel, Me.Owner.oSpecials.HullPerResidence) = False Then
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

    Private Function MaxPowerLU(ByVal lValue As Int32) As Int32
        If lValue < 120 Then
            Return 250
        ElseIf lValue < 170 Then
            Return 303
        ElseIf lValue < 300 Then
            Return 347
        ElseIf lValue < 500 Then
            Return 1042
        ElseIf lValue < 2000 Then
            Return 536
        ElseIf lValue < 5000 Then
            Return 1389
        ElseIf lValue < 10000 Then
            Return 1515
        ElseIf lValue < 200000 Then
            Return 3125
        ElseIf lValue < 400000 Then
            Return 6250
        ElseIf lValue < 600000 Then
            Return 12500
        ElseIf lValue < 1000000 Then
            Return 46900
        Else : Return 375000
        End If
    End Function
    Private Function HullToDPS(ByVal lValue As Int32) As Int32
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
    Private Function CheckHullSize(ByVal lValue As Int32) As Int32
        If lValue < 250 Then
            Return 90
        ElseIf lValue < 303 Then
            Return 120
        ElseIf lValue < 347 Then
            Return 170
        ElseIf lValue < 536 Then
            Return 300
        ElseIf lValue < 1042 Then
            Return 2000
        ElseIf lValue < 1389 Then
            Return 500
        ElseIf lValue < 1515 Then
            Return 5000
        ElseIf lValue < 3125 Then
            Return 10000
        ElseIf lValue < 6250 Then
            Return 200000
        ElseIf lValue < 12500 Then
            Return 400000
        ElseIf lValue < 46900 Then
            Return 60000
        ElseIf lValue < 375000 Then
            Return 1000000
        Else : Return 4000000
        End If
    End Function


    Protected Overrides Sub FinalizeResearch()
        Dim fMult As Single = 1.0F + (Owner.oSpecials.yShieldMaxHPBonus / 100.0F)
        MaxHitPoints = CInt(Math.Ceiling(MaxHitPoints * fMult))

        fMult = 1.0F + (Owner.oSpecials.yShieldProjHullSizeBonus / 100.0F)
        lProjectionHullSize = CInt(Math.Ceiling(lProjectionHullSize * fMult))

        fMult = 1.0F - (Owner.oSpecials.yShieldRechargeFreqBonus / 100.0F)
        RechargeFreq = CShort(RechargeFreq * fMult)

        fMult = 1.0F + (Owner.oSpecials.yShieldRechargeRateBonus / 100.0F)
        RechargeRate = CInt(Math.Ceiling(RechargeRate * fMult))
    End Sub

    Public Function GetMaxProperty(ByVal lPropID As Int32) As Int32
        Return GetMaxVal(CoilMineral.GetPropertyValue(lPropID), AcceleratorMineral.GetPropertyValue(lPropID), CasingMineral.GetPropertyValue(lPropID))
    End Function

    Public Overrides Function GetNonOwnerMsg() As Byte()
        Dim yResult(52) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yResult, lPos) : lPos += 2
        Me.GetGUIDAsString.CopyTo(yResult, lPos) : lPos += 6

        ShieldName.CopyTo(yResult, lPos) : lPos += 20
        System.BitConverter.GetBytes(MaxHitPoints).CopyTo(yResult, lPos) : lPos += 4
        System.BitConverter.GetBytes(RechargeRate).CopyTo(yResult, lPos) : lPos += 4
        System.BitConverter.GetBytes(RechargeFreq).CopyTo(yResult, lPos) : lPos += 4
        System.BitConverter.GetBytes(lProjectionHullSize).CopyTo(yResult, lPos) : lPos += 4
        System.BitConverter.GetBytes(PowerRequired).CopyTo(yResult, lPos) : lPos += 4
        System.BitConverter.GetBytes(HullRequired).CopyTo(yResult, lPos) : lPos += 4
        yResult(lPos) = Me.HullTypeID : lPos += 1

        Return yResult
    End Function

    Public Overrides Function TechnologyScore() As Integer
        If mlStoredTechScore = Int32.MinValue Then
            Try
                Dim fRF As Single = Me.RechargeFreq / 30.0F
                mlStoredTechScore = CInt(((MaxHitPoints * (RechargeRate / fRF)) * Me.lProjectionHullSize) / (Me.PowerRequired + Me.HullRequired)) \ 10
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
        lPos = 30
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eNameOnly Then lPos += 10
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then lPos += 28
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then lPos += 88 '16
        ReDim yMsg(lPos - 1)
        lPos = 0

        'Default Attributes
        GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        Me.ShieldName.CopyTo(yMsg, lPos) : lPos += 20

        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
            'at least settings 1 
            System.BitConverter.GetBytes(PowerRequired).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(HullRequired).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = ColorValue : lPos += 1
            yMsg(lPos) = Me.HullTypeID : lPos += 1
        End If
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
            'at least settings 2  
            System.BitConverter.GetBytes(MaxHitPoints).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(RechargeRate).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(RechargeFreq).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lProjectionHullSize).CopyTo(yMsg, lPos) : lPos += 4
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
            System.BitConverter.GetBytes(lCasingMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lCoilMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lAcceleratorMineralID).CopyTo(yMsg, lPos) : lPos += 4

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

            '======= end of header =========
            ReDim .ShieldName(19)
            Array.Copy(yData, lPos, .ShieldName, 0, 20) : lPos += 20
            .MaxHitPoints = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .RechargeRate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .RechargeFreq = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lProjectionHullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lCasingMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lCoilMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lAcceleratorMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .PowerRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .HullRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ColorValue = yData(lPos) : lPos += 1
            .HullTypeID = yData(lPos) : lPos += 1

            If .Owner Is Nothing = False Then
                Dim lFirstIdx As Int32 = -1
                For X As Int32 = 0 To .Owner.mlShieldUB
                    If .Owner.mlShieldIdx(X) = .ObjectID Then
                        .Owner.moShield(X) = Me
                        Return
                    ElseIf lFirstIdx = -1 AndAlso .Owner.mlShieldIdx(X) = -1 Then
                        lFirstIdx = X
                    End If
                Next X

                If lFirstIdx = -1 Then
                    lFirstIdx = .Owner.mlShieldUB + 1
                    ReDim Preserve .Owner.mlShieldIdx(lFirstIdx)
                    ReDim Preserve .Owner.moShield(lFirstIdx)
                    .Owner.mlShieldUB = lFirstIdx
                End If
                .Owner.moShield(lFirstIdx) = Me
                .Owner.mlShieldIdx(lFirstIdx) = Me.ObjectID
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
        Dim oComputer As New ShieldTechComputer
        With oComputer
            Dim decRInt As Decimal = Math.Max(1, RechargeFreq) / 30D
            .decHPS = RechargeRate / decRInt
            .blMaxHP = MaxHitPoints
            .blProjHull = lProjectionHullSize

            .lHullTypeID = Me.HullTypeID

            .lMineral1ID = lCoilMineralID 'lAcceleratorMineralID
            .lMineral2ID = lAcceleratorMineralID 'lCasingMineralID
            .lMineral3ID = lCasingMineralID 'lCoilMineralID
            .lMineral4ID = -1
            .lMineral5ID = -1
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

            If .BuilderCostValueChange(Me.ObjTypeID, 0D, PopIntel, Me.Owner.oSpecials.HullPerResidence) = False Then
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
