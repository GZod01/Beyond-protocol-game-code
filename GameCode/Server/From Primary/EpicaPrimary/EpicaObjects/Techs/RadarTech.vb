Public Class RadarTech
    Inherits Epica_Tech

    Public RadarName(19) As Byte
    Public WeaponAcc As Byte
    Public ScanResolution As Byte
    Public OptimumRange As Byte
    Public MaximumRange As Byte
    Public DisruptionResistance As Byte

    Public RadarType As Byte
    Public JamImmunity As Byte
    Public JamStrength As Byte
    Public JamTargets As Byte
    Public JamEffect As Byte

    Public lEmitterMineralID As Int32
    Public lDetectionMineralID As Int32
    Public lCollectionMineralID As Int32
    Public lCasingMineralID As Int32

    Public PowerRequired As Int32
    Public HullRequired As Int32

    Private mySendString() As Byte

    Public Function GetObjAsString() As Byte()
        If moResearchCost Is Nothing OrElse moProductionCost Is Nothing Then
            CalculateBothProdCosts()
            mbStringReady = False
        End If

        'here we will return the entire object as a string
        'If mbStringReady = False Then
        ReDim mySendString(Epica_Tech.BASE_OBJ_STRING_SIZE + 53)

        Dim lPos As Int32
        MyBase.GetBaseObjAsString(False).CopyTo(mySendString, 0)
        lPos = Epica_Tech.BASE_OBJ_STRING_SIZE

        RadarName.CopyTo(mySendString, lPos) : lPos += 20
        mySendString(lPos) = WeaponAcc : lPos += 1
        mySendString(lPos) = ScanResolution : lPos += 1
        mySendString(lPos) = OptimumRange : lPos += 1
        mySendString(lPos) = MaximumRange : lPos += 1
        mySendString(lPos) = DisruptionResistance : lPos += 1

        System.BitConverter.GetBytes(lCollectionMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lCasingMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lEmitterMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lDetectionMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(PowerRequired).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(HullRequired).CopyTo(mySendString, lPos) : lPos += 4

        mySendString(lPos) = RadarType : lPos += 1
        mySendString(lPos) = JamImmunity : lPos += 1
        mySendString(lPos) = JamStrength : lPos += 1
        mySendString(lPos) = JamTargets : lPos += 1
        mySendString(lPos) = JamEffect : lPos += 1

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
                sSQL = "INSERT INTO tblRadar (OwnerID, RadarName, WeaponAcc, ScanResolution, OptimumRange, " & _
                  "MaximumRange, DisruptionResist, CollectionMineralID, CasingMineralID, EmitterMineralID, " & _
                  "DetectionMineralID, PowerRequired, HullRequired" & _
                  ", CurrentSuccessChance, SuccessChanceIncrement, ResearchAttempts, RandomSeed, ResPhase, ErrorReasonCode" & _
                  ", RadarType, JamImmunity, JamStrength, JamTargets, JamEffect, MajorDesignFlaw, PopIntel, bArchived) VALUES (" & _
                  Owner.ObjectID & ", '" & MakeDBStr(BytesToString(RadarName)) & "', " & WeaponAcc & ", " & ScanResolution & ", " & _
                  OptimumRange & ", " & MaximumRange & ", " & DisruptionResistance & ", " & lCollectionMineralID & _
                  ", " & lCasingMineralID & ", " & lEmitterMineralID & ", " & lDetectionMineralID & ", " & _
                  PowerRequired & ", " & HullRequired & _
                  ", " & CurrentSuccessChance & ", " & SuccessChanceIncrement & ", " & ResearchAttempts & _
                  ", " & RandomSeed & ", " & ComponentDevelopmentPhase & ", " & Me.ErrorReasonCode & _
                  ", " & RadarType & ", " & JamImmunity & ", " & JamStrength & ", " & JamTargets & ", " & JamEffect & _
                  ", " & MajorDesignFlaw & ", " & PopIntel & ", " & yArchived & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblRadar SET OwnerID=" & Owner.ObjectID & ", RadarName='" & MakeDBStr(BytesToString(RadarName)) & "', WeaponAcc=" & _
                  WeaponAcc & ", ScanResolution=" & ScanResolution & ", OptimumRange = " & OptimumRange & _
                  ", MaximumRange = " & MaximumRange & ", DisruptionResist = " & DisruptionResistance & _
                  ",CollectionMineralID = " & lCollectionMineralID & ", CasingMineralID = " & _
                  lCasingMineralID & ", EmitterMineralID = " & lEmitterMineralID & ", DetectionMineralID = " & _
                  lDetectionMineralID & ", PowerRequired=" & PowerRequired & ", HullRequired = " & HullRequired & _
                  ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
                  SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & _
                  RandomSeed & ", ResPhase = " & ComponentDevelopmentPhase & ", ErrorReasonCode = " & Me.ErrorReasonCode & _
                  ", RadarType = " & RadarType & ", JamImmunity = " & JamImmunity & ", JamStrength = " & JamStrength & _
                  ", JamTargets = " & JamTargets & ", JamEffect = " & JamEffect & ", MajorDesignFlaw = " & _
                  MajorDesignFlaw & ", PopIntel = " & PopIntel & ", bArchived = " & yArchived & " WHERE RadarID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(RadarID) FROM tblRadar WHERE RadarName = '" & MakeDBStr(BytesToString(RadarName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing

                ''Ok, set our object in the master array
                'Dim bAdded As Boolean = False
                'For X As Int32 = 0 To glRadarUB
                '    If glRadarIdx(X) = -1 Then
                '        glRadarIdx(X) = ObjectID
                '        goRadar(X) = Me
                '        bAdded = True
                '        Exit For
                '    End If
                'Next X
                'If bAdded = True Then
                '    glRadarUB += 1
                '    ReDim Preserve glRadarIdx(glRadarUB)
                '    ReDim Preserve goRadar(glRadarUB)
                '    glRadarIdx(glRadarUB) = ObjectID
                '    goRadar(glRadarUB) = Me
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

            'UPDATE
            sSQL = "UPDATE tblRadar SET OwnerID=" & Owner.ObjectID & ", RadarName='" & MakeDBStr(BytesToString(RadarName)) & "', WeaponAcc=" & _
              WeaponAcc & ", ScanResolution=" & ScanResolution & ", OptimumRange = " & OptimumRange & _
              ", MaximumRange = " & MaximumRange & ", DisruptionResist = " & DisruptionResistance & _
              ",CollectionMineralID = " & lCollectionMineralID & ", CasingMineralID = " & _
              lCasingMineralID & ", EmitterMineralID = " & lEmitterMineralID & ", DetectionMineralID = " & _
              lDetectionMineralID & ", PowerRequired=" & PowerRequired & ", HullRequired = " & HullRequired & _
              ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
              SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & _
              RandomSeed & ", ResPhase = " & ComponentDevelopmentPhase & ", ErrorReasonCode = " & Me.ErrorReasonCode & _
              ", RadarType = " & RadarType & ", JamImmunity = " & JamImmunity & ", JamStrength = " & JamStrength & _
              ", JamTargets = " & JamTargets & ", JamEffect = " & JamEffect & ", MajorDesignFlaw = " & _
              MajorDesignFlaw & ", PopIntel = " & PopIntel & ", bArchived = " & yArchived & " WHERE RadarID = " & ObjectID
            oSB.AppendLine(sSQL)

            oSB.AppendLine(MyBase.GetFinalizeSaveText())
            Return oSB.ToString
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        End Try
        Return ""
    End Function

    Public Overrides Function ValidateDesign() As Boolean
        'Ensure that the player has playermineral objects of the minerals used
        Dim bMin1Good As Boolean = (lEmitterMineralID > 0) AndAlso Owner.IsMineralDiscovered(lEmitterMineralID)
        Dim bMin2Good As Boolean = (lDetectionMineralID > 0) AndAlso Owner.IsMineralDiscovered(lDetectionMineralID)
        Dim bMin3Good As Boolean = (lCollectionMineralID > 0) AndAlso Owner.IsMineralDiscovered(lCollectionMineralID)
        Dim bMin4Good As Boolean = (lCasingMineralID > 0) AndAlso Owner.IsMineralDiscovered(lCasingMineralID)

        If bMin1Good = False OrElse bMin2Good = False OrElse bMin3Good = False OrElse bMin4Good = False Then
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidMaterialsSelected
            Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
            LogEvent(LogEventType.PossibleCheat, "RadarTech.ValidateDesign invalid minerals: " & Me.Owner.ObjectID)
            Return False
        End If

        If OptimumRange > Owner.oSpecials.yRadarOptRange OrElse MaximumRange > Owner.oSpecials.yRadarMaxRange OrElse _
          ScanResolution > Owner.oSpecials.yRadarScanRes OrElse WeaponAcc > Owner.oSpecials.yRadarWpnAcc OrElse _
          DisruptionResistance > Owner.oSpecials.yRadarDisRes OrElse JamStrength > Owner.oSpecials.yJamStrength Then
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
            Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
            LogEvent(LogEventType.PossibleCheat, "RadarTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If

        If JamEffect <> 0 AndAlso JamStrength <> 0 Then ' AndAlso RadarType <> 0 Then
            Dim lValue As Int32 = Owner.oSpecials.yJamTargets

            If JamTargets = 128 Then
                If (lValue And 128) = 0 Then
                    Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
                    Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
                    LogEvent(LogEventType.PossibleCheat, "RadarTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                    Return False
                End If
            Else
                If (lValue And 128) <> 0 Then lValue -= 128
                If JamTargets > lValue Then
                    Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
                    Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
                    LogEvent(LogEventType.PossibleCheat, "RadarTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                    Return False
                End If
            End If

            If JamEffect > Owner.oSpecials.yJamEffectAvailable Then
                Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
                Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
                LogEvent(LogEventType.PossibleCheat, "RadarTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
            If JamImmunity > Owner.oSpecials.yJamImmunityAvailable Then
                Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
                Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
                LogEvent(LogEventType.PossibleCheat, "RadarTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
        End If

        If CheckDesignImpossibility() = True Then
            LogEvent(LogEventType.PossibleCheat, "RadarTech.ValidateDesign CheckDesignPossiblity Failed: " & Me.Owner.ObjectID)
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
            Return False
        End If

        Return bMin1Good AndAlso bMin2Good AndAlso bMin3Good AndAlso bMin4Good
    End Function

    Private Function CheckDesignImpossibility() As Boolean
        Dim oTB As New RadarTechComputer
        With Me
            oTB.blLockedProdCost = .blSpecifiedProdCost
            oTB.blLockedProdTime = .blSpecifiedProdTime
            oTB.blLockedResCost = .blSpecifiedResCost
            oTB.blLockedResTime = .blSpecifiedResTime

            oTB.lHullTypeID = .RadarType
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
            oTB.lMineral1ID = .lEmitterMineralID
            oTB.lMineral2ID = .lDetectionMineralID
            oTB.lMineral3ID = .lCollectionMineralID
            oTB.lMineral4ID = .lCasingMineralID
            oTB.lMineral5ID = -1
            oTB.lMineral6ID = -1

            oTB.blDisRes = .DisruptionResistance
            oTB.blJamStrength = .JamStrength
            'Bogus: Do not back out specials.  These bonus's are applied after the design is finalized.
            'oTB.blMaxRng = Math.Max(0, CInt(.MaximumRange) - CInt(Owner.oSpecials.yRadarMaxRangeBonus))
            'oTB.blOptRng = Math.Max(0, CInt(.OptimumRange) - CInt(Owner.oSpecials.yRadarOptRangeBonus))
            'oTB.blScanRes = Math.Max(0, CInt(.ScanResolution) - CInt(Owner.oSpecials.yRadarScanResBonus))
            'oTB.blWepAcc = Math.Max(0, CInt(.WeaponAcc) - CInt(Owner.oSpecials.yRadarWpnAccBonus))
            oTB.blMaxRng = .MaximumRange
            oTB.blOptRng = .OptimumRange
            oTB.blScanRes = .ScanResolution
            oTB.blWepAcc = .WeaponAcc

            oTB.blJamTargets = .JamTargets
            If .JamTargets = 128 Then oTB.blJamTargets = 20
            oTB.lJamType = .JamEffect
        End With
        Return oTB.IsDesignImpossible(Me.ObjTypeID, Owner)
    End Function


    Private moEmitterMaterial As Mineral = Nothing
    Public ReadOnly Property EmitterMaterial() As Mineral
        Get
            If moEmitterMaterial Is Nothing OrElse moEmitterMaterial.ObjectID <> lEmitterMineralID Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = lEmitterMineralID Then
                        moEmitterMaterial = goMineral(X)
                        Exit For
                    End If
                Next X
            End If
            Return moEmitterMaterial
        End Get
    End Property

    Private moDetectionMaterial As Mineral = Nothing
    Public ReadOnly Property DetectionMaterial() As Mineral
        Get
            If moDetectionMaterial Is Nothing OrElse moDetectionMaterial.ObjectID <> lDetectionMineralID Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = lDetectionMineralID Then
                        moDetectionMaterial = goMineral(X)
                        Exit For
                    End If
                Next X
            End If
            Return moDetectionMaterial
        End Get
    End Property

    Private moCollectionMaterial As Mineral = Nothing
    Public ReadOnly Property CollectionMaterial() As Mineral
        Get
            If moCollectionMaterial Is Nothing OrElse moCollectionMaterial.ObjectID <> lCollectionMineralID Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = lCollectionMineralID Then
                        moCollectionMaterial = goMineral(X)
                        Exit For
                    End If
                Next X
            End If
            Return moCollectionMaterial
        End Get
    End Property

    Private moCasingMaterial As Mineral = Nothing
    Public ReadOnly Property CasingMaterial() As Mineral
        Get
            If moCasingMaterial Is Nothing OrElse moCasingMaterial.ObjectID <> lCasingMineralID Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = lCasingMineralID Then
                        moCasingMaterial = goMineral(X)
                        Exit For
                    End If
                Next X
            End If
            Return moCasingMaterial
        End Get
    End Property

    Public Overrides Function SetFromDesignMsg(ByRef yData() As Byte) As Boolean
        Dim lPos As Int32 = 14  '2 for msg code, 2 for objtypeid, 4 for id, 6 for researcher guid
        Dim bRes As Boolean = False
        Try
            With Me
                .ObjTypeID = ObjectType.eRadarTech

                .WeaponAcc = yData(lPos) : lPos += 1
                .ScanResolution = yData(lPos) : lPos += 1
                .OptimumRange = yData(lPos) : lPos += 1
                .MaximumRange = yData(lPos) : lPos += 1
                .DisruptionResistance = yData(lPos) : lPos += 1
                .lEmitterMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lDetectionMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lCollectionMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lCasingMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                .RadarType = yData(lPos) : lPos += 1
                .JamImmunity = yData(lPos) : lPos += 1
                .JamStrength = yData(lPos) : lPos += 1
                .JamTargets = yData(lPos) : lPos += 1
                .JamEffect = yData(lPos) : lPos += 1

                ReDim .RadarName(19)
                Array.Copy(yData, lPos, .RadarName, 0, 20)
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
            End With
            bRes = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "RadarTech.SetFromDesignMsg: " & ex.Message)
        End Try

        Return bRes
    End Function

    ' Protected Overrides Sub CalculateBothProdCosts()
    '     If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
    '     With moResearchCost
    '         .ObjectID = Me.ObjectID
    '         .ObjTypeID = Me.ObjTypeID
    '         Erase .ItemCosts
    '         .ItemCostUB = -1
    '     End With
    '     If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
    '     With moProductionCost
    '         .ObjectID = Me.ObjectID
    '         .ObjTypeID = Me.ObjTypeID
    '         Erase .ItemCosts
    '         .ItemCostUB = -1
    '     End With

    '     Try

    '         Dim lEmitterComplexity As Int32 = EmitterMaterial.GetPropertyValue(eMinPropID.Complexity)
    '         Dim lDetectorComplexity As Int32 = DetectionMaterial.GetPropertyValue(eMinPropID.Complexity)
    '         Dim lCollectorComplexity As Int32 = CollectionMaterial.GetPropertyValue(eMinPropID.Complexity)
    '         Dim lCasingComplexity As Int32 = CasingMaterial.GetPropertyValue(eMinPropID.Complexity)

    '         Dim lE_ElectRes_Actual As Int32 = EmitterMaterial.GetPropertyValue(eMinPropID.ElectricalResist)
    '         Dim lE_ElectRes_Known As Int32 = Owner.GetMineralPropertyKnowledge(lEmitterMineralID, eMinPropID.ElectricalResist)
    '         Dim lE_MagProd_Actual As Int32 = EmitterMaterial.GetPropertyValue(eMinPropID.MagneticProduction)
    '         Dim lE_MagProd_Known As Int32 = Owner.GetMineralPropertyKnowledge(lEmitterMineralID, eMinPropID.MagneticProduction)
    '         Dim lE_MagReact_Actual As Int32 = EmitterMaterial.GetPropertyValue(eMinPropID.MagneticReaction)
    '         Dim lE_MagReact_Known As Int32 = Owner.GetMineralPropertyKnowledge(lEmitterMineralID, eMinPropID.MagneticReaction)

    '         Dim fE_ElectRes_Score As Single = Math.Max(GetLowLookup(lE_ElectRes_Actual \ 10, 4), 0.1F)
    '         Dim fE_MagProd_Score As Single = Math.Max(GetHighLookup(lE_MagProd_Actual \ 10, 4), 0.1F)
    '         Dim fE_MagReact_Score As Single = Math.Max(GetHighLookup(lE_MagReact_Actual \ 10, 4), 0.1F)

    '         Dim lD_MagReact_Actual As Int32 = DetectionMaterial.GetPropertyValue(eMinPropID.MagneticReaction)
    '         Dim lD_MagReact_Known As Int32 = Owner.GetMineralPropertyKnowledge(lDetectionMineralID, eMinPropID.MagneticReaction)
    '         Dim lD_MagProd_Actual As Int32 = DetectionMaterial.GetPropertyValue(eMinPropID.MagneticProduction)
    '         Dim lD_MagProd_Known As Int32 = Owner.GetMineralPropertyKnowledge(lDetectionMineralID, eMinPropID.MagneticProduction)
    '         Dim lD_SuperC_Actual As Int32 = DetectionMaterial.GetPropertyValue(eMinPropID.SuperconductivePoint)
    '         Dim lD_SuperC_Known As Int32 = Owner.GetMineralPropertyKnowledge(lDetectionMineralID, eMinPropID.SuperconductivePoint)

    '         Dim lCo_ElectRes_Actual As Int32 = CollectionMaterial.GetPropertyValue(eMinPropID.ElectricalResist)
    '         Dim lCo_ElectRes_Known As Int32 = Owner.GetMineralPropertyKnowledge(lCollectionMineralID, eMinPropID.ElectricalResist)
    '         Dim lCo_Malleable_Actual As Int32 = CollectionMaterial.GetPropertyValue(eMinPropID.Malleable)
    '         Dim lCo_Malleable_Known As Int32 = Owner.GetMineralPropertyKnowledge(lCollectionMineralID, eMinPropID.Malleable)
    '         Dim lCo_MagProd_Actual As Int32 = CollectionMaterial.GetPropertyValue(eMinPropID.MagneticProduction)
    '         Dim lCo_MagProd_Known As Int32 = Owner.GetMineralPropertyKnowledge(lCollectionMineralID, eMinPropID.MagneticProduction)

    '         Dim lCa_ThermCond_Actual As Int32 = CasingMaterial.GetPropertyValue(eMinPropID.ThermalConductance)
    '         Dim lCa_ThermCond_Known As Int32 = Owner.GetMineralPropertyKnowledge(lCasingMineralID, eMinPropID.ThermalConductance)
    '         Dim lCa_ThermExp_Actual As Int32 = CasingMaterial.GetPropertyValue(eMinPropID.ThermalExpansion)
    '         Dim lCa_ThermExp_Known As Int32 = Owner.GetMineralPropertyKnowledge(lCasingMineralID, eMinPropID.ThermalExpansion)
    '         Dim lCa_Density_Actual As Int32 = CasingMaterial.GetPropertyValue(eMinPropID.Density)
    '         Dim lCa_Density_Known As Int32 = Owner.GetMineralPropertyKnowledge(lCasingMineralID, eMinPropID.Density)
    '         Dim lCa_MagProd_Actual As Int32 = CasingMaterial.GetPropertyValue(eMinPropID.MagneticProduction)
    '         Dim lCa_MagProd_Known As Int32 = Owner.GetMineralPropertyKnowledge(lCasingMineralID, eMinPropID.MagneticProduction)
    '         Dim lCa_MagReact_Actual As Int32 = CasingMaterial.GetPropertyValue(eMinPropID.MagneticReaction)
    '         Dim lCa_MagReact_Known As Int32 = Owner.GetMineralPropertyKnowledge(lCasingMineralID, eMinPropID.MagneticReaction)

    '         Dim fEmitterPowerRequired As Single
    '         If Me.RadarType = 0 Then
    '             fEmitterPowerRequired = 0
    '         Else
    '             fEmitterPowerRequired = ((CSng(Math.Pow(OptimumRange + 1, 1.3F)) * (MaximumRange + 1)) / 150.0F)
    '             If JamEffect <> 0 Then
    '                 If JamTargets = 0 Then
    '                     fEmitterPowerRequired += (CSng(Math.Sqrt((CSng(JamStrength) * CSng(OptimumRange)))) * 4)
    '                 Else
    '                     fEmitterPowerRequired += (CSng(Math.Sqrt((CSng(JamStrength) * CSng(JamTargets)))) * 4)
    '                 End If
    '             End If
    '             fEmitterPowerRequired += (50.0F * RandomSeed)
    '         End If
    '         fEmitterPowerRequired *= (fE_MagProd_Score * fE_MagReact_Score)

    '         Dim fEmitterSize As Single = ((((lE_ElectRes_Actual + 1) - lE_ElectRes_Known) / 2.0F) * fE_ElectRes_Score) * fEmitterPowerRequired

    '         Dim fDetectSensitivity As Single = 0
    '         If RadarType = 0 Then
    '             fDetectSensitivity = ((WeaponAcc / 4.0F) + 1) * ((OptimumRange + 1) * 2.0F) * (DisruptionResistance + 1.0F)
    '         Else
    '             fDetectSensitivity = ((WeaponAcc / 4.0F) + 1) * ((CSng(OptimumRange) + CSng(MaximumRange) + (DisruptionResistance * 3.0F) + 1.0F) * (ScanResolution + 1.0F))
    '         End If
    '         'Now, modify it...
    '         Dim fDetectSensMod As Single = (fDetectSensitivity * 2) / ((lD_MagReact_Actual + lD_MagReact_Known + 1) / ((lD_MagProd_Actual - (lD_MagProd_Known / 2.0F) + 1)))
    '         Dim fDetectorSize As Single = CSng(Math.Pow(fDetectSensMod, 1.3F))
    '         fDetectorSize /= ((lD_SuperC_Actual + 1 + lD_SuperC_Known) / 2.0F)
    '         If RadarType <> 0 Then fDetectorSize += 50

    '         Dim fMagneticNoise As Single = lD_MagProd_Actual + CSng(Math.Pow(lCo_MagProd_Actual, 2))
    '         fMagneticNoise += (lCa_MagProd_Actual * 10.0F)
    '         fMagneticNoise += (lCa_MagReact_Actual * 10.0F)
    '         fMagneticNoise += 1.0F
    '         If RadarType = 0 Then fMagneticNoise *= 1.54F

    '         Dim fCPU As Single
    '         If RadarType = 0 Then
    '             fCPU = (CSng(Math.Pow(CSng(OptimumRange + 1), 2)) + ((WeaponAcc + 1) * 5.0F)) * (CSng(Math.Pow(DisruptionResistance, 1.5F)) + 1)
    '         Else
    '             fCPU = ((CSng(OptimumRange) + CSng(MaximumRange) + CSng(DisruptionResistance * 5) + 1.0F) * _
    '               CSng(ScanResolution)) + ((CSng(WeaponAcc) + (DisruptionResistance * 3.0F)) * CSng(OptimumRange))
    '             If JamImmunity <> 0 Then fCPU *= 2
    '             fCPU += fMagneticNoise
    '         End If
    '         Dim fCollectorEfficiency As Single = fCPU / ((256.0F - lCo_ElectRes_Actual) * CSng((lCo_ElectRes_Known + 1) / (lCo_ElectRes_Actual + 1)))
    '         Dim fCPUPower As Single = fCPU / 100.0F

    '         Dim fHeatGenerated As Single = (fEmitterPowerRequired * 10.0F)
    '         fHeatGenerated += fDetectSensMod
    '         fHeatGenerated += (fCPU / fCollectorEfficiency)

    '         Dim fThermalRatio As Single = (fHeatGenerated / (lCa_ThermCond_Known + 1))
    'fThermalRatio += (fHeatGenerated / (lCa_ThermCond_Actual + 1))
    '         fThermalRatio /= 2.0F
    'fThermalRatio /= CSng((lCa_ThermExp_Actual - (lCa_ThermExp_Known / 2.0F) + 1.0F) / (lCa_Density_Actual + 1))

    '         Dim fCoolingEnergy As Single = (fThermalRatio * (RandomSeed + 0.5F))
    '         If RadarType = 0 Then fCoolingEnergy /= 100.0F Else fCoolingEnergy /= 80.0F

    '         Dim fCoolingSize As Single = CSng(Math.Sqrt(fThermalRatio))

    '         Dim fAvgAKScore As Single = 0
    '         Dim fTmpAKScore As Single = 0.0F
    '         Dim fMaxScore As Single = 0.0F

    '         If RadarType <> 0 Then
    '             fTmpAKScore = CSng((lE_ElectRes_Actual + 1) / (lE_ElectRes_Known + 1))
    '             fAvgAKScore += fTmpAKScore
    '             If fMaxScore < fTmpAKScore Then
    '                 fMaxScore = fTmpAKScore
    '                 MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop1
    '             End If

    '             fTmpAKScore = CSng((lE_MagProd_Actual + 1) / (lE_MagProd_Known + 1))
    '             fAvgAKScore += fTmpAKScore
    '             If fMaxScore < fTmpAKScore Then
    '                 fMaxScore = fTmpAKScore
    '                 MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop2
    '             End If

    '             fTmpAKScore = CSng((lE_MagReact_Actual + 1) / (lE_MagReact_Known + 1))
    '             fAvgAKScore += fTmpAKScore
    '             If fMaxScore < fTmpAKScore Then
    '                 fMaxScore = fTmpAKScore
    '                 MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop3
    '             End If
    '         End If
    '         fTmpAKScore = CSng((lD_MagReact_Actual + 1) / (lD_MagReact_Known + 1))
    '         fAvgAKScore += fTmpAKScore
    '         If fMaxScore < fTmpAKScore Then
    '             fMaxScore = fTmpAKScore
    '             MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop1
    '         End If

    '         fTmpAKScore = CSng((lD_MagProd_Actual + 1) / (lD_MagProd_Known + 1))
    '         fAvgAKScore += fTmpAKScore
    '         If fMaxScore < fTmpAKScore Then
    '             fMaxScore = fTmpAKScore
    '             MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop2
    '         End If

    '         fTmpAKScore = CSng((lD_SuperC_Actual + 1) / (lD_SuperC_Known + 1))
    '         fAvgAKScore += fTmpAKScore
    '         If fMaxScore < fTmpAKScore Then
    '             fMaxScore = fTmpAKScore
    '             MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop3
    '         End If

    '         fTmpAKScore = CSng((lCo_ElectRes_Actual + 1) / (lCo_ElectRes_Known + 1))
    '         fAvgAKScore += fTmpAKScore
    '         If fMaxScore < fTmpAKScore Then
    '             fMaxScore = fTmpAKScore
    '             MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop1
    '         End If

    '         fTmpAKScore = CSng((lCo_Malleable_Actual + 1) / (lCo_Malleable_Known + 1))
    '         fAvgAKScore += fTmpAKScore
    '         If fMaxScore < fTmpAKScore Then
    '             fMaxScore = fTmpAKScore
    '             MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop2
    '         End If

    '         fTmpAKScore = CSng((lCo_MagProd_Actual + 1) / (lCo_MagProd_Known + 1))
    '         fAvgAKScore += fTmpAKScore
    '         If fMaxScore < fTmpAKScore Then
    '             fMaxScore = fTmpAKScore
    '             MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop3
    '         End If

    '         fTmpAKScore = CSng((lCa_ThermCond_Actual + 1) / (lCa_ThermCond_Known + 1))
    '         fAvgAKScore += fTmpAKScore
    '         If fMaxScore < fTmpAKScore Then
    '             fMaxScore = fTmpAKScore
    '             MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop1
    '         End If

    '         fTmpAKScore = CSng((lCa_ThermExp_Actual + 1) / (lCa_ThermExp_Known + 1))
    '         fAvgAKScore += fTmpAKScore
    '         If fMaxScore < fTmpAKScore Then
    '             fMaxScore = fTmpAKScore
    '             MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop2
    '         End If

    '         fTmpAKScore = CSng((lCa_Density_Actual + 1) / (lCa_Density_Known + 1))
    '         fAvgAKScore += fTmpAKScore
    '         If fMaxScore < fTmpAKScore Then
    '             fMaxScore = fTmpAKScore
    '             MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop3
    '         End If

    '         fTmpAKScore = CSng((lCa_MagProd_Actual + 1) / (lCa_MagProd_Known + 1))
    '         fAvgAKScore += fTmpAKScore
    '         If fMaxScore < fTmpAKScore Then
    '             fMaxScore = fTmpAKScore
    '             MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop4
    '         End If

    '         fTmpAKScore = CSng((lCa_MagReact_Actual + 1) / (lCa_MagReact_Known + 1))
    '         fAvgAKScore += fTmpAKScore
    '         If fMaxScore < fTmpAKScore Then
    '             fMaxScore = fTmpAKScore
    '             MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop5
    '         End If

    '         If RadarType <> 0 Then
    '             fAvgAKScore /= 14.0F
    '         Else : fAvgAKScore /= 11.0F
    '         End If

    '         Dim fWpnAccScore As Single = GetLowLookup(WeaponAcc \ 10, 1)
    '         Dim fScanResScore As Single = GetLowLookup(ScanResolution \ 10, 1)
    '         Dim fOptRngScore As Single = GetLowLookup(OptimumRange \ 10, 0)
    '         Dim fMaxRngScore As Single = GetLowLookup(MaximumRange \ 10, 0)
    '         Dim fDisResScore As Single = GetLowLookup(DisruptionResistance \ 10, 2)
    '         Dim fJamStrScore As Single = GetLowLookup(JamStrength \ 10, 2)
    '         Dim fOverallScore As Single

    '         fOverallScore = (200 - fWpnAccScore) + (200 - fDisResScore) + (200 - fOptRngScore)
    '         If RadarType <> 0 Then
    '             fOverallScore += (200 - fScanResScore) + (200 - fMaxRngScore)
    '             If JamEffect <> 0 Then fOverallScore += (200 - fJamStrScore)
    '         End If
    '         Dim fL14 As Single
    '         If fOverallScore < 200000 Then
    '             'VLOOKUP(ROUNDUP(L13/10000*-1,0),LowLookup,2)
    '             fL14 = GetLowLookup(CInt(Math.Ceiling(fOverallScore / 10000)) * -1, 0)
    '         Else : fL14 = CInt(Math.Ceiling(fOverallScore / 10000.0F * -1))  'ROUNDUP(L13/10000*-1,0)
    '         End If
    '         fOverallScore = fL14 * CInt(Math.Ceiling(fOverallScore / 100000.0F * -1.0F))  '(ROUNDUP(L13 / 100000 * -1, 0))
    '         'fOverallScore = GetLowLookup(CInt(fOverallScore / 10000 * -1.0F), 0)
    '         If RadarType = 0 Then fOverallScore -= 100
    '         Me.CurrentSuccessChance = CInt(-fOverallScore)

    'Me.PowerRequired = CInt(Math.Ceiling(((fEmitterPowerRequired + fCPUPower + fCoolingEnergy) / 3.0F) * (RandomSeed + 0.5F)))
    '         'Me.HullRequired = CInt(((fEmitterSize + fDetectorSize + fCoolingSize) * 1.05F) / 30.0F)
    '         Me.HullRequired = 10

    '         Dim dTemp As Double
    '         dTemp = ((Math.Pow(DisruptionResistance, 2) + 1) + fCollectorEfficiency + fDetectSensMod)
    '         If RadarType <> 0 AndAlso JamEffect <> 0 Then
    '             If JamTargets = 0 Then
    '                 dTemp += Math.Pow(CSng(JamStrength) + CSng(OptimumRange), 2)
    '             Else : dTemp += ((CSng(JamStrength) + CSng(OptimumRange)) * CSng(JamTargets))
    '             End If
    '         End If
    '         dTemp *= (RandomSeed + 0.5F)
    '         dTemp *= Math.Sqrt(fAvgAKScore)
    '         If RadarType = 0 Then dTemp *= 1.1

    '         Dim fItoAC As Single = lEmitterComplexity + lDetectorComplexity + lCollectorComplexity + lCasingComplexity
    '         fItoAC /= 4.0F
    '         Dim fAvgACVal As Single = fItoAC
    '         fItoAC = Math.Max(1.0F, fItoAC - (PopIntel))


    'Dim blResearchCredits As Int64 = CLng(Math.Ceiling(dTemp))
    ''multiply credits by 10k outside of the points cost
    'blResearchCredits *= 10000
    'Dim blResearchPoints As Int64 = (blResearchCredits * CLng(Math.Ceiling(fItoAC)) \ 100L) + CLng(10000000 * RandomSeed)

    'blResearchPoints = GetResearchPointsCost(blResearchPoints)
    '         blResearchCredits *= Me.ResearchAttempts

    '         Dim blProductionCredits As Int64 = CLng(WeaponAcc) + CLng(OptimumRange) + CLng(DisruptionResistance)
    '         If RadarType <> 0 Then
    '             blProductionCredits += CLng(ScanResolution) + CLng(MaximumRange)
    '             If JamEffect <> 0 Then
    '                 If JamTargets = 0 Then
    '                     blProductionCredits += (CLng(OptimumRange) * CLng(JamStrength))
    '                 Else : blProductionCredits += (CLng(JamTargets) * CLng(JamStrength))
    '                 End If
    '             End If
    '         Else : blProductionCredits = CLng(blProductionCredits * 1.2F)
    '         End If
    '         blProductionCredits = CLng(blProductionCredits * fItoAC)
    'Dim blTemp As Int64 = CLng((((fCPU * fMagneticNoise) / (lCo_Malleable_Known + 1)) / (256.0F - fAvgACVal)) * (fDetectSensitivity / fDetectSensMod))
    '         If RadarType = 0 Then
    '             blProductionCredits *= CLng(blTemp * 0.9F)
    '         Else : blProductionCredits *= CLng(blTemp * 1.1F)
    'End If
    'blProductionCredits = CLng(blProductionCredits * (RandomSeed + 0.5F))
    '         Dim blProductionPoints As Int64 = (blProductionCredits * CLng(Math.Ceiling(fItoAC)) \ 10L) + CLng(1000000 * RandomSeed)

    'Dim lEmitterCosts As Int32 = CInt(Math.Ceiling(fEmitterSize * 1.1F * (RandomSeed + 0.5F)))
    'Dim lDetectorCosts As Int32 = CInt(Math.Ceiling(fDetectorSize * 1.1F * (RandomSeed + 0.5F)))
    'Dim lCollectorCosts As Int32 = CInt(Math.Ceiling((fCPU / (lCo_Malleable_Actual + lCo_Malleable_Known + 1)) * (RandomSeed + 0.5F)))
    '         Dim lCasingCosts As Int32 = 26  'formula was based on hull required

    '         If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
    '         With moResearchCost
    '             .ObjectID = Me.ObjectID
    '             .ObjTypeID = Me.ObjTypeID
    '             .ColonistCost = 0 : .EnlistedCost = 0 : .OfficerCost = 0
    '             .CreditCost = blResearchCredits
    '             Erase .ItemCosts
    '             .ItemCostUB = -1
    '             .PointsRequired = blResearchPoints

    '             .AddProductionCostItem(lEmitterMineralID, ObjectType.eMineral, lEmitterCosts)
    '             .AddProductionCostItem(lDetectionMineralID, ObjectType.eMineral, lDetectorCosts)
    '             .AddProductionCostItem(lCollectionMineralID, ObjectType.eMineral, lCollectorCosts)
    '             .AddProductionCostItem(lCasingMineralID, ObjectType.eMineral, lCasingCosts)
    '         End With
    '         If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
    '         With moProductionCost
    '             .ObjectID = Me.ObjectID
    '             .ObjTypeID = Me.ObjTypeID
    '             .ColonistCost = 0 : .EnlistedCost = 0 : .OfficerCost = 0
    '             .CreditCost = blProductionCredits
    '             Erase .ItemCosts
    '             .ItemCostUB = -1
    '             .PointsRequired = blProductionPoints

    '             .AddProductionCostItem(lEmitterMineralID, ObjectType.eMineral, lEmitterCosts)
    '             .AddProductionCostItem(lDetectionMineralID, ObjectType.eMineral, lDetectorCosts)
    '             .AddProductionCostItem(lCollectionMineralID, ObjectType.eMineral, lCollectorCosts)
    '             .AddProductionCostItem(lCasingMineralID, ObjectType.eMineral, lCasingCosts)
    '         End With

    '     Catch ex As Exception
    '         Me.ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
    '         Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
    '     End Try
    ' End Sub
    Protected Overrides Sub CalculateBothProdCosts()
        If moResearchCost Is Nothing OrElse moProductionCost Is Nothing Then ComponentDesigned()
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


    '		'=========== EMITTER ===========
    '		Dim uEm_ElectRes As MaterialPropertyItem2
    '		With uEm_ElectRes
    '			.lMineralID = EmitterMaterial.ObjectID
    '			.lPropertyID = eMinPropID.ElectricalResist

    '			.lActualValue = EmitterMaterial.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.3F
    '			.lGoalValue = 10
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
    '		Dim uEm_MagProd As MaterialPropertyItem2
    '		With uEm_MagProd
    '			.lMineralID = EmitterMaterial.ObjectID
    '			.lPropertyID = eMinPropID.MagneticProduction

    '			.lActualValue = EmitterMaterial.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.6F
    '			.lGoalValue = 130
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
    '		Dim uEm_MagReact As MaterialPropertyItem2
    '		With uEm_MagReact
    '			.lMineralID = EmitterMaterial.ObjectID
    '			.lPropertyID = eMinPropID.MagneticReaction

    '			.lActualValue = EmitterMaterial.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 130
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
    '		Dim lEmComplexity As Int32 = EmitterMaterial.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fEmInc As Single = (uEm_ElectRes.FinalScore + uEm_MagProd.FinalScore + uEm_MagReact.FinalScore)
    '		fEmInc *= CSng(lEmComplexity / PopIntel)
    '		'=========== END OF EMITTER =============

    '		'=========== DETECTOR ===================
    '		Dim uDe_MagReact As MaterialPropertyItem2
    '		With uDe_MagReact
    '			.lMineralID = DetectionMaterial.ObjectID
    '			.lPropertyID = eMinPropID.MagneticReaction

    '			.lActualValue = DetectionMaterial.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.3F
    '			.lGoalValue = 157 '150
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
    '		Dim uDe_MagProd As MaterialPropertyItem2
    '		With uDe_MagProd
    '			.lMineralID = DetectionMaterial.ObjectID
    '			.lPropertyID = eMinPropID.MagneticProduction

    '			.lActualValue = DetectionMaterial.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.3F
    '			.lGoalValue = 10
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
    '		Dim uDe_SuperC As MaterialPropertyItem2
    '		With uDe_SuperC
    '			.lMineralID = DetectionMaterial.ObjectID
    '			.lPropertyID = eMinPropID.SuperconductivePoint

    '			.lActualValue = DetectionMaterial.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.4F
    '			.lGoalValue = 121 ' 120
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
    '		Dim lDeComplexity As Int32 = DetectionMaterial.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fDeInc As Single = (uDe_MagProd.FinalScore + uDe_MagReact.FinalScore + uDe_SuperC.FinalScore)
    '		fDeInc *= CSng(lDeComplexity / PopIntel)
    '		'================ END OF DETECTOR =================

    '		'================ COLLECTOR =======================
    '		Dim uCo_ElectRes As MaterialPropertyItem2
    '		With uCo_ElectRes
    '			.lMineralID = CollectionMaterial.ObjectID
    '			.lPropertyID = eMinPropID.ElectricalResist

    '			.lActualValue = CollectionMaterial.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 1
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
    '		Dim uCo_Malleable As MaterialPropertyItem2
    '		With uCo_Malleable
    '			.lMineralID = CollectionMaterial.ObjectID
    '			.lPropertyID = eMinPropID.Malleable

    '			.lActualValue = CollectionMaterial.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 103 '100
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
    '		Dim uCo_MagProd As MaterialPropertyItem2
    '		With uCo_MagProd
    '			.lMineralID = CollectionMaterial.ObjectID
    '			.lPropertyID = eMinPropID.MagneticProduction

    '			.lActualValue = CollectionMaterial.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.8F
    '			.lGoalValue = 10
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
    '		Dim lCoComplexity As Int32 = CollectionMaterial.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fCoInc As Single = (uCo_ElectRes.FinalScore + uCo_MagProd.FinalScore + uCo_Malleable.FinalScore)
    '		fCoInc *= CSng(lCoComplexity / PopIntel)
    '		'================ END OF COLLECTOR ================

    '		'================ BEGIN CASING ====================
    '		Dim uCa_ThermCond As MaterialPropertyItem2
    '		With uCa_ThermCond
    '			.lMineralID = CasingMaterial.ObjectID
    '			.lPropertyID = eMinPropID.ThermalConductance

    '			.lActualValue = CasingMaterial.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 108 '100
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
    '		Dim uCa_ThermExp As MaterialPropertyItem2
    '		With uCa_ThermExp
    '			.lMineralID = CasingMaterial.ObjectID
    '			.lPropertyID = eMinPropID.ThermalExpansion

    '			.lActualValue = CasingMaterial.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 10
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
    '		Dim uCa_Density As MaterialPropertyItem2
    '		With uCa_Density
    '			.lMineralID = CasingMaterial.ObjectID
    '			.lPropertyID = eMinPropID.Density

    '			.lActualValue = CasingMaterial.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.2F
    '			.lGoalValue = 112 '100
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
    '		Dim uCa_MagProd As MaterialPropertyItem2
    '		With uCa_MagProd
    '			.lMineralID = CasingMaterial.ObjectID
    '			.lPropertyID = eMinPropID.MagneticProduction

    '			.lActualValue = CasingMaterial.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.5F
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
    '		Dim uCa_MagReact As MaterialPropertyItem2
    '		With uCa_MagReact
    '			.lMineralID = CasingMaterial.ObjectID
    '			.lPropertyID = eMinPropID.MagneticReaction

    '			.lActualValue = CasingMaterial.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 10
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop5 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat4_Prop5
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim lCaComplexity As Int32 = CasingMaterial.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fCaInc As Single = (uCa_Density.FinalScore + uCa_MagProd.FinalScore + uCa_MagReact.FinalScore + uCa_ThermCond.FinalScore + uCa_ThermExp.FinalScore)
    '		fCaInc *= CSng(lCaComplexity / PopIntel)
    '		'================ END OF CASING ===================

    '		If fHighestScore < 10.0F AndAlso bNotStudyFlaw = False Then
    '			MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eGoodDesign
    '		End If

    '		Dim fModifiedSeed As Single = 0.8F + (RandomSeed * 0.7F)
    '		Dim fSumIncs As Single = fEmInc + fCaInc + fCoInc + fDeInc
    '		Dim fAverageAK As Single = uEm_ElectRes.AKScore + uEm_MagProd.AKScore + uEm_MagReact.AKScore
    '		fAverageAK += uDe_MagProd.AKScore + uDe_MagReact.AKScore + uDe_MagReact.AKScore
    '		fAverageAK += uCo_ElectRes.AKScore + uCo_MagProd.AKScore + uCo_Malleable.AKScore
    '		fAverageAK += uCa_Density.AKScore + uCa_MagProd.AKScore + uCa_MagReact.AKScore + uCa_ThermCond.AKScore + uCa_ThermExp.AKScore
    '		fAverageAK /= 14.0F

    '		'Our modifier scores....
    '		Dim fMineralModifier As Single = fSumIncs / 4.0F
    '		Dim fWeaponAccMod As Single = fMineralModifier * CSng(WeaponAcc)
    '		Dim fScanResMod As Single = fCoInc * CSng(ScanResolution)
    '		Dim fOptRangeMod As Single = fDeInc * (CSng(OptimumRange) * CSng(OptimumRange))
    '		Dim fMaxRangeMod As Single = fCaInc * CSng(MaximumRange)
    '		Dim fDisResMod As Single = fCaInc * CSng(Math.Pow(DisruptionResistance, 0.5F + fModifiedSeed))
    '		Dim fJamMod As Single = 0.0F
    '		If JamImmunity <> 0 OrElse JamEffect <> 0 Then
    '			fJamMod = fMineralModifier * CSng(JamStrength)
    '		End If

    '		'now, the meat
    '		Dim fAverageComplexity As Single = lCaComplexity + lCoComplexity + lDeComplexity + lEmComplexity
    '		fAverageComplexity /= 4.0F

    '		Me.CurrentSuccessChance = CInt(fSumIncs * -1)

    '		'Hull Required...
    '		'=B6/5+IF(B9<>0,B11/10,0)
    '		If JamEffect <> 0 Then
    '			Me.HullRequired = CInt((CSng(MaximumRange) / 5.0F) + (CSng(JamStrength) / 10.0F))
    '		Else : Me.HullRequired = CInt(CSng(MaximumRange) / 5.0F)
    '		End If

    '		Dim lEnlisted As Int32 = CInt(Math.Floor(CInt(ScanResolution) * fDeInc / 100.0F))
    '		Dim lOfficers As Int32 = lEnlisted \ 5

    '		'Research Credits=IF(B1=1,,)
    '		Dim dblTemp As Double
    '		If RadarType = 1 Then
    '			'(B49+B44+B48)*10000+(B45+B46+B47)*1000*randomseed
    '			dblTemp = (fJamMod + fWeaponAccMod + fDisResMod) * 10000 + (fScanResMod + fOptRangeMod + fMaxRangeMod) * 1000 * fModifiedSeed
    '		Else
    '			'(B49+B44+B48)*1000+(B45+B46+B47)*100*randomseed
    '			dblTemp = (fJamMod + fWeaponAccMod + fDisResMod) * 1000 + (fScanResMod + fOptRangeMod + fMaxRangeMod) * 100 * fModifiedSeed
    '		End If
    '		Dim blResearchCredits As Int64 = CLng(dblTemp)

    '		'Research points  =B56*(B52/intelligence)*(0.5+randomseed)*20*E42
    '		Dim blResearchPoints As Int64 = CLng(dblTemp * (fAverageComplexity / PopIntel) * (0.5F + fModifiedSeed) * 20.0F * fAverageAK)

    '		'ProductionCredits=(B49+1)*(B44+1)*(B48+1)*(B47+1)
    '		dblTemp = (fJamMod + 1) * (fWeaponAccMod + 1) * (fDisResMod + 1) * (fMaxRangeMod + 1)
    '		Dim blProductionCredits As Int64 = CLng(Math.Max(fModifiedSeed * 500000, dblTemp))

    '		'Production Time=B58+SUM(B43:B49)*100*(randomseed*9)
    '		dblTemp += (fMineralModifier + fWeaponAccMod + fScanResMod + fOptRangeMod + fMaxRangeMod + fDisResMod + fJamMod) * 100 * (fModifiedSeed * 9)
    '		Dim blProductionPoints As Int64 = CLng(dblTemp)

    '		Dim lEmitterCosts As Int32 = CInt(fEmInc * 2)
    '		Dim lDetectorCosts As Int32 = CInt(fDeInc) ' * 1)
    '		Dim lCollectorCosts As Int32 = CInt(fCoInc * 10)
    '		Dim lCasingCosts As Int32 = CInt(fCaInc * 10)

    '		Me.HullRequired = Math.Max(Me.HullRequired, 10)

    '		If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
    '		With moResearchCost
    '			.ObjectID = Me.ObjectID
    '			.ObjTypeID = Me.ObjTypeID
    '			.ColonistCost = 0 : .EnlistedCost = 0 : .OfficerCost = 0
    '			.CreditCost = blResearchCredits
    '			Erase .ItemCosts
    '			.ItemCostUB = -1
    '			.PointsRequired = blResearchPoints

    '			.AddProductionCostItem(lEmitterMineralID, ObjectType.eMineral, lEmitterCosts)
    '			.AddProductionCostItem(lDetectionMineralID, ObjectType.eMineral, lDetectorCosts)
    '			.AddProductionCostItem(lCollectionMineralID, ObjectType.eMineral, lCollectorCosts)
    '			.AddProductionCostItem(lCasingMineralID, ObjectType.eMineral, lCasingCosts)
    '		End With
    '		If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
    '		With moProductionCost
    '			.ObjectID = Me.ObjectID
    '			.ObjTypeID = Me.ObjTypeID
    '			.ColonistCost = 0 : .EnlistedCost = lEnlisted : .OfficerCost = lOfficers
    '			.CreditCost = blProductionCredits
    '			Erase .ItemCosts
    '			.ItemCostUB = -1
    '			.PointsRequired = blProductionPoints

    '			.AddProductionCostItem(lEmitterMineralID, ObjectType.eMineral, Me.HullRequired)
    '			.AddProductionCostItem(lDetectionMineralID, ObjectType.eMineral, Me.HullRequired)
    '			.AddProductionCostItem(lCollectionMineralID, ObjectType.eMineral, Me.HullRequired)
    '			.AddProductionCostItem(lCasingMineralID, ObjectType.eMineral, Me.HullRequired)
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

            Dim oComputer As New RadarTechComputer
            With oComputer
                .blDisRes = DisruptionResistance
                .blJamStrength = JamStrength
                .blMaxRng = MaximumRange
                .blOptRng = OptimumRange
                .blScanRes = ScanResolution
                .blWepAcc = WeaponAcc
                .blJamTargets = JamTargets
                If JamTargets = 128 Then .blJamTargets = 20
                .lJamType = Me.JamEffect

                .lHullTypeID = RadarType

                .lMineral1ID = lEmitterMineralID
                .lMineral2ID = lDetectionMineralID
                .lMineral3ID = lCollectionMineralID
                .lMineral4ID = lCasingMineralID
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

                If .BuilderCostValueChange(Me.ObjTypeID, 1350000000D, PopIntel, Me.Owner.oSpecials.HullPerResidence) = False Then
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
        'Apply bonuses
        'Dim fMult As Single = 1.0F + (Owner.oSpecials.yRadarMaxRangeBonus / 100.0F)
        Dim lValue As Int32 = CInt(MaximumRange) + CInt(Owner.oSpecials.yRadarMaxRangeBonus) 'CInt(Math.Ceiling(MaximumRange * fMult))
        If lValue > 255 Then lValue = 255
        If lValue < 0 Then lValue = 0
        MaximumRange = CByte(lValue)

        'fMult = 1.0F + (Owner.oSpecials.yRadarOptRangeBonus / 100.0F)
        lValue = CInt(OptimumRange) + CInt(Owner.oSpecials.yRadarOptRangeBonus) 'CInt(Math.Ceiling(OptimumRange * fMult))
        If lValue > 255 Then lValue = 255
        If lValue < 0 Then lValue = 0
        OptimumRange = CByte(lValue)

        'fMult = 1.0F + (Owner.oSpecials.yRadarScanResBonus / 100.0F)
        'lValue = CInt(Math.Ceiling(ScanResolution * fMult))
        lValue = CInt(ScanResolution) + CInt(Owner.oSpecials.yRadarScanResBonus)
        If lValue > 255 Then lValue = 255
        If lValue < 0 Then lValue = 0
        ScanResolution = CByte(lValue)

        'fMult = 1.0F + (Owner.oSpecials.yRadarWpnAccBonus / 100.0F)
        'lValue = CInt(Math.Ceiling(WeaponAcc * fMult))
        lValue = CInt(WeaponAcc) + CInt(Owner.oSpecials.yRadarWpnAccBonus)
        If lValue > 255 Then lValue = 255
        If lValue < 0 Then lValue = 0
        WeaponAcc = CByte(lValue)
    End Sub

    Public Overrides Function GetNonOwnerMsg() As Byte()
        Dim yResult(40) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yResult, lPos) : lPos += 2
        Me.GetGUIDAsString.CopyTo(yResult, lPos) : lPos += 6

        RadarName.CopyTo(yResult, lPos) : lPos += 20
        yResult(lPos) = WeaponAcc : lPos += 1
        yResult(lPos) = ScanResolution : lPos += 1
        yResult(lPos) = OptimumRange : lPos += 1
        yResult(lPos) = MaximumRange : lPos += 1
        yResult(lPos) = DisruptionResistance : lPos += 1
        System.BitConverter.GetBytes(PowerRequired).CopyTo(yResult, lPos) : lPos += 4
        System.BitConverter.GetBytes(HullRequired).CopyTo(yResult, lPos) : lPos += 4

        Return yResult
    End Function

    Public Overrides Function TechnologyScore() As Integer
        If mlStoredTechScore = Int32.MinValue Then
            Try
                Dim lTemp As Int32
                If Me.RadarType = 0 Then
                    lTemp = CInt(WeaponAcc) + CInt(OptimumRange) + CInt(DisruptionResistance)
                Else
                    lTemp = CInt(WeaponAcc) + CInt(ScanResolution) + CInt(OptimumRange) + CInt(MaximumRange) + CInt(DisruptionResistance) + CInt(JamStrength)
                    If JamImmunity <> 0 Then
                        lTemp *= 2
                    End If
                    If JamStrength <> 0 AndAlso JamEffect <> 0 Then
                        If JamTargets <> 0 Then
                            lTemp += JamTargets
                        Else : lTemp += OptimumRange
                        End If
                    End If
                End If

                Dim dTemp As Double = Math.Pow(lTemp, 2) / Math.Max(1, Me.PowerRequired)
                mlStoredTechScore = CInt(dTemp / 1000.0F)
            Catch
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
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eNameOnly Then lPos += 9
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then lPos += 21
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then lPos += 92 '16
        ReDim yMsg(lPos - 1)
        lPos = 0

        'Default Attributes
        GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        Me.RadarName.CopyTo(yMsg, lPos) : lPos += 20

        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
            'at least settings 1 
            yMsg(lPos) = RadarType : lPos += 1
            System.BitConverter.GetBytes(PowerRequired).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(HullRequired).CopyTo(yMsg, lPos) : lPos += 4
        End If
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
            'at least settings 2  
            yMsg(lPos) = WeaponAcc : lPos += 1
            yMsg(lPos) = ScanResolution : lPos += 1
            yMsg(lPos) = OptimumRange : lPos += 1
            yMsg(lPos) = MaximumRange : lPos += 1
            yMsg(lPos) = DisruptionResistance : lPos += 1
            yMsg(lPos) = JamImmunity : lPos += 1
            yMsg(lPos) = JamStrength : lPos += 1
            yMsg(lPos) = JamTargets : lPos += 1
            yMsg(lPos) = JamEffect : lPos += 1
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
            System.BitConverter.GetBytes(lCollectionMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lCasingMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lEmitterMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lDetectionMineralID).CopyTo(yMsg, lPos) : lPos += 4

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
            ReDim .RadarName(19)
            Array.Copy(yData, lPos, .RadarName, 0, 20) : lPos += 20
            .WeaponAcc = yData(lPos) : lPos += 1
            .ScanResolution = yData(lPos) : lPos += 1
            .OptimumRange = yData(lPos) : lPos += 1
            .MaximumRange = yData(lPos) : lPos += 1
            .DisruptionResistance = yData(lPos) : lPos += 1

            .lCollectionMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lCasingMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lEmitterMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lDetectionMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .PowerRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .HullRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            .RadarType = yData(lPos) : lPos += 1
            .JamImmunity = yData(lPos) : lPos += 1
            .JamStrength = yData(lPos) : lPos += 1
            .JamTargets = yData(lPos) : lPos += 1
            .JamEffect = yData(lPos) : lPos += 1


            If .Owner Is Nothing = False Then
                Dim lFirstIdx As Int32 = -1
                For X As Int32 = 0 To .Owner.mlRadarUB
                    If .Owner.mlRadarIdx(X) = .ObjectID Then
                        .Owner.moRadar(X) = Me
                        Return
                    ElseIf lFirstIdx = -1 AndAlso .Owner.mlRadarIdx(X) = -1 Then
                        lFirstIdx = X
                    End If
                Next X

                If lFirstIdx = -1 Then
                    lFirstIdx = .Owner.mlRadarUB + 1
                    ReDim Preserve .Owner.mlRadarIdx(lFirstIdx)
                    ReDim Preserve .Owner.moRadar(lFirstIdx)
                    .Owner.mlRadarUB = lFirstIdx
                End If
                .Owner.moRadar(lFirstIdx) = Me
                .Owner.mlRadarIdx(lFirstIdx) = Me.ObjectID
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

        Dim oComputer As New RadarTechComputer
        With oComputer
            .blDisRes = DisruptionResistance
            .blJamStrength = JamStrength
            .blMaxRng = MaximumRange
            .blOptRng = OptimumRange
            .blScanRes = ScanResolution
            .blWepAcc = WeaponAcc
            .blJamTargets = JamTargets
            If JamTargets = 128 Then .blJamTargets = 20
            .lJamType = Me.JamEffect

            .lHullTypeID = RadarType

            .lMineral1ID = lEmitterMineralID
            .lMineral2ID = lDetectionMineralID
            .lMineral3ID = lCollectionMineralID
            .lMineral4ID = lCasingMineralID
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

            If .BuilderCostValueChange(Me.ObjTypeID, 1350000000D, PopIntel, Me.Owner.oSpecials.HullPerResidence) = False Then
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
