Public Class ArmorTech
    Inherits Epica_Tech

    Public ArmorName(19) As Byte

    Public yBeamResist As Byte
    Public yImpactResist As Byte
    Public yPiercingResist As Byte
    Public yMagneticResist As Byte
    Public yChemicalResist As Byte
    Public yBurnResist As Byte
    Public yRadarResist As Byte
    Public lHullUsagePerPlate As Int32
    Public lHPPerPlate As Int32
    Public lOuterLayerMineralID As Int32
    Public lMiddleLayerMineralID As Int32
    Public lInnerLayerMineralID As Int32

    Public lIntegrity As Int32

    Private mySendString() As Byte

    Public ReadOnly Property lDisplayedIntegrity() As Int32
        Get
            Dim lPercTime As Int32 = 100 - CInt(yIntegrityRoll)
            Return lPercTime
        End Get
    End Property
    Public ReadOnly Property yIntegrityRoll() As Byte
        Get
            Dim lTemp As Int32 = lIntegrity \ 2
            If lTemp > 255 Then lTemp = 255
            If lTemp < 0 Then lTemp = 0
            Return CByte(lTemp)
        End Get
    End Property

    Public Function GetObjAsString() As Byte()
        If moResearchCost Is Nothing OrElse moProductionCost Is Nothing Then
            CalculateBothProdCosts()
            mbStringReady = False
        End If

        'here we will return the entire object as a string
        'If mbStringReady = False Then
        ReDim mySendString(Epica_Tech.BASE_OBJ_STRING_SIZE + 50) '36)
        Dim lPos As Int32

        MyBase.GetBaseObjAsString(False).CopyTo(mySendString, 0)
        lPos = Epica_Tech.BASE_OBJ_STRING_SIZE

        ArmorName.CopyTo(mySendString, lPos) : lPos += 20
        mySendString(lPos) = yPiercingResist : lPos += 1
        mySendString(lPos) = yImpactResist : lPos += 1
        mySendString(lPos) = yBeamResist : lPos += 1
        mySendString(lPos) = yMagneticResist : lPos += 1
        mySendString(lPos) = yBurnResist : lPos += 1
        mySendString(lPos) = yChemicalResist : lPos += 1
        mySendString(lPos) = yRadarResist : lPos += 1
        System.BitConverter.GetBytes(lHPPerPlate).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lHullUsagePerPlate).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lOuterLayerMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lMiddleLayerMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lInnerLayerMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lDisplayedIntegrity).CopyTo(mySendString, lPos) : lPos += 4
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
                sSQL = "INSERT INTO tblArmor (ArmorName, OwnerID, PiercingResist, ImpactResist, BeamResist, " & _
                  "ECMResist, FlameResist, ChemicalResist, DetectionResist, HitPoints, HullUsagePerPlate, " & _
                  "OuterLayerMineralID, MiddleLayerMineralID, InnerLayerMineralID" & _
                  ", CurrentSuccessChance, SuccessChanceIncrement, ResearchAttempts, RandomSeed, ResPhase" & _
                  ", ErrorReasonCode, Integrity, MajorDesignFlaw, PopIntel, bArchived, VersionNumber) VALUES ('" & MakeDBStr(BytesToString(ArmorName)) & "', " & Owner.ObjectID & ", " & _
                  yPiercingResist & ", " & yImpactResist & ", " & yBeamResist & ", " & yMagneticResist & ", " & yBurnResist & _
                  ", " & yChemicalResist & ", " & yRadarResist & ", " & lHPPerPlate & ", " & lHullUsagePerPlate & _
                  ", " & lOuterLayerMineralID & ", " & lMiddleLayerMineralID & ", " & lInnerLayerMineralID & _
                  ", " & CurrentSuccessChance & ", " & SuccessChanceIncrement & ", " & ResearchAttempts & _
                  ", " & RandomSeed & ", " & ComponentDevelopmentPhase & ", " & ErrorReasonCode & ", " & lIntegrity & _
                  ", " & MajorDesignFlaw & ", " & PopIntel & ", " & yArchived & ", " & lVersionNum & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblArmor SET ArmorName = '" & MakeDBStr(BytesToString(ArmorName)) & "', OwnerID = " & Owner.ObjectID & _
                  ", PiercingResist = " & yPiercingResist & ", ImpactResist = " & yImpactResist & ", BeamResist = " & _
                  yBeamResist & ", ECMResist = " & yMagneticResist & ", FlameResist = " & yBurnResist & ", ChemicalResist = " & _
                  yChemicalResist & ", DetectionResist = " & yRadarResist & ", HitPoints = " & lHPPerPlate & _
                  ", HullUsagePerPlate = " & lHullUsagePerPlate & ", OuterLayerMineralID = " & lOuterLayerMineralID & _
                  ", MiddleLayerMineralID = " & lMiddleLayerMineralID & ", InnerLayerMineralID = " & _
                  lInnerLayerMineralID & _
                  ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
                  SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & _
                  RandomSeed & ", ResPhase = " & ComponentDevelopmentPhase & ", ErrorReasonCode = " & ErrorReasonCode & _
                  ", Integrity = " & lIntegrity & ", MajorDesignFlaw = " & MajorDesignFlaw & ", PopIntel = " & PopIntel & _
                  ", bArchived = " & yArchived & ", VersionNumber = " & lVersionNum & " WHERE ArmorID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(ArmorID) FROM tblArmor WHERE ArmorName = '" & MakeDBStr(BytesToString(ArmorName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If
            oComm = Nothing

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

        If ObjectID = -1 Then
            SaveObject()
            Return ""
        End If

        Try

            Dim oSB As New System.Text.StringBuilder

            'UPDATE
            sSQL = "UPDATE tblArmor SET ArmorName = '" & MakeDBStr(BytesToString(ArmorName)) & "', OwnerID = " & Owner.ObjectID & _
              ", PiercingResist = " & yPiercingResist & ", ImpactResist = " & yImpactResist & ", BeamResist = " & _
              yBeamResist & ", ECMResist = " & yMagneticResist & ", FlameResist = " & yBurnResist & ", ChemicalResist = " & _
              yChemicalResist & ", DetectionResist = " & yRadarResist & ", HitPoints = " & lHPPerPlate & _
              ", HullUsagePerPlate = " & lHullUsagePerPlate & ", OuterLayerMineralID = " & lOuterLayerMineralID & _
              ", MiddleLayerMineralID = " & lMiddleLayerMineralID & ", InnerLayerMineralID = " & _
              lInnerLayerMineralID & _
              ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
              SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & _
              RandomSeed & ", ResPhase = " & ComponentDevelopmentPhase & ", ErrorReasonCode = " & ErrorReasonCode & _
              ", Integrity = " & lIntegrity & ", MajorDesignFlaw = " & MajorDesignFlaw & ", PopIntel = " & PopIntel & _
              ", bArchived = " & yArchived & " WHERE ArmorID = " & ObjectID
            oSB.AppendLine(sSQL)
            oSB.AppendLine(MyBase.GetFinalizeSaveText())
            Return oSB.ToString
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        End Try
        Return ""
    End Function

    Public Overrides Function ValidateDesign() As Boolean
        'ensure that the player knows of the minerals being used
        Dim bMin1Good As Boolean = (lOuterLayerMineralID > 0) AndAlso Owner.IsMineralDiscovered(lOuterLayerMineralID)
        Dim bMin2Good As Boolean = (lMiddleLayerMineralID > 0) AndAlso Owner.IsMineralDiscovered(lMiddleLayerMineralID)
        Dim bMin3Good As Boolean = (lInnerLayerMineralID > 0) AndAlso Owner.IsMineralDiscovered(lInnerLayerMineralID)

        If lHullUsagePerPlate < 1 OrElse lHPPerPlate < 1 Then
            ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
            Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
            LogEvent(LogEventType.PossibleCheat, "ArmorTech.ValidateDesign hullusageperplate or hpperplate are invalid: " & Me.Owner.ObjectID)
            Return False
        End If

        If bMin1Good = False OrElse bMin2Good = False OrElse bMin3Good = False Then
            ErrorReasonCode = TechBuilderErrorReason.eInvalidMaterialsSelected
            Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
            LogEvent(LogEventType.PossibleCheat, "ArmorTech.ValidateDesign Minerals are bad: " & Me.Owner.ObjectID)
            Return False
        End If

        If lHullUsagePerPlate > 0 Then
            If lHPPerPlate / lHullUsagePerPlate > 30 Then
                ErrorReasonCode = TechBuilderErrorReason.eInvalidMaterialsSelected
                Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
                LogEvent(LogEventType.PossibleCheat, "ArmorTech.ValidateDesign Hull to HP Ratio Exceeds 30: " & Me.Owner.ObjectID)
                Return False
            End If
        End If

        Return True
    End Function

    Private moOuter As Mineral = Nothing
    Private ReadOnly Property OuterMineral() As Mineral
        Get
            If (moOuter Is Nothing OrElse moOuter.ObjectID <> lOuterLayerMineralID) AndAlso lOuterLayerMineralID > 0 Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = lOuterLayerMineralID Then
                        moOuter = goMineral(X)
                        Exit For
                    End If
                Next X
            End If
            Return moOuter
        End Get
    End Property

    Private moMiddle As Mineral = Nothing
    Private ReadOnly Property MiddleMineral() As Mineral
        Get
            If (moMiddle Is Nothing OrElse moMiddle.ObjectID <> lMiddleLayerMineralID) AndAlso lMiddleLayerMineralID > 0 Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = lMiddleLayerMineralID Then
                        moMiddle = goMineral(X)
                        Exit For
                    End If
                Next X
            End If
            Return moMiddle
        End Get
    End Property

    Private moInner As Mineral = Nothing
    Private ReadOnly Property InnerMineral() As Mineral
        Get
            If (moInner Is Nothing OrElse moInner.ObjectID <> lInnerLayerMineralID) AndAlso lInnerLayerMineralID > 0 Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = lInnerLayerMineralID Then
                        moInner = goMineral(X)
                        Exit For
                    End If
                Next X
            End If
            Return moInner
        End Get
    End Property

    '#Region " Excel Spreadsheet's K Value Calculations "
    '    Private Function GetAverageKVal() As Single
    '        Dim fTemp As Single = 0.0F
    '        fTemp += GetBeamResistKVal()
    '        fTemp += GetImpactResistKVal()
    '        fTemp += GetPiercingResistKVal()
    '        fTemp += GetMagneticResistKVal()
    '        fTemp += GetChemicalResistKVal()
    '        fTemp += GetBurnResistKVal()
    '        fTemp += GetRadarResistKVal()

    '        Return fTemp / 7.0F
    '    End Function

    '    Private Function GetBeamResistKVal() As Single
    '        Dim BeamResTimes3 As Int32 = yBeamResist * 3I       'G2

    '        Dim KVal As Single = 0.0F
    '        'K2 = SUM(H2 to J2)

    '        'H2 = IF ( G2 < D2, 0, (G2 - D2) ^ 1.3 )
    '        Dim lOuterReflect As Int32                          'D2
    '        lOuterReflect = OuterMineral.GetPropertyValue(eMinPropID.Reflection)
    '        If Not (BeamResTimes3 < lOuterReflect) Then
    '            KVal += CSng(Math.Pow((BeamResTimes3 - lOuterReflect), 1.3F))
    '        End If

    '        'I2 = IF ( G2 < E2, 0, (G2 - E2) ^ 1.1 )
    '        Dim lOuterRefract As Int32                          'E2
    '        lOuterRefract = OuterMineral.GetPropertyValue(eMinPropID.Refraction)
    '        If Not (BeamResTimes3 < lOuterRefract) Then
    '            KVal += CSng(Math.Pow((BeamResTimes3 - lOuterRefract), 1.1F))

    '            'F2 = E2....
    '            'J2 = IF ( G2 < F2, 0, (G2 - F2) ^ 1.2 )            
    '            KVal += CSng(Math.Pow((BeamResTimes3 - lOuterRefract), 1.2F))
    '        End If

    '        Return KVal
    '    End Function

    '    Private Function GetImpactResistKVal() As Single
    '        Dim ImpactResTimes3 As Int32 = yImpactResist * 3I       'G3

    '        Dim KVal As Single = 0.0F                               'K3 = Sum(H3:J3)

    '        'H3 = If ( G3 < D3, 0, (G3 - D3) ^ 2)
    '        Dim lOuterDensity As Int32                              'D3
    '        lOuterDensity = OuterMineral.GetPropertyValue(eMinPropID.Density)
    '        If Not (ImpactResTimes3 < lOuterDensity) Then
    '            KVal += CSng(Math.Pow((ImpactResTimes3 - lOuterDensity), 2))
    '        End If

    '        'I3 = If ( G3 < E3, 0, (G3 - E3) ^ 1.5)
    '        Dim lMiddleHardness As Int32                            'E3
    '        lMiddleHardness = MiddleMineral.GetPropertyValue(eMinPropID.Hardness)
    '        If Not (ImpactResTimes3 < lMiddleHardness) Then
    '            KVal += CSng(Math.Pow((ImpactResTimes3 - lMiddleHardness), 1.5))
    '        End If

    '        'J3 = If ( G3 < F3, 0, (G3 - F3) ^ 1.4)
    '        Dim lInnerCompress As Int32                             'F3
    '        lInnerCompress = InnerMineral.GetPropertyValue(eMinPropID.Compressibility)
    '        If Not (ImpactResTimes3 < lInnerCompress) Then
    '            KVal += CSng(Math.Pow((ImpactResTimes3 - lInnerCompress), 1.4))
    '        End If

    '        Return KVal
    '    End Function

    '    Private Function GetPiercingResistKVal() As Single
    '        Dim PierceResTimes3 As Int32 = yPiercingResist * 3I     'G4 = yPiercingResist
    '        Dim KVal As Single = 0.0F                               'K4 = SUM(H4:J4)

    '        'D4 = Outer.HardnessActual
    '        Dim lOuterHardness As Int32 = OuterMineral.GetPropertyValue(eMinPropID.Hardness)
    '        'H4 = IF ( G4 < D4, 0, (G4 - D4) ^ 2 )
    '        If Not (PierceResTimes3 < lOuterHardness) Then
    '            KVal += CSng(Math.Pow((PierceResTimes3 - lOuterHardness), 2))
    '        End If

    '        'E4 = Middle.MalleableActual
    '        Dim lMiddleMalleable As Int32 = MiddleMineral.GetPropertyValue(eMinPropID.Malleable)
    '        'I4 = IF ( G4 < E4, 0, (G4 - E4) ^ 1.4 )
    '        If Not (PierceResTimes3 < lMiddleMalleable) Then
    '            KVal += CSng(Math.Pow((PierceResTimes3 - lMiddleMalleable), 1.4))
    '        End If

    '        'F4 = Inner.CompressActual
    '        Dim lInnerCompress As Int32 = InnerMineral.GetPropertyValue(eMinPropID.Compressibility)
    '        'J4 = IF ( G4 < F4, 0, (G4 - F4) ^ 1.2 )
    '        If Not (PierceResTimes3 < lInnerCompress) Then
    '            KVal += CSng(Math.Pow((PierceResTimes3 - lInnerCompress), 1.2))
    '        End If

    '        Return KVal
    '    End Function

    '    Private Function GetMagneticResistKVal() As Single
    '        Dim lMagTimes3 As Int32 = yMagneticResist * 3I          'G5 = yMagneticResist * 3
    '        Dim KVal As Single = 0.0F                                 'K5 = SUM(H5:J5)

    '        'D5 = Outer.MagReactActual
    '        Dim lOuterMagReact As Int32 = OuterMineral.GetPropertyValue(eMinPropID.MagneticReaction)
    '        'H5 = IF ( G5 < D5, 0, (G5 - D5) ^ 1.8 )
    '        If Not (lMagTimes3 < lOuterMagReact) Then
    '            KVal += CSng(Math.Pow(lMagTimes3 - lOuterMagReact, 1.8))
    '        End If

    '        'E5 = Middle.QuantumActual
    '        Dim lMiddleQuantum As Int32 = MiddleMineral.GetPropertyValue(eMinPropID.Quantum)
    '        'I5 = IF ( G5 < E5, 0, (G5 - E5) ^ 1.1 )
    '        If Not (lMagTimes3 < lMiddleQuantum) Then
    '            KVal += CSng(Math.Pow(lMagTimes3 - lMiddleQuantum, 1.1))
    '        End If

    '        'F5 = Inner.MagReactActual
    '        Dim lInnerMagReact As Int32 = InnerMineral.GetPropertyValue(eMinPropID.MagneticReaction)
    '        'J5 = IF ( G5 < F5, 0, (G5 - F5) ^ 1.2 )
    '        If Not (lMagTimes3 < lInnerMagReact) Then
    '            KVal += CSng(Math.Pow(lMagTimes3 - lInnerMagReact, 1.2))
    '        End If

    '        Return KVal
    '    End Function

    '    Private Function GetChemicalResistKVal() As Single
    '        Dim lChemResTimes3 As Int32 = yChemicalResist * 3I              'G6 = ChemicalResist * 3
    '        Dim KVal As Single = 0.0F                                       'K6 = SUM(H6:J6)

    '        'D6 = Outer.ChemReactActual
    '        Dim lOuterChemReact As Int32 = OuterMineral.GetPropertyValue(eMinPropID.ChemicalReactance)
    '        'H6 = IF ( G6 < D6, 0, (G6 - D6) ^ 1.3 )
    '        If Not (lChemResTimes3 < lOuterChemReact) Then
    '            KVal += CSng(Math.Pow(lChemResTimes3 - lOuterChemReact, 1.3))
    '        End If

    '        'E6 = Middle.CompressActual
    '        Dim lMiddleCompress As Int32 = MiddleMineral.GetPropertyValue(eMinPropID.Compressibility)
    '        'I6 = IF ( G6 < E6, 0, (G6 - E6) ^ 1.1 )
    '        If Not (lChemResTimes3 < lMiddleCompress) Then
    '            KVal += CSng(Math.Pow(lChemResTimes3 - lMiddleCompress, 1.1))
    '        End If

    '        'F6 = Inner.ChemReactActual
    '        Dim lInnerChemReact As Int32 = InnerMineral.GetPropertyValue(eMinPropID.ChemicalReactance)
    '        'J6 = IF ( G6 < F6, 0, (G6 - F6) ^ 1.2 )
    '        If Not (lChemResTimes3 < lInnerChemReact) Then
    '            KVal += CSng(Math.Pow(lChemResTimes3 - lInnerChemReact, 1.2))
    '        End If

    '        Return KVal
    '    End Function

    '    Private Function GetBurnResistKVal() As Single
    '        Dim lBurnResTimes3 As Int32 = yBurnResist * 3I                      'G7 = BurnResist * 3
    '        Dim KVal As Single = 0.0F                                           'K7 = SUM(H7:J7)

    '        'D7 = Outer.CombustActual
    '        Dim lOuterCombust As Int32 = OuterMineral.GetPropertyValue(eMinPropID.Combustiveness)
    '        'H7 = IF ( G7 < D7, 0, (G7 - D7) ^ 1.5 )
    '        If Not (lBurnResTimes3 < lOuterCombust) Then
    '            KVal += CSng(Math.Pow(lBurnResTimes3 - lOuterCombust, 1.5))
    '        End If

    '        'E7 = Middle.DensityActual
    '        Dim lMiddleDensity As Int32 = MiddleMineral.GetPropertyValue(eMinPropID.Density)
    '        'I7 = IF ( G7 < E7, 0, (G7 - E7) ^ 1.1 )
    '        If Not (lBurnResTimes3 < lMiddleDensity) Then
    '            KVal += CSng(Math.Pow(lBurnResTimes3 - lMiddleDensity, 1.1))
    '        End If

    '        'F7 = Inner.CombustActual
    '        Dim lInnerCombust As Int32 = InnerMineral.GetPropertyValue(eMinPropID.Combustiveness)
    '        'J7 = IF ( G7 < F7, 0, (G7 - F7) ^ 1.2 )
    '        If Not (lBurnResTimes3 < lInnerCombust) Then
    '            KVal += CSng(Math.Pow(lBurnResTimes3 - lInnerCombust, 1.2))
    '        End If

    '        Return KVal

    '    End Function

    '    Private Function GetRadarResistKVal() As Single
    '        Dim lRadarResTimes3 As Int32 = yRadarResist * 3I                    'G8 = RadarResist * 3
    '        Dim KVal As Single = 0.0F                                           'K8 = SUM(H8:J8)

    '        'D8 = Outer.RefractActual
    '        Dim lOuterRefract As Int32 = OuterMineral.GetPropertyValue(eMinPropID.Refraction)
    '        'H8 = IF ( G8 < D8, 0, (G8 - D8) ^ 2 )
    '        If Not (lRadarResTimes3 < lOuterRefract) Then
    '            KVal += CSng(Math.Pow(lRadarResTimes3 - lOuterRefract, 2))
    '        End If

    '        'E8 = Middle.RefractActual
    '        Dim lMiddleRefract As Int32 = MiddleMineral.GetPropertyValue(eMinPropID.Refraction)
    '        'I8 = IF ( G8 < E8, 0, (G8 - E8) ^ 1.1 )
    '        If Not (lRadarResTimes3 < lMiddleRefract) Then
    '            KVal += CSng(Math.Pow(lRadarResTimes3 - lMiddleRefract, 1.1))
    '        End If

    '        'F8 = Inner.RefractActual
    '        Dim lInnerRefract As Int32 = InnerMineral.GetPropertyValue(eMinPropID.Refraction)
    '        'J8 = IF ( G8 < F8, 0, (G8 - F8) ^ 1.2 )
    '        If Not (lRadarResTimes3 < lInnerRefract) Then
    '            KVal += CSng(Math.Pow(lRadarResTimes3 - lInnerRefract, 1.2))
    '        End If

    '        Return KVal
    '    End Function

    '    Private Function GetSummedResistKVal() As Int32
    '        Dim fTmp As Single = GetBeamResistKVal() + GetImpactResistKVal() + GetPiercingResistKVal()
    '        fTmp += GetMagneticResistKVal() + GetChemicalResistKVal() + GetBurnResistKVal() + GetRadarResistKVal()
    '        Return CInt(fTmp)
    '    End Function
    '#End Region

    Public Overrides Function SetFromDesignMsg(ByRef yData() As Byte) As Boolean
        Dim lPos As Int32 = 14  '2 for msg code, 2 for objtypeid, 4 for id, 6 for researcher
        Dim bRes As Boolean = False
        Try
            With Me
                .ObjTypeID = ObjectType.eArmorTech
                .yBeamResist = yData(lPos) : lPos += 1
                .yImpactResist = yData(lPos) : lPos += 1
                .yPiercingResist = yData(lPos) : lPos += 1
                .yMagneticResist = yData(lPos) : lPos += 1
                .yChemicalResist = yData(lPos) : lPos += 1
                .yBurnResist = yData(lPos) : lPos += 1
                .yRadarResist = yData(lPos) : lPos += 1
                .lHullUsagePerPlate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lHPPerPlate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lOuterLayerMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lMiddleLayerMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lInnerLayerMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                ReDim .ArmorName(19)
                Array.Copy(yData, lPos, .ArmorName, 0, 20)
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
            LogEvent(LogEventType.CriticalError, "ArmorTech.SetFromDesignMsg: " & ex.Message)
        End Try

        Return bRes
    End Function

    Protected Overrides Sub CalculateBothProdCosts()
        If moResearchCost Is Nothing OrElse moProductionCost Is Nothing Then
            ComponentDesigned()
        End If

        'Dim bValid As Boolean = True

        'If moResearchCost Is Nothing Then
        '    moResearchCost = New ProductionCost

        '    Try

        '        moResearchCost.ObjectID = Me.ObjectID
        '        moResearchCost.ObjTypeID = Me.ObjTypeID

        '        Dim fAvgKCost As Single = GetAverageKVal() * RandomSeed
        '        If fAvgKCost < 101.0F Then fAvgKCost = 101

        '        Dim f_I24 As Single = 100.0F / GetPopIntel()

        '        Dim f_d16 As Single = CSng(Math.Pow((fAvgKCost - 100), 1.6))

        '        Dim f_J18 As Single = CSng(lHPPerPlate / lHullUsagePerPlate)

        '        Dim f_J20 As Single = f_J18 / l_NORMALIZED_HP_PER_HULL
        '        Dim f_d17 As Single = f_d16 * f_J20 * f_I24
        '        moResearchCost.CreditCost = CInt((f_d17 + lHPPerPlate) * RandomSeed)

        '        Dim lFinal As Int32
        '        lFinal = Owner.GetMineralPropertyKnowledge(lOuterLayerMineralID, eMinPropID.Reflection)
        '        lFinal = Math.Min(Owner.GetMineralPropertyKnowledge(lOuterLayerMineralID, eMinPropID.Density), lFinal)
        '        lFinal = Math.Min(Owner.GetMineralPropertyKnowledge(lOuterLayerMineralID, eMinPropID.Hardness), lFinal)
        '        lFinal = Math.Min(Owner.GetMineralPropertyKnowledge(lOuterLayerMineralID, eMinPropID.MagneticReaction), lFinal)
        '        lFinal = Math.Min(Owner.GetMineralPropertyKnowledge(lOuterLayerMineralID, eMinPropID.ChemicalReactance), lFinal)
        '        lFinal = Math.Min(Owner.GetMineralPropertyKnowledge(lOuterLayerMineralID, eMinPropID.Combustiveness), lFinal)
        '        lFinal = Math.Min(Owner.GetMineralPropertyKnowledge(lOuterLayerMineralID, eMinPropID.Refraction), lFinal)
        '        Me.CurrentSuccessChance = lFinal
        '        Me.SuccessChanceIncrement = 1

        '        moResearchCost.ColonistCost = 0 : moResearchCost.EnlistedCost = 0 : moResearchCost.OfficerCost = 0

        '        'Points Required
        '        moResearchCost.PointsRequired = GetResearchPointsCost(moResearchCost.CreditCost)
        '        moResearchCost.CreditCost *= Me.ResearchAttempts
        '        'TODO: Reset Research Attempts?

        '        'Minerals Required
        '        With moResearchCost
        '            .ItemCostUB = -1
        '            Erase .ItemCosts

        '            .AddProductionCostItem(lOuterLayerMineralID, ObjectType.eMineral, CInt(Math.Ceiling(Math.Pow(lHullUsagePerPlate, 1.3))))
        '            .AddProductionCostItem(lMiddleLayerMineralID, ObjectType.eMineral, CInt(Math.Ceiling(lHullUsagePerPlate * 0.3F)))
        '            .AddProductionCostItem(lInnerLayerMineralID, ObjectType.eMineral, CInt(Math.Ceiling(lHullUsagePerPlate * 0.3F)))
        '        End With
        '    Catch
        '        ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
        '        bValid = False
        '    End Try
        'End If

        'If moProductionCost Is Nothing Then
        '    moProductionCost = New ProductionCost
        '    moProductionCost.ObjectID = Me.ObjectID
        '    moProductionCost.ObjTypeID = Me.ObjTypeID
        '    moProductionCost.ColonistCost = 0
        '    moProductionCost.EnlistedCost = 0
        '    moProductionCost.OfficerCost = 0

        '    Try

        '        Dim lTmp As Int32 = GetSummedResistKVal()
        '        Dim fJ_20 As Single = (CSng(lHPPerPlate / lHullUsagePerPlate) / l_NORMALIZED_HP_PER_HULL)
        '        lTmp \= 1000I
        '        lTmp *= CInt(Math.Pow(fJ_20, 2))
        '        lTmp = CInt(lTmp * Math.Pow((100.0F / GetPopIntel()), 1.6))
        '        lTmp = CInt(Math.Ceiling(lTmp * RandomSeed))
        '        moProductionCost.CreditCost = lTmp

        '        Dim fTmp As Single = GetAverageKVal()
        '        If fTmp * RandomSeed > 101 Then
        '            fTmp *= RandomSeed
        '        Else : fTmp = 101.0F
        '        End If
        '        fTmp *= (lHullUsagePerPlate * CSng(Math.Pow(fJ_20, 2)))
        '        fTmp = fTmp / CSng(Math.Pow((100.0F / GetPopIntel()), 1.2) * RandomSeed) / 10.0F
        '        moProductionCost.PointsRequired = CLng(Math.Ceiling(fTmp))

        '        With moProductionCost
        '            .ItemCostUB = -1
        '            Erase .ItemCosts

        '            'Outer layer = HullUsagePerPlate ^ 1.3
        '            .AddProductionCostItem(lOuterLayerMineralID, ObjectType.eMineral, CInt(Math.Ceiling(Math.Pow(lHullUsagePerPlate, 1.3))))
        '            'Middle Layer = RoundUp(HullUsagePerPlate * .3)
        '            .AddProductionCostItem(lMiddleLayerMineralID, ObjectType.eMineral, CInt(Math.Ceiling(lHullUsagePerPlate * 0.3F)))
        '            'Inner Layer = RoundUp(HullUsagePerPlate * .3)
        '            .AddProductionCostItem(lInnerLayerMineralID, ObjectType.eMineral, CInt(Math.Ceiling(lHullUsagePerPlate * 0.3F)))
        '        End With
        '    Catch
        '        ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
        '        bValid = False
        '    End Try
        'End If

        'If bValid = False Then Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
    End Sub

    ' Public Overrides Sub ComponentDesigned()
    '     Try
    '         'ensure our factors are initialized (if they are already, the routine will break out)
    '         InitializeFactors()

    '         If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
    '         With moResearchCost
    '             .ObjectID = Me.ObjectID
    '             .ObjTypeID = Me.ObjTypeID
    '             Erase .ItemCosts
    '             .ItemCostUB = -1
    '         End With
    '         If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
    '         With moProductionCost
    '             .ObjectID = Me.ObjectID
    '             .ObjTypeID = Me.ObjTypeID
    '             Erase .ItemCosts
    '             .ItemCostUB = -1
    '         End With

    '         Dim fResistRands() As Single
    '         Dim lScores() As Int32
    '         ReDim lScores(eFactorType.eLastFactor - 1)
    '         ReDim fResistRands(eFactorType.eLastFactor - 1)

    '         'Use the same sequence of rands for this technology
    '         Rnd(-1)
    '         Randomize(RandomSeed)
    '         For X As Int32 = 0 To fResistRands.GetUpperBound(0)
    '             fResistRands(X) = Rnd()
    '         Next X

    '         Dim lAKScoreValue As Int32 = 0
    '         Dim lAKScoreValueCnt As Int32 = 0

    '         For X As Int32 = 0 To lFactorUB
    '             With muArmorFactors(X)
    '                 Dim lActual As Int32
    '                 Dim lKnown As Int32

    '                 If .lArmorLayer = 0 Then        'Outer layer
    '                     lActual = OuterMineral.GetPropertyValue(.lFactorID)
    '                     lKnown = Owner.GetMineralPropertyKnowledge(OuterMineral.ObjectID, .lFactorID)
    '                 ElseIf .lArmorLayer = 1 Then    'Middle Layer
    '                     lActual = MiddleMineral.GetPropertyValue(.lFactorID)
    '                     lKnown = Owner.GetMineralPropertyKnowledge(MiddleMineral.ObjectID, .lFactorID)
    '                 Else                            'Inner Layer
    '                     lActual = InnerMineral.GetPropertyValue(.lFactorID)
    '                     lKnown = Owner.GetMineralPropertyKnowledge(InnerMineral.ObjectID, .lFactorID)
    '                 End If

    '                 If lActual <> lKnown Then
    '                     If MajorDesignFlaw = 0 Then
    '                         If .lArmorLayer = 0 Then
    '                             MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop1
    '                         ElseIf .lArmorLayer = 1 Then
    '                             MajorDesignFlaw = eComponentDesignFlaw.eMat2_Prop1
    '                         Else : MajorDesignFlaw = eComponentDesignFlaw.eMat3_Prop1
    '                         End If
    '                     End If
    '                 End If

    '                 lAKScoreValue += ((lKnown + 1) \ (lActual + 1)) * 100
    '                 lAKScoreValueCnt += 1

    '                 Dim lPosMult As Int32 = lActual - (lActual - lKnown)
    '                 Dim lNegMult As Int32 = lActual

    '                 'Now, work the magic...
    '                 If .lFactorValue < 0 Then
    '                     lScores(.lFactorType) += (.lFactorValue * lNegMult)
    '                 Else : lScores(.lFactorType) += (.lFactorValue * lPosMult)
    '                 End If
    '             End With
    '         Next X

    '         'Now, our scores should be known, let's get our multipliers
    '         For X As Int32 = 0 To lScores.GetUpperBound(0)
    '             lScores(X) = lMaxFactorScore(X) - lScores(X)
    '         Next X

    '         'Now, we have our multipliers, let's get our ResistanceCosts
    '         Dim lResistCosts(lScores.GetUpperBound(0)) As Int32
    '         lResistCosts(eFactorType.BeamFactor) = yBeamResist * lScores(eFactorType.BeamFactor)
    '         lResistCosts(eFactorType.BurnFactor) = yBurnResist * lScores(eFactorType.BurnFactor)
    '         lResistCosts(eFactorType.ImpactFactor) = yImpactResist * lScores(eFactorType.ImpactFactor)
    '         lResistCosts(eFactorType.MagneticFactor) = yMagneticResist * lScores(eFactorType.MagneticFactor)
    '         lResistCosts(eFactorType.PiercingFactor) = yPiercingResist * lScores(eFactorType.PiercingFactor)
    '         'lResistCosts(eFactorType.RadarFactor) = yRadarResist * lScores(eFactorType.RadarFactor)
    '         lResistCosts(eFactorType.ToxicFactor) = yChemicalResist * lScores(eFactorType.ToxicFactor)

    '         Dim lMaxCost As Int32 = 0
    '         If MajorDesignFlaw = 0 Then
    '             For X As Int32 = 0 To lScores.GetUpperBound(0)
    '                 If lResistCosts(X) > lMaxCost Then
    '                     lMaxCost = lResistCosts(X)
    '                     MajorDesignFlaw = CByte(X + 16)
    '                 End If
    '             Next X
    '         End If

    '         'Now, we can also calculate our integrity
    '         '(((lScore / lMaxScore) * ResistValue) * ResistsRandom)
    '         Dim fIntegrity As Single = 0.0F
    '         lIntegrity = 0
    '         For X As Int32 = 0 To lScores.GetUpperBound(0)
    '             fIntegrity += ((CSng(lScores(X) / lMaxFactorScore(X)) * GetFactorsResist(X)) * fResistRands(X))
    '         Next X
    '         lIntegrity = CInt(fIntegrity)

    '         'Now, let's get the rest of the data
    '         Dim fHPPerHull As Single = CSng(Me.lHPPerPlate / Me.lHullUsagePerPlate)
    '         Dim fHPPerHullMult As Single = fHPPerHull / Me.Owner.oSpecials.yArmorHullToHP

    '         Dim fLookupResult As Single = GetLowLookup(CInt(Math.Floor(fHPPerHullMult)), 3)
    '         Dim lTotalResistCosts As Int32 = 0
    '         For X As Int32 = 0 To lResistCosts.GetUpperBound(0)
    '             lTotalResistCosts += lResistCosts(X)
    '         Next X
    'fLookupResult *= Math.Max(lTotalResistCosts, Math.Max(200 * ((RandomSeed * 0.7F) + 0.8F), 1))

    '         Dim blResearchCostCredits As Int64 = CLng((fLookupResult + 1) * 100 * (RandomSeed + 0.5F))
    '         Dim blResearchCostPoints As Int64 = blResearchCostCredits \ 1000
    '         Dim blProdCostCredits As Int64 = CLng(((((fLookupResult / 1000.0F) * Me.lHPPerPlate) / Me.lHullUsagePerPlate) + 1) * (RandomSeed + 0.5F))
    '         Dim blProdCostPoints As Int64 = blProdCostCredits \ 1000

    '         lAKScoreValue \= lAKScoreValueCnt

    '         Dim fAvgResist As Single = 0.0F
    '         fAvgResist = (CInt(yBeamResist) + CInt(yImpactResist) + CInt(yPiercingResist) + CInt(yMagneticResist) + CInt(yChemicalResist) + CInt(yBurnResist) + CInt(yRadarResist))
    '         fAvgResist /= eFactorType.eLastFactor

    '         Me.CurrentSuccessChance = CInt(lAKScoreValue - fAvgResist - GetLowLookup(CInt(Math.Floor(fHPPerHullMult)), 3))
    '         Me.SuccessChanceIncrement = 1

    '         If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
    '         Dim bValid As Boolean = True
    '         Try

    '             moResearchCost.ObjectID = Me.ObjectID
    '             moResearchCost.ObjTypeID = Me.ObjTypeID
    '             moResearchCost.CreditCost = blResearchCostCredits
    '             moResearchCost.ColonistCost = 0 : moResearchCost.EnlistedCost = 0 : moResearchCost.OfficerCost = 0
    '             moResearchCost.PointsRequired = GetResearchPointsCost(blResearchCostPoints)
    '             moResearchCost.CreditCost *= Me.ResearchAttempts
    '             'TODO: Reset Research Attempts?

    '             'Minerals Required
    '             With moResearchCost
    '                 .ItemCostUB = -1
    '                 Erase .ItemCosts

    '                 .AddProductionCostItem(lOuterLayerMineralID, ObjectType.eMineral, CInt(Math.Ceiling(Math.Pow(lHullUsagePerPlate, 1.3))))
    '                 .AddProductionCostItem(lMiddleLayerMineralID, ObjectType.eMineral, CInt(Math.Ceiling(lHullUsagePerPlate * 0.3F)))
    '                 .AddProductionCostItem(lInnerLayerMineralID, ObjectType.eMineral, CInt(Math.Ceiling(lHullUsagePerPlate * 0.3F)))
    '             End With
    '         Catch
    '             ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
    '             bValid = False
    '         End Try

    '         If moProductionCost Is Nothing Then moProductionCost = New ProductionCost

    '         Try
    '             moProductionCost.ObjectID = Me.ObjectID
    '             moProductionCost.ObjTypeID = Me.ObjTypeID
    '             moProductionCost.ColonistCost = 0
    '             moProductionCost.EnlistedCost = 0
    '             moProductionCost.OfficerCost = 0
    '             moProductionCost.CreditCost = blProdCostCredits
    '             moProductionCost.PointsRequired = blProdCostPoints

    '             With moProductionCost
    '                 .ItemCostUB = -1
    '                 Erase .ItemCosts

    '                 'Outer layer = HullUsagePerPlate ^ 1.3
    '                 .AddProductionCostItem(lOuterLayerMineralID, ObjectType.eMineral, CInt(Math.Ceiling(Math.Pow(lHullUsagePerPlate, 1.3))))
    '                 'Middle Layer = RoundUp(HullUsagePerPlate * .3)
    '                 .AddProductionCostItem(lMiddleLayerMineralID, ObjectType.eMineral, CInt(Math.Ceiling(lHullUsagePerPlate * 0.3F)))
    '                 'Inner Layer = RoundUp(HullUsagePerPlate * .3)
    '                 .AddProductionCostItem(lInnerLayerMineralID, ObjectType.eMineral, CInt(Math.Ceiling(lHullUsagePerPlate * 0.3F)))
    '             End With
    '         Catch
    '             ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
    '             bValid = False
    '         End Try
    '         If bValid = False Then Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
    '     Catch
    '         ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
    '         Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
    '     End Try

    ' End Sub

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

            Dim oComputer As New ArmorTechComputer
            With oComputer
                .blBeam = yBeamResist
                .blBurn = yBurnResist
                .blHPPerPlate = lHPPerPlate
                .blHullUsagePerPlate = lHullUsagePerPlate
                .blImpact = yImpactResist
                .blMagnetic = yMagneticResist
                .blPiercing = yPiercingResist
                .blToxic = yChemicalResist
                .lHPRatioSpecial = Me.Owner.oSpecials.yArmorHullToHP
                .lHullTypeID = 0
                .lMineral1ID = lOuterLayerMineralID
                .lMineral2ID = lMiddleLayerMineralID
                .lMineral3ID = lInnerLayerMineralID
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

                If .ArmorBuilderCostValueChange(Me.ObjTypeID, 50D, PopIntel) = False Then
                    ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
                    Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
                    Return
                End If
            End With

            Me.CurrentSuccessChance = 100 'CInt(lAKScoreValue - fAvgResist - GetLowLookup(CInt(Math.Floor(fHPPerHullMult)), 3))
            Me.SuccessChanceIncrement = 1


            Dim lSumOfResist As Int32 = CInt(yBeamResist) + CInt(yBurnResist) + CInt(yChemicalResist) + CInt(yImpactResist) + CInt(yMagneticResist) + CInt(yPiercingResist) '+ CInt(yRadarResist)
            '=100 - ((H6-150)^1.4)
            Dim lSumAllDA As Int32 = oComputer.lSumAllDA
            lSumAllDA = Math.Max(1, 200 - lSumAllDA)
            lSumOfResist = (lSumOfResist - lSumAllDA)
            If lSumOfResist < 0 Then
                lIntegrity = 0
            Else
                lIntegrity = Math.Min(100, CInt(Math.Pow(lSumOfResist, 1.2)))
            End If
            lIntegrity *= 2

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

    'Private Function GetFactorsResist(ByVal lFactor As Int32) As Int32
    '    Select Case lFactor
    '        Case eFactorType.BeamFactor
    '            Return yBeamResist
    '        Case eFactorType.BurnFactor
    '            Return yBurnResist
    '        Case eFactorType.ImpactFactor
    '            Return yImpactResist
    '        Case eFactorType.MagneticFactor
    '            Return yMagneticResist
    '        Case eFactorType.PiercingFactor
    '            Return yPiercingResist
    '            'Case eFactorType.RadarFactor
    '            '    Return yRadarResist
    '        Case eFactorType.ToxicFactor
    '            Return yChemicalResist
    '        Case Else
    '            Return 0
    '    End Select
    'End Function

    Protected Overrides Sub FinalizeResearch()
        Dim lBR As Int32 = yBeamResist
        Dim lIR As Int32 = yImpactResist
        Dim lPR As Int32 = yPiercingResist
        Dim lMR As Int32 = yMagneticResist
        Dim lCR As Int32 = yChemicalResist
        Dim lFR As Int32 = yBurnResist

        Dim lSum As Int32 = lBR
        lSum += lIR
        lSum += lPR
        lSum += lMR
        lSum += lCR
        lSum += lFR
        If lSum > 200 Then
            LogEvent(LogEventType.Warning, "Design of ArmorTech over 200 total resist belonging to " & Me.Owner.ObjectID & ", TechID: " & Me.ObjectID)
        End If

        'Ok, check for 2x techs
        If (Owner.oSpecials.lNoResistAt2X And Player_Specials.elNo2XValues.eNoBurn2XImpact) <> 0 Then
            If lFR = 0 Then lIR *= 2
        End If
        If (Owner.oSpecials.lNoResistAt2X And Player_Specials.elNo2XValues.eNoChem2XBurn) <> 0 Then
            If lCR = 0 Then lFR *= 2
        End If
        If (Owner.oSpecials.lNoResistAt2X And Player_Specials.elNo2XValues.eNoImpact2XBurn) <> 0 Then
            If lIR = 0 Then lFR *= 2
        End If
        If (Owner.oSpecials.lNoResistAt2X And Player_Specials.elNo2XValues.eNoMag2XChem) <> 0 Then
            If lMR = 0 Then lCR *= 2
        End If
        If (Owner.oSpecials.lNoResistAt2X And Player_Specials.elNo2XValues.eNoPierce2XBeam) <> 0 Then
            If lPR = 0 Then lBR *= 2
        End If

        'Ok, adjust the resistance values now...
        Dim fMult As Single = 1.0F + (Owner.oSpecials.yBeamResistImprove / 100.0F)
        Dim lValue As Int32 = CInt(Math.Ceiling(lBR * fMult))
        If lValue > 100 Then lValue = 100
        If lValue < 0 Then lValue = 0
        yBeamResist = CByte(lValue)

        fMult = 1.0F + (Owner.oSpecials.yBurnResistImprove / 100.0F)
        lValue = CInt(Math.Ceiling(lFR * fMult))
        If lValue > 100 Then lValue = 100
        If lValue < 0 Then lValue = 0
        yBurnResist = CByte(lValue)

        fMult = 1.0F + (Owner.oSpecials.yChemResistImprove / 100.0F)
        lValue = CInt(Math.Ceiling(lCR * fMult))
        If lValue > 100 Then lValue = 100
        If lValue < 0 Then lValue = 0
        yChemicalResist = CByte(lValue)

        fMult = 1.0F + (Owner.oSpecials.yImpactResistImprove / 100.0F)
        lValue = CInt(Math.Ceiling(lIR * fMult))
        If lValue > 100 Then lValue = 100
        If lValue < 0 Then lValue = 0
        yImpactResist = CByte(lValue)

        fMult = 1.0F + (Owner.oSpecials.yMagneticResistImprove / 100.0F)
        lValue = CInt(Math.Ceiling(lMR * fMult))
        If lValue > 100 Then lValue = 100
        If lValue < 0 Then lValue = 0
        yMagneticResist = CByte(lValue)

        fMult = 1.0F + (Owner.oSpecials.yPierceResistImprove / 100.0F)
        lValue = CInt(Math.Ceiling(lPR * fMult))
        If lValue > 100 Then lValue = 100
        If lValue < 0 Then lValue = 0
        yPiercingResist = CByte(lValue)

        fMult = 1.0F + (Owner.oSpecials.yRadarResistImprove / 100.0F)
        lValue = CInt(Math.Ceiling(yRadarResist * fMult))
        If lValue > 100 Then lValue = 100
        If lValue < 0 Then lValue = 0
        yRadarResist = CByte(lValue)
    End Sub

    Public Function GetMaxProperty(ByVal lPropID As Int32) As Int32
        Return GetMaxVal(OuterMineral.GetPropertyValue(lPropID), MiddleMineral.GetPropertyValue(lPropID), InnerMineral.GetPropertyValue(lPropID))
    End Function

    Public Overrides Function GetNonOwnerMsg() As Byte()
        Dim yResult(46) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yResult, lPos) : lPos += 2
        Me.GetGUIDAsString.CopyTo(yResult, lPos) : lPos += 6

        ArmorName.CopyTo(yResult, lPos) : lPos += 20
        yResult(lPos) = yPiercingResist : lPos += 1
        yResult(lPos) = yImpactResist : lPos += 1
        yResult(lPos) = yBeamResist : lPos += 1
        yResult(lPos) = yMagneticResist : lPos += 1
        yResult(lPos) = yBurnResist : lPos += 1
        yResult(lPos) = yChemicalResist : lPos += 1
        yResult(lPos) = yRadarResist : lPos += 1
        System.BitConverter.GetBytes(lHPPerPlate).CopyTo(yResult, lPos) : lPos += 4
        System.BitConverter.GetBytes(lHullUsagePerPlate).CopyTo(yResult, lPos) : lPos += 4
        System.BitConverter.GetBytes(lDisplayedIntegrity).CopyTo(yResult, lPos) : lPos += 4

        Return yResult
    End Function

    '#Region "  Resist Factors  "
    '    Private Enum eFactorType As Integer
    '        BeamFactor = 0
    '        ImpactFactor = 1
    '        PiercingFactor = 2
    '        MagneticFactor = 3
    '        ToxicFactor = 4
    '        BurnFactor = 5
    '        'RadarFactor = 6

    '        eLastFactor          'should be the last factor ALWAYS
    '    End Enum

    '    Private Structure ArmorFactor
    '        Public lFactorID As Int32           'property ID of the factor
    '        Public lFactorType As eFactorType
    '        Public lFactorValue As Int32        'factor value
    '        Public lArmorLayer As Int32         '0 = outer, 1 = middle, 2 = inner
    '    End Structure

    '    Private Shared muArmorFactors() As ArmorFactor
    '    Private Shared lFactorUB As Int32 = -1
    '    Private Shared lFactorScore() As Int32
    '    Private Shared lMaxFactorScore() As Int32

    '    Private Shared Sub InitializeFactors()
    '        If lFactorUB <> -1 Then Return

    '        ReDim lFactorScore(eFactorType.eLastFactor - 1)
    '        ReDim lMaxFactorScore(eFactorType.eLastFactor - 1)
    '        For X As Int32 = 0 To lFactorScore.GetUpperBound(0)
    '            lFactorScore(X) = 0
    '        Next X

    '        Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand("SELECT * FROM tblArmorFactor", goCN)
    '        Dim oData As OleDb.OleDbDataReader = oComm.ExecuteReader()

    '        While oData.Read()
    '            lFactorUB += 1
    '            ReDim Preserve muArmorFactors(lFactorUB)
    '            With muArmorFactors(lFactorUB)
    '                .lFactorID = CInt(oData("FactorID"))
    '                .lFactorType = CType(CInt(oData("FactorType")), eFactorType)
    '                .lFactorValue = CInt(oData("FactorValue"))
    '                .lArmorLayer = CInt(oData("ArmorLayer"))

    '                If .lFactorValue > 0 Then lFactorScore(.lFactorType) += .lFactorValue
    '            End With
    '        End While
    '        oData.Close()
    '        oData = Nothing
    '        oComm.Dispose()
    '        oComm = Nothing

    '        For X As Int32 = 0 To lFactorScore.GetUpperBound(0)
    '            lMaxFactorScore(X) = lFactorScore(X) * 255
    '        Next X
    '    End Sub

    '#End Region

    Public Overrides Function TechnologyScore() As Integer
        If mlStoredTechScore = Int32.MinValue Then
            Try
                Dim lTemp As Int32 = CInt(yBeamResist) + CInt(yBurnResist) + CInt(yChemicalResist) + CInt(yImpactResist) + CInt(yMagneticResist) + CInt(yPiercingResist) + CInt(yRadarResist)
                lTemp *= Me.lHPPerPlate
                lTemp \= Me.lHullUsagePerPlate
                mlStoredTechScore = (lTemp * Me.lIntegrity) \ 1000
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
        lPos = 30
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eNameOnly Then lPos += 8
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then lPos += 23
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then lPos += 88 '12
        ReDim yMsg(lPos - 1)
        lPos = 0

        'Default attributes
        GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        Me.ArmorName.CopyTo(yMsg, lPos) : lPos += 20

        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
            'at least settings 1 which is HP Per Plate and Hull Usage Per plate
            System.BitConverter.GetBytes(lHPPerPlate).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lHullUsagePerPlate).CopyTo(yMsg, lPos) : lPos += 4
        End If
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
            'at least settings 2 which is the armor resistances and integrity
            yMsg(lPos) = yPiercingResist : lPos += 1
            yMsg(lPos) = yImpactResist : lPos += 1
            yMsg(lPos) = yBeamResist : lPos += 1
            yMsg(lPos) = yMagneticResist : lPos += 1
            yMsg(lPos) = yBurnResist : lPos += 1
            yMsg(lPos) = yChemicalResist : lPos += 1
            yMsg(lPos) = yRadarResist : lPos += 1
            System.BitConverter.GetBytes(lDisplayedIntegrity).CopyTo(yMsg, lPos) : lPos += 4

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
            System.BitConverter.GetBytes(lOuterLayerMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lMiddleLayerMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lInnerLayerMineralID).CopyTo(yMsg, lPos) : lPos += 4

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

            ReDim .ArmorName(19)
            Array.Copy(yData, lPos, .ArmorName, 0, 20) : lPos += 20
            .yPiercingResist = yData(lPos) : lPos += 1
            .yImpactResist = yData(lPos) : lPos += 1
            .yBeamResist = yData(lPos) : lPos += 1
            .yMagneticResist = yData(lPos) : lPos += 1
            .yBurnResist = yData(lPos) : lPos += 1
            .yChemicalResist = yData(lPos) : lPos += 1
            .yRadarResist = yData(lPos) : lPos += 1

            .lHPPerPlate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lHullUsagePerPlate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lOuterLayerMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lMiddleLayerMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lInnerLayerMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            Dim lDispInt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'ok, this is our displayed integrity... add 100
            lDispInt += 100
            'now... to get our final integrity... multiply by 2
            lDispInt *= 2
            lIntegrity = lDispInt

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

            If .Owner Is Nothing = False Then
                Dim lFirstIdx As Int32 = -1
                For X As Int32 = 0 To .Owner.mlArmorUB
                    If .Owner.mlArmorIdx(X) = .ObjectID Then
                        .Owner.moArmor(X) = Me
                        Return
                    ElseIf lFirstIdx = -1 AndAlso .Owner.mlArmorIdx(X) = -1 Then
                        lFirstIdx = X
                    End If
                Next X

                If lFirstIdx = -1 Then
                    lFirstIdx = .Owner.mlArmorUB + 1
                    ReDim Preserve .Owner.mlArmorIdx(lFirstIdx)
                    ReDim Preserve .Owner.moArmor(lFirstIdx)
                    .Owner.mlArmorUB = lFirstIdx
                End If
                .Owner.moArmor(lFirstIdx) = Me
                .Owner.mlArmorIdx(lFirstIdx) = Me.ObjectID
            End If
        End With


    End Sub
End Class
