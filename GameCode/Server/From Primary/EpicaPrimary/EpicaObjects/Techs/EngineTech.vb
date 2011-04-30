Public Class EngineTech
    Inherits Epica_Tech

    Public EngineName(19) As Byte

    Public Thrust As Int32
    Public Maneuver As Byte
    Public Speed As Byte
    Public PowerProd As Int32

    Public lStructuralBodyMineralID As Int32
    Public lStructuralFrameMineralID As Int32
    Public lStructuralMeldMineralID As Int32

    Public lDriveBodyMineralID As Int32
    Public lDriveFrameMineralID As Int32
    Public lDriveMeldMineralID As Int32

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
        ReDim mySendString(Epica_Tech.BASE_OBJ_STRING_SIZE + 59)

        Dim lPos As Int32
        MyBase.GetBaseObjAsString(False).CopyTo(mySendString, 0)
        lPos = Epica_Tech.BASE_OBJ_STRING_SIZE

        EngineName.CopyTo(mySendString, lPos) : lPos += 20

        System.BitConverter.GetBytes(Thrust).CopyTo(mySendString, lPos) : lPos += 4
        mySendString(lPos) = Maneuver : lPos += 1
        mySendString(lPos) = Speed : lPos += 1
        System.BitConverter.GetBytes(PowerProd).CopyTo(mySendString, lPos) : lPos += 4

        System.BitConverter.GetBytes(lStructuralBodyMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lStructuralFrameMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lStructuralMeldMineralID).CopyTo(mySendString, lPos) : lPos += 4

        System.BitConverter.GetBytes(lDriveBodyMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lDriveFrameMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lDriveMeldMineralID).CopyTo(mySendString, lPos) : lPos += 4

        'System.BitConverter.GetBytes(lFuelCompositionMineralID).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(lFuelCatalystMineralID).CopyTo(mySendString, lPos) : lPos += 4

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
                sSQL = "INSERT INTO tblEngine (EngineName, Thrust, Maneuver, Speed, PowerProd, " & _
                  "OwnerID, StructuralBodyMineralID, StructuralFrameMineralID, StructuralMeldMineralID, " & _
                  "DriveBodyMineralID, DriveFrameMineralID, DriveMeldMineralID, HullRequired, ColorValue" & _
                  ", CurrentSuccessChance, SuccessChanceIncrement, ResearchAttempts, RandomSeed, ResPhase" & _
                  ", ErrorReasonCode, MajorDesignFlaw, PopIntel, bArchived, HullTypeID) VALUES ('" & _
                  MakeDBStr(BytesToString(EngineName)) & "', " & Thrust & ", " & Maneuver & ", " & Speed & ", " & _
                  PowerProd & ", " & Owner.ObjectID & ", " & lStructuralBodyMineralID & ", " & _
                  lStructuralFrameMineralID & ", " & lStructuralMeldMineralID & ", " & lDriveBodyMineralID & _
                  ", " & lDriveFrameMineralID & ", " & lDriveMeldMineralID & ", " & HullRequired & ", " & ColorValue & _
                  ", " & CurrentSuccessChance & ", " & SuccessChanceIncrement & ", " & ResearchAttempts & _
                  ", " & RandomSeed & ", " & ComponentDevelopmentPhase & ", " & ErrorReasonCode & ", " & _
                  MajorDesignFlaw & ", " & PopIntel & ", " & yArchived & ", " & HullTypeID & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblEngine SET EngineName = '" & MakeDBStr(BytesToString(EngineName)) & "', Thrust = " & Thrust & ", Maneuver = " & _
                  Maneuver & ", Speed = " & Speed & ", PowerProd = " & PowerProd & ", OwnerID = " & _
                  Owner.ObjectID & ", StructuralBodyMineralID = " & lStructuralBodyMineralID & _
                  ", StructuralFrameMineralID = " & lStructuralFrameMineralID & ", StructuralMeldMineralID = " & _
                  lStructuralMeldMineralID & ", DriveBodyMineralID = " & lDriveBodyMineralID & _
                  ", DriveFrameMineralID = " & lDriveFrameMineralID & ", DriveMeldMineralID = " & _
                  lDriveMeldMineralID & ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
                  SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & _
                  RandomSeed & ", ResPhase = " & ComponentDevelopmentPhase & _
                  ", HullRequired = " & HullRequired & ", ColorValue = " & ColorValue & ", ErrorReasonCode = " & _
                  Me.ErrorReasonCode & ", MajorDesignFlaw = " & MajorDesignFlaw & ", PopIntel = " & PopIntel & _
                  ", bArchived = " & yArchived & ", HullTypeID = " & HullTypeID & " WHERE EngineID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(EngineID) FROM tblEngine WHERE EngineName = '" & MakeDBStr(BytesToString(EngineName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing

                ''Ok, set our object in the master array
                'Dim bAdded As Boolean = False
                'For X As Int32 = 0 To glEngineUB
                '    If glEngineIdx(X) = -1 Then
                '        glEngineIdx(X) = ObjectID
                '        goEngine(X) = Me
                '        bAdded = True
                '        Exit For
                '    End If
                'Next X
                'If bAdded = True Then
                '    glEngineUB += 1
                '    ReDim Preserve glEngineIdx(glEngineUB)
                '    ReDim Preserve goEngine(glEngineUB)
                '    glEngineIdx(glEngineUB) = ObjectID
                '    goEngine(glEngineUB) = Me
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

        If ObjectID = -1 Then
            SaveObject()
            Return ""
        End If

        Try
            Dim oSB As New System.Text.StringBuilder

            'UPDATE
            sSQL = "UPDATE tblEngine SET EngineName = '" & MakeDBStr(BytesToString(EngineName)) & "', Thrust = " & Thrust & ", Maneuver = " & _
              Maneuver & ", Speed = " & Speed & ", PowerProd = " & PowerProd & ", OwnerID = " & _
              Owner.ObjectID & ", StructuralBodyMineralID = " & lStructuralBodyMineralID & _
              ", StructuralFrameMineralID = " & lStructuralFrameMineralID & ", StructuralMeldMineralID = " & _
              lStructuralMeldMineralID & ", DriveBodyMineralID = " & lDriveBodyMineralID & _
              ", DriveFrameMineralID = " & lDriveFrameMineralID & ", DriveMeldMineralID = " & _
              lDriveMeldMineralID & ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
              SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & _
              RandomSeed & ", ResPhase = " & ComponentDevelopmentPhase & _
              ", HullRequired = " & HullRequired & ", ColorValue = " & ColorValue & ", ErrorReasonCode = " & _
              Me.ErrorReasonCode & ", MajorDesignFlaw = " & MajorDesignFlaw & ", PopIntel = " & PopIntel & _
              ", bArchived = " & yArchived & ", HullTypeID = " & HullTypeID & " WHERE EngineID = " & ObjectID
            oSB.AppendLine(sSQL)
            oSB.AppendLine(MyBase.GetFinalizeSaveText())
            Return oSB.ToString
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        End Try
        Return ""
    End Function

    Private moDriveMeldMineral As Mineral
    Private Function DriveMeldMineral() As Mineral
        If moDriveMeldMineral Is Nothing OrElse moDriveMeldMineral.ObjectID <> lDriveMeldMineralID Then
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = lDriveMeldMineralID Then
                    moDriveMeldMineral = goMineral(X)
                    Exit For
                End If
            Next X
        End If
        Return moDriveMeldMineral
    End Function

    Private moStructureBodyMineral As Mineral
    Private Function StructureBodyMineral() As Mineral
        If moStructureBodyMineral Is Nothing OrElse moStructureBodyMineral.ObjectID <> lStructuralBodyMineralID Then
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = lStructuralBodyMineralID Then
                    moStructureBodyMineral = goMineral(X)
                    Exit For
                End If
            Next X
        End If
        Return moStructureBodyMineral
    End Function

    Private moDriveBodyMineral As Mineral
    Private Function DriveBodyMineral() As Mineral
        If moDriveBodyMineral Is Nothing OrElse moDriveBodyMineral.ObjectID <> lDriveBodyMineralID Then
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = lDriveBodyMineralID Then
                    moDriveBodyMineral = goMineral(X)
                    Exit For
                End If
            Next X
        End If
        Return moDriveBodyMineral
    End Function

    Private moStructureMeldMineral As Mineral
    Private Function StructureMeldMineral() As Mineral
        If moStructureMeldMineral Is Nothing OrElse moStructureMeldMineral.ObjectID <> lStructuralMeldMineralID Then
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = lStructuralMeldMineralID Then
                    moStructureMeldMineral = goMineral(X)
                    Exit For
                End If
            Next X
        End If
        Return moStructureMeldMineral
    End Function

    Private moDriveFrameMineral As Mineral
    Private Function DriveFrameMineral() As Mineral
        If moDriveFrameMineral Is Nothing OrElse moDriveFrameMineral.ObjectID <> lDriveFrameMineralID Then
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = lDriveFrameMineralID Then
                    moDriveFrameMineral = goMineral(X)
                    Exit For
                End If
            Next X
        End If
        Return moDriveFrameMineral
    End Function

    Private moStructureFrameMineral As Mineral
    Private Function StructureFrameMineral() As Mineral
        If moStructureFrameMineral Is Nothing OrElse moStructureFrameMineral.ObjectID <> lStructuralFrameMineralID Then
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = lStructuralFrameMineralID Then
                    moStructureFrameMineral = goMineral(X)
                    Exit For
                End If
            Next X
        End If
        Return moStructureFrameMineral
    End Function

    Public Overrides Function ValidateDesign() As Boolean
        'Player needs to set all mineral ID's
        Dim bMinGood As Boolean = _
            (lStructuralBodyMineralID > 0 AndAlso Owner.IsMineralDiscovered(lStructuralBodyMineralID)) AndAlso _
            (lStructuralFrameMineralID > 0 AndAlso Owner.IsMineralDiscovered(lStructuralFrameMineralID)) AndAlso _
            (lStructuralMeldMineralID > 0 AndAlso Owner.IsMineralDiscovered(lStructuralMeldMineralID)) AndAlso _
            (lDriveBodyMineralID > 0 AndAlso Owner.IsMineralDiscovered(lDriveBodyMineralID)) AndAlso _
            (lDriveFrameMineralID > 0 AndAlso Owner.IsMineralDiscovered(lDriveFrameMineralID)) AndAlso _
            (lDriveMeldMineralID > 0 AndAlso Owner.IsMineralDiscovered(lDriveMeldMineralID))

        'Do any other validation here
        Dim lMaxManeuver As Int32 = Owner.oSpecials.yMaxManeuver
        Dim lMaxMaxSpeed As Int32 = Owner.oSpecials.yMaxSpeed
        Dim lPowerThrustLimit As Int32 = Owner.oSpecials.lPowerThrustLimit

        If Maneuver > lMaxManeuver OrElse Speed > lMaxMaxSpeed OrElse PowerProd > lPowerThrustLimit OrElse Thrust > lPowerThrustLimit Then
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
            LogEvent(LogEventType.PossibleCheat, "EngineTech.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If

        If bMinGood = False Then
            LogEvent(LogEventType.PossibleCheat, "EngineTech.ValidateDesign invalid value: " & Me.Owner.ObjectID)
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidMaterialsSelected
        Else
            Me.ErrorReasonCode = TechBuilderErrorReason.eNoError
        End If

        If CheckDesignImpossibility() = True Then
            LogEvent(LogEventType.PossibleCheat, "EngineTech.ValidateDesign CheckDesignPossiblity Failed: " & Me.Owner.ObjectID)
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
            Return False
        End If

        Return bMinGood

    End Function

    Private Function CheckDesignImpossibility() As Boolean
        Dim oTB As New EngineTechComputer()
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
            oTB.lMineral1ID = .lStructuralBodyMineralID
            oTB.lMineral2ID = .lStructuralFrameMineralID
            oTB.lMineral3ID = .lStructuralMeldMineralID
            oTB.lMineral4ID = .lDriveBodyMineralID
            oTB.lMineral5ID = .lDriveFrameMineralID
            oTB.lMineral6ID = .lDriveMeldMineralID

            'Bogus: Do not back out specials.  These bonus's are applied after the design is finalized.
            'oTB.blMan = Math.Max(0, CInt(.Maneuver) - CInt(Owner.oSpecials.yManeuverBonus))
            'If .Maneuver > 0 Then oTB.blMan = Math.Max(1, oTB.blMan)
            'oTB.blMaxSpd = Math.Max(0, CInt(.Speed) - CInt(Owner.oSpecials.yMaxSpeedBonus))
            'If .Speed > 0 Then oTB.blMaxSpd = Math.Max(1, oTB.blMaxSpd)
            'Dim fMult As Single = 1.0F + (Owner.oSpecials.yPowerBonus / 100.0F)

            oTB.blMaxSpd = .Speed
            oTB.blPowGen = .PowerProd
            oTB.blThrustGen = .Thrust
            oTB.blMan = .Maneuver

            oTB.blThrustGen = .Thrust
        End With
        Return oTB.IsDesignImpossible(Me.ObjTypeID, Owner)
    End Function

    Public Overrides Function SetFromDesignMsg(ByRef yData() As Byte) As Boolean
        Dim lPos As Int32 = 14  '2 for msg code, 2 for objtypeid, 4 for id, 6 for researcher guid
        Dim bRes As Boolean = False
        Try
            With Me
                .ObjTypeID = ObjectType.eEngineTech
                .PowerProd = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .Thrust = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                Speed = yData(lPos) : lPos += 1
                .Maneuver = yData(lPos) : lPos += 1
                .lStructuralBodyMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lStructuralFrameMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lStructuralMeldMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lDriveBodyMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lDriveFrameMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lDriveMeldMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                '.lFuelCompositionMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                '.lFuelCatalystMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .ColorValue = yData(lPos) : lPos += 1

                ReDim .EngineName(19)
                Array.Copy(yData, lPos, .EngineName, 0, 20)
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
            LogEvent(LogEventType.PossibleCheat, "EngineTech.SetFromDesignMsg: " & ex.Message)
        End Try

        Return bRes
    End Function

    Protected Overrides Sub CalculateBothProdCosts()
        If moResearchCost Is Nothing OrElse moProductionCost Is Nothing Then ComponentDesigned()

        'If moResearchCost Is Nothing OrElse moProductionCost Is Nothing Then
        '    If moResearchCost Is Nothing Then moResearchCost = New ProductionCost
        '    moResearchCost.ObjectID = Me.ObjectID
        '    moResearchCost.ObjTypeID = Me.ObjTypeID

        '    Dim f_D2 As Single
        '    Dim f_G2 As Single
        '    Dim f_H2 As Single
        '    Dim f_D3 As Single
        '    Dim f_H3 As Single
        '    Dim f_H4 As Single
        '    Dim f_E2 As Single
        '    Dim f_D4 As Single
        '    Dim f_H5 As Single

        '    Dim lStrBodyMagProd As Int32 = StructureBodyMineral.GetPropertyValue(eMinPropID.MagneticProduction)
        '    Dim lStrBodyMagProdKnow As Int32 = Owner.GetMineralPropertyKnowledge(lStructuralBodyMineralID, eMinPropID.MagneticProduction)
        '    Dim lDrvBodyMagReact As Int32 = DriveBodyMineral.GetPropertyValue(eMinPropID.MagneticReaction)
        '    Dim lDrvBodyMagReactKnow As Int32 = Owner.GetMineralPropertyKnowledge(lDriveBodyMineralID, eMinPropID.MagneticReaction)
        '    Dim lStrMeldMeltPt As Int32 = StructureMeldMineral.GetPropertyValue(eMinPropID.MeltingPoint)
        '    Dim lStrMeldMeltPtKnow As Int32 = Owner.GetMineralPropertyKnowledge(lStructuralMeldMineralID, eMinPropID.MeltingPoint)
        '    Dim lStrMeldMall As Int32 = StructureMeldMineral.GetPropertyValue(eMinPropID.Malleable)
        '    Dim lStrFrameMall As Int32 = StructureFrameMineral.GetPropertyValue(eMinPropID.Malleable)
        '    Dim lDrvFrameSuperC As Int32 = DriveFrameMineral.GetPropertyValue(eMinPropID.SuperconductivePoint)
        '    Dim lDrvFrameSuperCKnow As Int32 = Owner.GetMineralPropertyKnowledge(lDriveFrameMineralID, eMinPropID.SuperconductivePoint)
        '    Dim lDrvFrameHard As Int32 = DriveFrameMineral.GetPropertyValue(eMinPropID.Hardness)
        '    Dim lFuelCompTempSens As Int32 = FuelCompMineral.GetPropertyValue(eMinPropID.TemperatureSensitivity)
        '    Dim lFuelCatThermCond As Int32 = FuelCatalystMineral.GetPropertyValue(eMinPropID.ThermalConductance)
        '    Dim lFuelCompComb As Int32 = FuelCompMineral.GetPropertyValue(eMinPropID.Combustiveness)
        '    Dim lFuelCatChemReact As Int32 = FuelCatalystMineral.GetPropertyValue(eMinPropID.ChemicalReactance)

        '    Dim bValid As Boolean = True

        '    If lDrvBodyMagReact = 0 Then
        '        bValid = False
        '        Me.ErrorReasonCode = TechBuilderErrorReason.eEngine_DriveBodyMagReactZero
        '    ElseIf lStrFrameMall = 0 Then
        '        bValid = False
        '        Me.ErrorReasonCode = TechBuilderErrorReason.eEngine_StructureFrameMalleableZero
        '    ElseIf lFuelCompComb = 0 Then
        '        bValid = False
        '        Me.ErrorReasonCode = TechBuilderErrorReason.eEngine_FuelCompositionCombustionZero
        '    ElseIf lFuelCatChemReact = 0 Then
        '        bValid = False
        '        Me.ErrorReasonCode = TechBuilderErrorReason.eEngine_FuelCatalystChemReactZero
        '    End If

        '    Me.CurrentSuccessChance = CInt(Math.Abs(Math.Abs(Me.GetPopIntel - _
        '      Math.Max(DriveMeldMineral.GetPropertyValue(eMinPropID.Complexity), _
        '      Math.Max(StructureBodyMineral.GetPropertyValue(eMinPropID.Complexity), _
        '      DriveBodyMineral.GetPropertyValue(eMinPropID.Complexity)))) + 1) + Me.RandomSeed * 5)

        '    With moResearchCost
        '        .ColonistCost = 0 'Me.CurrentSuccessChance
        '        .EnlistedCost = 0
        '        .OfficerCost = 0

        '        Try

        '            If bValid = True Then

        '                'ABS ( ABS ( StructBody.MagProdActual / DriveBody.MagReactActual - (( 255 - StructMeld.MeltingActual) * .001 )) + ((100 - DriveFrame.SuperPointActual) / 1000 ) )
        '                f_D2 = CSng(Math.Abs(Math.Abs(lStrBodyMagProd / lDrvBodyMagReact - ((255.0F - lStrMeldMeltPt) * 0.001F)) + ((100.0F - lDrvFrameSuperC) / 1000.0F)))

        '                '((StructBody.MagProdKnown + DriveBody.MagReactKnown) / 2) / .StartingSuccessChance
        '                f_G2 = ((lStrBodyMagProdKnow + lDrvBodyMagReactKnow) / 2.0F) / Me.CurrentSuccessChance

        '                If f_D2 = 0 OrElse f_G2 = 0 Then
        '                    bValid = False
        '                    Me.ErrorReasonCode = TechBuilderErrorReason.eEngine_MagneticsOutOfAlignment
        '                End If

        '                If bValid = True Then
        '                    'DesiredPower / G2 / D2
        '                    f_H2 = Me.PowerProd / f_G2 / f_D2

        '                    'ABS ( DriveFrame.HardActual / StructFrame.MalleableActual ) - ((StructMeld.MalleableActual - 255 ) * .001 ))
        '                    f_D3 = CSng(Math.Abs(lDrvFrameHard / lStrFrameMall) - ((lStrMeldMall - 255.0F) * 0.001F))

        '                    'DesiredThrust * D3
        '                    f_H3 = Me.Thrust * f_D3

        '                    f_E2 = Math.Abs(Math.Abs(GetPopIntel() - _
        '                      Math.Max(DriveMeldMineral.GetPropertyValue(eMinPropID.Complexity), _
        '                      Math.Max(StructureBodyMineral.GetPropertyValue(eMinPropID.Complexity), _
        '                      DriveBodyMineral.GetPropertyValue(eMinPropID.Complexity)))) + 1) + RandomSeed * 5
        '                    '( FuelComp.TempSensActual * FuelCat.ThermCondActual ) / ( FuelComp.CombustActual * FuelCat.ChemReactActual )

        '                    f_D4 = CSng((lFuelCompTempSens * lFuelCatThermCond) / (lFuelCompComb * lFuelCatChemReact))
        '                    f_H4 = mySpeed * f_E2 * f_D4

        '                    f_H5 = Me.Maneuver * (Me.Maneuver * Me.Thrust / (Math.Abs(195 - DriveFrameMineral.GetPropertyValue(eMinPropID.Hardness) / f_D3)))
        '                    'H5 = D5 * DesiredManeuver

        '                    '=(H2 + H3 + H4 + H5) / 4
        '                    .CreditCost = CInt((f_H2 + f_H3 + f_H4 + f_H5) / 4.0F)

        '                    '.PointsRequired = .CreditCost * 100
        '                    .PointsRequired = GetResearchPointsCost(.CreditCost * 100)
        '                    .CreditCost *= Me.ResearchAttempts
        '                    'TODO: Reset ResearchAttempts?

        '                    'Minerals...
        '                    .ItemCostUB = -1
        '                    ReDim .ItemCosts(.ItemCostUB)

        '                    Dim lTemp As Int32
        '                    HullRequired = 0
        '                    '  [Structural Alloy Body].Quantity = Ceiling(G2 * (DesiredPower / 100))
        '                    lTemp = CInt(Math.Ceiling(f_G2 * (Me.PowerProd / 100.0F)))
        '                    HullRequired += lTemp
        '                    .AddProductionCostItem(lStructuralBodyMineralID, ObjectType.eMineral, lTemp)

        '                    '  [Structural Alloy Frame].Quantity = Ceiling(DesiredThrust / .StartingSuccessChance)
        '                    lTemp = CInt(Math.Ceiling(Me.Thrust / Me.CurrentSuccessChance))
        '                    HullRequired += lTemp
        '                    .AddProductionCostItem(lStructuralFrameMineralID, ObjectType.eMineral, lTemp)

        '                    '  [Structural Alloy Meld].Quantity = DesiredThrust
        '                    .AddProductionCostItem(lStructuralMeldMineralID, ObjectType.eMineral, Me.Thrust \ 100)

        '                    '  [Drive Alloy Body].Quantity = Ceiling(DesiredPower / .StartingSuccessChance / 4)
        '                    lTemp = CInt(Math.Ceiling(Me.PowerProd / Me.CurrentSuccessChance / 4))
        '                    HullRequired += lTemp
        '                    .AddProductionCostItem(lDriveBodyMineralID, ObjectType.eMineral, lTemp)

        '                    '  [Drive Alloy Frame].Quantity = Ceiling(DesiredThrust / ABS (1 - D3) / 10)
        '                    lTemp = CInt(Math.Ceiling(Me.Thrust / Math.Abs(1 - f_D3) / 10))
        '                    HullRequired += lTemp
        '                    .AddProductionCostItem(lDriveFrameMineralID, ObjectType.eMineral, lTemp)

        '                    '  [Drive Alloy Meld].Quantity = Ceiling( H3 / .StartingSuccessChance )
        '                    .AddProductionCostItem(lDriveMeldMineralID, ObjectType.eMineral, CInt(Math.Ceiling(f_H3 / Me.CurrentSuccessChance)))

        '                    '  [Fuel Composition].Quantity = Ceiling( D4 )
        '                    .AddProductionCostItem(lFuelCompositionMineralID, ObjectType.eMineral, CInt(Math.Ceiling(f_D4)))

        '                    '  [Fuel Catalyst].Quantity = Ceiling( D5 / D3 )
        '                    .AddProductionCostItem(lFuelCatalystMineralID, ObjectType.eMineral, CInt(Math.Ceiling((Me.Maneuver * Me.Thrust / (Math.Abs(195 - DriveFrameMineral.GetPropertyValue(eMinPropID.Hardness)) / f_D3)) / f_D3)))

        '                    'Hull required...
        '                    HullRequired = CInt(HullRequired * RandomSeed)
        '                End If
        '            End If
        '        Catch
        '            bValid = False
        '            Me.ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
        '        End Try
        '    End With



        '    '===================================== BEGIN PRODUCTION COST!!! ===========================
        '    If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
        '    moProductionCost.ObjectID = Me.ObjectID
        '    moProductionCost.ObjTypeID = Me.ObjTypeID
        '    moProductionCost.ColonistCost = 0
        '    moProductionCost.EnlistedCost = 0
        '    moProductionCost.OfficerCost = 0
        '    moProductionCost.PointsRequired = 0
        '    Try
        '        With moProductionCost
        '            'Add Mineral costs...
        '            .ItemCostUB = -1
        '            ReDim .ItemCosts(.ItemCostUB)

        '            For X As Int32 = 0 To moResearchCost.ItemCostUB
        '                .AddProductionCostItem(moResearchCost.ItemCosts(X).ItemID, moResearchCost.ItemCosts(X).ItemTypeID, moResearchCost.ItemCosts(X).QuantityNeeded)
        '                .PointsRequired += .ItemCosts(X).QuantityNeeded
        '            Next X

        '            .PointsRequired *= 123
        '        End With

        '        If bValid = True Then
        '            moProductionCost.CreditCost = CInt(Math.Ceiling((f_H2 + f_H3 + f_H4 + f_H5) * (RandomSeed + 1))) * 50
        '        End If
        '    Catch
        '        bValid = False
        '        Me.ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
        '    End Try

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

    '	Dim bNotStudyFlaw As Boolean = False

    '	Try

    '		'D5
    '		Dim blMaxValue As Int64 = Math.Max((CLng(Me.PowerProd) * 100L), CLng(Me.Thrust) * 5L)
    '		blMaxValue = Math.Max(blMaxValue, CInt(Me.Speed) * CInt(Me.Speed) * CInt(Me.Speed))
    '		blMaxValue = Math.Max(blMaxValue, CInt(Me.Maneuver) * CInt(Me.Maneuver) * CInt(Me.Maneuver))

    '		'to keep track of the major design flaw
    '		Dim fHighestScore As Single = 0

    '		'======== Structural Body ===========
    '		Dim uSB_MagProd As MaterialPropertyItem2
    '		With uSB_MagProd
    '			.lMineralID = StructureBodyMineral.ObjectID
    '			.lPropertyID = eMinPropID.MagneticProduction

    '			.lActualValue = StructureBodyMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
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
    '		Dim uSB_Density As MaterialPropertyItem2
    '		With uSB_Density
    '			.lMineralID = StructureBodyMineral.ObjectID
    '			.lPropertyID = eMinPropID.Density

    '			.lActualValue = StructureBodyMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.8F
    '			.lGoalValue = 147 '200
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
    '		Dim uSB_MeltingPt As MaterialPropertyItem2
    '		With uSB_MeltingPt
    '			.lMineralID = StructureBodyMineral.ObjectID
    '			.lPropertyID = eMinPropID.MeltingPoint

    '			.lActualValue = StructureBodyMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 149 '200
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
    '		Dim lSBComplexity As Int32 = StructureBodyMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fSBInc As Single = CSng(lSBComplexity / Me.PopIntel)
    '		fSBInc *= (uSB_Density.FinalScore + uSB_MagProd.FinalScore + uSB_MeltingPt.FinalScore)
    '		'========= End of Structural Body =============

    '		'========= Structural Frame ============
    '		Dim uSF_Malleable As MaterialPropertyItem2
    '		With uSF_Malleable
    '			.lMineralID = StructureFrameMineral.ObjectID
    '			.lPropertyID = eMinPropID.Malleable

    '			.lActualValue = StructureFrameMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.3F
    '			.lGoalValue = 137 '200
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
    '		Dim uSF_Hardness As MaterialPropertyItem2
    '		With uSF_Hardness
    '			.lMineralID = StructureFrameMineral.ObjectID
    '			.lPropertyID = eMinPropID.Hardness

    '			.lActualValue = StructureFrameMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.6F
    '			.lGoalValue = 144 '200
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
    '		Dim uSF_Compress As MaterialPropertyItem2
    '		With uSF_Compress
    '			.lMineralID = StructureFrameMineral.ObjectID
    '			.lPropertyID = eMinPropID.Compressibility

    '			.lActualValue = StructureFrameMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 10
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
    '		Dim lSFComplexity As Int32 = StructureFrameMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fSFInc As Single = CSng(lSFComplexity / Me.PopIntel)
    '		fSFInc *= (uSF_Compress.FinalScore + uSF_Hardness.FinalScore + uSF_Malleable.FinalScore)
    '		'========= End of Structural Frame =====

    '		'========= Structural Meld =============
    '		Dim uSM_Malleable As MaterialPropertyItem2
    '		With uSM_Malleable
    '			.lMineralID = StructureMeldMineral.ObjectID
    '			.lPropertyID = eMinPropID.Malleable

    '			.lActualValue = StructureMeldMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.3F
    '			.lGoalValue = 10
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
    '		Dim uSM_Hardness As MaterialPropertyItem2
    '		With uSM_Hardness
    '			.lMineralID = StructureMeldMineral.ObjectID
    '			.lPropertyID = eMinPropID.Hardness

    '			.lActualValue = StructureMeldMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.6F
    '			.lGoalValue = 148 '200
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
    '		Dim uSM_MeltingPt As MaterialPropertyItem2
    '		With uSM_MeltingPt
    '			.lMineralID = StructureMeldMineral.ObjectID
    '			.lPropertyID = eMinPropID.MeltingPoint

    '			.lActualValue = StructureMeldMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 142 ' 200
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
    '		Dim lSMComplexity As Int32 = StructureMeldMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fSMInc As Single = CSng(lSMComplexity / Me.PopIntel)
    '		fSMInc *= (uSM_Hardness.FinalScore + uSM_Malleable.FinalScore + uSM_MeltingPt.FinalScore)
    '		'========= End of Structural meld ======

    '		'========= Drive Body ==================
    '		Dim uDB_MagReact As MaterialPropertyItem2
    '		With uDB_MagReact
    '			.lMineralID = DriveBodyMineral.ObjectID
    '			.lPropertyID = eMinPropID.MagneticReaction

    '			.lActualValue = DriveBodyMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.3F
    '			.lGoalValue = 168 '200
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
    '		Dim uDB_Compress As MaterialPropertyItem2
    '		With uDB_Compress
    '			.lMineralID = DriveBodyMineral.ObjectID
    '			.lPropertyID = eMinPropID.Compressibility

    '			.lActualValue = DriveBodyMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.6F
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
    '		Dim uDB_Density As MaterialPropertyItem2
    '		With uDB_Density
    '			.lMineralID = DriveBodyMineral.ObjectID
    '			.lPropertyID = eMinPropID.Density

    '			.lActualValue = DriveBodyMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
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
    '		Dim lDBComplexity As Int32 = DriveBodyMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fDBInc As Single = CSng(lDBComplexity / Me.PopIntel)
    '		fDBInc *= (uDB_Compress.FinalScore + uDB_Density.FinalScore + uDB_MagReact.FinalScore)
    '		'========= End of Drive Body ===========

    '		'========= Drive Frame =================
    '		Dim uDF_SuperC As MaterialPropertyItem2
    '		With uDF_SuperC
    '			.lMineralID = DriveFrameMineral.ObjectID
    '			.lPropertyID = eMinPropID.SuperconductivePoint

    '			.lActualValue = DriveFrameMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.3F
    '			.lGoalValue = 165 '200
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
    '		Dim uDF_Hardness As MaterialPropertyItem2
    '		With uDF_Hardness
    '			.lMineralID = DriveFrameMineral.ObjectID
    '			.lPropertyID = eMinPropID.Hardness

    '			.lActualValue = DriveFrameMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.6F
    '			.lGoalValue = 143 '200
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
    '		Dim uDF_Compress As MaterialPropertyItem2
    '		With uDF_Compress
    '			.lMineralID = DriveFrameMineral.ObjectID
    '			.lPropertyID = eMinPropID.Compressibility

    '			.lActualValue = DriveFrameMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 10
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
    '		Dim lDFComplexity As Int32 = DriveFrameMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fDFInc As Single = CSng(lDFComplexity / Me.PopIntel)
    '		fDFInc *= (uDF_Compress.FinalScore + uDF_Hardness.FinalScore + uDF_SuperC.FinalScore)
    '		'========== end of drive frame =========

    '		'========== Drive Meld =================
    '		Dim uDM_ChemReact As MaterialPropertyItem2
    '		With uDM_ChemReact
    '			.lMineralID = DriveMeldMineral.ObjectID
    '			.lPropertyID = eMinPropID.ChemicalReactance

    '			.lActualValue = DriveMeldMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.3F
    '			.lGoalValue = 147 '200
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat6_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat6_Prop1
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uDM_Malleable As MaterialPropertyItem2
    '		With uDM_Malleable
    '			.lMineralID = DriveMeldMineral.ObjectID
    '			.lPropertyID = eMinPropID.Malleable

    '			.lActualValue = DriveMeldMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.6F
    '			.lGoalValue = 10
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat6_Prop2 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat6_Prop2
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim uDM_Combust As MaterialPropertyItem2
    '		With uDM_Combust
    '			.lMineralID = DriveMeldMineral.ObjectID
    '			.lPropertyID = eMinPropID.Combustiveness

    '			.lActualValue = DriveMeldMineral.GetPropertyValue(.lPropertyID)
    '			.lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
    '			.fNormalize = 0.1F
    '			.lGoalValue = 10
    '			If .lActualValue <> .lKnownValue Then
    '				bNotStudyFlaw = True
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat6_Prop3 Or eComponentDesignFlaw.eShift_Not_Study
    '			End If
    '			If .FinalScore > fHighestScore AndAlso bNotStudyFlaw = False Then
    '				MajorDesignFlaw = eComponentDesignFlaw.eMat6_Prop3
    '				If .lActualValue < .lGoalValue Then MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eShift_Should_Be_Higher
    '				fHighestScore = .FinalScore
    '			End If
    '		End With
    '		Dim lDMComplexity As Int32 = DriveMeldMineral.GetPropertyValue(eMinPropID.Complexity)
    '		Dim fDMInc As Single = CSng(lDMComplexity / Me.PopIntel)
    '		fDMInc *= (uDM_ChemReact.FinalScore + uDM_Combust.FinalScore + uDM_Malleable.FinalScore)
    '		'========== end of drive meld ==========

    '		If fHighestScore < 10.0F AndAlso bNotStudyFlaw = False Then
    '			MajorDesignFlaw = MajorDesignFlaw Or eComponentDesignFlaw.eGoodDesign
    '		End If

    '		'============= beginning of the bottom of the sheet =============
    '		'First, modify our randomseed to be in the correct range
    '		Dim fModifiedSeed As Single = (RandomSeed * 0.7F) + 0.8F
    '		'Normalized value =IF(thrust>0,((thrust*maneuver*maxspeed*power))*B40/IF(thrust>0,VLOOKUP(thrust,lookupstuff,4,TRUE),14000000),1350000000)
    '		Dim blNormalized As Int64 = 0
    '		If Thrust > 0 Then
    '			blNormalized = CLng(Thrust) * CLng(Maneuver) * CLng(Speed) * CLng(PowerProd) * CLng(GunsPerHull())
    '			blNormalized = CLng(blNormalized / GetThrustNormalizer())
    '		Else : blNormalized = 1350000000
    '		End If
    '		'Hedge =IF(B41-1350000000<1,IF(thrust <> 0,POWER(B40,5),power*20),B41-1343000000)
    '		Dim blHedge As Int64 = 0
    '		If blNormalized - 1350000000 < 1 Then
    '			If Thrust <> 0 Then
    '				blHedge = CLng(Math.Pow(GunsPerHull(), 2)) * PowerProd
    '			Else : blHedge = CLng(PowerProd) * 20L
    '			End If
    '		Else : blHedge = blNormalized - 1350000000
    '		End If
    '		'AverageComplexity = AVERAGE(B37,B32,B27,B22,B17,B12)
    '		Dim fAvgComplexity As Single = lSBComplexity + lSFComplexity + lSMComplexity + lDBComplexity + lDFComplexity + lDMComplexity
    '		fAvgComplexity /= 6.0F
    '		'Starting success chance =(ABS(ROUNDDOWN(IF(thrust+maxspeed+maneuver = 0, -1*(power/10),B45/Intelligence*randomseed*(B42/1000000)),0))*-1)-1
    '		If Thrust = 0 AndAlso Speed = 0 AndAlso Maneuver = 0 Then
    '			Me.CurrentSuccessChance = -1 * (PowerProd \ 10)
    '		Else
    '			'B45/Intelligence*randomseed*(B42/1000000)
    '			Me.CurrentSuccessChance = CInt(Math.Floor(fAvgComplexity / PopIntel * fModifiedSeed * (blHedge / 1000000.0F)))
    '		End If
    '		Me.CurrentSuccessChance = (Math.Abs(Me.CurrentSuccessChance) * -1) - 1
    '		'Hull Required =IF(SUM(B2:B4) = 0,power/1000,1)*(1+ROUNDUP(0.2*thrust+(B42/10000)*SBInc,0))
    '		Me.HullRequired = CInt(1 + Math.Ceiling(0.2F * Thrust + (blHedge / 10000.0F) * fSBInc))
    '		If Thrust = 0 AndAlso Speed = 0 AndAlso Maneuver = 0 Then
    '			Me.HullRequired = CInt(Me.HullRequired * (PowerProd / 1000.0F))
    '		End If
    '		'Enlisted Required=ROUNDUP(SFInc/10*((power+thrust)/1000),0)
    '		Dim lEnlisted As Int32 = CInt(Math.Ceiling(fSFInc / 10.0F * ((CInt(PowerProd) + CInt(Thrust)) / 1000.0F)))
    '		'Officers required = rounddown(enlisted / 5)
    '		Dim lOfficers As Int32 = lEnlisted \ 5

    '		'Average AK (on sheet it is average ka but that is misnamed) = Average(AK Scores)
    '		Dim fAverageAK As Single = uSB_Density.AKScore + uSB_MagProd.AKScore + uSB_MeltingPt.AKScore
    '		fAverageAK += uSF_Compress.AKScore + uSF_Hardness.AKScore + uSF_Malleable.AKScore
    '		fAverageAK += uSM_Hardness.AKScore + uSM_Malleable.AKScore + uSM_MeltingPt.AKScore
    '		fAverageAK += uDB_Compress.AKScore + uDB_Density.AKScore + uDB_MagReact.AKScore
    '		fAverageAK += uDF_Compress.AKScore + uDF_Hardness.AKScore + uDF_SuperC.AKScore
    '		fAverageAK += uDM_ChemReact.AKScore + uDM_Combust.AKScore + uDM_Malleable.AKScore
    '		fAverageAK /= 18.0F

    '		'fIncSumItoac = SUM(DMInc,DFInc,DBInc,K22,SFInc,SBInc) * (B45/Intelligence)
    '		Dim fIncSumIToAC As Single = (fDMInc + fDFInc + fDBInc + fSMInc + fSFInc + fSBInc)
    '		fIncSumIToAC *= (fAvgComplexity / PopIntel)

    '		'Research Credits = 
    '		'IF( ROUNDUP ( fIncSumIToAC * E38 * randomseed,0) * 
    '		'fIncSumIToAC*E38*randomseed*IF(B42<10000,10000,1) * 
    '		'IF(thrust>0,1,0.1)+D5>10000000,
    '		'randomseed*ROUNDUP(fIncSumIToAC*E38*randomseed,0) * fIncSumIToAC * 
    '		'E38*randomseed*IF(B42<10000,10000,1)*IF(thrust>0,1,0.1)+D5,randomseed*10000000)
    '		Dim dblTemp As Double = Math.Ceiling(fIncSumIToAC * fAverageAK * fModifiedSeed) * fIncSumIToAC * fAverageAK * fModifiedSeed
    '		If blHedge < 10000 Then dblTemp *= 10000
    '		If Thrust < 1 Then dblTemp *= 0.1
    '		dblTemp += blMaxValue
    '		If dblTemp < 10000000 Then dblTemp = fModifiedSeed * 10000000
    '		Dim blResearchCredits As Int64 = CLng(dblTemp)

    '		'Research Credits =VLOOKUP(ROUNDUP(fIncSumItoAC *E38 * randomseed,0) ,lowlookup,5)
    '		' * fIncSumIToAC * E38 *randomseed * IF(B42<10000,10000,1)*IF(thrust>0,1,power/10)+D5
    '		'Dim dblTemp As Double = GetLowLookup(CInt(Math.Ceiling(fIncSumIToAC * fAverageAK * fModifiedSeed)), 3)
    '		'dblTemp *= fIncSumIToAC * fAverageAK * fModifiedSeed
    '		'If blHedge < 10000 Then dblTemp *= 10000
    '		'If Thrust = 0 Then dblTemp *= 0.1F
    '		'Dim blResearchCredits As Int64 = CLng(dblTemp) + blMaxValue

    '		'ResearchPoints = ABS(B42)*IF(B42<100000,100000,1)*AVERAGE(DMInc,DFInc,DBInc,K22,SFInc,SBInc)
    '		dblTemp = (fDMInc + fDFInc + fDBInc + fSMInc + fSFInc + fSBInc) / 6
    '		Dim blResearchPoints As Int64 = Math.Abs(blHedge)
    '		If blHedge < 100000 Then blResearchPoints *= 100000
    '		blResearchPoints = CLng(blResearchPoints * dblTemp)

    '		'MSC - 1/07/08 - put in to make lesser units easier on players (especially nubs)
    '		Dim blCheckVal As Int64 = CLng(Thrust) + CLng(PowerProd) + CLng(Speed) + CLng(Maneuver)
    '		If blCheckVal <= 400 Then
    '			If blResearchPoints > 476928000 Then		'2 days
    '				blResearchPoints = CLng(317952000 * fModifiedSeed)		'so that max is 2 days
    '			End If
    '		End If

    '		blResearchPoints = Math.Max(blResearchPoints, CLng(180000 * fModifiedSeed))

    '		'ProductionCredits=IF(thrust>0,ABS(B46*1000000)*randomseed,ABS(B46*randomseed*B45/Intelligence)*1000)
    '		Dim blProductionCredits As Int64 = 0
    '		If Thrust > 0 Then
    '			blProductionCredits = CLng(Math.Abs(Me.CurrentSuccessChance * 1000000) * fModifiedSeed)
    '		Else
    '			blProductionCredits = CLng(Math.Abs(Me.CurrentSuccessChance * fModifiedSeed * fAvgComplexity / PopIntel) * 1000)
    '		End If

    '		'Ok, slight, detour, do the element costs
    '		Dim lSBCost As Int32 = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fSBInc))	 '=ROUNDUP(B47*SBInc,0)
    '		Dim lSFCost As Int32 = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fSFInc))	'=ROUNDUP(B47*0.2*SFInc,0)
    '		Dim lSMCost As Int32 = CInt(Math.Ceiling(Me.HullRequired * 0.05F * fSMInc))	'=ROUNDUP(B47*0.05*K22,0)
    '		Dim lDBCost As Int32 = CInt(Math.Ceiling(lSBCost * 0.05F * fDBInc))	'=ROUNDUP(I41*0.8*DBInc,0)
    '		Dim lDFCost As Int32 = CInt(Math.Ceiling(lSFCost * 0.05F * fDFInc))	'=ROUNDUP(I42*0.8*DFInc,0)
    '		Dim lDMCost As Int32 = CInt(Math.Ceiling(lSMCost * 0.05F * fDMInc))	 '=ROUNDUP(I43*DMInc,0)

    '		lSBCost = Math.Max(lSBCost \ 10, 1) : lSFCost = Math.Max(lSFCost \ 10, 1) : lSMCost = Math.Max(lSMCost \ 10, 1)
    '		lDBCost = Math.Max(lDBCost \ 10, 1) : lDFCost = Math.Max(lDFCost \ 10, 1) : lDMCost = Math.Max(lDMCost \ 10, 1)

    '		'ProductionTime=IF(B47>400000,0.3,1)*SUM(I41:I46)*randomseed*B47
    '		Dim blProductionPoints As Int64 = CLng(lSBCost) + CLng(lSFCost) + CLng(lSMCost) + CLng(lDBCost) + CLng(lDFCost) + CLng(lDMCost)
    '		blProductionPoints *= Me.HullRequired
    '		If HullRequired > 100000 Then blProductionPoints = CLng(blProductionPoints * 0.3F)
    '		blProductionPoints = CLng(blProductionPoints * fModifiedSeed)

    '		blProductionPoints = Math.Max(blProductionPoints, CLng(900000 * fModifiedSeed))


    '		'Now, we fill in our research and production cost...
    '		If moResearchCost Is Nothing Then moResearchCost = New ProductionCost
    '		With moResearchCost
    '			.ColonistCost = 0
    '			.CreditCost = blResearchCredits
    '			.EnlistedCost = 0
    '			.ObjectID = Me.ObjectID
    '			.ObjTypeID = Me.ObjTypeID
    '			.OfficerCost = 0
    '			.PointsRequired = GetResearchPointsCost(blResearchPoints)
    '			'.CreditCost *= Me.ResearchAttempts

    '			Erase .ItemCosts
    '			.ItemCostUB = -1
    '			.AddProductionCostItem(lStructuralBodyMineralID, ObjectType.eMineral, lSBCost * 5)
    '			.AddProductionCostItem(lStructuralFrameMineralID, ObjectType.eMineral, lSFCost * 5)
    '			.AddProductionCostItem(lStructuralMeldMineralID, ObjectType.eMineral, lSMCost * 5)
    '			.AddProductionCostItem(lDriveBodyMineralID, ObjectType.eMineral, lDBCost * 5)
    '			.AddProductionCostItem(lDriveFrameMineralID, ObjectType.eMineral, lDFCost * 5)
    '			.AddProductionCostItem(lDriveMeldMineralID, ObjectType.eMineral, lDMCost * 5)
    '		End With
    '		If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
    '		With moProductionCost
    '			.ColonistCost = 0
    '			.CreditCost = blProductionCredits
    '			.EnlistedCost = lEnlisted
    '			.ObjectID = Me.ObjectID
    '			.ObjTypeID = Me.ObjTypeID
    '			.OfficerCost = lOfficers
    '			.PointsRequired = blProductionPoints

    '			Erase .ItemCosts
    '			.ItemCostUB = -1
    '			.AddProductionCostItem(lStructuralBodyMineralID, ObjectType.eMineral, lSBCost)
    '			.AddProductionCostItem(lStructuralFrameMineralID, ObjectType.eMineral, lSFCost)
    '			.AddProductionCostItem(lStructuralMeldMineralID, ObjectType.eMineral, lSMCost)
    '			.AddProductionCostItem(lDriveBodyMineralID, ObjectType.eMineral, lDBCost)
    '			.AddProductionCostItem(lDriveFrameMineralID, ObjectType.eMineral, lDFCost)
    '			.AddProductionCostItem(lDriveMeldMineralID, ObjectType.eMineral, lDMCost)
    '		End With
    '	Catch
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

            Dim oComputer As New EngineTechComputer
            With oComputer

                .blMan = Maneuver
                .blMaxSpd = Speed
                .blPowGen = PowerProd
                .blThrustGen = Thrust
                .lHullTypeID = HullTypeID

                .lMineral1ID = lStructuralBodyMineralID
                .lMineral2ID = lStructuralFrameMineralID
                .lMineral3ID = lStructuralMeldMineralID
                .lMineral4ID = lDriveBodyMineralID
                .lMineral5ID = lDriveFrameMineralID
                .lMineral6ID = lDriveMeldMineralID

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
                    .AddProductionCostItem(oComputer.lMineral6ID, ObjectType.eMineral, oComputer.lResultMin6)
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
                    .AddProductionCostItem(oComputer.lMineral6ID, ObjectType.eMineral, oComputer.lResultMin6)

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


#Region "  Lookup Stuff  "
    Private Function GunsPerHull() As Int32
        If Thrust = 0 AndAlso Speed = 0 AndAlso Maneuver = 0 Then
            Return 5
        ElseIf Thrust < 91 Then
            Return 2
        ElseIf Thrust < 121 Then
            Return 5
            'ElseIf Thrust < 171 Then
            '	Return 3
            'ElseIf Thrust < 201 Then
            '	Return 1
        ElseIf Thrust < 201 Then
            Return 3
        ElseIf Thrust < 301 Then
            Return 4
        ElseIf Thrust < 501 Then
            Return 12
        ElseIf Thrust < 2001 Then
            Return 8
        ElseIf Thrust < 8001 Then
            Return 12
        ElseIf Thrust < 20001 Then
            Return 15
        ElseIf Thrust < 40001 Then
            Return 20
        ElseIf Thrust < 80001 Then
            Return 20
        ElseIf Thrust < 250000 Then
            Return 25
        ElseIf Thrust < 1000000 Then
            Return 20
        Else : Return 40
        End If
    End Function
    Private Function GetThrustNormalizer() As Single
        If Thrust > 0 Then
            If Thrust < 91 Then
                Return 2
            ElseIf Thrust < 121 Then
                Return 1.33333F
            ElseIf Thrust < 171 Then
                Return 1.88889F
            ElseIf Thrust < 201 Then
                Return 2.22222F
            ElseIf Thrust < 301 Then
                Return 3.33333F
            ElseIf Thrust < 501 Then
                Return 5.55556F
            ElseIf Thrust < 2001 Then
                Return 22.2222F
            ElseIf Thrust < 8001 Then
                Return 55.5556F
            ElseIf Thrust < 20001 Then
                Return 111.1111F
            ElseIf Thrust < 40001 Then
                Return 2222.222F
            ElseIf Thrust < 80001 Then
                Return 4444.444F
            ElseIf Thrust < 250001 Then
                Return 6666.667F
            ElseIf Thrust < 1000000 Then
                Return 11111.11F
            Else : Return 44444.44F
            End If
        Else
            Return 14000000
        End If
    End Function
#End Region

    Protected Overrides Sub FinalizeResearch()
        'Dim fMult As Single = 1.0F + (Owner.oSpecials.yMaxSpeedBonus / 100.0F)
        'Dim lValue As Int32 = CInt(Math.Ceiling(Speed * fMult))
        If Me.HullTypeID <> eyHullType.Facility AndAlso Me.HullTypeID <> eyHullType.SpaceStation Then
            Dim lValue As Int32 = CInt(Speed) + CInt(Owner.oSpecials.yMaxSpeedBonus)
            If lValue > 255 Then lValue = 255
            If lValue < 0 Then lValue = 0
            Speed = CByte(lValue)

            'fMult = 1.0F + (Owner.oSpecials.yManeuverBonus / 100.0F)
            'lValue = CInt(Math.Ceiling(Maneuver * fMult))
            lValue = CInt(Maneuver) + CInt(Owner.oSpecials.yManeuverBonus)
            If lValue > 255 Then lValue = 255
            If lValue < 0 Then lValue = 0
            Maneuver = CByte(lValue)
        End If

        Dim fMult As Single = 1.0F + (Owner.oSpecials.yPowerBonus / 100.0F)
        PowerProd = CInt(Math.Ceiling(PowerProd * fMult))
    End Sub

    Public Function GetMaxProperty(ByVal lPropID As Int32) As Int32
        Return GetMaxVal(DriveMeldMineral.GetPropertyValue(lPropID), StructureBodyMineral.GetPropertyValue(lPropID), _
          DriveBodyMineral.GetPropertyValue(lPropID), StructureMeldMineral.GetPropertyValue(lPropID), _
          DriveFrameMineral.GetPropertyValue(lPropID), StructureFrameMineral.GetPropertyValue(lPropID))
    End Function

    Public Overrides Function GetNonOwnerMsg() As Byte()
        Dim yResult(42) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yResult, lPos) : lPos += 2
        Me.GetGUIDAsString.CopyTo(yResult, lPos) : lPos += 6

        EngineName.CopyTo(yResult, lPos) : lPos += 20
        System.BitConverter.GetBytes(Thrust).CopyTo(yResult, lPos) : lPos += 4
        yResult(lPos) = Maneuver : lPos += 1
        yResult(lPos) = Speed : lPos += 1
        System.BitConverter.GetBytes(PowerProd).CopyTo(yResult, lPos) : lPos += 4
        System.BitConverter.GetBytes(HullRequired).CopyTo(yResult, lPos) : lPos += 4
        yResult(lPos) = HullTypeID : lPos += 1

        Return yResult
    End Function

    Public Overrides Function TechnologyScore() As Integer
        If mlStoredTechScore = Int32.MinValue Then
            Try
                Dim lTemp As Int32 = (CInt(Maneuver) + CInt(Speed)) * (Me.Thrust \ 1000I) * (Me.PowerProd \ 1000I)
                If Me.HullRequired < 1 Then
                    mlStoredTechScore = lTemp \ 1000I
                Else : mlStoredTechScore = (lTemp \ Me.HullRequired) \ 1000I
                End If
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
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eNameOnly Then lPos += 6
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then lPos += 22
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then lPos += 100 '24
        ReDim yMsg(lPos - 1)
        lPos = 0

        'Default Attributes
        GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        Me.EngineName.CopyTo(yMsg, lPos) : lPos += 20

        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
            'at least settings 1 
            System.BitConverter.GetBytes(HullRequired).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = ColorValue : lPos += 1
            yMsg(lPos) = HullTypeID : lPos += 1
        End If
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
            'at least settings 2  
            System.BitConverter.GetBytes(Thrust).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = Maneuver : lPos += 1
            yMsg(lPos) = Speed : lPos += 1
            System.BitConverter.GetBytes(PowerProd).CopyTo(yMsg, lPos) : lPos += 4
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
            System.BitConverter.GetBytes(lStructuralBodyMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lStructuralFrameMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lStructuralMeldMineralID).CopyTo(yMsg, lPos) : lPos += 4

            System.BitConverter.GetBytes(lDriveBodyMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lDriveFrameMineralID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lDriveMeldMineralID).CopyTo(yMsg, lPos) : lPos += 4

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
            ReDim .EngineName(19)
            Array.Copy(yData, lPos, .EngineName, 0, 20) : lPos += 20
            .Thrust = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .Maneuver = yData(lPos) : lPos += 1
            .Speed = yData(lPos) : lPos += 1
            .PowerProd = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            .lStructuralBodyMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lStructuralFrameMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lStructuralMeldMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lDriveBodyMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lDriveFrameMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lDriveMeldMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            .HullRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ColorValue = yData(lPos) : lPos += 1

            .HullTypeID = yData(lPos) : lPos += 1

            If .Owner Is Nothing = False Then
                Dim lFirstIdx As Int32 = -1
                For X As Int32 = 0 To .Owner.mlEngineUB
                    If .Owner.mlEngineIdx(X) = .ObjectID Then
                        .Owner.moEngine(X) = Me
                        Return
                    ElseIf lFirstIdx = -1 AndAlso .Owner.mlEngineIdx(X) = -1 Then
                        lFirstIdx = X
                    End If
                Next X

                If lFirstIdx = -1 Then
                    lFirstIdx = .Owner.mlEngineUB + 1
                    ReDim Preserve .Owner.mlEngineIdx(lFirstIdx)
                    ReDim Preserve .Owner.moEngine(lFirstIdx)
                    .Owner.mlEngineUB = lFirstIdx
                End If
                .Owner.moEngine(lFirstIdx) = Me
                .Owner.mlEngineIdx(lFirstIdx) = Me.ObjectID
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

        Dim oComputer As New EngineTechComputer
        With oComputer

            .blMan = Maneuver
            .blMaxSpd = Speed
            .blPowGen = PowerProd
            .blThrustGen = Thrust
            .lHullTypeID = HullTypeID

            .lMineral1ID = lStructuralBodyMineralID
            .lMineral2ID = lStructuralFrameMineralID
            .lMineral3ID = lStructuralMeldMineralID
            .lMineral4ID = lDriveBodyMineralID
            .lMineral5ID = lDriveFrameMineralID
            .lMineral6ID = lDriveMeldMineralID

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
