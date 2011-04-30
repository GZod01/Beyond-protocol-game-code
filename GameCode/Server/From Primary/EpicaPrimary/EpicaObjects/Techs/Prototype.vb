Public Class Prototype
    Inherits Epica_Tech

    Private Const ml_NOISE_COST_MULT As Int32 = 1231

    Private Structure WeaponPlacement
        Public PW_ID As Int32
        Public WeaponTechID As Int32
        Public SlotX As Byte
        Public SlotY As Byte
        Public WeaponGroupTypeID As WeaponGroupType

        Public ArcID As UnitArcs

        Public Function SaveObject(ByVal lPrototypeID As Int32) As Boolean
            Dim bResult As Boolean = False
            Dim sSQL As String
            Dim oComm As OleDb.OleDbCommand = Nothing

            'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

            Try
                If PW_ID < 1 Then
                    'INSERT
                    sSQL = "INSERT INTO tblPrototypeWeapon (PrototypeID, WeaponID, SlotX, SlotY, ArcID, WeaponGroupTypeID) VALUES (" & _
                      lPrototypeID & ", " & WeaponTechID & ", " & SlotX & ", " & SlotY & ", " & CByte(ArcID) & ", " & _
                      CByte(WeaponGroupTypeID) & ")"
                Else
                    'UPDATE
                    sSQL = "UPDATE tblPrototypeWeapon SET PrototypeID = " & lPrototypeID & ", WeaponID = " & WeaponTechID & _
                      ", SlotX = " & SlotX & ", SlotY = " & SlotY & ", ArcID = " & CByte(ArcID) & ", WeaponGroupTypeID = " & _
                      CByte(WeaponGroupTypeID) & " WHERE PW_ID = " & PW_ID
                End If

                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Err.Raise(-1, "SaveObject", "No records affected when saving object!")
                End If
                If PW_ID < 1 Then
                    Dim oData As OleDb.OleDbDataReader
                    oComm = Nothing
                    sSQL = "SELECT MAX(PW_ID) FROM tblPrototypeWeapon WHERE PrototypeID = " & lPrototypeID
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oData = oComm.ExecuteReader(CommandBehavior.Default)
                    If oData.Read Then
                        PW_ID = CInt(oData(0))
                    End If
                    oData.Close()
                    oData = Nothing
                    oComm.Dispose()
                    oComm = Nothing
                End If

                bResult = True
            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "Unable to save PrototypeWeapon: " & ex.Message)
            Finally
                If oComm Is Nothing = False Then oComm.Dispose()
                oComm = Nothing
            End Try

            Return bResult
        End Function
        Public Function GetSaveObjectText(ByVal lPrototypeID As Int32) As String
            Dim sSQL As String

            Try
                If PW_ID < 1 Then
                    'INSERT
                    sSQL = "INSERT INTO tblPrototypeWeapon (PrototypeID, WeaponID, SlotX, SlotY, ArcID, WeaponGroupTypeID) VALUES (" & _
                      lPrototypeID & ", " & WeaponTechID & ", " & SlotX & ", " & SlotY & ", " & CByte(ArcID) & ", " & _
                      CByte(WeaponGroupTypeID) & ")"
                Else
                    'UPDATE
                    sSQL = "UPDATE tblPrototypeWeapon SET PrototypeID = " & lPrototypeID & ", WeaponID = " & WeaponTechID & _
                      ", SlotX = " & SlotX & ", SlotY = " & SlotY & ", ArcID = " & CByte(ArcID) & ", WeaponGroupTypeID = " & _
                      CByte(WeaponGroupTypeID) & " WHERE PW_ID = " & PW_ID
                End If
                Return sSQL

            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "Unable to save PrototypeWeapon: " & ex.Message)
            End Try

            Return ""
        End Function
    End Structure

    Public PrototypeName(19) As Byte          'takes from oHullTech.HullName by default

    Public lEngineTech As Int32
    Public lArmorTech As Int32
    Public lHullTech As Int32
    Public lRadarTech As Int32
    Public lShieldTech As Int32

    Public ForeArmorUnits As Int32
    Public AftArmorUnits As Int32
    Public LeftArmorUnits As Int32
    Public RightArmorUnits As Int32
    Public ProductionHull As Int32
    Public MaxCrew As Int32 = 0
    'Public MinCrew As Int32

    Private moWeapons() As WeaponPlacement
    Public WeaponUB As Int32 = -1

    Public moExpectedDef As Epica_Entity_Def = Nothing

    Public ResultingDef As Epica_Entity_Def = Nothing

    Private mlNoiseCnt As Int32 = -1

    Private mySendString() As Byte

    Public Function GetObjAsString() As Byte()

        'here we will return the entire object as a string
        'If mbStringReady = False Then

        Dim yDef() As Byte = Nothing

        If Me.ComponentDevelopmentPhase > eComponentDevelopmentPhase.eComponentDesign Then
            If moExpectedDef Is Nothing Then
                ComponentDesigned()
            End If
            yDef = moExpectedDef.GetObjAsString()

            ReDim mySendString(Epica_Tech.BASE_OBJ_STRING_SIZE_EXCLUDED + 71 + ((WeaponUB + 1) * 7) + (yDef.Length))
        Else
            ReDim mySendString(Epica_Tech.BASE_OBJ_STRING_SIZE_EXCLUDED + 71 + ((WeaponUB + 1) * 7))
        End If


        Dim lPos As Int32

        MyBase.GetBaseObjAsString(True).CopyTo(mySendString, 0)
        lPos = Epica_Tech.BASE_OBJ_STRING_SIZE_EXCLUDED

        PrototypeName.CopyTo(mySendString, lPos) : lPos += 20
        System.BitConverter.GetBytes(lEngineTech).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lArmorTech).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(lHangarTech).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lHullTech).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lRadarTech).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lShieldTech).CopyTo(mySendString, lPos) : lPos += 4

        System.BitConverter.GetBytes(ForeArmorUnits).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(AftArmorUnits).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(LeftArmorUnits).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(RightArmorUnits).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(ProductionHull).CopyTo(mySendString, lPos) : lPos += 4

        System.BitConverter.GetBytes(MaxCrew).CopyTo(mySendString, lPos) : lPos += 4
        'System.BitConverter.GetBytes(MinCrew).CopyTo(mySendString, lPos) : lPos += 2
        'System.BitConverter.GetBytes(MaxCrew).CopyTo(mySendString, lPos) : lPos += 2

        System.BitConverter.GetBytes((WeaponUB + 1I)).CopyTo(mySendString, lPos) : lPos += 4

        For X As Int32 = 0 To WeaponUB
            With moWeapons(X)
                System.BitConverter.GetBytes(.WeaponTechID).CopyTo(mySendString, lPos) : lPos += 4
                mySendString(lPos) = .SlotX : lPos += 1
                mySendString(lPos) = .SlotY : lPos += 1
                mySendString(lPos) = .WeaponGroupTypeID : lPos += 1
            End With
        Next X

        If moExpectedDef Is Nothing = False AndAlso moExpectedDef.ObjTypeID = ObjectType.eFacilityDef Then
            System.BitConverter.GetBytes(CType(moExpectedDef, FacilityDef).WorkerFactor).CopyTo(mySendString, lPos) : lPos += 4
        Else : System.BitConverter.GetBytes(0I).CopyTo(mySendString, lPos) : lPos += 4
        End If

        If Me.ComponentDevelopmentPhase > eComponentDevelopmentPhase.eComponentDesign AndAlso yDef Is Nothing = False Then
            yDef.CopyTo(mySendString, lPos)
        End If

        'mbStringReady = True
        'End If
        Return mySendString
    End Function

    Public Sub AddWeaponPlacement(ByVal lWeaponTechID As Int32, ByVal ySlotX As Byte, ByVal ySlotY As Byte, ByVal eWeaponGroupTypeID As WeaponGroupType)
        WeaponUB += 1
        ReDim Preserve moWeapons(WeaponUB)
        With moWeapons(WeaponUB)
            .PW_ID = -1
            .WeaponTechID = lWeaponTechID
            .SlotX = ySlotX
            .SlotY = ySlotY
            .WeaponGroupTypeID = eWeaponGroupTypeID


            If oHullTech Is Nothing = False Then
                Dim lTmpArcID As HullTech.SlotType = oHullTech.GetWeaponsMainSlotType(oHullTech.GetSlotGrpNum(ySlotX, ySlotY)) 'oHullTech.GetSlotType(ySlotX, ySlotY)
                Select Case lTmpArcID
                    Case HullTech.SlotType.eAllArc
                        .ArcID = UnitArcs.eAllArcs
                    Case HullTech.SlotType.eFront
                        .ArcID = UnitArcs.eForwardArc
                    Case HullTech.SlotType.eLeft
                        .ArcID = UnitArcs.eLeftArc
                    Case HullTech.SlotType.eRear
                        .ArcID = UnitArcs.eBackArc
                    Case HullTech.SlotType.eRight
                        .ArcID = UnitArcs.eRightArc
                    Case Else
                        .ArcID = UnitArcs.eForwardArc
                End Select
            End If
        End With

    End Sub

    Public Sub AddWeaponPlacement(ByVal lWeaponTechID As Int32, ByVal ySlotX As Byte, ByVal ySlotY As Byte, ByVal eWeaponGroupTypeID As WeaponGroupType, ByVal eArcID As UnitArcs)
        WeaponUB += 1
        ReDim Preserve moWeapons(WeaponUB)
        With moWeapons(WeaponUB)
            .PW_ID = -1
            .WeaponTechID = lWeaponTechID
            .SlotX = ySlotX
            .SlotY = ySlotY
            .WeaponGroupTypeID = eWeaponGroupTypeID
            .ArcID = eArcID
        End With
    End Sub

    Public Overrides Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblPrototype (PrototypeName, OwnerID, EngineID, ArmorID, ForeArmorUnits, " & _
                  "AftArmorUnits, LeftArmorUnits, RightArmorUnits, HullID, RadarID, ShieldID, MaxCrew, " & _
                  "MinCrew, CurrentSuccessChance, SuccessChanceIncrement, ResearchAttempts, RandomSeed, ResPhase, " & _
                  "ProductionHull, PopIntel, bArchived, VersionNumber) VALUES ('" & MakeDBStr(BytesToString(PrototypeName)) & "', " & Owner.ObjectID & ", "
                If oEngineTech Is Nothing = False Then
                    sSQL = sSQL & oEngineTech.ObjectID & ", "
                Else : sSQL = sSQL & "-1, "
                End If
                If oArmorTech Is Nothing = False Then
                    sSQL = sSQL & oArmorTech.ObjectID & ", "
                Else : sSQL = sSQL & "-1, "
                End If
                sSQL = sSQL & ForeArmorUnits & ", " & AftArmorUnits & ", " & LeftArmorUnits & ", " & RightArmorUnits & ", "

                'If oHangarTech Is Nothing = False Then
                '    sSQL = sSQL & oHangarTech.ObjectID & ", "
                'Else : sSQL = sSQL & "-1, "
                'End If
                'sSQL &= "-1, "

                sSQL = sSQL & oHullTech.ObjectID & ", "
                If oRadarTech Is Nothing = False Then
                    sSQL = sSQL & oRadarTech.ObjectID & ", "
                Else : sSQL = sSQL & "-1, "
                End If
                If oShieldTech Is Nothing = False Then
                    sSQL = sSQL & oShieldTech.ObjectID & ", "
                Else : sSQL = sSQL & "-1, "
                End If
                sSQL = sSQL & MaxCrew & ", 0, " & CurrentSuccessChance & ", " & SuccessChanceIncrement & ", " & ResearchAttempts & _
                  ", " & RandomSeed & ", " & ComponentDevelopmentPhase & ", " & ProductionHull & ", " & PopIntel & ", " & yArchived & ", " & lVersionNum & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblPrototype SET PrototypeName = '" & MakeDBStr(BytesToString(PrototypeName)) & "', OwnerID=" & Owner.ObjectID & _
                  ", EngineID = "
                If oEngineTech Is Nothing = False Then
                    sSQL = sSQL & oEngineTech.ObjectID & ", ArmorID="
                Else : sSQL = sSQL & "-1, ArmorID="
                End If
                If oArmorTech Is Nothing = False Then
                    sSQL = sSQL & oArmorTech.ObjectID & ", ForeArmorUnits="
                Else : sSQL = sSQL & "-1, ForeArmorUnits="
                End If
                sSQL = sSQL & ForeArmorUnits & ", AftArmorUnits=" & AftArmorUnits & ", LeftArmorUnits=" & _
                  LeftArmorUnits & ", RightArmorUnits = " & RightArmorUnits & ", HullID = "

                'If oHangarTech Is Nothing = False Then
                '    sSQL = sSQL & oHangarTech.ObjectID & ", HullID = "
                'Else : sSQL = sSQL & "-1, HullID ="
                'End If
                'sSQL &= "-1, HullID ="

                sSQL = sSQL & oHullTech.ObjectID & ", RadarID = "
                If oRadarTech Is Nothing = False Then
                    sSQL = sSQL & oRadarTech.ObjectID & ", ShieldID = "
                Else : sSQL = sSQL & "-1, ShieldID = "
                End If
                If oShieldTech Is Nothing = False Then
                    sSQL = sSQL & oShieldTech.ObjectID & ", MaxCrew ="
                Else : sSQL = sSQL & "-1, MaxCrew ="
                End If
                sSQL = sSQL & MaxCrew & ", MinCrew = 0, CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
                  SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & _
                  RandomSeed & ", ResPhase = " & ComponentDevelopmentPhase & ", ProductionHull = " & ProductionHull & _
                  ", PopIntel = " & PopIntel & ", bArchived = " & yArchived & ", VersionNumber = " & lVersionNum & " WHERE PrototypeID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(PrototypeID) FROM tblPrototype WHERE PrototypeName = '" & MakeDBStr(BytesToString(PrototypeName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
                oComm.Dispose()
                oComm = Nothing
            End If

            'TODO: This is a quick fix... PW_ID should be getting set on the load in GlobalVars but it isn't
            sSQL = "DELETE FROM tblPrototypeWeapon WHERE PrototypeID = " & Me.ObjectID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing

            'Save the PrototypeWeapons
            For X As Int32 = 0 To WeaponUB
                moWeapons(X).PW_ID = -1
                If moWeapons(X).SaveObject(Me.ObjectID) = False Then
                    'TODO: What should we do???
                End If
            Next X

            'MyBase.FinalizeSave()

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
            sSQL = "UPDATE tblPrototype SET PrototypeName = '" & MakeDBStr(BytesToString(PrototypeName)) & "', OwnerID=" & Owner.ObjectID & _
              ", EngineID = "
            If oEngineTech Is Nothing = False Then
                sSQL = sSQL & oEngineTech.ObjectID & ", ArmorID="
            Else : sSQL = sSQL & "-1, ArmorID="
            End If
            If oArmorTech Is Nothing = False Then
                sSQL = sSQL & oArmorTech.ObjectID & ", ForeArmorUnits="
            Else : sSQL = sSQL & "-1, ForeArmorUnits="
            End If
            sSQL = sSQL & ForeArmorUnits & ", AftArmorUnits=" & AftArmorUnits & ", LeftArmorUnits=" & _
              LeftArmorUnits & ", RightArmorUnits = " & RightArmorUnits & ", HullID = "
            sSQL = sSQL & oHullTech.ObjectID & ", RadarID = "
            If oRadarTech Is Nothing = False Then
                sSQL = sSQL & oRadarTech.ObjectID & ", ShieldID = "
            Else : sSQL = sSQL & "-1, ShieldID = "
            End If
            If oShieldTech Is Nothing = False Then
                sSQL = sSQL & oShieldTech.ObjectID & ", MaxCrew ="
            Else : sSQL = sSQL & "-1, MaxCrew ="
            End If
            sSQL = sSQL & MaxCrew & ", MinCrew = 0, CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
              SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & _
              RandomSeed & ", ResPhase = " & ComponentDevelopmentPhase & ", ProductionHull = " & ProductionHull & _
              ", PopIntel = " & PopIntel & ", bArchived = " & yArchived & " WHERE PrototypeID = " & ObjectID

            oSB.AppendLine(sSQL)

            'TODO: This is a quick fix... PW_ID should be getting set on the load in GlobalVars but it isn't
            sSQL = "DELETE FROM tblPrototypeWeapon WHERE PrototypeID = " & Me.ObjectID
            oSB.AppendLine(sSQL)

            'Save the PrototypeWeapons
            For X As Int32 = 0 To WeaponUB
                moWeapons(X).PW_ID = -1
                oSB.AppendLine(moWeapons(X).GetSaveObjectText(Me.ObjectID))
            Next X

            Return oSB.ToString
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        End Try
        Return ""
    End Function

    Public Overrides Function ValidateDesign() As Boolean
        Dim lPowerUsed As Int32 = 0
        Dim lEnlisted As Int32 = 0
        Dim lOfficers As Int32 = 0
        Dim lColonists As Int32 = 0

        'First, check hull, every prototype requires a hulltech
        If oHullTech Is Nothing Then Return False
        If oHullTech.GetProductionCost Is Nothing Then Return False
        With oHullTech.GetProductionCost
            lEnlisted += .EnlistedCost
            lOfficers += .OfficerCost
            lColonists += .ColonistCost
        End With

        Dim yHullTypeID As Byte = HullTech.GetHullTypeID(oHullTech.yTypeID, oHullTech.ySubTypeID)

        'Now, check our engine...
        If oEngineTech Is Nothing = False Then
            If oHullTech.yTypeID = 2 Then
                If oEngineTech.Speed <> 0 OrElse oEngineTech.Thrust <> 0 OrElse oEngineTech.Maneuver <> 0 Then
                    LogEvent(LogEventType.PossibleCheat, "Prototype Designed with moving engine on a facility. OwnerID: " & Me.Owner.ObjectID)
                    Return False
                End If
            End If

            If oEngineTech.HullTypeID <> yHullTypeID AndAlso oEngineTech.HullTypeID <> 255 Then
                LogEvent(LogEventType.PossibleCheat, "Prototype designed with invalid engine for hull. OwnerID: " & Me.Owner.ObjectID)
                Return False
            End If

            'Ok, there IS an engine... check hull requirements
            If oEngineTech.HullRequired > oHullTech.GetHullAllocation(HullTech.SlotConfig.eEngines) Then
                LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
            If oEngineTech.GetProductionCost Is Nothing Then Return False
            With oEngineTech.GetProductionCost
                lEnlisted += .EnlistedCost
                lOfficers += .OfficerCost
                lColonists += .ColonistCost
            End With
        End If

        'Get our hull allotments for armor per side
        Dim lFrontArmorHull As Int32 = oHullTech.GetSideHullAllocation(HullTech.SlotType.eFront, HullTech.SlotConfig.eArmorConfig)
        Dim lLeftArmorHull As Int32 = oHullTech.GetSideHullAllocation(HullTech.SlotType.eLeft, HullTech.SlotConfig.eArmorConfig)
        Dim lRightArmorHull As Int32 = oHullTech.GetSideHullAllocation(HullTech.SlotType.eRight, HullTech.SlotConfig.eArmorConfig)
        Dim lRearArmorHull As Int32 = oHullTech.GetSideHullAllocation(HullTech.SlotType.eRear, HullTech.SlotConfig.eArmorConfig)

        If oArmorTech Is Nothing = True AndAlso (ForeArmorUnits > 0 OrElse LeftArmorUnits > 0 OrElse RightArmorUnits > 0 OrElse AftArmorUnits > 0) = True Then
            LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If

        Dim fTotalArmorHull As Single = 0.0F
        If oArmorTech Is Nothing = False Then
            If ForeArmorUnits * oArmorTech.lHullUsagePerPlate > lFrontArmorHull Then
                LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
            If LeftArmorUnits * oArmorTech.lHullUsagePerPlate > lLeftArmorHull Then
                LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
            If RightArmorUnits * oArmorTech.lHullUsagePerPlate > lRightArmorHull Then
                LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
            If AftArmorUnits * oArmorTech.lHullUsagePerPlate > lRearArmorHull Then
                LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
            fTotalArmorHull = (ForeArmorUnits + LeftArmorUnits + RightArmorUnits + AftArmorUnits) * oArmorTech.lHullUsagePerPlate

            If oArmorTech.GetProductionCost Is Nothing = False AndAlso fTotalArmorHull <> 0 Then
                With oArmorTech.GetProductionCost
                    lEnlisted += .EnlistedCost
                    lOfficers += .OfficerCost
                    lColonists += .ColonistCost
                End With
            End If
        End If

        If oHullTech Is Nothing = False Then lPowerUsed += oHullTech.PowerRequired

        'Check Radar
        If oRadarTech Is Nothing = False Then
            If oRadarTech.HullRequired > oHullTech.GetHullAllocation(HullTech.SlotConfig.eRadar) Then
                LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
            lPowerUsed += oRadarTech.PowerRequired

            If oRadarTech.RadarType <> yHullTypeID AndAlso oRadarTech.RadarType <> 255 Then
                LogEvent(LogEventType.PossibleCheat, "Prototype designed with invalid radar for hull. OwnerID: " & Me.Owner.ObjectID)
                Return False
            End If

            If oRadarTech.GetProductionCost Is Nothing Then
                LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
            With oRadarTech.GetProductionCost
                lEnlisted += .EnlistedCost
                lOfficers += .OfficerCost
                lColonists += .ColonistCost
            End With
        End If

        'Check Shield
        If oShieldTech Is Nothing = False Then
            If oShieldTech.HullRequired > oHullTech.GetHullAllocation(HullTech.SlotConfig.eShields) Then
                LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
            If oShieldTech.lProjectionHullSize < oHullTech.HullSize Then
                LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
            lPowerUsed += oShieldTech.PowerRequired

            If oShieldTech.HullTypeID <> yHullTypeID AndAlso oShieldTech.HullTypeID <> 255 Then
                LogEvent(LogEventType.PossibleCheat, "Prototype designed with invalid shield for hull. OwnerID: " & Me.Owner.ObjectID)
                Return False
            End If

            If oShieldTech.GetProductionCost Is Nothing Then
                LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
            With oShieldTech.GetProductionCost
                lEnlisted += .EnlistedCost
                lOfficers += .OfficerCost
                lColonists += .ColonistCost
            End With
        End If

        'Check Weapons
        Dim lWpnCnt As Int32 = 0
        For X As Int32 = 0 To WeaponUB
            Dim oWpn As BaseWeaponTech = CType(Owner.GetTech(moWeapons(X).WeaponTechID, ObjectType.eWeaponTech), BaseWeaponTech)
            If oWpn Is Nothing Then oWpn = CType(Owner.GetPlayerTechKnowledgeTech(moWeapons(X).WeaponTechID, ObjectType.eWeaponTech, PlayerTechKnowledge.KnowledgeType.eSettingsLevel2), BaseWeaponTech)

            If oWpn Is Nothing = False Then
                lWpnCnt += 1

                If oWpn.WeaponClassTypeID = WeaponClassType.eBomb Then
                    If moWeapons(X).WeaponGroupTypeID <> WeaponGroupType.BombWeaponGroup Then
                        LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign Bomb weapon used in non bomb weapon group. Player: " & Me.Owner.ObjectID)
                        Return False
                    End If
                End If

                If oWpn.ComponentDevelopmentPhase <> eComponentDevelopmentPhase.eResearched Then
                    LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign Weapon is not researched. Player: " & Me.Owner.ObjectID)
                    Return False
                End If

                If moWeapons(X).WeaponGroupTypeID = WeaponGroupType.BombWeaponGroup Then
                    If oWpn.WeaponClassTypeID <> WeaponClassType.eBomb AndAlso oWpn.WeaponClassTypeID <> WeaponClassType.eEnergyBeam Then
                        LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign BombWeaponGroup assigned but non bomb or energy beam used. Player: " & Me.Owner.ObjectID)
                        Return False
                    End If
                    If oHullTech.yTypeID <> 1 OrElse oHullTech.ySubTypeID <> 3 Then
                        LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign Bomb weapon on non frigate! Player: " & Me.Owner.ObjectID)
                        Return False
                    End If
                End If

                Dim lHullReq As Int32 = oWpn.HullRequired
                If moWeapons(X).WeaponGroupTypeID = WeaponGroupType.PointDefenseWeaponGroup Then lHullReq = CInt(lHullReq * 0.5F)

                If lHullReq > oHullTech.GetWeaponHullAllotment(moWeapons(X).SlotX, moWeapons(X).SlotY) Then
                    LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                    Return False
                End If

                If oWpn.yHullTypeID <> yHullTypeID AndAlso oWpn.yHullTypeID <> 255 Then
                    LogEvent(LogEventType.PossibleCheat, "Prototype designed with invalid weapon for hull. OwnerID: " & Me.Owner.ObjectID)
                    Return False
                End If

                lPowerUsed += oWpn.PowerRequired
                If oWpn.GetProductionCost Is Nothing Then
                    LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                    Return False
                End If
                With oWpn.GetProductionCost
                    lEnlisted += .EnlistedCost
                    lOfficers += .OfficerCost
                    lColonists += .ColonistCost
                End With
            End If
        Next X

        If lWpnCnt > HullTech.MaxWpnSlots(oHullTech.yTypeID, oHullTech.ySubTypeID) Then
            LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If

        Dim lPowerMult As Int32 = 1
        If oHullTech Is Nothing = False Then
            Dim oModelDef As ModelDef = goModelDefs.GetModelDef(oHullTech.ModelID)
            If oModelDef Is Nothing = False Then
                If oModelDef.lSpecialTraitID = elModelSpecialTrait.PowerGen2 Then
                    lPowerMult = 2
                ElseIf oModelDef.lSpecialTraitID = elModelSpecialTrait.PowerGen3 Then
                    lPowerMult = 3
                End If
            End If
        End If
        If oEngineTech Is Nothing = True AndAlso lPowerUsed > 0 AndAlso oHullTech.yTypeID <> 2 Then
            LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        ElseIf oEngineTech Is Nothing = False AndAlso lPowerUsed > oEngineTech.PowerProd * lPowerMult AndAlso oHullTech.yTypeID <> 2 Then
            LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
            Return False
        End If

        Dim lBaseHullPerResidence As Int32 = Owner.oSpecials.HullPerResidence
        If oHullTech.HullSize <= 750 Then lBaseHullPerResidence = Math.Min(2, lBaseHullPerResidence)
        Dim lRequiredCrewQrtrs As Int32 = (lOfficers + lEnlisted + lColonists + MaxCrew) * lBaseHullPerResidence

        If oHullTech.HullSize <= 750 Then
            lRequiredCrewQrtrs = Math.Min(2, lRequiredCrewQrtrs)
        End If

        If oHullTech Is Nothing = False Then
            If ProductionHull + fTotalArmorHull + lRequiredCrewQrtrs > oHullTech.GetHullAllocation(HullTech.SlotConfig.eArmorConfig) Then
                LogEvent(LogEventType.PossibleCheat, "Prototype.ValidateDesign invalid values: " & Me.Owner.ObjectID)
                Return False
            End If
        End If

        Return True
    End Function

    Public Overrides Function SetFromDesignMsg(ByRef yData() As Byte) As Boolean
        Dim lPos As Int32 = 14
        Dim bRes As Boolean = False
        Dim lTemp As Int32
        Dim lCnt As Int32 = 0
        Dim ySlotX As Byte
        Dim ySlotY As Byte
        Dim yGroupType As Byte

        '2 bytes for hull tech type id
        '4 for id
        '4 bytes for researcher id, 2 bytes for type id

        If Me.Owner Is Nothing Then Return False

        Try
            ObjTypeID = ObjectType.ePrototype

            '20 bytes for name
            ReDim Me.PrototypeName(19)
            Array.Copy(yData, lPos, Me.PrototypeName, 0, 20)
            lPos += 20

            'EngineID (4)
            lEngineTech = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'ArmorID (4)
            lArmorTech = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'HangarID (4)
            'lHangarTech = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'HullID (4)
            lHullTech = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'RadarID (4)
            lRadarTech = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            'ShieldID (4)
            lShieldTech = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            'Armor Sides (4 each)
            ForeArmorUnits = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            AftArmorUnits = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            LeftArmorUnits = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            RightArmorUnits = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            'Prod slots
            ProductionHull = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            MaxCrew = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            'Cnt (4)
            lCnt = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            WeaponUB = -1
            For X As Int32 = 0 To lCnt - 1
                'WeaponTechID (4)
                lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                'SlotX (1)
                ySlotX = yData(lPos) : lPos += 1
                'SlotY (1)
                ySlotY = yData(lPos) : lPos += 1
                'WeaponGroupType (1)
                yGroupType = yData(lPos) : lPos += 1
                Me.AddWeaponPlacement(lTemp, ySlotX, ySlotY, CType(yGroupType, WeaponGroupType))
            Next X

            bRes = True

        Catch ex As Exception
            bRes = False
            GlobalVars.LogEvent(LogEventType.CriticalError, "PrototypeTech.SetFromDesignMsg Error: " & ex.Message)
        End Try

        Return bRes
    End Function

#Region "  Helper Functions  "
    Private moEngineTech As EngineTech = Nothing
    Public ReadOnly Property oEngineTech() As EngineTech
        Get
            If lEngineTech < 1 AndAlso moEngineTech Is Nothing = False Then moEngineTech = Nothing
            If (moEngineTech Is Nothing OrElse moEngineTech.ObjectID <> lEngineTech) AndAlso lEngineTech > 0 Then
                If Me.Owner Is Nothing = False Then
                    moEngineTech = CType(Me.Owner.GetTech(lEngineTech, ObjectType.eEngineTech), EngineTech)
                    If moEngineTech Is Nothing Then moEngineTech = CType(Me.Owner.GetPlayerTechKnowledgeTech(lEngineTech, ObjectType.eEngineTech, PlayerTechKnowledge.KnowledgeType.eSettingsLevel2), EngineTech)
                End If
            End If
            Return moEngineTech
        End Get
    End Property

    Private moArmorTech As ArmorTech = Nothing
    Public ReadOnly Property oArmorTech() As ArmorTech
        Get
            If lArmorTech < 1 AndAlso moArmorTech Is Nothing = False Then moArmorTech = Nothing
            If (moArmorTech Is Nothing OrElse moArmorTech.ObjectID <> lArmorTech) AndAlso lArmorTech > 0 Then
                If Me.Owner Is Nothing = False Then
                    moArmorTech = CType(Me.Owner.GetTech(lArmorTech, ObjectType.eArmorTech), ArmorTech)
                    If moArmorTech Is Nothing Then moArmorTech = CType(Me.Owner.GetPlayerTechKnowledgeTech(lArmorTech, ObjectType.eArmorTech, PlayerTechKnowledge.KnowledgeType.eSettingsLevel2), ArmorTech)
                End If
            End If
            Return moArmorTech
        End Get
    End Property

    Private moHullTech As HullTech = Nothing
    Public ReadOnly Property oHullTech() As HullTech
        Get
            If lHullTech < 1 AndAlso moHullTech Is Nothing = False Then moHullTech = Nothing
            If (moHullTech Is Nothing OrElse moHullTech.ObjectID <> lHullTech) AndAlso lHullTech > 0 Then
                If Me.Owner Is Nothing = False Then
                    moHullTech = CType(Me.Owner.GetTech(lHullTech, ObjectType.eHullTech), HullTech)
                End If
            End If
            Return moHullTech
        End Get
    End Property

    Private moRadarTech As RadarTech = Nothing
    Public ReadOnly Property oRadarTech() As RadarTech
        Get
            If lRadarTech < 1 AndAlso moRadarTech Is Nothing = False Then moRadarTech = Nothing
            If (moRadarTech Is Nothing OrElse moRadarTech.ObjectID <> lRadarTech) AndAlso lRadarTech > 0 Then
                If Me.Owner Is Nothing = False Then
                    moRadarTech = CType(Me.Owner.GetTech(lRadarTech, ObjectType.eRadarTech), RadarTech)
                    If moRadarTech Is Nothing Then moRadarTech = CType(Me.Owner.GetPlayerTechKnowledgeTech(lRadarTech, ObjectType.eRadarTech, PlayerTechKnowledge.KnowledgeType.eSettingsLevel2), RadarTech)
                End If
            End If
            Return moRadarTech
        End Get
    End Property

    Private moShieldTech As ShieldTech = Nothing
    Public ReadOnly Property oShieldTech() As ShieldTech
        Get
            If lShieldTech < 1 AndAlso moShieldTech Is Nothing = False Then moShieldTech = Nothing
            If (moShieldTech Is Nothing OrElse moShieldTech.ObjectID <> lShieldTech) AndAlso lShieldTech > 0 Then
                If Me.Owner Is Nothing = False Then
                    moShieldTech = CType(Me.Owner.GetTech(lShieldTech, ObjectType.eShieldTech), ShieldTech)
                    If moShieldTech Is Nothing Then moShieldTech = CType(Me.Owner.GetPlayerTechKnowledgeTech(lShieldTech, ObjectType.eShieldTech, PlayerTechKnowledge.KnowledgeType.eSettingsLevel2), ShieldTech)
                End If
            End If
            Return moShieldTech
        End Get
    End Property

#End Region

    Public Function GetTotalCrew() As Int32
        If moProductionCost Is Nothing = False Then
            Return moProductionCost.ColonistCost + moProductionCost.EnlistedCost + moProductionCost.OfficerCost '+ MaxCrew
        Else
            Dim oTmpCost As ProductionCost = GetCurrentProductionCost()
            If oTmpCost Is Nothing = False Then
                Return oTmpCost.ColonistCost + oTmpCost.EnlistedCost + oTmpCost.OfficerCost '+ MaxCrew
            End If
            Return MaxCrew
        End If
    End Function

    Protected Overrides Sub CalculateBothProdCosts()

        If oHullTech Is Nothing Then
            Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
            Return
        End If

        'If mlNoiseCnt = -1 Then ComponentDesigned()


        'Ok, let's get our values...
        Dim lTotalCrew As Int32 = 0
        If oHullTech Is Nothing = False Then
            With oHullTech.GetProductionCost
                lTotalCrew += .EnlistedCost + .OfficerCost + .ColonistCost
            End With
        End If
        If oEngineTech Is Nothing = False Then
            With oEngineTech.GetProductionCost
                lTotalCrew += .EnlistedCost + .OfficerCost + .ColonistCost
            End With
        End If
        If oRadarTech Is Nothing = False Then
            With oRadarTech.GetProductionCost
                lTotalCrew += .EnlistedCost + .OfficerCost + .ColonistCost
            End With
        End If
        If oShieldTech Is Nothing = False Then
            With oShieldTech.GetProductionCost
                lTotalCrew += .EnlistedCost + .OfficerCost + .ColonistCost
            End With
        End If

        For lWpnIdx As Int32 = 0 To WeaponUB
            Dim oWpn As BaseWeaponTech = CType(Owner.GetTech(moWeapons(lWpnIdx).WeaponTechID, ObjectType.eWeaponTech), BaseWeaponTech)
            If oWpn Is Nothing Then oWpn = CType(Owner.GetPlayerTechKnowledgeTech(moWeapons(lWpnIdx).WeaponTechID, ObjectType.eWeaponTech, PlayerTechKnowledge.KnowledgeType.eSettingsLevel2), BaseWeaponTech)
            If oWpn Is Nothing = False Then
                Dim oTempProdCost As ProductionCost = oWpn.GetProductionCost()
                lTotalCrew += oTempProdCost.EnlistedCost
                lTotalCrew += oTempProdCost.OfficerCost
                lTotalCrew += oTempProdCost.ColonistCost
            End If
        Next lWpnIdx

        lTotalCrew += MaxCrew
        Dim lAddtEnl As Int32 = 0
        Dim lAddtOff As Int32 = 0
        If MaxCrew <> 0 Then
            lAddtOff = MaxCrew \ 6
            lAddtEnl = (lAddtOff * 5) + (MaxCrew - (lAddtOff * 5) - lAddtOff)
        End If

        Dim fModifiedSeed As Single = 0.8F + (0.7F * RandomSeed)

        If moResearchCost Is Nothing Then moResearchCost = New ProductionCost
        With moResearchCost
            .ColonistCost = 0
            .EnlistedCost = 0
            .OfficerCost = 0
            .ObjectID = Me.ObjectID
            .ObjTypeID = Me.ObjTypeID

            'research time = crewcount * 60 seconds 
            'research cost = max(1000000,hullsize * crewcount) * randomseed
            .PointsRequired = CLng(Math.Max(1, lTotalCrew)) * 60000L
            .CreditCost = CLng(Math.Max(1000000L, CLng(oHullTech.HullSize) * CLng(Math.Max(1, lTotalCrew))) * fModifiedSeed)

            If Me.Owner.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                .PointsRequired \= 100
                .CreditCost \= 100
            End If

            '.CreditCost = CLng(ml_NOISE_COST_MULT * (mlNoiseCnt * (RandomSeed + 0.5F))) + (oHullTech.HullSize \ 2000)
            '.PointsRequired = .CreditCost
            '.CreditCost += oHullTech.GetResearchCost.CreditCost
            '.PointsRequired += oHullTech.GetResearchCost.PointsRequired

            .ItemCostUB = -1
            ReDim .ItemCosts(-1)
        End With

        If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
        moProductionCost.ObjectID = Me.ObjectID
        moProductionCost.ObjTypeID = Me.ObjTypeID

        moProductionCost.ColonistCost = 0
        moProductionCost.EnlistedCost = 0
        moProductionCost.OfficerCost = 0

        With moProductionCost
            .CreditCost = 0 ' CLng(ml_NOISE_COST_MULT * (mlNoiseCnt * (RandomSeed + 0.5F))) + (oHullTech.HullSize \ 2000)
            .PointsRequired = 0 '.CreditCost

            'production cost = crewcount*(1+AverageArmorAllocation)*SlotCountFromHull* randomseed * size 
            'Dim fAverageArmorAlloc As Single = Me.ForeArmorUnits + Me.LeftArmorUnits + Me.AftArmorUnits + Me.RightArmorUnits
            'fAverageArmorAlloc /= 4.0F
            '.CreditCost = CLng(Math.Max(lTotalCrew, 1) * (1 + fAverageArmorAlloc) * Math.Max(1, oHullTech.GetSlotTypeGroupCount()) * fModifiedSeed * Math.Max(oHullTech.HullSize, 1))
            '.CreditCost = Math.Max(oHullTech.GetProductionCost.CreditCost, .CreditCost \ 30)

            'BaseCrewCost = Math.Max(Hull.Type.MaxWeaponCount+1,Prototype.CrewCount)
            Dim fBaseCrewCost As Single = Math.Max(HullTech.MaxWpnSlots(oHullTech.yTypeID, oHullTech.ySubTypeID) + 1, lTotalCrew) '  oHullTech.HullSize) + 1, lTotalCrew)
            'BaseHullCost = BaseCrewCost * Hull.StructHitPoints * randomseed
            Dim fBaseHullCost As Single = Math.Max(1, fBaseCrewCost * oHullTech.StructuralHitPoints * fModifiedSeed)
            'ProtoCost = BaseHullCost * SlotTypeGroupCount
            .CreditCost = CLng(Math.Ceiling(Math.Max(1, (fModifiedSeed * oHullTech.GetProductionCost.CreditCost)) + (fBaseHullCost * Math.Max(1, oHullTech.GetSlotTypeGroupCount()))))


            'Production Time = max(20 seconds, size/3000 seconds) * slottypegroupcount
            Dim lPointsPerSec As Int64
            If oHullTech.yTypeID = 2 Then
                lPointsPerSec = 3000
            Else : lPointsPerSec = 15000
            End If
            .PointsRequired = Math.Max(20L * lPointsPerSec, CLng(CLng(oHullTech.HullSize) / (3000L * lPointsPerSec))) * CLng(oHullTech.GetSlotTypeGroupCount)

            .ItemCostUB = -1
            If oHullTech Is Nothing = False Then
                '.CreditCost += oHullTech.GetProductionCost.CreditCost
                '.EnlistedCost += oHullTech.GetProductionCost.EnlistedCost
                '.OfficerCost += oHullTech.GetProductionCost.OfficerCost
                '.ColonistCost += oHullTech.GetProductionCost.ColonistCost
                '.PointsRequired += oHullTech.GetProductionCost.PointsRequired

                For X As Int32 = 0 To oHullTech.GetProductionCost.ItemCostUB
                    .AddProductionCostItem(oHullTech.GetProductionCost.ItemCosts(X).ItemID, oHullTech.GetProductionCost.ItemCosts(X).ItemTypeID, oHullTech.GetProductionCost.ItemCosts(X).QuantityNeeded)
                Next X

                'Dim lSeconds As Int32 = (oHullTech.HullSize \ 10000) + 10
                'If oHullTech.yTypeID = 2 Then
                '	'facility... requires a unit to build therefore, 1 second is 300 points
                '	.PointsRequired += (lSeconds * 3000)
                '	If oHullTech.ModelID = 16 Then
                '		'MSC - 11/21/07 - by request of JHC, turrets to have additional build time
                '		.PointsRequired += 18000
                '	End If
                'Else
                '	'other... requires anything... therefore, 1 second is 15k points
                '	.PointsRequired += (lSeconds * 15000)
                'End If
            End If

            If Me.Owner.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
                .PointsRequired \= 100
                .CreditCost \= 100
            End If

            If oEngineTech Is Nothing = False Then
                '.PointsRequired += CLng(oEngineTech.HullRequired) * 1000L
                .AddProductionCostItem(oEngineTech.ObjectID, oEngineTech.ObjTypeID, 1)
                Dim oTempProdCost As ProductionCost = oEngineTech.GetProductionCost()
                .EnlistedCost += oTempProdCost.EnlistedCost
                .OfficerCost += oTempProdCost.OfficerCost
                .ColonistCost += oTempProdCost.ColonistCost
            End If
            If oRadarTech Is Nothing = False Then
                '.PointsRequired += CLng(oRadarTech.PowerRequired) * 5L
                .AddProductionCostItem(oRadarTech.ObjectID, oRadarTech.ObjTypeID, 1)
                Dim oTempProdCost As ProductionCost = oRadarTech.GetProductionCost()
                .EnlistedCost += oTempProdCost.EnlistedCost
                .OfficerCost += oTempProdCost.OfficerCost
                .ColonistCost += oTempProdCost.ColonistCost
            End If
            If oShieldTech Is Nothing = False Then
                '.PointsRequired += (CLng(oShieldTech.PowerRequired) * CLng(oShieldTech.HullRequired)) \ 2L
                .AddProductionCostItem(oShieldTech.ObjectID, oShieldTech.ObjTypeID, 1)
                Dim oTempProdCost As ProductionCost = oShieldTech.GetProductionCost()
                .EnlistedCost += oTempProdCost.EnlistedCost
                .OfficerCost += oTempProdCost.OfficerCost
                .ColonistCost += oTempProdCost.ColonistCost
            End If

            If oArmorTech Is Nothing = False Then
                Dim lTemp As Int32 = ForeArmorUnits + LeftArmorUnits + AftArmorUnits + RightArmorUnits
                '.PointsRequired += CLng(lTemp) * 10L
                .AddProductionCostItem(oArmorTech.ObjectID, oArmorTech.ObjTypeID, lTemp)
            End If

            .EnlistedCost += lAddtEnl
            .OfficerCost += lAddtOff

            For lWpnIdx As Int32 = 0 To WeaponUB
                .AddProductionCostItem(moWeapons(lWpnIdx).WeaponTechID, ObjectType.eWeaponTech, 1)

                Dim oWpn As BaseWeaponTech = CType(Owner.GetTech(moWeapons(lWpnIdx).WeaponTechID, ObjectType.eWeaponTech), BaseWeaponTech)
                If oWpn Is Nothing Then oWpn = CType(Owner.GetPlayerTechKnowledgeTech(moWeapons(lWpnIdx).WeaponTechID, ObjectType.eWeaponTech, PlayerTechKnowledge.KnowledgeType.eSettingsLevel2), BaseWeaponTech)
                If oWpn Is Nothing = False Then
                    Dim oTempProdCost As ProductionCost = oWpn.GetProductionCost()
                    .EnlistedCost += oTempProdCost.EnlistedCost
                    .OfficerCost += oTempProdCost.OfficerCost
                    .ColonistCost += oTempProdCost.ColonistCost
                End If
            Next lWpnIdx
        End With
    End Sub

    Private Function GetVersionDPSPenalty() As Int32
        Dim lVersions(-1) As Int32
        Dim lVersionUB As Int32 = -1
        If oHullTech Is Nothing = False Then
            Dim lNewVer As Int32 = oHullTech.lVersionNum
            Dim bFound As Boolean = False
            For X As Int32 = 0 To lVersionUB
                If lVersions(X) = lNewVer Then
                    bFound = True
                    Exit For
                End If
            Next X
            If bFound = False Then
                lVersionUB += 1
                ReDim Preserve lVersions(lVersionUB)
                lVersions(lVersionUB) = lNewVer
            End If
        End If

        If oEngineTech Is Nothing = False Then
            Dim lNewVer As Int32 = oEngineTech.lVersionNum
            Dim bFound As Boolean = False
            For X As Int32 = 0 To lVersionUB
                If lVersions(X) = lNewVer Then
                    bFound = True
                    Exit For
                End If
            Next X
            If bFound = False Then
                lVersionUB += 1
                ReDim Preserve lVersions(lVersionUB)
                lVersions(lVersionUB) = lNewVer
            End If
        End If

        If oArmorTech Is Nothing = False Then
            Dim lNewVer As Int32 = oArmorTech.lVersionNum
            Dim bFound As Boolean = False
            For X As Int32 = 0 To lVersionUB
                If lVersions(X) = lNewVer Then
                    bFound = True
                    Exit For
                End If
            Next X
            If bFound = False Then
                lVersionUB += 1
                ReDim Preserve lVersions(lVersionUB)
                lVersions(lVersionUB) = lNewVer
            End If
        End If

        If oRadarTech Is Nothing = False Then
            Dim lNewVer As Int32 = oRadarTech.lVersionNum
            Dim bFound As Boolean = False
            For X As Int32 = 0 To lVersionUB
                If lVersions(X) = lNewVer Then
                    bFound = True
                    Exit For
                End If
            Next X
            If bFound = False Then
                lVersionUB += 1
                ReDim Preserve lVersions(lVersionUB)
                lVersions(lVersionUB) = lNewVer
            End If
        End If

        If oShieldTech Is Nothing = False Then
            Dim lNewVer As Int32 = oShieldTech.lVersionNum
            Dim bFound As Boolean = False
            For X As Int32 = 0 To lVersionUB
                If lVersions(X) = lNewVer Then
                    bFound = True
                    Exit For
                End If
            Next X
            If bFound = False Then
                lVersionUB += 1
                ReDim Preserve lVersions(lVersionUB)
                lVersions(lVersionUB) = lNewVer
            End If
        End If

        'And Weapons...
        For lWpn As Int32 = 0 To WeaponUB
            Dim oWpn As BaseWeaponTech = CType(Me.Owner.GetTech(moWeapons(lWpn).WeaponTechID, ObjectType.eWeaponTech), BaseWeaponTech)
            If oWpn Is Nothing = False Then
                Dim lNewVer As Int32 = oWpn.lVersionNum
                Dim bFound As Boolean = False
                For X As Int32 = 0 To lVersionUB
                    If lVersions(X) = lNewVer Then
                        bFound = True
                        Exit For
                    End If
                Next X
                If bFound = False Then
                    lVersionUB += 1
                    ReDim Preserve lVersions(lVersionUB)
                    lVersions(lVersionUB) = lNewVer
                End If
            End If
        Next lWpn

        'Then go through and look at the VersionRels... determine the Maximum percentage drop
        Dim lMult As Int32 = 100
        For X As Int32 = 0 To lVersionUB
            For Y As Int32 = 0 To lVersionUB
                If X = Y Then Continue For
                For lIdx As Int32 = 0 To VersionListUB
                    If VersionList(lIdx) Is Nothing = False AndAlso VersionList(lIdx).lVersionNumber = lVersions(X) AndAlso VersionList(lIdx).lOtherVersion = lVersions(Y) Then
                        lMult = Math.Min(lMult, VersionList(lIdx).lNoisePerc)
                        Exit For
                    End If
                Next lIdx
            Next Y
        Next X
        'return that value
        Return lMult
    End Function

    Public Overrides Sub ComponentDesigned()
        If oHullTech Is Nothing Then
            Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
            Return
        End If

        mlNoiseCnt = 0

        Dim bHasMicroTech As Boolean = False

        'Now, we need to calculate everything based on noises...
        'We always have a hull...
        Dim lHullStructDensity As Int32 = oHullTech.StructureMineral.GetPropertyValue(eMinPropID.Density)
        Dim lHullStructCompress As Int32 = oHullTech.StructureMineral.GetPropertyValue(eMinPropID.Compressibility)
        Dim lHullStructHardness As Int32 = oHullTech.StructureMineral.GetPropertyValue(eMinPropID.Hardness)

        'Might not always have an armor...
        Dim lArmorMaxDensity As Int32 = 0
        Dim lArmorMaxMagProd As Int32 = 0
        Dim lArmorMaxMagReact As Int32 = 0
        Dim lArmorMaxReflection As Int32 = 0
        Dim lArmorMaxThermCond As Int32 = 0
        If oArmorTech Is Nothing = False Then
            lArmorMaxDensity = oArmorTech.GetMaxProperty(eMinPropID.Density)
            lArmorMaxMagProd = oArmorTech.GetMaxProperty(eMinPropID.MagneticProduction)
            lArmorMaxMagReact = oArmorTech.GetMaxProperty(eMinPropID.MagneticReaction)
            lArmorMaxReflection = oArmorTech.GetMaxProperty(eMinPropID.Reflection)
            lArmorMaxThermCond = oArmorTech.GetMaxProperty(eMinPropID.ThermalConductance)

            If (oArmorTech.MajorDesignFlaw And Epica_Tech.eComponentDesignFlaw.eMicroTech) = Epica_Tech.eComponentDesignFlaw.eMicroTech Then
                bHasMicroTech = True
            End If
        End If

        'Might not always have a shield...
        Dim lShieldMaxMagProd As Int32 = 0
        Dim lShieldCoilTempSens As Int32 = 0
        Dim lShieldAccTempSens As Int32 = 0
        Dim lShieldAccReflection As Int32 = 0
        If oShieldTech Is Nothing = False Then
            lShieldMaxMagProd = oShieldTech.GetMaxProperty(eMinPropID.MagneticProduction)
            lShieldCoilTempSens = oShieldTech.CoilMineral.GetPropertyValue(eMinPropID.TemperatureSensitivity)
            lShieldAccTempSens = oShieldTech.AcceleratorMineral.GetPropertyValue(eMinPropID.TemperatureSensitivity)
            lShieldAccReflection = oShieldTech.AcceleratorMineral.GetPropertyValue(eMinPropID.Reflection)

            If (oShieldTech.MajorDesignFlaw And Epica_Tech.eComponentDesignFlaw.eMicroTech) = Epica_Tech.eComponentDesignFlaw.eMicroTech Then
                bHasMicroTech = True
            End If
        End If

        'Might not always have an engine...
        Dim lEngineMaxMagProd As Int32 = 0
        If oEngineTech Is Nothing = False Then
            lEngineMaxMagProd = oEngineTech.GetMaxProperty(eMinPropID.MagneticProduction)

            If (oEngineTech.MajorDesignFlaw And Epica_Tech.eComponentDesignFlaw.eMicroTech) = Epica_Tech.eComponentDesignFlaw.eMicroTech Then
                bHasMicroTech = True
            End If
        End If

        'Might not always have a radar...
        Dim lRadarEmitterTempSens As Int32 = 0
        If oRadarTech Is Nothing = False Then
            lRadarEmitterTempSens = oRadarTech.EmitterMaterial.GetPropertyValue(eMinPropID.TemperatureSensitivity)

            'If (oRadarTech.MajorDesignFlaw And Epica_Tech.eComponentDesignFlaw.eMicroTech) = Epica_Tech.eComponentDesignFlaw.eMicroTech Then
            '    bHasMicroTech = True
            'End If
        End If

        Dim lPrototypeHullSizeResult As Int32 = CalculateActualHullUsed()

        'We have everything now...
        Dim fTemp As Single = 0.0F

        'Engine specifics...
        Dim lThrust As Int32 = 0
        Dim lSpeed As Int32 = 0
        Dim lManeuver As Int32 = 0
        Dim lPower As Int32 = 0
        If oEngineTech Is Nothing = False Then
            'THRUST
            fTemp = (lHullStructDensity + lArmorMaxDensity) / 2.0F
            If fTemp > 20 Then
                mlNoiseCnt += 1
                lThrust = CInt(Math.Floor(oEngineTech.Thrust - (oEngineTech.Thrust * (fTemp / 2000.0F))))
            Else : lThrust = oEngineTech.Thrust
            End If

            'SPEED
            fTemp = CSng(lThrust / lPrototypeHullSizeResult)
            lSpeed = CInt(fTemp * CInt(oEngineTech.Speed))
            If lSpeed <> CInt(oEngineTech.Speed) Then mlNoiseCnt += 1

            'MANEUVER
            If lHullStructCompress > 40 Then
                fTemp -= (lHullStructCompress / 400.0F)
                mlNoiseCnt += 1
            End If
            lManeuver = CInt(fTemp * CInt(oEngineTech.Maneuver))

            'POWER
            fTemp = lArmorMaxMagProd
            If fTemp > 20 Then
                lPower = CInt(Math.Floor(oEngineTech.PowerProd - (oEngineTech.PowerProd * (fTemp / 2000.0F))))
                mlNoiseCnt += 1
            Else : lPower = oEngineTech.PowerProd
            End If
        End If

        'Shield specifics
        Dim lShieldHP As Int32 = 0
        Dim lRechargeRate As Int32 = 0
        Dim lRechargeFreq As Int32 = 0          'this value is given to us as a number of seconds... reduce it down
        If oShieldTech Is Nothing = False Then
            'MAX HP
            fTemp = Math.Max(lEngineMaxMagProd, lArmorMaxMagProd)
            lShieldHP = oShieldTech.MaxHitPoints
            'If fTemp > 20 Then
            '    lShieldHP = CInt(Math.Floor(lShieldHP - (lShieldHP * (fTemp / 2000.0F))))
            '    mlNoiseCnt += 1
            'End If

            'RECHARGE RATE - shares fTemp of MaxHP
            lRechargeRate = oShieldTech.RechargeRate
            'If fTemp > 20 Then
            '    lRechargeRate = CInt(Math.Floor(lRechargeRate - (lRechargeRate * (fTemp / 2000.0F))))
            '    mlNoiseCnt += 1
            'End If

            'RECHARGE FREQ
            lRechargeFreq = oShieldTech.RechargeFreq
            'fTemp = Math.Max(CInt(fTemp), lArmorMaxMagReact)
            'If fTemp > 20 Then
            '    lRechargeFreq = CInt(Math.Floor(lRechargeFreq - (lRechargeFreq * (fTemp / 2000.0F))))
            '    mlNoiseCnt += 1
            'End If
        End If

        'Armor specifics
        Dim lImpactResist As Int32 = 0
        If oArmorTech Is Nothing = False Then
            lImpactResist = CInt(oArmorTech.yImpactResist)
            fTemp = 100 - ((lHullStructDensity + lHullStructHardness) / 2.0F)
            If fTemp > 20 Then
                lImpactResist = CInt(Math.Floor(lImpactResist - (lImpactResist * (fTemp / 2000.0F))))
                mlNoiseCnt += 1
            End If
        End If

        'Radar Specifics
        Dim lWeaponAcc As Int32 = 0
        Dim lScanRes As Int32 = 0
        Dim lOptRange As Int32 = 0
        Dim lMaxRange As Int32 = 0
        Dim lDisRes As Int32 = 0
        If oRadarTech Is Nothing = False Then
            'WEAPON ACCURACY
            lWeaponAcc = CInt(oRadarTech.WeaponAcc)
            fTemp = (lArmorMaxReflection + lArmorMaxDensity + lHullStructDensity) / 3.0F
            If fTemp > 20 Then
                lWeaponAcc = CInt(Math.Floor(lWeaponAcc - (lWeaponAcc * (fTemp / 2000.0F))))
                mlNoiseCnt += 1
            End If

            'SCAN RESOLUTION
            lScanRes = CInt(oRadarTech.ScanResolution)
            fTemp = (lArmorMaxThermCond + lArmorMaxMagProd) / 2.0F
            If fTemp > 20 Then
                lScanRes = CInt(Math.Floor(lScanRes - (lScanRes * (fTemp / 2000.0F))))
                mlNoiseCnt += 1
            End If

            'OPTIMUM RANGE and MAXIMUM RANGE
            fTemp = Math.Max(lEngineMaxMagProd, lShieldMaxMagProd)
            lOptRange = CInt(oRadarTech.OptimumRange)
            lMaxRange = CInt(oRadarTech.MaximumRange)
            If fTemp > 20 Then
                lOptRange = CInt(Math.Floor(lOptRange - (lOptRange * (fTemp / 2000.0F))))
                lMaxRange = CInt(Math.Floor(lMaxRange - (lMaxRange * (fTemp / 2000.0F))))
                mlNoiseCnt += 2
            End If

            'DISRUPTION RESIST
            fTemp = Math.Max(lHullStructDensity, lArmorMaxDensity)
            lDisRes = CInt(oRadarTech.DisruptionResistance)
            If fTemp > 20 Then
                lDisRes = CInt(Math.Floor(lDisRes - (lDisRes * (fTemp / 2000.0F))))
                mlNoiseCnt += 1
            End If
        End If

        'WEAPONS...
        Dim lWpnROF(WeaponUB) As Int32
        Dim lWpnOptRange(WeaponUB) As Int32
        Dim lWpnMaxDmg(WeaponUB) As Int32

        Dim lForeWpnCnt As Int32 = 0
        Dim lLeftWpnCnt As Int32 = 0
        Dim lRightWpnCnt As Int32 = 0
        Dim lRearWpnCnt As Int32 = 0

        For X As Int32 = 0 To WeaponUB
            Dim oWpn As BaseWeaponTech = CType(Owner.GetTech(moWeapons(X).WeaponTechID, ObjectType.eWeaponTech), BaseWeaponTech)
            If oWpn Is Nothing Then
                'ok, check the ptk
                oWpn = CType(Owner.GetPlayerTechKnowledgeTech(moWeapons(X).WeaponTechID, ObjectType.eWeaponTech, PlayerTechKnowledge.KnowledgeType.eSettingsLevel1), BaseWeaponTech)
            End If

            If oWpn Is Nothing = False Then
                If (oWpn.MajorDesignFlaw And Epica_Tech.eComponentDesignFlaw.eMicroTech) = Epica_Tech.eComponentDesignFlaw.eMicroTech Then
                    bHasMicroTech = True
                End If

                'TODO: define this more when more weapons are ready
                Select Case oWpn.WeaponClassTypeID
                    Case WeaponClassType.eBomb
                        With CType(oWpn, BombWeaponTech)
                            lWpnROF(X) = CInt(.iROF)
                            lWpnOptRange(X) = CInt(.yRange)
                        End With
                    Case WeaponClassType.eEnergyBeam
                        With CType(oWpn, BeamWeaponTech)
                            lWpnROF(X) = CInt(.ROF)
                            lWpnOptRange(X) = CInt(.MaxRange)
                            lWpnMaxDmg(X) = CInt(.MaxDamage)
                        End With
                    Case WeaponClassType.eEnergyPulse
                        With CType(oWpn, PulseWeaponTech)
                            lWpnROF(X) = CInt(.ROF)
                            lWpnOptRange(X) = CInt(.MaxRange)
                        End With
                    Case WeaponClassType.eMine
                    Case WeaponClassType.eMissile
                        With CType(oWpn, MissileWeaponTech)
                            If .ROF > 0 Then
                                lWpnROF(X) = .ROF
                            End If
                            lWpnOptRange(X) = .FlightTime 'CInt((CInt(.FlightTime) * CInt(.MaxSpeed)))
                        End With
                    Case WeaponClassType.eProjectile
                        With CType(oWpn, ProjectileWeaponTech)
                            lWpnROF(X) = CInt(.ROF)
                            lWpnOptRange(X) = CInt(.MaxRange)
                        End With
                End Select



                Select Case moWeapons(X).ArcID
                    Case UnitArcs.eLeftArc
                        lLeftWpnCnt += 1
                    Case UnitArcs.eRightArc
                        lRightWpnCnt += 1
                    Case UnitArcs.eBackArc, UnitArcs.eAllArcs
                        lRearWpnCnt += 1
                    Case UnitArcs.eForwardArc
                        lForeWpnCnt += 1
                End Select

                'Now... do our calcs
                'ROF
                fTemp = Math.Max(lShieldCoilTempSens, lShieldAccTempSens)
                If oShieldTech Is Nothing = False Then
                    fTemp *= Math.Min(CSng(oShieldTech.PowerRequired / oWpn.PowerRequired), 255.0F)
                End If
                If fTemp > 20 Then
                    lWpnROF(X) = CInt(Math.Floor(lWpnROF(X) + (lWpnROF(X) * (fTemp / 2000.0F))))

                    If X = 0 Then mlNoiseCnt += 1
                End If

                'OPTIMUM RANGE
                If lShieldAccReflection > 20 Then
                    lWpnOptRange(X) = CInt(Math.Floor(lWpnOptRange(X) - (lWpnOptRange(X) * (fTemp / 2000.0F))))

                    If X = 0 Then mlNoiseCnt += 1
                End If

                'TODO: Redetermine this noise system
                'MAXIMUM DAMAGE
                'If lRadarEmitterTempSens > 20 Then
                '    lWpnMaxDmg(X) = CInt(Math.Floor(lWpnMaxDmg(X) - (lWpnMaxDmg(X) * (fTemp / 2000.0F))))

                '    If X = 0 Then mlNoiseCnt += 1
                'End If
            End If
        Next X

        Dim lSpecTraitID As elModelSpecialTrait = elModelSpecialTrait.NoSpecialTrait
        Dim oModelDef As ModelDef = goModelDefs.GetModelDef(oHullTech.ModelID)
        If oModelDef Is Nothing = False Then
            lSpecTraitID = CType(oModelDef.lSpecialTraitID, elModelSpecialTrait)
        End If

        'Now, fill in our entity def...
        If oHullTech.yTypeID = 2 Then
            moExpectedDef = New FacilityDef
        Else : moExpectedDef = New Epica_Entity_Def
        End If
        With moExpectedDef
            .DefName = Me.PrototypeName
            .oPrototype = Me
            .OwnerID = Owner.ObjectID

            .ObjectID = -1

            .lExtendedFlags = 0
            If bHasMicroTech = True Then
                .lExtendedFlags = .lExtendedFlags Or Epica_Entity_Def.eExtendedFlagValues.ContainsMicroTech
            End If

            If oHullTech.yTypeID = 2 Then
                .ObjTypeID = ObjectType.eFacilityDef
            Else : .ObjTypeID = ObjectType.eUnitDef
            End If

            'ArmorHPs
            If oArmorTech Is Nothing = False Then
                .Q1_MaxHP = oArmorTech.lHPPerPlate * Me.ForeArmorUnits
                .Q2_MaxHP = oArmorTech.lHPPerPlate * Me.LeftArmorUnits
                .Q3_MaxHP = oArmorTech.lHPPerPlate * Me.AftArmorUnits
                .Q4_MaxHP = oArmorTech.lHPPerPlate * Me.RightArmorUnits

                If lSpecTraitID = elModelSpecialTrait.Armor1000 Then
                    .Q1_MaxHP *= 10 '00
                    .Q2_MaxHP *= 10 '00
                    .Q3_MaxHP *= 10 '00
                    .Q4_MaxHP *= 10 '00
                End If
                If oHullTech Is Nothing = False AndAlso oHullTech.yTypeID = 2 AndAlso oHullTech.ySubTypeID <> 9 Then
                    .Q1_MaxHP *= 10
                    .Q2_MaxHP *= 10
                    .Q3_MaxHP *= 10
                    .Q4_MaxHP *= 10
                End If
            End If

            If oEngineTech Is Nothing = False Then
                lManeuver = Math.Min(oEngineTech.Maneuver, lManeuver)
                lSpeed = Math.Min(oEngineTech.Speed, lSpeed)
                Select Case lSpecTraitID
                    Case elModelSpecialTrait.Maneuver1
                        If lManeuver > 0 Then lManeuver += 1
                    Case elModelSpecialTrait.Maneuver10
                        If lManeuver > 0 Then lManeuver += 10
                    Case elModelSpecialTrait.Maneuver10Critical1
                        If lManeuver > 0 Then lManeuver += 10
                    Case elModelSpecialTrait.Maneuver10Critical2
                        If lManeuver > 0 Then lManeuver += 10
                    Case elModelSpecialTrait.Maneuver2
                        If lManeuver > 0 Then lManeuver += 2
                    Case elModelSpecialTrait.Maneuver3
                        If lManeuver > 0 Then lManeuver += 3
                    Case elModelSpecialTrait.Maneuver30
                        If lManeuver > 0 Then lManeuver += 30
                    Case elModelSpecialTrait.Maneuver4
                        If lManeuver > 0 Then lManeuver += 4
                    Case elModelSpecialTrait.Maneuver5
                        If lManeuver > 0 Then lManeuver += 5
                    Case elModelSpecialTrait.Maneuver5Critical2
                        If lManeuver > 0 Then lManeuver += 5
                    Case elModelSpecialTrait.Speed1
                        If lSpeed > 0 Then lSpeed += 1
                    Case elModelSpecialTrait.Speed2
                        If lSpeed > 0 Then lSpeed += 2
                    Case elModelSpecialTrait.Speed5Critical2
                        If lSpeed > 0 Then lSpeed += 5
                    Case elModelSpecialTrait.SpeedAndManeuver1
                        If lManeuver > 0 Then lManeuver += 1
                        If lSpeed > 0 Then lSpeed += 1
                End Select 
            End If

            If lManeuver < 0 Then lManeuver = 0
            If lManeuver > 255 Then lManeuver = 255
            .Maneuver = CByte(lManeuver)

            If lSpeed < 0 Then lSpeed = 0
            If lSpeed > 255 Then lSpeed = 255
            .MaxSpeed = CByte(lSpeed)

            If .MaxSpeed > 0 AndAlso .Maneuver = 0 Then .Maneuver = 1

            .Structure_MaxHP = oHullTech.StructuralHitPoints

            .HullSize = oHullTech.HullSize
            Dim lCargoCap As Int32 = oHullTech.GetHullAllocation(HullTech.SlotConfig.eCargoBay)
            Dim lHangarCap As Int32 = oHullTech.GetHullAllocation(HullTech.SlotConfig.eHangar)
            Select Case lSpecTraitID
                Case elModelSpecialTrait.Cargo10
                    lCargoCap = CInt(lCargoCap * 1.1F)
                Case elModelSpecialTrait.Cargo20
                    lCargoCap = CInt(lCargoCap * 1.2F)
                Case elModelSpecialTrait.Cargo5
                    lCargoCap = CInt(lCargoCap * 1.05F)
                Case elModelSpecialTrait.CargoAndHangar10
                    lCargoCap = CInt(lCargoCap * 1.1F)
                    lHangarCap = CInt(lHangarCap * 1.1F)
                Case elModelSpecialTrait.CargoAndHangar3
                    lCargoCap = CInt(lCargoCap * 1.03F)
                    lHangarCap = CInt(lHangarCap * 1.03F)
                Case elModelSpecialTrait.CargoAndHangar6
                    lCargoCap = CInt(lCargoCap * 1.06F)
                    lHangarCap = CInt(lHangarCap * 1.06F)
                Case elModelSpecialTrait.Hangar10
                    lHangarCap = CInt(lHangarCap * 1.1F)
                Case elModelSpecialTrait.Hangar5
                    lHangarCap = CInt(lHangarCap * 1.05F)

                    'put radar's here too
                Case elModelSpecialTrait.ScanRange10
                    If lOptRange > 0 Then lOptRange += 10
                    If lMaxRange > 0 Then lMaxRange += 10
                Case elModelSpecialTrait.ScanRange15
                    If lOptRange > 0 Then lOptRange += 15
                    If lMaxRange > 0 Then lMaxRange += 15
                Case elModelSpecialTrait.Cargo200
                    lCargoCap = CInt(lCargoCap * 3)
            End Select
            .Cargo_Cap = lCargoCap
            .Hangar_Cap = lHangarCap
            .Fuel_Cap = oHullTech.GetHullAllocation(HullTech.SlotConfig.eFuelBay)

            If lWeaponAcc < 0 Then lWeaponAcc = 0
            If lWeaponAcc > 255 Then lWeaponAcc = 255
            .Weapon_Acc = CByte(lWeaponAcc)

            'all facilities receive a +50 PD bonus
            If oHullTech.yTypeID = 2 Then lScanRes += 50
            If lScanRes < 0 Then lScanRes = 0
            If lScanRes > 255 Then lScanRes = 255
            .ScanResolution = CByte(lScanRes)

            If lOptRange < 0 Then lOptRange = 0
            If lOptRange > 255 Then lOptRange = 255
            .OptRadarRange = CByte(lOptRange)

            If lMaxRange < 0 Then lMaxRange = 0
            If lMaxRange > 255 Then lMaxRange = 255
            .MaxRadarRange = CByte(lMaxRange)

            If lDisRes < 0 Then lDisRes = 0
            If lDisRes > 255 Then lDisRes = 255
            .DisruptionResistance = CByte(lDisRes)

            'Jamming abilities
            If oRadarTech Is Nothing = False AndAlso (oRadarTech.JamEffect <> 0 AndAlso oRadarTech.JamStrength <> 0) AndAlso lSpecTraitID <> elModelSpecialTrait.NoJammer Then
                .JamEffect = oRadarTech.JamEffect
                .JamImmunity = oRadarTech.JamImmunity
                .JamStrength = oRadarTech.JamStrength
                .JamTargets = oRadarTech.JamTargets
            Else
                .JamEffect = 0
                .JamImmunity = 0
                .JamStrength = 0
                .JamTargets = 0
            End If

            'armor resists
            If oArmorTech Is Nothing = False Then
                Dim lPR As Int32 = oArmorTech.yPiercingResist
                Dim lIR As Int32 = lImpactResist
                Dim lBR As Int32 = oArmorTech.yBeamResist
                Dim lMR As Int32 = oArmorTech.yMagneticResist
                Dim lFR As Int32 = oArmorTech.yBurnResist
                Dim lCR As Int32 = oArmorTech.yChemicalResist

                If lPR > 100 Then lPR = 100
                If lIR > 100 Then lIR = 100
                If lBR > 100 Then lBR = 100
                If lMR > 100 Then lMR = 100
                If lFR > 100 Then lFR = 100
                If lCR > 100 Then lCR = 100

                .PiercingResist = CByte(100 - lPR)
                .ImpactResist = CByte(100 - lIR)
                .BeamResist = CByte(100 - lBR)
                .ECMResist = CByte(100 - lMR)
                .FlameResist = CByte(100 - lFR)
                .ChemicalResist = CByte(100 - lCR)
                .DetectionResist = CByte(100 - oArmorTech.yRadarResist)
            End If

            'Shield Values
            .Shield_MaxHP = lShieldHP
            .ShieldRecharge = lRechargeRate
            .ShieldRechargeFreq = lRechargeFreq '* 30        'Multiply by 30 to change seconds to cycles

            .ModelID = oHullTech.ModelID

            Dim lFrontCrits As Int32 = 0
            Dim lLeftCrits As Int32 = 0
            Dim lRightCrits As Int32 = 0
            Dim lRearCrits As Int32 = 0

            'Fill the SideCrits of the entity def
            oHullTech.FillBaseCrits(lFrontCrits, lLeftCrits, lRearCrits, lRightCrits, oEngineTech Is Nothing = False, _
               oRadarTech Is Nothing = False, oShieldTech Is Nothing = False, lForeWpnCnt, _
               lLeftWpnCnt, lRearWpnCnt, lRightWpnCnt)

            .lSideCrits(UnitArcs.eForwardArc) = lFrontCrits
            .lSideCrits(UnitArcs.eLeftArc) = lLeftCrits
            .lSideCrits(UnitArcs.eRightArc) = lRightCrits
            .lSideCrits(UnitArcs.eBackArc) = lRearCrits

            'Now, define the EntityDefHangarDoors
            oHullTech.SetEntityDefsHangarDoors(moExpectedDef)

            If oHullTech.yTypeID = 2 Then
                'Facility...
                .RequiredProductionTypeID = ProductionType.eFacility

                Dim lProdFactor As Int32 = 0
                Dim lWorkerFactor As Int32 = 0
                Dim lPowerFactor As Int32 = 0

                'Ok, get the power factor
                Dim lPowerUsed As Int32 = 0
                If oRadarTech Is Nothing = False Then lPowerUsed += oRadarTech.PowerRequired
                If oShieldTech Is Nothing = False Then lPowerUsed += oShieldTech.PowerRequired
                lPowerUsed += oHullTech.PowerRequired

                For X As Int32 = 0 To WeaponUB
                    Dim oWpn As BaseWeaponTech = CType(Owner.GetTech(moWeapons(X).WeaponTechID, ObjectType.eWeaponTech), BaseWeaponTech)
                    If oWpn Is Nothing = False Then
                        lPowerUsed += oWpn.PowerRequired
                    End If
                Next X

                If oEngineTech Is Nothing = False Then lPowerFactor = oEngineTech.PowerProd * -1
                If lSpecTraitID = elModelSpecialTrait.PowerGen2 Then
                    lPowerFactor *= 2
                ElseIf lSpecTraitID = elModelSpecialTrait.PowerGen3 Then
                    lPowerFactor *= 3
                End If
                lPowerFactor += lPowerUsed

                'MSC: We only care about the mesh portion of the model here...
                Dim iMesh As Int16 = (.ModelID And 255S)

                Select Case oHullTech.ySubTypeID
                    Case 0  'CC
                        .ProductionTypeID = ProductionType.eCommandCenterSpecial
                        'lProdFactor = ProductionHull \ Owner.HullPerResidence
                        lProdFactor = ProductionHull \ Owner.oSpecials.HullPerResidence
                        If lProdFactor <> 0 Then
                            lWorkerFactor = 5000
                        Else
                            lWorkerFactor = 0
                            .ProductionTypeID = 0       'clear the production type
                        End If
                    Case 1  'mine
                        If (.ModelID And 255) <> 148 Then
                            .ProductionTypeID = ProductionType.eMining

                            'lWorkerFactor = oHullTech.GetHullAllocation(HullTech.SlotConfig.eCrewQuarters) \ Owner.MiningWorkers
                            'lWorkerFactor = ProductionHull \ Owner.MiningWorkers
                            'lWorkerFactor = ProductionHull \ Owner.oSpecials.MiningWorkers
                            lWorkerFactor = 0
                            'reduce the power factory by the hull tech's power required
                            lPowerFactor -= oHullTech.PowerRequired

                            'NOTE: the 5 is a constant we can manipulate
                            lProdFactor = CInt((ProductionHull / 100) / (5.0F + ((1.0F - RandomSeed) - 0.25F)))
                            lProdFactor = Math.Max(lProdFactor, 10)
                            'lPowerFactor += (lProdFactor * 4)
                        End If
                    Case 2  'other
                        'warehouse and defense, we only care about warehouse
                        If iMesh = 11 Then
                            .ProductionTypeID = ProductionType.eWareHouse
                            lWorkerFactor = .Hangar_Cap \ 1000
                        End If
                    Case 3  'Personnel
                        If iMesh = 6 Then
                            .ProductionTypeID = ProductionType.eEnlisted
                            'Dim lEnlCap As Int32 = oHullTech.GetHullAllocation(HullTech.SlotConfig.eCrewQuarters) \ 3
                            Dim lEnlCap As Int32 = ProductionHull \ 3
                            'Dim fEnlPerTrainer As Single = Owner.BaseEnlistedPerTrainer - ((RandomSeed - 0.5F) * 3.0F)
                            Dim fEnlPerTrainer As Single = Owner.oSpecials.BaseEnlistedPerTrainer - ((RandomSeed - 0.5F) * 3.0F)
                            lWorkerFactor = CInt(Math.Floor(lEnlCap / fEnlPerTrainer))
                            'lProdFactor = lWorkerFactor \ Owner.EnlistedTrainingFactor
                            lProdFactor = lWorkerFactor \ Owner.oSpecials.EnlistedTrainingFactor
                            lPowerFactor += (lProdFactor \ 2)
                        ElseIf iMesh = 7 Then
                            .ProductionTypeID = ProductionType.eOfficers

                            Dim lUnusedRadarHull As Int32 = oHullTech.GetHullAllocation(HullTech.SlotConfig.eRadar)
                            If oRadarTech Is Nothing = False Then
                                lUnusedRadarHull -= oRadarTech.HullRequired
                            End If

                            'NOTE: the 3 could be a constant to modify this data
                            'Dim lOffOut As Int32 = oHullTech.GetHullAllocation(HullTech.SlotConfig.eCrewQuarters) \ 3I
                            Dim lOffOut As Int32 = ProductionHull \ 3I

                            'lWorkerFactor = lOffOut \ Owner.BaseOfficerPerTrainer
                            lWorkerFactor = lOffOut \ Owner.oSpecials.BaseOfficerPerTrainer


                            lProdFactor = CInt(Math.Sqrt(lWorkerFactor * Owner.oSpecials.OfficerTrainingFactor))
                            ''Dim lWrkrHrs As Int32 = lWorkerFactor * Owner.OfficerTrainingFactor
                            'Dim lWrkrHrs As Int32 = lWorkerFactor * Owner.oSpecials.OfficerTrainingFactor
                            'Dim lRadarTime As Int32 = CInt(Math.Ceiling(lWrkrHrs * 0.2F))
                            'Dim fRadarTimeMult As Single = 1.0F
                            'If lRadarTime > lUnusedRadarHull Then fRadarTimeMult = CSng(lUnusedRadarHull / lRadarTime)
                            'fRadarTimeMult = 1.0F - fRadarTimeMult + (RandomSeed / 10.0F)
                            'Dim fTotTime As Single = lWrkrHrs * (1 + fRadarTimeMult)

                            'NOTE: 200 is another constant we can manipulate
                            'lProdFactor = CInt(Math.Floor((lWrkrHrs / fTotTime) * 200))
                            lPowerFactor += (lProdFactor * 2)
                        End If
                    Case 4  'Power Gen
                        .ProductionTypeID = ProductionType.ePowerCenter

                        lProdFactor = lPowerFactor
                        'lWorkerFactor = lProdFactor \ Owner.PowerGenProdToWorker
                        lWorkerFactor = lProdFactor \ Owner.oSpecials.PowerGenProdToWorker
                    Case 5  'Production
                        If iMesh = 13 OrElse iMesh = 102 Then
                            .ProductionTypeID = ProductionType.eAerialProduction
                        ElseIf iMesh = 137 Then
                            .ProductionTypeID = ProductionType.eNavalProduction
                        Else
                            .ProductionTypeID = ProductionType.eProduction
                        End If

                        'NOTE: the 6 is a constant we can manipulate
                        lWorkerFactor = ProductionHull \ 6
                        ''NOTE: the 10 is a constant we can manipulate
                        'Dim lWrkrsPerShift As Int32 = .Hangar_Cap \ 10
                        'If lWrkrsPerShift < 1 Then lWrkrsPerShift = 1
                        'Dim lCalcShifts As Int32 = CInt(Math.Ceiling(lWorkerFactor / lWrkrsPerShift))
                        'Dim lShifts As Int32 = Math.Min(CInt(Math.Ceiling(lWorkerFactor / lWrkrsPerShift)), 3I)
                        ''Get the actual shift workers
                        'If lCalcShifts <> lShifts Then lWrkrsPerShift = lWorkerFactor \ lShifts
                        ''NOTE: We can change 280 or 7 in this equation for altered results
                        'Dim lWeekHrs As Int32 = 280 - (7 * lShifts)
                        ''NOTE: 10080 is a constant we can manipulate
                        'lProdFactor = (lWeekHrs * lWrkrsPerShift) \ 10080
                        'If lShifts = 0 Then lProdFactor = 0
                        lProdFactor = ProductionHull \ 700

                        lPowerFactor += (lProdFactor * 11)
                    Case 6  'Refining
                        .ProductionTypeID = ProductionType.eRefining
                        'lProdFactor = 200
                        lProdFactor = Math.Max(ProductionHull \ 500, 1)
                        lWorkerFactor = 1200
                        lPowerFactor += lProdFactor
                    Case 7  'Research
                        .ProductionTypeID = ProductionType.eResearch

                        'NOTE: the 10 is a constant we can manipulate
                        'lWorkerFactor = (oHullTech.GetHullAllocation(HullTech.SlotConfig.eCrewQuarters) \ Owner.ResearchCrewQtrs) * 10
                        'lWorkerFactor = (ProductionHull \ Owner.ResearchCrewQtrs) * 10
                        lWorkerFactor = (ProductionHull \ Owner.oSpecials.ResearchCrewQtrs) * 10
                        'lProdFactor = lWorkerFactor \ Owner.ResearchProdFactor
                        lProdFactor = lWorkerFactor \ Owner.oSpecials.ResearchProdFactor

                        lPowerFactor += (lProdFactor * 15)
                    Case 8  'Residence
                        .ProductionTypeID = ProductionType.eColonists
                        'lProdFactor = oHullTech.GetHullAllocation(HullTech.SlotConfig.eCrewQuarters) \ Owner.HullPerResidence
                        'lProdFactor = ProductionHull \ Owner.HullPerResidence
                        lProdFactor = ProductionHull \ Owner.oSpecials.HullPerResidence
                        lWorkerFactor = lProdFactor \ 40
                        lPowerFactor += ((lProdFactor * 3) \ 10)
                    Case 9  'SpaceStation

                        'check if this is a tradepost...
                        If iMesh = 24 Then
                            .ProductionTypeID = ProductionType.eTradePost
                            lProdFactor = 1000
                            lWorkerFactor = 10000       'per slot
                            lPowerFactor = 0
                            '.Cargo_Cap += 1000000000
                            '.Hangar_Cap += 1000000000
                            '.lSideCrits(UnitArcs.eBackArc) = .lSideCrits(UnitArcs.eBackArc) Or elUnitStatus.eCargoBayOperational
                            '.lSideCrits(UnitArcs.eForwardArc) = .lSideCrits(UnitArcs.eForwardArc) Or elUnitStatus.eCargoBayOperational
                            '.lSideCrits(UnitArcs.eRightArc) = .lSideCrits(UnitArcs.eRightArc) Or elUnitStatus.eCargoBayOperational
                            '.lSideCrits(UnitArcs.eLeftArc) = .lSideCrits(UnitArcs.eLeftArc) Or elUnitStatus.eCargoBayOperational
                        Else
                            .ProductionTypeID = ProductionType.eSpaceStationSpecial
                            'TODO: Figure this out
                            lProdFactor = 1200
                            lWorkerFactor = 50000
                        End If
                    Case 10 'Tradepost
                        .ProductionTypeID = ProductionType.eTradePost
                        lProdFactor = 1000
                        lWorkerFactor = 10000
                        '.Cargo_Cap += 1000000000
                        '.Hangar_Cap += 1000000000
                        '.lSideCrits(UnitArcs.eBackArc) = .lSideCrits(UnitArcs.eBackArc) Or elUnitStatus.eCargoBayOperational
                        '.lSideCrits(UnitArcs.eForwardArc) = .lSideCrits(UnitArcs.eForwardArc) Or elUnitStatus.eCargoBayOperational
                        '.lSideCrits(UnitArcs.eRightArc) = .lSideCrits(UnitArcs.eRightArc) Or elUnitStatus.eCargoBayOperational
                        '.lSideCrits(UnitArcs.eLeftArc) = .lSideCrits(UnitArcs.eLeftArc) Or elUnitStatus.eCargoBayOperational
                        'For Z As Int32 = 0 To .lSideCrits.GetUpperBound(0)
                        '    .lSideCrits(Z) = .lSideCrits(Z) Or (elUnitStatus.eCargoBayOperational Or elUnitStatus.eHangarOperational)
                        'Next
                End Select

                If oHullTech.ySubTypeID <> 9 Then
                    .MaxSpeed = 0 : .Maneuver = 0
                End If

                'Now, set our resulting prod and worker factor
                CType(moExpectedDef, FacilityDef).ProdFactor = Math.Max(Math.Abs(lProdFactor), 1)
                CType(moExpectedDef, FacilityDef).WorkerFactor = Math.Abs(lWorkerFactor)
                CType(moExpectedDef, FacilityDef).PowerFactor = lPowerFactor

            Else
                'Unit... easy enough
                If (oHullTech.yChassisType And ChassisType.eGroundBased) <> 0 Then
                    .RequiredProductionTypeID = ProductionType.eProduction
                ElseIf (oHullTech.yChassisType And ChassisType.eAtmospheric) <> 0 Then
                    .RequiredProductionTypeID = ProductionType.eAerialProduction
                ElseIf (oHullTech.yChassisType And ChassisType.eNavalBased) <> 0 Then
                    .RequiredProductionTypeID = ProductionType.eNavalProduction
                Else : .RequiredProductionTypeID = ProductionType.eSpaceStationSpecial
                End If

                If oHullTech.yTypeID = 7 Then
                    'Utility
                    Select Case oHullTech.ySubTypeID
                        Case 0
                            .ProductionTypeID = ProductionType.eFacility
                        Case 1
                            .ProductionTypeID = ProductionType.eFacility
                        Case 3
                            .ProductionTypeID = ProductionType.eMining
                    End Select
                ElseIf oHullTech.yTypeID = 8 Then
                    If oHullTech.ySubTypeID = 6 Then .ProductionTypeID = ProductionType.eFacility
                End If

            End If

            'set our chassis type
            .yChassisType = oHullTech.yChassisType

            .ArmorIntegrity = 0
            If oArmorTech Is Nothing = False Then .ArmorIntegrity = oArmorTech.yIntegrityRoll()

            Dim lTmpFXColor As Int32 = 0
            If oEngineTech Is Nothing = False Then lTmpFXColor = oEngineTech.ColorValue * 16
            If oShieldTech Is Nothing = False Then lTmpFXColor += oShieldTech.ColorValue
            If lTmpFXColor > 255 Then lTmpFXColor = 255
            If lTmpFXColor < 0 Then lTmpFXColor = 0
            .yFXColors = CByte(lTmpFXColor)

            'TODO: Missing Items:
            'FuelEfficiency - Obsolete
            'MaxFacilitySize - Obsolete

            'Now, set its production cost
            .ProductionRequirements = New ProductionCost
        End With

        'now, our critical hit chances
        oHullTech.GetSlotCriticalChances(moExpectedDef)

        'If moProductionCost Is Nothing = True Then CalculateBothProdCosts()
        CalculateBothProdCosts()
        With moExpectedDef.ProductionRequirements
            .ColonistCost = moProductionCost.ColonistCost
            .CreditCost = moProductionCost.CreditCost
            .EnlistedCost = moProductionCost.EnlistedCost
            For X As Int32 = 0 To moProductionCost.ItemCostUB
                .AddProductionCostItem(moProductionCost.ItemCosts(X).ItemID, moProductionCost.ItemCosts(X).ItemTypeID, moProductionCost.ItemCosts(X).QuantityNeeded)
            Next X
            .ObjectID = moExpectedDef.ObjectID
            .ObjTypeID = moExpectedDef.ObjTypeID
            .OfficerCost = moProductionCost.OfficerCost
            .PointsRequired = moProductionCost.PointsRequired
            .ProductionCostType = 0
        End With

        'Now, fill the WeaponDefs
        Dim bFrontFirst As Boolean = True
        Dim bLeftFirst As Boolean = True
        Dim bRightFirst As Boolean = True
        Dim bRearFirst As Boolean = True

        Dim oTmpWpnDefs() As WeaponDef = Nothing
        Dim yTmpWpnDefGroups() As WeaponGroupType = Nothing
        Dim lTmpWpnDefUB As Int32 = -1
        Dim lTmpWpnDefIdx() As Int32 = Nothing        'the ID of the wEapon tech

        'Ok, now, put together our weapon defs
        For X As Int32 = 0 To Me.WeaponUB
            Dim bFound As Boolean = False
            For Y As Int32 = 0 To lTmpWpnDefUB
                If lTmpWpnDefIdx(Y) = moWeapons(X).WeaponTechID AndAlso yTmpWpnDefGroups(Y) = moWeapons(X).WeaponGroupTypeID Then
                    bFound = True
                    Exit For
                End If
            Next Y

            If bFound = False Then
                lTmpWpnDefUB += 1
                ReDim Preserve oTmpWpnDefs(lTmpWpnDefUB)
                ReDim Preserve lTmpWpnDefIdx(lTmpWpnDefUB)
                ReDim Preserve yTmpWpnDefGroups(lTmpWpnDefUB)

                lTmpWpnDefIdx(lTmpWpnDefUB) = moWeapons(X).WeaponTechID
                yTmpWpnDefGroups(lTmpWpnDefUB) = moWeapons(X).WeaponGroupTypeID

                Dim oWpn As BaseWeaponTech = CType(Owner.GetTech(moWeapons(X).WeaponTechID, ObjectType.eWeaponTech), BaseWeaponTech)
                If oWpn Is Nothing Then
                    'ok, check the ptk
                    oWpn = CType(Owner.GetPlayerTechKnowledgeTech(moWeapons(X).WeaponTechID, ObjectType.eWeaponTech, PlayerTechKnowledge.KnowledgeType.eSettingsLevel1), BaseWeaponTech)
                End If
                If oWpn Is Nothing = False Then
                    oTmpWpnDefs(lTmpWpnDefUB) = oWpn.GetWeaponDefResult()

                    'Now, adjust it by our values...
                    'rof, optrng, maxdmg
                    With oTmpWpnDefs(lTmpWpnDefUB)
                        .ROF = CShort(lWpnROF(X))
                        If lWpnOptRange(X) < 1 Then lWpnOptRange(X) = 1
                        If lWpnOptRange(X) > Int16.MaxValue Then lWpnOptRange(X) = Int16.MaxValue
                        .Range = CShort(lWpnOptRange(X))
                        '.BeamMaxDmg = CShort(lWpnMaxDmg(X))

                        If moWeapons(X).WeaponGroupTypeID = WeaponGroupType.PointDefenseWeaponGroup Then
                            .BeamMaxDmg \= 2 : .BeamMinDmg \= 2
                            .ChemicalMaxDmg \= 2 : .ChemicalMinDmg \= 2
                            .ECMMaxDmg \= 2 : .ECMMinDmg \= 2
                            .FlameMaxDmg \= 2 : .FlameMinDmg \= 2
                            .ImpactMaxDmg \= 2 : .ImpactMinDmg \= 2
                            .PiercingMaxDmg \= 2 : .PiercingMinDmg \= 2
                            Dim lAcc As Int32 = .Accuracy
                            lAcc = Math.Min(255, lAcc + 25)
                            .Accuracy = CByte(lAcc)
                            Dim lRng As Int32 = .Range
                            lRng = CInt(lRng * 0.25F)
                            If .Range > 0 And lRng = 0 Then lRng = 1
                            .Range = CShort(lRng)
                        End If
                    End With
                Else : lTmpWpnDefUB -= 1  'Should neverh appen...
                End If
            End If
        Next X

        Dim lDPSMult As Int32 = GetVersionDPSPenalty()

        moExpectedDef.WeaponDefUB = Me.WeaponUB
        ReDim moExpectedDef.WeaponDefs(Me.WeaponUB)
        For X As Int32 = 0 To Me.WeaponUB
            moExpectedDef.WeaponDefs(X) = New UnitWeaponDef
            With moExpectedDef.WeaponDefs(X)
                .mlAmmoCap = -1
                Dim oWpn As BaseWeaponTech = CType(Owner.GetTech(moWeapons(X).WeaponTechID, ObjectType.eWeaponTech), BaseWeaponTech)
                If oWpn Is Nothing Then
                    'ok, check the ptk
                    oWpn = CType(Owner.GetPlayerTechKnowledgeTech(moWeapons(X).WeaponTechID, ObjectType.eWeaponTech, PlayerTechKnowledge.KnowledgeType.eSettingsLevel1), BaseWeaponTech)
                End If
                'TODO: This was causing issues if fSpaceRemaining / oWpn.GetAmmoSize() results in an overflow
                'If oWpn.GetAmmoSize() > 0.0F Then
                '    Dim fSpaceRemaining As Single = oHullTech.GetWeaponHullAllotment(Me.moWeapons(X).SlotX, Me.moWeapons(X).SlotY) - oWpn.HullRequired
                '    Dim lAmmoCap As Int32 = CInt(fSpaceRemaining / oWpn.GetAmmoSize())
                '    If lAmmoCap < 1 Then .mlAmmoCap = 1 Else .mlAmmoCap = lAmmoCap
                'End If

                Select Case moWeapons(X).ArcID
                    Case UnitArcs.eBackArc, UnitArcs.eAllArcs
                        If bRearFirst = True Then
                            .lEntityStatusGroup = elUnitStatus.eAftWeapon1
                        Else : .lEntityStatusGroup = elUnitStatus.eAftWeapon2
                        End If
                        bRearFirst = Not bRearFirst
                    Case UnitArcs.eForwardArc
                        If bFrontFirst = True Then
                            .lEntityStatusGroup = elUnitStatus.eForwardWeapon1
                        Else : .lEntityStatusGroup = elUnitStatus.eForwardWeapon2
                        End If
                        bFrontFirst = Not bFrontFirst
                    Case UnitArcs.eLeftArc
                        If bLeftFirst = True Then
                            .lEntityStatusGroup = elUnitStatus.eLeftWeapon1
                        Else : .lEntityStatusGroup = elUnitStatus.eLeftWeapon2
                        End If
                        bLeftFirst = Not bLeftFirst
                    Case UnitArcs.eRightArc
                        If bRightFirst = True Then
                            .lEntityStatusGroup = elUnitStatus.eRightWeapon1
                        Else : .lEntityStatusGroup = elUnitStatus.eRightWeapon2
                        End If
                        bRightFirst = Not bRightFirst
                End Select

                .ObjectID = -1
                .ObjTypeID = ObjectType.eWeaponDef
                .oUnitDef = moExpectedDef

                .yArcID = moWeapons(X).ArcID

                For Y As Int32 = 0 To lTmpWpnDefUB
                    If lTmpWpnDefIdx(Y) = moWeapons(X).WeaponTechID AndAlso yTmpWpnDefGroups(Y) = moWeapons(X).WeaponGroupTypeID Then
                        .oWeaponDef = oTmpWpnDefs(Y)
                        Exit For
                    End If
                Next Y
                If .oWeaponDef Is Nothing AndAlso oWpn Is Nothing = False Then
                    .oWeaponDef = oWpn.GetWeaponDefResult()
                    If moWeapons(X).WeaponGroupTypeID = WeaponGroupType.PointDefenseWeaponGroup Then
                        With .oWeaponDef
                            .BeamMaxDmg \= 2 : .BeamMinDmg \= 2
                            .ChemicalMaxDmg \= 2 : .ChemicalMinDmg \= 2
                            .ECMMaxDmg \= 2 : .ECMMinDmg \= 2
                            .FlameMaxDmg \= 2 : .FlameMinDmg \= 2
                            .ImpactMaxDmg \= 2 : .ImpactMinDmg \= 2
                            .PiercingMaxDmg \= 2 : .PiercingMinDmg \= 2
                            Dim lAcc As Int32 = .Accuracy
                            lAcc = Math.Min(255, lAcc + 25)
                            .Accuracy = CByte(lAcc)
                            Dim lRng As Int32 = .Range
                            lRng = CInt(lRng * 0.25F)
                            If .Range > 0 And lRng = 0 Then lRng = 1
                            .Range = CShort(lRng)
                        End With
                    End If

                    If lDPSMult <> 100 Then
                        With .oWeaponDef
                            Dim fVal As Single = lDPSMult / 100.0F
                            .BeamMaxDmg = CInt(.BeamMaxDmg * fVal) : .BeamMinDmg = CInt(.BeamMinDmg * fVal)
                            .ChemicalMaxDmg = CInt(.ChemicalMaxDmg * fVal) : .ChemicalMinDmg = CInt(.ChemicalMinDmg * fVal)
                            .ECMMaxDmg = CInt(.ECMMaxDmg * fVal) : .ECMMinDmg = CInt(.ECMMinDmg * fVal)
                            .FlameMaxDmg = CInt(.FlameMaxDmg * fVal) : .FlameMinDmg = CInt(.FlameMinDmg * fVal)
                            .ImpactMaxDmg = CInt(.ImpactMaxDmg * fVal) : .ImpactMinDmg = CInt(.ImpactMinDmg * fVal)
                            .PiercingMaxDmg = CInt(.PiercingMaxDmg * fVal) : .PiercingMinDmg = CInt(.PiercingMinDmg * fVal)
                        End With
                    End If

                End If

                'TODO: Almost like this should be on the UnitDefWeapon object instead of the weapon def itself
                If .oWeaponDef Is Nothing = False Then .oWeaponDef.WpnGroup = moWeapons(X).WeaponGroupTypeID
            End With
        Next X

        If Me.Owner.yPlayerPhase = eyPlayerPhase.eInitialPhase AndAlso Me.Owner.lTutorialStep >= 164 AndAlso Me.Owner.lTutorialStep <= 174 Then
            moExpectedDef.Maneuver = 7
            moExpectedDef.MaxSpeed = CByte(Math.Max(moExpectedDef.MaxSpeed, 15))
        End If


    End Sub

    Protected Overrides Sub FinalizeResearch()
        If moExpectedDef Is Nothing Then ComponentDesigned()

        If moExpectedDef Is Nothing Then
            LogEvent(LogEventType.CriticalError, "FinalizeResearch on Prototype produced no Entity Def!")
            Return
        End If

        moExpectedDef.OwnerID = Me.Owner.ObjectID

        'Take the expecteddef and store it in our array... save it... send it to the region servers... and to the player?
        If moExpectedDef.SaveObject() = True Then
            'Ok... save the production cost...
            moExpectedDef.ProductionRequirements.ObjectID = moExpectedDef.ObjectID
            moExpectedDef.ProductionRequirements.ObjTypeID = moExpectedDef.ObjTypeID
            moExpectedDef.ProductionRequirements.ProductionCostType = 0
            If moExpectedDef.ProductionRequirements.SaveObject() = False Then
                LogEvent(LogEventType.CriticalError, "FinalizeResearch.Prototype's Def's Production Costs did not Save!")
                Return
            End If

            'create what remains when the entitydef is destroyed
            CreateEntityDefMinerals()

            Me.ResultingDef = moExpectedDef

            'Now, find a place in the global array
            Dim lIdx As Int32 = -1

            moExpectedDef.DataChanged()
            Me.DataChanged()

            If moExpectedDef.ObjTypeID = ObjectType.eFacilityDef Then
                For X As Int32 = 0 To glFacilityDefUB
                    If glFacilityDefIdx(X) = -1 Then
                        lIdx = X
                        Exit For
                    End If
                Next X

                If lIdx = -1 Then
                    ReDim Preserve goFacilityDef(glFacilityDefUB + 1)
                    ReDim Preserve glFacilityDefIdx(glFacilityDefUB + 1)
                    glFacilityDefUB += 1
                    lIdx = glFacilityDefUB
                End If

                goFacilityDef(lIdx) = CType(moExpectedDef, FacilityDef)
                glFacilityDefIdx(lIdx) = moExpectedDef.ObjectID

                Dim yOutMsg() As Byte = goMsgSys.GetAddObjectMessage(goFacilityDef(lIdx), GlobalMessageCode.eAddObjectCommand)
                goMsgSys.BroadcastToDomains(yOutMsg)

                Owner.SendPlayerMessage(yOutMsg, False, AliasingRights.eAddProduction Or AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)

                yOutMsg = goMsgSys.GetAddObjectMessage(moExpectedDef.ProductionRequirements, GlobalMessageCode.eAddObjectCommand)
                Owner.SendPlayerMessage(yOutMsg, False, AliasingRights.eAddProduction Or AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)

                'Dim yProdCost() As Byte = goFacilityDef(lIdx).ProductionRequirements.GetObjAsString()
                'Dim lTemp As Int32 = yOutMsg.GetUpperBound(0)
                'ReDim Preserve yOutMsg(yOutMsg.GetUpperBound(0) + yProdCost.GetUpperBound(0))
                'yProdCost.CopyTo(yOutMsg, lTemp + 1)
                'goMsgSys.SendMsgToOperator(yOutMsg)
            Else
                For X As Int32 = 0 To glUnitDefUB
                    If glUnitDefIdx(X) = -1 Then
                        lIdx = X
                        Exit For
                    End If
                Next X

                If lIdx = -1 Then
                    ReDim Preserve glUnitDefIdx(glUnitDefUB + 1)
                    ReDim Preserve goUnitDef(glUnitDefUB + 1)
                    glUnitDefUB += 1
                    lIdx = glUnitDefUB
                End If

                goUnitDef(lIdx) = moExpectedDef
                glUnitDefIdx(lIdx) = moExpectedDef.ObjectID

                Dim yOutMsg() As Byte = goMsgSys.GetAddObjectMessage(goUnitDef(lIdx), GlobalMessageCode.eAddObjectCommand)
                goMsgSys.BroadcastToDomains(yOutMsg)
                Owner.SendPlayerMessage(yOutMsg, False, AliasingRights.eAddProduction Or AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)
                yOutMsg = goMsgSys.GetAddObjectMessage(moExpectedDef.ProductionRequirements, GlobalMessageCode.eAddObjectCommand)
                Owner.SendPlayerMessage(yOutMsg, False, AliasingRights.eAddProduction Or AliasingRights.eAddResearch Or AliasingRights.eViewResearch Or AliasingRights.eViewTechDesigns)

                'Dim yProdCost() As Byte = goUnitDef(lIdx).ProductionRequirements.GetObjAsString()
                'Dim lTemp As Int32 = yOutMsg.GetUpperBound(0)
                'ReDim Preserve yOutMsg(yOutMsg.GetUpperBound(0) + yProdCost.GetUpperBound(0))
                'yProdCost.CopyTo(yOutMsg, lTemp + 1)
                'goMsgSys.SendMsgToOperator(yOutMsg)
            End If

            'on more save to verify it saved properly
            moExpectedDef.SaveObject()

            'ok, determine the prototype's overall DPS
            Dim decDPS As Decimal = 0D
            For X As Int32 = 0 To moExpectedDef.WeaponDefUB
                decDPS += moExpectedDef.WeaponDefs(X).oWeaponDef.GetOverallDPS()
            Next X
            Dim lMaxDPS As Int32
            Dim lMaxGuns As Int32
            TechBuilderComputer.GetTypeValues(HullTech.GetHullTypeID(oHullTech.yTypeID, oHullTech.ySubTypeID), 0, lMaxGuns, lMaxDPS, 0, 0, 0)
            lMaxDPS *= lMaxGuns
            If decDPS > lMaxDPS Then
                LogEvent(LogEventType.Warning, "Prototype designed with abnormally high DPS: " & Me.ObjectID & ". Player: " & Me.Owner.ObjectID)
            End If


        Else
            LogEvent(LogEventType.CriticalError, "FinalizeResearch.Prototype's Def did not Save!")
            Return
        End If
    End Sub

    Private Function CalculateActualHullUsed() As Int32
        Dim lResult As Int32 = 0

        If oEngineTech Is Nothing = False Then lResult += oEngineTech.HullRequired
        'If oHangarTech Is Nothing = False Then lResult += oHangarTech.HullRequired
        If oRadarTech Is Nothing = False Then lResult += oRadarTech.HullRequired
        If oShieldTech Is Nothing = False Then lResult += oShieldTech.HullRequired

        For X As Int32 = 0 To WeaponUB
            Dim oTmpWpn As BaseWeaponTech = CType(Owner.GetTech(moWeapons(X).WeaponTechID, ObjectType.eWeaponTech), BaseWeaponTech)
            If oTmpWpn Is Nothing = False Then lResult += oTmpWpn.HullRequired
        Next X

        If oArmorTech Is Nothing = False Then
            lResult += (ForeArmorUnits * oArmorTech.lHullUsagePerPlate)
            lResult += (LeftArmorUnits * oArmorTech.lHullUsagePerPlate)
            lResult += (RightArmorUnits * oArmorTech.lHullUsagePerPlate)
            lResult += (AftArmorUnits * oArmorTech.lHullUsagePerPlate)
        End If

        lResult += oHullTech.GetHullAllocation(HullTech.SlotConfig.eCargoBay)
        lResult += oHullTech.GetHullAllocation(HullTech.SlotConfig.eHangar)

        Return lResult
    End Function

    Public Overrides Function GetNonOwnerMsg() As Byte()
        Dim yResult(27) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yResult, lPos) : lPos += 2
        Me.GetGUIDAsString.CopyTo(yResult, lPos) : lPos += 6
        Me.PrototypeName.CopyTo(yResult, lPos) : lPos += 20

        Return yResult
    End Function

    Public Overrides Function TechnologyScore() As Integer
        Return 1000I
    End Function

    Public Overrides Function GetPlayerTechKnowledgeMsg(ByVal yTechLvl As PlayerTechKnowledge.KnowledgeType) As Byte()
        'TODO: ???
        Return Nothing
    End Function

    Private Sub CreateEntityDefMinerals()
        moExpectedDef.lEntityDefMineralUB = -1

        If oHullTech Is Nothing = False AndAlso oHullTech.GetProductionCost Is Nothing = False Then
            With oHullTech.GetProductionCost
                For X As Int32 = 0 To .ItemCostUB
                    If .ItemCosts(X).ItemTypeID = ObjectType.eMineral Then
                        moExpectedDef.AddEntityDefMineral(.ItemCosts(X).ItemID, .ItemCosts(X).QuantityNeeded \ 10)
                    End If
                Next X
            End With
        End If

        If oEngineTech Is Nothing = False AndAlso oEngineTech.GetProductionCost Is Nothing = False Then
            With oEngineTech.GetProductionCost
                For X As Int32 = 0 To .ItemCostUB
                    If .ItemCosts(X).ItemTypeID = ObjectType.eMineral Then
                        moExpectedDef.AddEntityDefMineral(.ItemCosts(X).ItemID, .ItemCosts(X).QuantityNeeded \ 10)
                    End If
                Next X
            End With
        End If
        If oRadarTech Is Nothing = False AndAlso oRadarTech.GetProductionCost Is Nothing = False Then
            With oRadarTech.GetProductionCost
                For X As Int32 = 0 To .ItemCostUB
                    If .ItemCosts(X).ItemTypeID = ObjectType.eMineral Then
                        moExpectedDef.AddEntityDefMineral(.ItemCosts(X).ItemID, .ItemCosts(X).QuantityNeeded \ 10)
                    End If
                Next X
            End With
        End If
        If oShieldTech Is Nothing = False AndAlso oShieldTech.GetProductionCost Is Nothing = False Then
            With oShieldTech.GetProductionCost
                For X As Int32 = 0 To .ItemCostUB
                    If .ItemCosts(X).ItemTypeID = ObjectType.eMineral Then
                        moExpectedDef.AddEntityDefMineral(.ItemCosts(X).ItemID, .ItemCosts(X).QuantityNeeded \ 10)
                    End If
                Next X
            End With
        End If

        For lWpnIdx As Int32 = 0 To WeaponUB
            Dim oWpn As BaseWeaponTech = CType(Owner.GetTech(moWeapons(lWpnIdx).WeaponTechID, ObjectType.eWeaponTech), BaseWeaponTech)
            If oWpn Is Nothing Then oWpn = CType(Owner.GetPlayerTechKnowledgeTech(moWeapons(lWpnIdx).WeaponTechID, ObjectType.eWeaponTech, PlayerTechKnowledge.KnowledgeType.eSettingsLevel2), BaseWeaponTech)
            If oWpn Is Nothing = False AndAlso oWpn.GetProductionCost Is Nothing = False Then
                With oWpn.GetProductionCost
                    For X As Int32 = 0 To .ItemCostUB
                        If .ItemCosts(X).ItemTypeID = ObjectType.eMineral Then
                            moExpectedDef.AddEntityDefMineral(.ItemCosts(X).ItemID, .ItemCosts(X).QuantityNeeded \ 10)
                        End If
                    Next X
                End With
            End If
        Next lWpnIdx
    End Sub

    Public Sub CreateDismantledCaches(ByRef oDomainSocket As NetSock, ByVal lStatus As Int32, ByVal lLocX As Int32, ByVal lLocZ As Int32, ByVal lParentID As Int32, ByVal iParentTypeID As Int16, ByVal lArmorSide1 As Int32, ByVal lArmorSide2 As Int32, ByVal lArmorSide3 As Int32, ByVal lArmorSide4 As Int32, ByVal yCacheType As Byte)
        'If oArmorTech Is Nothing = False AndAlso oArmorTech.Owner Is Nothing = False Then
        '    Dim lPlates As Int32 = 0
        '    lPlates += (lArmorSide1 \ oArmorTech.lHPPerPlate)
        '    lPlates += (lArmorSide2 \ oArmorTech.lHPPerPlate)
        '    lPlates += (lArmorSide3 \ oArmorTech.lHPPerPlate)
        '    lPlates += (lArmorSide4 \ oArmorTech.lHPPerPlate)
        '    Dim lCacheIdx As Int32 = AddComponentCache(lParentID, iParentTypeID, lPlates, lLocX, lLocZ, oArmorTech.ObjectID, oArmorTech.ObjTypeID, oArmorTech.Owner.ObjectID, yCacheType)
        '    If lCacheIdx <> -1 Then
        '        Dim yData() As Byte = goMsgSys.GetAddObjectMessage(goComponentCache(lCacheIdx), GlobalMessageCode.eAddObjectCommand)
        '        If yData Is Nothing = False Then oDomainSocket.SendData(yData)
        '    End If
        'End If

        If oEngineTech Is Nothing = False AndAlso oEngineTech.Owner Is Nothing = False AndAlso (lStatus And elUnitStatus.eEngineOperational) <> 0 Then
            Dim lCacheIdx As Int32 = AddComponentCache(lParentID, iParentTypeID, 1, lLocX, lLocZ, oEngineTech.ObjectID, oEngineTech.ObjTypeID, oEngineTech.Owner.ObjectID, yCacheType)
            If lCacheIdx <> -1 Then
                Dim yData() As Byte = goMsgSys.GetAddObjectMessage(goComponentCache(lCacheIdx), GlobalMessageCode.eAddObjectCommand)
                If yData Is Nothing = False Then oDomainSocket.SendData(yData)
            End If
        End If

        If oRadarTech Is Nothing = False AndAlso oRadarTech.Owner Is Nothing = False AndAlso (lStatus And elUnitStatus.eRadarOperational) <> 0 Then
            Dim lCacheIdx As Int32 = AddComponentCache(lParentID, iParentTypeID, 1, lLocX, lLocZ, oRadarTech.ObjectID, oRadarTech.ObjTypeID, oRadarTech.Owner.ObjectID, yCacheType)
            If lCacheIdx <> -1 Then
                Dim yData() As Byte = goMsgSys.GetAddObjectMessage(goComponentCache(lCacheIdx), GlobalMessageCode.eAddObjectCommand)
                If yData Is Nothing = False Then oDomainSocket.SendData(yData)
            End If
        End If

        If oShieldTech Is Nothing = False AndAlso oShieldTech.Owner Is Nothing = False AndAlso (lStatus And elUnitStatus.eShieldOperational) <> 0 Then
            Dim lCacheIdx As Int32 = AddComponentCache(lParentID, iParentTypeID, 1, lLocX, lLocZ, oShieldTech.ObjectID, oShieldTech.ObjTypeID, oShieldTech.Owner.ObjectID, yCacheType)
            If lCacheIdx <> -1 Then
                Dim yData() As Byte = goMsgSys.GetAddObjectMessage(goComponentCache(lCacheIdx), GlobalMessageCode.eAddObjectCommand)
                If yData Is Nothing = False Then oDomainSocket.SendData(yData)
            End If
        End If

        If ResultingDef Is Nothing = False Then
            For X As Int32 = 0 To ResultingDef.WeaponDefUB
                If (ResultingDef.WeaponDefs(X).lEntityStatusGroup And lStatus) <> 0 Then
                    With ResultingDef.WeaponDefs(X)
                        If .oWeaponDef Is Nothing = False AndAlso .oWeaponDef.RelatedWeapon Is Nothing = False Then
                            If .oWeaponDef.RelatedWeapon.Owner Is Nothing = False AndAlso .oWeaponDef.RelatedWeapon.Owner.ObjectID <> 0 Then
                                With .oWeaponDef.RelatedWeapon
                                    Dim lCacheIdx As Int32 = AddComponentCache(lParentID, iParentTypeID, 1, lLocX, lLocZ, .ObjectID, .ObjTypeID, .Owner.ObjectID, yCacheType)
                                    If lCacheIdx <> -1 Then
                                        Dim yData() As Byte = goMsgSys.GetAddObjectMessage(goComponentCache(lCacheIdx), GlobalMessageCode.eAddObjectCommand)
                                        If yData Is Nothing = False Then oDomainSocket.SendData(yData)
                                    End If
                                End With
                            End If
                        End If
                    End With
                End If
            Next X
        End If
    End Sub

    Public Function GetRandomWeaponComponentOnSide(ByVal ySide As UnitArcs) As BaseWeaponTech
        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To Me.WeaponUB
            If moWeapons(X).ArcID = ySide Then
                lCnt += 1
            End If
        Next X
        If lCnt <> 0 Then
            lCnt = CInt(Rnd() * lCnt)
            For X As Int32 = 0 To Me.WeaponUB
                If moWeapons(X).ArcID = ySide Then
                    If lCnt = 0 Then
                        Return CType(Owner.GetTech(moWeapons(X).WeaponTechID, ObjectType.eWeaponTech), BaseWeaponTech)
                    Else : lCnt -= 1
                    End If
                End If
            Next X
        End If
        Return Nothing
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

            ReDim .PrototypeName(19)
            Array.Copy(yData, lPos, .PrototypeName, 0, 20) : lPos += 20

            'System.BitConverter.GetBytes(lEngineTech).CopyTo(mySendString, lPos) : lPos += 4
            .lEngineTech = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lArmorTech = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lHullTech = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lRadarTech = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lShieldTech = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ForeArmorUnits = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .AftArmorUnits = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .LeftArmorUnits = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .RightArmorUnits = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ProductionHull = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            '.MinCrew = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .MaxCrew = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            WeaponUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
            ReDim moWeapons(WeaponUB)
            For X As Int32 = 0 To WeaponUB
                With moWeapons(X)
                    .WeaponTechID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .SlotX = yData(lPos) : lPos += 1
                    .SlotY = yData(lPos) : lPos += 1
                    .WeaponGroupTypeID = CType(yData(lPos), WeaponGroupType) : lPos += 1
                End With
            Next X

            If .Owner Is Nothing = False Then
                Dim lFirstIdx As Int32 = -1
                For X As Int32 = 0 To .Owner.mlPrototypeUB
                    If .Owner.mlPrototypeIdx(X) = .ObjectID Then
                        .Owner.moPrototype(X) = Me
                        Return
                    ElseIf lFirstIdx = -1 AndAlso .Owner.mlPrototypeIdx(X) = -1 Then
                        lFirstIdx = X
                    End If
                Next X

                If lFirstIdx = -1 Then
                    lFirstIdx = .Owner.mlPrototypeUB + 1
                    ReDim Preserve .Owner.mlPrototypeIdx(lFirstIdx)
                    ReDim Preserve .Owner.moPrototype(lFirstIdx)
                    .Owner.mlPrototypeUB = lFirstIdx
                End If
                .Owner.moPrototype(lFirstIdx) = Me
                .Owner.mlPrototypeIdx(lFirstIdx) = Me.ObjectID
            End If
 
        End With

        
    End Sub

    Public Sub CheckForMicroTech()
        Dim bHasMicroTech As Boolean = False

        'If moExpectedDef Is Nothing = False Then
        '    If (moExpectedDef.lExtendedFlags And Epica_Entity_Def.eExtendedFlagValues.ContainsMicroTech) <> 0 Then
        '        moExpectedDef.lExtendedFlags = moExpectedDef.lExtendedFlags Xor Epica_Entity_Def.eExtendedFlagValues.ContainsMicroTech
        '        moExpectedDef.SaveObject()
        '    End If
        'End If

        'If oArmorTech Is Nothing = False AndAlso (oArmorTech.MajorDesignFlaw And Epica_Tech.eComponentDesignFlaw.eMicroTech) = Epica_Tech.eComponentDesignFlaw.eMicroTech Then
        '    bHasMicroTech = True
        'End If
        'If oShieldTech Is Nothing = False AndAlso (oShieldTech.MajorDesignFlaw And Epica_Tech.eComponentDesignFlaw.eMicroTech) = Epica_Tech.eComponentDesignFlaw.eMicroTech Then
        '    bHasMicroTech = True
        'End If
        'If oEngineTech Is Nothing = False AndAlso (oEngineTech.MajorDesignFlaw And Epica_Tech.eComponentDesignFlaw.eMicroTech) = Epica_Tech.eComponentDesignFlaw.eMicroTech Then
        '    bHasMicroTech = True
        'End If
        'If oRadarTech Is Nothing = False AndAlso (oRadarTech.MajorDesignFlaw And Epica_Tech.eComponentDesignFlaw.eMicroTech) = Epica_Tech.eComponentDesignFlaw.eMicroTech Then
        '    bHasMicroTech = True
        'End If

        For X As Int32 = 0 To WeaponUB
            Dim oWpn As BaseWeaponTech = CType(Owner.GetTech(moWeapons(X).WeaponTechID, ObjectType.eWeaponTech), BaseWeaponTech)
            If oWpn Is Nothing Then
                'ok, check the ptk
                oWpn = CType(Owner.GetPlayerTechKnowledgeTech(moWeapons(X).WeaponTechID, ObjectType.eWeaponTech, PlayerTechKnowledge.KnowledgeType.eSettingsLevel1), BaseWeaponTech)
            End If
            If oWpn Is Nothing = False Then
                If (oWpn.MajorDesignFlaw And Epica_Tech.eComponentDesignFlaw.eMicroTech) = Epica_Tech.eComponentDesignFlaw.eMicroTech Then
                    bHasMicroTech = True
                End If
            End If
        Next X

        If bHasMicroTech = True Then
            If moExpectedDef Is Nothing = False Then
                Dim lFlags As Int32 = moExpectedDef.lExtendedFlags
                moExpectedDef.lExtendedFlags = moExpectedDef.lExtendedFlags Or Epica_Entity_Def.eExtendedFlagValues.ContainsMicroTech
                If moExpectedDef.lExtendedFlags <> lFlags Then moExpectedDef.SaveObject()
            End If
            'Else
            '    If moExpectedDef Is Nothing = False Then
            '        Dim lFlags As Int32 = moExpectedDef.lExtendedFlags
            '        If (lFlags And Epica_Entity_Def.eExtendedFlagValues.ContainsMicroTech) <> 0 Then
            '            lFlags = lFlags Xor Epica_Entity_Def.eExtendedFlagValues.ContainsMicroTech
            '            moExpectedDef.lExtendedFlags = lFlags
            '            moExpectedDef.SaveObject()
            '        End If
            '    End If
        End If
    End Sub
End Class
