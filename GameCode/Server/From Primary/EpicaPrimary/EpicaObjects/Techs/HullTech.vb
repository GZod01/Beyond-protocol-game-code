Public Class HullTech
    Inherits Epica_Tech

    Private Enum eyHullRequirements As Byte
        RequiresAnEngine = 1
        EngineIsToBeInRear = 2
        RequiresFiftyThousandStructHP = 4
        RequiresBayDoor = 8
    End Enum

    Private Const ml_GRID_SIZE_WH As Int32 = 30

    'NOTE: Should mimic what the client has in UIHullSlots.vb
    Public Enum SlotType As Byte
        eUnused = 0
        eFront = 1
        eLeft = 2
        eRear = 3
        eRight = 4
        eAllArc = 5

        eNoChange = 255
    End Enum
    Public Enum SlotConfig As Byte
        eArmorConfig = 0
        eCrewQuarters = 1
        eCargoBay = 2
        eRadar = 3
        eEngines = 4
        eHangar = 5
        eShields = 6
        eWeapons = 7
        eFuelBay = 8
        eHangarDoor = 9
    End Enum

    Private Structure HullSlot
        Public SlotConfig As SlotConfig
        Public SlotGroup As Byte
        Public SlotType As SlotType

        Public bVerified As Boolean         'Used in validation
    End Structure

    Private Structure HullValPair
        Public ValueConfig As SlotConfig
        Public ValueGroupNum As Int32
        Public Value As Int32
    End Structure

    Public HullName(19) As Byte
    Private miModelID As Short
    Public HullSize As Int32
    Public StructuralMineralID As Int32
    Public StructuralHitPoints As Int32

    Public yTypeID As Byte
    Public ySubTypeID As Byte

    Public yChassisType As Byte

    Private moSlots(,) As HullSlot

    Private mlPowerRequired As Int32 = -1
    Public ReadOnly Property PowerRequired() As Int32
        Get
            If mlPowerRequired = -1 Then
                If Me.StructureMineral Is Nothing Then Return 0

                'Get power requirement for cargo cap
                Dim lCargoCap As Int32 = 0
                If (yChassisType And ChassisType.eSpaceBased) <> 0 Then
                    lCargoCap = GetHullAllocation(SlotConfig.eCargoBay)
                    If lCargoCap <> 0 Then
                        lCargoCap = CInt(Math.Ceiling(Math.Sqrt(lCargoCap) * 1.4F))
                    End If
                ElseIf (yChassisType And ChassisType.eNavalBased) <> 0 Then
                    lCargoCap = GetHullAllocation(SlotConfig.eCargoBay)
                    If lCargoCap <> 0 Then
                        lCargoCap = CInt(Math.Ceiling(Math.Sqrt(lCargoCap) * 1.1F))
                    End If
                End If

                'Get power requirement for hangar
                Dim lHangarCap As Int32 = GetHullAllocation(SlotConfig.eHangar)
                If lHangarCap = 0 Then
                    mlPowerRequired = CInt(lCargoCap * (RandomSeed + 0.4F))
                    Return mlPowerRequired
                Else : lHangarCap = CInt(Math.Ceiling(Math.Sqrt(lHangarCap) * 0.8F))
                End If

                'Now, go through all the doors
                Dim lGroup() As Int32 = Nothing
                Dim lSize() As Int32 = Nothing
                Dim lGroupUB As Int32 = -1
                Dim lTotal As Int32

                'Load our door groups
                For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                        If moSlots(X, Y).SlotType <> SlotType.eUnused Then
                            lTotal += 1

                            'Ok, now, is this a HangarDoor?
                            If moSlots(X, Y).SlotConfig = SlotConfig.eHangarDoor Then
                                'Ok, it is... find our group
                                Dim lIdx As Int32 = -1
                                For lTempIdx As Int32 = 0 To lGroupUB
                                    If lGroup(lTempIdx) = moSlots(X, Y).SlotGroup Then
                                        lIdx = lTempIdx
                                        Exit For
                                    End If
                                Next lTempIdx

                                If lIdx = -1 Then
                                    lGroupUB += 1
                                    ReDim Preserve lGroup(lGroupUB)
                                    ReDim Preserve lSize(lGroupUB)
                                    lIdx = lGroupUB
                                End If

                                lGroup(lIdx) = moSlots(X, Y).SlotGroup
                                lSize(lIdx) += 1
                            End If
                        End If
                    Next X
                Next Y

                'Get our HullPerSlot
                Dim fHullPerSlot As Single = CSng(HullSize / lTotal)
                Dim lDoorPower As Int32 = 0
                Dim lStructDens As Int32 = StructureMineral.GetPropertyValue(eMinPropID.Density)

                For lIdx As Int32 = 0 To lGroupUB
                    If lSize(lIdx) <> 0 Then
                        'Now, go through the doors and figure out their power
                        Dim fScore As Single = ((lSize(lIdx) * (lStructDens / 256.0F)) * lSize(lIdx)) / 2.0F
                        lDoorPower += CInt(Math.Ceiling(fScore / 100.0F))
                    End If
                Next lIdx

                mlPowerRequired = CInt((lCargoCap + lHangarCap + lDoorPower) * (RandomSeed + 0.4F))
            End If
            Return mlPowerRequired
        End Get
    End Property
    Public Sub ClearPowerRequired()
        mlPowerRequired = -1
    End Sub

    Private mySendString() As Byte

    Public Function GetObjAsString() As Byte()
        If moResearchCost Is Nothing OrElse moProductionCost Is Nothing Then
            CalculateBothProdCosts()
            mbStringReady = False
        End If

        'here we will return the entire object as a string
        'If mbStringReady = False Then
        Dim lCnt As Int32 = 0

        'Get our count of items to send...
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).SlotType <> SlotType.eUnused AndAlso moSlots(X, Y).SlotConfig <> SlotConfig.eArmorConfig Then lCnt += 1
            Next X
        Next Y

        ReDim mySendString(Epica_Tech.BASE_OBJ_STRING_SIZE + 44 + (lCnt * 4))

        Dim lPos As Int32
        MyBase.GetBaseObjAsString(False).CopyTo(mySendString, 0)
        lPos = Epica_Tech.BASE_OBJ_STRING_SIZE

        HullName.CopyTo(mySendString, lPos) : lPos += 20
        System.BitConverter.GetBytes(HullSize).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(StructuralMineralID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(StructuralHitPoints).CopyTo(mySendString, lPos) : lPos += 4
        mySendString(lPos) = yTypeID : lPos += 1
        mySendString(lPos) = ySubTypeID : lPos += 1
        System.BitConverter.GetBytes(ModelID).CopyTo(mySendString, lPos) : lPos += 2
        mySendString(lPos) = yChassisType : lPos += 1
        System.BitConverter.GetBytes(PowerRequired).CopyTo(mySendString, lPos) : lPos += 4

        System.BitConverter.GetBytes(lCnt).CopyTo(mySendString, lPos) : lPos += 4

        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).SlotType <> SlotType.eUnused AndAlso moSlots(X, Y).SlotConfig <> SlotConfig.eArmorConfig Then
                    mySendString(lPos) = CByte(X) : lPos += 1
                    mySendString(lPos) = CByte(Y) : lPos += 1
                    mySendString(lPos) = CByte(moSlots(X, Y).SlotConfig) : lPos += 1
                    mySendString(lPos) = CByte(moSlots(X, Y).SlotGroup) : lPos += 1
                End If
            Next X
        Next Y

        'mbStringReady = True
        ' End If
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
                sSQL = "INSERT INTO tblHull (OwnerID, ModelID, HullSize, StructuralMineralID, " & _
                  "StructuralHitPoints, HullName, TypeID, SubTypeID, ChassisType" & _
                  ", CurrentSuccessChance, SuccessChanceIncrement, ResearchAttempts, RandomSeed, ResPhase" & _
                  ", PopIntel, bArchived, VersionNumber) VALUES (" & _
                  Owner.ObjectID & ", " & ModelID & ", " & HullSize & ", " & _
                  StructuralMineralID & ", " & StructuralHitPoints & ", '" & MakeDBStr(BytesToString(HullName)) & "', " & _
                  yTypeID & ", " & ySubTypeID & ", " & yChassisType & _
                  ", " & CurrentSuccessChance & ", " & SuccessChanceIncrement & ", " & ResearchAttempts & _
                  ", " & RandomSeed & ", " & ComponentDevelopmentPhase & ", " & PopIntel & ", " & yArchived & ", " & Me.lVersionNum & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblHull SET OwnerID = " & Owner.ObjectID & ", ModelID = " & ModelID & ", HullSize = " & _
                  HullSize & ", StructuralMineralID = " & _
                  StructuralMineralID & ", StructuralHitPoints = " & StructuralHitPoints & ", HullName = '" & _
                  MakeDBStr(BytesToString(HullName)) & "', TypeID = " & yTypeID & _
                  ", SubTypeID = " & ySubTypeID & ", ChassisType = " & yChassisType & _
                  ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
                  SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & _
                  RandomSeed & ", ResPhase = " & ComponentDevelopmentPhase & ", PopIntel = " & PopIntel & _
                  ", bArchived = " & yArchived & ", VersionNumber = " & Me.lVersionNum & " WHERE HullID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(HullID) FROM tblHull WHERE HullName = '" & MakeDBStr(BytesToString(HullName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            MyBase.FinalizeSave()

            'Clear out the hardpoints for this hull...
            sSQL = "DELETE FROM tblHardPoint WHERE HullID = " & ObjectID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing

            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    If moSlots(X, Y).SlotConfig <> SlotConfig.eArmorConfig AndAlso moSlots(X, Y).SlotType <> SlotType.eUnused Then
                        Try
                            sSQL = "INSERT INTO tblHardPoint (HullID, SlotX, SlotY, SlotConfig, SlotGroup) VALUES (" & _
                               Me.ObjectID & ", " & X & ", " & Y & ", " & CInt(moSlots(X, Y).SlotConfig) & ", " & CInt(moSlots(X, Y).SlotGroup) & ")"
                            oComm = New OleDb.OleDbCommand(sSQL, goCN)
                            oComm.ExecuteNonQuery()
                        Catch ex As Exception
                            LogEvent(LogEventType.CriticalError, "Unable to save Hardpoint for Hull ID " & Me.ObjectID & " at slot (" & X & ", " & Y & ").")
                        Finally
                            oComm.Dispose()
                            oComm = Nothing
                        End Try
                    End If
                Next X
            Next Y

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
            sSQL = "UPDATE tblHull SET OwnerID = " & Owner.ObjectID & ", ModelID = " & ModelID & ", HullSize = " & _
              HullSize & ", StructuralMineralID = " & _
              StructuralMineralID & ", StructuralHitPoints = " & StructuralHitPoints & ", HullName = '" & _
              MakeDBStr(BytesToString(HullName)) & "', TypeID = " & yTypeID & _
              ", SubTypeID = " & ySubTypeID & ", ChassisType = " & yChassisType & _
              ", CurrentSuccessChance = " & CurrentSuccessChance & ", SuccessChanceIncrement = " & _
              SuccessChanceIncrement & ", ResearchAttempts = " & ResearchAttempts & ", RandomSeed = " & _
              RandomSeed & ", ResPhase = " & ComponentDevelopmentPhase & ", PopIntel = " & PopIntel & _
              ", bArchived = " & yArchived & " WHERE HullID = " & ObjectID
            oSB.AppendLine(sSQL)

            oSB.AppendLine(MyBase.GetFinalizeSaveText())

            'Clear out the hardpoints for this hull...
            sSQL = "DELETE FROM tblHardPoint WHERE HullID = " & ObjectID
            oSB.AppendLine(sSQL)

            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    If moSlots(X, Y).SlotConfig <> SlotConfig.eArmorConfig AndAlso moSlots(X, Y).SlotType <> SlotType.eUnused Then
                        Try
                            sSQL = "INSERT INTO tblHardPoint (HullID, SlotX, SlotY, SlotConfig, SlotGroup) VALUES (" & _
                               Me.ObjectID & ", " & X & ", " & Y & ", " & CInt(moSlots(X, Y).SlotConfig) & ", " & CInt(moSlots(X, Y).SlotGroup) & ")"
                            oSB.AppendLine(sSQL)
                        Catch ex As Exception
                            LogEvent(LogEventType.CriticalError, "Unable to save Hardpoint for Hull ID " & Me.ObjectID & " at slot (" & X & ", " & Y & ").")
                        End Try
                    End If
                Next X
            Next Y

            Return oSB.ToString
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        End Try
        Return ""
    End Function

    Public Overrides Function ValidateDesign() As Boolean

        Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered

        If StructureMineral Is Nothing Then
            Me.ErrorReasonCode = TechBuilderErrorReason.eHull_HardnessZero
            LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign: StructureMineral is nothing. Owner is: " & Me.Owner.ObjectID)
            Return False
        End If

        Dim bHasNavalFac As Boolean = Owner.oSpecials.yNavalAbility > 1
        Dim bHasNavalUnit As Boolean = Owner.oSpecials.yNavalAbility > 0

        'Follows the Client.UIHullSlots.HasErrorList rules...
        Dim lCrewReqVal As Int32 = 20
        With goModelDefs.GetModelDef(ModelID)
            If BytesToString(.FrameName).ToUpper.StartsWith("UNNAMED") = True Then Return False
            If HullSize < .MinHull OrElse HullSize > .MaxHull Then
                LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign. Player using a hull size value outside of range: " & Owner.ObjectID)
                Return False
            End If

            If .bRequiresClaim = True Then
                Dim bClaimed As Boolean = False
                Dim lCompareModelID As Int32 = (ModelID And 255)
                For X As Int32 = 0 To Me.Owner.Claimables.GetUpperBound(0)
                    If Owner.Claimables(X).lID = lCompareModelID AndAlso Owner.Claimables(X).iTypeID = ObjectType.eHullTech AndAlso (Owner.Claimables(X).yClaimFlag And eyClaimFlags.eClaimed) <> 0 Then
                        bClaimed = True
                        Exit For
                    End If
                Next X
                If bClaimed = False AndAlso .ModelID = 141 Then
                    If Me.Owner.lGuildID = 13 OrElse Me.Owner.lGuildID = 14 OrElse Me.Owner.lGuildID = 17 Then
                        bClaimed = True
                    End If
                End If
                If bClaimed = False Then '
                    LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign. Player using claimable hull without claiming it: " & Owner.ObjectID)
                    Return False
                End If
            End If

            lCrewReqVal = .CrewReqPerc

            If .FrameTypeID = 3 Then
                'naval
                If .TypeID = 2 Then
                    'fac
                    'If bHasNavalFac = False AndAlso (bHasNavalUnit = False OrElse (.ModelID And 255) <> 137) Then
                    If bHasNavalFac = False AndAlso (bHasNavalUnit = False OrElse .SubTypeID <> 5) Then
                        LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign. Player using naval fac but does not have special. " & Owner.ObjectID)
                        Return False
                    End If
                Else
                    'unit
                    If bHasNavalUnit = False Then
                        LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign. Player using naval unit but does not have special. " & Owner.ObjectID)
                        Return False
                    End If
                End If
            End If

            If Owner.oSpecials.lPowerThrustLimit < .MinHull AndAlso .TypeID <> 2 Then
                LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign. Player's PowerThrustLimit is less than Min Hull and TypeID is not a facility. Player: " & Owner.ObjectID)
                Return False
            End If
        End With

        Dim bShip As Boolean = False
        bShip = (yTypeID = 0 OrElse yTypeID = 1 OrElse yTypeID = 3 OrElse yTypeID = 6)

        Dim lCrewQuarters As Int32 = 0
        Dim bInsuffCrewQuarters As Boolean = False
        Dim fHullPerSlot As Single = CSng(HullSize / Me.TotalSlots)
        Dim bEngineOnRearEdge As Boolean = False

        Dim bHasEngine As Boolean = False
        Dim bHasHangarDoor As Boolean = False

        Dim lMaxWpnSlot As Int32 = 0

        Dim lWpnGrpsInAllArcs(-1) As Int32

        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                With moSlots(X, Y)
                    If .SlotType <> SlotType.eUnused Then
                        Select Case .SlotConfig
                            Case SlotConfig.eCrewQuarters
                                lCrewQuarters += 1
                            Case SlotConfig.eEngines
                                bHasEngine = True
                                If Y < ml_GRID_SIZE_WH - 1 AndAlso .SlotType = SlotType.eRear Then
                                    If moSlots(X, Y + 1).SlotType = SlotType.eUnused Then
                                        bEngineOnRearEdge = True
                                    End If
                                ElseIf Y = ml_GRID_SIZE_WH - 1 AndAlso .SlotType = SlotType.eRear Then
                                    bEngineOnRearEdge = True
                                End If
                            Case SlotConfig.eArmorConfig
                                If .SlotType = SlotType.eAllArc Then Return False 'cannot have armor in an all-arc
                            Case SlotConfig.eHangarDoor
                                bHasHangarDoor = True
                                If .SlotType = SlotType.eAllArc Then Return False 'cannot have hangar doors in all-arcs
                            Case SlotConfig.eWeapons
                                If .SlotGroup > lMaxWpnSlot Then lMaxWpnSlot = .SlotGroup
                                If .SlotGroup < 1 Then Return False
                                If .SlotType = SlotType.eAllArc Then
                                    Dim bFound As Boolean = False
                                    For Z As Int32 = 0 To lWpnGrpsInAllArcs.GetUpperBound(0)
                                        If lWpnGrpsInAllArcs(Z) = .SlotGroup Then
                                            bFound = True
                                            Exit For
                                        End If
                                    Next Z
                                    If bFound = False Then
                                        ReDim Preserve lWpnGrpsInAllArcs(lWpnGrpsInAllArcs.GetUpperBound(0) + 1)
                                        lWpnGrpsInAllArcs(lWpnGrpsInAllArcs.GetUpperBound(0)) = .SlotGroup
                                    End If
                                End If
                        End Select
                    End If
                End With
            Next X
        Next Y

        For Z As Int32 = 0 To lWpnGrpsInAllArcs.GetUpperBound(0)
            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    With moSlots(X, Y)
                        If .SlotType <> SlotType.eUnused AndAlso .SlotConfig = SlotConfig.eWeapons Then
                            If .SlotGroup = lWpnGrpsInAllArcs(Z) AndAlso .SlotType <> SlotType.eAllArc Then
                                LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign Player has invalid wpn grps: " & Me.Owner.ObjectID)
                                Return False
                            End If
                        End If
                    End With
                Next X
            Next Y
        Next Z

        If lMaxWpnSlot > HullTech.MaxWpnSlots(Me.yTypeID, Me.ySubTypeID) Then
            LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign Player has too many wpn groups: " & Me.Owner.ObjectID)
            Return False '  Me.HullSize) Then Return False
        End If

        If yTypeID = 2 AndAlso ySubTypeID = 9 Then
            'MSC: We only care about the mesh portion of this modelid
            If ((ModelID And 255) = 138 AndAlso bHasEngine = False) OrElse (bHasEngine = False OrElse bHasHangarDoor = False) Then
                LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign Invalid Space Station Design: " & Me.Owner.ObjectID)
                Return False
            End If
        End If

        'If yTypeID = 2 AndAlso ySubTypeID = 1 Then
        '    Dim lHullSlots As Int32 = Me.GetHullAllocation(SlotConfig.eHangar)
        '    Dim lHullDoorSlots As Int32 = Me.GetMaxHangarDoorSize()
        '    Dim lCargoHull As Int32 = Me.GetHullAllocation(SlotConfig.eCargoBay)

        '    If lHullSlots < 500 OrElse lHullDoorSlots * fHullPerSlot < 500 OrElse lCargoHull < 1000 Then
        '        LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign Invalid design for mining facility: " & Me.Owner.ObjectID)
        '        Return False
        '    End If
        'End If

        'Set Insufficient Crew Quarters...
        bInsuffCrewQuarters = (lCrewQuarters * fHullPerSlot) < ((lCrewReqVal / 100.0F) * HullSize)
        If bInsuffCrewQuarters = True Then Return False
        If bShip = True AndAlso bEngineOnRearEdge = False Then
            LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign Unit without engines in rear: " & Me.Owner.ObjectID)
            Return False
        End If

        'Containers except for Crew Quarters and Armor and cargo bay must be touching the same container/group
        Dim vpVals() As HullValPair = Nothing
        Dim lValUB As Int32 = -1
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                With moSlots(X, Y)
                    If .SlotType <> SlotType.eUnused AndAlso .SlotConfig <> SlotConfig.eCargoBay AndAlso .SlotConfig <> SlotConfig.eArmorConfig AndAlso .SlotConfig <> SlotConfig.eCrewQuarters Then
                        Dim bFound As Boolean = False
                        For lIdx As Int32 = 0 To lValUB
                            If vpVals(lIdx).ValueConfig = .SlotConfig AndAlso vpVals(lIdx).ValueGroupNum = .SlotGroup Then
                                bFound = True
                                Exit For
                            End If
                        Next lIdx
                        If bFound = False Then
                            lValUB += 1
                            ReDim Preserve vpVals(lValUB)
                            vpVals(lValUB).ValueConfig = .SlotConfig
                            vpVals(lValUB).ValueGroupNum = .SlotGroup
                            vpVals(lValUB).Value = 0
                        End If
                        .bVerified = False
                    End If
                End With
            Next X
        Next Y

        'Now, go back through our values
        For lIdx As Int32 = 0 To lValUB
            Dim bStarted As Boolean = False

            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    Dim bDecrementY As Boolean = False
                    Dim bDecrementX As Boolean = False

                    With moSlots(X, Y)
                        If .SlotConfig = vpVals(lIdx).ValueConfig AndAlso .SlotGroup = vpVals(lIdx).ValueGroupNum Then
                            'Is this the first one?
                            If bStarted = False Then
                                bStarted = True
                                'Go ahead and verify this one...
                                .bVerified = True
                            End If

                            'Is this square verified?
                            If .bVerified = True Then
                                'yes, verify its neighbors
                                If Y <> 0 AndAlso moSlots(X, Y - 1).SlotConfig = .SlotConfig AndAlso moSlots(X, Y - 1).SlotGroup = .SlotGroup Then
                                    If moSlots(X, Y - 1).bVerified = False Then bDecrementY = True
                                    moSlots(X, Y - 1).bVerified = True
                                End If
                                If Y <> ml_GRID_SIZE_WH - 1 AndAlso moSlots(X, Y + 1).SlotConfig = .SlotConfig AndAlso moSlots(X, Y + 1).SlotGroup = .SlotGroup Then
                                    moSlots(X, Y + 1).bVerified = True
                                End If
                                If X <> 0 AndAlso moSlots(X - 1, Y).SlotConfig = .SlotConfig AndAlso moSlots(X - 1, Y).SlotGroup = .SlotGroup Then
                                    If moSlots(X - 1, Y).bVerified = False Then bDecrementX = True
                                    moSlots(X - 1, Y).bVerified = True
                                End If
                                If X <> ml_GRID_SIZE_WH - 1 AndAlso moSlots(X + 1, Y).SlotConfig = .SlotConfig AndAlso moSlots(X + 1, Y).SlotGroup = .SlotGroup Then
                                    moSlots(X + 1, Y).bVerified = True
                                End If
                            End If
                        End If
                    End With

                    If bDecrementY = True Then
                        Y = Y - 2
                        Exit For
                    ElseIf bDecrementX = True Then
                        X = X - 2
                    End If
                Next X
            Next Y
        Next lIdx

        'Now... go through and ensure all squares are valid
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                With moSlots(X, Y)
                    If .bVerified = False AndAlso .SlotType <> SlotType.eUnused AndAlso .SlotConfig <> SlotConfig.eCargoBay AndAlso .SlotConfig <> SlotConfig.eArmorConfig AndAlso .SlotConfig <> SlotConfig.eCrewQuarters Then
                        'Ok... got one... so find it in vpVals
                        For lIdx As Int32 = 0 To lValUB
                            If vpVals(lIdx).ValueConfig = .SlotConfig AndAlso vpVals(lIdx).ValueGroupNum = .SlotGroup Then
                                LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign: Invalid Slot Config detected! Player: " & Me.Owner.ObjectID)
                                Return False
                            End If
                        Next lIdx
                    End If
                End With
            Next X
        Next Y

        'Check Bay Door Edges...
        Dim lBayGroups() As Int32 = Nothing
        Dim lBayGroupUB As Int32 = -1
        Dim bHasNoEdge() As Boolean = Nothing
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                With moSlots(X, Y)
                    If .SlotType <> SlotType.eUnused AndAlso .SlotConfig = SlotConfig.eHangarDoor Then
                        Dim lIdx As Int32 = -1
                        For lTempIdx As Int32 = 0 To lBayGroupUB
                            If lBayGroups(lTempIdx) = .SlotGroup Then
                                lIdx = lTempIdx
                                Exit For
                            End If
                        Next lTempIdx
                        If lIdx = -1 Then
                            lBayGroupUB += 1
                            ReDim Preserve lBayGroups(lBayGroupUB)
                            ReDim Preserve bHasNoEdge(lBayGroupUB)
                            lBayGroups(lBayGroupUB) = .SlotGroup
                            bHasNoEdge(lBayGroupUB) = False
                            lIdx = lBayGroupUB
                        End If

                        Dim bHasEdge As Boolean = False
                        If X > 0 AndAlso moSlots(X - 1, Y).SlotType = SlotType.eUnused Then
                            bHasEdge = True
                        ElseIf X < ml_GRID_SIZE_WH - 1 AndAlso moSlots(X + 1, Y).SlotType = SlotType.eUnused Then
                            bHasEdge = True
                        ElseIf Y > 0 AndAlso moSlots(X, Y - 1).SlotType = SlotType.eUnused Then
                            bHasEdge = True
                        ElseIf Y < ml_GRID_SIZE_WH - 1 AndAlso moSlots(X, Y + 1).SlotType = SlotType.eUnused Then
                            bHasEdge = True
                        End If

                        bHasNoEdge(lIdx) = bHasNoEdge(lIdx) OrElse (bHasEdge = False)
                    End If
                End With
            Next X
        Next Y
        'Now, check the edges
        For X As Int32 = 0 To lBayGroupUB
            If bHasNoEdge(X) = True Then
                LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign bay slots not on edge: " & Me.Owner.ObjectID)
                Return False
            End If
        Next X
        'Now, check hangar door touches a hangar
        For lIdx As Int32 = 0 To lBayGroupUB
            Dim bFound As Boolean = False

            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    With moSlots(X, Y)
                        If .SlotType <> SlotType.eUnused AndAlso .SlotConfig = SlotConfig.eHangarDoor AndAlso .SlotGroup = lBayGroups(lIdx) Then
                            'check the neighbors
                            If X <> 0 AndAlso moSlots(X - 1, Y).SlotConfig = SlotConfig.eHangar Then
                                bFound = True
                            ElseIf X <> ml_GRID_SIZE_WH - 1 AndAlso moSlots(X + 1, Y).SlotConfig = SlotConfig.eHangar Then
                                bFound = True
                            ElseIf Y <> 0 AndAlso moSlots(X, Y - 1).SlotConfig = SlotConfig.eHangar Then
                                bFound = True
                            ElseIf Y <> ml_GRID_SIZE_WH - 1 AndAlso moSlots(X, Y + 1).SlotConfig = SlotConfig.eHangar Then
                                bFound = True
                            End If
                        End If
                    End With

                    If bFound = True Then Exit For
                Next X
                If bFound = True Then Exit For
            Next Y

            If bFound = False Then
                LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign hangar slots not touching bay doors: " & Me.Owner.ObjectID)
                Return False
            End If
        Next lIdx

        'Check Weapon Group Edges...
        Dim lWpnGroups() As Int32 = Nothing
        Dim lWpnGroupUB As Int32 = -1
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                With moSlots(X, Y)
                    If .SlotType <> SlotType.eUnused AndAlso .SlotConfig = SlotConfig.eWeapons Then
                        Dim bFound As Boolean = False
                        For lIdx As Int32 = 0 To lWpnGroupUB
                            If lWpnGroups(lIdx) = .SlotGroup Then
                                bFound = True
                                Exit For
                            End If
                        Next lIdx
                        If bFound = False Then
                            lWpnGroupUB += 1
                            ReDim Preserve lWpnGroups(lWpnGroupUB)
                            lWpnGroups(lWpnGroupUB) = .SlotGroup
                        End If
                    End If
                End With
            Next X
        Next Y

        Dim lWpnGroupEdges(lWpnGroupUB) As Int32
        For lIdx As Int32 = 0 To lWpnGroupUB
            'bBreakout = False
            Dim bLeftArcWpn As Boolean = False
            Dim bForeArcWpn As Boolean = False
            Dim bRightArcWpn As Boolean = False
            Dim bRearArcWpn As Boolean = False
            Dim bAllArcWpn As Boolean = False

            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    With moSlots(X, Y)
                        If .SlotType <> SlotType.eUnused AndAlso .SlotConfig = SlotConfig.eWeapons AndAlso lWpnGroups(lIdx) = .SlotGroup Then
                            'ok... now... is this an edge???
                            If (X > 0 AndAlso moSlots(X - 1, Y).SlotType = SlotType.eUnused) OrElse (X = 0) Then bLeftArcWpn = True
                            If (X < ml_GRID_SIZE_WH - 1 AndAlso moSlots(X + 1, Y).SlotType = SlotType.eUnused) OrElse (X = ml_GRID_SIZE_WH - 1) Then bRightArcWpn = True
                            If (Y > 0 AndAlso moSlots(X, Y - 1).SlotType = SlotType.eUnused) OrElse (Y = 0) Then bForeArcWpn = True
                            If (Y < ml_GRID_SIZE_WH - 1 AndAlso moSlots(X, Y + 1).SlotType = SlotType.eUnused) OrElse (Y = ml_GRID_SIZE_WH - 1) Then bRearArcWpn = True
                            If .SlotType = SlotType.eAllArc Then bAllArcWpn = True
                        End If
                    End With
                Next X
            Next Y

            'If (bLeftArcWpn Xor bForeArcWpn Xor bRightArcWpn Xor bRearArcWpn) = False AndAlso bAllArcWpn = False Then
            If bLeftArcWpn = False AndAlso bForeArcWpn = False AndAlso bRightArcWpn = False AndAlso bRearArcWpn = False AndAlso bAllArcWpn = False Then
                LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign: invalid wpn placement. " & Me.Owner.ObjectID)
                Return False
            End If

        Next lIdx

        If (Me.ModelID And 255) = 148 Then
            If (Me.Owner.oSpecials.lSuperSpecials And Player_Specials.SuperSpecialID.eOrbitalMiningPlatform) = 0 Then
                LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign: OMP without tech. " & Me.Owner.ObjectID)
                Return False
            End If
        End If

        Me.ErrorReasonCode = TechBuilderErrorReason.eNoError
        Return True

    End Function

    Public Overrides Function SetFromDesignMsg(ByRef yData() As Byte) As Boolean
        Dim lPos As Int32 = 14
        Dim lCnt As Int32
        Dim bRes As Boolean = False

        '2 bytes for hull tech type id
        '4 bytes for hull id
        '4 bytes for researcher id, 2 bytes for type id

        Try

            ReDim moSlots(ml_GRID_SIZE_WH - 1, ml_GRID_SIZE_WH - 1)
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    moSlots(X, Y).SlotConfig = SlotConfig.eArmorConfig
                    moSlots(X, Y).bVerified = False
                Next Y
            Next X

            ObjTypeID = ObjectType.eHullTech

            '20 bytes for name
            ReDim Me.HullName(19)
            Array.Copy(yData, lPos, Me.HullName, 0, 20)
            lPos += 20

            '4 bytes for hullsize
            HullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            '4 bytes for struct mineral ID
            StructuralMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            '4 bytes for hp
            StructuralHitPoints = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            '1 byte for TypeID
            yTypeID = yData(lPos) : lPos += 1

            '1 byte for SubTypeID
            ySubTypeID = yData(lPos) : lPos += 1

            '4 bytes for modelid
            Dim iTmpVal As Int16 = CShort(System.BitConverter.ToInt32(yData, lPos)) : lPos += 4
            If iTmpVal <> 0 Then miModelID = 0 Else miModelID = -1
            ModelID = iTmpVal

            '1 byte for chassis type
            yChassisType = yData(lPos) : lPos += 1

            '4 bytes for slot count
            lCnt = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            Dim yX As Byte
            Dim yY As Byte
            For X As Int32 = 0 To lCnt - 1
                yX = yData(lPos) : lPos += 1
                yY = yData(lPos) : lPos += 1
                moSlots(yX, yY).SlotConfig = CType(yData(lPos), SlotConfig) : lPos += 1
                moSlots(yX, yY).SlotGroup = yData(lPos) : lPos += 1
            Next X

            bRes = True

        Catch ex As Exception
            bRes = False
            GlobalVars.LogEvent(LogEventType.CriticalError, "HullTech.SetFromDesignMsg Error: " & ex.Message)
        End Try

        Return bRes
    End Function

    Public Sub New()
        ReDim moSlots(ml_GRID_SIZE_WH - 1, ml_GRID_SIZE_WH - 1)
    End Sub

    Public Sub SetHullSlot(ByVal ySlotX As Byte, ByVal ySlotY As Byte, ByVal ySlotConfig As SlotConfig, ByVal ySlotGroup As Byte)
        If ySlotX > -1 AndAlso ySlotX < ml_GRID_SIZE_WH AndAlso ySlotY > -1 AndAlso ySlotY < ml_GRID_SIZE_WH Then
            moSlots(ySlotX, ySlotY).SlotConfig = ySlotConfig
            moSlots(ySlotX, ySlotY).SlotGroup = ySlotGroup
        End If
    End Sub

    Public Property ModelID() As Int16
        Get
            Return miModelID
        End Get
        Set(ByVal value As Int16)
            If value <> miModelID Then
                miModelID = value

                Dim oTmpModDef As ModelDef = goModelDefs.GetModelDef(miModelID)
                Dim X As Int32
                Dim Y As Int32
                Dim lTemp As Int32

                For X = 0 To ml_GRID_SIZE_WH - 1
                    For Y = 0 To ml_GRID_SIZE_WH - 1
                        moSlots(X, Y).SlotType = SlotType.eUnused
                    Next
                Next

                If oTmpModDef Is Nothing = False Then
                    'Front
                    For lIdx As Int32 = 0 To oTmpModDef.FrontLocs.GetUpperBound(0)
                        lTemp = oTmpModDef.FrontLocs(lIdx)
                        Y = lTemp \ 30
                        X = lTemp - (Y * 30)
                        moSlots(X, Y).SlotType = SlotType.eFront
                    Next lIdx

                    'Left
                    For lIdx As Int32 = 0 To oTmpModDef.LeftLocs.GetUpperBound(0)
                        lTemp = oTmpModDef.LeftLocs(lIdx)
                        Y = lTemp \ 30
                        X = lTemp - (Y * 30)
                        moSlots(X, Y).SlotType = SlotType.eLeft
                    Next lIdx

                    'Rear
                    For lIdx As Int32 = 0 To oTmpModDef.RearLocs.GetUpperBound(0)
                        lTemp = oTmpModDef.RearLocs(lIdx)
                        Y = lTemp \ 30
                        X = lTemp - (Y * 30)
                        moSlots(X, Y).SlotType = SlotType.eRear
                    Next lIdx

                    'Right
                    For lIdx As Int32 = 0 To oTmpModDef.RightLocs.GetUpperBound(0)
                        lTemp = oTmpModDef.RightLocs(lIdx)
                        Y = lTemp \ 30
                        X = lTemp - (Y * 30)
                        moSlots(X, Y).SlotType = SlotType.eRight
                    Next lIdx

                    'All Arc
                    For lIdx As Int32 = 0 To oTmpModDef.AllArcLocs.GetUpperBound(0)
                        lTemp = oTmpModDef.AllArcLocs(lIdx)
                        Y = lTemp \ 30
                        X = lTemp - (Y * 30)
                        moSlots(X, Y).SlotType = SlotType.eAllArc
                    Next lIdx
                Else
                    'TODO: indicate, somehow, who the cheater is would be good...
                    LogEvent(LogEventType.PossibleCheat, "Could not find ModelID " & miModelID & ".")
                End If
            End If
        End Set
    End Property

    Protected Overrides Sub CalculateBothProdCosts()
        If moResearchCost Is Nothing OrElse moProductionCost Is Nothing Then ComponentDesigned()
    End Sub

    Public Overrides Sub ComponentDesigned()
        Dim bValid As Boolean = True

        If StructureMineral Is Nothing Then
            Me.ErrorReasonCode = TechBuilderErrorReason.eHull_HardnessZero
            Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
            LogEvent(LogEventType.PossibleCheat, "HullTech.ValidateDesign: StructureMineral is nothing. Owner is: " & Me.Owner.ObjectID)
            Return
        End If

        Dim lHardness As Int32 = StructureMineral.GetPropertyValue(eMinPropID.Hardness)

        If lHardness = 0 Then
            Me.ErrorReasonCode = TechBuilderErrorReason.eHull_HardnessZero
            bValid = False
        End If

        If moResearchCost Is Nothing Then moResearchCost = New ProductionCost

        'Dim uMinProp As MaterialPropertyItem2
        'With uMinProp
        '    .lMineralID = Me.StructuralMineralID
        '    .lPropertyID = eMinPropID.Hardness

        '    .lActualValue = StructureMineral.GetPropertyValue(.lPropertyID)
        '    .lKnownValue = Owner.GetMineralPropertyKnowledge(.lMineralID, .lPropertyID)
        '    .fNormalize = 1.0F
        '    .lGoalValue = 250
        '    If .lActualValue <> .lKnownValue Then
        '        MajorDesignFlaw = eComponentDesignFlaw.eMat1_Prop1 Or eComponentDesignFlaw.eShift_Not_Study
        '    End If
        'End With

        Try
            'Adjust this in order to create more variance in the hull
            'Dim fMinVariance As Single = Math.Min(uMinProp.FinalScore, 10) * 0.01F '* CSng(StructureMineral.GetPropertyValue(eMinPropID.Complexity) / PopIntel)
            Dim fModifiedSeed As Single = (0.7F * RandomSeed) + 0.8F
            'Dim lMineralCost As Int32 = CInt(fMinVariance * Me.HullSize * fModifiedSeed)
            Dim fPower As Single = CSng(Math.Max(1.0F, StructuralHitPoints / HullSize))

            Dim fMineralMultiplier As Single = 1.0F + ((255.0F - StructureMineral.GetPropertyValue(eMinPropID.Hardness)) / 200.0F)

            ''facilities that are not space stations...
            'lMineralCost \= 100
            ''lMineralCost = Math.Max(100, lMineralCost)
            'lMineralCost = CInt(Math.Pow(lMineralCost, fPower))

            'lMineralCost \= 10
            'lMineralCost = Math.Max(lMineralCost, 100)

            Dim lMineralCost As Int32
            If StructuralHitPoints > HullSize Then
                lMineralCost = CInt(Math.Abs((StructuralHitPoints - HullSize) * 1.1))
            Else : lMineralCost = 1
            End If
            lMineralCost = CInt(lMineralCost * HullSize * 0.01 * (1.0F - ((StructureMineral.GetPropertyValue(eMinPropID.Hardness) / 250) * 0.7)))
            lMineralCost = Math.Max(lMineralCost, 100)

            'randomseed * Size * sum(GBValue,NValue,AValue,SValue) * (1+NumberofWarnings) * SlotTypeCount) 
            Dim dblTemp As Double = StructuralHitPoints
            Dim lMult As Int32 = 0
            If (yChassisType And ChassisType.eGroundBased) <> 0 Then
                lMult += 1
            End If
            If (yChassisType And ChassisType.eAtmospheric) <> 0 Then
                lMult += 1
            End If
            If (yChassisType And ChassisType.eSpaceBased) <> 0 Then
                lMult += 1
            End If
            If (yChassisType And ChassisType.eNavalBased) <> 0 Then
                lMult += 1
            End If

            dblTemp *= Math.Max(1, lMult)
            'Now, our slottypecount
            dblTemp *= Math.Max(1, GetSlotTypeGroupCount())

            lMult = 0
            If Me.HasConfigOnSide(SlotConfig.eArmorConfig, SlotType.eFront) = False Then lMult += 1
            If Me.HasConfigOnSide(SlotConfig.eArmorConfig, SlotType.eLeft) = False Then lMult += 1
            If Me.HasConfigOnSide(SlotConfig.eArmorConfig, SlotType.eRear) = False Then lMult += 1
            If Me.HasConfigOnSide(SlotConfig.eArmorConfig, SlotType.eRight) = False Then lMult += 1
            lMult += GetCrossSideWarningCount()
            dblTemp *= Math.Max(1, lMult)
            Dim blResPoints As Int64 = CLng(Math.Max(5400000, dblTemp) * fModifiedSeed)
            blResPoints = CLng(Math.Pow(blResPoints, fPower))

            '= Size*1000*randomseed*max(materialComplexity,130)/intelligence 
            'Dim blResCredits As Int64 = CLng(CLng(StructuralHitPoints) * (100.0F * fModifiedSeed) * fModifiedSeed * Math.Max(130, StructureMineral.GetPropertyValue(eMinPropID.Complexity)) / PopIntel)
            'blResCredits = CLng(Math.Pow(blResCredits, fPower))
            'Dim blProdCredits As Int64 = CLng((HullSize * StructureMineral.GetPropertyValue(eMinPropID.Complexity) / 100.0F) * fMineralMultiplier)
            'blProdCredits = CLng(Math.Pow(blProdCredits, fPower))
            'If yTypeID = 2 Then
            ' blProdCredits += CLng(Me.StructuralHitPoints) * 100
            'Else
            'blProdCredits += CLng(Me.StructuralHitPoints) * 1000
            'End If

            Dim blResCredits As Int64 = CLng((StructuralHitPoints * (100.0F * fModifiedSeed) * fModifiedSeed * 130) / PopIntel)
            blResCredits = CLng(Math.Pow(blResCredits, fPower))
            Dim blProdCredits As Int64 = CLng(HullSize * fMineralMultiplier)
            blProdCredits = CLng(Math.Pow(blProdCredits, fPower))
            If yTypeID = 2 Then
                If ySubTypeID = 9 Then
                    blProdCredits += CLng(Me.StructuralHitPoints) * 100L
                Else
                    blProdCredits += CLng(Me.StructuralHitPoints) * 10L
                End If
            Else
                blProdCredits += CLng(Me.StructuralHitPoints) * 1000L
            End If

            With moResearchCost
                .ObjectID = Me.ObjectID
                .ObjTypeID = Me.ObjTypeID
                .ColonistCost = 0
                .EnlistedCost = 0
                .OfficerCost = 0

                Try
                    If bValid = True Then
                        .CreditCost = blResCredits
                        .PointsRequired = blResPoints
                        .ItemCostUB = -1
                        .AddProductionCostItem(Me.StructuralMineralID, ObjectType.eMineral, lMineralCost)
                        'Dim HPExtra As Int32 = StructuralHitPoints - HullSize
                        'If HPExtra < 1 Then HPExtra = 1
                        ''.CreditCost = CLng((HullSize / lHardness) * (Math.Sqrt(255 - lHardness) * CLng(CLng(HPExtra) * CLng(HPExtra)))) \ 100
                        '.CreditCost = CLng((HullSize / lHardness) * (Math.Sqrt(255 - lHardness) * CLng(CLng(HPExtra) * CLng(HPExtra) * Math.Max(1, CLng(StructuralHitPoints \ 10))))) \ 100
                        'If .CreditCost < 100 Then .CreditCost = 100
                        '.PointsRequired = (.CreditCost * 10)

                        '.ItemCostUB = -1
                        '.AddProductionCostItem(Me.StructuralMineralID, ObjectType.eMineral, Math.Max((HullSize \ 1000) + ((StructuralHitPoints \ lHardness) + 1), 10))
                    End If
                Catch
                    bValid = False
                    Me.ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
                End Try
            End With

            If moProductionCost Is Nothing Then moProductionCost = New ProductionCost
            With moProductionCost
                .ObjectID = Me.ObjectID
                .ObjTypeID = Me.ObjTypeID
                .ColonistCost = 0
                .EnlistedCost = 0
                .OfficerCost = 0

                Try
                    .PointsRequired = 0
                    .CreditCost = blProdCredits
                    .ItemCostUB = -1
                    .AddProductionCostItem(Me.StructuralMineralID, ObjectType.eMineral, lMineralCost)

                    '.PointsRequired = moResearchCost.PointsRequired
                    '.CreditCost = moResearchCost.CreditCost \ 5
                    '.ItemCostUB = -1
                    '.AddProductionCostItem(Me.StructuralMineralID, ObjectType.eMineral, Math.Max((HullSize \ 1000) + ((StructuralHitPoints \ lHardness) + 1), 10))
                Catch
                    bValid = False
                    Me.ErrorReasonCode = TechBuilderErrorReason.eUnknownReason
                End Try
            End With
        Catch
            bValid = False
        End Try

        If bValid = False Then Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
    End Sub

    Private Function GetCrossSideWarningCount() As Int32
        Dim lResult As Int32 = 0
        Dim bFore As Boolean = False
        Dim bLeft As Boolean = False
        Dim bRear As Boolean = False
        Dim bRight As Boolean = False
        Dim bAll As Boolean = False

        WhatSidesHaveConfig(SlotConfig.eCargoBay, bFore, bLeft, bRear, bRight, bAll)
        Dim lTempCnt As Int32 = 0
        If bFore = True Then lTempCnt += 1 : If bLeft = True Then lTempCnt += 1 : If bRear = True Then lTempCnt += 1 : If bRight = True Then lTempCnt += 1 : If bAll = True Then lTempCnt += 1
        If lTempCnt > 1 Then lResult += 1

        WhatSidesHaveConfig(SlotConfig.eEngines, bFore, bLeft, bRear, bRight, bAll)
        lTempCnt = 0
        If bFore = True Then lTempCnt += 1 : If bLeft = True Then lTempCnt += 1 : If bRear = True Then lTempCnt += 1 : If bRight = True Then lTempCnt += 1 : If bAll = True Then lTempCnt += 1
        If lTempCnt > 1 Then lResult += 1

        WhatSidesHaveConfig(SlotConfig.eHangar, bFore, bLeft, bRear, bRight, bAll)
        lTempCnt = 0
        If bFore = True Then lTempCnt += 1 : If bLeft = True Then lTempCnt += 1 : If bRear = True Then lTempCnt += 1 : If bRight = True Then lTempCnt += 1 : If bAll = True Then lTempCnt += 1
        If lTempCnt > 1 Then lResult += 1

        WhatSidesHaveConfig(SlotConfig.eRadar, bFore, bLeft, bRear, bRight, bAll)
        lTempCnt = 0
        If bFore = True Then lTempCnt += 1 : If bLeft = True Then lTempCnt += 1 : If bRear = True Then lTempCnt += 1 : If bRight = True Then lTempCnt += 1 : If bAll = True Then lTempCnt += 1
        If lTempCnt > 1 Then lResult += 1

        WhatSidesHaveConfig(SlotConfig.eShields, bFore, bLeft, bRear, bRight, bAll)
        lTempCnt = 0
        If bFore = True Then lTempCnt += 1 : If bLeft = True Then lTempCnt += 1 : If bRear = True Then lTempCnt += 1 : If bRight = True Then lTempCnt += 1 : If bAll = True Then lTempCnt += 1
        If lTempCnt > 1 Then lResult += 1

        Return lResult
    End Function

    Public Function GetSlotTypeGroupCount() As Int32
        Dim lConfig() As Int32 = Nothing
        Dim lGroup() As Int32 = Nothing
        Dim lUB As Int32 = -1

        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).SlotType <> SlotType.eUnused Then

                    Dim lIdx As Int32 = -1
                    For Z As Int32 = 0 To lUB
                        If lConfig(Z) = moSlots(X, Y).SlotConfig Then
                            If moSlots(X, Y).SlotConfig = SlotConfig.eHangarDoor OrElse moSlots(X, Y).SlotConfig = SlotConfig.eWeapons Then
                                If lGroup(Z) = moSlots(X, Y).SlotGroup Then
                                    lIdx = Z
                                    Exit For
                                End If
                            Else : lIdx = Z
                            End If
                        End If
                    Next Z
                    If lIdx = -1 Then
                        lUB += 1
                        ReDim Preserve lConfig(lUB)
                        ReDim Preserve lGroup(lUB)
                        lConfig(lUB) = moSlots(X, Y).SlotConfig
                        lGroup(lUB) = moSlots(X, Y).SlotGroup
                    End If
                End If
            Next X
        Next Y

        Return lUB + 1

    End Function

    Protected Overrides Sub FinalizeResearch()

    End Sub

    Private moMineral As Mineral = Nothing
    Public ReadOnly Property StructureMineral() As Mineral
        Get
            If moMineral Is Nothing OrElse moMineral.ObjectID <> Me.StructuralMineralID Then
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) = Me.StructuralMineralID Then
                        moMineral = goMineral(X)
                        Exit For
                    End If
                Next X
            End If
            Return moMineral
        End Get
    End Property

    Public Function GetHullAllocation(ByVal lConfig As SlotConfig) As Int32
        Dim lTotal As Int32 = 0
        Dim lCnt As Int32 = 0

        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).SlotType <> SlotType.eUnused Then
                    lTotal += 1
                    If moSlots(X, Y).SlotConfig = lConfig Then
                        lCnt += 1
                    End If
                End If
            Next X
        Next Y

        If lTotal = 0 Then Return 0
        Dim fHullPerSlot As Single = CSng(HullSize / lTotal)

        Return CInt(lCnt * fHullPerSlot)
    End Function

    Private Function GetMaxHangarDoorSize() As Int32
        Dim lGroup() As Int32 = Nothing
        Dim lSize() As Int32 = Nothing
        Dim lGroupUB As Int32 = -1

        Dim lTotal As Int32

        'Load our groups
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).SlotType <> SlotType.eUnused Then
                    lTotal += 1

                    'Ok, now, is this a HangarDoor?
                    If moSlots(X, Y).SlotConfig = SlotConfig.eHangarDoor Then
                        'Ok, it is... find our group
                        Dim lIdx As Int32 = -1
                        For lTempIdx As Int32 = 0 To lGroupUB
                            If lGroup(lTempIdx) = moSlots(X, Y).SlotGroup Then
                                lIdx = lTempIdx
                                Exit For
                            End If
                        Next lTempIdx

                        If lIdx = -1 Then
                            lGroupUB += 1
                            ReDim Preserve lGroup(lGroupUB)
                            ReDim Preserve lSize(lGroupUB)
                            lIdx = lGroupUB
                        End If

                        lGroup(lIdx) = moSlots(X, Y).SlotGroup
                        lSize(lIdx) += 1
                    End If
                End If
            Next X
        Next Y

        Dim lMaxSize As Int32 = 0
        For X As Int32 = 0 To lGroupUB
            If lSize(X) > lMaxSize Then lMaxSize = lSize(X)
        Next X
        Return lMaxSize
    End Function

    Public Function HasConfigOnSide(ByVal lConfig As SlotConfig, ByVal lType As SlotType) As Boolean
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).SlotType = lType AndAlso moSlots(X, Y).SlotConfig = lConfig Then
                    Return True
                End If
            Next X
        Next Y
        Return False
    End Function

    Public Sub FillBaseCrits(ByRef lFront As Int32, ByRef lLeft As Int32, ByRef lRear As Int32, ByRef lRight As Int32, ByVal bHasEngine As Boolean, ByVal bHasRadar As Boolean, ByVal bHasShield As Boolean, ByVal lForeWpnCnt As Int32, ByVal lLeftWpnCnt As Int32, ByVal lRearWpnCnt As Int32, ByVal lRightWpnCnt As Int32)
        Dim bFront As Boolean = False
        Dim bLeft As Boolean = False
        Dim bRight As Boolean = False
        Dim bRear As Boolean = False
        Dim bAll As Boolean = False

        Dim bHasHangar As Boolean = (Me.GetHullAllocation(SlotConfig.eHangar) <> 0) AndAlso (Me.GetHullAllocation(SlotConfig.eHangarDoor) <> 0)

        'Clear slots...
        lFront = 0 : lLeft = 0 : lRear = 0 : lRight = 0

        'Now, fill them
        If bHasEngine = True Then
            WhatSidesHaveConfig(SlotConfig.eEngines, bFront, bLeft, bRear, bRight, bAll)

            If bAll = True Then
                lFront = lFront Or elUnitStatus.eEngineOperational
                lLeft = lLeft Or elUnitStatus.eEngineOperational
                lRight = lRight Or elUnitStatus.eEngineOperational
                lRear = lRear Or elUnitStatus.eEngineOperational
            Else
                If bFront = True Then lFront = lFront Or elUnitStatus.eEngineOperational
                If bLeft = True Then lLeft = lLeft Or elUnitStatus.eEngineOperational
                If bRear = True Then lRear = lRear Or elUnitStatus.eEngineOperational
                If bRight = True Then lRight = lRight Or elUnitStatus.eEngineOperational
            End If
        End If

        If Me.GetHullAllocation(SlotConfig.eFuelBay) <> 0 Then
            WhatSidesHaveConfig(SlotConfig.eFuelBay, bFront, bLeft, bRear, bRight, bAll)

            If bAll = True Then
                lFront = lFront Or elUnitStatus.eFuelBayOperational
                lLeft = lLeft Or elUnitStatus.eFuelBayOperational
                lRight = lRight Or elUnitStatus.eFuelBayOperational
                lRear = lRear Or elUnitStatus.eFuelBayOperational
            Else
                If bFront = True Then lFront = lFront Or elUnitStatus.eFuelBayOperational
                If bLeft = True Then lLeft = lLeft Or elUnitStatus.eFuelBayOperational
                If bRear = True Then lRear = lRear Or elUnitStatus.eFuelBayOperational
                If bRight = True Then lRight = lRight Or elUnitStatus.eFuelBayOperational
            End If
        End If

        If bHasShield = True Then
            WhatSidesHaveConfig(SlotConfig.eShields, bFront, bLeft, bRear, bRight, bAll)

            If bAll = True Then
                lFront = lFront Or elUnitStatus.eShieldOperational
                lLeft = lLeft Or elUnitStatus.eShieldOperational
                lRight = lRight Or elUnitStatus.eShieldOperational
                lRear = lRear Or elUnitStatus.eShieldOperational
            Else
                If bFront = True Then lFront = lFront Or elUnitStatus.eShieldOperational
                If bLeft = True Then lLeft = lLeft Or elUnitStatus.eShieldOperational
                If bRear = True Then lRear = lRear Or elUnitStatus.eShieldOperational
                If bRight = True Then lRight = lRight Or elUnitStatus.eShieldOperational
            End If
        End If

        If bHasRadar = True Then
            WhatSidesHaveConfig(SlotConfig.eRadar, bFront, bLeft, bRear, bRight, bAll)

            If bAll = True Then
                lFront = lFront Or elUnitStatus.eRadarOperational
                lLeft = lLeft Or elUnitStatus.eRadarOperational
                lRight = lRight Or elUnitStatus.eRadarOperational
                lRear = lRear Or elUnitStatus.eRadarOperational
            Else
                If bFront = True Then lFront = lFront Or elUnitStatus.eRadarOperational
                If bLeft = True Then lLeft = lLeft Or elUnitStatus.eRadarOperational
                If bRear = True Then lRear = lRear Or elUnitStatus.eRadarOperational
                If bRight = True Then lRight = lRight Or elUnitStatus.eRadarOperational
            End If
        End If

        If bHasHangar = True Then
            WhatSidesHaveConfig(SlotConfig.eHangar, bFront, bLeft, bRear, bRight, bAll)

            If bAll = True Then
                lFront = lFront Or elUnitStatus.eHangarOperational
                lLeft = lLeft Or elUnitStatus.eHangarOperational
                lRight = lRight Or elUnitStatus.eHangarOperational
                lRear = lRear Or elUnitStatus.eHangarOperational
            Else
                If bFront = True Then lFront = lFront Or elUnitStatus.eHangarOperational
                If bLeft = True Then lLeft = lLeft Or elUnitStatus.eHangarOperational
                If bRear = True Then lRear = lRear Or elUnitStatus.eHangarOperational
                If bRight = True Then lRight = lRight Or elUnitStatus.eHangarOperational
            End If
        End If

        If Me.GetHullAllocation(SlotConfig.eCargoBay) <> 0 Then
            WhatSidesHaveConfig(SlotConfig.eCargoBay, bFront, bLeft, bRear, bRight, bAll)
            If bAll = True Then
                lFront = lFront Or elUnitStatus.eCargoBayOperational
                lLeft = lLeft Or elUnitStatus.eCargoBayOperational
                lRight = lRight Or elUnitStatus.eCargoBayOperational
                lRear = lRear Or elUnitStatus.eCargoBayOperational
            Else
                If bFront = True Then lFront = lFront Or elUnitStatus.eCargoBayOperational
                If bLeft = True Then lLeft = lLeft Or elUnitStatus.eCargoBayOperational
                If bRight = True Then lRight = lRight Or elUnitStatus.eCargoBayOperational
                If bRear = True Then lRear = lRear Or elUnitStatus.eCargoBayOperational
            End If
        End If

        If lForeWpnCnt > 0 Then
            lFront = lFront Or elUnitStatus.eForwardWeapon1
            If lForeWpnCnt > 1 Then lFront = lFront Or elUnitStatus.eForwardWeapon2
        End If
        If lLeftWpnCnt > 0 Then
            lLeft = lLeft Or elUnitStatus.eLeftWeapon1
            If lLeftWpnCnt > 1 Then lLeft = lLeft Or elUnitStatus.eLeftWeapon2
        End If
        If lRearWpnCnt > 0 Then
            lRear = lRear Or elUnitStatus.eAftWeapon1
            If lRearWpnCnt > 1 Then lRear = lRear Or elUnitStatus.eAftWeapon2
        End If
        If lRightWpnCnt > 0 Then
            lRight = lRight Or elUnitStatus.eRightWeapon1
            If lRightWpnCnt > 1 Then lRight = lRight Or elUnitStatus.eRightWeapon2
        End If
    End Sub

    Private Sub WhatSidesHaveConfig(ByVal lConfig As SlotConfig, ByRef bFront As Boolean, ByRef bLeft As Boolean, ByRef bRear As Boolean, ByRef bRight As Boolean, ByRef bAll As Boolean)
        bLeft = False : bFront = False : bRear = False : bRight = False : bAll = False

        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).SlotType <> SlotType.eUnused AndAlso moSlots(X, Y).SlotConfig = lConfig Then
                    Select Case moSlots(X, Y).SlotType
                        Case SlotType.eFront : bFront = True
                        Case SlotType.eLeft : bLeft = True
                        Case SlotType.eRear : bRear = True
                        Case SlotType.eRight : bRight = True
                        Case SlotType.eAllArc
                            bAll = True
                            Return
                    End Select
                End If
            Next X
        Next Y
    End Sub

    Public Function GetSideHullAllocation(ByVal lType As SlotType, ByVal lConfig As SlotConfig) As Int32
        Dim lTotal As Int32 = 0
        Dim lCnt As Int32 = 0

        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).SlotType <> SlotType.eUnused Then
                    lTotal += 1
                    If moSlots(X, Y).SlotConfig = lConfig AndAlso moSlots(X, Y).SlotType = lType Then
                        lCnt += 1
                    End If
                End If
            Next X
        Next Y

        If lTotal = 0 Then Return 0
        Dim fHullPerSlot As Single = CSng(HullSize / lTotal)

        Return CInt(lCnt * fHullPerSlot)
    End Function

    Public Function GetWeaponHullAllotment(ByVal BaseSlotX As Byte, ByVal BaseSlotY As Byte) As Int32
        Dim lTotal As Int32 = 0
        Dim lCnt As Int32 = 0

        Dim lGroupNum As Int32 = 0

        If BaseSlotX < 0 OrElse BaseSlotX > moSlots.GetUpperBound(0) Then Return 0
        If BaseSlotY < 0 OrElse BaseSlotY > moSlots.GetUpperBound(1) Then Return 0

        lGroupNum = moSlots(BaseSlotX, BaseSlotY).SlotGroup

        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).SlotType <> SlotType.eUnused Then
                    lTotal += 1
                    If moSlots(X, Y).SlotConfig = SlotConfig.eWeapons AndAlso moSlots(X, Y).SlotGroup = lGroupNum Then
                        lCnt += 1
                    End If
                End If
            Next X
        Next Y
        If lTotal = 0 Then Return 0
        Dim fHullPerSlot As Single = CSng(HullSize / lTotal)

        Return CInt(lCnt * fHullPerSlot)
    End Function

    Public Function TotalSlots() As Int32
        Dim lCnt As Int32 = 0
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).SlotType <> SlotType.eUnused Then
                    lCnt += 1
                End If
            Next X
        Next Y
        Return lCnt
    End Function

    Public Sub SetEntityDefsHangarDoors(ByRef oEntityDef As Epica_Entity_Def)
        Dim lGroup() As Int32 = Nothing
        Dim lSize() As Int32 = Nothing
        Dim lGroupUB As Int32 = -1

        Dim lTotal As Int32

        oEntityDef.lHangarDoorUB = -1

        'Load our groups
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).SlotType <> SlotType.eUnused Then
                    lTotal += 1

                    'Ok, now, is this a HangarDoor?
                    If moSlots(X, Y).SlotConfig = SlotConfig.eHangarDoor Then
                        'Ok, it is... find our group
                        Dim lIdx As Int32 = -1
                        For lTempIdx As Int32 = 0 To lGroupUB
                            If lGroup(lTempIdx) = moSlots(X, Y).SlotGroup Then
                                lIdx = lTempIdx
                                Exit For
                            End If
                        Next lTempIdx

                        If lIdx = -1 Then
                            lGroupUB += 1
                            ReDim Preserve lGroup(lGroupUB)
                            ReDim Preserve lSize(lGroupUB)
                            lIdx = lGroupUB
                        End If

                        lGroup(lIdx) = moSlots(X, Y).SlotGroup
                        lSize(lIdx) += 1
                    End If
                End If
            Next X
        Next Y

        'Get our HullPerSlot
        Dim fHullPerSlot As Single = CSng(HullSize / lTotal)

        Dim lSideCnt(3) As Int32
        For lIdx As Int32 = 0 To lGroupUB
            'Reset our cnts
            lSideCnt(0) = 0 : lSideCnt(1) = 0 : lSideCnt(2) = 0 : lSideCnt(3) = 0

            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    If moSlots(X, Y).SlotType <> SlotType.eUnused AndAlso moSlots(X, Y).SlotConfig = SlotConfig.eHangarDoor AndAlso moSlots(X, Y).SlotGroup = lGroup(lIdx) Then
                        Select Case moSlots(X, Y).SlotType
                            Case SlotType.eFront
                                lSideCnt(0) += 1
                            Case SlotType.eLeft
                                lSideCnt(1) += 1
                            Case SlotType.eRear
                                lSideCnt(2) += 1
                            Case SlotType.eRight
                                lSideCnt(3) += 1
                        End Select
                    End If
                Next X
            Next Y

            'Now, get our max
            Dim lMax As Int32 = Int32.MinValue
            Dim lSide As Int32 = 0
            For Y As Int32 = 0 To 3
                If lSideCnt(Y) > lMax Then
                    lMax = lSideCnt(Y)
                    lSide = Y
                End If
            Next Y

            'Now, add our hangar door
            oEntityDef.AddHangarDoor(-1, CInt(fHullPerSlot * lSize(lIdx)), CByte(lSide))
        Next lIdx


    End Sub

    Public Overrides Function GetNonOwnerMsg() As Byte()
        Dim yResult(27) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eGetNonOwnerDetails).CopyTo(yResult, lPos) : lPos += 2
        Me.GetGUIDAsString.CopyTo(yResult, lPos) : lPos += 6
        Me.HullName.CopyTo(yResult, lPos) : lPos += 20

        Return yResult
    End Function

    Public Function GetSlotGrpNum(ByVal X As Int32, ByVal Y As Int32) As Int32
        Return moSlots(X, Y).SlotGroup
    End Function
    Public Function GetWeaponsMainSlotType(ByVal lWpnGrpNum As Int32) As SlotType

        Dim bOnEdge() As Boolean = {False, False, False, False}
        Dim lSideCnt() As Int32 = {0, 0, 0, 0}

        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                With moSlots(X, Y)
                    If .SlotType <> SlotType.eNoChange AndAlso .SlotType <> SlotType.eUnused Then
                        If .SlotConfig = SlotConfig.eWeapons AndAlso .SlotGroup = lWpnGrpNum Then
                            'ok, determine what we are
                            If .SlotType = SlotType.eAllArc Then Return SlotType.eAllArc

                            Select Case .SlotType
                                Case SlotType.eLeft
                                    lSideCnt(UnitArcs.eLeftArc) += 1
                                    If bOnEdge(UnitArcs.eLeftArc) = False Then bOnEdge(UnitArcs.eLeftArc) = SlotIsOnEdge(X, Y)
                                Case SlotType.eRear
                                    lSideCnt(UnitArcs.eBackArc) += 1
                                    If bOnEdge(UnitArcs.eBackArc) = False Then bOnEdge(UnitArcs.eBackArc) = SlotIsOnEdge(X, Y)
                                Case SlotType.eRight
                                    lSideCnt(UnitArcs.eRightArc) += 1
                                    If bOnEdge(UnitArcs.eRightArc) = False Then bOnEdge(UnitArcs.eRightArc) = SlotIsOnEdge(X, Y)
                                Case SlotType.eFront
                                    lSideCnt(UnitArcs.eForwardArc) += 1
                                    If bOnEdge(UnitArcs.eForwardArc) = False Then bOnEdge(UnitArcs.eForwardArc) = SlotIsOnEdge(X, Y)
                            End Select
                        End If
                    End If
                End With
            Next X
        Next Y

        Dim lMax As Int32 = 0
        Dim lSideVal As Int32 = -1
        For X As Int32 = 0 To 3
            If bOnEdge(X) = True Then
                If lSideCnt(X) > lMax Then
                    lMax = lSideCnt(X)
                    lSideVal = X
                End If
            End If
        Next X

        'did we end up with a side?
        If lSideVal <> -1 Then
            Select Case lSideVal
                Case UnitArcs.eLeftArc
                    Return SlotType.eLeft
                Case UnitArcs.eForwardArc
                    Return SlotType.eFront
                Case UnitArcs.eRightArc
                    Return SlotType.eRight
                Case UnitArcs.eBackArc
                    Return SlotType.eRear
                Case Else
                    Return SlotType.eUnused
            End Select
        Else
            Return SlotType.eUnused
        End If

    End Function

    Private Function SlotIsOnEdge(ByVal X As Int32, ByVal Y As Int32) As Boolean
        'Ok, check our edges
        If X <> 0 Then
            If moSlots(X - 1, Y).SlotType = SlotType.eUnused Then
                Return True
            End If
        End If
        If X <> ml_GRID_SIZE_WH - 1 Then
            If moSlots(X + 1, Y).SlotType = SlotType.eUnused Then
                Return True
            End If
        End If
        If Y <> 0 Then
            If moSlots(X, Y - 1).SlotType = SlotType.eUnused Then
                Return True
            End If
        End If
        If Y <> ml_GRID_SIZE_WH - 1 Then
            If moSlots(X, Y + 1).SlotType = SlotType.eUnused Then
                Return True
            End If
        End If
        Return False
    End Function

    Public Function GetSlotType(ByVal lSlotX As Int32, ByVal lSlotY As Int32) As SlotType
        If lSlotX > -1 AndAlso lSlotY > -1 Then
            If lSlotX <= moSlots.GetUpperBound(0) AndAlso lSlotY <= moSlots.GetUpperBound(1) Then
                Return moSlots(lSlotX, lSlotY).SlotType
            End If
        End If
        Return SlotType.eUnused
    End Function

    Public Overrides Function TechnologyScore() As Integer
        Return 1000I
    End Function

    Public Overrides Function GetPlayerTechKnowledgeMsg(ByVal yTechLvl As PlayerTechKnowledge.KnowledgeType) As Byte()
        Dim yMsg() As Byte = Nothing
        Dim lPos As Int32 = 0

        'set up our msg size
        lPos = 30
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eNameOnly Then lPos += 17
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
            Dim lCnt As Int32 = 0

            'Get our count of items to send...
            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    If moSlots(X, Y).SlotType <> SlotType.eUnused AndAlso moSlots(X, Y).SlotConfig <> SlotConfig.eArmorConfig Then lCnt += 1
                Next X
            Next Y

            lPos += (lCnt * 4) + 16
        End If
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then lPos += 4
        ReDim yMsg(lPos - 1)
        lPos = 0

        'Default Attributes
        GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        Me.HullName.CopyTo(yMsg, lPos) : lPos += 20

        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eNameOnly Then
            'at least settings 1 
            System.BitConverter.GetBytes(HullSize).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = yTypeID : lPos += 1
            yMsg(lPos) = ySubTypeID : lPos += 1
            System.BitConverter.GetBytes(ModelID).CopyTo(yMsg, lPos) : lPos += 2
            yMsg(lPos) = yChassisType : lPos += 1
            System.BitConverter.GetBytes(PowerRequired).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(StructuralHitPoints).CopyTo(yMsg, lPos) : lPos += 4
        End If
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
            'at least settings 2  
            Dim lCnt As Int32 = 0

            'Get our count of items to send...
            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    If moSlots(X, Y).SlotType <> SlotType.eUnused AndAlso moSlots(X, Y).SlotConfig <> SlotConfig.eArmorConfig Then lCnt += 1
                Next X
            Next Y

            System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4
            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    If moSlots(X, Y).SlotType <> SlotType.eUnused AndAlso moSlots(X, Y).SlotConfig <> SlotConfig.eArmorConfig Then
                        yMsg(lPos) = CByte(X) : lPos += 1
                        yMsg(lPos) = CByte(Y) : lPos += 1
                        yMsg(lPos) = CByte(moSlots(X, Y).SlotConfig) : lPos += 1
                        yMsg(lPos) = CByte(moSlots(X, Y).SlotGroup) : lPos += 1
                    End If
                Next X
            Next Y

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
            System.BitConverter.GetBytes(StructuralMineralID).CopyTo(yMsg, lPos) : lPos += 4
        End If

        Return yMsg
    End Function

    Public Sub GetSlotCriticalChances(ByRef oEntityDef As Epica_Entity_Def)
        Dim uHSTCC() As HullSlotTypeConfigChance = Nothing
        Dim lUB As Int32 = -1

        Dim lSlotScore(ml_GRID_SIZE_WH - 1, ml_GRID_SIZE_WH - 1) As Int32
        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).SlotType <> SlotType.eUnused Then
                    Dim lVal As Int32 = 100     'base 100
                    'now, each side I directly touch add 100
                    If X > 0 Then If moSlots(X - 1, Y).SlotType = SlotType.eUnused Then lVal += 100
                    If Y > 0 Then If moSlots(X, Y - 1).SlotType = SlotType.eUnused Then lVal += 100
                    If X < ml_GRID_SIZE_WH - 1 Then If moSlots(X + 1, Y).SlotType = SlotType.eUnused Then lVal += 100
                    If Y < ml_GRID_SIZE_WH - 1 Then If moSlots(X, Y + 1).SlotType = SlotType.eUnused Then lVal += 100
                    lSlotScore(X, Y) = lVal
                Else : lSlotScore(X, Y) = 0
                End If
            Next X
        Next Y

        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                If moSlots(X, Y).SlotType <> SlotType.eUnused AndAlso moSlots(X, Y).SlotConfig <> SlotConfig.eArmorConfig AndAlso moSlots(X, Y).SlotConfig <> SlotConfig.eCrewQuarters AndAlso moSlots(X, Y).SlotConfig <> SlotConfig.eFuelBay Then
                    Dim lType As SlotType = moSlots(X, Y).SlotType
                    Dim lConfig As SlotConfig = moSlots(X, Y).SlotConfig

                    'Find our HSTCC
                    Dim lIdx As Int32 = -1
                    For lTmp As Int32 = 0 To lUB
                        If uHSTCC(lTmp).lSlotType = lType AndAlso uHSTCC(lTmp).lSlotConfig = lConfig Then
                            lIdx = lTmp
                            Exit For
                        End If
                    Next lTmp
                    If lIdx = -1 Then
                        lUB += 1
                        ReDim Preserve uHSTCC(lUB)
                        lIdx = lUB
                        uHSTCC(lIdx).lSlotType = lType
                        uHSTCC(lIdx).lSlotConfig = lConfig
                    End If

                    'Now, take that score and interpolate the surrounding scores
                    Dim lDirectVal As Int32 = 0
                    Dim lDiagVal As Int32 = 0
                    If X > 0 Then
                        lDirectVal += lSlotScore(X - 1, Y)
                        If Y > 0 Then lDiagVal += lSlotScore(X - 1, Y - 1)
                        If Y < ml_GRID_SIZE_WH - 1 Then lDiagVal += lSlotScore(X - 1, Y + 1)
                    End If
                    If X < ml_GRID_SIZE_WH - 1 Then
                        lDirectVal += lSlotScore(X + 1, Y)
                        If Y > 0 Then lDiagVal += lSlotScore(X + 1, Y - 1)
                        If Y < ml_GRID_SIZE_WH - 1 Then lDiagVal += lSlotScore(X + 1, Y + 1)
                    End If
                    If Y > 0 Then lDirectVal += lSlotScore(X, Y - 1)
                    If Y < ml_GRID_SIZE_WH - 1 Then lDirectVal += lSlotScore(X, Y + 1)

                    Dim lInterVal As Int32 = (lSlotScore(X, Y) \ 2) + (lDirectVal \ 4) + (lDiagVal \ 8)
                    uHSTCC(lIdx).lOverallScore += lInterVal
                End If
            Next X
        Next Y

        'Now, go through our HSTCC's and get the final chance to critical
        For X As Int32 = 0 To lUB
            Dim lTotalScore As Int32 = 0
            For Y As Int32 = 0 To lUB
                If uHSTCC(X).lSlotType = uHSTCC(Y).lSlotType Then lTotalScore += uHSTCC(Y).lOverallScore
            Next Y
            uHSTCC(X).lChanceToCritical = CInt(CSng(uHSTCC(X).lOverallScore / lTotalScore) * 100)
            If uHSTCC(X).lChanceToCritical < 1 Then uHSTCC(X).lChanceToCritical = 1
            If uHSTCC(X).lChanceToCritical > 100 Then uHSTCC(X).lChanceToCritical = 100
        Next X

        'Now, populate the entity def's values
        For X As Int32 = 0 To lUB
            With uHSTCC(X)
                Dim ySide As UnitArcs
                Select Case .lSlotType
                    Case SlotType.eAllArc
                        ySide = UnitArcs.eAllArcs
                    Case SlotType.eFront
                        ySide = UnitArcs.eForwardArc
                    Case SlotType.eLeft
                        ySide = UnitArcs.eLeftArc
                    Case SlotType.eRear
                        ySide = UnitArcs.eBackArc
                    Case Else ' SlotType.eRight
                        ySide = UnitArcs.eRightArc
                End Select

                Dim lCritical As Int32
                Select Case .lSlotConfig
                    Case SlotConfig.eCargoBay
                        lCritical = elUnitStatus.eCargoBayOperational
                    Case SlotConfig.eEngines
                        lCritical = elUnitStatus.eEngineOperational
                    Case SlotConfig.eFuelBay
                        lCritical = elUnitStatus.eFuelBayOperational
                    Case SlotConfig.eHangar, SlotConfig.eHangarDoor         'TODO: Sure about the hangar door?
                        lCritical = elUnitStatus.eHangarOperational
                    Case SlotConfig.eRadar
                        lCritical = elUnitStatus.eRadarOperational
                    Case SlotConfig.eShields
                        lCritical = elUnitStatus.eShieldOperational
                    Case SlotConfig.eWeapons
                        Select Case ySide
                            Case UnitArcs.eLeftArc
                                lCritical = elUnitStatus.eLeftWeapon1 Or elUnitStatus.eLeftWeapon2
                            Case UnitArcs.eBackArc
                                lCritical = elUnitStatus.eAftWeapon1 Or elUnitStatus.eAftWeapon2
                            Case UnitArcs.eRightArc
                                lCritical = elUnitStatus.eRightWeapon1 Or elUnitStatus.eRightWeapon2
                            Case Else       'forward and all arc
                                lCritical = elUnitStatus.eForwardWeapon1 Or elUnitStatus.eForwardWeapon2
                        End Select
                End Select

                oEntityDef.AddCriticalHitChance(ySide, lCritical, CByte(.lChanceToCritical))
            End With
        Next X
        oEntityDef.ResortCriticalHitChances()
    End Sub

    'Public Shared Function MaxWpnSlots(ByVal lHull As Int32) As Int32
    '	If lHull < 120 Then
    '		Return 2
    '	ElseIf lHull < 170 Then
    '		Return 5
    '	ElseIf lHull < 200 Then
    '		Return 3
    '	ElseIf lHull < 300 Then
    '		Return 1
    '	ElseIf lHull < 500 Then
    '		Return 4
    '	ElseIf lHull < 2000 Then
    '		Return 12
    '	ElseIf lHull < 8000 Then
    '		Return 8
    '	ElseIf lHull < 20000 Then
    '		Return 12
    '	ElseIf lHull < 40000 Then
    '		Return 15
    '	ElseIf lHull < 80000 Then
    '		Return 20
    '	ElseIf lHull < 250000 Then
    '		Return 20
    '	ElseIf lHull < 1000000 Then
    '		Return 25
    '	ElseIf lHull < 4000000 Then
    '		Return 20
    '	Else : Return 40
    '	End If
    'End Function
    Public Shared Function MaxWpnSlots(ByVal lTypeID As Int32, ByVal lSubTypeID As Int32) As Int32
        Dim lHullTypeID As eyHullType = GetHullTypeID(lTypeID, lSubTypeID)

        Dim lMaxGuns As Int32 = 1
        TechBuilderComputer.GetTypeValues(lHullTypeID, 0, lMaxGuns, 0, 0, 0, 0)
        Return lMaxGuns
        'Select Case lTypeID
        '	Case 0
        '		If lSubTypeID = 0 Then
        '			Return 20	'battleship
        '		ElseIf lSubTypeID = 2 Then
        '			Return 25	'battlecruiser
        '		End If
        '	Case 1
        '		If lSubTypeID = 0 Then Return 15
        '		If lSubTypeID = 1 Then Return 20
        '		If lSubTypeID = 2 Then Return 20
        '		If lSubTypeID = 3 Then Return 12
        '		Return 8
        '	Case 2
        '		If lSubTypeID = 9 Then
        '			Return 40		'space station
        '		Else : Return 20	  'other facility
        '		End If
        '	Case 3
        '		If lSubTypeID = 0 Then Return 2 'small fighter
        '		If lSubTypeID = 1 Then Return 3 'medium fighter
        '		If lSubTypeID = 2 Then Return 4 'hvy fighter
        '	Case 4
        '		Return 5	'small vehicles
        '	Case 5
        '		Return 12	'tanks
        '	Case 6
        '		Return 4	'transports
        '	Case 7
        '		Return 1	'utility
        'End Select

        'Return 1
    End Function

    Private Structure HullSlotTypeConfigChance
        Public lSlotType As SlotType
        Public lSlotConfig As SlotConfig
        Public lOverallScore As Int32
        Public lChanceToCritical As Int32
    End Structure


    Public Shared Function GetHullTypeID(ByVal lTypeID As Int32, ByVal lSubTypeID As Int32) As Epica_Tech.eyHullType
        Select Case lTypeID
            Case 0  'capital
                If lSubTypeID = 2 Then Return eyHullType.BattleCruiser
                If lSubTypeID = 0 Then Return eyHullType.Battleship
            Case 1  'escort
                Select Case lSubTypeID
                    Case 0 : Return eyHullType.Corvette
                    Case 1 : Return eyHullType.Cruiser
                    Case 2 : Return eyHullType.Destroyer
                    Case 3 : Return eyHullType.Frigate
                    Case 4 : Return eyHullType.Escort
                End Select
            Case 2  'facility
                If lSubTypeID = 9 Then Return eyHullType.SpaceStation Else Return eyHullType.Facility
            Case 3  'fighter
                If lSubTypeID = 0 Then
                    Return eyHullType.LightFighter
                ElseIf lSubTypeID = 1 Then
                    Return eyHullType.MediumFighter
                ElseIf lSubTypeID = 2 Then
                    Return eyHullType.HeavyFighter
                End If
            Case 4  'small vehicle
                Return eyHullType.SmallVehicle
            Case 5  'tank
                Return eyHullType.Tank
            Case 6  'transport
                '????
                Return eyHullType.Utility
            Case 7  'utility vehicle
                Return eyHullType.Utility
            Case 8  'naval
                Select Case lSubTypeID
                    Case 0 : Return eyHullType.NavalBattleship
                    Case 1 : Return eyHullType.NavalCarrier
                    Case 2 : Return eyHullType.NavalCruiser
                    Case 3 : Return eyHullType.NavalDestroyer
                    Case 4 : Return eyHullType.NavalFrigate
                    Case 5 : Return eyHullType.NavalSub
                    Case 6 : Return eyHullType.Utility
                End Select
        End Select
        Return eyHullType.Utility
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
            ReDim .HullName(19)
            Array.Copy(yData, lPos, .HullName, 0, 20) : lPos += 20
            .HullSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .StructuralMineralID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .StructuralHitPoints = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .yTypeID = yData(lPos) : lPos += 1
            .ySubTypeID = yData(lPos) : lPos += 1
            .ModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .yChassisType = yData(lPos) : lPos += 1
            .mlPowerRequired = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

            Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            For X As Int32 = 0 To lCnt - 1
                Dim lX As Int32 = yData(lPos) : lPos += 1
                Dim lY As Int32 = yData(lPos) : lPos += 1
                Dim lConfig As SlotConfig = CType(yData(lPos), SlotConfig) : lPos += 1
                Dim yGroup As Byte = yData(lPos) : lPos += 1

                moSlots(lX, lY).SlotConfig = lConfig
                moSlots(lX, lY).SlotGroup = yGroup
            Next X


            If Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eResearched Then
                Dim oDef As New WeaponDef
                lPos = oDef.FillFromPrimaryMsg(yData, lPos)
            End If

            If .Owner Is Nothing = False Then
                Dim lFirstIdx As Int32 = -1
                For X As Int32 = 0 To .Owner.mlHullUB
                    If .Owner.mlHullIdx(X) = .ObjectID Then
                        .Owner.moHull(X) = Me
                        Return
                    ElseIf lFirstIdx = -1 AndAlso .Owner.mlHullIdx(X) = -1 Then
                        lFirstIdx = X
                    End If
                Next X

                If lFirstIdx = -1 Then
                    lFirstIdx = .Owner.mlHullUB + 1
                    ReDim Preserve .Owner.mlHullIdx(lFirstIdx)
                    ReDim Preserve .Owner.moHull(lFirstIdx)
                    .Owner.mlHullUB = lFirstIdx
                End If
                .Owner.moHull(lFirstIdx) = Me
                .Owner.mlHullIdx(lFirstIdx) = Me.ObjectID
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
End Class
