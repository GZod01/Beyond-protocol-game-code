Option Strict On

Public Class BombWeaponTech
    Inherits BaseWeaponTech

    Public lPayloadSize As Int32
    Public yAOE As Byte
    Public yGuidance As Byte
    Public iROF As Int16
    Public yRange As Byte
    Public yPayloadType As Byte

    Public lPayloadMatID As Int32
    Public lGuidanceMatID As Int32
    Public lCasingID As Int32

    Protected Overrides Sub CalculateBothProdCosts()

    End Sub


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

            Dim oComputer As New BombTechComputer
            With oComputer
                Dim fROF As Single = Math.Max(1, iROF) / 30.0F
                Dim lImpactMin As Int32 = 0
                Dim lImpactMax As Int32 = 0
                Dim lFlameMin As Int32 = 0
                Dim lFlameMax As Int32 = 0
                If yPayloadType = 0 Then
                    Dim lHalfPayload As Int32 = lPayloadSize \ 2
                    Dim lQtrPayload As Int32 = lHalfPayload \ 2
                    lImpactMax = lHalfPayload : lImpactMin = lQtrPayload
                    lFlameMax = lHalfPayload : lFlameMin = lQtrPayload
                End If

                .yPayloadType = yPayloadType

                Dim decNormalizer As Decimal = 1D
                Dim lMaxGuns As Int32 = 1
                Dim lMaxDPS As Int32 = 1
                Dim lMaxHullSize As Int32 = 1
                Dim lHullAvail As Int32 = 1
                Dim lMaxPower As Int32 = 1

                TechBuilderComputer.GetTypeValues(yHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)

                .decDPS = CDec((CDec(lImpactMax) + CDec(lFlameMax)) / fROF)
                .decDPS = Math.Max(.decDPS, (.decDPS + CDec(lImpactMax) + CDec(lFlameMax)) / 2D)

                .blAOE = yAOE
                .blGuidance = yGuidance
                .blPayloadSize = lPayloadSize
                .blRange = yRange
                .blROF = iROF
                .lHullTypeID = yHullTypeID

                .lMineral1ID = lCasingID
                .lMineral2ID = lGuidanceMatID
                .lMineral3ID = lPayloadMatID
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


    Protected Overrides Sub FinalizeResearch()

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
        System.BitConverter.GetBytes(iROF).CopyTo(yResult, lPos) : lPos += 2
        System.BitConverter.GetBytes(lPayloadSize).CopyTo(yResult, lPos) : lPos += 4
        yResult(lPos) = yRange : lPos += 1
        yResult(lPos) = yAOE : lPos += 1
        yResult(lPos) = yGuidance : lPos += 1
        yResult(lPos) = yPayloadType : lPos += 1

        Return yResult
    End Function

    Public Overrides Function GetPlayerTechKnowledgeMsg(ByVal yTechLvl As PlayerTechKnowledge.KnowledgeType) As Byte()
        Dim yMsg() As Byte = Nothing
        Dim lPos As Int32 = 0

        'set up our msg size
        lPos = 31
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eNameOnly Then lPos += 11
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then lPos += 22
        If yTechLvl > PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then lPos += 88 '12
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
            System.BitConverter.GetBytes(lPayloadSize).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = yAOE : lPos += 1
            yMsg(lPos) = yGuidance : lPos += 1
            yMsg(lPos) = yRange : lPos += 1
            yMsg(lPos) = yPayloadType : lPos += 1
            System.BitConverter.GetBytes(iROF).CopyTo(yMsg, lPos) : lPos += 2

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
            System.BitConverter.GetBytes(lPayloadMatID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lGuidanceMatID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lCasingID).CopyTo(yMsg, lPos) : lPos += 4

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
            sSQL &= iROF & ", ShotHullSize = " & lPayloadSize & ", OptimumRange = " & _
              yRange & ", PayloadType = " & yPayloadType & ", ExplosionRadius = " & _
               yAOE & ", Mineral1ID = " & lPayloadMatID & ", Mineral2ID = " & lGuidanceMatID & ", Mineral3ID = " & _
               lCasingID & ", ErrorReasonCode = " & ErrorReasonCode & ", MajorDesignFlaw = " & MajorDesignFlaw & ", PopIntel = " & _
               PopIntel & ", bArchived = " & yArchived & ", Accuracy = " & yGuidance & ", HullTypeID = " & yHullTypeID & " WHERE WeaponID = " & Me.ObjectID

            Return sSQL
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        End Try
        Return ""
    End Function

    Public Overrides Function GetWeaponDefResult() As WeaponDef
        Dim oWpnDef As WeaponDef = New WeaponDef
        Dim fTemp As Single = 0.0F

        With oWpnDef
            .Accuracy = yGuidance
            .AmmoReloadDelay = 0
            .AmmoSize = 0
            .AOERange = yAOE

            Dim lHalfPayload As Int32 = 0
            Dim lQtrPayload As Int32 = 0
            If yPayloadType = 0 Then
                lHalfPayload = lPayloadSize \ 2
                lQtrPayload = lHalfPayload \ 2
            End If

            .BeamMaxDmg = 0 : .BeamMinDmg = 0
            .PiercingMaxDmg = 0 : .PiercingMinDmg = 0
            .ImpactMaxDmg = lHalfPayload : .ImpactMinDmg = lQtrPayload
            .ChemicalMaxDmg = 0 : .ChemicalMinDmg = 0
            .ECMMaxDmg = 0 : .ECMMinDmg = 0
            .FlameMaxDmg = lHalfPayload : .FlameMinDmg = lQtrPayload

            .ObjectID = -1
            .ObjTypeID = ObjectType.eWeaponDef

            .Range = yRange
            .MaxSpeed = 100
            .Maneuver = 0

            .RelatedWeapon = Me
            .ROF = iROF
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
                   "ShotHullSize, OptimumRange, PayloadType, ExplosionRadius, Mineral1ID, Mineral2ID, Mineral3ID, ErrorReasonCode, " & _
                   "MajorDesignFlaw, PopIntel, bArchived, Accuracy, HullTypeID) VALUES ('" & _
                   MakeDBStr(BytesToString(WeaponName)) & "', " & CByte(WeaponClassTypeID) & ", " & CByte(WeaponTypeID) & ", " & _
                   PowerRequired & ", " & HullRequired & ", " & CurrentSuccessChance & ", " & SuccessChanceIncrement & ", " & _
                   ResearchAttempts & ", " & RandomSeed & ", " & CInt(ComponentDevelopmentPhase) & ", "
                If Me.Owner Is Nothing = False Then
                    sSQL &= Owner.ObjectID & ", "
                Else : sSQL &= "-1, "
                End If
                sSQL &= iROF & ", " & lPayloadSize & ", " & yRange & ", " & yPayloadType & ", " & yAOE & ", " & lPayloadMatID & _
                 ", " & lGuidanceMatID & ", " & lCasingID & ", " & ErrorReasonCode & ", " & MajorDesignFlaw & ", " & PopIntel & _
                 ", " & yArchived & ", " & yGuidance & ", " & yHullTypeID & ")"
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
                sSQL &= iROF & ", ShotHullSize = " & lPayloadSize & ", OptimumRange = " & _
                  yRange & ", PayloadType = " & yPayloadType & ", ExplosionRadius = " & _
                   yAOE & ", Mineral1ID = " & lPayloadMatID & ", Mineral2ID = " & lGuidanceMatID & ", Mineral3ID = " & _
                   lCasingID & ", ErrorReasonCode = " & ErrorReasonCode & ", MajorDesignFlaw = " & MajorDesignFlaw & ", PopIntel = " & _
                   PopIntel & ", bArchived = " & yArchived & ", Accuracy = " & yGuidance & ", HullTypeID = " & yHullTypeID & " WHERE WeaponID = " & Me.ObjectID
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
        Dim bResult As Boolean = False
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

                .lPayloadSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .yAOE = yData(lPos) : lPos += 1
                .yGuidance = yData(lPos) : lPos += 1
                .iROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .yRange = yData(lPos) : lPos += 1
                .yPayloadType = yData(lPos) : lPos += 1
                .lPayloadMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lGuidanceMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lCasingID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                ReDim .WeaponName(19)
                Array.Copy(yData, lPos, .WeaponName, 0, 20)
                lPos += 20
            End With

            bResult = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "ProjectileWeaponTech.SetFromDesignMsg: " & ex.Message)
        End Try

        Return bResult
    End Function

    Public Overrides Function TechnologyScore() As Integer
        If mlStoredTechScore = Int32.MinValue Then
            Try
                Dim fROF As Single = iROF / 30.0F
                mlStoredTechScore = CInt(((((lPayloadSize + CInt(yAOE)) / fROF) * CInt(yRange) * (1.0F + (CInt(yPayloadType) / 10.0F))) / (Me.PowerRequired + Me.HullRequired + 1)) * 10)
            Catch
                mlStoredTechScore = 1000
            End Try
        End If
        Return mlStoredTechScore
    End Function

    Public Overrides Function ValidateDesign() As Boolean
        Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidMaterialsSelected
        If lPayloadMatID < 1 OrElse Owner.IsMineralDiscovered(lPayloadMatID) = False Then
            LogEvent(LogEventType.PossibleCheat, "BombWeaponTech.ValidateDesign invalid material: " & Me.Owner.ObjectID)
            Return False
        End If
        If lGuidanceMatID < 1 OrElse Owner.IsMineralDiscovered(lGuidanceMatID) = False Then
            LogEvent(LogEventType.PossibleCheat, "BombWeaponTech.ValidateDesign invalid material: " & Me.Owner.ObjectID)
            Return False
        End If
        If lCasingID < 1 OrElse Owner.IsMineralDiscovered(lCasingID) = False Then
            LogEvent(LogEventType.PossibleCheat, "BombWeaponTech.ValidateDesign invalid material: " & Me.Owner.ObjectID)
            Return False
        End If

        Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
        If lPayloadSize < 2 OrElse lPayloadSize > 90000 Then
            LogEvent(LogEventType.PossibleCheat, "BombWeaponTech.ValidateDesign invalid value: " & Me.Owner.ObjectID)
            Return False
        End If
        If yAOE > Owner.oSpecials.MaxBombAOE Then
            LogEvent(LogEventType.PossibleCheat, "BombWeaponTech.ValidateDesign invalid value: " & Me.Owner.ObjectID)
            Return False
        End If
        If yGuidance > Owner.oSpecials.MaxBombGuidance Then
            LogEvent(LogEventType.PossibleCheat, "BombWeaponTech.ValidateDesign invalid value: " & Me.Owner.ObjectID)
            Return False
        End If
        If iROF < Owner.oSpecials.MinBombROF Then
            LogEvent(LogEventType.PossibleCheat, "BombWeaponTech.ValidateDesign invalid value: " & Me.Owner.ObjectID)
            Return False
        End If
        If yRange > Owner.oSpecials.MaxBombRange Then
            LogEvent(LogEventType.PossibleCheat, "BombWeaponTech.ValidateDesign invalid value: " & Me.Owner.ObjectID)
            Return False
        End If

        Dim yAllowedType As Byte = Owner.oSpecials.BombPayloadType
        Select Case yPayloadType
            Case 1
                If (yAllowedType And 1) = 0 Then
                    LogEvent(LogEventType.PossibleCheat, "BombWeaponTech.ValidateDesign invalid value: " & Me.Owner.ObjectID)
                    Return False
                End If
            Case 2
                If (yAllowedType And 2) = 0 Then
                    LogEvent(LogEventType.PossibleCheat, "BombWeaponTech.ValidateDesign invalid value: " & Me.Owner.ObjectID)
                    Return False
                End If
            Case 3
                If (yAllowedType And 4) = 0 Then
                    LogEvent(LogEventType.PossibleCheat, "BombWeaponTech.ValidateDesign invalid value: " & Me.Owner.ObjectID)
                    Return False
                End If
        End Select

        If CheckDesignImpossibility() = True Then
            LogEvent(LogEventType.PossibleCheat, "BombWeaponTech.ValidateDesign CheckDesignPossiblity Failed: " & Me.Owner.ObjectID)
            Me.ErrorReasonCode = TechBuilderErrorReason.eInvalidValuesEntered
            Return False
        End If

        Me.ErrorReasonCode = TechBuilderErrorReason.eNoError
        Return True
    End Function

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
            ReDim mySendString(BASE_WEAPON_MSG_HEADER_LENGTH + yDefMsg.Length + 21)
        Else : ReDim mySendString(BASE_WEAPON_MSG_HEADER_LENGTH + 21)
        End If

        Dim lPos As Int32 = MyBase.FillBaseWeaponMsgHdr()

        System.BitConverter.GetBytes(lPayloadSize).CopyTo(mySendString, lPos) : lPos += 4
        mySendString(lPos) = yAOE : lPos += 1
        mySendString(lPos) = yGuidance : lPos += 1
        mySendString(lPos) = yRange : lPos += 1
        mySendString(lPos) = yPayloadType : lPos += 1
        System.BitConverter.GetBytes(iROF).CopyTo(mySendString, lPos) : lPos += 2
        System.BitConverter.GetBytes(lPayloadMatID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lGuidanceMatID).CopyTo(mySendString, lPos) : lPos += 4
        System.BitConverter.GetBytes(lCasingID).CopyTo(mySendString, lPos) : lPos += 4


        If yDefMsg Is Nothing = False Then
            yDefMsg.CopyTo(mySendString, lPos) : lPos += yDefMsg.Length
        End If

        'mbStringReady = True
        'End If
        Return mySendString
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

            .lPayloadSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .yAOE = yData(lPos) : lPos += 1
            .yGuidance = yData(lPos) : lPos += 1
            .yRange = yData(lPos) : lPos += 1
            .yPayloadType = yData(lPos) : lPos += 1
            .iROF = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .lPayloadMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lGuidanceMatID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .lCasingID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

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

        Dim oComputer As New BombTechComputer
        With oComputer
            Dim fROF As Single = Math.Max(1, iROF) / 30.0F
            Dim lImpactMin As Int32 = 0
            Dim lImpactMax As Int32 = 0
            Dim lFlameMin As Int32 = 0
            Dim lFlameMax As Int32 = 0
            If yPayloadType = 0 Then
                Dim lHalfPayload As Int32 = lPayloadSize \ 2
                Dim lQtrPayload As Int32 = lHalfPayload \ 2
                lImpactMax = lHalfPayload : lImpactMin = lQtrPayload
                lFlameMax = lHalfPayload : lFlameMin = lQtrPayload
            End If

            .yPayloadType = yPayloadType

            Dim decNormalizer As Decimal = 1D
            Dim lMaxGuns As Int32 = 1
            Dim lMaxDPS As Int32 = 1
            Dim lMaxHullSize As Int32 = 1
            Dim lHullAvail As Int32 = 1
            Dim lMaxPower As Int32 = 1

            TechBuilderComputer.GetTypeValues(yHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)

            .decDPS = CDec((CDec(lImpactMax) + CDec(lFlameMax)) / fROF)
            .decDPS = Math.Max(.decDPS, (.decDPS + CDec(lImpactMax) + CDec(lFlameMax)) / 2D)

            .blAOE = yAOE
            .blGuidance = yGuidance
            .blPayloadSize = lPayloadSize
            .blRange = yRange
            .blROF = iROF
            .lHullTypeID = yHullTypeID

            .lMineral1ID = lCasingID
            .lMineral2ID = lGuidanceMatID
            .lMineral3ID = lPayloadMatID
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
