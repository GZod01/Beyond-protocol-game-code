Public MustInherit Class Epica_Tech
    Inherits Epica_GUID

    Public Enum eyHullType As Byte
        BattleCruiser = 11
        Battleship = 12
        Corvette = 8
        Cruiser = 10
        Destroyer = 9
        Escort = 6
        Facility = 14
        Frigate = 7
        HeavyFighter = 4
        LightFighter = 0
        MediumFighter = 2
        NavalBattleship = 15
        NavalCarrier = 16
        NavalCruiser = 17
        NavalDestroyer = 18
        NavalFrigate = 19
        NavalSub = 20
        SmallVehicle = 1
        SpaceStation = 13
        Tank = 5
        Utility = 3
    End Enum

    Public Enum eComponentDevelopmentPhase As Integer
        eInvalidDesign = -1         'The design is invalid
        eComponentDesign = 0        'initial phase, component is being designed
        eComponentResearching       'component is designed, now it is being researched
        eResearched                 'component was designed and was researched successfully 
    End Enum

    Public Enum eComponentDesignFlaw As Byte
        eNo_Particular_Design_Flaw = 0
        eMat1_Prop1 = 1
        eMat1_Prop2 = 2
        eMat1_Prop3 = 3
        eMat1_Prop4 = 4
        eMat1_Prop5 = 5
        eMat2_Prop1 = 6
        eMat2_Prop2 = 7
        eMat2_Prop3 = 8
        eMat2_Prop4 = 9
        eMat2_Prop5 = 10
        eMat3_Prop1 = 11
        eMat3_Prop2 = 12
        eMat3_Prop3 = 13
        eMat3_Prop4 = 14
        eMat3_Prop5 = 15
        eMat4_Prop1 = 16
        eMat4_Prop2 = 17
        eMat4_Prop3 = 18
        eMat4_Prop4 = 19
        eMat4_Prop5 = 20
        eMat5_Prop1 = 21
        eMat5_Prop2 = 22
        eMat5_Prop3 = 23
        eMat5_Prop4 = 24
        eMat5_Prop5 = 25
        eMat6_Prop1 = 26
        eMat6_Prop2 = 27
        eMat6_Prop3 = 28
        eMat6_Prop4 = 29
		eMat6_Prop5 = 30

        eMicroTech = 31

		eGoodDesign = 32
		eShift_Should_Be_Higher = 64
		eShift_Not_Study = 128
    End Enum

    Public ComponentDevelopmentPhase As eComponentDevelopmentPhase = eComponentDevelopmentPhase.eComponentDesign
    Protected moResearchCost As ProductionCost
    Protected Shared moDesignCost As ProductionCost
    Protected moProductionCost As ProductionCost
    Public Owner As Player

    'ok, all techs (generally) have a prod cost configuration with them, if these are -1 then there is no configuration for that value
    Public lSpecifiedHull As Int32 = -1
    Public lSpecifiedPower As Int32 = -1
    Public lSpecifiedColonists As Int32 = -1
    Public lSpecifiedEnlisted As Int32 = -1
    Public lSpecifiedOfficers As Int32 = -1
    Public blSpecifiedProdCost As Int64 = -1
    Public blSpecifiedProdTime As Int64 = -1
    Public blSpecifiedResCost As Int64 = -1
    Public blSpecifiedResTime As Int64 = -1
    Public lSpecifiedMin1 As Int32 = -1
    Public lSpecifiedMin2 As Int32 = -1
    Public lSpecifiedMin3 As Int32 = -1
    Public lSpecifiedMin4 As Int32 = -1
    Public lSpecifiedMin5 As Int32 = -1
    Public lSpecifiedMin6 As Int32 = -1

    Public MustOverride Function ValidateDesign() As Boolean
    
    Protected MustOverride Sub FinalizeResearch()
    Public MustOverride Sub ComponentDesigned()
    Protected MustOverride Sub CalculateBothProdCosts()

    Public MustOverride Function SaveObject() As Boolean
    Public MustOverride Function SetFromDesignMsg(ByRef yData() As Byte) As Boolean
    Public MustOverride Function GetNonOwnerMsg() As Byte()
    Public MustOverride Function TechnologyScore() As Int32
    Public MustOverride Function GetPlayerTechKnowledgeMsg(ByVal yTechLvl As PlayerTechKnowledge.KnowledgeType) As Byte()
    Public MustOverride Sub FillFromPrimaryAddMsg(ByVal yData() As Byte)

    Public CurrentSuccessChance As Int32
    Public SuccessChanceIncrement As Int32 = 1
    Public ResearchAttempts As Int32

    Public ErrorReasonCode As Byte
    Public MajorDesignFlaw As Byte 
    Public PopIntel As Int32 = 100

    Public yArchived As Byte = 0

    Public Shared TechVersionNum As Int32 = 5
    Public lVersionNum As Int32 = TechVersionNum

    Protected mlStoredTechScore As Int32 = Int32.MinValue

#Region "  Versioning  "
    Public Shared VersionList() As VersionRel
    Public Shared VersionListUB As Int32 = -1
    Public Shared Sub AddVersionRel(ByVal oRel As VersionRel)
        VersionListUB += 1
        ReDim Preserve VersionList(VersionListUB)
        VersionList(VersionListUB) = oRel
    End Sub
#End Region

    Private mfRandVal As Single = 0.0F
    Public Property RandomSeed() As Single
        Get
            If mfRandVal = 0.0F Then
                Randomize(Val(Now.ToString("MMddHHmm")))
                mfRandVal = Rnd() * 1.0F
            End If
            Return mfRandVal
        End Get
        Set(ByVal value As Single)
            mfRandVal = value
        End Set
    End Property

    ''' <summary>
    ''' Call this routine when production of research points reaches the researchpoints cost...
    ''' </summary>
    ''' <returns>Boolean indicating whether the research succeeded or not.</returns>
    ''' <remarks></remarks>
    Public Function ResearchComplete() As Boolean
        mbStringReady = False
        'bNeedsSaved = True
        If Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eComponentDesign Then
            'Always true and handled in this sub...
            ComponentDevelopmentPhase = eComponentDevelopmentPhase.eComponentResearching

            'clear our production and research costs
            'TODO: Make this better, right now, by deleting an recreating, we use up more ID's
            'If Me.moProductionCost Is Nothing = False Then Me.moProductionCost.DeleteMe()
            'If Me.moResearchCost Is Nothing = False Then Me.moResearchCost.DeleteMe()
            'Me.moProductionCost = Nothing : Me.moResearchCost = Nothing
            If moProductionCost Is Nothing = False Then moProductionCost.DeleteItems()
            If moResearchCost Is Nothing = False Then moResearchCost.DeleteItems()
            ResearchAttempts = 0

            'Tell the component it was designed
            ComponentDesigned()
        ElseIf Me.ComponentDevelopmentPhase <> eComponentDevelopmentPhase.eInvalidDesign Then
            'ResearchAttempts += 1       'increment our research attempts
            ''Actual Research is chance based
            'Dim lRoll As Int32 = CInt(Int(Rnd() * 100) + 1)
            ''Result must be less than or equal to our chance to succeed
            'If lRoll <= CurrentSuccessChance Then
            '    ComponentDevelopmentPhase = eComponentDevelopmentPhase.eResearched
            '    ComputeResults()        'every component has this routine which does the finalization work...
            '    Return True
            'Else
            '    CurrentSuccessChance += SuccessChanceIncrement
            '    Return False
            'End If
            ComponentDevelopmentPhase = eComponentDevelopmentPhase.eResearched
            FinalizeResearch()

            'Ok, now, send this tech to the operator
            Try
                Dim yOpMsg() As Byte = goMsgSys.GetAddObjectMessage(Me, GlobalMessageCode.eAddObjectCommand)
                Dim yResMsg() As Byte = Me.GetProductionCost.GetObjAsString()
                Dim yProdMsg() As Byte = Me.GetProductionCost.GetObjAsString()
                Dim yFinal(yOpMsg.Length + yResMsg.Length + yProdMsg.Length - 1) As Byte
                Dim lPos As Int32 = 0
                yOpMsg.CopyTo(yFinal, lPos) : lPos += yOpMsg.Length
                yResMsg.CopyTo(yFinal, lPos) : lPos += yResMsg.Length
                yProdMsg.CopyTo(yFinal, lPos) : lPos += yProdMsg.Length
            Catch
            End Try

            Me.Owner.TestCustomTitlePermissions_Research(Me.ObjTypeID)
        End If

        Try
            If mlPrimaryIdx <> -1 Then
                If glFacilityIdx(mlPrimaryIdx) = mlPrimaryID Then
                    Dim oColony As Colony = goFacility(mlPrimaryIdx).ParentColony
                    If oColony Is Nothing = False Then
                        AddToQueue(glCurrentCycle + 30, QueueItemType.eCheckColonyResearchQueue, oColony.ObjectID, 0, 0, 0, 0, 0, 0, 0)
                    End If
                End If
            End If
        Catch
        End Try

        'Ok, our research is complete, so set all assisters to clear and the primary researcher
        mlPrimaryIdx = -1
        Dim lTmp As Int32 = mlAssisterUB
        For X As Int32 = 0 To lTmp
            If mlAssisters(X) <> -1 AndAlso glFacilityIdx(mlAssisters(X)) <> -1 AndAlso glFacilityIdx(mlAssisters(X)) = mlAssisterID(X) Then
                Dim lVal As Int32 = mlAssisters(X)
                mlAssisters(X) = -1
                Dim oFac As Facility = goFacility(lVal)
                If oFac Is Nothing = False Then
                    oFac.CurrentProduction.lProdCount = 0
                    oFac.CurrentProduction.PointsProduced = 0
                    oFac.CurrentProduction.lLastUpdateCycle = glCurrentCycle
                    oFac.GetNextProduction(False)
                End If
            End If
        Next X
        If mlPrimaryIdx = -1 Then ReDim mlAssisters(-1)
        mlAssisterUB = -1

        Me.SaveObject()

        'Return true that research completed successfully
        Return True
    End Function

    Public Const BASE_OBJ_STRING_SIZE As Int32 = 93 '17
    Public Const BASE_OBJ_STRING_SIZE_EXCLUDED As Int32 = 17
    Protected Function GetBaseObjAsString(ByVal bExcludeSpecifieds As Boolean) As Byte()

        If bExcludeSpecifieds = True Then
            Dim yResult(BASE_OBJ_STRING_SIZE_EXCLUDED - 1) As Byte
            GetGUIDAsString.CopyTo(yResult, 0)
            System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(yResult, 6)
            System.BitConverter.GetBytes(ComponentDevelopmentPhase).CopyTo(yResult, 10)
            yResult(14) = ErrorReasonCode

            Dim lTemp As Int32 = GetResearcherCount()
            If lTemp > 255 Then lTemp = 255
            If lTemp < 0 Then lTemp = 0
            yResult(15) = CByte(lTemp)
            yResult(16) = MajorDesignFlaw
            Return yResult
        Else
            Dim yResult(BASE_OBJ_STRING_SIZE - 1) As Byte
            GetGUIDAsString.CopyTo(yResult, 0)
            System.BitConverter.GetBytes(Owner.ObjectID).CopyTo(yResult, 6)
            System.BitConverter.GetBytes(ComponentDevelopmentPhase).CopyTo(yResult, 10)
            yResult(14) = ErrorReasonCode

            Dim lTemp As Int32 = GetResearcherCount()
            If lTemp > 255 Then lTemp = 255
            If lTemp < 0 Then lTemp = 0
            yResult(15) = CByte(lTemp)
            yResult(16) = MajorDesignFlaw

            Dim lPos As Int32 = 17
            System.BitConverter.GetBytes(lSpecifiedHull).CopyTo(yResult, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedPower).CopyTo(yResult, lPos) : lPos += 4
            System.BitConverter.GetBytes(blSpecifiedResCost).CopyTo(yResult, lPos) : lPos += 8
            System.BitConverter.GetBytes(blSpecifiedResTime).CopyTo(yResult, lPos) : lPos += 8
            System.BitConverter.GetBytes(blSpecifiedProdCost).CopyTo(yResult, lPos) : lPos += 8
            System.BitConverter.GetBytes(blSpecifiedProdTime).CopyTo(yResult, lPos) : lPos += 8
            System.BitConverter.GetBytes(lSpecifiedColonists).CopyTo(yResult, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedEnlisted).CopyTo(yResult, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedOfficers).CopyTo(yResult, lPos) : lPos += 4

            System.BitConverter.GetBytes(lSpecifiedMin1).CopyTo(yResult, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedMin2).CopyTo(yResult, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedMin3).CopyTo(yResult, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedMin4).CopyTo(yResult, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedMin5).CopyTo(yResult, lPos) : lPos += 4
            System.BitConverter.GetBytes(lSpecifiedMin6).CopyTo(yResult, lPos) : lPos += 4

            Return yResult
        End If
    End Function

    Protected Function GetResearchPointsCost(ByVal blBaseCost As Int64) As Int64
        Dim bDone As Boolean = False
        Dim blResult As Int64 = 0
        Dim lRoll As Int32

        If ResearchAttempts = 0 Then
            If RandomSeed = 0 OrElse RandomSeed = 1.0F Then
                Randomize(Val(Now.ToString("MMddHHmm")))
                RandomSeed = Rnd() * 1.0F
            End If
            Rnd(-1)
            Randomize(RandomSeed)

            While bDone = False
                ResearchAttempts += 1
                blResult += blBaseCost
                lRoll = CInt(Int(Rnd() * 100) + 1)
                If lRoll <= CurrentSuccessChance Then
                    bDone = True
                Else
                    CurrentSuccessChance += SuccessChanceIncrement
                End If
            End While
            Randomize(Val(Now.ToString("MMddHHmm")))

        Else
            blResult = ResearchAttempts * blBaseCost
        End If

        Return blResult
    End Function

    Public Shared Function GetMinVal(ByVal ParamArray Values() As Int32) As Int32
        Dim lMinVal As Int32 = Int32.MaxValue

        For X As Int32 = 0 To Values.Length - 1
            lMinVal = Math.Min(lMinVal, Values(X))
        Next X

        Return lMinVal
    End Function

    Public Shared Function GetMaxVal(ByVal ParamArray Values() As Int32) As Int32
        Dim lMaxVal As Int32 = Int32.MinValue

        For X As Int32 = 0 To Values.Length - 1
            lMaxVal = Math.Max(lMaxVal, Values(X))
        Next

        Return lMaxVal
    End Function

    Public Function GetCurrentProductionCost() As ProductionCost
        If ComponentDevelopmentPhase = eComponentDevelopmentPhase.eComponentDesign Then
            If moDesignCost Is Nothing Then
                moDesignCost = New ProductionCost
                With moDesignCost
                    .ColonistCost = 0 : .OfficerCost = 0 : .EnlistedCost = 0
                    ReDim .ItemCosts(-1)
                    .ItemCostUB = -1
                    .CreditCost = 0
                    .PointsRequired = 60000
                End With
            End If

            moDesignCost.ObjectID = Me.ObjectID
            moDesignCost.ObjTypeID = Me.ObjTypeID
            Return moDesignCost
        ElseIf ComponentDevelopmentPhase = eComponentDevelopmentPhase.eComponentResearching Then
            Return GetResearchCost()
        Else
            Return GetProductionCost()
        End If
    End Function

    Public Function GetResearchCost() As ProductionCost
        If moResearchCost Is Nothing Then CalculateBothProdCosts()
        If moResearchCost Is Nothing = False Then moResearchCost.ProductionCostType = 1

        If moResearchCost Is Nothing = False AndAlso moResearchCost.CreditCost < 0 Then
            Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
            LogEvent(LogEventType.Warning, "Tech Research Costs < 0. TechID: " & Me.ObjectID & ", TechTypeID: " & Me.ObjTypeID)
        End If

        Return moResearchCost
    End Function

    Public Function GetProductionCost() As ProductionCost
        If moProductionCost Is Nothing Then CalculateBothProdCosts()
        If moProductionCost Is Nothing = False Then moProductionCost.ProductionCostType = 0

        If moProductionCost Is Nothing = False AndAlso moProductionCost.CreditCost < 0 Then
            Me.ComponentDevelopmentPhase = eComponentDevelopmentPhase.eInvalidDesign
            LogEvent(LogEventType.Warning, "Tech Production Costs < 0. TechID: " & Me.ObjectID & ", TechTypeID: " & Me.ObjTypeID)
        End If

        Return moProductionCost
    End Function

    Protected Sub FinalizeSave()
        If moResearchCost Is Nothing = False Then
            moResearchCost.ProductionCostType = 1
            moResearchCost.SaveObject()
        End If
        If moProductionCost Is Nothing = False Then
            moProductionCost.ProductionCostType = 0
            moProductionCost.SaveObject()
        End If

        Dim sTable As String = ""
        Dim sIDField As String = ""
        Select Case Me.ObjTypeID
            Case ObjectType.eArmorTech
                sTable = "tblArmor"
                sIDField = "ArmorID"
            Case ObjectType.eEngineTech
                sTable = "tblEngine"
                sIDField = "EngineID"
            Case ObjectType.eRadarTech
                sTable = "tblRadar"
                sIDField = "RadarID"
            Case ObjectType.eShieldTech
                sTable = "tblShield"
                sIDField = "ShieldID"
            Case ObjectType.eWeaponTech
                sTable = "tblWeapon"
                sIDField = "WeaponID"
        End Select

        If sTable <> "" Then
            Dim sSQL As String = "UPDATE " & sTable & " SET SpecifiedHull = " & lSpecifiedHull & ", SpecifiedPower = " & _
                lSpecifiedPower & ", SpecifiedResCost = " & blSpecifiedResCost & ", SpecifiedResTime = " & _
                blSpecifiedResTime & ", SpecifiedProdCost = " & blSpecifiedProdCost & ", SpecifiedProdTime = " & _
                blSpecifiedProdTime & ", SpecifiedColonist = " & lSpecifiedColonists & ", SpecifiedEnlisted = " & _
                lSpecifiedEnlisted & ", SpecifiedOfficer = " & lSpecifiedOfficers & ", SpecifiedMin1 = " & _
                lSpecifiedMin1 & ", SpecifiedMin2 = " & lSpecifiedMin2 & ", SpecifiedMin3 = " & lSpecifiedMin3 & _
                ", SpecifiedMin4 = " & lSpecifiedMin4 & ", SpecifiedMin5 = " & lSpecifiedMin5 & ", SpecifiedMin6 = " & _
                lSpecifiedMin6 & ", VersionNumber = " & lVersionNum & " WHERE " & sIDField & " = " & Me.ObjectID
            Dim oComm As New OleDb.OleDbCommand(sSQL, goCN)
            Try
                oComm.ExecuteNonQuery()
            Catch ex As Exception
                LogEvent(LogEventType.CriticalError, "Unable to save Epica_Tech.FinalizeSave of Specifieds: " & ex.Message)
            End Try
            oComm.Dispose()
            oComm = Nothing
        End If

        AddQuickTechnologyLookup(Me.ObjectID, Me.ObjTypeID, Me.Owner.ObjectID)
    End Sub

    Protected Function GetFinalizeSaveText() As String
        Dim oSB As New System.Text.StringBuilder

        If moResearchCost Is Nothing = False Then
            moResearchCost.ProductionCostType = 1
            oSB.AppendLine(moResearchCost.GetSaveObjectText())
        End If
        If moProductionCost Is Nothing = False Then
            moProductionCost.ProductionCostType = 0
            oSB.AppendLine(moProductionCost.GetSaveObjectText())
        End If

        Dim sTable As String = ""
        Dim sIDField As String = ""
        Select Case Me.ObjTypeID
            Case ObjectType.eArmorTech
                sTable = "tblArmor"
                sIDField = "ArmorID"
            Case ObjectType.eEngineTech
                sTable = "tblEngine"
                sIDField = "EngineID"
            Case ObjectType.eRadarTech
                sTable = "tblRadar"
                sIDField = "RadarID"
            Case ObjectType.eShieldTech
                sTable = "tblShield"
                sIDField = "ShieldID"
            Case ObjectType.eWeaponTech
                sTable = "tblWeapon"
                sIDField = "WeaponID"
        End Select

        If sTable <> "" Then
            Dim sSQL As String = "UPDATE " & sTable & " SET SpecifiedHull = " & lSpecifiedHull & ", SpecifiedPower = " & _
                lSpecifiedPower & ", SpecifiedResCost = " & blSpecifiedResCost & ", SpecifiedResTime = " & _
                blSpecifiedResTime & ", SpecifiedProdCost = " & blSpecifiedProdCost & ", SpecifiedProdTime = " & _
                blSpecifiedProdTime & ", SpecifiedColonist = " & lSpecifiedColonists & ", SpecifiedEnlisted = " & _
                lSpecifiedEnlisted & ", SpecifiedOfficer = " & lSpecifiedOfficers & ", SpecifiedMin1 = " & _
                lSpecifiedMin1 & ", SpecifiedMin2 = " & lSpecifiedMin2 & ", SpecifiedMin3 = " & lSpecifiedMin3 & _
                ", SpecifiedMin4 = " & lSpecifiedMin4 & ", SpecifiedMin5 = " & lSpecifiedMin5 & ", SpecifiedMin6 = " & _
                lSpecifiedMin6 & " WHERE " & sIDField & " = " & Me.ObjectID
            oSB.AppendLine(sSQL)
        End If

        Return oSB.ToString
    End Function

#Region " Multiple Researcher Code "
    'TODO: Right now, only facilities can research, if this should ever change, we will need to change this
    Private mlPrimaryIdx As Int32 = -1  'Index in the Facility array
    Private mlPrimaryID As Int32 = -1   'ID of the facility
    Private mlAssisters() As Int32
    Private mlAssisterID() As Int32     'ID of the assister
    Private mlAssisterUB As Int32 = -1

    ''' <summary>
    ''' Call this when an entity begins researching this tech as it will handle "Primary" vs. "Assists"
    ''' </summary>
    ''' <param name="lFacilityID"> The Object ID of the facility </param>
    ''' <remarks></remarks>
    Public Sub AssignResearcher(ByVal lFacilityID As Int32)
        Dim lFacilityIndex As Int32 = -1
        For X As Int32 = 0 To glFacilityUB
            If glFacilityIdx(X) = lFacilityID Then
                lFacilityIndex = X
                Exit For
            End If
        Next X
        If lFacilityIndex = -1 Then Return

        If mlPrimaryIdx = -1 OrElse glFacilityIdx(mlPrimaryIdx) = -1 OrElse glFacilityIdx(mlPrimaryIdx) <> mlPrimaryID Then
            'Ok, this researcher will be the primary
            mlPrimaryIdx = lFacilityIndex
            mlPrimaryID = lFacilityID
        Else
            'Ok, this researcher will be an assister
            Dim lIdx As Int32 = -1

            For X As Int32 = 0 To mlAssisterUB
                If mlAssisters(X) = lFacilityIndex AndAlso mlAssisterID(X) = lFacilityID Then
                    'already there... hmmm, for now, exit for
                    lIdx = X
                    Exit For
                ElseIf lIdx = -1 AndAlso mlAssisters(X) < 1 Then
                    lIdx = X
                End If
            Next X

            If lIdx = -1 Then
                mlAssisterUB += 1
                ReDim Preserve mlAssisters(mlAssisterUB)
                ReDim Preserve mlAssisterID(mlAssisterUB)
                lIdx = mlAssisterUB
            End If
            mlAssisters(lIdx) = lFacilityIndex
            mlAssisterID(lIdx) = lFacilityID

            'facility never ends assisting until the primary is done
            goFacility(mlAssisters(lIdx)).CurrentProduction.lFinishCycle = Int32.MaxValue

            If (Me.Owner.lCustomTitlePermission And elCustomRankPermissions.ChiefScientist) = 0 Then
                If MeetsCriteriaForChiefScientist() = True Then
                    Dim lOriginal As Int32 = Me.Owner.lCustomTitlePermission
                    Me.Owner.lCustomTitlePermission = Me.Owner.lCustomTitlePermission Or elCustomRankPermissions.ChiefScientist
                    Me.Owner.ProcessCustomTitlePermissionChange(lOriginal)
                End If
            End If
        End If

        'Ok, now, invoke the Primary's Recalc function
        goFacility(mlPrimaryIdx).RecalcProduction()

        If Me.Owner.lConnectedPrimaryID > -1 OrElse Me.Owner.HasOnlineAliases(AliasingRights.eViewResearch) = True Then
            Dim lCnt As Int32 = GetResearcherCount()
            Dim yCnt As Byte = CByte(lCnt)

            If lCnt > 255 Then lCnt = 255
            If lCnt < 0 Then lCnt = 0

            Dim yMsg(8) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eUpdateResearcherCnt).CopyTo(yMsg, 0)
            Me.GetGUIDAsString.CopyTo(yMsg, 2)
            yMsg(8) = yCnt
            Owner.SendPlayerMessage(yMsg, False, AliasingRights.eViewResearch)
        End If
    End Sub

    Private Function MeetsCriteriaForChiefScientist() As Boolean
        Dim lEnvirs(mlAssisterUB + 1) As Int32
        Dim lEnvirCnt(mlAssisterUB + 1) As Int32
        Dim lEnvirUB As Int32 = -1

        If mlPrimaryIdx > -1 AndAlso glFacilityIdx(mlPrimaryIdx) = mlPrimaryID Then
            Dim oFac As Facility = goFacility(mlPrimaryIdx)
            If oFac Is Nothing = False Then
                If oFac.EntityDef.ProdFactor > 400 Then
                    If oFac.ParentColony Is Nothing = False Then
                        Dim lParentID As Int32 = oFac.ParentColony.ObjectID
                        Dim bFound As Boolean = False
                        For Y As Int32 = 0 To lEnvirUB
                            If lEnvirs(Y) = lParentID Then
                                lEnvirCnt(Y) += 1
                                Exit For
                            End If
                        Next Y
                        If bFound = False Then
                            lEnvirUB += 1
                            lEnvirs(lEnvirUB) = lParentID
                            lEnvirCnt(lEnvirUB) += 1
                        End If
                    End If
                End If
            End If
        End If

        For X As Int32 = 0 To mlAssisterUB
            If mlAssisters(X) > -1 AndAlso glFacilityIdx(mlAssisters(X)) = mlAssisterID(X) Then
                Dim oFac As Facility = goFacility(mlAssisters(X))
                If oFac Is Nothing = False Then
                    If oFac.EntityDef.ProdFactor > 400 Then
                        If oFac.ParentColony Is Nothing = False Then
                            Dim lParentID As Int32 = oFac.ParentColony.ObjectID
                            Dim bFound As Boolean = False
                            For Y As Int32 = 0 To lEnvirUB
                                If lEnvirs(Y) = lParentID Then
                                    lEnvirCnt(Y) += 1
                                    Exit For
                                End If
                            Next Y
                            If bFound = False Then
                                lEnvirUB += 1
                                lEnvirs(lEnvirUB) = lParentID
                                lEnvirCnt(lEnvirUB) += 1
                            End If
                        End If
                    End If
                End If
            End If
        Next X

        Dim lCnt As Int32 = 0
        For X As Int32 = 0 To lEnvirUB
            If lEnvirs(X) > 0 Then
                If lEnvirCnt(X) > 1 Then
                    lCnt += 1
                End If
            End If
        Next X
        Return lCnt > 1
    End Function

    Public Sub RemoveResearcher(ByVal lFacilityID As Int32)
        If mlPrimaryIdx <> -1 AndAlso glFacilityIdx(mlPrimaryIdx) = lFacilityID AndAlso mlPrimaryID = lFacilityID Then
            'ok, the primary is being removed.... do we have assisters?
            For X As Int32 = 0 To mlAssisterUB
                If mlAssisters(X) <> -1 Then
                    'Ok, you'll do...
                    Dim oFac As Facility = goFacility(mlAssisters(X))
                    If oFac Is Nothing = False Then
                        Dim oPrimFac As Facility = goFacility(mlPrimaryIdx)
                        If oPrimFac Is Nothing = False Then oFac.CurrentProduction = oPrimFac.CurrentProduction
                        mlPrimaryIdx = mlAssisters(X)
                        mlPrimaryID = mlAssisterID(X)
                        mlAssisters(X) = -1
                        mlAssisterID(X) = -1
                        oFac.RecalcProduction()
                        Return
                    End If
                End If
            Next X

            'if we are here then there are no assisters... so we'll drop dead fred
            mlPrimaryIdx = -1
            mlAssisterUB = -1
            Erase mlAssisters

            'Now, if this tech is not past the designed stage, we need to delete...
            If ComponentDevelopmentPhase < eComponentDevelopmentPhase.eComponentResearching AndAlso ObjTypeID <> ObjectType.eSpecialTech Then
                If CanBeDeleted() = True Then
                    Owner.RemoveTech(ObjectID, ObjTypeID)
                    Dim yData(7) As Byte
                    Dim lPos As Int32 = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eDeleteDesign).CopyTo(yData, lPos) : lPos += 2
                    Me.GetGUIDAsString.CopyTo(yData, lPos) : lPos += 6

                    DeleteMe()
                    Owner.SendPlayerMessage(yData, False, AliasingRights.eViewResearch)
                End If
            End If
        Else
            'just an assister
            For X As Int32 = 0 To mlAssisterUB
                If mlAssisters(X) <> -1 AndAlso glFacilityIdx(mlAssisters(X)) = lFacilityID AndAlso mlAssisterID(X) = lFacilityID Then
                    mlAssisters(X) = -1
                    mlAssisterID(X) = -1
                    RecalcPrimarysProdFactor()
                    Return
                End If
            Next X
        End If

        If Me.Owner.lConnectedPrimaryID > -1 OrElse Me.Owner.HasOnlineAliases(AliasingRights.eViewResearch) = True Then
            Dim lCnt As Int32 = GetResearcherCount()
            Dim yCnt As Byte = CByte(lCnt)

            If lCnt > 255 Then lCnt = 255
            If lCnt < 0 Then lCnt = 0

            Dim yMsg(8) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eUpdateResearcherCnt).CopyTo(yMsg, 0)
            Me.GetGUIDAsString.CopyTo(yMsg, 2)
            yMsg(8) = yCnt
            Owner.SendPlayerMessage(yMsg, False, AliasingRights.eViewResearch)
        End If

    End Sub

    Public Function GetPrimaryResearcherProdFactor() As Int32
        Dim lProdFactor As Int32 = 0

        If mlPrimaryIdx = -1 OrElse glFacilityIdx(mlPrimaryIdx) = -1 OrElse glFacilityIdx(mlPrimaryIdx) <> mlPrimaryID Then Return 0
        Dim oFac As Facility = goFacility(mlPrimaryIdx)
        If oFac Is Nothing Then Return 0
        lProdFactor = oFac.mlProdPoints

        'Now, here we go
        Dim lCnt As Int32 = 1
        Dim fMult As Single = Me.Owner.oSpecials.myResearchFacAssistBonus * 0.01F

        For X As Int32 = 0 To mlAssisterUB
            If mlAssisters(X) <> -1 AndAlso glFacilityIdx(mlAssisters(X)) <> -1 AndAlso glFacilityIdx(mlAssisters(X)) = mlAssisterID(X) Then
                oFac = goFacility(mlAssisters(X))
                If oFac Is Nothing = False AndAlso oFac.Active = True Then
                    lProdFactor += CInt(Math.Floor((oFac.mlProdPoints * (fMult / lCnt)))) 'lProdFactor += (oFac.mlProdPoints \ lCnt)
                    lCnt += 1
                End If
            End If
        Next X

        If Me.Owner.yPlayerPhase <> eyPlayerPhase.eFullLivePhase Then
            Return CInt(lProdFactor * 100)
        End If
        Return CInt(lProdFactor * gfResTimeMult)
    End Function

    Public Function IsPrimaryResearcher(ByVal lFacilityID As Int32) As Boolean
        If mlPrimaryIdx = -1 Then Return False
        Return glFacilityIdx(mlPrimaryIdx) = lFacilityID
    End Function

    Public Sub RecalcPrimarysProdFactor()
        'Ok, this is caused by an assister's prod factor changing...
        Dim bPrimaryGood As Boolean = False
        If mlPrimaryIdx <> -1 AndAlso glFacilityIdx(mlPrimaryIdx) <> -1 AndAlso glFacilityIdx(mlPrimaryIdx) = mlPrimaryID Then
            Dim oFac As Facility = goFacility(mlPrimaryIdx)
            If oFac Is Nothing = False Then
                oFac.RecalcProduction()
                If oFac.bProducing = True AndAlso oFac.CurrentProduction Is Nothing = False Then bPrimaryGood = True
            End If
        End If
        If bPrimaryGood = False Then
            'ok, let's reassign the primary
            If mlPrimaryIdx <> -1 AndAlso glFacilityIdx(mlPrimaryIdx) <> -1 Then
                RemoveResearcher(glFacilityIdx(mlPrimaryIdx))
            Else
                'ok, need to make an assister the primary
                For X As Int32 = 0 To mlAssisterUB
                    If mlAssisters(X) <> -1 Then
                        'Ok, you'll do...
                        Dim oFac As Facility = goFacility(mlAssisters(X))
                        If oFac Is Nothing = False AndAlso oFac.CurrentProduction Is Nothing = False Then
                            mlPrimaryIdx = mlAssisters(X)
                            mlPrimaryID = mlAssisterID(X)
                            mlAssisters(X) = -1
                            mlAssisterID(X) = -1
                            oFac.RecalcProduction()
                            Exit For
                        End If
                    End If
                Next X
            End If

        End If
    End Sub

    Public Function GetPrimarysPointsProduced() As Int64
        If mlPrimaryIdx <> -1 AndAlso glFacilityIdx(mlPrimaryIdx) <> -1 AndAlso glFacilityIdx(mlPrimaryIdx) = mlPrimaryID Then
            Dim oFac As Facility = goFacility(mlPrimaryIdx)
            If oFac Is Nothing = False Then
                If oFac.CurrentProduction Is Nothing = False Then Return oFac.CurrentProduction.PointsProduced
            End If
        End If
        Return 0
    End Function

    Public Sub FillPrimarysProductionStatus(ByRef blPointsProd As Int64, ByRef lLastUpdate As Int32, ByRef lProd As Int32, ByRef lFinish As Int32)
        If mlPrimaryIdx <> -1 AndAlso glFacilityIdx(mlPrimaryIdx) <> -1 AndAlso glFacilityIdx(mlPrimaryIdx) = mlPrimaryID Then
            Dim oFac As Facility = goFacility(mlPrimaryIdx)
            If oFac Is Nothing = False AndAlso oFac.CurrentProduction Is Nothing = False Then
                With oFac.CurrentProduction
                    blPointsProd = .PointsProduced
                    lLastUpdate = .lLastUpdateCycle
                    lProd = oFac.mlProdPoints
                    lFinish = .lFinishCycle
                End With
            End If
        End If
    End Sub

    ''' <summary>
    ''' Gets the number of entities currently assigned to this research
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetResearcherCount() As Int32
        Dim lCnt As Int32 = 0
        If mlPrimaryIdx = -1 OrElse glFacilityIdx(mlPrimaryIdx) = -1 OrElse glFacilityIdx(mlPrimaryIdx) <> mlPrimaryID OrElse goFacility(mlPrimaryIdx).Active = False Then
            Return 0
        End If

        lCnt += 1
        For X As Int32 = 0 To mlAssisterUB
            If mlAssisters(X) <> -1 AndAlso glFacilityIdx(mlAssisters(X)) <> -1 AndAlso glFacilityIdx(mlAssisters(X)) = mlAssisterID(X) Then
                Dim oFac As Facility = goFacility(mlAssisters(X))
                If oFac Is Nothing = False AndAlso oFac.Active = True Then lCnt += 1
            End If
        Next X
        Return lCnt
    End Function

    Public Sub ReducePrimarysPointsProduced(ByVal fMult As Single)
        If mlPrimaryIdx <> -1 AndAlso glFacilityIdx(mlPrimaryIdx) <> -1 AndAlso glFacilityIdx(mlPrimaryIdx) = mlPrimaryID Then
            Dim oFac As Facility = goFacility(mlPrimaryIdx)
            If oFac Is Nothing = False Then
                If oFac.CurrentProduction Is Nothing = False Then
                    'Return oFac.CurrentProduction.PointsProduced
                    oFac.CurrentProduction.PointsProduced = CLng(oFac.CurrentProduction.PointsProduced * fMult)
                End If
            End If
        End If
    End Sub

    ''' <summary>
    ''' Takes an amount of claimed points to be applied to this research and adjusts the research's progress. It then returns any amount of claim points remaining
    ''' </summary>
    ''' <param name="blClaimPoints">The amount of claim points available from the claimable</param>
    ''' <returns>The amount of claim points remaining after applying the blClaimPoints to the research progress</returns>
    ''' <remarks></remarks>
    Public Function AdjustProductionFromClaim(ByVal blClaimPoints As Int64) As Int64
        Dim blRemains As Int64 = blClaimPoints

        'Force a recalc now
        RecalcPrimarysProdFactor()

        'Now, get the primary
        If mlPrimaryIdx = -1 OrElse glFacilityIdx(mlPrimaryIdx) = -1 OrElse glFacilityIdx(mlPrimaryIdx) <> mlPrimaryID Then Return blRemains
        Dim oFac As Facility = goFacility(mlPrimaryIdx)
        If oFac Is Nothing Then Return blRemains

        'Now, determine the primary's production
        If oFac.CurrentProduction Is Nothing = False Then
            oFac.RecalcProduction()
            oFac.RecalcProduction()

            Dim blApplyable As Int64 = oFac.CurrentProduction.PointsRequired - oFac.CurrentProduction.PointsProduced
            Dim blApplied As Int64 = Math.Min(blClaimPoints, blApplyable)

            LogEvent(LogEventType.Informational, Me.Owner.ObjectID & " applying claimable tech refund to " & Me.ObjectID & ", " & Me.ObjTypeID & ". Required: " & oFac.CurrentProduction.PointsRequired & ". Produced: " & oFac.CurrentProduction.PointsProduced & ". Claimable: " & blClaimPoints.ToString & ", Applyable: " & blApplyable & ", Applied: " & blApplied)

            oFac.CurrentProduction.PointsProduced += blApplied
            blRemains -= blApplied

            RecalcPrimarysProdFactor()
        End If

        Return blRemains
    End Function
#End Region

#Region "  High/Low Lookups from Spreadsheets  "
    Private Shared HighLookup() As LookupEntry
    Private Shared LowLookup() As LookupEntry
    Private Shared bLookupsInitialized As Boolean = False

    Private Structure LookupEntry
        Public lLookupIndex As Int32
        Public fValues() As Single
    End Structure

    Private Shared Sub InitializeLookups()
        If bLookupsInitialized = True Then Return

        If HighLookup Is Nothing OrElse HighLookup.GetUpperBound(0) = -1 OrElse LowLookup Is Nothing OrElse LowLookup.GetUpperBound(0) = -1 Then
            ReDim HighLookup(25)
            ReDim LowLookup(25)

            For X As Int32 = 0 To 25
                With HighLookup(X)
                    .lLookupIndex = X
                    ReDim .fValues(4)

                    'Squared
                    .fValues(0) = (25 - X) : .fValues(0) *= .fValues(0)

                    'Fib
                    If X = 25 Then
                        .fValues(1) = 0
                    ElseIf X = 24 OrElse X = 23 Then
                        .fValues(1) = 1
                    Else
                        Dim fPrev As Single = 1.0F
                        .fValues(1) = 1.0F

                        For Y As Int32 = 22 To X Step -1
                            Dim fTemp As Single = .fValues(1)
                            .fValues(1) += fPrev
                            fPrev = fTemp
                        Next Y
                    End If

                    'fibXBase
                    .fValues(2) = .fValues(1) * (25I - X)

                    'fibXSquare
                    .fValues(3) = .fValues(0) * .fValues(1)

                    'fib/square
                    If X = 25 Then
                        .fValues(4) = 0
                    Else : .fValues(4) = .fValues(1) / .fValues(0)
                    End If
                End With

                With LowLookup(X)
                    .lLookupIndex = X
                    ReDim .fValues(4)

                    'squared
                    .fValues(0) = X * X
                    'fib
                    If X = 0 Then
                        .fValues(1) = 0
                    ElseIf X = 1 OrElse X = 2 Then
                        .fValues(1) = 1
                    Else
                        Dim fPrev As Single = 1.0F
                        .fValues(1) = 1.0F
                        For Y As Int32 = 3 To X
                            Dim fTemp As Single = .fValues(1)
                            .fValues(1) += fPrev
                            fPrev = fTemp
                        Next Y
                    End If
                    'fibXBase
                    .fValues(2) = X * .fValues(1)
                    'fibXSquare
                    .fValues(3) = .fValues(0) * .fValues(1)
                    'fib/square
                    If X = 0 Then
                        .fValues(4) = 0
                    Else : .fValues(4) = .fValues(1) / .fValues(0)
                    End If
                End With
            Next X
        End If
        bLookupsInitialized = True
    End Sub

    Protected Shared Function GetHighLookup(ByVal lValue As Int32, ByVal lValueIdx As Int32) As Single
        InitializeLookups()
        If lValue < 0 Then lValue = 0
        If lValue > 25 Then lValue = 25

        Return HighLookup(lValue).fValues(lValueIdx)
    End Function

    Protected Shared Function GetLowLookup(ByVal lValue As Int32, ByVal lValueIdx As Int32) As Single
        InitializeLookups()
        If lValue < 0 Then lValue = 0
        If lValue > 25 Then lValue = 25

        Return LowLookup(lValue).fValues(lValueIdx)
    End Function

#End Region

    Protected Structure MaterialPropertyItem
        Public lMineralID As Int32
        Public lPropertyID As Int32

        Public lActualValue As Int32
        Public lKnownValue As Int32
        Public fMinAmt As Single
        Public fSuccess As Single
        Public fPower As Single
        Public fHull As Single
        Public lEnlisted As Int32
        Public lOfficer As Int32
        Public fResearch As Single
        Public fProduction As Single

        Public ReadOnly Property AKScore() As Single
            Get
                Return CSng((lActualValue + 1) / (lKnownValue + 1))
            End Get
        End Property
        Public ReadOnly Property KAScore() As Single
            Get
                Return CSng((lKnownValue + 1) / (lActualValue + 1))
            End Get
        End Property
    End Structure

    Protected Structure MaterialPropertyItem2
        Public lMineralID As Int32
        Public lPropertyID As Int32

        Public lGoalValue As Int32
        Public lActualValue As Int32
        Public lKnownValue As Int32

        Public fNormalize As Single

        Public ReadOnly Property FinalScore() As Single
            Get
                Return fNormalize * Difference
            End Get
        End Property
        Public ReadOnly Property Difference() As Single
            Get
                '=POWER(ABS(C9-B9)+1,IF(E9=1,1,1+E9))
                Dim fResult As Single = Math.Abs(lActualValue - lGoalValue) + 1
                If lActualValue = lKnownValue Then
                    Return fResult
                Else : Return CSng(Math.Pow(fResult, 1 + AKScore))
                End If
            End Get
        End Property
        Public ReadOnly Property AKScore() As Single
            Get
                Return CSng((lActualValue + 1) / (lKnownValue + 1))
            End Get
        End Property
        Public ReadOnly Property KAScore() As Single
            Get
                Return CSng((lKnownValue + 1) / (lActualValue + 1))
            End Get
        End Property
    End Structure

    Public Sub SetTechCost(ByRef oPC As ProductionCost)
        If oPC.ProductionCostType = 0 Then
            moProductionCost = oPC
        Else
            moResearchCost = oPC
        End If
    End Sub
    'Public Sub SetTechCost(ByVal lPC_ID As Int32, ByVal lObjID As Int32, ByVal iObjTypeID As Short, ByVal blCredits As Int64, ByVal lColonists As Int32, ByVal lEnlisted As Int32, ByVal lOfficers As Int32, ByVal blPoints As Int64, ByVal yType As Byte)
    '    If yType = 0 Then       'productioncost
    '        If moProductionCost Is Nothing Then moProductionCost = New ProductionCost()
    '        With moProductionCost
    '            .ColonistCost = lColonists
    '            .CreditCost = blCredits
    '            .EnlistedCost = lEnlisted
    '            .ItemCostUB = -1
    '            Erase .ItemCosts
    '            .ObjectID = lObjID
    '            .ObjTypeID = iObjTypeID
    '            .OfficerCost = lOfficers
    '            .PC_ID = lPC_ID
    '            .PointsRequired = blPoints
    '            .ProductionCostType = yType
    '        End With
    '    Else                    'researchcost
    '        If moResearchCost Is Nothing Then moResearchCost = New ProductionCost()
    '        With moResearchCost
    '            .ColonistCost = lColonists
    '            .CreditCost = blCredits
    '            .EnlistedCost = lEnlisted
    '            .ItemCostUB = -1
    '            Erase .ItemCosts
    '            .ObjectID = lObjID
    '            .ObjTypeID = iObjTypeID
    '            .OfficerCost = lOfficers
    '            .PC_ID = lPC_ID
    '            .PointsRequired = blPoints
    '            .ProductionCostType = yType
    '        End With
    '    End If
    'End Sub

    'Public Sub FinalizeTechCostsLoad()
    '    Dim oComm As OleDb.OleDbCommand
    '    Dim oData As OleDb.OleDbDataReader

    '    Dim sSQL As String = "SELECT pc.ProductionCostType, pci.ItemID, pci.ItemTypeID, pci.Quantity, pci.PCM_ID FROM tblProductionCostItem pci LEFT OUTER JOIN " & _
    '      " tblProductionCost pc ON pci.PC_ID = pc.PC_ID WHERE pc.ObjectID = " & Me.ObjectID & " AND pc.ObjTypeID = " & Me.ObjTypeID
    '    oComm = New OleDb.OleDbCommand(sSQL, goCN)
    '    oData = oComm.ExecuteReader()

    '    While oData.Read()
    '        Dim yType As Byte = CByte(oData("ProductionCostType"))
    '        If yType = 0 Then
    '            If moProductionCost Is Nothing = False Then

    '                moProductionCost.ItemCostUB += 1
    '                ReDim Preserve moProductionCost.ItemCosts(moProductionCost.ItemCostUB)
    '                moProductionCost.ItemCosts(moProductionCost.ItemCostUB) = New ProductionCostItem
    '                With moProductionCost.ItemCosts(moProductionCost.ItemCostUB)
    '                    .ItemID = CInt(oData("ItemID"))
    '                    .ItemTypeID = CShort(oData("ItemTypeID"))
    '                    .oProdCost = moProductionCost
    '                    .PC_ID = moProductionCost.PC_ID
    '                    .PCM_ID = CInt(oData("PCM_ID"))
    '                    .QuantityNeeded = CInt(oData("Quantity"))
    '                End With
    '            End If
    '        Else
    '            If moResearchCost Is Nothing = False Then
    '                moResearchCost.ItemCostUB += 1
    '                ReDim Preserve moResearchCost.ItemCosts(moResearchCost.ItemCostUB)
    '                moResearchCost.ItemCosts(moResearchCost.ItemCostUB) = New ProductionCostItem
    '                With moResearchCost.ItemCosts(moResearchCost.ItemCostUB)
    '                    .ItemID = CInt(oData("ItemID"))
    '                    .ItemTypeID = CShort(oData("ItemTypeID"))
    '                    .oProdCost = moResearchCost
    '                    .PC_ID = moResearchCost.PC_ID
    '                    .PCM_ID = CInt(oData("PCM_ID"))
    '                    .QuantityNeeded = CInt(oData("Quantity"))
    '                End With
    '            End If
    '        End If
    '    End While

    '    oData.Close()
    '    oData = Nothing
    '    oComm = Nothing
    'End Sub

    Public Function CanBeDeleted() As Boolean
        If Me.ObjTypeID = ObjectType.eSpecialTech Then Return False
        If Me.ComponentDevelopmentPhase < eComponentDevelopmentPhase.eResearched Then Return True Else Return False
    End Function

    Public Function DeleteMe() As Boolean
        Dim sSQL As String = ""
        Dim sTable As String
        Dim sField As String
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim bResult As Boolean = False

        If mlPrimaryIdx <> -1 AndAlso glFacilityIdx(mlPrimaryIdx) <> -1 AndAlso glFacilityIdx(mlPrimaryIdx) = mlPrimaryID Then
            Dim oFac As Facility = goFacility(mlPrimaryIdx)
            If oFac Is Nothing = False Then oFac.CurrentProduction = Nothing
        End If

        For X As Int32 = 0 To mlAssisterUB
            If mlAssisters(X) <> -1 AndAlso glFacilityIdx(mlAssisters(X)) <> -1 AndAlso glFacilityIdx(mlAssisters(X)) = mlAssisterID(X) Then
                Dim oFac As Facility = goFacility(mlAssisters(X))
                If oFac Is Nothing Then oFac.CurrentProduction = Nothing
            End If
        Next X

        Try
            Select Case Me.ObjTypeID
                Case ObjectType.eAlloyTech
                    sTable = "tblAlloy"
                    sField = "AlloyID"
                    sSQL = "DELETE FROM tblAlloyResultProperty WHERE AlloyID = " & Me.ObjectID
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oComm.ExecuteNonQuery()
                    oComm.Dispose()
                    oComm = Nothing
                Case ObjectType.eArmorTech
                    sTable = "tblArmor"
                    sField = "ArmorID"
                Case ObjectType.eEngineTech
                    sTable = "tblEngine"
                    sField = "EngineID"
                Case ObjectType.eHullTech
                    sTable = "tblHull"
                    sField = "HullID"
                    sSQL = "DELETE FROM tblHardpoint WHERE HullID = " & Me.ObjectID
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    oComm.ExecuteNonQuery()
                    oComm.Dispose()
                    oComm = Nothing
                Case ObjectType.ePrototype
                    sTable = "tblPrototype"
                    sField = "PrototypeID"
                Case ObjectType.eRadarTech
                    sTable = "tblRadar"
                    sField = "RadarID"
                Case ObjectType.eShieldTech
                    sTable = "tblShield"
                    sField = "ShieldID"
                Case ObjectType.eWeaponTech
                    sTable = "tblWeapon"
                    sField = "WeaponID"
                Case Else
                    Return False
            End Select
            sSQL = "DELETE FROM " & sTable & " WHERE " & sField & " = " & Me.ObjectID

            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing

            bResult = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Could not delete " & Me.ObjectID & " type " & Me.ObjTypeID & ": " & ex.Message)
        Finally
            If oComm Is Nothing = False Then oComm.Dispose()
            oComm = Nothing
        End Try
        'bNeedsSaved = False
        Return bResult
    End Function

    Public Function GetTechName() As Byte()
        Select Case Me.ObjTypeID
            Case ObjectType.eAlloyTech
                Return CType(Me, AlloyTech).AlloyName
            Case ObjectType.eArmorTech
                Return CType(Me, ArmorTech).ArmorName
            Case ObjectType.eEngineTech
                Return CType(Me, EngineTech).EngineName
            Case ObjectType.eHullTech
                Return CType(Me, HullTech).HullName
            Case ObjectType.ePrototype
                Return CType(Me, Prototype).PrototypeName
            Case ObjectType.eRadarTech
                Return CType(Me, RadarTech).RadarName
            Case ObjectType.eShieldTech
                Return CType(Me, ShieldTech).ShieldName
            Case ObjectType.eSpecialTech
                Return CType(Me, PlayerSpecialTech).oTech.TechName
            Case ObjectType.eWeaponTech
                Return CType(Me, BaseWeaponTech).WeaponName
            Case Else
                Return StringToBytes("Unknown")
        End Select
    End Function

    Public Sub SetTechName(ByVal sNewVal As String)
        Select Case Me.ObjTypeID
            Case ObjectType.eAlloyTech
                ReDim CType(Me, AlloyTech).AlloyName(19)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewVal).CopyTo(CType(Me, AlloyTech).AlloyName, 0)
            Case ObjectType.eArmorTech
                ReDim CType(Me, ArmorTech).ArmorName(19)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewVal).CopyTo(CType(Me, ArmorTech).ArmorName, 0)
            Case ObjectType.eEngineTech
                ReDim CType(Me, EngineTech).EngineName(19)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewVal).CopyTo(CType(Me, EngineTech).EngineName, 0)
            Case ObjectType.eHullTech
                ReDim CType(Me, HullTech).HullName(19)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewVal).CopyTo(CType(Me, HullTech).HullName, 0)
            Case ObjectType.ePrototype
                ReDim CType(Me, Prototype).PrototypeName(19)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewVal).CopyTo(CType(Me, Prototype).PrototypeName, 0)
            Case ObjectType.eRadarTech
                ReDim CType(Me, RadarTech).RadarName(19)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewVal).CopyTo(CType(Me, RadarTech).RadarName, 0)
            Case ObjectType.eShieldTech
                ReDim CType(Me, ShieldTech).ShieldName(19)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewVal).CopyTo(CType(Me, ShieldTech).ShieldName, 0)
            Case ObjectType.eWeaponTech
                ReDim CType(Me, BaseWeaponTech).WeaponName(19)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewVal).CopyTo(CType(Me, BaseWeaponTech).WeaponName, 0)
        End Select
    End Sub
End Class
