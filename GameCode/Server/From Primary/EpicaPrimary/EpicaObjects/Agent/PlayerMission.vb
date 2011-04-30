Option Strict On

Public Enum ePM_ErrorCode As Int32
	eUnknownError = Int32.MinValue
	eNoMissionSet = -2
	eNoErrorsFound = 1
	eInvalidMethod = -4
	eMissingSkillset = -5
	eMissingAssignment = -6
	eMissingAssignmentAgent = -7
	eAssignedAgentUnavail = -8
	eInvalidTargetSelection = -9
End Enum

Public Class PlayerMission
    Public PM_ID As Int32               'unique identifier

    Public oMission As Mission
    Public CasualtyCnt As Int32 = 0

    Public oPlayer As Player                'player who started this mission (agents belong to this player)
    Public oTarget As Player                'player who this mission is against

    'Specific item GUID... for example, a specific factory, unit, or something could be selected to be destroyed...
    '  a specific agent could be selected to be rescued.... etc...
    Public lTargetID As Int32
    Public iTargetTypeID As Int16
    Public lTargetID2 As Int32
    Public iTargetTypeID2 As Int16

    Public Modifier As Int32
    Public bAlarmThrown As Boolean = False

    Public oMissionGoals() As PlayerMissionGoal
    Public ySafeHouseSetting As Byte = 0        '0 = No Safe House, Non-zero is a bonus to rolls that benefit from safehouses
    Public oSafeHouseGoal As PlayerMissionGoal = Nothing

    Public oPhases() As PlayerMissionPhase
	Public lPhaseUB As Int32 = -1

	Public lMethodID As Int32 = -1

	Public bIsSuccess As Boolean = False		'indicates current setting, CurrentPhase always represents the correct value, but when all
	'goals are executed, this value is what determines the missions final result

    'Public bArchived As Boolean = False
    Public yArchived As Byte = 0

    Public lPreviousAsyncPhase As eMissionPhase = eMissionPhase.eInPlanning

    Public lCurrentPhase As eMissionPhase = eMissionPhase.eInPlanning

    Public bAuditAlertSent As Boolean = False

    Public Sub ProcessTests()
        If lCurrentPhase = eMissionPhase.eReinfiltrationPhase Then
            'PER AGENT USED IN MISSION THAT IS NOT CAPTURED OR DEAD
            '   If the alarm is set and counter agents are in the area, capture tests are required (General Counter Agents Excluded)

            If Me.oTarget Is Nothing OrElse oTarget.ObjectID <> Me.oPlayer.ObjectID Then

                Dim lAlarmMult As Int32 = 1
                Dim bCaptureTest As Boolean = False
                If bAlarmThrown = True Then
                    '       - captures during this period are temporary
                    '       - cover agents can intercept
                    bCaptureTest = True
                    '       AlarmMultiplier = 5
                    lAlarmMult = 2
                End If

                Dim lAliveCounterAgents As Int32 = 0
                If oTarget Is Nothing = False AndAlso oTarget.oSecurity Is Nothing = False Then
                    lAliveCounterAgents = oTarget.oSecurity.GetCounterAgentCount(Me.oMission.lInfiltrationType)
                End If
                If lAliveCounterAgents = 0 Then bCaptureTest = False

                For X As Int32 = 0 To oMission.GoalUB
                    If oMission.MethodIDs(X) = lMethodID AndAlso oMissionGoals(X).GoalExecuted = 0 Then
                        With oMissionGoals(X)
                            For Y As Int32 = 0 To .lAssignmentUB
                                If .oAssignments(Y) Is Nothing = False AndAlso .oAssignments(Y).oAgent Is Nothing = False Then
                                    If (.oAssignments(Y).oAgent.lAgentStatus And (AgentStatus.IsDead Or AgentStatus.HasBeenCaptured)) = 0 Then
                                        If bCaptureTest = True Then
                                            'ok, gotta run a capture test
                                            ReinfiltrationCaptureTest(.oAssignments(Y).oAgent)
                                        End If

                                        If .oAssignments(Y).oAgent Is Nothing Then Continue For
                                        '   Adjust suspicion based on this equation:
                                        '       (1 + Alive_Counter_Agent_Count_In_Area) * (1 + (Agent_Suspicion / 100)) * AlarmMultiplier
                                        Dim lVal As Int32 = .oAssignments(Y).oAgent.Suspicion
                                        lVal += CInt((1 + lAliveCounterAgents) * (1 + (lVal / 100.0F)) * lAlarmMult)
                                        If lVal > 255 Then lVal = 255
                                        If lVal < 0 Then lVal = 0
                                        .oAssignments(Y).oAgent.Suspicion = CByte(lVal)
                                    End If
                                End If
                            Next Y
                        End With
                    End If
                Next X
            End If

            'Now, test for phase change to throw us into mission complete status
            TestForPhaseChange()
        Else

            'Ok, first, the safehouse setting
            If Me.ySafeHouseSetting > 0 Then
                If oSafeHouseGoal.GoalExecuted = 0 Then
                    'ok, try it...
                    With oSafeHouseGoal
                        If .lLastProcessTest = Int32.MinValue Then
                            .lLastProcessTest = glCurrentCycle
                            Return
                        ElseIf glCurrentCycle - .lLastProcessTest < .oGoal.BaseTime Then
                            Return
                        End If

                        For Y As Int32 = 0 To .lAssignmentUB
                            If .oAssignments(Y).Completed = False Then
                                If (.oAssignments(Y).yStatus And AgentAssignmentStatus.Skipped) = 0 Then
                                    .oAssignments(Y).ProcessAgentAssignment()
                                Else
                                    .GoalExecuted = 1
                                    .CheckAssignmentsCompleted()
                                End If
                            End If
                        Next Y

                        If .GoalSuccess = 1 Then
                            .GoalExecuted = 1
                            .oGoal.HandleGoalCtrlID(Me, .oGoal.SuccessProgCtrlID, 0)
                        End If
                    End With
                    Return
                End If
            End If

            Dim bFound As Boolean = False
            For X As Int32 = 0 To oMission.GoalUB
                If oMission.MethodIDs(X) = lMethodID AndAlso oMission.Goals(X).MissionPhase = lCurrentPhase AndAlso oMissionGoals(X).GoalExecuted = 0 Then

                    If oMission.Goals(X).OrderNum > 0 Then
                        Dim lOrderNum As Int32 = oMission.Goals(X).OrderNum
                        Dim bBad As Boolean = False
                        For Y As Int32 = 0 To oMission.GoalUB
                            If oMission.MethodIDs(Y) = lMethodID AndAlso oMission.Goals(Y).MissionPhase = lCurrentPhase AndAlso oMission.Goals(Y).OrderNum < lOrderNum Then
                                If oMissionGoals(Y).GoalExecuted = 0 Then
                                    bBad = True
                                    Exit For
                                End If
                            End If
                        Next Y
                        If bBad = True Then Continue For
                    End If

                    bFound = True

                    With oMissionGoals(X)

                        If .lLastProcessTest = Int32.MinValue Then
                            .lLastProcessTest = glCurrentCycle
                            Continue For
                        ElseIf glCurrentCycle - .lLastProcessTest < .oGoal.BaseTime AndAlso gb_IS_TEST_SERVER = False Then
                            Continue For
                        End If

                        For Y As Int32 = 0 To .lAssignmentUB
                            If .oAssignments(Y).Completed = False Then
                                If (.oAssignments(Y).yStatus And AgentAssignmentStatus.Skipped) = 0 Then
                                    .oAssignments(Y).ProcessAgentAssignment()
                                Else
                                    .GoalExecuted = 1
                                    .CheckAssignmentsCompleted()
                                End If
                            End If
                        Next Y
                        'if our current phase > preparation then forcefully check goal success
                        If lCurrentPhase > eMissionPhase.ePreparationTime Then
                            '.GoalExecuted = 1
                            .CheckAssignmentsCompleted() 
                            If .GoalExecuted = 1 Then
                                If .GoalSuccess = 1 Then .oGoal.HandleGoalCtrlID(Me, .oGoal.SuccessProgCtrlID, X) Else .oGoal.HandleGoalCtrlID(Me, .oGoal.FailureProgCtrlID, X)
                            End If
                            TestForPhaseChange()
                        Else
                            'Ok, see if we need to move on to the next phase
                            If .GoalExecuted = 1 Then       'was goalsuccess
                                TestForPhaseChange()
                                Exit For            'this is new 06/14/08
                            End If
                        End If
                    End With
                End If
            Next X
            If bFound = False Then TestForPhaseChange()
        End If
    End Sub

    Private Sub TestForPhaseChange()
        Dim bResult As Boolean = True
        For X As Int32 = 0 To oMission.GoalUB
			If oMission.MethodIDs(X) = lMethodID AndAlso oMissionGoals(X).oGoal.MissionPhase = lCurrentPhase AndAlso oMissionGoals(X).GoalSuccess = 0 AndAlso oMissionGoals(X).GoalExecuted = 0 Then
				bResult = False
				Exit For
			End If
        Next X
        If bResult = True Then
            Select Case lCurrentPhase
                Case eMissionPhase.eFlippingTheSwitch
                    lCurrentPhase = eMissionPhase.eReinfiltrationPhase
                Case eMissionPhase.ePreparationTime
                    lCurrentPhase = eMissionPhase.eSettingTheStage
                    Me.Modifier += Me.oMission.Modifier
                    For X As Int32 = 0 To oMission.GoalUB
                        If oMission.MethodIDs(X) = lMethodID AndAlso oMissionGoals(X).oGoal.MissionPhase = eMissionPhase.ePreparationTime Then
                            If Me.oMissionGoals(X).GoalSuccess = 0 Then
                                'ok, determine how many items were skipped...
                                Dim lCnt As Int32 = 0
                                Dim fPreparedness As Single = 0.0F
                                For Y As Int32 = 0 To Me.oMissionGoals(X).lAssignmentUB
                                    If Me.oMissionGoals(X).oAssignments(Y) Is Nothing = False Then
                                        Dim fTemp As Single = CSng(Me.oMissionGoals(X).oAssignments(Y).PointsAccumulated / Math.Max(1, Me.oMissionGoals(X).oAssignments(Y).PointsRequired))
                                        fPreparedness += fTemp
                                        lCnt += 1
                                    End If
                                Next Y
                                If lCnt = 0 Then fPreparedness = 1 Else fPreparedness /= Math.Max(1, lCnt)
                                fPreparedness *= 50
                                Dim lFinalMod As Int32 = CInt(Math.Abs(fPreparedness))
                                Me.Modifier += -lFinalMod
                            End If
                        End If
                    Next X
				Case eMissionPhase.eReinfiltrationPhase
					If bIsSuccess = True Then
						lCurrentPhase = eMissionPhase.eMissionOverSuccess
					Else : lCurrentPhase = eMissionPhase.eMissionOverFailure
					End If
					MarkAgentsOffMission()
                Case eMissionPhase.eSettingTheStage
                    lCurrentPhase = eMissionPhase.eFlippingTheSwitch
			End Select

            If Me.oPlayer.lConnectedPrimaryID > -1 OrElse Me.oPlayer.HasOnlineAliases(AliasingRights.eViewAgents) = True Then
                Me.oPlayer.SendPlayerMessage(Me.GetAddObjectMessage(), False, AliasingRights.eViewAgents)
            End If
        End If
    End Sub

    Public Sub New(ByRef poMission As Mission, ByRef poPlayer As Player)
        oMission = poMission
        oPlayer = poPlayer

        ReDim oMissionGoals(oMission.GoalUB)
        For X As Int32 = 0 To oMission.GoalUB
            oMissionGoals(X) = New PlayerMissionGoal()
            With oMissionGoals(X)
                .oGoal = oMission.Goals(X)
                .oMission = Me
            End With
        Next X

        lPhaseUB = eMissionPhase.eReinfiltrationPhase
        ReDim oPhases(lPhaseUB)
        For X As Int32 = 0 To lPhaseUB
            oPhases(X) = New PlayerMissionPhase()
            With oPhases(X)
                .lPhase = CType(X, eMissionPhase)
                .oParent = Me
            End With
        Next X
    End Sub

    Public Function GetAvailableCoverAgent(ByVal lPhase As eMissionPhase) As Agent
		Dim lInfTarget As Int32 = -1
		Dim iInfTargetTypeID As Int16 = -1
		Select Case CType(oMission.ProgramControlID, eMissionResult)
            Case eMissionResult.eFindMineral, eMissionResult.eReconPlanetMap
                lInfTarget = lTargetID2 : iInfTargetTypeID = iTargetTypeID2
			Case eMissionResult.eGeologicalSurvey
				lInfTarget = lTargetID : iInfTargetTypeID = iTargetTypeID
			Case eMissionResult.eTutorialFindFactory
                lInfTarget = lTargetID : iInfTargetTypeID = ObjectType.ePlanet
            Case eMissionResult.eSearchAndRescueAgent
                Dim oTmpAgent As Agent = GetEpicaAgent(Me.lTargetID)
                If oTmpAgent Is Nothing = False AndAlso (oTmpAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then
                    lInfTarget = oTmpAgent.lCapturedBy
                    iInfTargetTypeID = ObjectType.ePlayer
                End If
			Case Else
				'Player....
				If oTarget Is Nothing = False Then
					lInfTarget = oTarget.ObjectID : iInfTargetTypeID = ObjectType.ePlayer
				End If
		End Select

        With oPhases(lPhase)
            For X As Int32 = 0 To .lCoverAgentUB
				If .lCoverAgentIdx(X) <> -1 AndAlso .yUsedAsCoverAgent(X) = 0 AndAlso (.oCoverAgents(X).lAgentStatus And AgentStatus.UsedAsCoverAgent) = 0 AndAlso (.oCoverAgents(X).lAgentStatus And AgentStatus.IsInfiltrated) <> 0 Then
					If .oCoverAgents(X).lTargetID = lInfTarget AndAlso .oCoverAgents(X).iTargetTypeID = iInfTargetTypeID Then
						Return .oCoverAgents(X)
					End If
				End If
            Next X
        End With
        Return Nothing
    End Function

	Public Sub HandleMissionCompletion()

        bIsSuccess = False

        'ok, mission is completed, determine what mission result is
        Select Case CType(oMission.ProgramControlID, eMissionResult)
            Case eMissionResult.eAcquireTechData
                If oTarget Is Nothing Then Return
                If lTargetID < 1 OrElse iTargetTypeID < 1 Then
                    'ok, need to get a targetid, targettypeid...
                    oTarget.GetRandomComponentTech(lTargetID, iTargetTypeID)
                End If

                Dim oTech As Epica_Tech = oTarget.GetTech(lTargetID, iTargetTypeID)
                If oTech Is Nothing = False Then
                    Dim oPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(oPlayer, oTech, PlayerTechKnowledge.KnowledgeType.eSettingsLevel2, False)
                    If oPTK Is Nothing = False Then oPTK.SendMsgToPlayer()
                End If
            Case eMissionResult.eAgencyAnalysis
                If oTarget Is Nothing = False Then
                    If oTarget.oSecurity Is Nothing = False Then
                        Dim sBody As String = oTarget.oSecurity.GetAgencyAnalysis()
                        If sBody Is Nothing = False AndAlso sBody.Trim <> "" Then
                            Dim oPC As PlayerComm = Me.oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                                sBody, "Agency Analysis for " & oTarget.sPlayerNameProper, Me.oPlayer.ObjectID, GetDateAsNumber(Now), False, Me.oPlayer.sPlayerNameProper, Nothing)
                            If oPC Is Nothing = False Then
                                oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                            End If
                        End If
                    End If
                End If
            Case eMissionResult.eAssassinateAgent
                If oTarget Is Nothing Then Return
                If lTargetID < 1 Then lTargetID = oTarget.GetRandomAgentID(True)
                Dim oAgent As Agent = GetEpicaAgent(lTargetID)
                If oAgent Is Nothing = False Then
                    oAgent.KillMe()
                    If (oAgent.oOwner.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                        Dim oPC As PlayerComm = oAgent.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                           "Your agent, " & BytesToString(oAgent.AgentName) & ", was mysteriously killed. An investigation is underway to determine cause and suspects. At this time, it appears to be a black op from a rival empire.", _
                           "Agent Killed", oAgent.oOwner.ObjectID, GetDateAsNumber(Now), False, oAgent.oOwner.sPlayerNameProper, Nothing)
                        If oPC Is Nothing = False Then
                            oAgent.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                        End If
                    End If
                End If
            Case eMissionResult.eAssassinateGovernor
                If oTarget Is Nothing Then Return
                Dim lColonyIdx As Int32 = -1
                If lTargetID < 1 Then
                    lTargetID = oTarget.GetRandomColonyID()
                End If
                If lTargetID <> -1 Then lColonyIdx = oTarget.GetColonyFromColonyID(lTargetID)
                If lColonyIdx = -1 OrElse glColonyIdx(lColonyIdx) = -1 OrElse glColonyIdx(lColonyIdx) <> lTargetID Then
                    'Indicate target colony could not be find
                    Dim oPC As PlayerComm = Me.oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                     "Your agent team reports that the target colony could not be located.", "Mission Failure", Me.oPlayer.ObjectID, _
                     GetDateAsNumber(Now), False, Me.oPlayer.sPlayerNameProper, Nothing)
                    If oPC Is Nothing = False Then
                        Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    End If
                Else
                    Dim lDuration As Int32 = CInt(Rnd() * 10) + 5     'in hours
                    lDuration *= 108000     'convert to cycles
                    Dim oColony As Colony = goColony(lColonyIdx)
                    If oColony Is Nothing = False Then

                        If oPlayer.HasItemIntel(oColony.ObjectID, oColony.ObjTypeID, oColony.Owner.ObjectID) = False Then
                            Dim oPII As PlayerItemIntel = oPlayer.AddPlayerItemIntel(oColony.ObjectID, oColony.ObjTypeID, PlayerItemIntel.PlayerItemIntelType.eLocation, oColony.Owner.ObjectID)
                            If oPII Is Nothing = False Then
                                oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                            End If
                        End If

                        'Determine if the colony already has an agent effect of assassinate governor...
                        If oColony.HasEffectActive(AgentEffectType.eGovernorMorale) = True Then
                            Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                             "The governor of the " & BytesToString(oColony.ColonyName) & " colony was already assassinated when your team arrived.", "Governor Assassinated", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                            If oPC Is Nothing = False Then
                                oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                            End If
                        Else
                            Dim oEffect As AgentEffect = oColony.AddAgentEffect(glCurrentCycle, lDuration, AgentEffectType.eGovernorMorale, oMission.BaseEffect, False, oPlayer.ObjectID)
                            If oEffect Is Nothing Then Return
                            oPlayer.AddImposedEffect(oEffect, oTarget, BytesToString(oColony.ColonyName))

                            Dim oPC As PlayerComm = oColony.Owner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                             "The governor of the " & BytesToString(oColony.ColonyName) & " colony has been assassinated. We are sure this is the work of some terrorist group affiliated with a rival empire." & _
                             vbCrLf & vbCrLf & "Production and research efficiency is expected to fall while the colony mourns.", _
                             "Governor Assassinated", oColony.Owner.ObjectID, GetDateAsNumber(Now), False, oColony.Owner.sPlayerNameProper, Nothing)
                            If oPC Is Nothing = False Then
                                oColony.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                            End If
                        End If
                    End If
                End If
            Case eMissionResult.eAudit
                oPlayer.oSecurity.PerformAudit()
            Case eMissionResult.eBadPublicity
                If oTarget Is Nothing Then Return
                Dim lDuration As Int32 = CInt(Rnd() * 168) + 24       'in hours
                lDuration *= 108000     'convert to cycles
                Dim oEffect As AgentEffect = oTarget.oBudget.AddAgentEffect(glCurrentCycle, lDuration, AgentEffectType.eTradeIncome, oMission.BaseEffect, True, oPlayer.ObjectID)
                If oEffect Is Nothing = False Then oPlayer.AddImposedEffect(oEffect, oTarget, "Trade Income")
            Case eMissionResult.eCaptureAgent
                If oTarget Is Nothing Then Return
                If lTargetID < 1 Then lTargetID = oTarget.GetRandomAgentID(False)
                Dim oAgent As Agent = GetEpicaAgent(lTargetID)
                If oAgent Is Nothing = False AndAlso (oAgent.lAgentStatus And (AgentStatus.IsDead Or AgentStatus.Dismissed Or AgentStatus.NewRecruit)) = 0 Then
                    oPlayer.oSecurity.CaptureAgent(oAgent)
                Else
                    'Indicate agent could not be found
                    Dim oPC As PlayerComm = Me.oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                     "Your agent team reports that the target agent to be captured could not be located.", "Mission Failure", Me.oPlayer.ObjectID, _
                     GetDateAsNumber(Now), False, Me.oPlayer.sPlayerNameProper, Nothing)
                    If oPC Is Nothing = False Then
                        Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    End If
                End If
            Case eMissionResult.eClutterCargoBay
                If oTarget Is Nothing Then Return
                Dim oColony As Colony = Nothing
                Dim bGood As Boolean = False
                If lTargetID > 0 Then
                    oColony = GetEpicaColony(lTargetID)
                    If oColony Is Nothing = False Then
                        If oPlayer.HasItemIntel(oColony.ObjectID, oColony.ObjTypeID, oColony.Owner.ObjectID) = False Then
                            Dim oPII As PlayerItemIntel = oPlayer.AddPlayerItemIntel(oColony.ObjectID, oColony.ObjTypeID, PlayerItemIntel.PlayerItemIntelType.eLocation, oColony.Owner.ObjectID)
                            If oPII Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                        End If

                        Dim lDuration As Int32 = CInt(Rnd() * 168) + 24       'in hours
                        lDuration *= 108000     'convert to cycles
                        Dim oEffect As AgentEffect = oColony.AddAgentEffect(glCurrentCycle, lDuration, AgentEffectType.eCargoBay, oMission.BaseEffect, True, oPlayer.ObjectID)
                        If oEffect Is Nothing = False Then oPlayer.AddImposedEffect(oEffect, oTarget, BytesToString(oColony.ColonyName))
                        bGood = True
                    End If
                End If
                If (Me.oPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                    Dim oPC As PlayerComm = Nothing
                    If bGood = False Then
                        oPC = Me.oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                            "Your agent team reports that the target colony could not be located to clutter cargo processes. The mission ultimately is a failure.", "Clutter Cargo Mission Failure", Me.oPlayer.ObjectID, _
                            GetDateAsNumber(Now), False, Me.oPlayer.sPlayerNameProper, Nothing)
                    Else
                        oPC = Me.oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                            "Your agent team reports that the target colony was located and the cargo processes have been cluttered. The mission is a success.", "Clutter Cargo Mission Success", Me.oPlayer.ObjectID, _
                            GetDateAsNumber(Now), False, Me.oPlayer.sPlayerNameProper, Nothing)
                    End If
                    If oPC Is Nothing = False Then
                        Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    End If
                End If
            Case eMissionResult.eComponentDesignList
                If oTarget Is Nothing Then Return
                oTarget.HandleComponentListEspionage(False, oPlayer)
            Case eMissionResult.eComponentResearchList
                If oTarget Is Nothing Then Return
                oTarget.HandleComponentListEspionage(True, oPlayer)
            Case eMissionResult.eCorruptProduction
                'TODO: What to do?
            Case eMissionResult.eDecreaseHousingMorale
                If oTarget Is Nothing Then Return
                If lTargetID < 1 Then
                    lTargetID = oTarget.GetRandomColonyID()
                End If
                If lTargetID <> -1 Then
                    Dim oColony As Colony = GetEpicaColony(lTargetID)
                    If oColony Is Nothing = False Then
                        Dim oPII As PlayerItemIntel = Me.oPlayer.AddPlayerItemIntel(oColony.ObjectID, oColony.ObjTypeID, PlayerItemIntel.PlayerItemIntelType.eLocation, oColony.Owner.ObjectID)
                        If oPII Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)

                        Dim lDuration As Int32 = CInt(Rnd() * 168) + 24       'in hours
                        lDuration *= 108000     'convert to cycles
                        Dim oEffect As AgentEffect = oColony.AddAgentEffect(glCurrentCycle, lDuration, AgentEffectType.eColonyHousingMorale, oMission.BaseEffect, True, oPlayer.ObjectID)
                        If oEffect Is Nothing = False Then oPlayer.AddImposedEffect(oEffect, oTarget, BytesToString(oColony.ColonyName))
                    End If
                End If
            Case eMissionResult.eDecreaseTaxMorale
                If oTarget Is Nothing Then Return
                If lTargetID < 1 Then
                    lTargetID = oTarget.GetRandomColonyID()
                End If
                If lTargetID <> -1 Then
                    Dim oColony As Colony = GetEpicaColony(lTargetID)
                    If oColony Is Nothing = False Then
                        'Dim lColonyIdx As Int32 = oColony.Owner.GetColonyFromColonyID(lTargetID)
                        Dim oPII As PlayerItemIntel = Me.oPlayer.AddPlayerItemIntel(oColony.ObjectID, oColony.ObjTypeID, PlayerItemIntel.PlayerItemIntelType.eLocation, oColony.Owner.ObjectID)
                        If oPII Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)

                        Dim lDuration As Int32 = CInt(Rnd() * 168) + 24       'in hours
                        lDuration *= 108000     'convert to cycles
                        Dim oEffect As AgentEffect = oColony.AddAgentEffect(glCurrentCycle, lDuration, AgentEffectType.eColonyTaxMorale, oMission.BaseEffect, True, oPlayer.ObjectID)
                        If oEffect Is Nothing = False Then oPlayer.AddImposedEffect(oEffect, oTarget, BytesToString(oColony.ColonyName))
                    End If
                End If
            Case eMissionResult.eDecreaseUnemploymentMorale
                If oTarget Is Nothing Then Return
                If lTargetID < 1 Then
                    lTargetID = oTarget.GetRandomColonyID()
                End If
                If lTargetID <> -1 Then
                    Dim oColony As Colony = GetEpicaColony(lTargetID)
                    If oColony Is Nothing = False Then
                        'Dim lColonyIdx As Int32 = oColony.Owner.GetColonyFromColonyID(lTargetID)
                        Dim oPII As PlayerItemIntel = Me.oPlayer.AddPlayerItemIntel(oColony.ObjectID, oColony.ObjTypeID, PlayerItemIntel.PlayerItemIntelType.eLocation, oColony.Owner.ObjectID)
                        If oPII Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)

                        Dim lDuration As Int32 = CInt(Rnd() * 168) + 24       'in hours
                        lDuration *= 108000     'convert to cycles
                        Dim oEffect As AgentEffect = oColony.AddAgentEffect(glCurrentCycle, lDuration, AgentEffectType.eColonyUnemploymentMorale, oMission.BaseEffect, True, oPlayer.ObjectID)
                        If oEffect Is Nothing = False Then oPlayer.AddImposedEffect(oEffect, oTarget, BytesToString(oColony.ColonyName))
                    End If
                End If
            Case eMissionResult.eDegradePay
                'TODO: to be implemented
            Case eMissionResult.eDestroyFacility
                If oTarget Is Nothing Then Return
                Dim oFacility As Facility = Nothing
                If lTargetID < 1 Then
                    Try
                        oFacility = CType(oTarget.GetRandomUnitOrFacility(1, 0), Facility)
                    Catch
                    End Try
                Else : oFacility = GetEpicaFacility(lTargetID)
                End If
                If oFacility Is Nothing = False Then

                    Dim uAttach() As PlayerComm.WPAttachment = Nothing

                    Dim lID As Int32 = -1
                    Dim iTypeID As Int16 = -1
                    With CType(oFacility.ParentObject, Epica_GUID)
                        lID = .ObjectID
                        iTypeID = .ObjTypeID
                        If .ObjTypeID = ObjectType.eFacility Then
                            With CType(CType(oFacility.ParentObject, Facility).ParentObject, Epica_GUID)
                                lID = .ObjectID
                                iTypeID = .ObjTypeID
                            End With
                        End If
                    End With
                    If lID > -1 AndAlso iTypeID > -1 Then
                        ReDim uAttach(0)
                        With uAttach(0)
                            'oPC.AddEmailAttachment(0, lID, iTypeId, oFacility.LocX, oFacility.LocZ, "Location")
                            .AttachNumber = 0
                            .EnvirID = lID
                            .EnvirTypeID = iTypeID
                            .LocX = oFacility.LocX
                            .LocZ = oFacility.LocZ
                            .sWPName = "Location"
                        End With
                    End If

                    Dim oPC As PlayerComm = oFacility.Owner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                     BytesToString(oFacility.EntityName) & " has been destroyed. The incident appears to be an act of terrorism and is suspected to be supported by a rival empire.", "Facility Sabotage", oFacility.Owner.ObjectID, _
                     GetDateAsNumber(Now), False, oFacility.Owner.sPlayerNameProper, uAttach)
                    If oPC Is Nothing = False Then
                        oFacility.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    End If
                    DestroyEntity(CType(oFacility, Epica_Entity), True, -1, True, "MissionComplete:DestroyFacility")
                End If
            Case eMissionResult.eDestroyCurrentSpecialProject
                If oTarget Is Nothing = False Then
                    Dim oTech As Epica_Tech = oTarget.GetTech(lTargetID, iTargetTypeID)
                    If oTech Is Nothing = False Then
                        oTech.ReducePrimarysPointsProduced(0)
                        Dim oPC As PlayerComm = oTarget.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "Someone has sabotaged our research on the special project '" & BytesToString(oTech.GetTechName()) & "'. Our scientists are scrambling to pick up the pieces and get research progress underway, however, initial assessments indicate a total loss in progress to date.", "Special Project Sabotaged", oTarget.ObjectID, GetDateAsNumber(Now), False, oTarget.sPlayerNameProper, Nothing)
                        If oPC Is Nothing = False Then
                            oTarget.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                        End If
                        oPC = Nothing
                        If (oPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                            oPC = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "Your agent team was successful in sabotaging the special project '" & BytesToString(oTech.GetTechName()) & "' belonging to " & oTarget.sPlayerNameProper, "Mission Success", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                            If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                        End If
                    End If
                End If
            Case eMissionResult.eDoorJamHangar
                If oTarget Is Nothing Then Return
                Dim oEntity As Epica_Entity = Nothing
                If lTargetID < 1 OrElse iTargetTypeID < 1 Then
                    Dim yType As Byte = 1   'default to fac
                    If iTargetTypeID = ObjectType.eUnit Then yType = 0 'change to 1 if needed
                    oEntity = oTarget.GetRandomUnitOrFacility(yType, 2)
                Else : oEntity = GetEpicaEntity(lTargetID, iTargetTypeID)
                End If
                If oEntity Is Nothing = False Then
                    Dim lDuration As Int32 = CInt(Rnd() * 168) + 24       'in hours
                    lDuration *= 108000     'convert to cycles
                    Dim oEffect As AgentEffect = oEntity.AddAgentEffect(glCurrentCycle, lDuration, AgentEffectType.eHangarBay, oMission.BaseEffect, True, oPlayer.ObjectID)
                    If oEffect Is Nothing = False Then oPlayer.AddImposedEffect(oEffect, oTarget, BytesToString(oEntity.EntityName))
                End If
                'Case eMissionResult.eFindCommandCenters
                '	Dim lColonyIdx As Int32 = oTarget.GetColonyFromParent(lTargetID, iTargetTypeID)
                '	'check lcolonyidx is valid
                '	If lColonyIdx = -1 OrElse glColonyIdx(lColonyIdx) = -1 OrElse glColonyIdx(lColonyIdx) <> lTargetID Then
                '		'TODO: Indicate target colony could not be find
                '	Else
                '		Dim oColony As Colony = goColony(lColonyIdx)
                '		If oColony Is Nothing = False Then
                '			For X As Int32 = 0 To oColony.ChildrenUB
                '				If oColony.lChildrenIdx(X) <> -1 AndAlso (oColony.oChildren(X).yProductionType = ProductionType.eCommandCenterSpecial OrElse oColony.oChildren(X).yProductionType = ProductionType.eSpaceStationSpecial) Then
                '					Dim oFac As Facility = oColony.oChildren(X)
                '					Dim oPII As PlayerItemIntel = oPlayer.AddPlayerItemIntel(oFac.ObjectID, ObjectType.eFacility, oFac.ServerIndex, PlayerItemIntel.PlayerItemIntelType.eExistance Or PlayerItemIntel.PlayerItemIntelType.eLocation Or PlayerItemIntel.PlayerItemIntelType.eStatus, oTarget.ObjectID)
                '					Exit For
                '				End If
                '			Next X
                '		End If
                '	End If
            Case eMissionResult.eDropInvulField
                If oTarget Is Nothing = False Then
                    If oTarget.yIronCurtainState <> eIronCurtainState.IronCurtainIsDown AndAlso oTarget.yIronCurtainState <> eIronCurtainState.IronCurtainIsFalling Then
                        'AddToQueue(glCurrentCycle + 60, QueueItemType.eIronCurtainFall, oTarget.ObjectID, -1, -1, -1, 0, 0, 0, Int32.MaxValue)

                        oTarget.lCurrent15MinutesRemaining = 0
                        oTarget.bInFullLockDown = False
                        oTarget.CancelIronCurtain()
                        oTarget.yIronCurtainState = eIronCurtainState.IronCurtainIsDown

                        AddToQueue(glCurrentCycle + 108000, QueueItemType.eIronCurtainRaise, oTarget.ObjectID, 0, 0, 0, 0, 0, 0, 0)
                        Dim oPC As PlayerComm = oTarget.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "The offline invulnerability field has been dropped due to sabotage. Your engineers believe they can restore it within one hour but your empire is defenseless until it is restored!", "Sabotage!", oTarget.ObjectID, GetDateAsNumber(Now), False, oTarget.sPlayerNameProper, Nothing)
                        If oPC Is Nothing = False Then
                            goMsgSys.SendOutboundEmail(oPC, oTarget, 0, 0, 0, 0, 0, 0, 0, "")
                            oTarget.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                        End If
                        oPC = Nothing
                        If (oPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                            oPC = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "Your agents have successfully sabotaged the invulnerability fields belonging to " & oTarget.sPlayerNameProper & ". The fields should remain down for 1 hour.", "Drop Invulnerability Mission Complete", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                            oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                        End If
                    End If
                End If
            Case eMissionResult.eFindMineral
                'Compile an Email with the locational Waypoints of each location that the mineral was found and send it to oPlayer
                Dim oSB As New System.Text.StringBuilder
                Dim oMineral As Mineral = GetEpicaMineral(lTargetID)
                If oMineral Is Nothing Then
                    LogEvent(LogEventType.PossibleCheat, "HandleMissionCompletion.FindMineral: Mineral not found (" & lTargetID & "). Player: " & Me.oPlayer.ObjectID)
                    Return
                End If

                If iTargetTypeID2 = ObjectType.eSolarSystem Then
                    Dim oSystem As SolarSystem = GetEpicaSystem(lTargetID2)
                    If oSystem Is Nothing Then
                        LogEvent(LogEventType.PossibleCheat, "HandleMissionCompletion.FindMineral: System not found (" & lTargetID2 & "). Player: " & Me.oPlayer.ObjectID)
                        Return
                    End If
                    oSB.AppendLine("Your agents have finished their scouting of " & BytesToString(oSystem.SystemName) & " for " & BytesToString(oMineral.MineralName) & ".")
                    oSystem = Nothing
                Else
                    Dim oPlanet As Planet = GetEpicaPlanet(lTargetID2)
                    If oPlanet Is Nothing Then
                        LogEvent(LogEventType.PossibleCheat, "HandleMissionCompletion.FindMineral: Planet not found (" & lTargetID2 & "). Player: " & Me.oPlayer.ObjectID)
                        Return
                    End If
                    oSB.AppendLine("Your agents have finished their scouting of " & BytesToString(oPlanet.PlanetName) & " for " & BytesToString(oMineral.MineralName) & ".")
                    oPlanet = Nothing
                End If

                Dim ptLocs() As Point = Nothing
                Dim ptGUID() As Point = Nothing
                Dim sPtName() As String = Nothing
                Dim lLocUB As Int32 = -1

                Dim lMinUB As Int32 = -1
                If glMineralCacheIdx Is Nothing = False Then lMinUB = Math.Min(glMineralCacheUB, glMineralCacheIdx.GetUpperBound(0))
                Try
                    For X As Int32 = 0 To lMinUB
                        If glMineralCacheIdx(X) <> -1 Then
                            Dim oCache As MineralCache = goMineralCache(X)
                            If oCache Is Nothing = False AndAlso oCache.oMineral Is Nothing = False AndAlso oCache.oMineral.ObjectID = lTargetID AndAlso oCache.CacheTypeID = MineralCacheType.eMineable Then
                                With CType(oCache.ParentObject, Epica_GUID)

                                    If (iTargetTypeID2 = ObjectType.eSolarSystem AndAlso .ObjTypeID = ObjectType.ePlanet AndAlso CType(oCache.ParentObject, Planet).ParentSystem Is Nothing = False AndAlso CType(oCache.ParentObject, Planet).ParentSystem.ObjectID = lTargetID2) OrElse (.ObjectID = lTargetID2 AndAlso .ObjTypeID = iTargetTypeID2) Then
                                        lLocUB += 1
                                        ReDim Preserve ptLocs(lLocUB)
                                        ReDim Preserve sPtName(lLocUB)
                                        ReDim Preserve ptGUID(lLocUB)
                                        ptLocs(lLocUB).X = oCache.LocX
                                        ptLocs(lLocUB).Y = oCache.LocZ
                                        ptGUID(lLocUB).X = .ObjectID
                                        ptGUID(lLocUB).Y = .ObjTypeID

                                        Dim sParentName As String = ""
                                        If .ObjTypeID = ObjectType.ePlanet Then
                                            With CType(oCache.ParentObject, Planet)
                                                sParentName = BytesToString(.PlanetName)
                                                If .ParentSystem Is Nothing = False Then
                                                    Dim sSysName As String = BytesToString(.ParentSystem.SystemName)
                                                    sParentName = sParentName.Replace(sSysName, "").Trim
                                                End If
                                                sParentName = sParentName.Replace("(S)", "").Trim
                                                sParentName &= " "
                                            End With
                                        End If

                                        sPtName(lLocUB) = sParentName & oCache.Concentration & "/" & oCache.Quantity
                                        If sPtName(lLocUB).Length > 20 Then sPtName(lLocUB) = sPtName(lLocUB).Substring(0, 20)
                                    End If
                                End With
                            End If
                        End If
                    Next X
                Catch
                End Try

                Dim lMax As Int32 = Math.Min(lLocUB + 1, 200)
                If lLocUB = -1 Then
                    oSB.AppendLine("After an exhaustive search, they could not find any mineable quantities of " & BytesToString(oMineral.MineralName) & ".")
                Else
                    oSB.AppendLine("After an exhaustive search, they have found " & (lLocUB + 1) & " mineable caches of " & BytesToString(oMineral.MineralName) & ".")
                    If lMax <> lLocUB + 1 Then
                        oSB.AppendLine("The first " & lMax & " caches have been attached as waypoints to this message.")
                    Else : oSB.AppendLine("All of the caches found have been attached as waypoints to this message.")
                    End If
                End If

                Dim uAttach(lMax - 1) As PlayerComm.WPAttachment
                For X As Int32 = 0 To lMax - 1
                    With uAttach(X)
                        'CByte(X) , ptGUID(X).X, CShort(ptGUID(X).Y), ptLocs(X).X, ptLocs(X).Y, sPtName(X))
                        .AttachNumber = CByte(X)
                        .EnvirID = ptGUID(X).X
                        .EnvirTypeID = CShort(ptGUID(X).Y)
                        .LocX = ptLocs(X).X
                        .LocZ = ptLocs(X).Y
                        .sWPName = sPtName(X)
                    End With
                Next X

                Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Survey Mission Completed for " & BytesToString(oMineral.MineralName), oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, uAttach)
                If oPC Is Nothing = False Then
                    If oPC.SaveObject(oPlayer.ObjectID) = False Then
                        LogEvent(LogEventType.CriticalError, "Survey mission resulted in a bad saveobjecto on the playercomm.")
                    Else
                        oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    End If
                    'For X As Int32 = 0 To lMax - 1
                    '    oPC.AddEmailAttachment(CByte(X), ptGUID(X).X, CShort(ptGUID(X).Y), ptLocs(X).X, ptLocs(X).Y, sPtName(X))
                    'Next X
                End If

                'Case eMissionResult.eFindProductionFac
                '	Dim lColonyIdx As Int32 = oTarget.GetColonyFromParent(lTargetID, iTargetTypeID)
                '	If lColonyIdx = -1 OrElse glColonyIdx(lColonyIdx) = -1 OrElse glColonyIdx(lColonyIdx) <> lTargetID Then
                '		'TODO: Indicate target colony could not be find
                '	Else
                '		Dim oColony As Colony = goColony(lColonyIdx)
                '		If oColony Is Nothing = False Then
                '			For X As Int32 = 0 To oColony.ChildrenUB
                '				If oColony.lChildrenIdx(X) <> -1 AndAlso (oColony.oChildren(X).yProductionType And ProductionType.eProduction) <> 0 Then
                '					Dim oFac As Facility = oColony.oChildren(X)
                '					Dim oPII As PlayerItemIntel = oPlayer.AddPlayerItemIntel(oFac.ObjectID, ObjectType.eFacility, oFac.ServerIndex, PlayerItemIntel.PlayerItemIntelType.eExistance Or PlayerItemIntel.PlayerItemIntelType.eLocation Or PlayerItemIntel.PlayerItemIntelType.eStatus, oTarget.ObjectID)
                '					If oPII Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                '				End If
                '			Next X
                '		End If
                '	End If
                'Case eMissionResult.eFindResearchFac
                '	Dim lColonyIdx As Int32 = oTarget.GetColonyFromParent(lTargetID, iTargetTypeID)
                '	If lColonyIdx = -1 OrElse glColonyIdx(lColonyIdx) = -1 OrElse glColonyIdx(lColonyIdx) <> lTargetID Then
                '		'TODO: Indicate target colony could not be find
                '	Else
                '		Dim oColony As Colony = goColony(lColonyIdx)
                '		If oColony Is Nothing = False Then
                '			For X As Int32 = 0 To oColony.ChildrenUB
                '				If oColony.lChildrenIdx(X) <> -1 AndAlso oColony.oChildren(X).yProductionType = ProductionType.eResearch Then
                '					Dim oFac As Facility = oColony.oChildren(X)
                '					Dim oPII As PlayerItemIntel = oPlayer.AddPlayerItemIntel(oFac.ObjectID, ObjectType.eFacility, oFac.ServerIndex, PlayerItemIntel.PlayerItemIntelType.eExistance Or PlayerItemIntel.PlayerItemIntelType.eLocation Or PlayerItemIntel.PlayerItemIntelType.eStatus, oTarget.ObjectID)
                '					If oPII Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                '				End If
                '			Next X
                '		End If
                '	End If
                'Case eMissionResult.eFindSpaceStations
                '	For X As Int32 = 0 To glFacilityUB
                '		If glFacilityIdx(X) <> -1 Then
                '			Dim oFac As Facility = goFacility(X)
                '			If oFac Is Nothing = False AndAlso oFac.Owner.ObjectID = oTarget.ObjectID Then
                '				If oFac.yProductionType = ProductionType.eSpaceStationSpecial Then
                '					If CType(oFac.ParentObject, Epica_GUID).ObjectID = lTargetID Then
                '						Dim oPII As PlayerItemIntel = oPlayer.AddPlayerItemIntel(oFac.ObjectID, ObjectType.eFacility, oFac.ServerIndex, PlayerItemIntel.PlayerItemIntelType.eExistance Or PlayerItemIntel.PlayerItemIntelType.eLocation Or PlayerItemIntel.PlayerItemIntelType.eStatus, oTarget.ObjectID)
                '						If oPII Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                '					End If
                '				End If
                '			End If
                '		End If
                '	Next X
            Case eMissionResult.eGeologicalSurvey
                'Compile an email of the data for the mineral list of the target planet and send it to oPlayer
                Dim oPlanet As Planet = GetEpicaPlanet(lTargetID)
                Dim oSB As New System.Text.StringBuilder

                oSB.AppendLine("Your agent team has finished their survey of " & BytesToString(oPlanet.PlanetName) & " and have compiled these results.")
                Dim lMinID() As Int32 = Nothing
                Dim sMinName() As String = Nothing
                Dim blMinAmt() As Int64 = Nothing
                Dim lMaxConc() As Int32 = Nothing
                Dim lMinUB As Int32 = -1

                For X As Int32 = 0 To glMineralCacheUB
                    If glMineralCacheIdx(X) <> -1 Then
                        Dim oCache As MineralCache = goMineralCache(X)
                        If oCache Is Nothing = False AndAlso oCache.CacheTypeID = MineralCacheType.eMineable Then
                            With CType(oCache.ParentObject, Epica_GUID)
                                If .ObjectID <> lTargetID OrElse .ObjTypeID <> iTargetTypeID Then Continue For
                            End With
                            Dim lIdx As Int32 = -1
                            Dim lID As Int32 = oCache.oMineral.ObjectID
                            For Y As Int32 = 0 To lMinUB
                                If lMinID(Y) = lID Then
                                    lIdx = Y
                                    Exit For
                                End If
                            Next Y
                            If lIdx = -1 Then
                                lMinUB += 1
                                ReDim Preserve lMinID(lMinUB)
                                ReDim Preserve sMinName(lMinUB)
                                ReDim Preserve blMinAmt(lMinUB)
                                ReDim Preserve lMaxConc(lMinUB)
                                lIdx = lMinUB
                                blMinAmt(lIdx) = 0
                                lMaxConc(lIdx) = 0
                                sMinName(lIdx) = BytesToString(oCache.oMineral.MineralName)
                                lMinID(lIdx) = lID
                            End If

                            blMinAmt(lIdx) += oCache.Quantity
                            If oCache.Concentration > lMaxConc(lIdx) Then lMaxConc(lIdx) = oCache.Concentration
                        End If
                    End If
                Next X

                For X As Int32 = 0 To lMinUB
                    oSB.AppendLine(sMinName(X) & " found. Estimated Quantity: " & blMinAmt(X).ToString("#,###") & ". Max Concentration: " & lMaxConc(X).ToString("#,###"))
                Next X
                If lMinUB = -1 Then oSB.AppendLine("The planet is mineral poor. No mineral elements could be found.")

                Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Survey Mission Completed", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            Case eMissionResult.eGetAgentList
                Dim bFoundAgents As Boolean = False
                If oTarget Is Nothing = True Then Return
                Dim oSB As New System.Text.StringBuilder
                oSB.AppendLine("Your agents have finished assessing the target's Agency and have found the following agents on their payroll:")
                For X As Int32 = 0 To oTarget.mlAgentUB
                    If oTarget.mlAgentIdx(X) > -1 Then
                        If glAgentIdx(oTarget.mlAgentIdx(X)) = oTarget.mlAgentID(X) Then
                            Dim oAgent As Agent = goAgent(oTarget.mlAgentIdx(X))
                            If oAgent Is Nothing = False AndAlso (oAgent.lAgentStatus And (AgentStatus.IsDead Or AgentStatus.Dismissed)) = 0 Then
                                Dim oPII As PlayerItemIntel = Me.oPlayer.AddPlayerItemIntel(oAgent.ObjectID, oAgent.ObjTypeID, PlayerItemIntel.PlayerItemIntelType.eStatus, Me.oTarget.ObjectID)
                                If oPII Is Nothing = False Then
                                    oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                                End If
                                oSB.AppendLine("     " & BytesToString(oAgent.AgentName))
                                bFoundAgents = True
                            End If
                        End If
                    End If
                Next X
                'TODO: figure out how to include captured agents effectively
                If bFoundAgents = True Then
                    Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Get Agent List Mission Result", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                    If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                End If

            Case eMissionResult.eGetAliasList
                If oTarget Is Nothing = False Then
                    Dim sBody As String = oTarget.GetAliasMissionList()
                    If sBody Is Nothing = False AndAlso sBody.Trim <> "" Then
                        Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sBody, "Alias List for " & oTarget.sPlayerNameProper, oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                        If oPC Is Nothing = False Then
                            oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                        End If
                    End If
                End If
            Case eMissionResult.eGetAlliesList
                If oTarget Is Nothing Then Return
                Dim bFoundAllies As Boolean = False
                Dim oSB As New System.Text.StringBuilder
                oSB.AppendLine("Your agents have finished assessing the target's enemy relations and have determined they are allied with following empires:")
                For X As Int32 = 0 To oTarget.PlayerRelUB
                    Dim oPR As PlayerRel = oTarget.GetPlayerRelByIndex(X)
                    If oPR.WithThisScore > elRelTypes.eNeutral Then
                        Dim oPII As PlayerItemIntel = oPlayer.AddPlayerItemIntel(oPR.oThisPlayer.ObjectID, ObjectType.ePlayer, PlayerItemIntel.PlayerItemIntelType.eStatus, oTarget.ObjectID)
                        If oPII Is Nothing = False Then
                            oPII.lValue = oPR.WithThisScore
                            oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                            oSB.AppendLine("     " & oPR.oThisPlayer.sPlayerNameProper)
                            bFoundAllies = True
                        End If
                    End If
                Next X
                If bFoundAllies = True Then
                    Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Get Ally List Mission Result", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                    If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                End If
            Case eMissionResult.eGetColonyBudget
                If oTarget Is Nothing Then Return
                'Return the budget of the target colony in an email and send it
                Dim sText As String = oTarget.oBudget.GetEnvirBudgetText(lTargetID, iTargetTypeID)

                Dim lColonyIdx As Int32 = oTarget.GetColonyFromParent(lTargetID, iTargetTypeID)
                If lColonyIdx <> -1 AndAlso glColonyIdx(lColonyIdx) <> -1 Then
                    Dim oColony As Colony = goColony(lColonyIdx)
                    If oColony Is Nothing = False Then
                        Dim oPII As PlayerItemIntel = oPlayer.AddPlayerItemIntel(oColony.ObjectID, ObjectType.eColony, PlayerItemIntel.PlayerItemIntelType.eLocation, oTarget.ObjectID)
                        If oPII Is Nothing = False Then
                            oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                        End If
                    End If
                End If

                Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sText, "Get Budget Mission Completed", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            Case eMissionResult.eGetEnemyList
                If oTarget Is Nothing Then Return
                Dim bFoundEnemies As Boolean = False
                Dim oSB As New System.Text.StringBuilder
                oSB.AppendLine("Your agents have finished assessing the target's enemy relations and have determined they are at war with following empires:")
                For X As Int32 = 0 To oTarget.PlayerRelUB
                    Dim oPR As PlayerRel = oTarget.GetPlayerRelByIndex(X)
                    If oPR.WithThisScore <= elRelTypes.eWar Then
                        Dim oPII As PlayerItemIntel = oPlayer.AddPlayerItemIntel(oPR.oThisPlayer.ObjectID, ObjectType.ePlayer, PlayerItemIntel.PlayerItemIntelType.eStatus, oTarget.ObjectID)
                        bFoundEnemies = True
                        oSB.AppendLine("     " & oPR.oThisPlayer.sPlayerNameProper)
                        If oPII Is Nothing = False Then
                            oPII.lValue = oPR.WithThisScore
                            oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                        End If
                    End If
                Next X
                If bFoundEnemies = True Then
                    Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Get Enemy List Mission Result", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                    If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                End If
            Case eMissionResult.eGetFacilityList
                'Returns a list of facilities in the environment, based on oMission.BaseEffect
                If iTargetTypeID2 = ObjectType.eSolarSystem AndAlso oTarget Is Nothing = False Then
                    If lTargetID = ProductionType.eSpaceStationSpecial Then
                        Dim lCnt As Int32 = 0
                        For X As Int32 = 0 To oTarget.mlColonyUB
                            If oTarget.mlColonyIdx(X) <> -1 AndAlso glColonyIdx(oTarget.mlColonyIdx(X)) = oTarget.mlColonyID(X) Then
                                Dim oColony As Colony = goColony(oTarget.mlColonyIdx(X))
                                If oColony Is Nothing = False Then
                                    If oColony.ParentObject Is Nothing = False AndAlso CType(oColony.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                                        lCnt += 1
                                    End If
                                End If
                            End If
                        Next X
                        lCnt = CInt(Math.Floor(Rnd() * lCnt))
                        For X As Int32 = 0 To oTarget.mlColonyUB
                            If oTarget.mlColonyIdx(X) <> -1 AndAlso glColonyIdx(oTarget.mlColonyIdx(X)) = oTarget.mlColonyID(X) Then
                                Dim oColony As Colony = goColony(oTarget.mlColonyIdx(X))
                                If oColony Is Nothing = False Then
                                    If oColony.ParentObject Is Nothing = False AndAlso CType(oColony.ParentObject, Epica_GUID).ObjTypeID = ObjectType.eFacility Then
                                        lCnt -= 1
                                        If lCnt < 1 Then
                                            Dim oPII As PlayerItemIntel = oPlayer.AddPlayerItemIntel(oColony.ObjectID, ObjectType.eColony, PlayerItemIntel.PlayerItemIntelType.eLocation, oTarget.ObjectID)
                                            If oPII Is Nothing = False Then
                                                oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                                            End If
                                            oPII = Nothing

                                            Dim oSB As New System.Text.StringBuilder
                                            oSB.AppendLine("Your agents have finished locating the target's Space Stations and have found the following:")
                                            Dim oFacility As Facility = GetEpicaFacility(CType(oColony.ParentObject, Epica_GUID).ObjectID)
                                            If oFacility Is Nothing = False Then
                                                oSB.AppendLine("     " & BytesToString(CType(oFacility, Epica_Entity).EntityName) & " located in " & BytesToString(oColony.ColonyName))
                                            Else
                                                oSB.AppendLine("     Facility located in " & BytesToString(oColony.ColonyName))
                                            End If
                                            Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Get Facility List Mission Result", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                                            oPC.AddEmailAttachment(1, CType(oFacility.ParentObject, Epica_GUID).ObjectID, ObjectType.eSolarSystem, oFacility.LocX, oFacility.LocZ, BytesToString(CType(oFacility, Epica_Entity).EntityName))
                                            If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        Next X
                    End If
                ElseIf oTarget Is Nothing = False Then
                    Dim lColonyIdx As Int32 = oTarget.GetColonyFromParent(lTargetID2, iTargetTypeID2)
                    'Verify lColonyIdx
                    If lColonyIdx <> -1 AndAlso glColonyIdx(lColonyIdx) <> -1 Then
                        Dim oColony As Colony = goColony(lColonyIdx)
                        If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID = oTarget.ObjectID Then
                            Dim oPII As PlayerItemIntel = oPlayer.AddPlayerItemIntel(oColony.ObjectID, ObjectType.eColony, PlayerItemIntel.PlayerItemIntelType.eLocation, oTarget.ObjectID)
                            If oPII Is Nothing = False Then
                                oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                            End If
                            oPII = Nothing
                            Dim bFoundFacility As Boolean = False
                            Dim oSB As New System.Text.StringBuilder
                            oSB.AppendLine("Your agents have finished assessing the target's facility list for " & BytesToString(oColony.ColonyName) & " and have found the following structures:")
                            For X As Int32 = 0 To oColony.ChildrenUB
                                If oColony.lChildrenIdx(X) > -1 AndAlso oColony.oChildren(X).yProductionType = Me.lTargetID Then
                                    Dim oFac As Facility = oColony.oChildren(X)
                                    oPII = oPlayer.AddPlayerItemIntel(oFac.ObjectID, ObjectType.eFacility, PlayerItemIntel.PlayerItemIntelType.eLocation, oTarget.ObjectID)
                                    If oPII Is Nothing = False Then
                                        oPII.lValue = oFac.EntityDef.ModelID
                                        oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                                    End If
                                    bFoundFacility = True
                                    oSB.AppendLine("     " & BytesToString(CType(oFac, Epica_Entity).EntityName))
                                End If
                            Next X
                            If bFoundFacility = True Then
                                Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Get Facility List Mission Result", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)

                                Dim iAttachmentNumber As Int32 = 1
                                For X As Int32 = 0 To oColony.ChildrenUB
                                    If oColony.lChildrenIdx(X) > -1 AndAlso oColony.oChildren(X).yProductionType = Me.lTargetID Then
                                        Dim oFac As Facility = oColony.oChildren(X)
                                        iAttachmentNumber += 1
                                        oPC.AddEmailAttachment(CByte(iAttachmentNumber), CType(oFac.ParentObject, Epica_GUID).ObjectID, ObjectType.ePlanet, oFac.LocX, oFac.LocZ, BytesToString(CType(oFac, Epica_Entity).EntityName))
                                        If iAttachmentNumber > 200 Then Exit For

                                    End If
                                Next X

                                If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                            End If
                        End If
                    End If
                End If
            Case eMissionResult.eGetMineralList
                If oTarget Is Nothing Then Return
                Dim oSB As New System.Text.StringBuilder
                oSB.AppendLine("Your agents have finished assessing the target's mineral list and have determined the following results:")
                oSB.AppendLine("(All values are approximations)")
                For X As Int32 = 0 To oTarget.lPlayerMineralUB
                    If oTarget.oPlayerMinerals(X) Is Nothing = False Then
                        Dim blTotal As Int64 = oTarget.GetAllColonyCargoForMineral(oTarget.oPlayerMinerals(X).Mineral.ObjectID)
                        oSB.AppendLine(BytesToString(oTarget.oPlayerMinerals(X).Mineral.MineralName) & " - " & blTotal.ToString("#,##0") & " empire-wide")
                    End If
                Next X
                Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Get Mineral List Mission Result", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            Case eMissionResult.eGetTradeTreaties
                'If oTarget Is Nothing = False Then oTarget.oBudget.HandleTradeTreatyEspionage(oPlayer)
                If oTarget Is Nothing = False AndAlso oTarget.oBudget Is Nothing = False Then
                    Dim bFoundTraders As Boolean = False
                    Dim oSB As New System.Text.StringBuilder
                    oSB.AppendLine("Your agents have finished assessing the target's trade relations and have determined they have trade income from the following empires:")
                    With oTarget.oBudget
                        If .lTradePlayerID Is Nothing OrElse .lTradeValue Is Nothing Then Return
                        For Y As Int32 = 0 To .lTradePlayerID.GetUpperBound(0)
                            Dim oTrader As Player = GetEpicaPlayer(.lTradePlayerID(Y))
                            If oTrader Is Nothing = False Then
                                oSB.AppendLine("     " & oTrader.sPlayerNameProper)
                                bFoundTraders = True
                            End If
                        Next Y
                    End With
                    If bFoundTraders = True Then
                        Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Get Trade Treties Mission Result", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                        If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    End If
                End If
            Case eMissionResult.eIFFSabotage
                'TODO: Handle this
            Case eMissionResult.eImpedeCurrentDevelopment
                If oTarget Is Nothing = False Then
                    'Dim oTech As Epica_Tech = oTarget.GetTech(lTargetID, iTargetTypeID)
                    'If oTech Is Nothing = False Then
                    '    oTech.ReducePrimarysPointsProduced(0.5F)
                    '    Dim oPC As PlayerComm = oTarget.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "Someone has sabotaged our research on the special project '" & BytesToString(oTech.GetTechName()) & "'. The scientists assigned to the project indicate that half of their work has been lost.", "Special Project Sabotaged", oTarget.ObjectID, GetDateAsNumber(Now), False, oTarget.sPlayerNameProper, Nothing)
                    '    If oPC Is Nothing = False Then
                    '        oTarget.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    '    End If
                    '    oPC = Nothing
                    '    If (oPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                    '        oPC = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "Your agent team was successful in sabotaging the special project '" & BytesToString(oTech.GetTechName()) & "' belonging to " & oTarget.sPlayerNameProper, "Mission Success", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                    '        If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    '    End If
                    'End If
                    Dim bFoundTech As Boolean = False
                    Dim oTargets() As Int32
                    ReDim oTargets(0)
                    For X As Int32 = 0 To oTarget.oSpecials.mlSpecialTechUB
                        If oTarget.oSpecials.mlSpecialTechIdx(X) <> -1 Then
                            If oTarget.oSpecials.moSpecialTech(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eComponentResearching AndAlso oTarget.oSpecials.moSpecialTech(X).GetResearcherCount > 0 Then
                                ReDim oTargets(oTargets.GetUpperBound(0) + 1)
                                oTargets(oTargets.GetUpperBound(0)) = oTarget.oSpecials.moSpecialTech(X).ObjectID
                                bFoundTech = True
                            End If
                        End If
                    Next

                    If bFoundTech = True Then
                        Dim lTargetID As Int32 = CInt(Rnd() * (oTargets.GetUpperBound(0) - 1)) + 1
                        Dim oTech As Epica_Tech = oTarget.GetTech(oTargets(lTargetID), ObjectType.eSpecialTech)

                        If oTech Is Nothing = False Then
                            oTech.ReducePrimarysPointsProduced(0.5F)
                            Dim oPC As PlayerComm = oTarget.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "Someone has sabotaged our research on the special project '" & BytesToString(oTech.GetTechName()) & "'. The scientists assigned to the project indicate that half of their work has been lost.", "Special Project Sabotaged", oTarget.ObjectID, GetDateAsNumber(Now), False, oTarget.sPlayerNameProper, Nothing)
                            If oPC Is Nothing = False Then
                                oTarget.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                            End If
                            oPC = Nothing
                            If (oPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                                oPC = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "Your agent team was successful in sabotaging the special project '" & BytesToString(oTech.GetTechName()) & "' belonging to " & oTarget.sPlayerNameProper, "Mission Success", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                                If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                            End If
                        End If
                    End If
                End If
            Case eMissionResult.eIncreasePowerNeeds
                If oTarget Is Nothing = False Then
                    If iTargetTypeID = ObjectType.eColony Then
                        Dim oColony As Colony = GetEpicaColony(lTargetID)
                        If oColony Is Nothing = False AndAlso oColony.Owner.ObjectID = oTarget.ObjectID Then
                            If oPlayer.HasItemIntel(oColony.ObjectID, oColony.ObjTypeID, oColony.Owner.ObjectID) = False Then
                                Dim oPII As PlayerItemIntel = oPlayer.AddPlayerItemIntel(oColony.ObjectID, oColony.ObjTypeID, PlayerItemIntel.PlayerItemIntelType.eLocation, oColony.Owner.ObjectID)
                                If oPII Is Nothing = False Then
                                    oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                                End If
                            End If

                            'ok, do it
                            Dim lDuration As Int32 = CInt(Rnd() * 168) + 24          'in hours
                            lDuration *= 108000     'convert to cycles
                            Dim oEffect As AgentEffect = oColony.AddAgentEffect(glCurrentCycle, lDuration, AgentEffectType.eIncreasePowerNeed, 10, True, oPlayer.ObjectID)
                            If oEffect Is Nothing = False Then oPlayer.AddImposedEffect(oEffect, oTarget, BytesToString(oColony.ColonyName))
                            oColony.UpdateAllValues(-1)
                        End If
                    End If
                End If
            Case eMissionResult.eLocateWormhole
                If oTarget Is Nothing = False Then
                    Dim lCnt As Int32 = 0
                    Dim bValid(oTarget.lWormholeUB) As Boolean
                    For X As Int32 = 0 To oTarget.lWormholeUB
                        If oTarget.oWormholes(X) Is Nothing = False AndAlso oPlayer.HasWormholeKnowledge(oTarget.oWormholes(X).ObjectID) = False Then
                            If oPlayer.oBudget.IsInEnvironment(oTarget.oWormholes(X).System1.ObjectID, ObjectType.eSolarSystem) = True OrElse _
                                oPlayer.oBudget.IsInEnvironment(oTarget.oWormholes(X).System2.ObjectID, ObjectType.eSolarSystem) = True Then
                                lCnt += 1
                                bValid(X) = True
                            End If
                        End If
                    Next X
                    If lCnt = 0 Then
                        If (oPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                            Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oTarget.sPlayerNameProper & " does not have any wormhole knowledge that you do not already possess.", "Locate Wormhole Completed", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                            oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents Or AliasingRights.eViewEmail)
                        End If
                    Else
                        Dim lVal As Int32 = CInt(Rnd() * (lCnt - 1)) + 1        'returns 1 to lCnt
                        Dim oWormhole As Wormhole = Nothing
                        For X As Int32 = 0 To oTarget.lWormholeUB
                            If bValid(X) = True Then
                                lVal -= 1
                                oWormhole = oTarget.oWormholes(X)
                                If lVal = 0 Then
                                    Exit For
                                End If
                            End If
                        Next X
                        If oWormhole Is Nothing = True Then
                            If (oPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                                Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oTarget.sPlayerNameProper & " does not have any wormhole knowledge that you do not already possess.", "Locate Wormhole Completed", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                                oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents Or AliasingRights.eViewEmail)
                            End If
                        Else
                            Dim oSys As SolarSystem = Nothing
                            Dim oSB As New System.Text.StringBuilder
                            If oWormhole.System1.SystemType < oWormhole.System2.SystemType Then
                                oSys = oWormhole.System1
                                oSB.AppendLine("Wormhole found from " & BytesToString(oWormhole.System1.SystemName) & " to " & BytesToString(oWormhole.System2.SystemName))
                            Else
                                oSys = oWormhole.System2
                                oSB.AppendLine("Wormhole found from " & BytesToString(oWormhole.System2.SystemName) & " to " & BytesToString(oWormhole.System1.SystemName))
                            End If
                            oPlayer.AddWormholeKnowledge(oWormhole, True, oSys, True)
                            If (oPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                                Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "Your agent team was successful in stealing some of " & oTarget.sPlayerNameProper & "'s wormhole knowledge." & vbCrLf & oSB.ToString, "Locate Wormhole Completed", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                                oPC.AddEmailAttachment(1, oWormhole.System1.ObjectID, ObjectType.eSolarSystem, oWormhole.LocX1, oWormhole.LocY1, "Wormhole " & BytesToString(oWormhole.System1.SystemName))
                                oPC.AddEmailAttachment(2, oWormhole.System2.ObjectID, ObjectType.eSolarSystem, oWormhole.LocX2, oWormhole.LocY2, "Wormhole " & BytesToString(oWormhole.System2.SystemName))
                                oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents Or AliasingRights.eViewEmail)
                            End If
                        End If
                    End If
                End If
            Case eMissionResult.eMilitaryCoup
                'TODO: Handle this
            Case eMissionResult.ePlantEvidence
                'TODO: Handle this
            Case eMissionResult.eProductionScore
                If oTarget Is Nothing = False Then
                    Dim oPlayerIntel As PlayerIntel = oPlayer.GetOrAddPlayerIntel(oTarget.ObjectID, False)
                    oPlayerIntel.ProductionScore = oTarget.ProductionScore
                    oPlayerIntel.ProductionUpdate = glCurrentCycle
                    oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPlayerIntel, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                End If
            Case eMissionResult.eReconPlanetMap
                'TODO: Handle this
            Case eMissionResult.eRelayBattleReports
                'TODO: Handle this
            Case eMissionResult.eSabotageProduction
                If oTarget Is Nothing = False Then
                    If oMission.lInfiltrationType = eInfiltrationType.eMiningInfiltration Then
                        'Get Unit that has cargo route including this facility
                        Dim lTargetPlayerID As Int32 = oTarget.ObjectID
                        Dim oFac As Facility = Nothing
                        If lTargetID < 1 Then
                            Try
                                oFac = CType(oTarget.GetRandomUnitOrFacility(1, 0), Facility)
                            Catch
                            End Try
                        Else : oFac = GetEpicaFacility(lTargetID)
                        End If
                        If oFac Is Nothing = False Then
                            For X As Int32 = 0 To glUnitUB
                                If glUnitIdx(X) <> -1 Then
                                    Dim oUnit As Unit = goUnit(X)
                                    If oUnit Is Nothing = False AndAlso oUnit.Owner.ObjectID = lTargetPlayerID Then
                                        If oUnit.RouteContainsFacility(lTargetID) = True Then

                                            Dim uAttach(0) As PlayerComm.WPAttachment
                                            Dim lID As Int32 = -1
                                            Dim iTypeID As Int16 = -1
                                            With CType(oUnit.ParentObject, Epica_GUID)
                                                lID = .ObjectID
                                                iTypeID = .ObjTypeID
                                                If .ObjTypeID = ObjectType.eFacility Then
                                                    With CType(CType(oUnit.ParentObject, Facility).ParentObject, Epica_GUID)
                                                        lID = .ObjectID
                                                        iTypeID = .ObjTypeID
                                                    End With
                                                End If
                                            End With

                                            With uAttach(0)
                                                'oPC.AddEmailAttachment(0, lID, iTypeId, oUnit.LocX, oUnit.LocZ, "Location")
                                                .AttachNumber = 0
                                                .EnvirID = lID
                                                .EnvirTypeID = iTypeID
                                                .LocX = oUnit.LocX
                                                .LocZ = oUnit.LocZ
                                                .sWPName = "Location"
                                            End With

                                            Dim oPC As PlayerComm = oUnit.Owner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                                             BytesToString(oUnit.EntityName) & " has been destroyed. The incident appears to be an act of terrorism and is suspected to be supported by a rival empire.", "Production Sabotage", oUnit.Owner.ObjectID, _
                                             GetDateAsNumber(Now), False, oUnit.Owner.sPlayerNameProper, uAttach)
                                            If oPC Is Nothing = False Then
                                                oUnit.Owner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                                            End If

                                            'destroy the unit
                                            DestroyEntity(CType(oUnit, Epica_Entity), True, -1, False, "MissionComplete:SabotageProd")
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next X
                        End If

                    ElseIf oMission.lInfiltrationType = eInfiltrationType.eAgencyInfiltration Then
                        oTarget.oSecurity.lAllMissionModifier += oMission.BaseEffect
                    Else
                        Dim oFac As Facility = Nothing
                        If lTargetID < 1 Then
                            Try
                                oFac = CType(oTarget.GetRandomUnitOrFacility(1, 0), Facility)
                            Catch
                            End Try
                        Else : oFac = GetEpicaFacility(lTargetID)
                        End If
                        If oFac Is Nothing = False AndAlso oFac.CurrentProduction Is Nothing = False AndAlso oFac.bProducing = True Then
                            oFac.CurrentProduction.PointsProduced = 0
                            Dim yRes As eResourcesResult = oFac.ParentColony.HasRequiredResources(oFac.CurrentProduction.ProdCost, oFac, eyHasRequiredResourcesFlags.NoFlags)
                            If yRes = eResourcesResult.Insufficient_Clear Then
                                oFac.CurrentProduction = Nothing
                            End If
                        End If
                    End If
                End If
            Case eMissionResult.eSearchAndRescueAgent
                'search and rescue was successful, return target agent to my agent list
                Dim oAgent As Agent = GetEpicaAgent(lTargetID)
                If oAgent Is Nothing = False Then
                    If (oAgent.lAgentStatus And AgentStatus.IsDead) = 0 Then
                        'send a msg to the owner
                        Dim oCaptor As Player = GetEpicaPlayer(oAgent.lCapturedBy)
                        Dim sCaptorStr As String = ""
                        If oCaptor Is Nothing = False Then
                            sCaptorStr = vbCrLf & vbCrLf & "The team found the agent in the custody of " & oCaptor.sPlayerNameProper & "."
                        End If
                        Dim oPC As PlayerComm = Me.oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                           "Your agent team reports that " & BytesToString(oAgent.AgentName) & " has been rescued and is on their way home." & sCaptorStr, _
                           "Rescue Success", Me.oPlayer.ObjectID, GetDateAsNumber(Now), False, Me.oPlayer.sPlayerNameProper, Nothing)
                        If oPC Is Nothing = False Then
                            Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                        End If
                        oPC = Nothing

                        If oCaptor Is Nothing = False Then
                            oPC = oCaptor.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, BytesToString(oAgent.AgentName) & _
                             " has been rescued by an agent team from your prison facilities. The whereabouts of the fugitive or the team that rescued the fugitive are unknown.", _
                             "Prisoner Escape", oCaptor.ObjectID, GetDateAsNumber(Now), False, oCaptor.sPlayerNameProper, Nothing)
                            If oPC Is Nothing = False Then
                                oCaptor.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                            End If
                        End If

                        'a benefit for completing this mission is that all agents receive a bonus to Loyalty
                        oAgent.oOwner.AdjustAllAgentLoyalty(10)

                        oAgent.ReturnHome()
                    End If
                End If
            Case eMissionResult.eShowQueueAndContents

                'NOTE: Can be ANY!!!

                'check our infiltration
                Select Case CType(oMission.lInfiltrationType, eInfiltrationType)
                    Case eInfiltrationType.eProductionInfiltration
                        'TODO: shows the current queue
                    Case eInfiltrationType.eResearchInfiltration
                        'TODO: shows the current research project
                    Case eInfiltrationType.eTradeInfiltration
                        'TODO: shows the current trade items for the tradepost (buy/sell orders as well as direct trades)
                    Case eInfiltrationType.eColonialInfiltration
                        'TODO: assess colonial morale
                    Case eInfiltrationType.eAgencyInfiltration
                        'TODO: shows the Agent List
                End Select
            Case eMissionResult.eSiphonItem
                Select Case CType(oMission.lInfiltrationType, eInfiltrationType)
                    Case eInfiltrationType.eTradeInfiltration
                        'TODO: Siphon trade profits
                        'Dim oFac as Facility = GetFac()
                        'oFac.AddAgentEffect(glCurrentCycle, lDuration, AgentEffectType.eSiphonTradepost, lAmount, True)
                    Case eInfiltrationType.eMiningInfiltration
                        'TODO: siphon Mining production (some is transported automatically)?
                        'Dim oFac as Facility = GetFac()
                        'oFac.AddAgentEffect(glCurrentCycle, lDuration, AgentEffectType.eSiphonMining, lAmount, True)
                    Case eInfiltrationType.eCorporationInfiltration
                        'TODO: siphon income of that corporation
                End Select
            Case eMissionResult.eSlowProduction
                Dim oFac As Facility = Nothing
                If lTargetID < 1 Then
                    Try
                        oFac = CType(oTarget.GetRandomUnitOrFacility(1, 0), Facility)
                    Catch
                    End Try
                Else : oFac = GetEpicaFacility(lTargetID)
                End If
                If oFac Is Nothing Then Return

                Dim lDuration As Int32 = CInt(Rnd() * 168) + 24       'in hours
                lDuration *= 108000     'convert to cycles
                Dim oEffect As AgentEffect = oFac.AddAgentEffect(glCurrentCycle, lDuration, AgentEffectType.eProduction, oMission.BaseEffect, True, oPlayer.ObjectID)
                If oTarget Is Nothing = False Then
                    If oEffect Is Nothing = False Then oPlayer.AddImposedEffect(oEffect, oTarget, BytesToString(oFac.EntityName))
                    If oMission.lInfiltrationType = eInfiltrationType.ePowerCenterInfiltration Then
                        oFac.ParentColony.UpdatePowerNeeds(-1)
                    End If
                End If
            Case eMissionResult.eSpecialProjectList
                Dim bFoundTech As Boolean = False
                Dim oSB As New System.Text.StringBuilder
                oSB.AppendLine("Your agents have finished assessing the target's Special Projects.")
                oSB.AppendLine("The following has also been uploaded to your Player Intel database.")
                If oTarget Is Nothing = False Then
                    For X As Int32 = 0 To oTarget.oSpecials.mlSpecialTechUB
                        If oTarget.oSpecials.mlSpecialTechIdx(X) <> -1 Then
                            If oTarget.oSpecials.moSpecialTech(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched OrElse (oTarget.oSpecials.moSpecialTech(X).ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eComponentResearching AndAlso oTarget.oSpecials.moSpecialTech(X).GetResearcherCount > 0) Then
                                Dim oPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(oPlayer, CType(oTarget.oSpecials.moSpecialTech(X), Epica_Tech), PlayerTechKnowledge.KnowledgeType.eNameOnly, False)
                                If oPTK Is Nothing = False Then
                                    oPTK.SendMsgToPlayer()
                                End If
                                bFoundTech = True
                                oSB.AppendLine("     " & BytesToString(oTarget.oSpecials.moSpecialTech(X).oTech.TechName))
                            End If
                        End If
                    Next X
                    If bFoundTech Then
                        Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Get Special Projects Mission Result", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                        If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    End If
                End If
            Case eMissionResult.eStealCargo
                Dim oFac As Facility = Nothing
                If lTargetID < 1 Then
                    Try
                        oFac = CType(oTarget.GetRandomUnitOrFacility(1, 0), Facility)
                    Catch
                    End Try
                Else : oFac = GetEpicaFacility(lTargetID)
                End If
                If oFac Is Nothing = False Then
                    If iTargetTypeID2 = ObjectType.eUnit Then
                        'hangar
                        For X As Int32 = 0 To oFac.lHangarUB
                            If oFac.lHangarIdx(X) = lTargetID2 Then
                                'TODO: Remove this item from the hangar and place it somewhere??? Where do we place it?
                                Exit For
                            End If
                        Next X
                    ElseIf (iTargetTypeID2 = ObjectType.eMineralCache OrElse iTargetTypeID2 = ObjectType.eMineral) AndAlso Colony.ProductionTypeSharesColonyCargo(oFac.yProductionType) = True AndAlso oFac.ParentColony Is Nothing = False Then
                        If lTargetID2 = -1 Then
                            lTargetID2 = oTarget.GetRandomKnownMineral()
                        End If
                        Dim oMineral As Mineral = Nothing
                        If lTargetID2 > -1 Then
                            oMineral = GetEpicaMineral(lTargetID2)
                            If oMineral Is Nothing Then
                                Dim oCache As MineralCache = GetEpicaMineralCache(lTargetID2)
                                If oCache Is Nothing = False Then
                                    If oCache.oMineral Is Nothing = False AndAlso oCache.ParentObject Is Nothing = False Then
                                        With CType(oCache.ParentObject, Epica_GUID)
                                            If .ObjTypeID = ObjectType.eFacility Then
                                                If CType(oCache.ParentObject, Facility).Owner.ObjectID <> oFac.Owner.ObjectID Then
                                                    LogEvent(LogEventType.PossibleCheat, "HandleMissionCompletion: Target Cache did not belong to player. " & oPlayer.ObjectID)
                                                Else
                                                    oMineral = oCache.oMineral
                                                End If
                                            ElseIf .ObjTypeID = ObjectType.eColony Then
                                                If CType(oCache.ParentObject, Colony).Owner.ObjectID <> oFac.Owner.ObjectID Then
                                                    LogEvent(LogEventType.PossibleCheat, "HandleMissionCompletion: Target cache does not belong to player. " & oPlayer.ObjectID)
                                                Else
                                                    oMineral = oCache.oMineral
                                                End If
                                            End If
                                        End With
                                    End If
                                End If
                            End If
                        End If
                        If oMineral Is Nothing = False Then
                            'pull mineral cache from the colony
                            Dim lPerFacCargoUsed As Int32 = oFac.ParentColony.GetPerFacilityCargoUsed(oFac.EntityDef.Cargo_Cap)
                            Dim lTotalMineralOnHand As Int32 = oFac.ParentColony.ColonyMineralQuantity(oMineral.ObjectID)
                            Dim lTotalCargoUsed As Int32 = oFac.ParentColony.TotalCargoCapUsed()
                            Dim lAmt As Int32 = CInt(lPerFacCargoUsed * (lTotalMineralOnHand / Math.Max(1, lTotalCargoUsed)))
                            If lAmt > 0 Then
                                oFac.ParentColony.AdjustColonyMineralCache(oMineral.ObjectID, -lAmt)
                                'Ok, find a colony with that amount available...
                                Dim oColony As Colony = oPlayer.GetColonyWithRoom(lAmt)
                                If oColony Is Nothing = False Then
                                    oColony.Owner.CheckFirstContactWithMineral(oMineral.ObjectID)
                                    oColony.AdjustColonyMineralCache(oMineral.ObjectID, lAmt)
                                    If (oPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                                        Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "Your agents successfully stole " & lAmt.ToString("#,##0") & " units of " & BytesToString(oMineral.MineralName) & " from " & oTarget.sPlayerNameProper & ". The cargo has been delivered to the " & BytesToString(oColony.ColonyName) & " colony.", "Steal Cargo Success", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                                        If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                                    End If
                                Else
                                    Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "Your agents were in the process to steal the target cargo but could not do so because you do not have a colony with enough cargo space.", "Mission Failed", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerName, Nothing)
                                    If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                                End If
                            End If
                        Else
                            Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, "Your agents were in the process of stealing cargo but could not find any that matched your criteria.", "Mission Failed", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerName, Nothing)
                            If oPC Is Nothing = False Then oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                        End If
                    Else
                        'cargo
                        If Colony.ProductionTypeSharesColonyCargo(oFac.yProductionType) = True Then
                        Else
                        End If
                    End If
                End If
            Case eMissionResult.eTutorialFindFactory
                'SpawnPirateFactory(Me.lTargetID)

                Dim oPC As PlayerComm = Me.oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                   "We have found the location of the pirate factory! Your agent was also successful in causing a malfunction in the factory's machinery which resulted in an explosion doing massive damage to the factory. The waypoint is attached to this email. For the honor of your empire!", _
                   "Pirate Factory", Me.oPlayer.ObjectID, GetDateAsNumber(Now), False, Me.oPlayer.sPlayerNameProper, Nothing)
                If oPC Is Nothing = False Then
                    oPC.AddEmailAttachment(1, Me.lTargetID, ObjectType.ePlanet, -15000, -16600, "Factory Loc")
                    Me.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                End If
                oPC = Nothing
        End Select

        bIsSuccess = True
	End Sub

	Public Sub MarkAgentsOnAMission()
		For X As Int32 = 0 To Me.oMission.GoalUB
			If oMission.MethodIDs(X) = Me.lMethodID Then
				For Y As Int32 = 0 To oMissionGoals(X).lAssignmentUB
					If oMissionGoals(X).oAssignments(Y) Is Nothing = False Then
						With oMissionGoals(X).oAssignments(Y)
							If .oAgent Is Nothing = False Then
								.oAgent.lAgentStatus = .oAgent.lAgentStatus Or AgentStatus.OnAMission
								.oAgent.oMission = Me

								If .oAgent.InfiltrationType <> Me.oMission.lInfiltrationType Then
									Dim lTemp As Int32 = .oAgent.Suspicion
									lTemp += 5
									If lTemp > 255 Then lTemp = 255
									.oAgent.Suspicion = CByte(lTemp)
									goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.ReduceSuspicion, .oAgent, Nothing, glCurrentCycle + Agent.ml_DEINFILTRATE_TIME)
								End If
							End If
						End With
					End If
				Next Y
			End If
		Next X
	End Sub

    Public Sub MarkAgentsOffMission()
        Try
            For X As Int32 = 0 To Me.oMission.GoalUB
                If oMission.MethodIDs(X) = Me.lMethodID Then
                    For Y As Int32 = 0 To oMissionGoals(X).lAssignmentUB
                        If oMissionGoals(X).oAssignments(Y) Is Nothing = False Then
                            With oMissionGoals(X).oAssignments(Y)
                                If .oAgent Is Nothing = False Then
                                    '2010-05-28: Fix exploit that allows you to queue up 3 missions.  1 exeuting, 2 waiting.  Abort #2 and this sets agents idle, popping #3.
                                    If .oAgent.oMission Is Nothing = True OrElse Me.PM_ID = .oAgent.oMission.PM_ID Then
                                        If (.oAgent.lAgentStatus And AgentStatus.OnAMission) <> 0 Then
                                            .oAgent.lAgentStatus = .oAgent.lAgentStatus Xor AgentStatus.OnAMission
                                            .oAgent.oMission = Nothing
                                        End If
                                    End If
                                End If
                            End With
                        End If
                    Next Y
                End If
            Next X
        Catch ex As Exception
            LogEvent(LogEventType.Warning, "PlayerMission.MarkAgentsOffMission: " & ex.Message)
        End Try
    End Sub

    Public Sub HandleMissionFailure()
        Me.lCurrentPhase = eMissionPhase.eReinfiltrationPhase
        ProcessTests()
        Me.lCurrentPhase = eMissionPhase.eMissionOverFailure
        bIsSuccess = False
        MarkAgentsOffMission()
    End Sub
 

	Public Sub AddPhaseCoverAgent(ByVal lPhase As Int32, ByRef oCoverAgent As Agent, ByVal yUsedAsCover As Byte)
		If oCoverAgent Is Nothing Then Return
		If lPhase > lPhaseUB Then Return
		If lPhase < 0 Then Return
		oPhases(lPhase).AddCoverAgent(oCoverAgent, yUsedAsCover)
	End Sub

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If PM_ID = -1 Then
                'INSERT
				sSQL = "INSERT INTO tblPlayerMission (MissionID, PlayerID, CasualtyCnt, TargetPlayerID, TargetID, TargetTypeID, Modifier, " & _
				  "yAlarmThrown, SafeHouseSetting, CurrentPhase, TargetID2, TargetTypeID2, MethodID, bArchived) VALUES (" & oMission.ObjectID & ", " & _
				  Me.oPlayer.ObjectID & ", " & CasualtyCnt & ", "
				If Me.oTarget Is Nothing = False Then
					sSQL &= Me.oTarget.ObjectID
				Else : sSQL &= "-1"
				End If
                sSQL &= ", " & lTargetID & ", " & iTargetTypeID & _
                 ", " & Modifier & ", " & CByte(bAlarmThrown) & ", " & ySafeHouseSetting & ", " & CByte(lCurrentPhase) & _
                 ", " & lTargetID2 & ", " & iTargetTypeID2 & ", " & lMethodID & ", " & yArchived & ")"
            Else
                'UPDATE
				sSQL = "UPDATE tblPlayerMission SET MissionID = " & oMission.ObjectID & ", PlayerID = " & Me.oPlayer.ObjectID & _
				  ", CasualtyCnt = " & CasualtyCnt & ", TargetPlayerID = "
				If Me.oTarget Is Nothing = False Then
					sSQL &= Me.oTarget.ObjectID
				Else : sSQL &= "-1"
				End If
                sSQL &= ", TargetID = " & lTargetID & ", TargetTypeID = " & iTargetTypeID & ", Modifier = " & Modifier & _
                 ", yAlarmThrown = " & CByte(bAlarmThrown) & ", SafeHouseSetting = " & ySafeHouseSetting & ", CurrentPhase = " & _
                 CByte(lCurrentPhase) & ", TargetID2 = " & lTargetID2 & ", TargetTypeID2 = " & iTargetTypeID2 & ", MethodID = " & _
                 lMethodID & ", bArchived = " & yArchived & " WHERE PM_ID = " & PM_ID
			End If
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			If oComm.ExecuteNonQuery() = 0 Then
				Err.Raise(-1, "SaveObject", "No records affected when saving object!")
			End If
			If PM_ID = -1 Then
				Dim oData As OleDb.OleDbDataReader
				oComm = Nothing
				sSQL = "SELECT MAX(pm_id) FROM tblPlayerMission WHERE MissionID = " & oMission.ObjectID & " AND PlayerID = " & Me.oPlayer.ObjectID
				oComm = New OleDb.OleDbCommand(sSQL, goCN)
				oData = oComm.ExecuteReader(CommandBehavior.Default)
				If oData.Read Then
					PM_ID = CInt(oData(0))
				End If
				oData.Close()
				oData = Nothing
			End If

			'oMissionGoals
			sSQL = "DELETE FROM tblPlayerMissionGoal WHERE PM_ID = " & Me.PM_ID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm.Dispose()
            oComm = Nothing

            If Me.ySafeHouseSetting > 0 Then
                If Me.oSafeHouseGoal Is Nothing = False Then
                    Me.oSafeHouseGoal.SaveObject()
                End If
            End If

			If oMissionGoals Is Nothing = False Then
				For X As Int32 = 0 To oMissionGoals.GetUpperBound(0)
					If oMissionGoals(X) Is Nothing = False Then oMissionGoals(X).SaveObject()
				Next X
			End If

			'oPhases
			If oPhases Is Nothing = False Then
				For X As Int32 = 0 To oPhases.GetUpperBound(0)
					If oPhases(X) Is Nothing = False Then oPhases(X).SaveObject()
				Next X
			End If

			bResult = True
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object PlayerMission " & PM_ID & ". Reason: " & Err.Description)
		Finally
			oComm = Nothing
		End Try
        Return bResult
	End Function

	Public Function DataIsValid(ByVal bExact As Boolean) As Int32
		Try
			'is omission set?
			If oMission Is Nothing Then Return ePM_ErrorCode.eNoMissionSet

			If bExact = True Then
				If lMethodID < 1 Then Return ePM_ErrorCode.eInvalidMethod

				Dim bGood As Boolean = False
				For X As Int32 = 0 To oMission.GoalUB
					If lMethodID = oMission.MethodIDs(X) Then
						bGood = True
						Exit For
					End If
				Next X
				If bGood = False Then Return ePM_ErrorCode.eInvalidMethod

				bGood = False
				'check our target and target types
                Select Case CType(oMission.ProgramControlID, eMissionResult)
                    Case eMissionResult.eAudit
                        oTarget = oPlayer
                    Case eMissionResult.eShowQueueAndContents, eMissionResult.eSlowProduction, eMissionResult.eSabotageProduction, eMissionResult.eCorruptProduction, eMissionResult.eDestroyFacility
                        'player, facility
                        If oTarget Is Nothing = False AndAlso oPlayer.GetPlayerRel(oTarget.ObjectID) Is Nothing = False Then
                            'ok, check the facility
                            'can be an any target
                            If lTargetID > 0 AndAlso iTargetTypeID = ObjectType.eFacility Then
                                Dim oFac As Facility = GetEpicaFacility(lTargetID)
                                If oFac Is Nothing OrElse oFac.Owner.ObjectID <> oTarget.ObjectID Then Return ePM_ErrorCode.eInvalidTargetSelection
                                'Else : Return ePM_ErrorCode.eInvalidTargetSelection
                            End If
                        Else : Return ePM_ErrorCode.eInvalidTargetSelection
                        End If
                    Case eMissionResult.eClutterCargoBay, eMissionResult.eIncreasePowerNeeds
                        If oTarget Is Nothing = False AndAlso oPlayer.GetPlayerRel(oTarget.ObjectID) Is Nothing = False Then
                            'ok, check the colony
                            If lTargetID > 0 Then
                                Dim oColony As Colony = Nothing
                                If iTargetTypeID = ObjectType.eColony Then oColony = GetEpicaColony(lTargetID)
                                If oColony Is Nothing OrElse oColony.Owner.ObjectID <> oTarget.ObjectID Then Return ePM_ErrorCode.eInvalidTargetSelection
                                If oPlayer.HasItemIntel(lTargetID, iTargetTypeID, oTarget.ObjectID) = False Then
                                    LogEvent(LogEventType.PossibleCheat, "Player does not have intel on targeted colony! Player: " & oPlayer.ObjectID)
                                    Return ePM_ErrorCode.eInvalidTargetSelection
                                End If
                            Else : Return ePM_ErrorCode.eInvalidTargetSelection
                            End If
                        Else : Return ePM_ErrorCode.eInvalidTargetSelection
                        End If
                    Case eMissionResult.eImpedeCurrentDevelopment
                        If oTarget Is Nothing = False AndAlso oPlayer.GetPlayerRel(oTarget.ObjectID) Is Nothing = False Then
                            'Impede nolonger requires a Target Tech.
                            Return ePM_ErrorCode.eNoErrorsFound
                            'If iTargetTypeID = ObjectType.eSpecialTech Then
                            '    Dim oTech As Epica_Tech = oTarget.GetTech(lTargetID, iTargetTypeID)
                            '    If oTech Is Nothing OrElse oTech.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                            '        Return ePM_ErrorCode.eInvalidTargetSelection
                            '    End If
                            'Else : Return ePM_ErrorCode.eInvalidTargetSelection
                            'End If
                        Else : Return ePM_ErrorCode.eInvalidTargetSelection
                        End If
                    Case eMissionResult.eDestroyCurrentSpecialProject
                        If oTarget Is Nothing = False AndAlso oPlayer.GetPlayerRel(oTarget.ObjectID) Is Nothing = False Then
                            If iTargetTypeID = ObjectType.eSpecialTech Then
                                Dim oTech As Epica_Tech = oTarget.GetTech(lTargetID, iTargetTypeID)
                                If oTech Is Nothing OrElse oTech.ComponentDevelopmentPhase = Epica_Tech.eComponentDevelopmentPhase.eResearched Then
                                    Return ePM_ErrorCode.eInvalidTargetSelection
                                End If
                            Else : Return ePM_ErrorCode.eInvalidTargetSelection
                            End If
                        Else : Return ePM_ErrorCode.eInvalidTargetSelection
                        End If
                    Case eMissionResult.eDoorJamHangar, eMissionResult.eStealCargo
                            'Player, unit or facility
                            If oTarget Is Nothing = False AndAlso oPlayer.GetPlayerRel(oTarget.ObjectID) Is Nothing = False Then
                                'ok, check the facility/unit
                                If lTargetID > 0 Then
                                    Dim oEntity As Epica_Entity = Nothing
                                    If iTargetTypeID = ObjectType.eFacility Then
                                        oEntity = GetEpicaFacility(lTargetID)
                                    ElseIf iTargetTypeID = ObjectType.eUnit Then
                                        oEntity = GetEpicaUnit(lTargetID)
                                    End If
                                    If oEntity Is Nothing OrElse oEntity.Owner.ObjectID <> oTarget.ObjectID Then Return ePM_ErrorCode.eInvalidTargetSelection
                                    If oPlayer.HasItemIntel(lTargetID, iTargetTypeID, oTarget.ObjectID) = False Then
                                        LogEvent(LogEventType.PossibleCheat, "Player does not have intel on targeted facility! Player: " & oPlayer.ObjectID)
                                        Return ePM_ErrorCode.eInvalidTargetSelection
                                    End If
                                    'Else : Return ePM_ErrorCode.eInvalidTargetSelection
                                End If
                            Else : Return ePM_ErrorCode.eInvalidTargetSelection
                            End If
                    Case eMissionResult.eAcquireTechData
                            'player, component
                            If oTarget Is Nothing = False AndAlso oPlayer.GetPlayerRel(oTarget.ObjectID) Is Nothing = False Then
                                'ok, check the component
                                If lTargetID > 0 AndAlso iTargetTypeID > 0 Then
                                    Dim oTech As Epica_Tech = oTarget.GetTech(lTargetID, iTargetTypeID)
                                    If oTech Is Nothing Then Return ePM_ErrorCode.eInvalidTargetSelection
                                    If oPlayer.HasTechKnowledge(oTech.ObjectID, oTech.ObjTypeID, PlayerTechKnowledge.KnowledgeType.eNameOnly) = False Then
                                        LogEvent(LogEventType.PossibleCheat, "PlayerMission.DataIsValid() Player does not have tech knowledge of target tech! Player: " & oPlayer.ObjectID)
                                        Return ePM_ErrorCode.eInvalidTargetSelection
                                    End If
                                End If
                            Else : Return ePM_ErrorCode.eInvalidTargetSelection
                            End If
                    Case eMissionResult.eGetFacilityList
                            'Player, productiontype
                            If oTarget Is Nothing OrElse oPlayer.GetPlayerRel(oTarget.ObjectID) Is Nothing Then
                                Return ePM_ErrorCode.eInvalidTargetSelection
                            End If
                            If lTargetID2 > 0 Then
                                If iTargetTypeID2 = ObjectType.ePlanet Then
                                    Dim oPlanet As Planet = GetEpicaPlanet(lTargetID2)
                                    If oPlanet Is Nothing Then Return ePM_ErrorCode.eInvalidTargetSelection
                                ElseIf iTargetTypeID2 = ObjectType.eSolarSystem Then
                                    Dim oSystem As SolarSystem = GetEpicaSystem(lTargetID2)
                                    If oSystem Is Nothing Then Return ePM_ErrorCode.eInvalidTargetSelection
                                Else : Return ePM_ErrorCode.eInvalidTargetSelection
                                End If
                            Else : Return ePM_ErrorCode.eInvalidTargetSelection
                            End If
                    Case eMissionResult.eSiphonItem
                            'Player, SiphonType
                            If oTarget Is Nothing OrElse oPlayer.GetPlayerRel(oTarget.ObjectID) Is Nothing Then
                                Return ePM_ErrorCode.eInvalidTargetSelection
                            End If
                    Case eMissionResult.ePlantEvidence
                            'player, otherplayerid, pm_id
                            If oTarget Is Nothing OrElse oPlayer.GetPlayerRel(oTarget.ObjectID) Is Nothing Then
                                Return ePM_ErrorCode.eInvalidTargetSelection
                            Else
                                If iTargetTypeID <> ObjectType.ePlayer OrElse oTarget.GetPlayerRel(lTargetID) Is Nothing Then Return ePM_ErrorCode.eInvalidTargetSelection
                                Dim oOtherPlayer As Player = GetEpicaPlayer(lTargetID)
                                If oOtherPlayer Is Nothing Then Return ePM_ErrorCode.eInvalidTargetSelection
                                Dim oPM As PlayerMission = GetEpicaPlayerMission(lTargetID2)
                                If oPM Is Nothing OrElse oPM.oPlayer.ObjectID <> oPlayer.ObjectID Then Return ePM_ErrorCode.eInvalidTargetSelection
                            End If
                    Case eMissionResult.eAssassinateAgent, eMissionResult.eCaptureAgent
                            'player, agentid (must be agent of target)
                            If oTarget Is Nothing OrElse oPlayer.GetPlayerRel(oTarget.ObjectID) Is Nothing Then Return ePM_ErrorCode.eInvalidTargetSelection

                            If lTargetID > 0 AndAlso iTargetTypeID > 0 Then
                                If iTargetTypeID <> ObjectType.eAgent Then Return ePM_ErrorCode.eInvalidTargetSelection
                                If oPlayer.HasItemIntel(lTargetID, iTargetTypeID, oTarget.ObjectID) = False Then
                                    LogEvent(LogEventType.PossibleCheat, "Player does not have intel on targeted agent! Player: " & oPlayer.ObjectID)
                                    Return ePM_ErrorCode.eInvalidTargetSelection
                                End If
                            End If
                    Case eMissionResult.eFindMineral
                            'No player, mineral
                            If oPlayer.IsMineralDiscovered(lTargetID) = False Then
                                LogEvent(LogEventType.PossibleCheat, "Player does not have knowledge of mineral! Player: " & oPlayer.ObjectID)
                                Return ePM_ErrorCode.eInvalidTargetSelection
                            End If
                    Case eMissionResult.eGetColonyBudget
                            If oTarget Is Nothing OrElse oPlayer.GetPlayerRel(oTarget.ObjectID) Is Nothing Then Return ePM_ErrorCode.eInvalidTargetSelection
                            If lTargetID > 0 AndAlso iTargetTypeID = ObjectType.ePlanet Then
                                Dim oPlanet As Planet = GetEpicaPlanet(lTargetID)
                                If oPlanet Is Nothing Then Return ePM_ErrorCode.eInvalidTargetSelection
                            Else : Return ePM_ErrorCode.eInvalidTargetSelection
                            End If
                    Case eMissionResult.eSearchAndRescueAgent
                            'no player, agent
                    Case eMissionResult.eReconPlanetMap
                            'no player, locx, locz
                    Case eMissionResult.eGeologicalSurvey, eMissionResult.eTutorialFindFactory
                            'no player, planet
                    Case eMissionResult.eStealCargo
                            'player, target fac/unit/any, target
                            If oTarget Is Nothing OrElse oPlayer.GetPlayerRel(oTarget.ObjectID) Is Nothing Then
                                Return ePM_ErrorCode.eInvalidTargetSelection
                            End If

                    Case Else
                            'Player only
                            If oTarget Is Nothing OrElse oPlayer.GetPlayerRel(oTarget.ObjectID) Is Nothing Then
                                Return ePM_ErrorCode.eInvalidTargetSelection
                            End If
                End Select

                'all assignments must exist...
                If Me.ySafeHouseSetting > 0 Then
                    If Me.oSafeHouseGoal Is Nothing Then Return ePM_ErrorCode.eMissingSkillset
                    Dim lResp As ePM_ErrorCode = ValidatePlayerMissionGoal(Me.oSafeHouseGoal)
                    If lResp <> ePM_ErrorCode.eNoErrorsFound Then Return lResp
                End If
				For X As Int32 = 0 To oMission.GoalUB
					If oMission.MethodIDs(X) = lMethodID Then
                        Dim lResp As ePM_ErrorCode = ValidatePlayerMissionGoal(oMissionGoals(X))
                        If lResp <> ePM_ErrorCode.eNoErrorsFound Then Return lResp
                    End If
				Next X
			End If
		Catch ex As Exception
			LogEvent(LogEventType.Warning, "PlayerMission.DataIsValid: " & ex.Message)
			Return ePM_ErrorCode.eUnknownError
		End Try
		Return ePM_ErrorCode.eNoErrorsFound
    End Function

    Private Function ValidatePlayerMissionGoal(ByRef oPMG As PlayerMissionGoal) As ePM_ErrorCode
        If oPMG.oSkillSet Is Nothing Then Return ePM_ErrorCode.eMissingSkillset
        'omissiongoals(x).oSkillSet.SkillUB
        Dim bSkillGood(oPMG.oSkillSet.SkillUB) As Boolean
        For Y As Int32 = 0 To oPMG.oSkillSet.SkillUB
            bSkillGood(Y) = False
        Next Y

        Dim bAgentDetained As Boolean = False
        For Y As Int32 = 0 To oPMG.lAssignmentUB
            Dim lSkillID As Int32 = oPMG.oAssignments(Y).oSkill.ObjectID
            If oPMG.oAssignments(Y).oAgent Is Nothing Then Return ePM_ErrorCode.eMissingAssignmentAgent
            If (oPMG.oAssignments(Y).oAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 OrElse (oPMG.oAssignments(Y).oAgent.lAgentStatus And AgentStatus.IsDead) <> 0 Then
                bAgentDetained = True
            Else
                For Z As Int32 = 0 To oPMG.oSkillSet.SkillUB
                    If oPMG.oSkillSet.Skills(Z).oSkill.ObjectID = lSkillID Then
                        bSkillGood(Z) = True
                        Exit For
                    End If
                Next Z

                If oPMG.oAssignments(Y).oAgent.GetSkillValue(lSkillID, False, 0) = 0 Then
                    'ok, find any other assignments in this goal that use this agent...
                    For Z As Int32 = 0 To oPMG.lAssignmentUB
                        If Y <> Z Then
                            If oPMG.oAssignments(Z) Is Nothing = False AndAlso oPMG.oAssignments(Z).oAgent Is Nothing = False Then
                                If oPMG.oAssignments(Z).oAgent.ObjectID = oPMG.oAssignments(Y).oAgent.ObjectID Then
                                    Dim lTmpSkill As Int32 = oPMG.oAssignments(Z).oSkill.ObjectID
                                    If oPMG.oAssignments(Z).oAgent.GetSkillValue(lTmpSkill, True, 0) = 0 Then
                                        Return ePM_ErrorCode.eAssignedAgentUnavail
                                    End If
                                End If
                            End If
                        End If
                    Next Z
                End If
            End If
        Next Y

        'Now, verify our goods
        For Y As Int32 = 0 To oPMG.oSkillSet.SkillUB
            If bSkillGood(Y) = False Then
                If bAgentDetained = False Then Return ePM_ErrorCode.eMissingAssignment Else Return ePM_ErrorCode.eAssignedAgentUnavail
            End If
        Next Y
        Return ePM_ErrorCode.eNoErrorsFound
    End Function

	Public Function AgentsReadyToExecuteMission() As Boolean

		Dim lInfTarget As Int32 = -1
		Dim iInfTargetTypeID As Int16 = -1
		Select Case CType(oMission.ProgramControlID, eMissionResult)
            Case eMissionResult.eFindMineral, eMissionResult.eReconPlanetMap
                lInfTarget = lTargetID2 : iInfTargetTypeID = iTargetTypeID2
			Case eMissionResult.eGeologicalSurvey
				lInfTarget = lTargetID : iInfTargetTypeID = iTargetTypeID
			Case eMissionResult.eTutorialFindFactory
                lInfTarget = lTargetID : iInfTargetTypeID = ObjectType.ePlanet
            Case eMissionResult.eSearchAndRescueAgent
                Dim oTmpAgent As Agent = GetEpicaAgent(Me.lTargetID)
                If oTmpAgent Is Nothing = False AndAlso (oTmpAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then
                    lInfTarget = oTmpAgent.lCapturedBy
                    iInfTargetTypeID = ObjectType.ePlayer
                End If
			Case Else
				'Player....
				If oTarget Is Nothing = False Then
					lInfTarget = oTarget.ObjectID : iInfTargetTypeID = ObjectType.ePlayer
				End If
		End Select

		If lInfTarget = -1 OrElse iInfTargetTypeID = -1 Then
			If Me.lCurrentPhase = eMissionPhase.eWaitingToExecute Then Me.lCurrentPhase = eMissionPhase.eInPlanning
			Return False
		End If

		For X As Int32 = 0 To oMission.GoalUB
			If oMission.MethodIDs(X) = lMethodID Then
				If oMissionGoals(X).oSkillSet Is Nothing Then
					If Me.lCurrentPhase = eMissionPhase.eWaitingToExecute Then Me.lCurrentPhase = eMissionPhase.eInPlanning
					Return False
				End If

				For Y As Int32 = 0 To oMissionGoals(X).lAssignmentUB
					Dim lSkillID As Int32 = oMissionGoals(X).oAssignments(Y).oSkill.ObjectID

					Dim oAgent As Agent = oMissionGoals(X).oAssignments(Y).oAgent

					If oAgent Is Nothing Then
						If Me.lCurrentPhase = eMissionPhase.eWaitingToExecute Then Me.lCurrentPhase = eMissionPhase.eInPlanning
						Return False
					End If
					If (oAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 OrElse (oAgent.lAgentStatus And AgentStatus.IsDead) <> 0 Then
						If Me.lCurrentPhase = eMissionPhase.eWaitingToExecute Then Me.lCurrentPhase = eMissionPhase.eInPlanning
						Return False
                    ElseIf (oAgent.lAgentStatus And AgentStatus.OnAMission) <> 0 AndAlso (oAgent.oMission Is Nothing = False AndAlso oAgent.oMission.PM_ID <> Me.PM_ID) Then
                        '(oAgent.oMission.lCurrentPhase And (eMissionPhase.eCancelled Or eMissionPhase.eMissionOverFailure Or eMissionPhase.eMissionOverSuccess Or eMissionPhase.eMissionPaused)) = 0) Then
                        Return False
					ElseIf (oAgent.lAgentStatus And AgentStatus.IsInfiltrated) = 0 Then
						Return False
					ElseIf oAgent.lTargetID <> lInfTarget OrElse oAgent.iTargetTypeID <> iInfTargetTypeID Then
						Return False
					End If
				Next Y
			End If
		Next X

		Return True
	End Function

	Public Sub RemoveAgentFromAssignments(ByVal lAgentID As Int32)
		If Me.oMission Is Nothing = False Then
			For X As Int32 = 0 To oMission.GoalUB
				If oMission.MethodIDs(X) = lMethodID Then
					For Y As Int32 = 0 To oMissionGoals(X).lAssignmentUB
						If oMissionGoals(X).oAssignments(Y) Is Nothing = False Then
							If oMissionGoals(X).oAssignments(Y).oAgent Is Nothing = False Then
								If oMissionGoals(X).oAssignments(Y).oAgent.ObjectID = lAgentID Then
									oMissionGoals(X).oAssignments(Y).oAgent = Nothing
									'TODO: is the mission still possible?
								End If
							End If
							If oMissionGoals(X).oAssignments(Y).oCoveringAgent Is Nothing = False Then
								If oMissionGoals(X).oAssignments(Y).oCoveringAgent.ObjectID = lAgentID Then
									oMissionGoals(X).oAssignments(Y).oCoveringAgent = Nothing
									'TODO: is the mission still possible?
								End If
							End If
						End If
					Next
				End If
			Next X
		End If
	End Sub

	Public Function GetAddObjectMessage() As Byte()
		Dim yPhaseCnt As Byte = 0
		Dim yGoalCnt As Byte = 0
		Dim lPhaseCvrCnt() As Int32 = Nothing
		Dim lAssignCnt() As Int32 = Nothing

		'determine our counts now
		ReDim lPhaseCvrCnt(lPhaseUB)
		For X As Int32 = 0 To lPhaseUB
			lPhaseCvrCnt(X) = 0
			If oPhases(X) Is Nothing = False Then
				yPhaseCnt += CByte(1)
				For Y As Int32 = 0 To oPhases(X).lCoverAgentUB
					If oPhases(X).lCoverAgentIdx(Y) <> -1 Then lPhaseCvrCnt(X) += 1
				Next Y
			End If
		Next X
		If oMission Is Nothing = False Then
			ReDim lAssignCnt(oMission.GoalUB)
			For X As Int32 = 0 To oMission.GoalUB
				lAssignCnt(X) = 0
				If oMissionGoals(X) Is Nothing = False AndAlso oMission.MethodIDs(X) = lMethodID Then
					yGoalCnt += CByte(1)
					For Y As Int32 = 0 To oMissionGoals(X).lAssignmentUB
						If oMissionGoals(X).oAssignments(Y) Is Nothing = False Then lAssignCnt(X) += 1
					Next Y
				End If
			Next X
		Else
			ReDim lAssignCnt(-1)
		End If

		'Ok, we have everything, let's determine total size
		'33 + (yPhaseCnt * 2) + (lSumOfPhaseCover * 4) + (yGoalCnt * 9) + (lSumAgentAssignment * 8)
		Dim lSumOfPhaseCover As Int32 = 0
		Dim lSumAgentAssignment As Int32 = 0
		For X As Int32 = 0 To lPhaseUB
			lSumOfPhaseCover += lPhaseCvrCnt(X)
		Next X
		For X As Int32 = 0 To lAssignCnt.GetUpperBound(0)
			lSumAgentAssignment += lAssignCnt(X)
        Next X

        Dim lSafeHouseSize As Int32 = 0
        If ySafeHouseSetting > 0 Then
            lSafeHouseSize = 26
        End If

		Dim yMsg() As Byte
        ReDim yMsg(35 + lSafeHouseSize + (CInt(yPhaseCnt) * 2) + (lSumOfPhaseCover * 4) + (CInt(yGoalCnt) * 9) + (lSumAgentAssignment * 11))
		Dim lPos As Int32 = 0

		System.BitConverter.GetBytes(GlobalMessageCode.eSubmitMission).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(PM_ID).CopyTo(yMsg, lPos) : lPos += 4

        'If bArchived = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 0
        yMsg(lPos) = yArchived
		lPos += 1

		If oMission Is Nothing = False Then
			System.BitConverter.GetBytes(oMission.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
		Else
			System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
		End If
		If oTarget Is Nothing = False Then
			System.BitConverter.GetBytes(oTarget.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
		Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
		End If
		System.BitConverter.GetBytes(lTargetID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(iTargetTypeID).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(lTargetID2).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(iTargetTypeID2).CopyTo(yMsg, lPos) : lPos += 2
		yMsg(lPos) = ySafeHouseSetting : lPos += 1
		System.BitConverter.GetBytes(lMethodID).CopyTo(yMsg, lPos) : lPos += 4
		yMsg(lPos) = lCurrentPhase : lPos += 1
		If bAlarmThrown = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 0
        lPos += 1

        If ySafeHouseSetting > 0 Then
            If oSafeHouseGoal Is Nothing = False Then
                With oSafeHouseGoal
                    'skillsetid
                    If .oSkillSet Is Nothing = False Then
                        System.BitConverter.GetBytes(.oSkillSet.SkillSetID).CopyTo(yMsg, lPos) : lPos += 4
                    Else : lPos += 4
                    End If
                    If .lAssignmentUB > -1 Then
                        'assignment 1 agentid
                        If .oAssignments(0).oAgent Is Nothing = False Then
                            System.BitConverter.GetBytes(.oAssignments(0).oAgent.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                        Else
                            System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                        End If
                        'assignment 1 skillid
                        If .oAssignments(0).oSkill Is Nothing = False Then
                            System.BitConverter.GetBytes(.oAssignments(0).oSkill.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                        Else
                            System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                        End If
                        System.BitConverter.GetBytes(.oAssignments(0).PointsAccumulated).CopyTo(yMsg, lPos) : lPos += 2
                        yMsg(lPos) = .oAssignments(0).yStatus : lPos += 1

                        If .lAssignmentUB > 0 Then
                            'assignment 2 agentid
                            If .oAssignments(1).oAgent Is Nothing = False Then
                                System.BitConverter.GetBytes(.oAssignments(1).oAgent.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                            Else
                                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                            End If
                            'assignment 2 skillid
                            If .oAssignments(1).oSkill Is Nothing = False Then
                                System.BitConverter.GetBytes(.oAssignments(1).oSkill.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                            Else
                                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                            End If
                            System.BitConverter.GetBytes(.oAssignments(1).PointsAccumulated).CopyTo(yMsg, lPos) : lPos += 2
                            yMsg(lPos) = .oAssignments(1).yStatus : lPos += 1
                        Else : lPos += 11
                        End If
                    Else : lPos += 22
                    End If

                End With
            Else : lPos += 26
            End If
        End If

		'ok, get the phases
		yMsg(lPos) = yPhaseCnt : lPos += 1
		For X As Int32 = 0 To lPhaseUB
			If oPhases(X) Is Nothing = False Then
				yMsg(lPos) = oPhases(X).lPhase : lPos += 1
				yMsg(lPos) = CByte(lPhaseCvrCnt(X)) : lPos += 1

				For Y As Int32 = 0 To oPhases(X).lCoverAgentUB
					If oPhases(X).lCoverAgentIdx(Y) <> -1 Then
						System.BitConverter.GetBytes(oPhases(X).lCoverAgentIdx(Y)).CopyTo(yMsg, lPos) : lPos += 4
					End If
				Next Y
			End If
		Next X

		'now for our goals
		yMsg(lPos) = yGoalCnt : lPos += 1
		If oMission Is Nothing = False Then
			For X As Int32 = 0 To oMission.GoalUB
				If oMissionGoals(X) Is Nothing = False AndAlso oMission.MethodIDs(X) = lMethodID Then
					If oMissionGoals(X).oGoal Is Nothing = False Then
						System.BitConverter.GetBytes(oMissionGoals(X).oGoal.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
					Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
					End If
					If oMissionGoals(X).oSkillSet Is Nothing = False Then
						System.BitConverter.GetBytes(oMissionGoals(X).oSkillSet.SkillSetID).CopyTo(yMsg, lPos) : lPos += 4
					Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
					End If


					yMsg(lPos) = CByte(lAssignCnt(X)) : lPos += 1
					For Y As Int32 = 0 To oMissionGoals(X).lAssignmentUB
						If oMissionGoals(X).oAssignments(Y) Is Nothing = False Then
							If oMissionGoals(X).oAssignments(Y).oAgent Is Nothing = False Then
								System.BitConverter.GetBytes(oMissionGoals(X).oAssignments(Y).oAgent.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
							Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
							End If
							If oMissionGoals(X).oAssignments(Y).oSkill Is Nothing = False Then
								System.BitConverter.GetBytes(oMissionGoals(X).oAssignments(Y).oSkill.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
							Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
							End If
							System.BitConverter.GetBytes(oMissionGoals(X).oAssignments(Y).PointsAccumulated).CopyTo(yMsg, lPos) : lPos += 2
							yMsg(lPos) = oMissionGoals(X).oAssignments(Y).yStatus : lPos += 1
						End If
					Next Y
				End If
			Next X
		End If

		Return yMsg
	End Function

	Public Function GetPMUpdateMsg() As Byte()
		Dim yGoalCnt As Byte = 0
		Dim lAssignCnt() As Int32 = Nothing
		If oMission Is Nothing = False Then
			ReDim lAssignCnt(oMission.GoalUB)
			For X As Int32 = 0 To oMission.GoalUB
				lAssignCnt(X) = 0
				If oMissionGoals(X) Is Nothing = False AndAlso oMission.MethodIDs(X) = lMethodID Then
					yGoalCnt += CByte(1)
					For Y As Int32 = 0 To oMissionGoals(X).lAssignmentUB
						If oMissionGoals(X).oAssignments(Y) Is Nothing = False Then lAssignCnt(X) += 1
					Next Y
				End If
			Next X
		Else
			ReDim lAssignCnt(-1)
		End If

		'Ok, we have everything, let's determine total size
		Dim lSumAgentAssignment As Int32 = 0
		For X As Int32 = 0 To lAssignCnt.GetUpperBound(0)
			lSumAgentAssignment += lAssignCnt(X)
        Next X

        Dim lSafehousesize As Int32 = 0
        If ySafeHouseSetting > 0 Then
            lSafehousesize = 26
        End If

		Dim yResp() As Byte
		Dim lPos As Int32 = 0

        ReDim yResp(8 + lSafehousesize + (yGoalCnt * 5) + (lSumAgentAssignment * 11))

		System.BitConverter.GetBytes(GlobalMessageCode.eGetPMUpdate).CopyTo(yResp, lPos) : lPos += 2
		System.BitConverter.GetBytes(PM_ID).CopyTo(yResp, lPos) : lPos += 4
		Dim yValue As Byte = Me.ySafeHouseSetting
		If bAlarmThrown = True Then yValue = CByte(yValue Or 128)
		yResp(lPos) = yValue : lPos += 1
        yResp(lPos) = Me.lCurrentPhase : lPos += 1

        If ySafeHouseSetting > 0 Then
            If oSafeHouseGoal Is Nothing = False Then
                With oSafeHouseGoal
                    'skillsetid
                    If .oSkillSet Is Nothing = False Then
                        System.BitConverter.GetBytes(.oSkillSet.SkillSetID).CopyTo(yResp, lPos) : lPos += 4
                    Else : lPos += 4
                    End If
                    If .lAssignmentUB > -1 Then
                        'assignment 1 agentid
                        If .oAssignments(0).oAgent Is Nothing = False Then
                            System.BitConverter.GetBytes(.oAssignments(0).oAgent.ObjectID).CopyTo(yResp, lPos) : lPos += 4
                        Else
                            System.BitConverter.GetBytes(-1I).CopyTo(yResp, lPos) : lPos += 4
                        End If
                        'assignment 1 skillid
                        If .oAssignments(0).oSkill Is Nothing = False Then
                            System.BitConverter.GetBytes(.oAssignments(0).oSkill.ObjectID).CopyTo(yResp, lPos) : lPos += 4
                        Else
                            System.BitConverter.GetBytes(-1I).CopyTo(yResp, lPos) : lPos += 4
                        End If
                        System.BitConverter.GetBytes(.oAssignments(0).PointsAccumulated).CopyTo(yResp, lPos) : lPos += 2
                        yResp(lPos) = .oAssignments(0).yStatus : lPos += 1

                        If .lAssignmentUB > 0 Then
                            'assignment 2 agentid
                            If .oAssignments(1).oAgent Is Nothing = False Then
                                System.BitConverter.GetBytes(.oAssignments(1).oAgent.ObjectID).CopyTo(yResp, lPos) : lPos += 4
                            Else
                                System.BitConverter.GetBytes(-1I).CopyTo(yResp, lPos) : lPos += 4
                            End If
                            'assignment 2 skillid
                            If .oAssignments(1).oSkill Is Nothing = False Then
                                System.BitConverter.GetBytes(.oAssignments(1).oSkill.ObjectID).CopyTo(yResp, lPos) : lPos += 4
                            Else
                                System.BitConverter.GetBytes(-1I).CopyTo(yResp, lPos) : lPos += 4
                            End If
                            System.BitConverter.GetBytes(.oAssignments(1).PointsAccumulated).CopyTo(yResp, lPos) : lPos += 2
                            yResp(lPos) = .oAssignments(1).yStatus : lPos += 1
                        Else : lPos += 11
                        End If
                    Else : lPos += 22
                    End If

                End With
            Else : lPos += 26
            End If
        End If

		'now for our goals
		yResp(lPos) = yGoalCnt : lPos += 1
		If oMission Is Nothing = False Then
			For X As Int32 = 0 To oMission.GoalUB
				If oMission.MethodIDs(X) = lMethodID AndAlso oMissionGoals(X) Is Nothing = False Then
					If oMissionGoals(X).oGoal Is Nothing = False Then
						System.BitConverter.GetBytes(oMissionGoals(X).oGoal.ObjectID).CopyTo(yResp, lPos) : lPos += 4
					Else : System.BitConverter.GetBytes(-1I).CopyTo(yResp, lPos) : lPos += 4
					End If

					yResp(lPos) = CByte(lAssignCnt(X)) : lPos += 1
					For Y As Int32 = 0 To oMissionGoals(X).lAssignmentUB
						If oMissionGoals(X).oAssignments(Y) Is Nothing = False Then
							If oMissionGoals(X).oAssignments(Y).oAgent Is Nothing = False Then
								System.BitConverter.GetBytes(oMissionGoals(X).oAssignments(Y).oAgent.ObjectID).CopyTo(yResp, lPos) : lPos += 4
							Else : System.BitConverter.GetBytes(-1I).CopyTo(yResp, lPos) : lPos += 4
							End If
							If oMissionGoals(X).oAssignments(Y).oSkill Is Nothing = False Then
								System.BitConverter.GetBytes(oMissionGoals(X).oAssignments(Y).oSkill.ObjectID).CopyTo(yResp, lPos) : lPos += 4
							Else : System.BitConverter.GetBytes(-1I).CopyTo(yResp, lPos) : lPos += 4
							End If
							System.BitConverter.GetBytes(oMissionGoals(X).oAssignments(Y).PointsAccumulated).CopyTo(yResp, lPos) : lPos += 2
							yResp(lPos) = oMissionGoals(X).oAssignments(Y).yStatus : lPos += 1
						End If
					Next Y
				End If
			Next X
		End If
		Return yResp
    End Function

    Private Sub ReinfiltrationCaptureTest(ByRef oAgent As Agent)
        If oTarget Is Nothing = False AndAlso oTarget.ObjectID = oAgent.oOwner.ObjectID Then Return

        Dim lSafeHouseMod As Int32 = 0
        If ySafeHouseSetting > 1 Then
            lSafeHouseMod = CInt(ySafeHouseSetting)
        End If

        Dim lAgentBaseVal As Int32 = oAgent.GetSkillValue(lSkillHardcodes.eDisguises, False, 0)
        Dim bDisguises As Boolean = lAgentBaseVal > 0
        lAgentBaseVal += lSafeHouseMod
        lAgentBaseVal += oAgent.GetSkillValue(lSkillHardcodes.eSpyGames, True, 0) + oAgent.GetSkillValue(lSkillHardcodes.eEscapeArtist, True, 0)

        If oTarget Is Nothing = False AndAlso oTarget.oSecurity Is Nothing = False Then
            For X As Int32 = 0 To oTarget.oSecurity.lCounterAgentUB
                If oTarget.oSecurity.lCounterAgentIdx(X) > 0 Then
                    Dim oCounter As Agent = oTarget.oSecurity.oCounterAgents(X)
                    If oCounter Is Nothing = False Then
                        If oCounter.InfiltrationType = oAgent.InfiltrationType AndAlso (oCounter.lAgentStatus And (AgentStatus.IsDead Or AgentStatus.Dismissed Or AgentStatus.HasBeenCaptured Or AgentStatus.Infiltrating Or AgentStatus.ReturningHome)) = 0 Then
                            Dim lCounterVal As Int32 = CInt(Rnd() * 100) + oCounter.GetSkillValue(lSkillHardcodes.eSpyGames, True, 0)
                            lCounterVal += oCounter.GetSkillValue(lSkillHardcodes.eTracking, False, 0)
                            lCounterVal += oCounter.GetSkillValue(lSkillHardcodes.ePerceptive, True, 0)
                            If bDisguises = True Then lCounterVal += oCounter.GetSkillValue(lSkillHardcodes.eDisguises, False, 0)
                            lCounterVal -= Me.Modifier

                            'Now, do our test...
                            Dim lFinalVal As Int32 = lAgentBaseVal + CInt(Rnd() * 100)
                            If lFinalVal < lCounterVal Then
                                'we are in a reinfiltration scenario, do we use cover agents here? dont think so
                                oCounter.oOwner.oSecurity.CaptureAgent(oAgent)
                                Return
                            End If

                        End If
                    End If
                End If
            Next X
        End If

    End Sub

    Public Sub AddAccomplicesToIntel(ByRef oToPlayer As Player, ByVal lSqueeledAgentID As Int32)
        If Me.oMission Is Nothing = False Then
            For X As Int32 = 0 To oMission.GoalUB
                If oMission.MethodIDs(X) = lMethodID Then
                    If oMissionGoals(X) Is Nothing = False Then
                        With oMissionGoals(X)
                            For Y As Int32 = 0 To .lAssignmentUB
                                If .oAssignments(Y) Is Nothing = False Then
                                    If .oAssignments(Y).oAgent Is Nothing = False Then
                                        If .oAssignments(Y).oAgent.ObjectID <> lSqueeledAgentID Then
                                            Dim oAgent As Agent = .oAssignments(Y).oAgent
                                            CheckAndAddAccompliceIntel(oToPlayer, oAgent)
                                        End If
                                    End If
                                    If .oAssignments(Y).oCoveringAgent Is Nothing = False Then
                                        If oMissionGoals(X).oAssignments(Y).oCoveringAgent.ObjectID <> lSqueeledAgentID Then
                                            Dim oAgent As Agent = .oAssignments(Y).oCoveringAgent
                                            CheckAndAddAccompliceIntel(oToPlayer, oAgent)
                                        End If
                                    End If
                                End If
                            Next Y
                        End With
                    End If
                End If
            Next X
        End If
    End Sub
    Private Sub CheckAndAddAccompliceIntel(ByRef oToPlayer As Player, ByRef oAgent As Agent)
        If oToPlayer.HasItemIntel(oAgent.ObjectID, oAgent.ObjTypeID, oAgent.oOwner.ObjectID) = False Then
            Dim oPII As PlayerItemIntel = oToPlayer.AddPlayerItemIntel(oAgent.ObjectID, oAgent.ObjTypeID, PlayerItemIntel.PlayerItemIntelType.eLocation, oAgent.oOwner.ObjectID)
            If oPII Is Nothing = False Then
                oToPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
            End If
        End If
    End Sub

    Public Sub DeleteMe()
        If Me.PM_ID > -1 Then
            Dim sSQL As String = "DELETE FROM tblPMPhaseCoverAgent WHERE PM_ID = " & Me.PM_ID
            Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand(sSQL, goCN)
            Try
                oComm.ExecuteNonQuery()
                oComm = Nothing
            Catch
            End Try

            Try
                sSQL = "DELETE FROM tblPlayerMissionGoal WHERE PM_ID = " & Me.PM_ID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oComm.ExecuteNonQuery()
                oComm = Nothing
            Catch
            End Try

            Try
                sSQL = "DELETE FROM tblPlayerMission WHERE PM_ID = " & Me.PM_ID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oComm.ExecuteNonQuery()
                oComm = Nothing
            Catch
            End Try
        End If
    End Sub
End Class

Public Class PlayerMissionPhase
    Public oParent As PlayerMission
    Public lPhase As eMissionPhase

    Public oCoverAgents() As Agent
    Public lCoverAgentIdx() As Int32
    Public yUsedAsCoverAgent() As Byte
    Public lCoverAgentUB As Int32 = -1

    Public Sub AddCoverAgent(ByRef oAgent As Agent, ByVal yUsed As Byte)
        Dim lID As Int32 = oAgent.ObjectID
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To lCoverAgentUB
            If lCoverAgentIdx(X) = lID Then
                yUsedAsCoverAgent(X) = yUsed
                Return
            End If
            If lIdx = -1 AndAlso lCoverAgentIdx(X) = -1 Then lIdx = X
        Next X
        If lIdx = -1 Then
            lCoverAgentUB += 1
            ReDim Preserve oCoverAgents(lCoverAgentUB)
            ReDim Preserve yUsedAsCoverAgent(lCoverAgentUB)
            ReDim Preserve lCoverAgentIdx(lCoverAgentUB)
            lIdx = lCoverAgentUB
        End If
        oCoverAgents(lIdx) = oAgent
        lCoverAgentIdx(lIdx) = oAgent.ObjectID
        yUsedAsCoverAgent(lIdx) = yUsed
    End Sub

    Public Sub RemoveCoverAgent(ByVal lAgentID As Int32)
        For X As Int32 = 0 To lCoverAgentUB
            If lCoverAgentIdx(X) = lAgentID Then
                lCoverAgentIdx(X) = -1
                oCoverAgents(X) = Nothing
            End If
        Next X
    End Sub

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try

            sSQL = "DELETE FROM tblPMPhaseCoverAgent WHERE PM_ID = " & oParent.PM_ID & " AND PhaseID = " & CInt(lPhase)
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()
            oComm.Dispose()
            oComm = Nothing

            For X As Int32 = 0 To lCoverAgentUB
                If lCoverAgentIdx(X) <> -1 Then
                    sSQL = "INSERT INTO tblPMPhaseCoverAgent (PM_ID, PhaseID, CoverAgentID, UsedAsCoverAgent) VALUES (" & _
                      oParent.PM_ID & ", " & CInt(lPhase) & ", " & oCoverAgents(X).ObjectID & ", " & yUsedAsCoverAgent(X) & ")"
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    If oComm.ExecuteNonQuery() = 0 Then
                        Err.Raise(-1, "PlayerMissionPhase.SaveObject", "PlayerMissionPhase.SaveObject: No Records saved!")
                    End If
                    oComm.Dispose()
                    oComm = Nothing
                End If
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object PlayerMissionPhase. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function
End Class