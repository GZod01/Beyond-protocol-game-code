Public Enum AgentAssignmentStatus As Byte
	NoStatus = 0
	Skipped = 1			'agent assignment was skipped
	Finished = 2		'same as the Completed boolean server-side (indicates completion only)
	Success = 4			'indicates a success. If this flag is not set and Finished is set, then the AA was a failure
End Enum

Public Class AgentAssignment
	Public oParent As PlayerMissionGoal
	Public oAgent As Agent
	Public oSkill As Skill
	Public PointsAccumulated As Int16
	Public PointsRequired As Int16
	'Public Completed As Boolean = False
	Public yStatus As Byte = 0		'uses AgentAssignmentStatus

	'Private mlPreviousTestCycle As Int32 = 0

	'This is not configured by the player but instead occurs during the process of the assignment and is used if oAgent is Captured
	Public oCoveringAgent As Agent = Nothing	'used if oAgent is captured and a Cover Agent assumes the role of this assignment

	Public Property Completed() As Boolean
		Get
			Return (yStatus And AgentAssignmentStatus.Finished) <> 0
		End Get
		Set(ByVal value As Boolean)
			If value = True Then
				yStatus = yStatus Or AgentAssignmentStatus.Finished
			Else
				If (yStatus And AgentAssignmentStatus.Finished) <> 0 Then
					yStatus = yStatus Xor AgentAssignmentStatus.Finished
				End If
			End If
		End Set
	End Property

	Public Property IsSuccess() As Boolean
		Get
			Return (yStatus And AgentAssignmentStatus.Success) <> 0
		End Get
		Set(ByVal value As Boolean)
			If value = True Then
				yStatus = yStatus Or AgentAssignmentStatus.Success
			Else
				If (yStatus And AgentAssignmentStatus.Success) <> 0 Then yStatus = yStatus Xor AgentAssignmentStatus.Success
			End If
		End Set
	End Property

	Private Sub Preparation(ByRef oActingAgent As Agent)
		Dim lValue As Int32 = oActingAgent.GetSkillValue(oSkill.ObjectID, False, oActingAgent.Resourcefulness)
		lValue += CInt(Rnd() * 100)

		Dim oSkillSetSkill As SkillSet_Skill = Nothing
		For X As Int32 = 0 To oParent.oSkillSet.SkillUB
			If oSkill.ObjectID = oParent.oSkillSet.Skills(X).oSkill.ObjectID Then
				oSkillSetSkill = oParent.oSkillSet.Skills(X)
				Exit For
			End If
		Next X
		If oSkillSetSkill Is Nothing = False Then lValue += oSkillSetSkill.ToHitModifier
        lValue += Me.oParent.oMission.Modifier

        If Me.oParent.oMission.ySafeHouseSetting > 1 Then
            lValue += CInt(Me.oParent.oMission.ySafeHouseSetting)
        End If

		'if the agent is working double for the current phase then check if the agent is able to process this assignment right now
		For X As Int32 = 0 To oParent.lAssignmentUB
			If oParent.oAssignments(X) Is Nothing = False Then
				If oParent.oAssignments(X).oAgent Is Nothing = False Then
					If oParent.oAssignments(X).oAgent.ObjectID = oAgent.ObjectID Then
						If oParent.oAssignments(X).oSkill Is Nothing = False AndAlso oParent.oAssignments(X).oSkill.ObjectID = oSkill.ObjectID Then
							Exit For
						ElseIf (oParent.oAssignments(X).yStatus And (AgentAssignmentStatus.Skipped Or AgentAssignmentStatus.Finished)) = 0 Then
							'ok, this agent is assigned elsewhere, we're going to break out
							Return
						End If
					End If
				Else : Return
				End If
			End If
		Next X

        'LogEvent(LogEventType.Informational, "AA.Preparation: " & lValue & " for " & BytesToString(oAgent.AgentName))

		If oSkill.MaxVal = 20 Then
			'1 to 120 base
			If lValue < 30 Then
				PointsAccumulated -= 3S
			ElseIf lValue < 40 Then
				PointsAccumulated -= 1S
			ElseIf lValue < 100 Then
				PointsAccumulated += 1S
			Else : PointsAccumulated += 3S
			End If
		Else
			'These use to be 10, 50, 180, else
			If lValue < 50 Then
				PointsAccumulated -= 3S
			ElseIf lValue < 60 Then
				PointsAccumulated -= 1S
			ElseIf lValue < 180 Then
				PointsAccumulated += 1S
			Else : PointsAccumulated += 3S
			End If
        End If
        If PointsAccumulated < -100 Then
            PointsAccumulated = -100
            Me.oParent.oMission.HandleMissionFailure()
            Return
        End If
        If gb_IS_TEST_SERVER = True Then PointsAccumulated = PointsRequired
		Completed = PointsAccumulated >= PointsRequired

		If Completed = True Then
			oParent.CheckAssignmentsCompleted()
		End If
	End Sub

	Private Sub Execution(ByRef oActingAgent As Agent)
        Dim lToHit As Int32 = oActingAgent.GetSkillValue(oSkill.ObjectID, False, 5)

		Dim oSkillSetSkill As SkillSet_Skill = Nothing
		For X As Int32 = 0 To oParent.oSkillSet.SkillUB
			If oSkill.ObjectID = oParent.oSkillSet.Skills(X).oSkill.ObjectID Then
				oSkillSetSkill = oParent.oSkillSet.Skills(X)
				Exit For
			End If
		Next X
		If oSkillSetSkill Is Nothing = False Then lToHit += oSkillSetSkill.ToHitModifier
		lToHit += Me.oParent.oMission.Modifier

		Dim lMaxVal As Int32 = oSkill.MaxVal
		Dim lMinVal As Int32 = oSkill.MinVal

		If lToHit < lMinVal Then lToHit = lMinVal
		If lToHit > lMaxVal Then lToHit = lMaxVal

		Dim lRoll As Int32 = CInt((Rnd() * (lMaxVal - lMinVal)) + lMinVal)
		If lRoll < lToHit Then
			'Ok, goal succeeded
            LogEvent(LogEventType.Informational, "AA.Execution Success for " & BytesToString(oActingAgent.AgentName))
			IsSuccess = True
		Else
			'ok, goal failed
            LogEvent(LogEventType.Informational, "AA.Execution Failed for " & BytesToString(oActingAgent.AgentName))
			IsSuccess = False
		End If
		Completed = True

	End Sub

	Public Sub ProcessAgentAssignment()

		'If glCurrentCycle - mlPreviousTestCycle > Me.oParent.oGoal.BaseTime Then
		'	mlPreviousTestCycle = glCurrentCycle
		'Else : Return
		'End If

		Dim oActingAgent As Agent = oAgent
		If oAgent Is Nothing OrElse (oAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 OrElse (oAgent.lAgentStatus And AgentStatus.IsDead) <> 0 Then
			oActingAgent = oCoveringAgent
		End If
		If oActingAgent Is Nothing Then 
            Me.oParent.oMission.HandleMissionFailure()
            Return
        End If

		If Me.oParent.oGoal.MissionPhase = eMissionPhase.ePreparationTime Then
			Preparation(oActingAgent)
		ElseIf Me.oParent.oGoal.MissionPhase > eMissionPhase.ePreparationTime Then
			Execution(oActingAgent)
		Else : Return
		End If

		If oParent.oMission.oTarget Is Nothing Then Return
        If oParent.oMission.oTarget.ObjectID = oParent.oMission.oPlayer.ObjectID Then Return

		'Check for risk of detection
		'1D100 < (Goal_Risk_of_Detection + (Agent_Infiltration_Level / 2)) - (Counter_Agent_Luck / 2) + Agent_Suspicion
        Dim lToHitNum As Int32 = CInt(oParent.oGoal.RiskOfDetection) + (CInt(oActingAgent.InfiltrationLevel) \ 2) + CInt(oActingAgent.Suspicion)
		Dim lTempValue As Int32 = oActingAgent.GetSkillValue(lSkillHardcodes.eStealth, True, 0)
		If lTempValue <> 0 Then
			If CInt(Rnd() * 100) < lTempValue Then lToHitNum -= 10
		End If
		lToHitNum -= oParent.oMission.Modifier			'use a minus here because higher difficulty increases counter agent ability

		With oParent.oMission.oTarget.oSecurity
			For X As Int32 = 0 To .lCounterAgentUB
				If .lCounterAgentIdx(X) <> -1 Then
					Dim lModVal As Int32 = 0

					If .oCounterAgents(X).InfiltrationType = eInfiltrationType.eGeneralInfiltration Then
						lModVal = .oCounterAgents(X).Luck \ 4
					ElseIf .oCounterAgents(X).InfiltrationType = oParent.oMission.oMission.lInfiltrationType Then
						lModVal = .oCounterAgents(X).Luck \ 2
					Else : Continue For
					End If

					lTempValue = .oCounterAgents(X).GetSkillValue(lSkillHardcodes.eDeductiveReasoning, True, 0)
					If lTempValue <> 0 Then
						If CInt(Rnd() * 100) < lTempValue Then
							'ok, add +15
							lModVal += 15
						End If
					End If
					lModVal += .oCounterAgents(X).GetSkillValue(lSkillHardcodes.eParanoid, True, 0)
                    'lModVal += .oCounterAgents(X).GetSkillValue(lSkillHardcodes.eSecurity, True, 0)
					lModVal += .oCounterAgents(X).GetSkillValue(lSkillHardcodes.eSpyGames, True, 0)

					lTempValue = .oCounterAgents(X).GetSkillValue(lSkillHardcodes.eTreachery, True, 0)
					If lTempValue <> 0 Then If CInt(Rnd() * 100) < lTempValue Then lModVal += 10


					If CInt(Rnd() * 100) < lToHitNum + lModVal OrElse PointsAccumulated < (PointsRequired * -3) Then
						RiskOfDetectionTestFailed(oActingAgent, .oCounterAgents(X))
                    End If

                    'Ok, finally, check for the safehouse setting to be compromised
                    If .oCounterAgents(X).InfiltrationType = oParent.oMission.oMission.lInfiltrationType Then
                        If Me.oParent.oMission.ySafeHouseSetting > 1 AndAlso Me.oParent.oMission.oSafeHouseGoal Is Nothing = False Then
                            lTempValue = .oCounterAgents(X).Resourcefulness
                            lTempValue += .oCounterAgents(X).GetSkillValue(lSkillHardcodes.eTracking, False, 0)
                            lTempValue += .oCounterAgents(X).GetSkillValue(lSkillHardcodes.eParanoid, True, 0)
                            lTempValue += .oCounterAgents(X).GetSkillValue(lSkillHardcodes.ePerceptive, True, 0)
                            lTempValue += .oCounterAgents(X).GetSkillValue(lSkillHardcodes.eSpyGames, True, 0)
                            lTempValue += Me.oParent.oMission.Modifier

                            lTempValue -= oActingAgent.GetSkillValue(lSkillHardcodes.eStealth, False, 0)
                            lTempValue -= CInt(oActingAgent.Resourcefulness)
                            lTempValue -= oActingAgent.GetSkillValue(lSkillHardcodes.eSpyGames, True, 0)

                            If Me.oParent.oMission.oSafeHouseGoal.oSkillSet Is Nothing = False Then
                                Dim lVal As Int32 = Me.oParent.oMission.oSafeHouseGoal.oSkillSet.ProgramControlID
                                If lVal = 1 Then
                                    'slum lords - harder to find
                                    lTempValue -= 20

                                    'Else
                                    'living large - easier to find... so it receives no bonus
                                End If
                            End If

                            lTempValue += CInt(Rnd() * 100 - 50)
                            If lTempValue > 0 Then
                                Me.oParent.oMission.ySafeHouseSetting = 0
                                If (oActingAgent.oOwner.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                                    Dim oPC As PlayerComm = oActingAgent.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                                                                    "The safehouse in the mission '" & BytesToString(Me.oParent.oMission.oMission.MissionName) & _
                                                                    "' against " & .oCounterAgents(X).oOwner.sPlayerNameProper & " has been compromised. Your agent team will have to make do without it.", "Safehouse Compromised", oActingAgent.oOwner.ObjectID, GetDateAsNumber(Now), False, oActingAgent.oOwner.sPlayerNameProper, Nothing)
                                    If oPC Is Nothing = False Then
                                        oActingAgent.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                                    End If
                                End If
                            End If

                        End If
                    End If
                    
                End If
            Next X
		End With

	End Sub

	Private Sub RiskOfDetectionTestFailed(ByRef oActingAgent As Agent, ByRef oDetector As Agent)
        'Ok, determine what phase I am in
        If oActingAgent.oOwner.ObjectID = oDetector.oOwner.ObjectID Then Return

		Select Case oParent.oGoal.MissionPhase
			Case eMissionPhase.ePreparationTime
				'during the preparation phase, nothing too severe, the agent receives suspicion (+3)
                Dim lVal As Int32 = CInt(oActingAgent.Suspicion) + 3
				If lVal > 255 Then lVal = 255
				oActingAgent.Suspicion = CByte(lVal)
				goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.ReduceSuspicion, oActingAgent, Nothing, glCurrentCycle + Agent.ml_DEINFILTRATE_TIME)
			Case eMissionPhase.eSettingTheStage, eMissionPhase.eFlippingTheSwitch, eMissionPhase.eReinfiltrationPhase
				'Ok, a cover agent/enforcer has discovered the agent doing something fishy
				' At this point, the agent can subdue the counter-agent by doing a "dagger" test
				Dim lResult As Int32 = oActingAgent.HandleDaggerTest(oDetector, eDaggerTestType.AvoidCapture)	'never returns 0!!!
				'Ok, if the result is positive, the agent wins
				If lResult > 0 Then
					'detection avoided... 
					If lResult > 20 Then
						oDetector.KillMe() 'counter-agent killed
						If (oDetector.oOwner.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                            Dim oPC As PlayerComm = oDetector.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                                "Your agent, " & BytesToString(oDetector.AgentName) & ", was mysteriously killed. An investigation is underway to determine cause and suspects. At this time, it appears to be a possible agent mission from a rival empire.", _
                                "Counter Agent Killed", oDetector.oOwner.ObjectID, GetDateAsNumber(Now), False, oDetector.oOwner.sPlayerNameProper, Nothing)
							If oPC Is Nothing = False Then
								oDetector.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
							End If
						End If
					End If
				Else
					'test for a cover agent that is available
					Dim oCover As Agent = oParent.oMission.GetAvailableCoverAgent(oParent.oGoal.MissionPhase)
					'is there a cover agent available?
					If oCover Is Nothing = False Then
                        oCover.lAgentStatus = oCover.lAgentStatus Or AgentStatus.UsedAsCoverAgent

                        Dim yResult As Agent.eyCoverAgentVsCounterResult = oCover.HandleCoverAgentVsCounter(oDetector, Me.oParent.oMission.Modifier)
                        If (yResult And Agent.eyCoverAgentVsCounterResult.AlarmThrown) <> 0 Then
                            lResult -= 21
                        End If
                        If (yResult And Agent.eyCoverAgentVsCounterResult.CounterAgentDies) <> 0 Then
                            oDetector.KillMe()
                            If (oDetector.oOwner.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                                Dim oPC As PlayerComm = oDetector.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                                  "Your agent, " & BytesToString(oDetector.AgentName) & ", was mysteriously killed. An investigation is underway to determine cause and suspects. At this time, it appears to be a possible agent mission from a rival empire.", _
                                  "Counter Agent Killed", oDetector.oOwner.ObjectID, GetDateAsNumber(Now), False, oDetector.oOwner.sPlayerNameProper, Nothing)
                                If oPC Is Nothing = False Then
                                    oDetector.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                                End If
                            End If
                        End If
                        If (yResult And Agent.eyCoverAgentVsCounterResult.CoverAgentDies) <> 0 Then
                            oCover.KillMe()
                            If (oCover.oOwner.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                                Dim sHowever As String = ""
                                If (yResult And Agent.eyCoverAgentVsCounterResult.CoverAgentWins) <> 0 Then
                                    sHowever = " The team reports that the intercept worked and that the original agent is still in place."
                                End If
                                Dim oPC As PlayerComm = oCover.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                                  "Your agent, " & BytesToString(oCover.AgentName) & ", was killed trying to intercept counter agents during their mission." & sHowever, "Cover Agent Killed", oDetector.oOwner.ObjectID, GetDateAsNumber(Now), False, oDetector.oOwner.sPlayerNameProper, Nothing)
                                If oPC Is Nothing = False Then
                                    oCover.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                                End If
                            End If
                        ElseIf (yResult And Agent.eyCoverAgentVsCounterResult.CoverAgentCaptured) <> 0 Then
                            oDetector.oOwner.oSecurity.CaptureAgent(oCover)
                        End If

                        'Now, determine the winner
                        If (yResult And Agent.eyCoverAgentVsCounterResult.CounterAgentWins) <> 0 Then
                            'ok, counter agent won, so the target agent still gets captured
                            oDetector.oOwner.oSecurity.CaptureAgent(oActingAgent)
                            If lResult < -20 Then
                                If oParent.oMission.bAlarmThrown = False Then
                                    oParent.oMission.bAlarmThrown = True
                                    oParent.oMission.Modifier += -20
                                End If
                            End If
                            If (yResult And Agent.eyCoverAgentVsCounterResult.CoverAgentDies) = 0 Then
                                If oCoveringAgent Is Nothing Then oCoveringAgent = oCover
                            End If
                        ElseIf (yResult And Agent.eyCoverAgentVsCounterResult.CoverAgentWins) <> 0 AndAlso (yResult And Agent.eyCoverAgentVsCounterResult.CoverAgentDies) = 0 Then
                            Dim oPC As PlayerComm = oCover.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                              "Your agent, " & BytesToString(oCover.AgentName) & " successfully intercepted a counter agent in their current mission saving " & BytesToString(oAgent.AgentName) & " from capture.", "Cover Agent Saves the Day", oCover.oOwner.ObjectID, GetDateAsNumber(Now), False, oCover.oOwner.sPlayerNameProper, Nothing)
                            If oPC Is Nothing = False Then
                                oCover.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                            End If
                        End If


                        'Dim lTmpResult As Int32 = oCover.HandleDaggerTest(oDetector, eDaggerTestType.AvoidCapture)
                        'If lTmpResult < 0 Then
                        '	'Counter agent won... modify lresult to ensure alarm is thrown
                        '	lResult = -21

                        '	'Ok, now, is the counter agent uber?
                        '	If Math.Abs(lTmpResult) > oCover.Dagger Then
                        '		oDetector.oOwner.oSecurity.CaptureAgent(oCover)
                        '	End If
                        'Else
                        '	'Cover agent won... kill the counter-agent
                        '	oDetector.KillMe()
                        '	If (oDetector.oOwner.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                        '		Dim oPC As PlayerComm = oDetector.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                        '				"Your agent, " & BytesToString(oDetector.AgentName) & ", was mysteriously killed. An investigation is underway to determine cause and suspects. At this time, it appears to be a possible agent mission from a rival empire.", _
                        '				"Counter Agent Killed", oDetector.oOwner.ObjectID, GetDateAsNumber(Now), False, oDetector.oOwner.sPlayerNameProper)
                        '		If oPC Is Nothing = False Then
                        '			oDetector.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                        '		End If
                        '	End If
                        '	Return
                        'End If
                    Else
                        'the agent is captured
                        oDetector.oOwner.oSecurity.CaptureAgent(oActingAgent)
                        If lResult < -20 Then
                            If oParent.oMission.bAlarmThrown = False Then
                                oParent.oMission.bAlarmThrown = True
                                oParent.oMission.Modifier += -20
                            End If
                        End If

                        If oCoveringAgent Is Nothing Then oCoveringAgent = oCover
                    End If

				End If
		End Select
	End Sub

End Class