Option Strict On

'Represents the security level of the player
Public Class PlayerSecurity
    Public lInfiltrationResistance() As Int32           'based on technology, government setup, military standards, etc...

	Public lAllMissionModifier As Int32 = 0				'modifier applied to all missions of this player

    Public oCounterAgents() As Agent
    Public lCounterAgentIdx() As Int32
	Public lCounterAgentUB As Int32 = -1

	Public oCapturedAgents() As Agent
	Public lCapturedAgentIdx() As Int32
	Public lCapturedAgentUB As Int32 = -1

    Public oParentPlayer As Player

	Public Sub AddCapturedAgent(ByRef oAgent As Agent)
		Dim lIdx As Int32 = -1
		For X As Int32 = 0 To lCapturedAgentUB
			If lCapturedAgentIdx(X) = -1 Then
				If lIdx = -1 Then lIdx = X
			ElseIf lCapturedAgentIdx(X) = oAgent.ObjectID Then
				Return
			End If
		Next X

		If lIdx = -1 Then
			lCapturedAgentUB += 1
			ReDim Preserve lCapturedAgentIdx(lCapturedAgentUB)
			ReDim Preserve oCapturedAgents(lCapturedAgentUB)
			lIdx = lCapturedAgentUB
		End If
		oCapturedAgents(lIdx) = oAgent
		lCapturedAgentIdx(lIdx) = oAgent.ObjectID
	End Sub

    Public Sub CaptureAgent(ByRef oAgent As Agent)
        'do the capture and notify both parties

        If oAgent.lCapturedBy > 0 AndAlso (oAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then
            Dim oPrevCaptor As Player = GetEpicaPlayer(oAgent.lCapturedBy)
            If oPrevCaptor Is Nothing = False Then
                Dim sBody As String = "Your agent network reports that the target agent to be captured could not be located.  No intel is available on the agent's current whereabouts."
                If (oAgent.ySpilledData And eSpilledData.AgentName) <> 0 Then
                    sBody &= vbCrLf & vbCrLf & "The agent's name was " & BytesToString(oAgent.AgentName) & "."
                Else
                    sBody &= vbCrLf & vbCrLf & "We never got the agent's name."
                End If
                Dim oCaptPC As PlayerComm = oPrevCaptor.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, sBody, "Missing Captured Agent", oPrevCaptor.ObjectID, GetDateAsNumber(Now), False, oPrevCaptor.sPlayerNameProper, Nothing)
                If oCaptPC Is Nothing = False Then
                    oCaptPC.SaveObject(oPrevCaptor.ObjectID)
                    oPrevCaptor.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oCaptPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                End If
            End If
            Exit Sub 'If the agent is captured, do not alert the target that we ReCaptured, nor roll an honor kill as they are already locked up in chains somewhere.
        End If

        AddCapturedAgent(oAgent)
        Dim bHonorKill As Boolean = False
        With oAgent

            Dim lHonor As Int32 = .GetSkillValue(lSkillHardcodes.eHonor, True, 0)
            If lHonor > 0 Then
                If Rnd() * 20 < lHonor AndAlso Rnd() * 100 < .Dagger Then
                    bHonorKill = True
                    Dim oKillPC As PlayerComm = oAgent.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                     "Intel reports that our agent " & BytesToString(oAgent.AgentName) & " is dead by their own hands (Honor Killing) while being captured by " & Me.oParentPlayer.sPlayerNameProper & ".", _
                     "Agent Killed", oAgent.oOwner.ObjectID, GetDateAsNumber(Now), False, oAgent.oOwner.sPlayerNameProper, Nothing)
                    If oKillPC Is Nothing = False Then oAgent.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oKillPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    oKillPC = Nothing
                    oKillPC = Me.oParentPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, "During the process of being captured, an enemy agent committed suicide instead of allowing themselves to be captured.", _
                        "Prisoner Killed", Me.oParentPlayer.ObjectID, GetDateAsNumber(Now), False, Me.oParentPlayer.sPlayerNameProper, Nothing)
                    If oKillPC Is Nothing = False Then oParentPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oKillPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                End If
            End If

            .lAgentStatus = .lAgentStatus Or AgentStatus.HasBeenCaptured
            .lCapturedBy = Me.oParentPlayer.ObjectID
            .lCapturedOn = glCurrentCycle
            .lPrisonTestCycles = CInt((2052000 * Rnd()) - (540000)) '5-24hr (19*rnd+5)
            .yHealth = 100
            .lInterrogatorID = -1
            .ySpilledData = eSpilledData.NoDataSpilled
            .RemoveMeFromMissions()
            If (.lAgentStatus And AgentStatus.OnAMission) <> 0 Then
                .lAgentStatus = .lAgentStatus Xor AgentStatus.OnAMission
            End If
            goAgentEngine.CancelAllAgentEvents(.ObjectID)

            .lTargetID = Me.oParentPlayer.ObjectID
            .InfiltrationLevel = 0
            .iTargetTypeID = ObjectType.ePlayer

            If bHonorKill = True Then .KillMe()
            If .lAgentStatus <> AgentStatus.IsDead Then goAgentEngine.AddAgentEvent(AgentEngine.EventTypeID.PrisonEscapeTest, oAgent, Nothing, glCurrentCycle + .lPrisonTestCycles)

            .oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oAgent, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
        End With

        If bHonorKill = False Then
            Dim oPC As PlayerComm = Nothing
            If (Me.oParentPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then oPC = Me.oParentPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, "We have captured an enemy agent doing some illegal activities. We have brought the agent to a secure location.", "Enemy Agent Captured", Me.oParentPlayer.ObjectID, GetDateAsNumber(Now), False, Me.oParentPlayer.sPlayerNameProper, Nothing)
            If oPC Is Nothing = False Then
                Me.oParentPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            oPC = Nothing

            If oAgent Is Nothing = False AndAlso (oAgent.oOwner.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then oPC = oAgent.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, "One of our agents (" & BytesToString(oAgent.AgentName) & ") is reported missing and is believed to have been captured by " & Me.oParentPlayer.sPlayerNameProper & ".", "Missing Agent", oAgent.oOwner.ObjectID, GetDateAsNumber(Now), False, oAgent.oOwner.sPlayerNameProper, Nothing)
            If oPC Is Nothing = False Then
                oAgent.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
            oPC = Nothing
        End If
		

        If Me.oParentPlayer.lConnectedPrimaryID > -1 OrElse Me.oParentPlayer.HasOnlineAliases(AliasingRights.eViewAgents) = True Then
            If oAgent Is Nothing = False Then Me.oParentPlayer.SendPlayerMessage(oAgent.GetCapturedAgentMsg, False, AliasingRights.eViewAgents)
        End If

	End Sub

    Public Sub AddCounterAgent(ByRef oAgent As Agent)
        Dim lIdx As Int32 = -1
        For X As Int32 = 0 To lCounterAgentUB
            If lCounterAgentIdx(X) = -1 Then
                If lIdx = -1 Then lIdx = X
            ElseIf lCounterAgentIdx(X) = oAgent.ObjectID Then
                Return
            End If
        Next X

        If lIdx = -1 Then
            lCounterAgentUB += 1
            ReDim Preserve lCounterAgentIdx(lCounterAgentUB)
            ReDim Preserve oCounterAgents(lCounterAgentUB)
            lIdx = lCounterAgentUB
        End If
        oCounterAgents(lIdx) = oAgent
        lCounterAgentIdx(lIdx) = oAgent.ObjectID
    End Sub

    Public Function GetCounterAgentCount(ByVal yInfType As eInfiltrationType) As Int32
        Dim lResult As Int32 = 0
        For X As Int32 = 0 To lCounterAgentUB
            If lCounterAgentIdx(X) <> -1 Then
                Dim oAgent As Agent = oCounterAgents(X)
                If oAgent Is Nothing = False AndAlso oAgent.InfiltrationType = yInfType Then
                    If (oAgent.lAgentStatus And AgentStatus.CounterAgent) <> 0 AndAlso (oAgent.lAgentStatus And (AgentStatus.HasBeenCaptured Or AgentStatus.IsDead Or AgentStatus.ReturningHome)) = 0 Then
                        lResult += 1
                        'for more balance, try this on
                        'If oAgent.InfiltrationLevel > 150 Then lResult += 1
                    End If
                End If
            End If
        Next X
        Return lResult
    End Function

    Public Function GetInfiltrationCounterAgentModifier(ByVal yInfType As eInfiltrationType, ByVal lInfDisguises As Int32, ByVal lInfPersuasive As Int32, ByVal lInfForger As Int32) As Int32
        Dim lCurUB As Int32 = -1
        If lCounterAgentIdx Is Nothing = False Then lCurUB = Math.Min(lCounterAgentUB, lCounterAgentIdx.GetUpperBound(0))
        Dim lCounterAgents As Int32 = 0
        Dim lResult As Int32 = 0
        For X As Int32 = 0 To lCurUB
            If lCounterAgentIdx(X) <> -1 Then
                Dim oAgent As Agent = oCounterAgents(X)
                If oAgent Is Nothing = False AndAlso oAgent.InfiltrationType = yInfType AndAlso (oAgent.lAgentStatus And (AgentStatus.Dismissed Or AgentStatus.HasBeenCaptured Or AgentStatus.Infiltrating Or AgentStatus.IsDead Or AgentStatus.ReturningHome)) = 0 Then
                    lCounterAgents += 1
                    With oAgent
                        'ok, get this agent's to hit number
                        'To hit = ((Luck + Dagger) / 2)
                        Dim fToHit As Single = (CSng(.Luck) + CSng(.Dagger)) / 2.0F
                        fToHit += .GetSkillValue(lSkillHardcodes.eParanoid, True, 0)

                        Dim lTemp As Int32 = .GetSkillValue(lSkillHardcodes.ePerceptive, True, 0)
                        If lInfDisguises > 0 OrElse lTemp > 0 Then
                            If Rnd() * 100 + (lTemp * 5) > Rnd() * 100 + lInfDisguises Then
                                fToHit += 20
                            End If
                        End If

                        If Rnd() * 100 < lInfPersuasive Then
                            fToHit -= 20
                        End If

                        If lInfForger > 0 Then
                            lTemp = .GetSkillValue(lSkillHardcodes.eForger, True, 0)
                            If Rnd() * 100 < lInfForger - lTemp Then fToHit -= 20
                        End If
                        fToHit += (.InfiltrationLevel / 2.5F)
                        fToHit /= 2.0F

                        If Rnd() * 100 < fToHit Then lResult -= 5
                    End With
                End If

            End If
        Next X

        Return lResult + lCounterAgents
    End Function

	Private Const ml_DAGGER_TEST_MAXDMG As Int32 = 40
	Private Const ml_RESOURCE_TEST_MAXDMG As Int32 = 5
	Private Const ml_PERSUASION_TEST_MAXDMG As Int32 = 5
	Private Const ml_POISONS_TEST_MAXDMG As Int32 = 70
	Private Const ml_TORTURE_TEST_MAXDMG As Int32 = 70
	Private Const ml_TREACHERY_TEST_MAXDMG As Int32 = 15

	Private Const ml_AGENT_NAME_CHANCE As Int32 = 75
	Private Const ml_INF_SPECS_CHANCE As Int32 = 60
	Private Const ml_OWNER_CHANCE As Int32 = 40
	Private Const ml_MISSION_CHANCE As Int32 = 35
	Private Const ml_TARGET_CHANCE As Int32 = 30
	Private Const ml_ACCOMPLICE_CHANCE As Int32 = 25

	Private Shared ml_MAX_DMG_VALUES() As Int32
	Private Shared Function GetAgentDmg(ByVal lTestIdx As Int32, ByVal lASB As Int32, ByVal lIMB As Int32) As Single
		If ml_MAX_DMG_VALUES Is Nothing Then
			ReDim ml_MAX_DMG_VALUES(5)
			ml_MAX_DMG_VALUES(0) = ml_DAGGER_TEST_MAXDMG
			ml_MAX_DMG_VALUES(1) = ml_RESOURCE_TEST_MAXDMG
			ml_MAX_DMG_VALUES(2) = ml_PERSUASION_TEST_MAXDMG
			ml_MAX_DMG_VALUES(3) = ml_POISONS_TEST_MAXDMG
			ml_MAX_DMG_VALUES(4) = ml_TORTURE_TEST_MAXDMG
			ml_MAX_DMG_VALUES(5) = ml_TREACHERY_TEST_MAXDMG
		End If

		Dim fValue As Single = (lASB + lIMB) / 100.0F
		Return fValue * ml_MAX_DMG_VALUES(lTestIdx)
	End Function

	Public Sub HandleInterrogationTest(ByRef oAgent As Agent, ByRef oInterrogator As Agent)
		'Ok, there are a series of tests that are run:
		'Interrogator vs. Agent
		'	Dagger vs. Dagger
		'	Resource vs. Loyalty
		'	Persuasion vs. (Loyalty + Honor)
		'	Poisons vs. (Loyalty + Endurance)
		'	Torture vs. (Loyalty + Endurance)
		'	Treachery vs. (Loyalty + Honor)
		'
		'Tests are only ran for skills that the Interrogator has
		'The Interrogator's Interrogation skill is added to his to-hit numbers for all tests
		'Each agent rolls against their modified to-hit number and attempts to get below the number
		'If the Interrogator misses, the value missed by is saved
		'If the Agent succeeds, the value succeeded by is saved
		'
		'Reason is that damage can be done to the agent
		'When the Interrogator fails any test, damage is done to the AGENT (not the Interrogator)
		'	Damage is based on amount missed listed above based on this formula:
		'		ASB = Agent Succeeds By
		'		IMB = Interrogator Missed By
		'		Damage Done = ((ASB + IMB) / 100) * <Damage for this test>
		'Values are stored in a single and added up at the end. The final sum is rounded up
		'
		'An Attempt Count is calculated based on: Interrogator's success - Agent's Success.
		'If the attempt count is positive, then a roll can be made against the information chart


		'Right now, there are only 6 possible tests
		Dim lUB As Int32 = 5
		Dim lIMB(lUB) As Int32
		Dim lASB(lUB) As Int32
		Dim lInterrogatorSuccess As Int32 = 0
		Dim lAgentSuccess As Int32 = 0
		Dim lRoll As Int32 = 0
		Dim lToHit As Int32 = 0
		Dim lIdx As Int32 = 0

		'Interrogator's Interrogation skill
		Dim lIntInt As Int32 = oInterrogator.GetSkillValue(lSkillHardcodes.eInterrogation, True, 0)

		For X As Int32 = 0 To lUB
			lIMB(X) = 0 : lASB(X) = 0
		Next X

		'ok, first, the dagger
		lRoll = CInt(Rnd() * 100)
		lToHit = CInt(oInterrogator.Dagger) + lIntInt
		lIdx = 0
		If lRoll < lToHit Then
			lInterrogatorSuccess += 1
		Else : lIMB(lIdx) = lRoll - lToHit
		End If
		lRoll = CInt(Rnd() * 100)
		lToHit = oAgent.Dagger
		If lRoll < lToHit Then
			lAgentSuccess += 1
			lASB(lIdx) = lToHit - lRoll
		End If

		'next is resource
		lRoll = CInt(Rnd() * 100)
		lIdx = 1
		lToHit = CInt(oInterrogator.Resourcefulness) + lIntInt
		If lRoll < lToHit Then
			lInterrogatorSuccess += 1
		Else : lIMB(1) = lRoll - lToHit
		End If
		lRoll = CInt(Rnd() * 100)
		lToHit = CInt(oAgent.Loyalty)
		If lRoll < lToHit Then
			lAgentSuccess += 1
			lASB(1) = lToHit - lRoll
		End If

		'Now, persuasion
		lIdx = 2
		lToHit = oInterrogator.GetSkillValue(lSkillHardcodes.ePersuasive, True, 0)
		If lToHit <> 0 Then
			lToHit += lIntInt
			lRoll = CInt(Rnd() * 100)
			If lRoll < lToHit Then
				lInterrogatorSuccess += 1
			Else : lIMB(lIdx) = lRoll - lToHit
			End If
			lRoll = CInt(Rnd() * 100)
			lToHit = CInt(oAgent.Loyalty) + oAgent.GetSkillValue(lSkillHardcodes.eHonor, True, 0)
			If lRoll < lToHit Then
				lAgentSuccess += 1
				lASB(lIdx) = lToHit - lRoll
			End If
		End If

		'Now, Poison
		lIdx += 1
		lToHit = oInterrogator.GetSkillValue(lSkillHardcodes.ePoisons, True, 0)
		If lToHit <> 0 Then
			lToHit += lIntInt
			lRoll = CInt(Rnd() * 100)
			If lRoll < lToHit Then
				lInterrogatorSuccess += 1
			Else : lIMB(lIdx) = lRoll - lToHit
			End If
			lRoll = CInt(Rnd() * 100)
			lToHit = CInt(oAgent.Loyalty) + oAgent.GetSkillValue(lSkillHardcodes.eEndurance, True, 0)
			If lRoll < lToHit Then
				lAgentSuccess += 1
				lASB(lIdx) = lToHit - lRoll
			End If
		End If

		'Now, torture
		lIdx += 1
		lToHit = oInterrogator.GetSkillValue(lSkillHardcodes.eTorture, True, 0)
		If lToHit <> 0 Then
			lToHit += lIntInt
			lRoll = CInt(Rnd() * 100)
			If lRoll < lToHit Then
				lInterrogatorSuccess += 1
			Else : lIMB(lIdx) = lRoll - lToHit
			End If
			lRoll = CInt(Rnd() * 100)
			lToHit = CInt(oAgent.Loyalty) + oAgent.GetSkillValue(lSkillHardcodes.eEndurance, True, 0)
			If lRoll < lToHit Then
				lAgentSuccess += 1
				lASB(lIdx) = lToHit - lRoll
			End If
		End If

		'Now, treachery
		lIdx += 1
		lToHit = oInterrogator.GetSkillValue(lSkillHardcodes.eTreachery, True, 0)
		If lToHit <> 0 Then
			lToHit += lIntInt
			lRoll = CInt(Rnd() * 100)
			If lRoll < lToHit Then
				lInterrogatorSuccess += 1
			Else : lIMB(lIdx) = lRoll - lToHit
			End If
			lRoll = CInt(Rnd() * 100)
			lToHit = CInt(oAgent.Loyalty) + oAgent.GetSkillValue(lSkillHardcodes.eHonor, True, 0)
			If lRoll < lToHit Then
				lAgentSuccess += 1
				lASB(lIdx) = lToHit - lRoll
			End If
		End If


		'================================================
		Dim bSendUpdate As Boolean = False

		'Ok, all tests are done... let's determine who won
		If lInterrogatorSuccess > lAgentSuccess Then
			'ok, interrogator won, determine our attempt count
			Dim lAttempts As Int32 = lInterrogatorSuccess - lAgentSuccess
			Dim yOriginalSpill As Byte = oAgent.ySpilledData
			For X As Int32 = 0 To lAttempts - 1
				'roll our 1d100
				lRoll = CInt(Rnd() * 100)
				'ok, determine what test we are running, it goes in this order:
				If (oAgent.ySpilledData And eSpilledData.AgentName) = 0 Then
					'agent name
					If lRoll < ml_AGENT_NAME_CHANCE Then oAgent.ySpilledData = oAgent.ySpilledData Or eSpilledData.AgentName
				ElseIf (oAgent.ySpilledData And eSpilledData.InfiltrationSpecifics) = 0 Then
					'infiltration specifics
					If lRoll < ml_INF_SPECS_CHANCE Then oAgent.ySpilledData = oAgent.ySpilledData Or eSpilledData.InfiltrationSpecifics
				ElseIf (oAgent.ySpilledData And eSpilledData.OwnerName) = 0 Then
					'owner
					If lRoll < ml_OWNER_CHANCE Then oAgent.ySpilledData = oAgent.ySpilledData Or eSpilledData.OwnerName
				ElseIf (oAgent.ySpilledData And eSpilledData.CurrentMission) = 0 Then
					'current mission
					If lRoll < ml_MISSION_CHANCE Then oAgent.ySpilledData = oAgent.ySpilledData Or eSpilledData.CurrentMission
				ElseIf (oAgent.ySpilledData And eSpilledData.MissionTarget) = 0 Then
					'mission target
					If lRoll < ml_TARGET_CHANCE Then oAgent.ySpilledData = oAgent.ySpilledData Or eSpilledData.MissionTarget
				ElseIf (oAgent.ySpilledData And eSpilledData.Accomplices) = 0 Then
					'accomplices
					If lRoll < ml_ACCOMPLICE_CHANCE Then
						oAgent.ySpilledData = oAgent.ySpilledData Or eSpilledData.Accomplices
                        If oAgent.oMission Is Nothing = False Then
                            oAgent.oMission.AddAccomplicesToIntel(Me.oParentPlayer, oAgent.ObjectID)
                        End If
					End If
				Else : Exit For
				End If
            Next X

            'now, check on additional data
            If Rnd() * 100 < 20 Then
                'ok, give some additional data...
                lRoll = CInt(Rnd() * 100)

                Dim sMsgBody As String = ""

                If lRoll < 5 Then
                    'tech knowledge - 5%
                    Dim lTechID As Int32 = -1
                    Dim iTechTypeID As Int16 = -1
                    oAgent.oOwner.GetRandomComponentTech(lTechID, iTechTypeID)
                    If lTechID <> -1 AndAlso iTechTypeID <> -1 Then
                        Dim oTech As Epica_Tech = oAgent.oOwner.GetTech(lTechID, iTechTypeID)
                        If oTech Is Nothing = False Then
                            Dim oPTK As PlayerTechKnowledge = PlayerTechKnowledge.CreateAndAddToPlayer(Me.oParentPlayer, oTech, PlayerTechKnowledge.KnowledgeType.eSettingsLevel1, False)
                            If oPTK Is Nothing = False Then
                                oPTK.SaveObject()
                                oPTK.SendMsgToPlayer()

                                If (Me.oParentPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                                    sMsgBody = "While interrogating " & BytesToString(oAgent.AgentName) & ", we have gained knowledge in a component belonging to " & oAgent.oOwner.sPlayerNameProper & ". The component is " & BytesToString(oTech.GetTechName()) & ". You can find the intel report in the agents area."
                                End If
                            End If
                        End If
                    End If
                ElseIf lRoll < 11 Then
                    'agent data - 6%
                    Dim lExposed As Int32 = oAgent.oOwner.GetRandomAgentID(False)
                    If lExposed > 0 Then
                        Dim oExposed As Agent = GetEpicaAgent(lExposed)
                        If oExposed Is Nothing = False Then
                            Dim oPII As PlayerItemIntel = Me.oParentPlayer.AddPlayerItemIntel(oExposed.ObjectID, oExposed.ObjTypeID, PlayerItemIntel.PlayerItemIntelType.eExistance, oAgent.oOwner.ObjectID)
                            If oPII Is Nothing = False Then
                                oPII.SaveObject()
                                Me.oParentPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)

                                If (Me.oParentPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                                    sMsgBody = "While interrogating " & BytesToString(oAgent.AgentName) & ", we have gained knowledge of a possible accomplice. We only got the name of this other agent which is: " & BytesToString(oExposed.AgentName) & "."
                                End If
                            End If
                        End If
                    End If
                ElseIf lRoll < 26 Then
                    'facility loc - 15%
                    Dim oEntity As Epica_Entity = oAgent.oOwner.GetRandomUnitOrFacility(1, 0)
                    If oEntity Is Nothing = False AndAlso oEntity.ObjTypeID = ObjectType.eFacility Then
                        Dim oPII As PlayerItemIntel = Me.oParentPlayer.AddPlayerItemIntel(oEntity.ObjectID, oEntity.ObjTypeID, PlayerItemIntel.PlayerItemIntelType.eLocation, oAgent.oOwner.ObjectID)
                        If oPII Is Nothing = False Then
                            oPII.lValue = CType(oEntity, Facility).EntityDef.ModelID
                            oPII.SaveObject()
                            Me.oParentPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)

                            If (Me.oParentPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                                sMsgBody = "While interrogating " & BytesToString(oAgent.AgentName) & ", they slipped and mentioned a facility. We have gathered the data into an intelligence report that you can view. The facility is named: " & BytesToString(oEntity.EntityName) & "."
                            End If
                        End If
                    End If
                ElseIf lRoll < 61 Then
                    'player stats - 35%
                    Dim oPI As PlayerIntel = Me.oParentPlayer.GetOrAddPlayerIntel(oAgent.oOwner.ObjectID, False)

                    Dim lWhichScore As Int32 = CInt(Rnd() * 120) \ 20
                    Dim sScoreName As String = ""
                    Dim yTypeID As Byte = 0
                    Select Case lWhichScore
                        Case 0
                            oPI.PopulationScore = oAgent.oOwner.PopulationScore
                            oPI.PopulationUpdate = GetDateAsNumber(Now)
                            yTypeID = 3
                            sScoreName = "Population"
                        Case 1
                            oPI.DiplomacyScore = oAgent.oOwner.DiplomacyScore
                            oPI.DiplomacyUpdate = GetDateAsNumber(Now)
                            yTypeID = 6
                            sScoreName = "Diplomacy"
                        Case 2
                            oPI.MilitaryScore = oAgent.oOwner.lMilitaryScore \ 50
                            oPI.MilitaryUpdate = GetDateAsNumber(Now)
                            yTypeID = 5
                            sScoreName = "Military"
                        Case 3
                            oPI.ProductionScore = oAgent.oOwner.ProductionScore
                            oPI.ProductionUpdate = GetDateAsNumber(Now)
                            yTypeID = 4
                            sScoreName = "Production"
                        Case 4
                            oPI.TechnologyScore = oAgent.oOwner.TechnologyScore
                            oPI.TechnologyUpdate = GetDateAsNumber(Now)
                            yTypeID = 1
                            sScoreName = "Technology"
                        Case Else
                            oPI.WealthScore = oAgent.oOwner.WealthScore
                            oPI.WealthUpdate = GetDateAsNumber(Now)
                            yTypeID = 2
                            sScoreName = "Wealth"
                    End Select
                    If yTypeID <> 0 Then
                        'goMsgSys.SendRequestGlobalPlayerScores(oAgent.oOwner.ObjectID, yTypeID, Me.oParentPlayer.ObjectID)
                    End If

                    oPI.SaveObject()

                    Dim bCausedSpill As Boolean = False
                    If (oAgent.ySpilledData And eSpilledData.OwnerName) = 0 Then
                        oAgent.ySpilledData = oAgent.ySpilledData Or eSpilledData.OwnerName
                        bCausedSpill = True
                    End If
                    If (Me.oParentPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                        sMsgBody = "While interrogating " & BytesToString(oAgent.AgentName) & ", we gained knowledge of "
                        If bCausedSpill = True Then sMsgBody &= "who the agent works for as well as "
                        sMsgBody &= " current " & sScoreName & " intelligence scores for " & oAgent.oOwner.sPlayerNameProper & "."
                    End If

                    Me.oParentPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPI, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
                Else
                    'colony loc - 39%
                    Dim lColonyID As Int32 = oAgent.oOwner.GetRandomColonyID
                    If lColonyID > -1 Then
                        Dim oColony As Colony = GetEpicaColony(lColonyID)
                        If oColony Is Nothing = False Then
                            Dim bCausedSpill As Boolean = False
                            If (oAgent.ySpilledData And eSpilledData.OwnerName) = 0 Then
                                oAgent.ySpilledData = oAgent.ySpilledData Or eSpilledData.OwnerName
                                bCausedSpill = True
                            End If

                            Dim oPII As PlayerItemIntel = Me.oParentPlayer.AddPlayerItemIntel(oColony.ObjectID, oColony.ObjTypeID, PlayerItemIntel.PlayerItemIntelType.eLocation, oAgent.oOwner.ObjectID)
                            If oPII Is Nothing = False Then
                                oPII.SaveObject()
                                Me.oParentPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)

                                If (Me.oParentPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                                    Dim sParentName As String = ""
                                    If oColony.ParentObject Is Nothing = False Then
                                        Dim yTmpVal() As Byte = GetEpicaObjectName(CType(oColony.ParentObject, Epica_GUID).ObjTypeID, oColony.ParentObject)
                                        If yTmpVal Is Nothing = False Then sParentName = BytesToString(yTmpVal)
                                    End If
                                    sMsgBody = "While interrogating " & BytesToString(oAgent.AgentName) & ", mention of a colony "
                                    If sParentName <> "" Then
                                        sMsgBody &= " located at " & sParentName & " "
                                    End If
                                    sMsgBody &= "named " & BytesToString(oColony.ColonyName) & "."
                                    If bCausedSpill = True Then
                                        sMsgBody &= " Our intelligence reports that colony belonging to " & oAgent.oOwner.sPlayerNameProper & "."
                                    End If
                                    sMsgBody &= " You can view the full intelligence report in the agent area."
                                End If
                            End If
                        End If
                    End If
                End If

                If sMsgBody <> "" Then
                    Dim oPC As PlayerComm = Me.oParentPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sMsgBody, "Interrogation Report", Me.oParentPlayer.ObjectID, GetDateAsNumber(Now), False, Me.oParentPlayer.sPlayerNameProper, Nothing)
                    If oPC Is Nothing = False Then
                        Me.oParentPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                    End If
                End If
            End If

            If oAgent.ySpilledData <> yOriginalSpill Then bSendUpdate = True
        ElseIf lInterrogatorSuccess = 0 AndAlso lAgentSuccess > 0 Then
            'does the agent have Escape Artist?
            lToHit = oAgent.GetSkillValue(lSkillHardcodes.eEscapeArtist, False, 0)
            If lToHit <> 0 Then
                bSendUpdate = True
                lRoll = CInt(Rnd() * 100)
                If lRoll < lToHit Then
                    Dim bIntKilled As Boolean = False
                    If oInterrogator.HandleDaggerTest(oAgent, eDaggerTestType.AvoidCapture) < 0 Then
                        'Kill Interrogator
                        oInterrogator.KillMe()
                        bIntKilled = True
                    End If

                    'Agent is no longer captured
                    'If (oAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then oAgent.lAgentStatus = oAgent.lAgentStatus Xor AgentStatus.HasBeenCaptured
                    'Agent is a fugitive (suspicion = 200)
                    'oAgent.Suspicion = 200
                    'Agent Infiltration = 0
                    'oAgent.InfiltrationLevel = 0

                    'Survival gives a bonus (possibly escape)

                    oAgent.SetAsFugitive(oInterrogator.oOwner, False)
                    Dim yMsg(6) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eCaptureKillAgent).CopyTo(yMsg, 0)
                    System.BitConverter.GetBytes(oAgent.ObjectID).CopyTo(yMsg, 2)
                    yMsg(6) = 3
                    oInterrogator.oOwner.SendPlayerMessage(yMsg, False, AliasingRights.eViewAgents)

                    'Capturer is notified
                    If (oInterrogator.oOwner.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                        Dim sAgent As String = BytesToString(oAgent.AgentName)
                        Dim sIntDeath As String = ""
                        If bIntKilled = True Then
                            sIntDeath = vbCrLf & vbCrLf & "The agent interrogating " & sAgent & ", " & BytesToString(oInterrogator.AgentName) & " was killed during the escape."
                        End If
                        Dim oPC As PlayerComm = oInterrogator.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                           sAgent & " has escaped captivity. Somehow the fugitive broke through security and moved outside of the confines of the compound." & _
                           vbCrLf & vbCrLf & "Authorities have been alerted and are in pursuit." & sIntDeath, "Prisoner Escape", oInterrogator.oOwner.ObjectID, _
                           GetDateAsNumber(Now), False, oInterrogator.oOwner.sPlayerNameProper, Nothing)
                        If oPC Is Nothing = False Then
                            oInterrogator.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                        End If
                    End If
                Else
                    'agent takes half damage
                    Dim lTmpDmg As Int32 = oAgent.yHealth
                    lTmpDmg \= 2
                    If lTmpDmg < 1 Then lTmpDmg = 1
                    oAgent.yHealth = CByte(lTmpDmg)
                End If
            End If
        End If

        'Ok, now, test for agent damage...
        Dim fDmg As Single = 0.0F
        For X As Int32 = 0 To lUB
            If lIMB(X) <> 0 Then
                fDmg += GetAgentDmg(X, lASB(X), lIMB(X))
            End If
        Next X
        Dim lTotalDmg As Int32 = CInt(Math.Ceiling(fDmg))
        If lTotalDmg > 0 Then
            bSendUpdate = True
            Dim lHP As Int32 = CInt(oAgent.yHealth) - lTotalDmg
            If lHP < 0 Then lHP = 0
            If lHP > 100 Then lHP = 100
            oAgent.yHealth = CByte(lHP)
        End If

        If oAgent.yHealth < 1 Then
            oAgent.KillMe()
            Dim oPC As PlayerComm = oAgent.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
             "Intel reports that our agent " & BytesToString(oAgent.AgentName) & " is dead by the hands of " & _
             Me.oParentPlayer.sPlayerNameProper & ". It appears that some questionable interrogation methods were conducted which resulted in the agent's death.", _
             "Agent Killed", oAgent.oOwner.ObjectID, GetDateAsNumber(Now), False, oAgent.oOwner.sPlayerNameProper, Nothing)
            If oPC Is Nothing = False Then oAgent.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            oPC = Nothing
            oPC = Me.oParentPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, "Your agent " & BytesToString(oInterrogator.AgentName) & _
               " regrets to inform you that the agent " & BytesToString(oAgent.AgentName) & " whom " & BytesToString(oInterrogator.AgentName) & _
               " was interrogating is dead. The death was an accident and happened during the interrogation process." & vbCrLf & vbCrLf & _
               "This information has leaked publicly and your agent network is sure that the government responsible for " & _
               BytesToString(oAgent.AgentName) & " is aware of their death.", "Prisoner Killed", Me.oParentPlayer.ObjectID, _
               GetDateAsNumber(Now), False, Me.oParentPlayer.sPlayerNameProper, Nothing)
            If oPC Is Nothing = False Then oParentPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
        End If

        'ok, check bsendupdate
        'If bSendUpdate = True Then
        If Me.oParentPlayer.lConnectedPrimaryID > -1 OrElse oParentPlayer.HasOnlineAliases(AliasingRights.eViewAgents) = True Then
            oParentPlayer.SendPlayerMessage(oAgent.GetCapturedAgentMsg, False, AliasingRights.eViewAgents)
        End If
        'End If

    End Sub

    Public Function GetAgencyAnalysis() As String
        Dim oSB As New System.Text.StringBuilder
        oSB.AppendLine("Your agents have finished analyzing the counter agent network of " & oParentPlayer.sPlayerNameProper & ". As of the date and time of this message, here are the assessments:")
        oSB.AppendLine()

        Dim bResistHeader As Boolean = False
        If lInfiltrationResistance Is Nothing = False Then
            For X As Int32 = 0 To lInfiltrationResistance.GetUpperBound(0)
                If lInfiltrationResistance(X) <> 0 Then
                    If bResistHeader = False Then
                        bResistHeader = True
                        oSB.AppendLine("Empire Wide Resistances (Lower is easier):")
                    End If
                    oSB.AppendLine("  " & GetInfiltrationTypeName(CType(X, eInfiltrationType)) & ": " & lInfiltrationResistance(X))
                End If
            Next X
            If bResistHeader = True Then oSB.AppendLine()
        End If

        Dim lCnts(eInfiltrationType.eLastInfiltrationType - 1) As Int32
        For X As Int32 = 0 To lCounterAgentUB
            If lCounterAgentIdx(X) <> -1 Then
                Dim oAgent As Agent = oCounterAgents(X)
                If oAgent Is Nothing = False Then
                    If (oAgent.lAgentStatus And AgentStatus.CounterAgent) <> 0 AndAlso (oAgent.lAgentStatus And (AgentStatus.HasBeenCaptured Or AgentStatus.IsDead Or AgentStatus.ReturningHome)) = 0 Then
                        lCnts(oAgent.InfiltrationType) += 1
                    End If
                End If
            End If
        Next X

        Dim bCounterHeader As Boolean = False
        For X As Int32 = 0 To lCnts.GetUpperBound(0)
            If lCnts(X) > 0 Then
                If bCounterHeader = False Then
                    bCounterHeader = True
                    oSB.AppendLine("Counter Agent Assignments:")
                End If
                oSB.AppendLine("  " & GetInfiltrationTypeName(CType(X, eInfiltrationType)) & ": " & lCnts(X))
            End If
        Next X

        If bCounterHeader = False AndAlso bResistHeader = False Then
            oSB.AppendLine("The target empire looks completely defenseless. Their civilization has fallen into complacency and would be easy targets for insurgency.")
        End If

        Return oSB.ToString

    End Function

    Private Shared Function GetInfiltrationTypeName(ByVal yType As eInfiltrationType) As String
        Select Case yType
            Case eInfiltrationType.eAgencyInfiltration
                Return "Agency"
            Case eInfiltrationType.eCapitalShipInfiltration
                Return "Capital Ship"
            Case eInfiltrationType.eColonialInfiltration
                Return "Colonial"
            Case eInfiltrationType.eCombatUnitInfiltration
                Return "Combat Unit"
            Case eInfiltrationType.eCorporationInfiltration
                Return "Corporation"
            Case eInfiltrationType.eFederalInfiltration
                Return "Goverment"
            Case eInfiltrationType.eGeneralInfiltration
                Return "General"
            Case eInfiltrationType.eMilitaryInfiltration
                Return "Military"
            Case eInfiltrationType.eMiningInfiltration
                Return "Mining"
            Case eInfiltrationType.ePlanetInfiltration
                Return "Planet"
            Case eInfiltrationType.ePowerCenterInfiltration
                Return "Power Center"
            Case eInfiltrationType.eProductionInfiltration
                Return "Production"
            Case eInfiltrationType.eResearchInfiltration
                Return "Research"
            Case eInfiltrationType.eSolarSystemInfiltration
                Return "Solar System"
            Case eInfiltrationType.eTradeInfiltration
                Return "Trade"
            Case Else
                Return "Unknown"
        End Select
    End Function

    Public Sub AlertInfiltratedOfAudit()
        Try
            Dim lAlertedPlayers(-1) As Int32
            Dim lCurUB As Int32 = -1
            If glAgentIdx Is Nothing = False Then lCurUB = Math.Min(glAgentUB, glAgentIdx.GetUpperBound(0))
            Dim lMePlayerID As Int32 = Me.oParentPlayer.ObjectID
            For X As Int32 = 0 To lCurUB
                If glAgentIdx(X) > 0 Then
                    Dim oAgent As Agent = goAgent(X)
                    If oAgent Is Nothing = False AndAlso oAgent.oOwner.ObjectID <> lMePlayerID AndAlso oAgent.lTargetID = lMePlayerID AndAlso oAgent.iTargetTypeID = ObjectType.ePlayer Then
                        If (oAgent.lAgentStatus And AgentStatus.IsInfiltrated) <> 0 Then
                            If oAgent.InfiltrationLevel > 150 Then
                                If CInt(Rnd() * 100) < CInt(oAgent.InfiltrationLevel) - 150I Then
                                    Dim bFound As Boolean = False
                                    For Y As Int32 = 0 To lAlertedPlayers.GetUpperBound(0)
                                        If lAlertedPlayers(Y) = oAgent.oOwner.ObjectID Then
                                            bFound = True
                                            Exit For
                                        End If
                                    Next Y
                                    If bFound = False Then
                                        ReDim Preserve lAlertedPlayers(lAlertedPlayers.GetUpperBound(0) + 1)
                                        lAlertedPlayers(lAlertedPlayers.GetUpperBound(0)) = oAgent.oOwner.ObjectID
                                        If (oAgent.oOwner.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                                            Dim oPC As PlayerComm = oAgent.oOwner.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, _
                                                "Your agent, " & BytesToString(oAgent.AgentName) & ", infiltrated within " & Me.oParentPlayer.sPlayerNameProper & _
                                                " has gotten word that there is an audit coming soon. The audit may reveal your agent activities within the empire belonging to " & _
                                                Me.oParentPlayer.sPlayerNameProper & ".", "Audit Notification", oAgent.oOwner.ObjectID, GetDateAsNumber(Now), False, oAgent.oOwner.sPlayerNameProper, Nothing)
                                            If oPC Is Nothing = False Then
                                                oAgent.oOwner.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next X
        Catch
        End Try
    End Sub

    Public Sub PerformAudit()
        If oParentPlayer Is Nothing Then Return

        Dim lBudgetItems As Int32 = 0
        Dim lColonyItems As Int32 = 0
        Dim lFacilityItems As Int32 = 0
        Dim lAgentItems As Int32 = 0

        Dim oSB As New System.Text.StringBuilder()
        oSB.AppendLine("The audit on your empire to discover and remove any negative effects placed on you by enemy agents is over.")
        oSB.AppendLine()
        oSB.AppendLine("The following items were discovered:")

        'audit removes many negative effects currently in place on the player
        If oParentPlayer.oBudget Is Nothing = False Then
            '  75% budget items
            lBudgetItems = oParentPlayer.oBudget.AuditRemoveAgentEffects(75)
            If lBudgetItems > 0 Then
                oSB.AppendLine("   " & lBudgetItems.ToString & " effects on our imperial budget." & vbCrLf)
            End If
        End If

        '  75% colony items except governor morale
        '  85% facility production items
        For X As Int32 = 0 To oParentPlayer.mlColonyUB
            If oParentPlayer.mlColonyIdx(X) > -1 AndAlso glColonyIdx(oParentPlayer.mlColonyIdx(X)) = oParentPlayer.mlColonyID(X) Then
                Dim oColony As Colony = goColony(oParentPlayer.mlColonyIdx(X))
                If oColony Is Nothing = False Then
                    Dim lColTemp As Int32 = oColony.AuditRemoveEffects(75)
                    lColonyItems += lColTemp
                    Dim lFacTemp As Int32 = oColony.AuditRemoveFactoryProdEffects(85)
                    lFacilityItems += lFacTemp

                    If lColTemp > 0 Then oSB.AppendLine("   " & lColTemp.ToString & " effects on the " & BytesToString(oColony.ColonyName) & " colony.")
                    If lFacTemp > 0 Then oSB.AppendLine("   " & lFacTemp.ToString & " effects on facilities within the " & BytesToString(oColony.ColonyName) & " colony.")
                End If
            End If
        Next X

        'furthermore, it attempts to route out infiltrated agents
        Try
            Dim lCurUB As Int32 = -1
            If glAgentIdx Is Nothing = False Then lCurUB = Math.Min(glAgentUB, glAgentIdx.GetUpperBound(0))
            Dim lMePlayerID As Int32 = Me.oParentPlayer.ObjectID
            For X As Int32 = 0 To lCurUB
                If glAgentIdx(X) > 0 Then
                    Dim oAgent As Agent = goAgent(X)
                    If oAgent Is Nothing = False AndAlso oAgent.oOwner.ObjectID <> lMePlayerID AndAlso oAgent.lTargetID = lMePlayerID AndAlso oAgent.iTargetTypeID = ObjectType.ePlayer Then
                        If (oAgent.lAgentStatus And AgentStatus.IsInfiltrated) <> 0 Then
                            Dim lLegit As Int32 = oAgent.GetSkillValue(lSkillHardcodes.eLegitimateCover, True, 0)
                            If lLegit > 0 Then Continue For

                            Dim lVal As Int32 = oAgent.InfiltrationLevel
                            lVal = lVal - 100
                            If oAgent.Suspicion <> 0 OrElse Rnd() * 100 > lVal Then
                                'agent gets to roll against...
                                'dagger
                                lVal = oAgent.Dagger
                                If Rnd() * 100 < lVal Then Continue For 'avoided capture

                                'resourcefulness
                                lVal = oAgent.Resourcefulness
                                If Rnd() * 100 < lVal - 5 Then Continue For 'avoided capture

                                'luck
                                lVal = oAgent.Luck
                                If Rnd() * 100 < lVal Then Continue For 'avoided capture

                                'disguises -15 infiltration level and +20 suspicion
                                lVal = oAgent.GetSkillValue(lSkillHardcodes.eDisguises, True, 0)
                                If lVal > 0 Then
                                    If Rnd() * 100 < lVal Then
                                        lVal = oAgent.InfiltrationLevel
                                        lVal -= 15
                                        If lVal < 1 Then lVal = 1
                                        oAgent.InfiltrationLevel = CByte(lVal)
                                        lVal = oAgent.Suspicion
                                        lVal += CInt(Rnd() * 15) + 5
                                        If lVal > 255 Then lVal = 255
                                        oAgent.Suspicion = CByte(lVal)
                                        Continue For
                                    End If
                                End If

                                'escape artist -30 infiltration level and +40 suspicion
                                lVal = oAgent.GetSkillValue(lSkillHardcodes.eEscapeArtist, True, 0)
                                If lVal > 0 Then
                                    If Rnd() * 100 < lVal Then
                                        lVal = oAgent.InfiltrationLevel
                                        lVal -= 30
                                        If lVal < 1 Then lVal = 1
                                        oAgent.InfiltrationLevel = CByte(lVal)
                                        lVal = oAgent.Suspicion
                                        lVal += CInt(Rnd() * 25) + 15
                                        If lVal > 255 Then lVal = 255
                                        oAgent.Suspicion = CByte(lVal)
                                        Continue For
                                    End If
                                End If

                                'im caught
                                lAgentItems += 1
                                CaptureAgent(oAgent)
                            End If
                        End If
                    End If
                End If
            Next X
        Catch
        End Try

        'Ok, create our email and send it...
        If lAgentItems > 0 Then
            oSB.AppendLine("We captured " & lAgentItems.ToString & " enemy agents that were infiltrated within our empire.")
        End If
        If lBudgetItems = 0 AndAlso lColonyItems = 0 AndAlso lFacilityItems = 0 AndAlso lAgentItems = 0 Then
            oSB.AppendLine("  No indications that our empire is being negatively affected by insurgency could be detected.")
        Else
            oSB.AppendLine()
            oSB.AppendLine("While this is an exhaustive list for this audit, it is not an exhaustive list of insurgency activity within your empire. Further audit missions may be necessary.")
        End If

        If (oParentPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
            Dim oPC As PlayerComm = oParentPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Audit Completed", oParentPlayer.ObjectID, GetDateAsNumber(Now), False, oParentPlayer.sPlayerNameProper, Nothing)
            If oPC Is Nothing = False Then
                oParentPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
            End If
        End If
    End Sub
End Class