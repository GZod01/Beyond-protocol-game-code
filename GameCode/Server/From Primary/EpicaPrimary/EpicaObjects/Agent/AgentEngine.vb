Public Class AgentEngine
	Public Enum EventTypeID As Int32
		AgentDeinfiltrateComplete
		AgentDismissed
		AgentFirstInfiltrate
		AgentReInfiltrate
		ProcessMission	'mission is underway, reprocess it causes the mission to do its deal
		AgentCheckIn
		InterrogateTest
		ReduceSuspicion
        RecruitDismiss  'a recruit has been a recruit for 48 hours, remove the recruit
        PrisonEscapeTest
	End Enum

	Private Structure uAgentEvent
		Public lEventTypeID As EventTypeID
		Public oAgent As Agent					'reference to the agent
		Public oMission As PlayerMission		'reference to the mission
		Public lAgentID As Int32
		Public lPM_ID As Int32
	End Structure

	Private moEvents() As uAgentEvent
	Private mlEventCycle() As Int32
	Private mlEventUB As Int32 = -1

	Public Const ml_MISSION_CHECK_DELAY As Int32 = 900	'30 seconds
	Private mlLastMissionStartCheck As Int32 = -1

	Public Function AddAgentEvent(ByVal lTypeID As EventTypeID, ByRef oAgent As Agent, ByRef oMission As PlayerMission, ByVal lStartCycle As Int32) As Int32

		'LogEvent(LogEventType.Informational, "AddAgentEvent: " & lTypeID.ToString)

		Dim lAgentID As Int32 = -1
		Dim lPM_ID As Int32 = -1
		If oAgent Is Nothing = False Then lAgentID = oAgent.ObjectID
		If oMission Is Nothing = False Then lPM_ID = oMission.PM_ID

		Dim lIdx As Int32 = -1
		Try
			For X As Int32 = 0 To mlEventUB
				If mlEventCycle(X) <> Int32.MinValue Then
					With moEvents(X)
						If .lEventTypeID = lTypeID AndAlso .lAgentID = lAgentID AndAlso .lPM_ID = lPM_ID Then
							Return X
						End If
					End With
				ElseIf lIdx = -1 Then
					lIdx = X
				End If
			Next X
		Catch
			lIdx = -1
		End Try
		If lIdx = -1 Then
			lIdx = mlEventUB + 1
			ReDim Preserve moEvents(lIdx)
			ReDim Preserve mlEventCycle(lIdx)
			mlEventUB += 1
		End If
		mlEventCycle(lIdx) = Int32.MaxValue
		With moEvents(lIdx)
			.lEventTypeID = lTypeID
			.oAgent = oAgent
			.oMission = oMission
			.lAgentID = lAgentID
			.lPM_ID = lPM_ID
		End With
		mlEventCycle(lIdx) = lStartCycle

		Return lIdx
	End Function

	Public Sub ProcessQueue()

		If glCurrentCycle - mlLastMissionStartCheck > 90 Then		'3 seconds
			mlLastMissionStartCheck = glCurrentCycle
			For X As Int32 = 0 To glPlayerMissionUB
				If glPlayerMissionIdx(X) <> -1 AndAlso goPlayerMission(X).lCurrentPhase = eMissionPhase.eWaitingToExecute Then
					'ok, check our assignments
					If goPlayerMission(X).AgentsReadyToExecuteMission = True Then
						'ok, let's do it
						goPlayerMission(X).MarkAgentsOnAMission()
                        goPlayerMission(X).lCurrentPhase = eMissionPhase.ePreparationTime

                        If goPlayerMission(X).oMission Is Nothing = False AndAlso goPlayerMission(X).oMission.ProgramControlID = eMissionResult.eAudit Then
                            If goPlayerMission(X).bAuditAlertSent = False Then
                                goPlayerMission(X).bAuditAlertSent = True
                                goPlayerMission(X).oPlayer.oSecurity.AlertInfiltratedOfAudit()
                            End If
                        End If

                        If goPlayerMission(X).oPlayer.lConnectedPrimaryID > -1 OrElse goPlayerMission(X).oPlayer.HasOnlineAliases(AliasingRights.eViewAgents) = True Then
                            goPlayerMission(X).oPlayer.SendPlayerMessage(goPlayerMission(X).GetAddObjectMessage, False, AliasingRights.eViewAgents)
                        End If

						If goPlayerMission(X).oPlayer.yPlayerPhase = eyPlayerPhase.eInitialPhase Then
							'ok, create the pirate factory and begin building units out of it
							CreatePirateFacility(goPlayerMission(X).oPlayer.ObjectID, goPlayerMission(X).oPlayer.lStartedEnvirID)
						End If

						AddAgentEvent(EventTypeID.ProcessMission, Nothing, goPlayerMission(X), glCurrentCycle)
					End If
				End If
			Next X
		End If

		For X As Int32 = 0 To mlEventUB
			If mlEventCycle(X) <> Int32.MinValue AndAlso glCurrentCycle > mlEventCycle(X) Then
				Dim uEvent As uAgentEvent = moEvents(X)

				With uEvent
					Select Case .lEventTypeID
						Case EventTypeID.AgentDeinfiltrateComplete
							If .oAgent Is Nothing = False Then
								'LogEvent(LogEventType.Informational, "AE.ProcessQueue: Agent Returning Home Complete: " & BytesToString(.oAgent.AgentName))
								.oAgent.lAgentStatus = .oAgent.lAgentStatus Xor AgentStatus.ReturningHome
								Dim yResp(9) As Byte
								System.BitConverter.GetBytes(GlobalMessageCode.eDeinfiltrateAgent).CopyTo(yResp, 0)
								System.BitConverter.GetBytes(.oAgent.ObjectID).CopyTo(yResp, 2)
								System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yResp, 6)
								.oAgent.oOwner.SendPlayerMessage(yResp, False, AliasingRights.eViewAgents)
							End If
							mlEventCycle(X) = Int32.MinValue
						Case EventTypeID.AgentDismissed
							If .oAgent Is Nothing = False Then
								'Ok, finish dismissing the agent
								If (.oAgent.lAgentStatus And AgentStatus.Dismissed) <> 0 Then
									'LogEvent(LogEventType.Informational, "AE.ProcessQueue: Agent Dismissed Complete: " & BytesToString(.oAgent.AgentName))
									'dismiss the agent... 
									Dim yResp(9) As Byte
									System.BitConverter.GetBytes(GlobalMessageCode.eSetAgentStatus).CopyTo(yResp, 0)
									System.BitConverter.GetBytes(.oAgent.ObjectID).CopyTo(yResp, 2)
									System.BitConverter.GetBytes(Int32.MaxValue).CopyTo(yResp, 6)
									.oAgent.oOwner.SendPlayerMessage(yResp, False, AliasingRights.eViewAgents)

									'and delete the agent
									.oAgent.DeleteMe()
								End If
							End If
							mlEventCycle(X) = Int32.MinValue
						Case EventTypeID.AgentFirstInfiltrate
							If .oAgent Is Nothing = False Then
								If .oAgent.lArrivalStart + .oAgent.lArrivalCycles < glCurrentCycle Then
									'LogEvent(LogEventType.Informational, "AE.ProcessQueue: Agent First Infiltrate: " & BytesToString(.oAgent.AgentName))
									Dim yResult As Byte = .oAgent.AttemptInfiltration()

									'Clear the infiltrating flags
									.oAgent.lArrivalStart = -1
									.oAgent.lArrivalCycles = -1
									.oAgent.lAgentStatus = .oAgent.lAgentStatus Xor AgentStatus.Infiltrating

									'yResult can equal: 0 = no target, 1 = no success returning home, 2 = captured, 255 = success
									' on 0 and 2, we do nothing, no target is handled already and Captured is handled in Agent.CaptureAgent()
									If yResult = 1 Then
										.oAgent.ReturnHome()
									ElseIf yResult = 255 Then

										If .oAgent.lTargetID = .oAgent.oOwner.ObjectID AndAlso .oAgent.iTargetTypeID = ObjectType.ePlayer Then
											.oAgent.lAgentStatus = .oAgent.lAgentStatus Or AgentStatus.CounterAgent
											.oAgent.oOwner.oSecurity.AddCounterAgent(.oAgent)
										ElseIf .oAgent.iTargetTypeID = ObjectType.ePlayer AndAlso .oAgent.InfiltrationType <> eInfiltrationType.eGeneralInfiltration Then
											'ok, add an event
											Me.AddAgentEvent(EventTypeID.AgentCheckIn, .oAgent, Nothing, glCurrentCycle)
										End If

										'success
										.oAgent.lAgentStatus = .oAgent.lAgentStatus Or AgentStatus.IsInfiltrated
										Dim yMsg(9) As Byte
										System.BitConverter.GetBytes(GlobalMessageCode.eSetAgentStatus).CopyTo(yMsg, 0)
										System.BitConverter.GetBytes(.oAgent.ObjectID).CopyTo(yMsg, 2)
										System.BitConverter.GetBytes(AgentStatus.IsInfiltrated).CopyTo(yMsg, 6)
										.oAgent.oOwner.SendPlayerMessage(yMsg, False, AliasingRights.eViewAgents)
										Me.AddAgentEvent(EventTypeID.AgentReInfiltrate, .oAgent, Nothing, glCurrentCycle + Agent.ml_DEINFILTRATE_TIME)
									End If
                                    .oAgent.SaveObject()
									mlEventCycle(X) = Int32.MinValue
								Else : mlEventCycle(X) = .oAgent.lArrivalStart + .oAgent.lArrivalCycles
								End If
							Else : mlEventCycle(X) = Int32.MinValue
							End If
						Case EventTypeID.AgentReInfiltrate
							Dim lNewCycle As Int32 = Int32.MinValue
							If .oAgent Is Nothing = False Then
								.oAgent.ReprocessInfiltrateTest()
								If .oAgent.InfiltrationLevel < 200 AndAlso .oAgent.iTargetTypeID = ObjectType.ePlayer AndAlso .oAgent.lTargetID > 0 Then
									lNewCycle = glCurrentCycle + Agent.ml_DEINFILTRATE_TIME
								End If
								'LogEvent(LogEventType.Informational, "AE.ProcessQueue: Agent Reinfiltrate Test : " & BytesToString(.oAgent.AgentName) & " @ " & .oAgent.InfiltrationLevel.ToString)
							End If
							mlEventCycle(X) = lNewCycle
						Case EventTypeID.ProcessMission
							Dim lNewCycle As Int32 = Int32.MinValue
							If .oMission Is Nothing = False Then
								'LogEvent(LogEventType.Informational, "AE.ProcessQueue: Process Mission")
                                .oMission.ProcessTests()
                                .oMission.SaveObject()
								If .oMission.lCurrentPhase > eMissionPhase.eInPlanning AndAlso .oMission.lCurrentPhase <= eMissionPhase.eReinfiltrationPhase Then
                                    lNewCycle = glCurrentCycle + ml_MISSION_CHECK_DELAY
                                    If gb_IS_TEST_SERVER = True Then lNewCycle = glCurrentCycle + 1
								Else
									Dim yMsg(10) As Byte
									System.BitConverter.GetBytes(GlobalMessageCode.eAgentMissionCompleted).CopyTo(yMsg, 0)
									System.BitConverter.GetBytes(.lPM_ID).CopyTo(yMsg, 2)
									System.BitConverter.GetBytes(.oMission.oMission.ObjectID).CopyTo(yMsg, 6)
									yMsg(10) = .oMission.lCurrentPhase
									.oMission.oPlayer.SendPlayerMessage(yMsg, False, AliasingRights.eViewAgents)
								End If
							End If
							mlEventCycle(X) = lNewCycle
						Case EventTypeID.AgentCheckIn
							'ok, an agent is checking in...
							Dim lNewCycle As Int32 = Int32.MinValue
							If (.oAgent.lAgentStatus And AgentStatus.IsInfiltrated) <> 0 AndAlso (.oAgent.lAgentStatus And (AgentStatus.Dismissed Or AgentStatus.IsDead Or AgentStatus.HasBeenCaptured Or AgentStatus.ReturningHome Or AgentStatus.OnAMission)) = 0 Then
								If .oAgent.HandleAgentCheckIn() = True Then
									lNewCycle = Agent.GetCycleDelayByFreq(.oAgent.yReportFreq)
									lNewCycle += glCurrentCycle
								End If
							End If
							mlEventCycle(X) = lNewCycle
						Case EventTypeID.InterrogateTest
							Dim lNewCycle As Int32 = Int32.MinValue
							If .oAgent Is Nothing = False Then
								If (.oAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 AndAlso (.oAgent.lAgentStatus And AgentStatus.IsDead) = 0 Then
									Dim lIntID As Int32 = .oAgent.lInterrogatorID
									If lIntID <> -1 Then
										Dim oInt As Agent = GetEpicaAgent(lIntID)
										If oInt Is Nothing = False Then
											'ok, has an interrogator and an agent...
											.oAgent.yInterrogationState = 2
                                            oInt.oOwner.oSecurity.HandleInterrogationTest(.oAgent, oInt)
                                            .oAgent.SaveObject()

											'If oInt.oOwner.oSecurity.HandleInterrogationTest(.oAgent, oInt) = True Then
											'lNewCycle = glCurrentCycle + ml_MISSION_CHECK_DELAY
											'End If
										End If
									End If
								End If
							End If
							mlEventCycle(X) = lNewCycle
						Case EventTypeID.ReduceSuspicion
							Dim lNewCycle As Int32 = Int32.MinValue
							If .oAgent Is Nothing = False Then
								If (.oAgent.lAgentStatus And (AgentStatus.HasBeenCaptured Or AgentStatus.IsDead Or AgentStatus.ReturningHome)) = 0 Then
									.oAgent.ReduceSuspicion()
									If .oAgent.Suspicion <> 0 Then
										lNewCycle = glCurrentCycle + Agent.ml_DEINFILTRATE_TIME
									End If
								End If
							End If
							mlEventCycle(X) = lNewCycle
						Case EventTypeID.RecruitDismiss
							Dim lNewCycle As Int32 = Int32.MinValue
							If .oAgent Is Nothing = False Then
								If (.oAgent.lAgentStatus And AgentStatus.NewRecruit) <> 0 Then
									'LogEvent(LogEventType.Informational, "AE.ProcessQueue: Agent Recruit Dismissed (24hr): " & BytesToString(.oAgent.AgentName))
									'dismiss the agent... 
									Dim yResp(9) As Byte
									System.BitConverter.GetBytes(GlobalMessageCode.eSetAgentStatus).CopyTo(yResp, 0)
									System.BitConverter.GetBytes(.oAgent.ObjectID).CopyTo(yResp, 2)
									System.BitConverter.GetBytes(Int32.MaxValue).CopyTo(yResp, 6)
									.oAgent.oOwner.SendPlayerMessage(yResp, False, AliasingRights.eViewAgents)

									'and delete the agent
									.oAgent.DeleteMe()
								End If
							End If
                            mlEventCycle(X) = lNewCycle
                        Case EventTypeID.PrisonEscapeTest
                            'Do an escape test.  If failed, hit health.  If pass, they escape.
                            Dim lNewCycle As Int32 = Int32.MinValue
                            If .oAgent Is Nothing = False Then
                                If (.oAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then
                                    If .oAgent.PrisonEscapeTest() = False Then
                                        .oAgent.lPrisonTestCycles = CInt((2052000 * Rnd()) - (540000)) '5-24hr (19*rnd+5)
                                        lNewCycle = glCurrentCycle + .oAgent.lPrisonTestCycles
                                    End If
                                End If
                            End If
                            mlEventCycle(X) = lNewCycle
						Case Else
                            mlEventCycle(X) = Int32.MinValue
					End Select
				End With
			End If
		Next X
	End Sub

	Public Sub CancelAgentDismissed(ByVal lAgentID As Int32)
		For X As Int32 = 0 To mlEventUB
			Try
				If mlEventCycle(X) <> Int32.MinValue Then
					If moEvents(X).lAgentID = lAgentID Then
						If moEvents(X).lEventTypeID = EventTypeID.AgentDismissed Then
							mlEventCycle(X) = Int32.MinValue
						End If
					End If
				End If
			Catch

			End Try
		Next X
	End Sub
	Public Sub CancelMissionEvent(ByVal lEventType As EventTypeID, ByVal lMissionID As Int32)
		For X As Int32 = 0 To mlEventUB
			Try
				If mlEventCycle(X) <> Int32.MinValue Then
					If moEvents(X).lPM_ID = lMissionID AndAlso moEvents(X).lEventTypeID = lEventType Then
						mlEventCycle(X) = Int32.MinValue
					End If
				End If
			Catch
			End Try
		Next X
	End Sub
	Public Sub CancelAgentEvent(ByVal lEventType As EventTypeID, ByVal lAgentID As Int32)
		For X As Int32 = 0 To mlEventUB
			Try
				If mlEventCycle(X) <> Int32.MinValue Then
					If moEvents(X).lAgentID = lAgentID AndAlso moEvents(X).lEventTypeID = lEventType Then
						mlEventCycle(X) = Int32.MinValue
					End If
				End If
			Catch
			End Try
		Next X
	End Sub
	Public Sub CancelAllAgentEvents(ByVal lAgentID As Int32)
		For X As Int32 = 0 To mlEventUB
			Try
				If mlEventCycle(X) <> Int32.MinValue Then
					If moEvents(X).lAgentID = lAgentID Then
						mlEventCycle(X) = Int32.MinValue
					End If
				End If
			Catch
			End Try
		Next X
	End Sub

	Public Sub CancelPlayerAgency(ByVal lPlayerID As Int32)
		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing Then Return

		'Ok... we need to clear all agents belonging to the player
		For X As Int32 = 0 To oPlayer.mlAgentUB
			Try
				If oPlayer.mlAgentIdx(X) <> -1 AndAlso glAgentIdx(oPlayer.mlAgentIdx(X)) = oPlayer.mlAgentID(X) Then
					Dim oAgent As Agent = goAgent(oPlayer.mlAgentIdx(X))
					CancelAllAgentEvents(oAgent.ObjectID)
					If oAgent Is Nothing = False Then oAgent.DeleteMe()
				End If
			Catch
			End Try

		Next X

		'  we need to cancel all missions that the player is currently running
		Try
			Dim lCurUB As Int32 = Math.Min(glPlayerMissionUB, glPlayerMissionIdx.GetUpperBound(0))
			For X As Int32 = 0 To lCurUB
				If glPlayerMissionIdx(X) <> -1 Then
					Dim oPM As PlayerMission = goPlayerMission(X)
					If oPM Is Nothing = False Then
						If oPM.oPlayer.ObjectID = lPlayerID OrElse (oPM.oTarget Is Nothing = False AndAlso oPM.oTarget.ObjectID = lPlayerID) Then
							'fail the mission
							oPM.lCurrentPhase = eMissionPhase.eCancelled
							oPM.MarkAgentsOffMission()
						End If
					End If
				End If
			Next X

            'All agents that are infiltrated on this player return to their owner
            lCurUB = -1
            If glAgentIdx Is Nothing = False Then lCurUB = Math.Min(glAgentUB, glAgentIdx.GetUpperBound(0))
			For X As Int32 = 0 To lCurUB
				If glAgentIdx(X) <> -1 Then
					Dim oAgent As Agent = goAgent(X)
					If oAgent Is Nothing = False Then
						If oAgent.lTargetID = lPlayerID AndAlso oAgent.iTargetTypeID = ObjectType.ePlayer Then
							oAgent.ReturnHome()
						End If
					End If
				End If
			Next X
		Catch
		End Try

	End Sub

	Public Sub FastForwardMissionEvent(ByVal lPM_ID As Int32)
		For X As Int32 = 0 To mlEventUB
			Try
				If mlEventCycle(X) <> Int32.MinValue Then
					If moEvents(X).lPM_ID = lPM_ID Then
						If moEvents(X).lEventTypeID = EventTypeID.ProcessMission Then
							mlEventCycle(X) = glCurrentCycle + 1
							If moEvents(X).oMission Is Nothing = False Then
								With moEvents(X).oMission
									If .oMissionGoals Is Nothing = False Then
										For Y As Int32 = 0 To .oMissionGoals.GetUpperBound(0)
											If .oMissionGoals(Y) Is Nothing = False AndAlso .oMissionGoals(Y).oAssignments Is Nothing = False Then
												For Z As Int32 = 0 To .oMissionGoals(Y).oAssignments.GetUpperBound(0)
													If .oMissionGoals(Y).oAssignments(Z) Is Nothing = False Then
                                                        .oMissionGoals(Y).oAssignments(Z).PointsAccumulated = .oMissionGoals(Y).oAssignments(Z).PointsRequired - 1S
                                                        .oMissionGoals(Y).lLastProcessTest = 0
													End If
												Next Z
											End If
										Next Y
									End If
								End With
							End If
						End If
					End If
				End If
			Catch

			End Try
		Next X
	End Sub
End Class
