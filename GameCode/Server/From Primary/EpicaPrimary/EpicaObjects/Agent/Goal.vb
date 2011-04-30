Public Enum eMissionPhase As Byte
    eInPlanning = 0
    ePreparationTime = 1
    eSettingTheStage = 2
    eFlippingTheSwitch = 3
	eReinfiltrationPhase = 4

	eWaitingToExecute = 10
	eCancelled = 11
	eMissionOverFailure = 12
	eMissionOverSuccess = 13

	eMissionPaused = 128			'logically OR'd with other values
End Enum

Public Enum eGoalProgCtrlID As int32
	ExecuteMissionProgCtrlSuccess = 1
	ExecuteMissionProgCtrlFailure = 2
	MissionFailure = 4						'mission is over due to failure, no other action is taken. No other goals are attempted.
	ThrowAlarm = 8							'mission alarm is thrown, retry of the goal is possible
	AddTargetKnowledge = 16					'knowledge of the target is added (for example, target = facility)
	ThrowAlarmPlusViolence = 32				'throw the alarm, if the alarm is already thrown, violence insues (all agents roll daggers, some may die, etc...)
	CaptureTeam = 64						'all team members active in this goal are captured (no tests required)
	CaptureTests = 128						'tests are ran for each member of the team active in this goal
	KillMembers = 256						'random tests kill off members of the team (useful for failing to exit a compound before explosions)
    EraseInfiltration = 512                 'all infiltration levels are reset to 1 for members active in this goal
    SafeHouseModifier = 1024                'safe house modifier is applied
End Enum

Public Class Goal
    Inherits Epica_GUID

    Public Const ml_SAFEHOUSE_GOAL_ID As Int32 = 10000

    Public GoalName(19) As Byte
    Public GoalDesc(254) As Byte
    Public BaseTime As Int32
    Public RiskOfDetection As Byte
    Public MissionPhase As eMissionPhase
	Public SuccessProgCtrlID As Int32
	Public FailureProgCtrlID As Int32
	Public OrderNum As Int32

    Public SkillSets() As SkillSet
    Public SkillSetUB As Int32 = -1

    Private mySendString() As Byte

    Public Function GetObjAsString() As Byte()
        'here we will return the entire object as a string
        If mbStringReady = False Then
			ReDim mySendString(298)
            Dim lPos As Int32 = 0

            GetGUIDAsString.CopyTo(mySendString, lPos) : lPos += 6
            GoalName.CopyTo(mySendString, lPos) : lPos += 20
            System.BitConverter.GetBytes(BaseTime).CopyTo(mySendString, lPos) : lPos += 4
            mySendString(lPos) = RiskOfDetection : lPos += 1
            mySendString(lPos) = MissionPhase : lPos += 1
			System.BitConverter.GetBytes(SuccessProgCtrlID).CopyTo(mySendString, lPos) : lPos += 4
			System.BitConverter.GetBytes(FailureProgCtrlID).CopyTo(mySendString, lPos) : lPos += 4
            GoalDesc.CopyTo(mySendString, lPos) : lPos += 255
            System.BitConverter.GetBytes((SkillSetUB + 1)).CopyTo(mySendString, lPos) : lPos += 4

            For X As Int32 = 0 To SkillSetUB
                Dim yTemp() As Byte = SkillSets(X).GetSkillsetAsBytes()
                If yTemp Is Nothing = False Then
                    ReDim Preserve mySendString(mySendString.GetUpperBound(0) + yTemp.Length)
                    yTemp.CopyTo(mySendString, lPos) : lPos += yTemp.Length
                End If
            Next X

            mbStringReady = True
        End If
        Return mySendString
	End Function

	'Public Function NewSaveObject() As Boolean
	'	Dim oSP As New SPMaker("sp_SaveGoal")

	'	oSP.AddParameter("GoalName", BytesToString(GoalName))
	'	oSP.AddParameter("GoalDesc", BytesToString(GoalDesc))
	'	oSP.AddParameter("BaseTime", BaseTime.ToString)
	'	oSP.AddParameter("RiskOfDetection", RiskOfDetection.ToString)
	'	oSP.AddParameter("MissionPhase", CByte(MissionPhase).ToString)
	'	oSP.AddParameter("SuccessProgCtrlID", SuccessProgCtrlID.ToString)
	'	oSP.AddParameter("FailureProgCtrlID", FailureProgCtrlID.ToString)

	'	oSP.SetResult("OutGoalID")
	'	If oSP.Execute = False Then Return False
	'	ObjectID = oSP.GetResult()
	'	oSP.DisposeMe()
	'	oSP = Nothing

	'	For X As Int32 = 0 To SkillSetUB
	'		If SkillSets(X) Is Nothing = False Then
	'			SkillSets(X).oGoal = Me
	'			SkillSets(X).SaveObject()
	'		End If
	'	Next X

	'	Return True
	'End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
				sSQL = "INSERT INTO tblGoal (GoalName, GoalDesc, BaseTime, RiskOfDetection, MissionPhase, SuccessProgCtrlID, FailureProgCtrlID) " & _
			   " VALUES ('" & MakeDBStr(BytesToString(GoalName)) & "', '" & MakeDBStr(BytesToString(GoalDesc)) & "', " & _
			   BaseTime & ", " & RiskOfDetection & ", " & CByte(MissionPhase) & ", " & SuccessProgCtrlID & ", " & FailureProgCtrlID & ")"
            Else
                'UPDATE
				sSQL = "UPDATE tblGoal SET GoalName ='" & MakeDBStr(BytesToString(GoalName)) & "', GoalDesc = '" & _
				MakeDBStr(BytesToString(GoalDesc)) & "', BaseTime = " & BaseTime & ", RiskOfDetection = " & RiskOfDetection & _
				", MissionPhase = " & CByte(MissionPhase) & ", SuccessProgCtrlID = " & SuccessProgCtrlID & ", FailureProgCtrlID = " & _
				FailureProgCtrlID & " WHERE GoalID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(GoalID) FROM tblGoal WHERE GoalName = '" & MakeDBStr(BytesToString(GoalName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            For X As Int32 = 0 To SkillSetUB
                If SkillSets(X) Is Nothing = False Then
                    SkillSets(X).oGoal = Me
                    SkillSets(X).SaveObject()
                End If
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Sub AddSkillSet(ByRef oSkillSet As SkillSet)
        SkillSetUB += 1
        ReDim Preserve SkillSets(SkillSetUB)
        SkillSets(SkillSetUB) = oSkillSet
    End Sub

    Public Function GetOrAddSkillSet(ByVal lSkillSetID As Int32, ByVal bNoAdd As Boolean) As SkillSet
        For X As Int32 = 0 To SkillSetUB
            If SkillSets(X).SkillSetID = lSkillSetID Then Return SkillSets(X)
        Next X

        If bNoAdd = True Then Return Nothing

        Dim oSkillSet As New SkillSet()
        With oSkillSet
            .oGoal = Me
            .SkillUB = -1
            .SkillSetID = lSkillSetID
        End With
        AddSkillSet(oSkillSet)
        Return oSkillSet
    End Function

	Public Sub HandleGoalCtrlID(ByRef oPM As PlayerMission, ByVal lCtrlID As Int32, ByVal lGoalIndex As Int32)
		If (lCtrlID And eGoalProgCtrlID.AddTargetKnowledge) <> 0 Then
			'always the first target is the knowledge
			With oPM

				'Dim lServerIdx As Int32 = -1
				'If .iTargetTypeID = ObjectType.eFacility Then
				'	Dim oFac As Facility = GetEpicaFacility(.lTargetID)
				'	If oFac Is Nothing = False Then lServerIdx = oFac.ServerIndex
				'	oFac = Nothing
				'ElseIf .iTargetTypeID = ObjectType.ePlayer Then
				'	Return
				'ElseIf .iTargetTypeID = ObjectType.eMineral Then
				'	Dim oMin As Mineral = GetEpicaMineral(.lTargetID)
				'	If oMin Is Nothing = False Then lServerIdx = oMin.ServerIndex
				'	oMin = Nothing
				'End If
				Dim oPII As PlayerItemIntel = .oPlayer.AddPlayerItemIntel(.lTargetID, .iTargetTypeID, PlayerItemIntel.PlayerItemIntelType.eLocation, .oTarget.ObjectID)
				If oPII Is Nothing = False Then
					.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPII, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewAgents)
				End If
			End With
		End If
		If (lCtrlID And eGoalProgCtrlID.CaptureTeam) <> 0 Then
			'Capture all agents on the mission active in this goal
			For X As Int32 = 0 To oPM.oMissionGoals(lGoalIndex).lAssignmentUB
				If oPM.oMissionGoals(lGoalIndex).oAssignments Is Nothing = False AndAlso oPM.oMissionGoals(lGoalIndex).oAssignments(X).oAgent Is Nothing = False Then
					oPM.oTarget.oSecurity.CaptureAgent(oPM.oMissionGoals(lGoalIndex).oAssignments(X).oAgent)
				End If
			Next X
		End If
		If (lCtrlID And eGoalProgCtrlID.CaptureTests) <> 0 Then
            'Run capture tests on all agents on the mission active in this goal
            If oPM.oTarget Is Nothing = False AndAlso oPM.oTarget.oSecurity Is Nothing = False Then
                Dim oSec As PlayerSecurity = oPM.oTarget.oSecurity
                For X As Int32 = 0 To oPM.oMissionGoals(lGoalIndex).lAssignmentUB
                    If oPM.oMissionGoals(lGoalIndex).oAssignments Is Nothing = False AndAlso oPM.oMissionGoals(lGoalIndex).oAssignments(X).oAgent Is Nothing = False Then
                        With oPM.oMissionGoals(lGoalIndex).oAssignments(X).oAgent
                            Dim bBreakout As Boolean = False
                            For Y As Int32 = 0 To oSec.lCounterAgentUB
                                If oSec.lCounterAgentIdx(Y) > 0 Then
                                    Dim oCounter As Agent = oSec.oCounterAgents(Y)
                                    If oCounter Is Nothing = False AndAlso oCounter.InfiltrationType = .InfiltrationType AndAlso (oCounter.lAgentStatus And (AgentStatus.HasBeenCaptured Or AgentStatus.IsDead Or AgentStatus.Infiltrating Or AgentStatus.ReturningHome)) = 0 Then
                                        Dim lResult As Int32 = .HandleDaggerTest(oCounter, eDaggerTestType.AvoidCapture)
                                        If lResult < 0 Then
                                            bBreakout = True
                                            oSec.CaptureAgent(oPM.oMissionGoals(lGoalIndex).oAssignments(X).oAgent)
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next Y
                            If bBreakout = True Then Exit For
                        End With
                    End If
                Next X
            End If
		End If
		If (lCtrlID And eGoalProgCtrlID.EraseInfiltration) <> 0 Then
			'all agents that are active on this goal have their infiltration set to 1
			For X As Int32 = 0 To oPM.oMissionGoals(lGoalIndex).lAssignmentUB
				If oPM.oMissionGoals(lGoalIndex).oAssignments Is Nothing = False AndAlso oPM.oMissionGoals(lGoalIndex).oAssignments(X).oAgent Is Nothing = False Then
					oPM.oMissionGoals(lGoalIndex).oAssignments(X).oAgent.InfiltrationLevel = 1
				End If
			Next X
		End If
		If (lCtrlID And eGoalProgCtrlID.ExecuteMissionProgCtrlFailure) <> 0 Then
            oPM.HandleMissionFailure() 
		End If
		If (lCtrlID And eGoalProgCtrlID.ExecuteMissionProgCtrlSuccess) <> 0 Then
			oPM.HandleMissionCompletion()
		End If
		If (lCtrlID And eGoalProgCtrlID.KillMembers) <> 0 Then
			'Kill all agents that are active on this goal
			Dim sNames As String = ""
			For X As Int32 = 0 To oPM.oMissionGoals(lGoalIndex).lAssignmentUB
				If oPM.oMissionGoals(lGoalIndex).oAssignments Is Nothing = False AndAlso oPM.oMissionGoals(lGoalIndex).oAssignments(X).oAgent Is Nothing = False Then
					oPM.oMissionGoals(lGoalIndex).oAssignments(X).oAgent.KillMe()
					sNames &= "  " & BytesToString(oPM.oMissionGoals(lGoalIndex).oAssignments(X).oAgent.AgentName) & vbCrLf
				End If
			Next X
			If sNames <> "" Then
				If (oPM.oPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                    Dim oPC As PlayerComm = oPM.oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, -1, _
                        "Reports have come in that the agent team working on the " & BytesToString(oPM.oMission.MissionName) & _
                      " mission were involved in some sort of catastrophic incident. The following casualties are confirmed: " & sNames, _
                      "Incident Report", oPM.oPlayer.ObjectID, GetDateAsNumber(Now), False, oPM.oPlayer.sPlayerNameProper, Nothing)
					If oPC Is Nothing = False Then oPM.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
				End If
			End If
		End If
		If (lCtrlID And eGoalProgCtrlID.MissionFailure) <> 0 Then
			oPM.HandleMissionFailure() 
		End If
		If (lCtrlID And eGoalProgCtrlID.ThrowAlarm) <> 0 Then
			If oPM.bAlarmThrown = False Then
				oPM.bAlarmThrown = True
				oPM.Modifier += -20
			End If
		End If
		If (lCtrlID And eGoalProgCtrlID.ThrowAlarmPlusViolence) <> 0 Then
			If oPM.bAlarmThrown = False Then
				oPM.bAlarmThrown = True
				oPM.Modifier += -20
			Else
                'Violence!! Some agents may die... etc...
                Dim oMeSB As New System.Text.StringBuilder
                Dim oCounterSB As New System.Text.StringBuilder

                Dim lMeAgentKill As Int32 = 0
                Dim lMeAgentHurt As Int32 = 0
                Dim lCounterAgentKill As Int32 = 0
                Dim lCounterAgentHurt As Int32 = 0
                Dim bMeHeader As Boolean = False
                Dim bCounterHeader As Boolean = False
                Dim bHadCounterAgent As Boolean = False

                Dim bCounterInjured(oPM.oTarget.oSecurity.lCounterAgentUB) As Boolean

                oMeSB.AppendLine("Your agent team tripped an alarm and soon after got in over their head. A lot of security showed up quickly and gunfire was exchanged.")
                oCounterSB.AppendLine("An enemy agent team was detected attempting to infiltrate our empire. Security was alerted, responded quickly and gunfire was exchanged.")

                For X As Int32 = 0 To oPM.oMissionGoals(lGoalIndex).lAssignmentUB
                    Dim oAssign As AgentAssignment = oPM.oMissionGoals(lGoalIndex).oAssignments(X)
                    If oAssign Is Nothing = False Then
                        If oAssign.oAgent Is Nothing = False Then
                            Dim bMeInjured As Boolean = False

                            For Y As Int32 = 0 To oPM.oTarget.oSecurity.lCounterAgentUB
                                If oPM.oTarget.oSecurity.lCounterAgentIdx(Y) > -1 Then
                                    Dim oCounter As Agent = oPM.oTarget.oSecurity.oCounterAgents(Y)
                                    If oCounter Is Nothing = False AndAlso oCounter.InfiltrationType = oPM.oMission.lInfiltrationType Then
                                        If (oCounter.lAgentStatus And AgentStatus.CounterAgent) <> 0 AndAlso (oCounter.lAgentStatus And (AgentStatus.HasBeenCaptured Or AgentStatus.IsDead Or AgentStatus.Infiltrating Or AgentStatus.ReturningHome)) = 0 Then
                                            bHadCounterAgent = True
                                            Dim lResult As Int32 = oAssign.oAgent.HandleDaggerTest(oCounter, eDaggerTestType.AvoidCapture)
                                            If lResult < 0 Then
                                                'counter won
                                                If bMeHeader = False Then
                                                    bMeHeader = True
                                                    oMeSB.AppendLine()
                                                    oMeSB.AppendLine("The following agents were injured or killed in the engagement:")
                                                End If

                                                Dim lHealth As Int32 = oAssign.oAgent.yHealth
                                                lHealth -= (CInt(Rnd() * CInt(oCounter.Dagger)) \ 10) + 1
                                                If lHealth < 1 Then
                                                    oAssign.oAgent.KillMe()
                                                    lMeAgentKill += 1
                                                    oMeSB.AppendLine("  " & BytesToString(oAssign.oAgent.AgentName) & " - Dead")
                                                Else
                                                    oAssign.oAgent.yHealth = CByte(lHealth)
                                                    bMeInjured = True
                                                End If
                                            Else
                                                If bCounterHeader = False Then
                                                    bCounterHeader = True
                                                    oCounterSB.AppendLine()
                                                    oCounterSB.AppendLine("The following agents were injured or killed in the engagement:")
                                                End If
                                                'agent won
                                                Dim lHealth As Int32 = oCounter.yHealth
                                                lHealth -= (CInt(Rnd() * CInt(oAssign.oAgent.Dagger)) \ 10) + 1
                                                If lHealth < 1 Then
                                                    oCounter.KillMe()
                                                    lCounterAgentKill += 1
                                                    oCounterSB.AppendLine("  " & BytesToString(oCounter.AgentName) & " - Dead")
                                                Else
                                                    oCounter.yHealth = CByte(lHealth)
                                                    bCounterInjured(Y) = True
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next Y

                            If bMeInjured = True Then
                                oMeSB.AppendLine("  " & BytesToString(oAssign.oAgent.AgentName) & " - Injured (" & oAssign.oAgent.yHealth & ")")
                                lMeAgentHurt += 1
                            End If
                        End If
                    End If
                Next X

                For X As Int32 = 0 To oPM.oTarget.oSecurity.lCounterAgentUB
                    If bCounterInjured(X) = True Then
                        Dim oAgent As Agent = oPM.oTarget.oSecurity.oCounterAgents(X)
                        If oAgent Is Nothing = False Then
                            If (oAgent.lAgentStatus And AgentStatus.IsDead) = 0 Then
                                oCounterSB.AppendLine("  " & BytesToString(oAgent.AgentName) & " - Injured (" & oAgent.yHealth & ")")
                                lCounterAgentHurt += 1
                            End If
                        End If
                    End If
                Next X

                If lMeAgentHurt > 0 OrElse lMeAgentKill > 0 Then
                    oCounterSB.AppendLine()
                    oCounterSB.AppendLine("The enemy agent team suffered " & lMeAgentKill & " casualties and " & lMeAgentHurt & " injured.")
                Else
                    oCounterSB.AppendLine()
                    oCounterSB.AppendLine("The enemy agent team got away unharmed.")
                End If

                If lCounterAgentHurt > 0 OrElse lCounterAgentKill > 0 Then
                    oMeSB.AppendLine()
                    oMeSB.AppendLine("The counter agents assigned to that area suffered " & lCounterAgentKill & " casualties and " & lCounterAgentHurt & " injured.")
                ElseIf bHadCounterAgent = True Then
                    oMeSB.AppendLine()
                    oMeSB.AppendLine("The counter agent group got away unharmed.")
                Else
                    oMeSB.AppendLine()
                    oMeSB.AppendLine("There was light resistance in the engagement.")
                End If

                Dim oPC As PlayerComm = Nothing
                If (oPM.oPlayer.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                    oPC = oPM.oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oMeSB.ToString, "Agent Confrontation", oPM.oPlayer.ObjectID, GetDateAsNumber(Now), False, oPM.oPlayer.sPlayerNameProper, Nothing)
                    If oPC Is Nothing = False Then oPM.oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                End If
                If oPM.oTarget Is Nothing = False AndAlso (oPM.oTarget.iInternalEmailSettings And eEmailSettings.eAgentUpdates) <> 0 Then
                    oPC = oPM.oTarget.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oCounterSB.ToString, "Agent Confrontation", oPM.oTarget.ObjectID, GetDateAsNumber(Now), False, oPM.oTarget.sPlayerNameProper, Nothing)
                    If oPC Is Nothing = False Then oPM.oTarget.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                End If

            End If
        End If
        If (lCtrlID And eGoalProgCtrlID.SafeHouseModifier) <> 0 Then
            'ok, determine which safe house modifier to use
            If oPM.oSafeHouseGoal Is Nothing = False AndAlso oPM.oSafeHouseGoal.oSkillSet Is Nothing = False Then
                Dim lVal As Int32 = oPM.oSafeHouseGoal.oSkillSet.ProgramControlID
                If lVal = 1 Then
                    'slum lords
                    Dim lRoll As Int32 = CInt(Rnd() * 100)
                    If lRoll < 30 Then
                        oPM.ySafeHouseSetting = 10
                    ElseIf lRoll < 60 Then
                        oPM.ySafeHouseSetting = 7
                    Else
                        oPM.ySafeHouseSetting = 5
                    End If
                Else
                    'living large
                    Dim lRoll As Int32 = CInt(Rnd() * 100)
                    If lRoll < 30 Then
                        oPM.ySafeHouseSetting = 20
                    ElseIf lRoll < 60 Then
                        oPM.ySafeHouseSetting = 15
                    Else
                        oPM.ySafeHouseSetting = 10
                    End If
                End If
            End If 
        End If
    End Sub

    Public Shared Function GetSafehouseGoal() As Goal
        For X As Int32 = 0 To glGoalUB
            If glGoalIdx(X) = ml_SAFEHOUSE_GOAL_ID Then
                Return goGoal(X)
            End If
        Next X
        Return Nothing
    End Function
End Class
