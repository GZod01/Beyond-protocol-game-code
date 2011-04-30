Public Class PlayerMissionGoal
    Public oMission As PlayerMission
    Public oGoal As Goal
    Public oSkillSet As SkillSet        'skillset being ran for this goal

	Public GoalSuccess As Byte = 0
	Public GoalExecuted As Byte = 0		'only set if this references a goal that is not in the preparation phase

	Public lLastProcessTest As Int32 = Int32.MinValue

    Public oAssignments() As AgentAssignment
    Public lAssignmentUB As Int32 = -1

    Public Function AddAgentAssignment(ByRef oAgent As Agent, ByRef oSkill As Skill) As AgentAssignment
		Dim iPointRequirement As Int16 = 15
		If oSkillSet Is Nothing = False Then
			For X As Int32 = 0 To oSkillSet.SkillUB
				If oSkillSet.Skills(X) Is Nothing = False AndAlso oSkillSet.Skills(X).oSkill Is Nothing = False AndAlso oSkillSet.Skills(X).oSkill.ObjectID = oSkill.ObjectID Then
					iPointRequirement = oSkillSet.Skills(X).PointRequirement
					Exit For
				End If
			Next X
		End If

		'Ensure we do not already have this skill...
		For X As Int32 = 0 To lAssignmentUB
			If oAssignments(X) Is Nothing = False AndAlso oAssignments(X).oSkill.ObjectID = oSkill.ObjectID Then
				With oAssignments(X)
					.oAgent = oAgent
					.oParent = Me
					.PointsRequired = iPointRequirement
				End With
				Return oAssignments(X)
			End If
		Next X

		lAssignmentUB += 1
		ReDim Preserve oAssignments(lAssignmentUB)
		oAssignments(lAssignmentUB) = New AgentAssignment
		With oAssignments(lAssignmentUB)
			.oAgent = oAgent
			.oParent = Me
			.oSkill = oSkill
			.PointsRequired = iPointRequirement
		End With
        Return oAssignments(lAssignmentUB)
    End Function

	Public Sub CheckAssignmentsCompleted()
		Dim bSuccess As Boolean = True
		For X As Int32 = 0 To lAssignmentUB
            If oAssignments(X).Completed = False AndAlso (oAssignments(X).yStatus And AgentAssignmentStatus.Skipped) = 0 Then Return
            If oAssignments(X).IsSuccess = False Then bSuccess = False
        Next X
		Me.GoalExecuted = 1
		If Me.oGoal.MissionPhase = eMissionPhase.ePreparationTime Then
			Me.GoalSuccess = 1
		Else
			If bSuccess = True Then Me.GoalSuccess = 1 Else Me.GoalSuccess = 0
		End If

	End Sub

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try
            'Always insert
            For X As Int32 = 0 To lAssignmentUB
                If oAssignments(X) Is Nothing = False AndAlso oAssignments(X).oAgent Is Nothing = False AndAlso oGoal Is Nothing = False AndAlso oMission Is Nothing = False AndAlso oSkillSet Is Nothing = False AndAlso oAssignments(X).oSkill Is Nothing = False Then
                    sSQL = "INSERT INTO tblPlayerMissionGoal (PM_ID, GoalID, SkillSetID, AgentID, SkillID, PointsAccumulated, AAStatus, CoveringAgentID) VALUES (" & _
                     oMission.PM_ID & ", " & oGoal.ObjectID & ", " & oSkillSet.SkillSetID & ", " & oAssignments(X).oAgent.ObjectID & _
                     ", " & oAssignments(X).oSkill.ObjectID & ", " & oAssignments(X).PointsAccumulated & ", " & oAssignments(X).yStatus & ", "
                    If oAssignments(X).oCoveringAgent Is Nothing = False Then
                        sSQL &= oAssignments(X).oCoveringAgent.ObjectID
                    Else : sSQL &= "-1"
                    End If
                    sSQL &= ")"

                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    If oComm.ExecuteNonQuery() = 0 Then
                        Err.Raise(-1, "SaveObject", "PlayerMissionGoal.SaveObject: No records affected!")
                    End If
                    oComm.Dispose()
                    oComm = Nothing
                End If
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object PlayerMissionGoal. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

End Class
