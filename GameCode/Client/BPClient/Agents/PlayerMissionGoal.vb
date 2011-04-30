Public Class PlayerMissionGoal
    Public oMission As PlayerMission
    Public oGoal As Goal
    Public oSkillSet As SkillSet        'skillset being ran for this goal

    Public GoalSuccess As Byte = 0

    Public oAssignments() As AgentAssignment
    Public lAssignmentUB As Int32 = -1

    Public Function AddAgentAssignment(ByRef oAgent As Agent, ByRef oSkill As Skill) As AgentAssignment
        'Ensure we do not already have this skill...
        For X As Int32 = 0 To lAssignmentUB
            If oAssignments(X) Is Nothing = False AndAlso oAssignments(X).oSkill.ObjectID = oSkill.ObjectID Then
                With oAssignments(X)
                    .oAgent = oAgent
                    .oParent = Me 
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
        End With
        Return oAssignments(lAssignmentUB)
    End Function

    Public Function AgentAssignedUsingNaturalTalents(ByVal lAgentID As Int32) As Boolean
        For X As Int32 = 0 To lAssignmentUB
            If oAssignments(X) Is Nothing = False Then
                If oAssignments(X).oAgent Is Nothing = False AndAlso oAssignments(X).oAgent.ObjectID = lAgentID Then
                    If oAssignments(X).oSkill Is Nothing = False Then
                        Dim oAgent As Agent = oAssignments(X).oAgent
                        Dim lSkillID As Int32 = oAssignments(X).oSkill.ObjectID

                        Dim yProf As Byte = oAgent.GetSkillProficiency(lSkillID)
                        If yProf = 0 Then
                            'assumed to be naturally talented
                            Return True
                        End If
                    End If
                End If
            End If
        Next X
        Return False
    End Function
End Class