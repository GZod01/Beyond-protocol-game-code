Public Class SkillSet
    'this object contains several skill and base to hit combinations
    Public SkillSetID As Int32
    Public oGoal As Goal            'parent goal

    Public ProgramControlID As Int32

    Public Skills() As SkillSet_Skill
    Public SkillUB As Int32 = -1

	Public Sub AddSkill(ByRef oSkill As Skill, ByVal iToHitModifier As Int16, ByVal lSuccessProgCtrlID As Int32, ByVal lFailureProgCtrlID As Int32, ByVal iPointRequirement As Int16, ByVal iAgentGroupingID As Int16)
		SkillUB = SkillUB + 1
		ReDim Preserve Skills(SkillUB)
		Skills(SkillUB) = New SkillSet_Skill
		With Skills(SkillUB)
			.oSkillSet = Me
			.oSkill = oSkill
			.ToHitModifier = iToHitModifier
			.PointRequirement = iPointRequirement
			.SuccessProgCtrlID = lSuccessProgCtrlID
			.FailureProgCtrlID = lFailureProgCtrlID
			.AgentGroupID = iAgentGroupingID
		End With
	End Sub

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If SkillSetID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblSkillSet (GoalID, ProgramControlID) VALUES (" & oGoal.ObjectID & ", " & ProgramControlID & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblSkillSet SET GoalID = " & oGoal.ObjectID & ", ProgramControlID = " & ProgramControlID & " WHERE SkillSetID = " & Me.SkillSetID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If SkillSetID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(SkillSetID) FROM tblSkillset WHERE GoalID = " & oGoal.ObjectID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    SkillSetID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            For X As Int32 = 0 To SkillUB
                If Skills(X) Is Nothing = False Then
                    Skills(X).oSkillSet = Me
                    Skills(X).SaveObject()
                End If
            Next X
            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save skillset " & SkillSetID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Function GetSkillsetAsBytes() As Byte()
        Dim yResp(11 + ((SkillUB + 1) * 8)) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(SkillSetID).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(ProgramControlID).CopyTo(yResp, lPos) : lPos += 4
        System.BitConverter.GetBytes(SkillUB + 1).CopyTo(yResp, lPos) : lPos += 4
        For X As Int32 = 0 To SkillUB
            With Skills(X)
                System.BitConverter.GetBytes(.oSkill.ObjectID).CopyTo(yResp, lPos) : lPos += 4
                System.BitConverter.GetBytes(.AgentGroupID).CopyTo(yResp, lPos) : lPos += 2
                System.BitConverter.GetBytes(.ToHitModifier).CopyTo(yResp, lPos) : lPos += 2
            End With
        Next X
        Return yResp
    End Function
End Class
