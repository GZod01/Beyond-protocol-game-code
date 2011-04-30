Public Class SkillSet_Skill
    Public oSkillSet As SkillSet
    Public oSkill As Skill
    Public SuccessProgCtrlID As Int32
    Public FailureProgCtrlID As Int32
	Public PointRequirement As Int16
    ''' <summary>
    ''' Positive values that are duplicated within the same skillset indicate the SAME AGENT must be assigned to all SkillSet_Skills of this ID.
    ''' Negative values that are duplicated within the same skillset indicate that DIFFERENT agents must be assigned to all SkillSet_Skills of this ID.
    ''' 0 values follow no rule.
    ''' </summary>
    ''' <remarks></remarks>
    Public AgentGroupID As Int16            'This attribute indicates grouping
    Public ToHitModifier As Int16           'applied to using the skill

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            'Ok, attempt an update first
            sSQL = "UPDATE tblSkillSet_Skill SET SuccessProgCtrlID = " & SuccessProgCtrlID & ", FailureProgCtrlID = " & FailureProgCtrlID & _
              ", PointRequirement = " & PointRequirement & ", AgentGroupingID = " & AgentGroupID & ", ToHitModifier = " & ToHitModifier & _
              " WHERE SkillSetID = " & oSkillSet.SkillSetID & " AND SkillID = " & oSkill.ObjectID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                'Ok, try to insert then
                sSQL = "INSERT INTO tblSkillSet_Skill (SkillSetID, SkillID, SuccessProgCtrlID, FailureProgCtrlID, PointRequirement, " & _
                  "AgentGroupingID, ToHitModifier) VALUES (" & oSkillSet.SkillSetID & ", " & oSkill.ObjectID & ", " & SuccessProgCtrlID & _
                  ", " & FailureProgCtrlID & ", " & PointRequirement & ", " & AgentGroupID & ", " & ToHitModifier & ")"
                oComm.Dispose()
                oComm = Nothing
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Err.Raise(-1, "SaveObject", "No records affected when saving object!")
                End If
            End If
 
            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save skillset_skill " & oSkillSet.SkillSetID & ", " & oSkill.ObjectID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function
End Class