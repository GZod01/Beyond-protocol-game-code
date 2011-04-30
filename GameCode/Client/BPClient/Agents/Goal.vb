Public Class Goal
    Inherits Base_GUID

    Public Const ml_SAFEHOUSE_GOAL_ID As Int32 = 10000

    Public sGoalName As String
    Public sGoalDesc As String
    Public BaseTime As Int32
    Public RiskOfDetection As Byte
    Public MissionPhase As eMissionPhase
	Public SuccessProgCtrlID As Int32
	Public FailureProgCtrlID As Int32

    Public SkillSets() As SkillSet
    Public SkillSetUB As Int32 = -1

    Private mbHasSkillSets As Boolean = False

    Public Sub FillFromMsg(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode

        With Me
            .ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            .sGoalName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
            .BaseTime = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .RiskOfDetection = yData(lPos) : lPos += 1
            .MissionPhase = CType(yData(lPos), eMissionPhase) : lPos += 1
			.SuccessProgCtrlID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			.FailureProgCtrlID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            .sGoalDesc = GetStringFromBytes(yData, lPos, 255) : lPos += 255

            .SkillSetUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
        End With

        ReDim SkillSets(SkillSetUB)
        For X As Int32 = 0 To SkillSetUB
            SkillSets(X) = New SkillSet()
            With SkillSets(X)
                .oGoal = Me
                .SkillSetID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .ProgramControlID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .SkillUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
                ReDim .Skills(.SkillUB)
            End With
            For Y As Int32 = 0 To SkillSets(X).SkillUB
                SkillSets(X).Skills(Y) = New SkillSet_Skill()
                With SkillSets(X).Skills(Y)
                    .oSkillSet = SkillSets(X)
                    Dim lSkillID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    For Z As Int32 = 0 To glSkillUB
                        If goSkills(Z).ObjectID = lSkillID Then
                            .oSkill = goSkills(Z)
                            Exit For
                        End If
                    Next Z
                    .AgentGroupID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .ToHitModifier = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                End With
            Next Y
        Next X
    End Sub

    Public Sub AddSkillSet(ByRef oSkillSet As SkillSet)
        SkillSetUB += 1
        ReDim Preserve SkillSets(SkillSetUB)
        SkillSets(SkillSetUB) = oSkillSet
    End Sub

    Public Shared Function GetRiskAssessment(ByVal yValue As Byte) As String
        If yValue < 10 Then Return "LOW"
        If yValue < 20 Then Return "MILD"
        If yValue < 40 Then Return "HIGH"
        Return "EXTREME"
    End Function

    Public Shared Function GetRiskCaptionColor(ByVal yValue As Byte) As System.Drawing.Color
        If yValue < 10 Then Return System.Drawing.Color.FromArgb(255, 64, 128, 64)
        If yValue < 20 Then Return muSettings.InterfaceBorderColor
        If yValue < 40 Then Return System.Drawing.Color.FromArgb(255, 255, 255, 0)
        Return System.Drawing.Color.FromArgb(255, 255, 64, 64)
    End Function

    Public Shared Function GetSafehouseGoal() As Goal
        For X As Int32 = 0 To glGoalUB
            If glGoalIdx(X) = ml_SAFEHOUSE_GOAL_ID Then
                Return goGoals(X)
            End If
        Next X
        Return Nothing
    End Function

    Public Function GetSkillset(ByVal lSkillsetID As Int32) As SkillSet
        For X As Int32 = 0 To SkillSetUB
            If SkillSets(X) Is Nothing = False AndAlso SkillSets(X).SkillSetID = lSkillsetID Then Return SkillSets(X)
        Next X
        Return Nothing
    End Function
End Class
