Public Class SkillSet
    'this object contains several skill and base to hit combinations
    Public SkillSetID As Int32
    Public oGoal As Goal            'parent goal

    Public sSkillSetName As String

    Public ProgramControlID As Int32

    Public Skills() As SkillSet_Skill
	Public SkillUB As Int32 = -1

	Public Function AddSkill(ByRef oSkill As Skill) As SkillSet_Skill
		SkillUB += 1
		ReDim Preserve Skills(SkillUB)
		Skills(SkillUB) = New SkillSet_Skill
		With Skills(SkillUB)
			.oSkill = oSkill
			.oSkillSet = Me
		End With

		Return Skills(SkillUB)
	End Function
End Class
