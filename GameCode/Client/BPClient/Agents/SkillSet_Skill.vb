Public Class SkillSet_Skill
    Public oSkillSet As SkillSet
    Public oSkill As Skill
	Public SuccessProgCtrlID As Int32
	Public FailureProgCtrlID As Int32
	Public PointRequirement As Int32		   'int32?
    ''' <summary>
    ''' Positive values that are duplicated within the same skillset indicate the SAME AGENT must be assigned to all SkillSet_Skills of this ID.
    ''' Negative values that are duplicated within the same skillset indicate that DIFFERENT agents must be assigned to all SkillSet_Skills of this ID.
    ''' 0 values follow no rule.
    ''' </summary>
    ''' <remarks></remarks>
    Public AgentGroupID As Int16            'This attribute indicates grouping
    Public ToHitModifier As Int16           'applied to using the skill
End Class