Public Enum AgentAssignmentStatus As Byte
	NoStatus = 0
	Skipped = 1			'agent assignment was skipped
	Finished = 2		'same as the Completed boolean server-side (indicates completion only)
	Success = 4			'indicates a success. If this flag is not set and Finished is set, then the AA was a failure
End Enum
Public Class AgentAssignment
	Public oParent As PlayerMissionGoal
	Public oAgent As Agent
	Public oSkill As Skill

	Public yStatus As Byte = 0		'uses AgentAssignmentStatus

	Public PointsRequired As Int32
	Public PointsProduced As Int32
End Class