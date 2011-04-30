Option Strict On

Public Enum RankPermissions As Int32
	AcceptApplicant = 1
	AcceptEvents = 2
	BuildGuildBase = 4
	ChangeMOTD = 8
	ChangeRankNames = 16
	ChangeRankPermissions = 32
	ChangeRankVotingWeight = 64
	ChangeRecruitment = 128
	CreateEvents = 256
	CreateRanks = 512
	DeleteEvents = 1024
	DeleteRanks = 2048
	DemoteMember = 4096
    'ChangeRankTaxRates = 8192
	PromoteMember = 16384
	ProposeVotes = 32768
	RejectMember = 65536
	RemoveMember = 131072
	ViewBankLog = 262144
	ViewContentsLowSec = 524288
	ViewContentsHiSec = 1048576
	ViewEventAttachments = 2097152
	ViewEvents = 4194304
	ViewGuildBase = 8388608
	ViewVotesHistory = 16777216
	ViewVotesInProgress = 33554432
	WithdrawLowSec = 67108864
	WithdrawHiSec = 134217728
	InviteMember = 268435456
	ModifyGuildRelation = 536870912
End Enum

Public Class GuildRank
	Public lRankID As Int32
	Public sRankName As String = "Requesting..."
	Public lVoteStrength As Int32
	Public yPosition As Byte

	Public lMembersInRank As Int32 = 0

	Public TaxRatePercentage As Byte = 0			'per interval in Guild
	Public TaxRatePercType As eyGuildTaxPercType = eyGuildTaxPercType.CashFlow
	Public TaxRateFlat As Int32 = 0					'per interval in Guild

	Public lRankPermissions As RankPermissions

	Public ReadOnly Property ListBoxDisplay() As String
		Get
			Dim sMems As String = lMembersInRank.ToString("#,##0").PadLeft(6, " "c)

			Return sRankName.PadRight(24, " "c) & sMems
		End Get
	End Property
End Class
