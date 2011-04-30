Option Strict On

Public Enum eyVoteValue As Byte
	AbstainVote = 0
	NoVote = 1
	YesVote = 2
End Enum

Public Enum eyTypeOfVote As Byte
	Charter = 0
	Membership = 1
	VotingWeights = 2
	AddRankPermissions = 3
	RemoveRankPermission = 4
    FreeFormVote = 5
End Enum

Public Class GuildVote
	Public VoteID As Int32 = -1
	Public ProposedByID As Int32 = -1
	Public dtVoteStarts As Date				'in GMT
	'Public dtVoteEnds As Date				'in GMT
	Public lVoteDuration As Int32			'in hours

	Public yTypeOfVote As Byte
	Public lSelectedItem As Int32
	Public lNewValue As Int32
	Public sNewValueText As String

	Public sSummary As String

	Public yVoteState As Byte = 0

	Public yPlayerVote As eyVoteValue = eyVoteValue.AbstainVote

	Public ReadOnly Property ListBoxText() As String
		Get
			Return dtVoteStarts.ToLocalTime.ToShortDateString & " - " & GetTypeOfVoteText(yTypeOfVote, lSelectedItem)
		End Get
	End Property

	'returns GMT time
	Public ReadOnly Property dtVoteEnds() As Date
		Get
			Return dtVoteStarts.AddHours(lVoteDuration)
		End Get
	End Property

	Public Shared Function GetTypeOfVoteText(ByVal yTypeOfVote As Byte, ByVal lSelectedItem As Int32) As String
		Select Case yTypeOfVote
			Case 0
				Select Case lSelectedItem
					Case elGuildFlags.AcceptMemberToGuild_RA
						Return "Charter - Accept Member Rights"
					Case elGuildFlags.AutomaticTradeAgreements
						Return "Charter - Open Trade"
					Case elGuildFlags.ChangeVotingWeight_RA
						Return "Charter - Change Vote Weight Rights"
					Case elGuildFlags.CreateRank_RA
						Return "Charter - Create Rank Rights"
					Case elGuildFlags.DeleteRank_RA
						Return "Charter - Delete Rank Rights"
					Case elGuildFlags.DemoteMember_RA
						Return "Charter - Demote Member Rights"
					Case elGuildFlags.ShareUnitVision
						Return "Charter - Share Unit/Facility Vision"
					Case elGuildFlags.PromoteGuildMember_RA
						Return "Charter - Promote Member Rights"
					Case elGuildFlags.RemoveGuildMember_RA
						Return "Charter - Remove Member Rights"
					Case elGuildFlags.RequirePeaceBetweenMembers
						Return "Charter - Require Peace Between Members"
                        'Case elGuildFlags.UnifiedForeignPolicy
                        '	Return "Charter - Unified Foreign Policy"
					Case -1
						Return "Charter - Disband"
					Case -2
						Return "Charter - Tax Rate Interval"
					Case 3
						Return "Charter - Vote Weight Type"
					Case Else
						Return "Charter"
				End Select
			Case 1
				Select Case lSelectedItem
					Case 1
						Return "Membership - Accept Member"
					Case 2
						Return "Membership - Demote Member"
					Case 4
						Return "Membership - Reject/Remove Member"
					Case 3
						Return "Membership - Promote Member"
					Case 5
						Return "Membership - Create Rank"
					Case 6
						Return "Membership - Delete Rank"
					Case 7 
						Return "Membership - Actively Recruiting"
					Case Else
						Return "Membership"
				End Select
			Case 2
				Dim sRankName As String = ""
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					sRankName = goCurrentPlayer.oGuild.GetRankName(lSelectedItem)
				End If
				If sRankName <> "" Then
					Return "Change Voting Weight - " & sRankName
				Else : Return "Change Voting Weight"
				End If
			Case 3
				Dim sRankName As String = ""
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					sRankName = goCurrentPlayer.oGuild.GetRankName(lSelectedItem)
				End If
				If sRankName <> "" Then
					Return "Add Rank Permission - " & sRankName
				Else : Return "Add Rank Permission"
				End If
			Case 4
				Dim sRankName As String = ""
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					sRankName = goCurrentPlayer.oGuild.GetRankName(lSelectedItem)
				End If
				If sRankName <> "" Then
					Return "Remove Rank Permission - " & sRankName
				Else : Return "Remove Rank Permission"
                End If
            Case Else
                Return ""
        End Select
		Return ""
	End Function
End Class
