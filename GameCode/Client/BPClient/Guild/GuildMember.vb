Option Strict On

Public Enum GuildMemberState As Byte
	Unaffiliated = 0				'nothing
	Invited = 1						'player has been invited to join the guild
	Applied = 2						'player has applied to join the guild
	Approved = 4					'application/invitation to join the guild was approved
	Rejected = 8					'application to join the guild was rejected
	AcceptedGuildFormInvite = 16	'player was invited to form the guild and accepted the invitation
End Enum

Public Class GuildMember
	Public lMemberID As Int32		'player ID of the member

	Public yMemberState As GuildMemberState = GuildMemberState.Unaffiliated

	Public sPlayerName As String = ""
	Public yPlayerTitle As Byte
    Private mlRankID As Int32 = -1
    Public Property lRankID() As Int32
        Get
            Return mlRankID
        End Get
        Set(ByVal value As Int32)
            mlRankID = value
            msMemberListText = ""
        End Set
    End Property
	Public yPlayerGender As Byte = 0

	Public dtLastOnline As Date			'in GMT
	Public dtJoined As Date				'in GMT

	Public sDetails As String = "Requesting..."

	Private mbRequestedDetails As Boolean = False
	Public ReadOnly Property DetailsText() As String
		Get
			If mbRequestedDetails = False Then
				mbRequestedDetails = True
				Dim yMsg(6) As Byte
				Dim lPos As Int32 = 0
				System.BitConverter.GetBytes(GlobalMessageCode.eGuildRequestDetails).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(Me.lMemberID).CopyTo(yMsg, lPos) : lPos += 4
				yMsg(lPos) = eyGuildRequestDetailsType.MemberDetails : lPos += 1

				If goUILib Is Nothing = False Then goUILib.SendMsgToPrimary(yMsg)
			End If
			Return sDetails
		End Get
	End Property

	Private msMemberListText As String = ""
	Public ReadOnly Property MemberListText() As String
		Get
			Dim sTemp As String = GetCacheObjectValue(lMemberID, ObjectType.ePlayer)
			If msMemberListText.Contains(sTemp) = False Then msMemberListText = ""
			If sPlayerName <> sTemp Then
				msMemberListText = ""
				sPlayerName = sTemp
			End If
			'If sPlayerName <> GetCacheObjectValue(lMemberID, ObjectType.ePlayer) Then
			'	msMemberListText = ""
			'End If
			If msMemberListText = "" Then
				msMemberListText = Player.GetPlayerNameWithTitle(yPlayerTitle, sPlayerName, yPlayerGender = 1).PadRight(30, " "c)

				Dim sRank As String = ""
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					If lRankID > -1 Then
						sRank = goCurrentPlayer.oGuild.GetRankName(lRankID)
					Else

						If (yMemberState And GuildMemberState.Applied) <> 0 Then
							sRank = "Applicant"
						ElseIf (yMemberState And GuildMemberState.Invited) <> 0 Then
							sRank = "Invited"
						End If

						If (yMemberState And GuildMemberState.Approved) <> 0 Then
							sRank &= " - Approved"
						ElseIf (yMemberState And GuildMemberState.Rejected) <> 0 Then
							sRank &= " - Rejected"
						End If
						If sRank = "" Then sRank = "Unaffiliated"
					End If
				End If
				msMemberListText &= sRank.PadRight(20, " "c)
			End If
			Return msMemberListText
		End Get
	End Property
End Class
