Option Strict On

Public Structure SenateVote
	Public VoterName As String
	Public yVote As eyVoteValue
End Structure

Public Structure SystemSenateVote
	Public VoterID As Int32
	Public VoterName As String
	Public yVote As eyVoteValue
	Public yRating As Byte

	Public Function GetNodeText() As String
        Return VoterName & " (" & yRating.ToString & ")"
	End Function
End Structure

Public Class SystemVote
	Public SystemID As Int32
	Public ySystemVote As eyVoteValue		'indicates how the entire system is voting
	Public yVoteRating As Byte

	Public oVotes() As SystemSenateVote

	Private mbRequestedDetails As Boolean = False
	Public Sub RequestDetails(ByVal lProposalID As Int32)
		If mbRequestedDetails = False Then
			mbRequestedDetails = True
			'send msg requesting the specific votes for this system
            Dim yMsg(14) As Byte
			Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eGetSenateObjectDetails).CopyTo(yMsg, lPos) : lPos += 2
            lPos += 4       'leave room for playerid
			yMsg(lPos) = eySenateRequestDetailsType.SystemVote : lPos += 1
			System.BitConverter.GetBytes(SystemID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(lProposalID).CopyTo(yMsg, lPos) : lPos += 4
			goUILib.SendMsgToPrimary(yMsg)
		End If
	End Sub

	Public Sub HandleDetailsMsg(ByRef yData() As Byte)
        Dim lPos As Int32 = 2           'for msgcode
        lPos += 13       'for type, systemid, and proposalid and playerid ph

		Dim lItemCnt As Int32 = yData(lPos) : lPos += 1
		yVoteRating = CByte(lItemCnt)

		lItemCnt = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim oNew(lItemCnt - 1) As SystemSenateVote
		Dim lVotesFor As Int32 = 0
		Dim lVotesTotal As Int32 = 0
		For X As Int32 = 0 To lItemCnt - 1
			Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim sPlayerName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
			Dim yVoteValue As Byte = yData(lPos) : lPos += 1
			Dim yPlayerRating As Byte = yData(lPos) : lPos += 1

			With oNew(X)
				.VoterID = lPlayerID
				.VoterName = sPlayerName
				.yRating = yPlayerRating
				.yVote = CType(yVoteValue, eyVoteValue)
				lVotesTotal += .yRating
				If .yVote = eyVoteValue.YesVote Then lVotesFor += .yRating
			End With
		Next X
		Dim lHalfVotes As Int32 = CInt(yVoteRating) \ 2
		If lVotesFor > lHalfVotes Then
			ySystemVote = eyVoteValue.YesVote
		ElseIf lVotesTotal - lVotesFor > lHalfVotes Then
			ySystemVote = eyVoteValue.NoVote
		Else
			ySystemVote = eyVoteValue.AbstainVote
		End If
		oVotes = oNew
	End Sub
End Class