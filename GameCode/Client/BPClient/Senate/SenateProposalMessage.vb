Option Strict On

Public Structure SenateProposalMessage
	Public sPosterName As String
	Public lPosterID As Int32
	Public sMsgData As String '= "Requesting..."
	Public lPostedOn As Int32			'date in GMT
	Public bRequestedDetails As Boolean
	Public Sub RequestDetails(ByVal lProposalID As Int32)
		If bRequestedDetails = False Then
			bRequestedDetails = True
            Dim yMsg(18) As Byte
			Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eGetSenateObjectDetails).CopyTo(yMsg, lPos) : lPos += 2
            lPos += 4       'leave room for playerid
			yMsg(lPos) = eySenateRequestDetailsType.ProposalMsg : lPos += 1
			System.BitConverter.GetBytes(lPosterID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(lPostedOn).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(lProposalID).CopyTo(yMsg, lPos) : lPos += 4
			goUILib.SendMsgToPrimary(yMsg)
		End If
	End Sub

	Public Function GetListBoxText() As String
		Return sPosterName.PadRight(25, " "c) & GetDateDisplayString(Date.SpecifyKind(GetDateFromNumber(lPostedOn), DateTimeKind.Utc).ToLocalTime, True, True)
	End Function
End Structure
