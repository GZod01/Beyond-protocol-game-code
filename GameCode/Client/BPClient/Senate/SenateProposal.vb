Option Strict On

Public Enum eyProposalState As Byte
	EmperorsChamber = 0
	AwaitingApproval = 1		'emperor's chamber voting is done, waiting approval from devs
	OnSenateFloor = 2			'dev approved on senate floor ready for votes
    DevDeclined = 4             'developers declined, returned to emperors chamber
    PassedFloor = 8
    FailedFloor = 16
    Archived = 32
    Implemented = 64
	SmallMsgBitShift = 128
End Enum

Public Enum eySenateRequestDetailsType As Byte
	SystemVote = 0
	SenateProposal = 1
    ProposalMsg = 2
    SenateObject = 3
    EmpChmbrMsgList = 4
    EmpChmbrMsg = 5
End Enum

Public Class SenateProposal
	Inherits Base_GUID

	Public yProposalState As eyProposalState
	Public sTitle As String = "Requesting..."
	Public sDescription As String = "Requesting..."

	Public sProposedByName As String = "Requesting..."
	Public lProposedBy As Int32
	Public lProposedOn As Int32			'date as number in gmt

	Public lVotesFor As Int32 = 0		'not saved, cummulated over time
	Public lVotesAgainst As Int32 = 0	'not saved, cummulated over time

	Public lMsgs As Int32 = 0

    Public yPlayerVote As eyVoteValue = eyVoteValue.AbstainVote
    Public yPlayerPriority As Byte = 0

	Public muMsgs() As SenateProposalMessage
	Public muVotes() As SenateVote
	Public moSystemVotes() As SystemVote		'not saved, kept track of as players change their votes and planet/system ownership changes

    Public yDefaultVote As Byte = eyVoteValue.NoVote
    Public lDeliveryEstimate As Int32 = 0
    Public lVotingEndDate As Int32 = 0
    Public lVotingStartDate As Int32 = 0
    Public lRequiredVoteCnt As Int32 = 0

    Public fAvgVotePriority As Single = 0.0F

	Private mbRequestedDetails As Boolean = False
	Public Sub RequestDetails()
		If mbRequestedDetails = False Then
			mbRequestedDetails = True
            Dim yMsg(10) As Byte
			Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eGetSenateObjectDetails).CopyTo(yMsg, lPos) : lPos += 2
            lPos += 4       'leave room for PlayerId
			yMsg(lPos) = eySenateRequestDetailsType.SenateProposal : lPos += 1
			System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
			goUILib.SendMsgToPrimary(yMsg)
		End If
	End Sub

	Public Sub FillFromSmallMsg(ByRef yData() As Byte)
        Dim lPos As Int32 = 2   'for msgcode
        lPos += 4       'for playerid ph
		ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		ObjTypeID = ObjectType.eSenateLaw : lPos += 2
		yProposalState = CType(yData(lPos), eyProposalState) : lPos += 1
		If (yProposalState And eyProposalState.SmallMsgBitShift) <> 0 Then
			yProposalState = yProposalState Xor eyProposalState.SmallMsgBitShift
		End If
		lVotesFor = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lVotesAgainst = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lProposedBy = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		lMsgs = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lTemp As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        sTitle = GetStringFromBytes(yData, lPos, lTemp) : lPos += lTemp

        Me.lVotingEndDate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim yPriority As Byte = yData(lPos) : lPos += 1
        fAvgVotePriority = (yPriority / 10.0F)

	End Sub

	Public Sub FillFromBigMsg(ByRef yData() As Byte)
        Dim lPos As Int32 = 2    'for msgcode
        lPos += 4    'for playerid ph

		Me.ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Me.ObjTypeID = ObjectType.eSenateLaw : lPos += 2

		yProposalState = CType(yData(lPos), eyProposalState) : lPos += 1
		If (yProposalState And eyProposalState.SmallMsgBitShift) <> 0 Then
			yProposalState = yProposalState Xor eyProposalState.SmallMsgBitShift
		End If
		Dim lTemp As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		sTitle = GetStringFromBytes(yData, lPos, lTemp) : lPos += lTemp
		lTemp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		sDescription = GetStringFromBytes(yData, lPos, lTemp) : lPos += lTemp

		sProposedByName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		lProposedBy = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lProposedOn = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim yTemp As Byte = yData(lPos) : lPos += 1

        If (yTemp And 8) <> 0 Then
            yTemp = CByte(yTemp Xor 8)
            yPlayerPriority = 1
        End If
        If (yTemp And 16) <> 0 Then
            yTemp = CByte(yTemp Xor 16)
            yPlayerPriority = 2
        End If
        If (yTemp And 32) <> 0 Then
            yTemp = CByte(yTemp Xor 32)
            yPlayerPriority = 3
        End If
        If (yTemp And 64) <> 0 Then
            yTemp = CByte(yTemp Xor 64)
            yPlayerPriority = 4
        End If
        If (yTemp And 128) <> 0 Then
            yTemp = CByte(yTemp Xor 128)
            yPlayerPriority = 5
        End If
        yPlayerVote = CType(yTemp, eyVoteValue)


        yDefaultVote = yData(lPos) : lPos += 1
        lDeliveryEstimate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lVotingStartDate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lVotingEndDate = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lRequiredVoteCnt = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim lVoteCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		If (Me.yProposalState And eyProposalState.OnSenateFloor) = 0 Then
			'emperors chambers
			ReDim muVotes(lVoteCnt - 1)
			For X As Int32 = 0 To lVoteCnt - 1
				muVotes(X) = New SenateVote()
				With muVotes(X)
					.VoterName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
					.yVote = CType(yData(lPos), eyVoteValue) : lPos += 1
				End With
			Next X
		Else
			'On Senate Floor
			ReDim moSystemVotes(lVoteCnt - 1)
			For X As Int32 = 0 To lVoteCnt - 1
				moSystemVotes(X) = New SystemVote()
				With moSystemVotes(X)
					.SystemID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.ySystemVote = CType(yData(lPos), eyVoteValue) : lPos += 1
					.yVoteRating = yData(lPos) : lPos += 1
				End With
			Next X
		End If

		lMsgs = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		ReDim muMsgs(lMsgs - 1)
		For X As Int32 = 0 To lMsgs - 1
			With muMsgs(X)
				.sPosterName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
				.lPosterID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.lPostedOn = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.sMsgData = "Requesting..."
				.bRequestedDetails = False
			End With
		Next X

	End Sub

	Public Function GetListBoxText() As String

        Dim sEndDate As String = GetDateFromNumber(lVotingEndDate).ToString("MM/dd")


        If (yProposalState And eyProposalState.Archived) <> 0 Then
            If (yProposalState And eyProposalState.PassedFloor) <> 0 Then
                If (yProposalState And eyProposalState.Implemented) <> 0 Then
                    Return Mid$(sTitle, 1, 50).PadRight(55, " "c) & "Implemented".PadRight(15, " "c) & (lMsgs.ToString & " Statements").PadRight(15, " "c) & sEndDate
                Else
                    Return Mid$(sTitle, 1, 50).PadRight(55, " "c) & "Passed".PadRight(15, " "c) & (lMsgs.ToString & " Statements").PadRight(15, " "c) & sEndDate
                End If
            Else
                Return Mid$(sTitle, 1, 50).PadRight(55, " "c) & "Did Not Pass".PadRight(15, " "c) & (lMsgs.ToString & " Statements").PadRight(15, " "c) & sEndDate
            End If
        ElseIf (yProposalState And eyProposalState.OnSenateFloor) <> 0 Then
            If (yProposalState And eyProposalState.PassedFloor) <> 0 Then
                Return Mid$(sTitle, 1, 50).PadRight(55, " "c) & "Passed".PadRight(15, " "c) & (lMsgs.ToString & " Statements").PadRight(15, " "c) & sEndDate
            ElseIf (yProposalState And eyProposalState.FailedFloor) <> 0 Then
                Return Mid$(sTitle, 1, 50).PadRight(55, " "c) & "Did Not Pass".PadRight(15, " "c) & (lMsgs.ToString & " Statements").PadRight(15, " "c) & sEndDate
            Else
                Return Mid$(sTitle, 1, 50).PadRight(55, " "c) & (lVotesFor.ToString & " / " & Math.Abs(lVotesAgainst).ToString).PadRight(15, " "c) & (lMsgs.ToString & " Statements").PadRight(15, " "c) & sEndDate
            End If
        Else
            sEndDate = fAvgVotePriority.ToString("0.0F")
            If (yProposalState And eyProposalState.DevDeclined) <> 0 Then
                Return Mid$(sTitle, 1, 50).PadRight(55, " "c) & "Dev Declined".PadRight(15, " "c) & (lMsgs.ToString & " Statements").PadRight(15, " "c) & sEndDate
            Else
                Return Mid$(sTitle, 1, 50).PadRight(55, " "c) & lVotesFor.ToString.PadRight(15, " "c) & (lMsgs.ToString & " Statements").PadRight(15, " "c) & sEndDate
            End If
        End If

    End Function

	Public Sub ReSortData()
		ReSortMsgs()
		'ReSortVotes()
	End Sub

	Private Sub ReSortMsgs()
		Dim lSorted() As Int32 = Nothing
		Dim lSortedUB As Int32 = -1
		If muMsgs Is Nothing Then Return
		Dim lItemUB As Int32 = muMsgs.GetUpperBound(0)

		For X As Int32 = 0 To lItemUB
			Dim lIdx As Int32 = -1

			Dim sNewValue As String = muMsgs(X).lPostedOn.ToString
			For Y As Int32 = 0 To lSortedUB
				Dim sLeftValue As String = muMsgs(lSorted(Y)).lPostedOn.ToString

				If sLeftValue.ToUpper < sNewValue.ToUpper Then
					lIdx = Y
					Exit For
				End If
			Next Y
			lSortedUB += 1
			ReDim Preserve lSorted(lSortedUB)
			If lIdx = -1 Then
				lSorted(lSortedUB) = X
			Else
				For Y As Int32 = lSortedUB To lIdx + 1 Step -1
					lSorted(Y) = lSorted(Y - 1)
				Next Y
				lSorted(lIdx) = X
			End If
		Next X

		If lSortedUB <> lMsgs - 1 Then Return

		If lSortedUB > -1 Then
			Dim uVals(lMsgs - 1) As SenateProposalMessage
			For X As Int32 = 0 To lSortedUB
				uVals(X) = muMsgs(lSorted(X))
			Next X
			muMsgs = uVals
		End If
	End Sub
	'Private Sub ReSortVotes()
	'	Dim lSorted() As Int32 = Nothing
	'	Dim lSortedUB As Int32 = -1
	'	If muVotes Is Nothing Then Return
	'	Dim lItemUB As Int32 = muVotes.GetUpperBound(0)

	'	For X As Int32 = 0 To lItemUB
	'		Dim lIdx As Int32 = -1

	'		Dim sNewValue As String = muVotes(X).VoterName
	'		For Y As Int32 = 0 To lSortedUB
	'			Dim sLeftValue As String = muVotes(lSorted(Y)).VoterName

	'			If sLeftValue.ToUpper > sNewValue.ToUpper Then
	'				lIdx = Y
	'				Exit For
	'			End If
	'		Next Y
	'		lSortedUB += 1
	'		ReDim Preserve lSorted(lSortedUB)
	'		If lIdx = -1 Then
	'			lSorted(lSortedUB) = X
	'		Else
	'			For Y As Int32 = lSortedUB To lIdx + 1 Step -1
	'				lSorted(Y) = lSorted(Y - 1)
	'			Next Y
	'			lSorted(lIdx) = X
	'		End If
	'	Next X

	'	If lSortedUB <> muVotes.GetUpperBound(0) Then Return

	'	If lSortedUB > -1 Then
	'		Dim uVals(lSortedUB) As SenateVote
	'		For X As Int32 = 0 To lSortedUB
	'			uVals(X) = muVotes(lSorted(X))
	'		Next X
	'		muVotes = uVals
	'	End If
	'End Sub

    Public Sub ClearRequestedDetails()
        mbRequestedDetails = False
    End Sub
End Class
