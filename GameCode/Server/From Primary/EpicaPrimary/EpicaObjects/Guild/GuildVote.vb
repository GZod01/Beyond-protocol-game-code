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

Public Enum eyVoteState As Byte
	eInProgress = 0
	eProcessed = 1
	eVoteFailed = 254
	eVotePassed = 255
End Enum

Public Class GuildVote
	Public VoteID As Int32 = -1
	Public ProposedByID As Int32 = -1
	Public dtVoteStarts As Date = Date.MinValue				'in GMT
	Public lVoteDuration As Int32 = 8		 'in hours
	Public yTypeOfVote As Byte
	Public lSelectedItem As Int32
	Public lNewValue As Int32
	Public yNewValueText() As Byte = Nothing
	Public ySummary() As Byte = Nothing
	Public yVoteState As Byte = 0

	Public lMemberVoteID() As Int32
	Public yMemberVoteValue() As eyVoteValue
	Public lMemberVoteUB As Int32 = -1

	Public Function GetPlayerVote(ByVal lPlayerID As Int32) As eyVoteValue
		For X As Int32 = 0 To lMemberVoteUB
			If lMemberVoteID(X) = lPlayerID Then Return yMemberVoteValue(X)
		Next X
		Return eyVoteValue.AbstainVote
	End Function

	Public Function SaveObject(ByVal lGuildID As Int32) As Boolean
		Dim bResult As Boolean = False
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

		Try
			Dim lVoteStarts As Int32 = 0
			If dtVoteStarts <> Date.MinValue Then lVoteStarts = GetDateAsNumber(dtVoteStarts)
			
			If VoteID = -1 Then
				'INSERT
				sSQL = "INSERT INTO tblGuildVote (GuildID, ProposedByID, VoteStarts, VoteDuration, TypeOfVote, SelectedItem, NewValueNumber, " & _
				  "NewValueText, VoteSummary, VoteState) VALUES (" & lGuildID & ", " & ProposedByID & ", " & lVoteStarts & ", " & lVoteDuration & ", " & _
				  CInt(yTypeOfVote) & ", " & lSelectedItem & ", " & lNewValue & ", '"
				If yNewValueText Is Nothing = False Then sSQL &= MakeDBStr(BytesToString(yNewValueText))
				sSQL &= "', '"
				If ySummary Is Nothing = False Then sSQL &= MakeDBStr(BytesToString(ySummary))
				sSQL &= "', " & yVoteState & ")"
			Else
				'UPDATE
				sSQL = "UPDATE tblGuildVote SET ProposedByID = " & ProposedByID & ", VoteStarts = " & lVoteStarts & ", VoteDuration = " & _
				 lVoteDuration & ", TypeOfVote = " & CInt(yTypeOfVote) & ", SelectedItem = " & lSelectedItem & ", NewValueNumber = " & _
				 lNewValue & ", NewValueText = '"
				If yNewValueText Is Nothing = False Then sSQL &= MakeDBStr(BytesToString(yNewValueText))
				sSQL &= "', VoteSummary = '"
				If ySummary Is Nothing = False Then sSQL &= MakeDBStr(BytesToString(ySummary))
				sSQL &= "', VoteState = " & yVoteState & " WHERE VoteID = " & Me.VoteID
			End If
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			If oComm.ExecuteNonQuery() = 0 Then
				Err.Raise(-1, "SaveObject", "No records affected when saving object!")
			End If
			If VoteID < 1 Then
				Dim oData As OleDb.OleDbDataReader
				oComm = Nothing
				sSQL = "SELECT MAX(VoteID) FROM tblGuildVote WHERE GuildID = " & lGuildID
				oComm = New OleDb.OleDbCommand(sSQL, goCN)
				oData = oComm.ExecuteReader(CommandBehavior.Default)
				If oData.Read Then
					VoteID = CInt(oData(0))
				End If
				oData.Close()
				oData = Nothing
			End If

			oComm = Nothing
			sSQL = "DELETE FROM tblGuildMemberVote WHERE VoteID = " & Me.VoteID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm = Nothing

			For X As Int32 = 0 To lMemberVoteUB
				Try
					If lMemberVoteID(X) <> -1 Then
						sSQL = "INSERT INTO tblGuildMemberVote (VoteID, MemberID, VoteValue) VALUES (" & Me.VoteID & ", " & lMemberVoteID(X) & ", " & CInt(yMemberVoteValue(X)) & ")"
						oComm = New OleDb.OleDbCommand(sSQL, goCN)
						oComm.ExecuteNonQuery()
						oComm = Nothing
					End If
				Catch ex As Exception
					LogEvent(LogEventType.CriticalError, "Unable to save GuildMemberVote. Reason: " & Err.Description)
				End Try
			Next X

			bResult = True
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object GuildVote. Reason: " & Err.Description)
		Finally
			oComm = Nothing
		End Try
		Return bResult
	End Function

	Public Sub SetMemberVote(ByVal lMemberID As Int32, ByVal yVoteValue As eyVoteValue)
		For X As Int32 = 0 To lMemberVoteUB
			If lMemberVoteID(X) = lMemberID Then
				yMemberVoteValue(X) = yVoteValue
				Return
			End If
		Next X
		If lMemberVoteID Is Nothing Then ReDim lMemberVoteID(-1)
		SyncLock lMemberVoteID
			ReDim Preserve lMemberVoteID(lMemberVoteUB + 1)
			ReDim Preserve yMemberVoteValue(lMemberVoteUB + 1)
			lMemberVoteID(lMemberVoteUB + 1) = lMemberID
			yMemberVoteValue(lMemberVoteUB + 1) = yVoteValue
			lMemberVoteUB += 1
		End SyncLock
	End Sub

	Public Function GetVoteObjectAsAdd() As Byte()
		Dim yMsg(302) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eProposeGuildVote).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(VoteID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(ProposedByID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(GetDateAsNumber(dtVoteStarts)).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lVoteDuration).CopyTo(yMsg, lPos) : lPos += 4
		yMsg(lPos) = yTypeOfVote : lPos += 1
		System.BitConverter.GetBytes(lSelectedItem).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lNewValue).CopyTo(yMsg, lPos) : lPos += 4
		If yNewValueText Is Nothing = False Then yNewValueText.CopyTo(yMsg, lPos)
		lPos += 20
		If ySummary Is Nothing = False Then ySummary.CopyTo(yMsg, lPos)
		lPos += 255
		yMsg(lPos) = yVoteState : lPos += 1

		Return yMsg
	End Function

	Public Sub ProcessVote(ByRef oGuild As Guild)
		Dim blVotesFor As Int64 = 0
		Dim blVotesAgainst As Int64 = 0

		For X As Int32 = 0 To lMemberVoteUB
			If lMemberVoteID(X) > 0 Then
				Dim oMember As GuildMember = oGuild.GetMember(lMemberVoteID(X))
				If oMember Is Nothing OrElse (oMember.yMemberState And GuildMemberState.Approved) = 0 Then Continue For

				Select Case oGuild.yVoteWeightType
					Case eyVoteWeightType.AgeOfPlayer
						Dim lJoinDate As Int32 = oMember.oMember.lJoinedGuildOn
						If lJoinDate <> 0 Then
							Dim dtJoin As Date = GetDateFromNumber(lJoinDate)
							Dim lDaysWithGuild As Int32 = CInt(Now.ToUniversalTime.Subtract(dtJoin).TotalDays)

							If yMemberVoteValue(X) = eyVoteValue.YesVote Then
								blVotesFor += lDaysWithGuild
							ElseIf yMemberVoteValue(X) = eyVoteValue.NoVote Then
								blVotesAgainst += lDaysWithGuild
							End If
						End If
					Case eyVoteWeightType.PopulationBased
						If yMemberVoteValue(X) = eyVoteValue.YesVote Then
							blVotesFor += oMember.oMember.blMaxPopulation
						ElseIf yMemberVoteValue(X) = eyVoteValue.NoVote Then
							blVotesAgainst += oMember.oMember.blMaxPopulation
						End If
					Case eyVoteWeightType.RankBased
						Dim oRank As GuildRank = oGuild.GetRank(oMember.oMember.lGuildRankID)
						If oRank Is Nothing = False Then
							If yMemberVoteValue(X) = eyVoteValue.YesVote Then
								blVotesFor += oRank.lVoteStrength
							ElseIf yMemberVoteValue(X) = eyVoteValue.NoVote Then
								blVotesAgainst += oRank.lVoteStrength
							End If
						End If
					Case Else
						If yMemberVoteValue(X) = eyVoteValue.YesVote Then
							blVotesFor += 1
						ElseIf yMemberVoteValue(X) = eyVoteValue.NoVote Then
							blVotesAgainst += 1
						End If
				End Select
			End If
		Next X

		If blVotesFor > blVotesAgainst Then
			yVoteState = eyVoteState.eVotePassed
			oGuild.SendMsgToGuildMembers(GetVoteObjectAsAdd)
		Else
			yVoteState = eyVoteState.eVoteFailed
			oGuild.SendMsgToGuildMembers(GetVoteObjectAsAdd)
			Return
		End If

		Select Case CType(yTypeOfVote, eyTypeOfVote)
			Case eyTypeOfVote.AddRankPermissions
				'selected item is a rankid
				Dim oRank As GuildRank = oGuild.GetRank(lSelectedItem)
				If oRank Is Nothing Then Return
				'newvalue is the rank permission to add
				oRank.lRankPermissions = oRank.lRankPermissions Or lNewValue
				oGuild.SendMsgToGuildMembers(oRank.GetGuildDetailsMsg)
			Case eyTypeOfVote.Charter
				'selected item is the law
				Select Case lSelectedItem
					Case -1
						'disband
					Case -2
						'tax rate interval
						'ok, change the guild tax rate interval
						oGuild.yGuildTaxRateInterval = CType(lNewValue, eyGuildInterval)
					Case -3
						'vote weight type
						oGuild.yVoteWeightType = CType(lNewValue, eyVoteWeightType)
					Case Else
						'guild flags
						If lNewValue = 0 Then
							'remove the flag
							If (oGuild.lBaseGuildFlags And lSelectedItem) <> 0 Then oGuild.lBaseGuildFlags = CType(oGuild.lBaseGuildFlags Xor lSelectedItem, elGuildFlags)
						Else
							'adding the guild flag
							oGuild.lBaseGuildFlags = CType(oGuild.lBaseGuildFlags Or lSelectedItem, elGuildFlags)
						End If
				End Select
				oGuild.SaveObject()
				oGuild.SendMsgToGuildMembers(goMsgSys.GetAddObjectMessage(oGuild, GlobalMessageCode.eAddObjectCommand))
			Case eyTypeOfVote.FreeFormVote
				'no action, return
				Return
			Case eyTypeOfVote.Membership
				'selected item is the action to take
				Select Case lSelectedItem
					Case 1		'accept member
						Dim oMember As GuildMember = oGuild.GetMember(lNewValue)
						'if target is applied then acting must be within the guild and must have rights to approve apps... then approve app
						If oMember Is Nothing = False Then
							If (oMember.yMemberState And GuildMemberState.Applied) <> 0 Then
								If oMember.oMember.oGuild Is Nothing = False Then
									Return
								End If

								oMember.yMemberState = oMember.yMemberState Or GuildMemberState.Approved
								oMember.oMember.lGuildID = oGuild.ObjectID
								oGuild.MemberJoined(lNewValue, False)
							End If
						End If 
					Case 2		'demote member
						Dim oTarget As Player = GetEpicaPlayer(lNewValue)
						If oTarget Is Nothing = False AndAlso oTarget.oGuild Is Nothing = False AndAlso oTarget.oGuild.ObjectID = oGuild.ObjectID Then
							Dim oCurrRank As GuildRank = oGuild.GetRank(oTarget.lGuildRankID)
							If oCurrRank Is Nothing Then Return

							Dim oNextRank As GuildRank = oGuild.GetNextRankPosition(1, oTarget.lGuildRankID)
							If oNextRank Is Nothing = False Then
								oTarget.lGuildRankID = oNextRank.lRankID
								Dim yData(13) As Byte
								Dim lPos As Int32 = 0
								System.BitConverter.GetBytes(GlobalMessageCode.eGuildMemberStatus).CopyTo(yData, lPos) : lPos += 2
								System.BitConverter.GetBytes(oGuild.ObjectID).CopyTo(yData, lPos) : lPos += 4
								System.BitConverter.GetBytes(oTarget.ObjectID).CopyTo(yData, lPos) : lPos += 4
								System.BitConverter.GetBytes(-1I).CopyTo(yData, lPos) : lPos += 4
								oGuild.SendMsgToGuildMembers(yData)
							End If
						End If
					Case 3		'promote member
						Dim oTarget As Player = GetEpicaPlayer(lNewValue)
						If oTarget Is Nothing = False AndAlso oTarget.oGuild Is Nothing = False AndAlso oTarget.oGuild.ObjectID = oGuild.ObjectID Then
							Dim oCurrRank As GuildRank = oGuild.GetRank(oTarget.lGuildRankID)
							If oCurrRank Is Nothing Then Return

							Dim oNextRank As GuildRank = oGuild.GetNextRankPosition(-1, oTarget.lGuildRankID)
							If oNextRank Is Nothing = False Then
								oTarget.lGuildRankID = oNextRank.lRankID
								Dim yData(13) As Byte
								Dim lPos As Int32 = 0
								System.BitConverter.GetBytes(GlobalMessageCode.eGuildMemberStatus).CopyTo(yData, lPos) : lPos += 2
								System.BitConverter.GetBytes(oGuild.ObjectID).CopyTo(yData, lPos) : lPos += 4
								System.BitConverter.GetBytes(oTarget.ObjectID).CopyTo(yData, lPos) : lPos += 4
								System.BitConverter.GetBytes(-2I).CopyTo(yData, lPos) : lPos += 4
								oGuild.SendMsgToGuildMembers(yData)
							End If
						End If
					Case 4		'remove member
						'ok, we have the rights... do it
						Dim oTarget As Player = GetEpicaPlayer(lNewValue)
						If oTarget Is Nothing = False AndAlso oTarget.oGuild Is Nothing = False AndAlso oTarget.oGuild.ObjectID = oGuild.ObjectID Then
							Dim oSB As New System.Text.StringBuilder()
							oSB.AppendLine("You have been removed from " & BytesToString(oGuild.yGuildName) & " by vote.")
                            Dim oPC As PlayerComm = oTarget.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Guild Removal", oTarget.ObjectID, GetDateAsNumber(Now), False, oTarget.sPlayerNameProper, Nothing)
							If oPC Is Nothing = False Then
								oTarget.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
							End If
							oPC = Nothing

							oGuild.RemoveMember(lNewValue)

							'Now, remove the member
							oTarget.lGuildID = -1
							oTarget.lGuildRankID = -1

							Dim yData(13) As Byte
							Dim lPos As Int32 = 0
							System.BitConverter.GetBytes(GlobalMessageCode.eGuildMemberStatus).CopyTo(yData, lPos) : lPos += 2
							System.BitConverter.GetBytes(oGuild.ObjectID).CopyTo(yData, lPos) : lPos += 4
							System.BitConverter.GetBytes(oTarget.ObjectID).CopyTo(yData, lPos) : lPos += 4
							System.BitConverter.GetBytes(Int32.MinValue).CopyTo(yData, lPos) : lPos += 4
							oGuild.SendMsgToGuildMembers(yData)

							oTarget.SendPlayerMessage(yData, False, 0)
						End If
						
					Case 5		'create rank
						Dim lNextPos As Int32 = -1
						For X As Int32 = 0 To oGuild.lRankUB
							If oGuild.oRanks(X) Is Nothing = False Then
								If oGuild.oRanks(X).yPosition >= lNextPos Then
									lNextPos = CInt(oGuild.oRanks(X).yPosition) + 1
								End If
							End If
						Next X
						If lNextPos < 0 Then lNextPos = 0
						If lNextPos > 255 Then lNextPos = 255
						Dim oRank As New GuildRank
						With oRank
							.lRankID = -1
							.lRankPermissions = 0
							.lVoteStrength = 0
							.ParentGuild = oGuild
							.TaxRateFlat = 0
							.TaxRatePercentage = 0
							.TaxRatePercType = eyGuildTaxPercType.CashFlow
							.yPosition = CByte(lNextPos)
							ReDim .yRankName(19)
							Array.Copy(yNewValueText, 0, .yRankName, 0, Math.Min(20, yNewValueText.Length))

							If .SaveObject(oGuild.ObjectID) = False Then Return
						End With
						oGuild.AddRank(oRank)
						oGuild.SendMsgToGuildMembers(goMsgSys.GetAddObjectMessage(oGuild, GlobalMessageCode.eAddObjectCommand))
					Case 6		'delete rank
						oGuild.RemoveRank(lNewValue)
					Case 7		'whether guild is recruiting
						If lNewValue = 0 Then
							If (oGuild.iRecruitFlags And eiRecruitmentFlags.GuildRecruiting) <> 0 Then
								oGuild.iRecruitFlags = oGuild.iRecruitFlags Xor eiRecruitmentFlags.GuildRecruiting
							End If
						Else
							oGuild.iRecruitFlags = oGuild.iRecruitFlags Or eiRecruitmentFlags.GuildRecruiting
						End If
				End Select
				'newvalue is the id of the member to take action on???
			Case eyTypeOfVote.RemoveRankPermission
				'selected item is a rankid
				'newvalue is the rank permission to remove
				Dim oRank As GuildRank = oGuild.GetRank(lSelectedItem)
				If oRank Is Nothing Then Return
				'newvalue is the rank permission to add
				If (oRank.lRankPermissions And lNewValue) <> 0 Then
					oRank.lRankPermissions = oRank.lRankPermissions Xor lNewValue
				End If
				oGuild.SendMsgToGuildMembers(oRank.GetGuildDetailsMsg)
			Case eyTypeOfVote.VotingWeights
				'selected item is a rankid
				Dim oRank As GuildRank = oGuild.GetRank(lSelectedItem)
				If oRank Is Nothing = False Then
					If yNewValueText Is Nothing = False Then
						oRank.lVoteStrength = Math.Max(CInt(Val(BytesToString(yNewValueText))), lNewValue)
					Else
						oRank.lVoteStrength = lNewValue
					End If
					oGuild.SendMsgToGuildMembers(oRank.GetGuildDetailsMsg)
				End If
				'value is a text value
		End Select
	End Sub
End Class

'Public Class VoteProposal
'    Inherits Epica_GUID

'    Public Enum eVoteTypes As Byte
'        eRuleChange = 0             'ID1 is the rule ID (eGuildRuleID), ID2 is the new value
'        eImpeach = 1                'ID1 is the playerid to impeach
'        eForeignPolicyChange = 2    'ID1 is the entityid, ID2 is the entitytypeid, ID3 is the new relationship
'        ePromoteDemoteMember = 3    'ID1 is the playerid, ID2 is the new rank id
'        eNominateDelegate = 4       'ID1 is the playerid
'        eInitiateTrade = 5          'ID1 is the trade ID
'        eElectDelegate = 6          'special, ID1, 2 and 3 are the parties up for election. A 4th can be 'penciled' in? <shrug>

'        eFreeform = 255
'    End Enum
'    Public Enum eGuildRuleID As Byte
'        RequirePeaceBetweenMembers = 1
'        ForeignPolicySharing = 2
'        ForeignPolicyVoting = 3
'        AutomaticTradeAgreements = 4
'        VotingWeight = 5
'        MemberOpenBorders = 6           'not sure of this one
'        ImpeachmentRule = 7
'        MemberTaxRate = 8
'        TaxPeriodInterval = 9
'        DelegateRevotePeriod = 10

'        MemberRankName = 11
'        MemberRankRights = 12
'        MemberRankVotingWeight = 13
'    End Enum
'    Public Enum ePublicVotingType As Byte
'        eNoDisplay = 0              'votes are concealed even after the vote ends
'        eUponVoteEnd = 1            'votes are concealed until after the vote ends and the result is completed
'        eAlways = 2                 'votes are not concealed
'    End Enum
'    Public Enum eVoteState As Byte          'the state in which this proposal is in
'        eInitialProposal = 0                'set when the proposal is initiated
'        eUnderDebate = 1                    'after the proposal is submitted but before the Vote Start date
'        eInProgress = 2                     'voting has started
'        ePollsClosed = 3                    'voting has ended but before execution
'        eProposalResultExecuted = 4         'voting has ended and the result has been executed
'        eProposalScrapped = 255             'the proposal is dismissed
'    End Enum

'    Public ParentEntity As Epica_GUID   'we use an Epica_GUID because of the possibility to expand this to Senates

'    Public VoteState As Byte

'    Public VoteType As Byte             'uses eVoteTypes
'    Public lID1 As Int32                'data id
'    Public lID2 As Int32                'data id
'    Public lID3 As Int32                'data id

'    Public yProposalText() As Byte      'a textual description of the proposal as entered by the initiator
'    Public dtProposedOn As DateTime     'date the proposal was proposed
'    Public dtVoteStart As DateTime      'date the voting begins. Votes cannot be submitted until this date. After this date, the contents of the proposal are locked
'    Public dtVoteEnd As DateTime        'date that voting ends. Votes cannot be submitted after this date. The proposal's result is executed.
'    Public lInitiatedByID As Int32      'player id that initiated the proposal

'    Public yPublicVoting As Byte        'indicates what type of vote this is and how individual player votes will be displayed (ePublicVotingType)

'    Private Structure VoteMsg
'        Public lPlayerID As Int32
'        Public yMessage() As Byte
'        Public dtPosted As DateTime
'    End Structure
'    Private muMsgs() As VoteMsg         'messages are like forum postings. Changes to the proposal in any way will create a message indicating the change
'    Private mlMsgUB As Int32 = -1

'    Private Structure PlayerVote
'        Public lPlayerID As Int32
'        Public lVoteValue As Int32      'the value, -1 is abstain
'        Public dtDateVoted As DateTime
'    End Structure
'    Private muVotes() As PlayerVote
'    Private mlVoteUB As Int32 = -1

'    Public Function GetObjectAsString() As Byte()
'        Dim yMsg() As Byte
'        Dim lPos As Int32 = 0
'        If yProposalText Is Nothing = False Then
'            ReDim yMsg(46 + yProposalText.Length)
'        Else : ReDim yMsg(46)
'        End If

'        GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
'        ParentEntity.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
'        yMsg(lPos) = VoteState : lPos += 1
'        yMsg(lPos) = VoteType : lPos += 1
'        yMsg(lPos) = yPublicVoting : lPos += 1
'        System.BitConverter.GetBytes(lID1).CopyTo(yMsg, lPos) : lPos += 4
'        System.BitConverter.GetBytes(lID2).CopyTo(yMsg, lPos) : lPos += 4
'        System.BitConverter.GetBytes(lID3).CopyTo(yMsg, lPos) : lPos += 4

'        Dim lTemp As Int32 = CInt(Val(dtProposedOn.ToString("yyMMddHHmm")))
'        System.BitConverter.GetBytes(lTemp).CopyTo(yMsg, lPos) : lPos += 4
'        lTemp = CInt(Val(dtVoteStart.ToString("yyMMddHHmm")))
'        System.BitConverter.GetBytes(lTemp).CopyTo(yMsg, lPos) : lPos += 4
'        lTemp = CInt(Val(dtVoteEnd.ToString("yyMMddHHmm")))
'        System.BitConverter.GetBytes(lTemp).CopyTo(yMsg, lPos) : lPos += 4
'        System.BitConverter.GetBytes(lInitiatedByID).CopyTo(yMsg, lPos) : lPos += 4

'        If yProposalText Is Nothing = False Then
'            System.BitConverter.GetBytes(yProposalText.Length).CopyTo(yMsg, lPos) : lPos += 4
'            yProposalText.CopyTo(yMsg, lPos) : lPos += yProposalText.Length
'        Else
'            System.BitConverter.GetBytes(0I).CopyTo(yMsg, lPos) : lPos += 4
'        End If

'        Return yMsg
'    End Function
'End Class