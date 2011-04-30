Option Strict On

#Region "  Enumerations  "
Public Enum eyGuildState As Byte
	Proposed = 0
	Forming = 1
	Formed = 2
	Disbanded = 255
End Enum

Public Enum eyGuildRequestDetailsType As Byte
	EventDetails = 0
	MemberDetails = 1
	GuildRel = 2
	GuildRank = 3
End Enum

Public Enum eiRecruitmentFlags As Int16
	GuildRecruiting = 1
	RecruitingTutorial = 2
	WillTrain = 4
	RecruitDiplomacy = 8
	RecruitEspionage = 16
	RecruitMilitary = 32
	RecruitProduction = 64
	RecruitTrade = 128
	RecruitResearch = 256
End Enum

Public Enum elGuildFlags As Int32
	RequirePeaceBetweenMembers = 1
	AutomaticTradeAgreements = 2

	ShareUnitVision = 4				'indicates that members can see other members units and facilities
    'UnifiedForeignPolicy = 8		'indicates that the members of the guild all share the same relationship lists

	'Indicates whether these items appear in the rank configuration screen
	DemoteMember_RA = 16
	AcceptMemberToGuild_RA = 32
	PromoteGuildMember_RA = 64
	RemoveGuildMember_RA = 128
	CreateRank_RA = 256
	DeleteRank_RA = 512
	ChangeVotingWeight_RA = 1024
End Enum

Public Enum eyVoteWeightType As Byte
	RankBased = 0		'each rank is assigned a vote weight value... all players within a rank are granted that rank's weight when voting
	PopulationBased = 1	'every player's vote weight is their current population size
	StaticValue = 2		'every player's vote is 1
	AgeOfPlayer = 3		'Player Vote Weight = the number of days the player has been with the guild rounded up
End Enum

Public Enum eyGuildInterval As Byte
	Daily = 0
	Weekly = 1
	SemiMonthly = 2	'twice a month
	Monthly = 3
	EveryTwoMonths = 4		'bi-monthly, every 2 months
	Quarterly = 5			'every 3 months
	SemiAnnually = 6		'every 6 months
	Annually = 7			'every 12 months
End Enum

'What the guild tax rate percentage is based on
Public Enum eyGuildTaxPercType As Byte
	CashFlow = 0
	TotalIncome = 1
	Treasury = 2
	'TODO: TradeIncome
End Enum
#End Region

Public Class Guild
	Inherits Base_GUID

	Public sName As String
	Public sMOTD As String
	Public sBillboard As String

	Public lIcon As Int32

	Public dtFormed As Date = Date.MinValue	'in GMT time
	Public dtJoined As Date = Date.MinValue	'in GMT time
	Public dtLastGuildRuleChange As Date = Date.MinValue		'in GMT Time

	'Public lCurrentRankID As Int32

	Public lGuildHallID As Int32 = -1
	Public lGuildHallLocID As Int32 = -1
	Public iGuildHallLocTypeID As Int16 = -1
	Public lGuildHallLocX As Int32 = 0
	Public lGuildHallLocZ As Int32 = 0

	Public lBaseGuildRules As elGuildFlags = CType(&HFFFFFFFF, elGuildFlags)

	Public yState As eyGuildState
	Public iRecruitFlags As eiRecruitmentFlags = 0

	Public yVoteWeightType As eyVoteWeightType = eyVoteWeightType.RankBased

	Public yGuildTaxRateInterval As eyGuildInterval = eyGuildInterval.Monthly
	Public yGuildTaxBaseMonth As Byte = 1
	Public yGuildTaxBaseDay As Byte = 1

	Public blTreasury As Int64 = 0

	Public moEvents(-1) As GuildEvent
	Public moRanks(-1) As GuildRank
	Public moVotes(-1) As GuildVote
	Public moMembers(-1) As GuildMember
    'Public moRels(-1) As GuildRel

	Public Sub HandleGuildEventSmall(ByRef yData() As Byte)
		Dim lPos As Int32 = 2		'for msgcode
		lPos += 4		'for guild id

		Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim oEvents(lCnt - 1) As GuildEvent
		For X As Int32 = 0 To lCnt - 1
			oEvents(X) = New GuildEvent()
			lPos = oEvents(X).FillFromSmallMsg(yData, lPos)
		Next X

		moEvents = oEvents
	End Sub

	Private mlCurrentRankID As Int32 = -1
	Private mlLastRankLookup As Int32 = -1
	Public ReadOnly Property lCurrentRankID() As Int32
		Get
			If mlCurrentRankID = -1 OrElse mlLastRankLookup = -1 OrElse glCurrentCycle - mlLastRankLookup > 30 Then
				mlLastRankLookup = glCurrentCycle
				If moMembers Is Nothing = False Then
					For X As Int32 = 0 To moMembers.GetUpperBound(0)
                        If moMembers(X) Is Nothing = False AndAlso moMembers(X).lMemberID = glPlayerID Then
                            mlCurrentRankID = moMembers(X).lRankID
                            Exit For
                        End If
					Next X
				End If
			End If
			Return mlCurrentRankID
		End Get
	End Property

	Public Function HasPermission(ByVal lPerm As RankPermissions) As Boolean
		Dim oRank As GuildRank = GetRankByID(lCurrentRankID)
		If oRank Is Nothing = False Then
			Return (oRank.lRankPermissions And lPerm) <> 0
		End If
		Return False
	End Function

	Public Function GetRankByID(ByVal lRankID As Int32) As GuildRank
		For X As Int32 = 0 To moRanks.GetUpperBound(0)
			If moRanks(X) Is Nothing = False AndAlso moRanks(X).lRankID = lRankID Then
				Return moRanks(X)
			End If
		Next X
		Return Nothing
	End Function

	Public Function GetRankName(ByVal lRankID As Int32) As String
		Dim oRank As GuildRank = GetRankByID(lRankID)
		If oRank Is Nothing = False Then
			Return oRank.sRankName
		End If
		Return "Requesting..."
	End Function

	Public Function GetDisplayedGuildHallLoc() As String
		If HasPermission(RankPermissions.ViewGuildBase) = False Then
			Return "CLASSIFIED"
		End If
		If lGuildHallLocID = -1 OrElse iGuildHallLocTypeID = -1 Then
			Return "Facility Not Built/Placed"
		End If
		Return GetCacheObjectValue(lGuildHallLocID, iGuildHallLocTypeID)
	End Function

	Public Function GetEvent(ByVal lEventID As Int32) As GuildEvent
		If moEvents Is Nothing = False Then
			For X As Int32 = 0 To moEvents.GetUpperBound(0)
				If moEvents(X) Is Nothing = False Then
					If moEvents(X).EventID = lEventID Then Return moEvents(X)
				End If
			Next X
		End If
		Return Nothing
	End Function

	Public Function GetVote(ByVal lVoteID As Int32) As GuildVote
		If moVotes Is Nothing = False Then
			For X As Int32 = 0 To moVotes.GetUpperBound(0)
				If moVotes(X) Is Nothing = False Then
					If moVotes(X).VoteID = lVoteID Then Return moVotes(X)
				End If
			Next X
		End If
		Return Nothing
	End Function

	Public Function GetMember(ByVal lMemberID As Int32) As GuildMember
		If moMembers Is Nothing = False Then
			For X As Int32 = 0 To moMembers.GetUpperBound(0)
				If moMembers(X) Is Nothing = False Then
					If moMembers(X).lMemberID = lMemberID Then Return moMembers(X)
				End If
			Next X
		End If
		Return Nothing
	End Function

    'Public Function GetRel(ByVal lID As Int32, ByVal iTypeID As Int16) As GuildRel
    '	If moRels Is Nothing = False Then
    '		For X As Int32 = 0 To moRels.GetUpperBound(0)
    '			If moRels(X) Is Nothing = False Then
    '				If moRels(X).lEntityID = lID AndAlso moRels(X).iEntityTypeID = iTypeID Then Return moRels(X)
    '			End If
    '		Next X
    '	End If
    '	Return Nothing
    'End Function

	'Public Shared Sub SetupTestGuild()

	'	If goCurrentPlayer Is Nothing Then
	'		goCurrentPlayer = New Player()
	'		goCurrentPlayer.ObjectID = 1
	'		goCurrentPlayer.ObjTypeID = ObjectType.ePlayer
	'		goCurrentPlayer.PlayerName = "Enoch Dagor"
	'	End If

	'	Dim oGuild As New Guild()
	'	With oGuild
	'		.dtFormed = Date.SpecifyKind(Now.Subtract(New TimeSpan(73, 12, 4)), DateTimeKind.Local).ToUniversalTime
	'		.dtJoined = Date.SpecifyKind(.dtFormed, DateTimeKind.Local).ToUniversalTime
	'		.iGuildHallLocTypeID = 3
	'		.lGuildHallLocID = 1
	'		.lGuildHallLocX = 15
	'		.lGuildHallLocZ = 15
	'		.lCurrentRankID = 1
	'		.lIcon = 1381444

	'		ReDim .moEvents(2)
	'		.moEvents(0) = New GuildEvent
	'		With .moEvents(0)
	'			.dtStartsAt = Date.SpecifyKind(Now.Subtract(New TimeSpan(20, 0, 0)).ToUniversalTime, DateTimeKind.Utc)
	'			.lDuration = 240
	'			.dtPostedOn = .dtStartsAt
	'			.EventID = 1
	'			.lPostedBy = 1
	'			.sDetails = "These are some test details for this guild event."
	'			.sTitle = "Test Event 0"
	'			.yEventIcon = 5
	'			.yEventType = 0
	'			.yMembersCanAccept = 1
	'			.yRecurrence = 0
	'			.ySendAlerts = 0
	'		End With
	'		.moEvents(1) = New GuildEvent
	'		With .moEvents(1)
	'			.dtStartsAt = Date.SpecifyKind(Now.Subtract(New TimeSpan(8, 0, 0)).ToUniversalTime, DateTimeKind.Utc)
	'			.lDuration = 5
	'			.dtPostedOn = .dtStartsAt
	'			.EventID = 2
	'			.lPostedBy = 1
	'			.sDetails = "Taxes are due"
	'			.sTitle = "Test Event 1"
	'			.yEventIcon = 1
	'			.yEventType = 0
	'			.yMembersCanAccept = 1
	'			.yRecurrence = 0
	'			.ySendAlerts = 0
	'		End With
	'		.moEvents(2) = New GuildEvent
	'		With .moEvents(2)
	'			.dtStartsAt = Date.SpecifyKind(Now.Add(New TimeSpan(2, 0, 0, 0)).ToUniversalTime, DateTimeKind.Utc)
	'			.lDuration = 120
	'			.dtPostedOn = .dtStartsAt
	'			.EventID = 3
	'			.lPostedBy = 1
	'			.sDetails = "Guild-wide meeting to discuss the impending war with the horsemen."
	'			.sTitle = "Test Event 2"
	'			.yEventIcon = 7
	'			.yEventType = 0
	'			.yMembersCanAccept = 1
	'			.yRecurrence = 0
	'			.ySendAlerts = 0
	'		End With

	'		ReDim .moMembers(1)
	'		.moMembers(0) = New GuildMember()
	'		With .moMembers(0)
	'			.dtLastOnline = Now
	'			.lMemberID = 1
	'			.sPlayerName = "Enoch Dagor"
	'			.lRankID = 1
	'			.yMemberState = GuildMemberState.Approved Or GuildMemberState.Applied
	'			.yPlayerTitle = 2
	'			.yPlayerGender = 1
	'		End With
	'		.moMembers(1) = New GuildMember()
	'		With .moMembers(1)
	'			.dtLastOnline = Now
	'			.lMemberID = 2
	'			.sPlayerName = "Csaj"
	'			.lRankID = 2
	'			.yMemberState = GuildMemberState.Approved Or GuildMemberState.Applied
	'			.yPlayerTitle = 2
	'			.yPlayerGender = 1
	'		End With

	'		ReDim .moRanks(2)
	'		.moRanks(0) = New GuildRank()
	'		With .moRanks(0)
	'			.lMembersInRank = 1
	'			.lRankID = 1
	'			.lRankPermissions = RankPermissions.AcceptApplicant Or RankPermissions.AcceptEvents Or RankPermissions.BuildGuildBase Or _
	'			  RankPermissions.ChangeMOTD Or RankPermissions.ChangeRankNames Or RankPermissions.ChangeRankPermissions Or _
	'			  RankPermissions.ChangeRankVotingWeight Or RankPermissions.ChangeRecruitment Or RankPermissions.CreateEvents Or _
	'			  RankPermissions.CreateRanks Or RankPermissions.DeleteEvents Or RankPermissions.DeleteRanks Or RankPermissions.DemoteMember Or _
	'			  RankPermissions.GuildChatChannel Or RankPermissions.InviteMember Or RankPermissions.PromoteMember Or _
	'			  RankPermissions.ProposeVotes Or RankPermissions.RejectMember Or RankPermissions.RemoveMember Or RankPermissions.ViewBankLog Or _
	'			  RankPermissions.ViewContentsHiSec Or RankPermissions.ViewContentsLowSec Or _
	'			  RankPermissions.ViewEventAttachments Or RankPermissions.ViewEvents Or RankPermissions.ViewGuildBase Or _
	'			  RankPermissions.ViewVotesHistory Or RankPermissions.ViewVotesInProgress Or RankPermissions.WithdrawHiSec Or _
	'			  RankPermissions.WithdrawLowSec
	'			.lVoteStrength = 5
	'			.sRankName = "Prime Minister"
	'			.yPosition = 0
	'		End With
	'		.moRanks(1) = New GuildRank()
	'		With .moRanks(1)
	'			.lMembersInRank = 1
	'			.lRankID = 2
	'			.lRankPermissions = CType(2147483647, RankPermissions)
	'			.lVoteStrength = 3
	'			.sRankName = "Council Member"
	'			.yPosition = 1
	'		End With
	'		.moRanks(2) = New GuildRank
	'		With .moRanks(2)
	'			.lMembersInRank = 0
	'			.lRankID = 3
	'			.lRankPermissions = CType(0, RankPermissions)
	'			.lVoteStrength = 0
	'			.sRankName = "Bottom Feeder"
	'			.yPosition = 2
	'		End With

	'		ReDim .moVotes(0)
	'		.moVotes(0) = New GuildVote
	'		With .moVotes(0)
	'			.dtVoteStarts = Date.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime
	'			.lVoteDuration = 24
	'			.lNewValue = 1
	'			.lSelectedItem = 1
	'			.ProposedByID = 1
	'			.sSummary = "Test Vote 0"
	'			.VoteID = 1
	'			.yPlayerVote = eyVoteValue.YesVote
	'			.yTypeOfVote = 0
	'		End With

	'		'.moVotes
	'		.ObjectID = 1
	'		.ObjTypeID = ObjectType.eGuild
	'		.sMOTD = "We are in desperate need of Power Engines. Contact Csaj regarding these." & vbCrLf & "Relations with the Horsemen have gotten worse... war is expected."
	'		.sName = "House Dagor"
	'		.yState = eyGuildState.Formed
	'	End With

	'	goCurrentPlayer.oGuild = oGuild

	'End Sub

	Public Function GetNextRankPosition(ByVal lPosDir As Int32, ByVal lStartRankID As Int32) As GuildRank
		If moRanks Is Nothing Then ReDim moRanks(-1)
		Dim lRankPos As Int32 = Int32.MinValue
		For X As Int32 = 0 To moRanks.GetUpperBound(0)
			If moRanks(X) Is Nothing = False AndAlso moRanks(X).lRankID = lStartRankID Then
				lRankPos = moRanks(X).yPosition
				Exit For
			End If
		Next X
		If lRankPos = Int32.MinValue Then Return Nothing

		'now, find next rank pos
		Dim oResult As GuildRank = Nothing
		Dim lCurrBest As Int32 = 0
		lCurrBest = Int32.MaxValue

		For X As Int32 = 0 To moRanks.GetUpperBound(0)
			If moRanks(X) Is Nothing = False AndAlso moRanks(X).lRankID <> lStartRankID Then
				'k, now check our direction
				Dim lTmp As Int32 = CInt(moRanks(X).yPosition)
				If lPosDir < 0 Then
					'ok, looking for next lowest rank position
					If lTmp < lRankPos Then
						If lRankPos - lTmp < lCurrBest Then
							oResult = moRanks(X)
							lCurrBest = lRankPos - lTmp
						End If
					End If
				Else
					'ok, looking for next highest rank position
					If lTmp > lRankPos Then
						If lTmp - lRankPos < lCurrBest Then
							oResult = moRanks(X)
							lCurrBest = lTmp - lRankPos
						End If
					End If
				End If
			End If
		Next X

		Return oResult
	End Function

	Public Function GetLowestRank() As GuildRank
		If Me.moRanks Is Nothing = False Then
			Dim yLowest As Byte = 0
			Dim oRank As GuildRank = Nothing
			For X As Int32 = 0 To moRanks.GetUpperBound(0)
				If moRanks(X) Is Nothing = False AndAlso moRanks(X).yPosition >= yLowest Then
					yLowest = moRanks(X).yPosition
					oRank = moRanks(X)
				End If
			Next X
			Return oRank
		End If
		Return Nothing
	End Function

	Public Function GetCharterText() As String
		'Ok, let's get the base line rules
		Dim oSB As New System.Text.StringBuilder

		oSB.AppendLine("Last Charter Change: " & dtLastGuildRuleChange.ToLocalTime.ToString)

		If yState = eyGuildState.Formed Then
			'oSB.AppendLine("In accordance to the laws set forth by the Galactic Senate as endowed by the populace of the entire Realm of Human Civilization, it is to be that " & Me.sName & _
			'  " be recognized as a legally formed body and collection of governments.")
			oSB.AppendLine("Formed On: " & dtFormed.ToLocalTime.ToString)

			oSB.AppendLine(vbCrLf & "It is declared and set forth by the Galactic Senate, as endowed by the populace of the entire Realm of Human Civilization, that " & _
			 Me.sName & " be recognized as a government body of governments and be permitted all the rules bestowed to such governments. The Galactic Senate " & _
			 "upholds the following laws and regulations within said government body of governments contained and described in this document.")
		Else
			oSB.AppendLine("In Process Of Senate Recognition")
		End If

		oSB.AppendLine()

		If (lBaseGuildRules And elGuildFlags.ShareUnitVision) <> 0 Then
			oSB.AppendLine(Me.sName & " shares military intelligence with all members automatically." & vbCrLf)
		End If
        'If (lBaseGuildRules And elGuildFlags.UnifiedForeignPolicy) <> 0 Then
        '	oSB.AppendLine("Members of " & Me.sName & " are required to share the relationship standings and foreign policy of " & Me.sName & "." & vbCrLf)
        '	oSB.AppendLine("Relationships of new members will be assimilated into the relationship standings of " & Me.sName & "." & vbCrLf)
        'End If
		If (lBaseGuildRules And elGuildFlags.AutomaticTradeAgreements) <> 0 Then
			oSB.AppendLine("Members of " & Me.sName & " are afforded open trade agreements with other members." & vbCrLf)
		End If
		If (lBaseGuildRules And elGuildFlags.RequirePeaceBetweenMembers) <> 0 Then
			oSB.AppendLine("All members of " & Me.sName & " must not have hostile relations with fellow members. Upon acceptance, new members that have hostile relations with existing members will immediately cease hostilities." & vbCrLf)
		Else
			oSB.AppendLine("It is understood that no guarantee of mutual peace or cease of hostilities between members of " & Me.sName & " has been made." & vbCrLf)
		End If

		If (lBaseGuildRules And elGuildFlags.AcceptMemberToGuild_RA) <> 0 Then
			oSB.AppendLine("The ability to accept members is assignable to individual ranks." & vbCrLf)
		Else : oSB.AppendLine("The ability to accept members requires a vote." & vbCrLf)
		End If
		If (lBaseGuildRules And elGuildFlags.ChangeVotingWeight_RA) <> 0 Then
			oSB.AppendLine("The ability to change rank voting weights is assignable to individual ranks." & vbCrLf)
		Else : oSB.AppendLine("The ability to change rank voting weights requires a vote." & vbCrLf)
		End If
		If (lBaseGuildRules And elGuildFlags.CreateRank_RA) <> 0 Then
			oSB.AppendLine("The ability to create ranks is assignable to individual ranks." & vbCrLf)
		Else : oSB.AppendLine("The ability to create ranks requires a vote." & vbCrLf)
		End If
		If (lBaseGuildRules And elGuildFlags.DeleteRank_RA) <> 0 Then
			oSB.AppendLine("The ability to delete ranks is assignable to individual ranks." & vbCrLf)
		Else : oSB.AppendLine("The ability to delete ranks requires a vote." & vbCrLf)
		End If
		If (lBaseGuildRules And elGuildFlags.DemoteMember_RA) <> 0 Then
			oSB.AppendLine("The ability to demote a member's ranks is assignable to individual ranks." & vbCrLf)
		Else : oSB.AppendLine("To demote a member in rank requires a vote." & vbCrLf)
		End If
		If (lBaseGuildRules And elGuildFlags.PromoteGuildMember_RA) <> 0 Then
			oSB.AppendLine("Promotion of a member rank is assignable to individual ranks." & vbCrLf)
		Else : oSB.AppendLine("Promotion of a member rank requires a vote." & vbCrLf)
		End If
		If (lBaseGuildRules And elGuildFlags.RemoveGuildMember_RA) <> 0 Then
			oSB.AppendLine("Forceful removal of a member from " & Me.sName & " is assignable to individual ranks." & vbCrLf)
		Else : oSB.AppendLine("Forceful removal of a member from " & Me.sName & " requires a vote." & vbCrLf)
		End If

		Select Case Me.yVoteWeightType
			Case eyVoteWeightType.RankBased
				oSB.AppendLine("Voting Weight per member is based on Rank of the member.")
			Case eyVoteWeightType.PopulationBased
				oSB.AppendLine("Voting Weight per member is based on the total population of the member's domain.")
			Case eyVoteWeightType.StaticValue
				oSB.AppendLine("Voting Weight is static per member meaning each member vote counts as 1 vote.")
			Case eyVoteWeightType.AgeOfPlayer
				oSB.AppendLine("Voting Weight per member is based on the time that the member has been part of " & Me.sName)
		End Select
		oSB.AppendLine()

		Dim sText As String = ""
		Select Case Me.yGuildTaxRateInterval
			Case eyGuildInterval.Annually
				sText = "annually on every " & yGuildTaxBaseDay & " day of " & GetMonthName(yGuildTaxBaseMonth)
			Case eyGuildInterval.Daily
				sText = "daily"
			Case eyGuildInterval.EveryTwoMonths
				sText = "every two months on the " & yGuildTaxBaseDay & " day"
			Case eyGuildInterval.Monthly
				sText = "every " & yGuildTaxBaseDay & " day of every month"
			Case eyGuildInterval.Quarterly
				sText = "every three months on the " & yGuildTaxBaseDay & " day"
			Case eyGuildInterval.SemiAnnually
				sText = "every six months on the " & yGuildTaxBaseDay & " day"
			Case eyGuildInterval.SemiMonthly
				Dim lFirst As Int32 = yGuildTaxBaseDay
				Dim lSecond As Int32 = yGuildTaxBaseDay
				If lFirst > 15 Then
					lFirst -= 15
				Else : lSecond += 15
				End If

				sText = "on the " & lFirst & " day and " & lSecond & " day of every month"
			Case eyGuildInterval.Weekly
				sText = "every " & GetDayOfWeekName(yGuildTaxBaseDay) & " of every week"
		End Select
		oSB.AppendLine(Me.sName & " will collect payment in the form of taxes from members " & sText & vbCrLf)

		If Me.moRanks Is Nothing = False Then
			oSB.AppendLine("It is decreed that the following ranks be identified: ")
			For X As Int32 = 0 To moRanks.GetUpperBound(0)
				If moRanks(X) Is Nothing = False Then
					With moRanks(X)
						oSB.AppendLine("Rank: " & .sRankName)
						If yVoteWeightType = eyVoteWeightType.RankBased Then
							oSB.AppendLine("Vote Weight: " & .lVoteStrength)
						End If

						If .TaxRateFlat <> 0 Then
							oSB.AppendLine("Tax Flat Rate: " & .TaxRateFlat.ToString("#,##0"))
						End If
						If .TaxRatePercentage <> 0 Then
							Dim sType As String = ""
							Select Case .TaxRatePercType
								Case eyGuildTaxPercType.CashFlow
									sType = "Cashflow"
								Case eyGuildTaxPercType.TotalIncome
									sType = "Total Income"
								Case eyGuildTaxPercType.Treasury
									sType = "Total Treasury"
							End Select
							oSB.AppendLine("Taxed " & .TaxRatePercentage & "% of the member's " & sType)
						End If
					End With
					oSB.AppendLine()
				End If
			Next X
		End If

		Return oSB.ToString
	End Function

	Public Shared Sub FillComboWithGuildIntervalList(ByRef cbo As UIComboBox)
		cbo.Clear()

		cbo.AddItem("Daily")
		cbo.ItemData(cbo.NewIndex) = eyGuildInterval.Daily
		cbo.AddItem("Weekly")
		cbo.ItemData(cbo.NewIndex) = eyGuildInterval.Weekly
		cbo.AddItem("Semi-Monthly")
		cbo.ItemData(cbo.NewIndex) = eyGuildInterval.SemiMonthly
		cbo.AddItem("Monthly")
		cbo.ItemData(cbo.NewIndex) = eyGuildInterval.Monthly
		cbo.AddItem("Bi-Monthly")
		cbo.ItemData(cbo.NewIndex) = eyGuildInterval.EveryTwoMonths
		cbo.AddItem("Quarterly")
		cbo.ItemData(cbo.NewIndex) = eyGuildInterval.Quarterly
		cbo.AddItem("Semi-Annually")
		cbo.ItemData(cbo.NewIndex) = eyGuildInterval.SemiAnnually
		cbo.AddItem("Annually")
		cbo.ItemData(cbo.NewIndex) = eyGuildInterval.Annually
	End Sub

	Public Function MemberInGuild(ByVal lPlayerID As Int32) As Boolean
		If moMembers Is Nothing = False Then
			Dim lCurUB As Int32 = moMembers.GetUpperBound(0)
			For X As Int32 = 0 To lCurUB
				If moMembers(X) Is Nothing = False Then
					If moMembers(X).lMemberID = lPlayerID Then
                        Return (moMembers(X).yMemberState And GuildMemberState.Approved) <> 0
					End If
				End If
			Next X
		End If
		Return False
	End Function

	Public Sub FillFromMsg(ByRef yData() As Byte)
		Dim lPos As Int32 = 2		'for msgcodee

		Me.ObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Me.ObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

		sName = GetStringFromBytes(yData, lPos, 50) : lPos += 50

		Dim lVal As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		If lVal <> 0 Then
			dtFormed = Date.SpecifyKind(GetDateFromNumber(lVal), DateTimeKind.Utc)
		Else : dtFormed = Date.MinValue
		End If
		lVal = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		If lVal <> 0 Then
			dtLastGuildRuleChange = Date.SpecifyKind(GetDateFromNumber(lVal), DateTimeKind.Utc)
		Else : dtLastGuildRuleChange = Date.MinValue
		End If

		lIcon = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		lBaseGuildRules = CType(System.BitConverter.ToInt32(yData, lPos), elGuildFlags) : lPos += 4
		yState = CType(yData(lPos), eyGuildState) : lPos += 1
		iRecruitFlags = CType(System.BitConverter.ToInt16(yData, lPos), eiRecruitmentFlags) : lPos += 2

		yVoteWeightType = CType(yData(lPos), eyVoteWeightType) : lPos += 1
		yGuildTaxRateInterval = CType(yData(lPos), eyGuildInterval) : lPos += 1
		yGuildTaxBaseMonth = yData(lPos) : lPos += 1
		yGuildTaxBaseDay = yData(lPos) : lPos += 1

		blTreasury = System.BitConverter.ToInt64(yData, lPos) : lPos += 8

		lGuildHallID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		If lGuildHallID > 0 Then
			lGuildHallLocID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			iGuildHallLocTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			lGuildHallLocX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			lGuildHallLocZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Else
			lGuildHallLocID = -1
			iGuildHallLocTypeID = -1S
			lGuildHallLocX = 0
			lGuildHallLocZ = 0
		End If


		Dim lTextLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		sMOTD = GetStringFromBytes(yData, lPos, lTextLen) : lPos += lTextLen
		lTextLen = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		sBillboard = GetStringFromBytes(yData, lPos, lTextLen) : lPos += lTextLen

		Dim lMemberCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim oMembers(lMemberCnt - 1) As GuildMember
		For X As Int32 = 0 To lMemberCnt - 1
			oMembers(X) = New GuildMember()
			Dim lMemberID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			If lMemberID > 0 Then
				With oMembers(X)
					.lMemberID = lMemberID
					.yMemberState = CType(yData(lPos), GuildMemberState) : lPos += 1
					.lRankID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.yPlayerTitle = yData(lPos) : lPos += 1
					.yPlayerGender = yData(lPos) : lPos += 1
					lVal = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					If lVal <> 0 Then
						.dtJoined = Date.SpecifyKind(GetDateFromNumber(lVal), DateTimeKind.Utc)
					Else : .dtJoined = Date.MinValue
					End If
					lVal = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					If lVal <> 0 Then
						.dtLastOnline = Date.SpecifyKind(GetDateFromNumber(lVal), DateTimeKind.Utc)
					Else : .dtLastOnline = Date.MinValue
					End If

					If lMemberID = glPlayerID Then
						'Me.lCurrentRankID = .lRankID
						Me.dtJoined = .dtJoined
					End If

				End With
			Else
				lPos += 15
				oMembers(X) = Nothing
			End If
		Next X
		moMembers = oMembers

        'Dim lRelCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        'Dim oRels(lRelCnt - 1) As GuildRel
        'For X As Int32 = 0 To lRelCnt - 1
        '	Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        '	Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        '	Dim yRelToUs As Byte = yData(lPos) : lPos += 1
        '	Dim yRelToThem As Byte = yData(lPos) : lPos += 1
        '	Dim lIcon As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        '	If lEntityID > 0 Then
        '		oRels(X) = New GuildRel
        '		With oRels(X)
        '			.lEntityID = lEntityID
        '			.iEntityTypeID = iEntityTypeID
        '			.yRelTowardsUs = yRelToUs
        '			.yRelTowardsThem = yRelToThem
        '			.lIcon = lIcon
        '		End With
        '	End If

        'Next X
        'moRels = oRels

		Dim lRankCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim oRanks(lRankCnt - 1) As GuildRank
		For X As Int32 = 0 To lRankCnt - 1
			oRanks(X) = New GuildRank()
			With oRanks(X)
				.lRankID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.sRankName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
				.lRankPermissions = CType(System.BitConverter.ToInt32(yData, lPos), RankPermissions) : lPos += 4
				.lVoteStrength = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.yPosition = yData(lPos) : lPos += 1
				.TaxRatePercentage = yData(lPos) : lPos += 1
				.TaxRatePercType = CType(yData(lPos), eyGuildTaxPercType) : lPos += 1
				.TaxRateFlat = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.lMembersInRank = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			End With
		Next X
		moRanks = oRanks


        RecheckGuildMembers()

    End Sub

    Public Sub RecheckGuildMembers()
        'Now, need to go through the currentenvir and set all units bGuildMember setting
        Try
            Dim oEnvir As BaseEnvironment = goCurrentEnvir
            If oEnvir Is Nothing = False Then
                Dim lUB As Int32 = -1
                If oEnvir.lEntityIdx Is Nothing = False Then lUB = Math.Min(oEnvir.lEntityUB, oEnvir.lEntityIdx.GetUpperBound(0))
                For X As Int32 = 0 To lUB
                    If oEnvir.lEntityIdx(X) > -1 Then
                        Dim oEntity As BaseEntity = oEnvir.oEntity(X)
                        If oEntity Is Nothing = False Then
                            oEntity.bGuildMember = MemberInGuild(oEntity.OwnerID)
                        End If
                    End If
                Next X
            End If
        Catch
        End Try
    End Sub

	Public Sub MoveRank(ByVal lRankID As Int32, ByVal yMoveDir As Byte)
		'0 for nothing, 1 for up, 2 for down, 255 for remove
		If yMoveDir = 1 Then
			'move up... find next lowest rank number
			Dim oRank As GuildRank = GetRankByID(lRankID)
			If oRank Is Nothing Then Return
			Dim oNextRank As GuildRank = GetNextRankPosition(-1, lRankID)
			If oNextRank Is Nothing Then Return
			Dim lTmp As Int32 = oNextRank.yPosition
			oNextRank.yPosition = oRank.yPosition
			oRank.yPosition = CByte(lTmp)
		Else
			'move down
			Dim oRank As GuildRank = GetRankByID(lRankID)
			If oRank Is Nothing Then Return
			Dim oNextRank As GuildRank = GetNextRankPosition(1, lRankID)
			If oNextRank Is Nothing Then Return
			Dim lTmp As Int32 = oNextRank.yPosition
			oNextRank.yPosition = oRank.yPosition
			oRank.yPosition = CByte(lTmp)
		End If
	End Sub

	Public Sub RemoveRank(ByVal lRankID As Int32)
		Dim oNextRank As GuildRank = GetNextRankPosition(1, lRankID)
		If oNextRank Is Nothing Then oNextRank = GetNextRankPosition(-1, lRankID)
		If oNextRank Is Nothing Then Return

		If moMembers Is Nothing = False Then
			For X As Int32 = 0 To moMembers.GetUpperBound(0)
				If moMembers(X) Is Nothing = False Then
					If moMembers(X).lRankID = lRankID Then
						moMembers(X).lRankID = oNextRank.lRankID
						oNextRank.lMembersInRank += 1
					End If
				End If
			Next X
		End If

		'If lCurrentRankID = lRankID Then lCurrentRankID = oNextRank.lRankID
		If moRanks Is Nothing = False Then
			For X As Int32 = 0 To moRanks.GetUpperBound(0)
				If moRanks(X) Is Nothing = False AndAlso moRanks(X).lRankID = lRankID Then
					moRanks(X) = Nothing
				End If
			Next X
		End If
	End Sub

	Public Sub RemoveMember(ByVal lMemberID As Int32)
		If moMembers Is Nothing Then Return
		For X As Int32 = 0 To moMembers.GetUpperBound(0)
			If moMembers(X) Is Nothing = False AndAlso moMembers(X).lMemberID = lMemberID Then
				moMembers(X) = Nothing
			End If
		Next X
	End Sub

	Public Sub AddVote(ByRef oVote As GuildVote)
		If moVotes Is Nothing Then ReDim moVotes(-1)
		For X As Int32 = 0 To moVotes.GetUpperBound(0)
			If moVotes(X) Is Nothing Then
				moVotes(X) = oVote
				Return
			End If
		Next X
		ReDim Preserve moVotes(moVotes.GetUpperBound(0) + 1)
		moVotes(moVotes.GetUpperBound(0)) = oVote
	End Sub

	Public Sub AddEvent(ByRef oEvent As GuildEvent)
		If moEvents Is Nothing Then ReDim moEvents(-1)
		For X As Int32 = 0 To moEvents.GetUpperBound(0)
			If moEvents(X) Is Nothing Then
				moEvents(X) = oEvent
				Return
			End If
		Next X
		ReDim Preserve moEvents(moEvents.GetUpperBound(0) + 1)
		moEvents(moEvents.GetUpperBound(0)) = oEvent
	End Sub

	Public Sub AddMember(ByRef oMember As GuildMember)
		If moMembers Is Nothing Then ReDim moMembers(-1)
		Dim lIdx As Int32 = -1
		For X As Int32 = 0 To moMembers.GetUpperBound(0)
			If moMembers(X) Is Nothing = False Then
				If moMembers(X).lMemberID = oMember.lMemberID Then
					moMembers(X) = oMember
					Return
				End If
			ElseIf lIdx <> -1 Then
				lIdx = X
			End If
		Next X
		If lIdx = -1 Then
			ReDim Preserve moMembers(moMembers.GetUpperBound(0) + 1)
			moMembers(moMembers.GetUpperBound(0)) = oMember
		Else
			moMembers(lIdx) = oMember
		End If
	End Sub

	Public Sub RemoveEvent(ByVal lEventID As Int32)
		If moEvents Is Nothing Then Return
		For X As Int32 = 0 To moEvents.GetUpperBound(0)
			If moEvents(X) Is Nothing = False Then
				If moEvents(X).EventID = lEventID Then
					moEvents(X) = Nothing
				End If
			End If
		Next X
	End Sub

	Public Sub HandleGuildRequestDetails(ByRef yData() As Byte)
		Dim lPos As Int32 = 2
		Dim yType As Byte = yData(lPos) : lPos += 1
		Select Case CType(yType, eyGuildRequestDetailsType)
			Case eyGuildRequestDetailsType.EventDetails
				Dim lEventID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				Dim oEvent As GuildEvent = GetEvent(lEventID)
				If oEvent Is Nothing Then Return
				With oEvent
					.sTitle = GetStringFromBytes(yData, lPos, 50) : lPos += 50
					.lPostedBy = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.dtPostedOn = Date.SpecifyKind(GetDateFromNumber(System.BitConverter.ToInt32(yData, lPos)), DateTimeKind.Utc) : lPos += 4
					.dtStartsAt = Date.SpecifyKind(GetDateFromNumber(System.BitConverter.ToInt32(yData, lPos)), DateTimeKind.Utc) : lPos += 4
					.lDuration = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.yEventIcon = yData(lPos) : lPos += 1
					Dim yVal As Byte = yData(lPos) : lPos += 1
					.yEventType = CByte(yVal And 15)
					If (yVal And 16) <> 0 Then .yMembersCanAccept = 1 Else .yMembersCanAccept = 0
					.ySendAlerts = CByte((yVal And 224) \ 32)
					.yRecurrence = yData(lPos) : lPos += 1

					Dim lLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.sDetails = GetStringFromBytes(yData, lPos, lLen) : lPos += lLen
				End With
			Case eyGuildRequestDetailsType.GuildRel
                'Dim lEntityID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                'Dim iEntityTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                'Dim oRel As GuildRel = GetRel(lEntityID, iEntityTypeID)
                'If oRel Is Nothing Then Return
                'With oRel
                '	.yRelTowardsUs = yData(lPos) : lPos += 1
                '	.yRelTowardsThem = yData(lPos) : lPos += 1
                '	.lIcon = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                '	.lWhoMadeFirstContact = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                '	.lWhoFirstContactWasMadeWith = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                '	Dim lTmp As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                '	If lTmp = 0 Then .dtWhenFirstContactMade = Date.MinValue Else .dtWhenFirstContactMade = Date.SpecifyKind(GetDateFromNumber(lTmp), DateTimeKind.Utc)
                '	.lLocationIDOfFC = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                '	.iLocationTypeIDOfFC = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                '	.lLocXOfFC = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                '	.lLocZOfFC = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

                '	lTmp = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                '	.sNotes = GetStringFromBytes(yData, lPos, lTmp) : lPos += lTmp
                'End With
			Case eyGuildRequestDetailsType.MemberDetails
				Dim lMemberID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				Dim oMember As GuildMember = GetMember(lMemberID)
				If oMember Is Nothing Then Return

				With oMember
					.yMemberState = CType(yData(lPos), GuildMemberState) : lPos += 1
					If (.yMemberState And GuildMemberState.Applied) <> 0 Then
						Dim lTechScore As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lDipScore As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lMilScore As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lPopScore As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lProdScore As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lWealthScore As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lTotalScore As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lColonyCount As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lStartedEnvirID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim iStartedEnvirTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
						Dim lTutPhaseWaves As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lPlayedTimeInTutOne As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim lTimeInWaves As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						Dim blMaxPop As Int64 = System.BitConverter.ToInt64(yData, lPos) : lPos += 8

						.sDetails = "This player has applied for guild membership and has provided the following portfolio:" & vbCrLf & _
						 "Player Scores" & vbCrLf & "Technology: " & lTechScore.ToString("#,##0") & vbCrLf & _
						 "Diplomacy: " & lDipScore.ToString("#,##0") & vbCrLf & "Population: " & lPopScore.ToString("#,##0") & vbCrLf & _
						 "Production: " & lProdScore.ToString("#,##0") & vbCrLf & "Wealth: " & lWealthScore.ToString("#,##0") & vbCrLf & _
						 "TOTAL: " & lTotalScore.ToString("#,##0") & vbCrLf & vbCrLf & "Colonies: " & lColonyCount.ToString("#,##0") & _
						 vbCrLf & "Homeworld: " & GetCacheObjectValue(lStartedEnvirID, iStartedEnvirTypeID) & vbCrLf & _
						 "Time spent in Tutorial Phase One: " & GetDurationFromSeconds(lPlayedTimeInTutOne, True) & vbCrLf & _
						 "Time spent in Waves: " & GetDurationFromSeconds(lTimeInWaves, True) & vbCrLf & _
						 "Number of Waves: " & lTutPhaseWaves.ToString("#,##0") & vbCrLf & vbCrLf & "Maximum Population: " & blMaxPop.ToString("#,##0")
					Else
						Dim lTotalScore As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
						.sDetails = "Because we invited this player, we can only see the total score which is : " & lTotalScore.ToString("#,##0")
					End If
				End With
			Case eyGuildRequestDetailsType.GuildRank
				Dim lRankID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				Dim oRank As GuildRank = GetRankByID(lRankID)
				Dim bAdded As Boolean = False
				If oRank Is Nothing Then
					bAdded = True
					oRank = New GuildRank()
				End If
				With oRank
					.lRankID = lRankID
					.sRankName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
					.lRankPermissions = CType(System.BitConverter.ToInt32(yData, lPos), RankPermissions) : lPos += 4
					.lVoteStrength = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.yPosition = yData(lPos) : lPos += 1
					.TaxRatePercentage = yData(lPos) : lPos += 1
					.TaxRatePercType = CType(yData(lPos), eyGuildTaxPercType) : lPos += 1
					.TaxRateFlat = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					.lMembersInRank = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				End With
				If bAdded = True Then
					If moRanks Is Nothing Then ReDim moRanks(-1)
					ReDim Preserve moRanks(moRanks.GetUpperBound(0) + 1)
					moRanks(moRanks.GetUpperBound(0)) = oRank
				End If
		End Select
	End Sub
End Class

