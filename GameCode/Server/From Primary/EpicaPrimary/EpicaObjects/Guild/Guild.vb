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
	Inherits Epica_GUID

	Public yGuildName(49) As Byte
	Public yMOTD() As Byte
	Public yBillboard() As Byte

	Public sSearchableName As String

	Public dtFormed As Date = Date.MinValue						'in GMT
	Public dtLastGuildRuleChange As Date = Date.MinValue		'in GMT

	Public lGuildHallID As Int32 = -1
	Private moGuildHall As Facility = Nothing
	Public ReadOnly Property oGuildHall() As Facility
		Get
			If moGuildHall Is Nothing OrElse moGuildHall.ObjectID <> lGuildHallID Then
				moGuildHall = GetEpicaFacility(lGuildHallID)
				If moGuildHall Is Nothing Then lGuildHallID = -1
			ElseIf moGuildHall.ServerIndex < 0 OrElse glFacilityIdx(moGuildHall.ServerIndex) <> moGuildHall.ObjectID Then
				moGuildHall = Nothing
			End If
			Return moGuildHall
		End Get
	End Property

	Public lBaseGuildFlags As elGuildFlags
	Public lIcon As Int32 = 0

	Public yState As eyGuildState = eyGuildState.Proposed
	Public iRecruitFlags As eiRecruitmentFlags = 0

	Public yVoteWeightType As eyVoteWeightType = eyVoteWeightType.RankBased

	Public yGuildTaxRateInterval As eyGuildInterval = eyGuildInterval.Monthly
	Public yGuildTaxBaseMonth As Byte = 1
    Public yGuildTaxBaseDay As Byte = 1
    Public dtLastTaxInterval As Date = Date.MinValue

    Public blTreasury As Int64 = 0
    Public blLastTaxIncome As Int64 = 0

	'Public MemberCount As Int32 = 0
	'Public ReadOnly Property MemberCount() As Int32
	'	Get
	'		Dim lCnt As Int32 = 0
	'		For X As Int32 = 0 To lMemberUB
	'			If oMembers(X) Is Nothing = False Then
	'				If (oMembers(X).yMemberState And GuildMemberState.Approved) <> 0 Then lCnt += 1
	'			End If
	'		Next X
	'		Return lCnt
	'	End Get
	'End Property

	Public oEvents(-1) As GuildEvent
	Public lEventUB As Int32 = -1
	Public oRanks(-1) As GuildRank
	Public lRankUB As Int32 = -1
	Public oVotes(-1) As GuildVote
	Public lVoteUB As Int32 = -1
	Public oMembers(-1) As GuildMember
    Public lMemberUB As Int32 = -1
    Public oBankLog(-1) As GuildTransLog
    Public lBankLogUB As Int32 = -1

    'Public oRels(-1) As GuildRel
    'Public lRelUB As Int32 = -1
    
	Public Function GetMemberCountOfRank(ByVal lRankID As Int32) As Int32
		Dim lCnt As Int32 = 0
		For X As Int32 = 0 To lMemberUB
			If oMembers(X) Is Nothing = False Then
				If oMembers(X).oMember Is Nothing = False AndAlso oMembers(X).oMember.lGuildRankID = lRankID Then lCnt += 1
			End If
		Next X
		Return lCnt
    End Function

    Public Sub AddBankLogItem(ByRef oItem As GuildTransLog)
        SyncLock oBankLog
            ReDim Preserve oBankLog(lBankLogUB + 1)

            For X As Int32 = lBankLogUB + 1 To 1 Step -1
                oBankLog(X) = oBankLog(X - 1)
            Next X
            oBankLog(0) = oItem
            lBankLogUB += 1
        End SyncLock
    End Sub

	Public Sub AddRank(ByRef oRank As GuildRank)
		For X As Int32 = 0 To lRankUB
			If oRanks(X) Is Nothing Then
				oRanks(X) = oRank
				Return
			End If
		Next X
		SyncLock oRanks
			ReDim Preserve oRanks(lRankUB + 1)
			oRanks(lRankUB + 1) = oRank
			lRankUB += 1
		End SyncLock
	End Sub

	Public Sub AddEvent(ByRef oEvent As GuildEvent)
		oEvent.ParentGuild = Me
		For X As Int32 = 0 To lEventUB
			If oEvents(X) Is Nothing Then
				oEvents(X) = oEvent
				Return
			End If
		Next X
		SyncLock oEvents
			ReDim Preserve oEvents(lEventUB + 1)
			oEvents(lEventUB + 1) = oEvent
			lEventUB += 1
		End SyncLock
	End Sub

    Public Sub AddVote(ByRef oVote As GuildVote)
        For X As Int32 = 0 To lVoteUB
            If oVotes(X) Is Nothing Then
                oVotes(X) = oVote
                Return
            End If
        Next X
        SyncLock oVotes
            ReDim Preserve oVotes(lVoteUB + 1)
            oVotes(lVoteUB + 1) = oVote
            lVoteUB += 1
        End SyncLock
    End Sub

	Public Sub AddMember(ByRef oMember As GuildMember)
		For X As Int32 = 0 To lMemberUB
			If oMembers(X) Is Nothing Then
				oMembers(X) = oMember
				Return
			End If
		Next X
		SyncLock oMembers
			ReDim Preserve oMembers(lMemberUB + 1)
			oMembers(lMemberUB + 1) = oMember
			lMemberUB += 1
		End SyncLock
	End Sub

	Public Sub AddMemberVoteValue(ByVal lVoteID As Int32, ByVal lMemberID As Int32, ByVal yValue As eyVoteValue)
		For X As Int32 = 0 To lVoteUB
			If oVotes(X) Is Nothing = False AndAlso oVotes(X).VoteID = lVoteID Then
				oVotes(X).SetMemberVote(lMemberID, yValue)
				Exit For
			End If
		Next X
	End Sub

	Public Function GetApprovedMemberCount() As Int32
		Dim lMemberCount As Int32 = 0
		For X As Int32 = 0 To lMemberUB
			If oMembers(X) Is Nothing = False AndAlso (oMembers(X).yMemberState And GuildMemberState.Approved) <> 0 Then
				lMemberCount += 1
			End If
		Next X
		Return lMemberCount
	End Function

	Public Function GetObjAsString() As Byte()
		'here we will return the entire object as a string
		Dim lMemberCount As Int32 = 0
		For X As Int32 = 0 To lMemberUB
			If oMembers(X) Is Nothing = False Then
				lMemberCount += 1
			End If
		Next X

        'Dim lRelCnt As Int32 = 0
        'For X As Int32 = 0 To lRelUB
        '	If oRels(X) Is Nothing = False Then lRelCnt += 1
        'Next X
		Dim lRankCnt As Int32 = 0
		For X As Int32 = 0 To lRankUB
			If oRanks(X) Is Nothing = False Then lRankCnt += 1
		Next X

		Dim lPos As Int32 = 0

		Dim lMOTDLen As Int32 = 0
		Dim lBillboardLen As Int32 = 0

		If yMOTD Is Nothing = False Then lMOTDLen = yMOTD.Length
		If yBillboard Is Nothing = False Then lBillboardLen = yBillboard.Length

		Dim lGuildHallLen As Int32 = 4
		If oGuildHall Is Nothing = False AndAlso oGuildHall.ParentObject Is Nothing = False Then
			lGuildHallLen += 14
		End If

        Dim yMsg(102 + (lMemberCount * 19) + lMOTDLen + lBillboardLen + (lRankCnt * 43) + lGuildHallLen) As Byte

		Me.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
		yGuildName.CopyTo(yMsg, lPos) : lPos += 50
		Dim lVal As Int32 = 0
		If dtFormed <> Date.MinValue Then lVal = GetDateAsNumber(dtFormed)
		System.BitConverter.GetBytes(lVal).CopyTo(yMsg, lPos) : lPos += 4
		If dtLastGuildRuleChange <> Date.MinValue Then lVal = GetDateAsNumber(dtLastGuildRuleChange)
		System.BitConverter.GetBytes(lVal).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lIcon).CopyTo(yMsg, lPos) : lPos += 4

		System.BitConverter.GetBytes(lBaseGuildFlags).CopyTo(yMsg, lPos) : lPos += 4
		yMsg(lPos) = yState : lPos += 1
		System.BitConverter.GetBytes(iRecruitFlags).CopyTo(yMsg, lPos) : lPos += 2
		yMsg(lPos) = yVoteWeightType : lPos += 1
		yMsg(lPos) = yGuildTaxRateInterval : lPos += 1
		yMsg(lPos) = yGuildTaxBaseMonth : lPos += 1
		yMsg(lPos) = yGuildTaxBaseDay : lPos += 1

		System.BitConverter.GetBytes(blTreasury).CopyTo(yMsg, lPos) : lPos += 8

		If oGuildHall Is Nothing = False AndAlso oGuildHall.ParentObject Is Nothing = False Then
			System.BitConverter.GetBytes(oGuildHall.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
			CType(oGuildHall.ParentObject, Epica_GUID).GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
			System.BitConverter.GetBytes(oGuildHall.LocX).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(oGuildHall.LocZ).CopyTo(yMsg, lPos) : lPos += 4
		Else
			System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
		End If

		System.BitConverter.GetBytes(lMOTDLen).CopyTo(yMsg, lPos) : lPos += 4
		If yMOTD Is Nothing = False Then yMOTD.CopyTo(yMsg, lPos)
		lPos += lMOTDLen
		System.BitConverter.GetBytes(lBillboardLen).CopyTo(yMsg, lPos) : lPos += 4
		If yBillboard Is Nothing = False Then yBillboard.CopyTo(yMsg, lPos)
		lPos += lBillboardLen

		System.BitConverter.GetBytes(lMemberCount).CopyTo(yMsg, lPos) : lPos += 4
		Dim lBasePos As Int32 = lPos
		For X As Int32 = 0 To lMemberUB
			If oMembers(X) Is Nothing = False Then
				Dim yTmp() As Byte = oMembers(X).GetObjectSmallString()
				If yTmp Is Nothing = False Then
					yTmp.CopyTo(yMsg, lPos) : lPos += yTmp.Length
				End If
			End If
		Next X

		If lPos <> lBasePos + (lMemberCount * 19) Then
			lPos = lBasePos + (lMemberCount * 19)
		End If

        'System.BitConverter.GetBytes(lRelCnt).CopyTo(yMsg, lPos) : lPos += 4
        'For X As Int32 = 0 To lRelUB
        '	If oRels(X) Is Nothing = False Then
        '		Dim yTmp() As Byte = oRels(X).GetObjAsSmallString()
        '		If yTmp Is Nothing = False Then
        '			yTmp.CopyTo(yMsg, lPos) : lPos += yTmp.Length
        '		End If
        '	End If
        'Next X

		System.BitConverter.GetBytes(lRankCnt).CopyTo(yMsg, lPos) : lPos += 4
		For X As Int32 = 0 To lRankUB
            If oRanks(X) Is Nothing = False Then
                oRanks(X).ParentGuild = Me
                Dim yTmp() As Byte = oRanks(X).GetObjectAsString()
                If yTmp Is Nothing = False Then
                    yTmp.CopyTo(yMsg, lPos) : lPos += yTmp.Length
                End If
            End If
		Next X

		Return yMsg
	End Function
 
	Public Function SaveObject() As Boolean
		Dim bResult As Boolean = False
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

		Try
			If ObjectID = -1 Then
				'INSERT
				sSQL = "INSERT INTO tblGuild (GuildName, MOTD, DateFormed, DateLastRuleChange, GuildHallID, BaseGuildFlags, GuildState, "
                sSQL &= "RecruitFlags, VoteWeightType, TaxRateInterval, TaxBaseMonth, TaxBaseDay, Billboard, GuildIcon, Treasury, DateTaxed) VALUES ('" & MakeDBStr(BytesToString(yGuildName)) & "', '"
				If yMOTD Is Nothing = False Then sSQL &= MakeDBStr(BytesToString(yMOTD)) & "', " Else sSQL &= "', "
				Dim lVal As Int32 = 0
				If dtFormed <> Date.MinValue Then lVal = GetDateAsNumber(dtFormed)
				sSQL &= lVal.ToString & ", "
				If dtLastGuildRuleChange <> Date.MinValue Then lVal = GetDateAsNumber(dtLastGuildRuleChange)
				sSQL &= lVal.ToString & ", "
				If oGuildHall Is Nothing = False Then sSQL &= oGuildHall.ObjectID & ", " Else sSQL &= "-1, "
				sSQL &= CInt(lBaseGuildFlags) & ", " & CInt(yState) & ", " & CInt(iRecruitFlags) & ", " & CInt(yVoteWeightType) & ", "
				sSQL &= CInt(yGuildTaxRateInterval) & ", " & CInt(yGuildTaxBaseMonth) & ", " & CInt(yGuildTaxBaseDay) & ", '"
				If yBillboard Is Nothing = False Then sSQL &= MakeDBStr(BytesToString(yBillboard))
                sSQL &= "', " & lIcon & ", " & blTreasury '& ")"
                If dtLastTaxInterval <> Date.MinValue Then lVal = GetDateAsNumber(dtLastTaxInterval) Else lVal = 0
                sSQL &= ", " & lVal & ")"
			Else
				'UPDATE
				sSQL = "UPDATE tblGuild SET GuildName = '" & MakeDBStr(BytesToString(yGuildName)) & "', MOTD = '"
				If yMOTD Is Nothing = False Then sSQL &= MakeDBStr(BytesToString(yMOTD))
				Dim lFormed As Int32 = 0
				If dtFormed <> Date.MinValue Then lFormed = GetDateAsNumber(dtFormed)
				Dim lLastRule As Int32 = 0
				If dtLastGuildRuleChange <> Date.MinValue Then lLastRule = GetDateAsNumber(dtLastGuildRuleChange) Else lLastRule = lFormed
				sSQL &= "', DateFormed = " & lFormed & ", DateLastRuleChange = " & lLastRule & ", GuildHallID = "
				If oGuildHall Is Nothing = False Then sSQL &= oGuildHall.ObjectID Else sSQL &= "-1"
				sSQL &= ", BaseGuildFlags = " & CInt(lBaseGuildFlags) & ", GuildState = " & CInt(yState) & ", RecruitFlags = " & CInt(iRecruitFlags)
				sSQL &= ", VoteWeightType = " & CInt(yVoteWeightType) & ", TaxRateInterval = " & CInt(yGuildTaxRateInterval) & ", TaxBaseMonth = "
				sSQL &= CInt(yGuildTaxBaseMonth) & ", TaxBaseDay = " & CInt(yGuildTaxBaseDay) & ", Billboard = '"
				If yBillboard Is Nothing = False Then sSQL &= MakeDBStr(BytesToString(yBillboard))
                sSQL &= "', GuildIcon = " & lIcon & ", Treasury = " & blTreasury & ", DateTaxed = "
                If dtLastTaxInterval <> Date.MinValue Then lFormed = GetDateAsNumber(dtLastTaxInterval) Else lFormed = 0
                sSQL &= lFormed & "  WHERE GuildID = " & Me.ObjectID
            End If
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			If oComm.ExecuteNonQuery() = 0 Then
				Err.Raise(-1, "SaveObject", "No records affected when saving object!")
			End If
			If ObjectID = -1 Then
				Dim oData As OleDb.OleDbDataReader
				oComm = Nothing
				sSQL = "SELECT MAX(GuildID) FROM tblGuild WHERE GuildName = '" & MakeDBStr(BytesToString(yGuildName)) & "'"
				oComm = New OleDb.OleDbCommand(sSQL, goCN)
				oData = oComm.ExecuteReader(CommandBehavior.Default)
				If oData.Read Then
					ObjectID = CInt(oData(0))
				End If
				oData.Close()
				oData = Nothing
			End If

			'Save remaining data...
			For X As Int32 = 0 To lEventUB
				If oEvents(X) Is Nothing = False Then
					oEvents(X).SaveEvent(Me.ObjectID)
				End If
			Next X
			For X As Int32 = 0 To lRankUB
				If oRanks(X) Is Nothing = False Then
					oRanks(X).SaveObject(Me.ObjectID)
				End If
			Next X
			For X As Int32 = 0 To lMemberUB
				If oMembers(X) Is Nothing = False Then
					oMembers(X).SaveObject(Me.ObjectID)
				End If
			Next X
			For X As Int32 = 0 To lVoteUB
				If oVotes(X) Is Nothing = False Then
					oVotes(X).SaveObject(Me.ObjectID)
				End If
			Next X
            'For X As Int32 = 0 To lRelUB
            '	If oRels(X) Is Nothing = False Then
            '		oRels(X).SaveObject(Me.ObjectID)
            '	End If
            'Next X


			bResult = True
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
		Finally
			oComm = Nothing
		End Try
		Return bResult
	End Function

	Public Function RankHasPermission(ByVal lRankID As Int32, ByVal lPermission As RankPermissions) As Boolean
		If lPermission = RankPermissions.AcceptApplicant AndAlso (lBaseGuildFlags And elGuildFlags.AcceptMemberToGuild_RA) = 0 Then Return False
		If lPermission = RankPermissions.ChangeRankVotingWeight AndAlso (lBaseGuildFlags And elGuildFlags.ChangeVotingWeight_RA) = 0 Then Return False
		If lPermission = RankPermissions.CreateRanks AndAlso (lBaseGuildFlags And elGuildFlags.CreateRank_RA) = 0 Then Return False
		If lPermission = RankPermissions.DeleteRanks AndAlso (lBaseGuildFlags And elGuildFlags.DeleteRank_RA) = 0 Then Return False
		If lPermission = RankPermissions.DemoteMember AndAlso (lBaseGuildFlags And elGuildFlags.DemoteMember_RA) = 0 Then Return False
		If lPermission = RankPermissions.PromoteMember AndAlso (lBaseGuildFlags And elGuildFlags.PromoteGuildMember_RA) = 0 Then Return False
		If (lPermission = RankPermissions.RemoveMember OrElse lPermission = RankPermissions.RejectMember) AndAlso (lBaseGuildFlags And elGuildFlags.RemoveGuildMember_RA) = 0 Then Return False

		For X As Int32 = 0 To lRankUB
			If oRanks(X) Is Nothing = False Then
				If oRanks(X).lRankID = lRankID Then
					Return (oRanks(X).lRankPermissions And lPermission) <> 0
				End If
			End If
		Next X
		Return False
	End Function

	Public Sub SendMsgToGuildMembers(ByRef yData() As Byte)

		If Me.yState = eyGuildState.Forming Then
			For X As Int32 = 0 To lMemberUB
				If oMembers(X) Is Nothing = False AndAlso oMembers(X).oMember Is Nothing = False Then
					If (oMembers(X).yMemberState And GuildMemberState.AcceptedGuildFormInvite) <> 0 Then
						oMembers(X).oMember.SendPlayerMessage(yData, False, 0)
					End If
				End If
			Next X
		Else
			For X As Int32 = 0 To lMemberUB
				If oMembers(X) Is Nothing = False AndAlso oMembers(X).oMember Is Nothing = False Then
					If (oMembers(X).yMemberState And GuildMemberState.Approved) <> 0 Then
						oMembers(X).oMember.SendPlayerMessage(yData, False, 0)
					End If
				End If
			Next X
		End If
		
	End Sub
    Public Sub SendMsgToGuildMembersWithOutbound(ByRef yData() As Byte, ByVal sOutEmail As String, ByVal sSubject As String, ByVal lFromPlayerID As Int32)

        Dim lNowAsNum As Int32 = GetDateAsNumber(Now)
        Dim iMsgCode As Int16 = System.BitConverter.ToInt16(yData, 0)

        For X As Int32 = 0 To lMemberUB
            If oMembers(X) Is Nothing = False AndAlso oMembers(X).oMember Is Nothing = False Then
                If (oMembers(X).yMemberState And GuildMemberState.Approved) <> 0 Then
                    With oMembers(X).oMember
                        If .lConnectedPrimaryID = glServerID Then
                            If .oSocket Is Nothing = False AndAlso .oSocket.IsConnected = True AndAlso .bInPlayerRequestDetails = False Then
                                .oSocket.SendData(yData)
                            Else
                                Dim oPC As PlayerComm = .AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sOutEmail, sSubject, lFromPlayerID, lNowAsNum, False, .sPlayerNameProper, Nothing)
                                If oPC Is Nothing = False Then
                                    goMsgSys.SendOutboundEmail(oPC, oMembers(X).oMember, iMsgCode, -1, -1, -1, -1, -1, -1, "")
                                End If
                            End If
                        End If
                    End With
                    'oMembers(X).oMember.SendPlayerMessage(yData, False, 0)
                End If
            End If
        Next X

    End Sub

	Public Sub SendMsgToGuildMembersByPermission(ByRef yData() As Byte, ByVal lPermission As RankPermissions)
		For X As Int32 = 0 To lMemberUB
			If oMembers(X) Is Nothing = False AndAlso oMembers(X).oMember Is Nothing = False Then
				If (oMembers(X).yMemberState And GuildMemberState.Approved) <> 0 Then
					If RankHasPermission(oMembers(X).oMember.lGuildRankID, lPermission) = True Then oMembers(X).oMember.SendPlayerMessage(yData, False, 0)
				End If
			End If
		Next X
	End Sub

	Public Function HandleRequestDetails(ByVal yRequestType As eyGuildRequestDetailsType, ByVal lItemID As Int32, ByVal iTypeID As Int16, ByVal lRankID As Int32, ByVal lPlayerID As Int32) As Byte()
		'take the rankid and check if the requested item type is able to be sent down...
		Select Case yRequestType
			Case eyGuildRequestDetailsType.EventDetails
				If RankHasPermission(lRankID, RankPermissions.ViewEvents) = False Then
					LogEvent(LogEventType.PossibleCheat, "Player requesting Event Details without permission: " & lPlayerID)
					Return Nothing
				End If

				Dim oEvent As GuildEvent = GetEvent(lItemID)
				If oEvent Is Nothing = False Then
					Dim yTmp() As Byte = oEvent.GetFullObjectString()
					If yTmp Is Nothing Then Return Nothing

					Dim yResp(2 + yTmp.Length) As Byte
					Dim lPos As Int32 = 0
					System.BitConverter.GetBytes(GlobalMessageCode.eGuildRequestDetails).CopyTo(yResp, lPos) : lPos += 2
					yResp(lPos) = yRequestType : lPos += 1
					yTmp.CopyTo(yResp, lPos) : lPos += yTmp.Length
					Return yResp
				End If
			Case eyGuildRequestDetailsType.GuildRel
                'Dim oRel As GuildRel = GetRel(lItemID, iTypeID)
                'If oRel Is Nothing = False Then
                '	Dim yTmp() As Byte = oRel.GetFullObjAsString()
                '	If yTmp Is Nothing Then Return Nothing

                '	Dim yResp(2 + yTmp.Length) As Byte
                '	Dim lPos As Int32 = 0
                '	System.BitConverter.GetBytes(GlobalMessageCode.eGuildRequestDetails).CopyTo(yResp, lPos) : lPos += 2
                '	yResp(lPos) = yRequestType : lPos += 1
                '	yTmp.CopyTo(yResp, lPos) : lPos += yTmp.Length
                '	Return yResp
                'End If
			Case eyGuildRequestDetailsType.MemberDetails
				If RankHasPermission(lRankID, RankPermissions.AcceptApplicant) = False Then Return Nothing

				Dim oMember As GuildMember = GetMember(lItemID)
				If oMember Is Nothing = False Then
					Dim yTmp() As Byte = oMember.GetObjectAsString()
					If yTmp Is Nothing Then Return Nothing

					Dim yResp(2 + yTmp.Length) As Byte
					Dim lPos As Int32 = 0
					System.BitConverter.GetBytes(GlobalMessageCode.eGuildRequestDetails).CopyTo(yResp, lPos) : lPos += 2
					yResp(lPos) = yRequestType : lPos += 1
					yTmp.CopyTo(yResp, lPos) : lPos += yTmp.Length
					Return yResp
				End If
		End Select
		Return Nothing
	End Function

	Public Sub SendGuildTreasuryUpdate()
		Dim yMsg(9) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eUpdateGuildTreasury).CopyTo(yMsg, 0)
		System.BitConverter.GetBytes(blTreasury).CopyTo(yMsg, 2)
		SendMsgToGuildMembers(yMsg)
	End Sub

	Public Function GetVote(ByVal lProposalID As Int32) As GuildVote
		For X As Int32 = 0 To lVoteUB
			If oVotes(X) Is Nothing = False AndAlso oVotes(X).VoteID = lProposalID Then Return oVotes(X)
		Next X
		Return Nothing
	End Function
	Public Function GetRank(ByVal lRankID As Int32) As GuildRank
		For X As Int32 = 0 To lRankUB
			If oRanks(X) Is Nothing = False AndAlso oRanks(X).lRankID = lRankID Then Return oRanks(X)
		Next X
		Return Nothing
	End Function
	Public Function GetMember(ByVal lMemberID As Int32) As GuildMember
		For X As Int32 = 0 To lMemberUB
			If oMembers(X) Is Nothing = False AndAlso oMembers(X).lMemberID = lMemberID Then Return oMembers(X)
		Next X
		Return Nothing
	End Function
	Public Function GetEvent(ByVal lEventID As Int32) As GuildEvent
		For X As Int32 = 0 To lEventUB
			If oEvents(X) Is Nothing = False AndAlso oEvents(X).EventID = lEventID Then Return oEvents(X)
		Next X
		Return Nothing
	End Function
    Public Function GetNextRankPosition(ByVal lPosDir As Int32, ByVal lStartRankID As Int32) As GuildRank
        Dim lRankPos As Int32 = Int32.MinValue
        For X As Int32 = 0 To lRankUB
            If oRanks(X) Is Nothing = False AndAlso oRanks(X).lRankID = lStartRankID Then
                lRankPos = oRanks(X).yPosition
                Exit For
            End If
        Next X
        If lRankPos = Int32.MinValue Then Return Nothing

        'now, find next rank pos
        Dim oResult As GuildRank = Nothing
        Dim lCurrBest As Int32 = 0
        lCurrBest = Int32.MaxValue

        For X As Int32 = 0 To lRankUB
            If oRanks(X) Is Nothing = False AndAlso oRanks(X).lRankID <> lStartRankID Then
                'k, now check our direction
                Dim lTmp As Int32 = CInt(oRanks(X).yPosition)
                If lPosDir < 0 Then
                    'ok, looking for next lowest rank position
                    If lTmp < lRankPos Then
                        If lRankPos - lTmp < lCurrBest Then
                            oResult = oRanks(X)
                            lCurrBest = lRankPos - lTmp
                        End If
                    End If
                Else
                    'ok, looking for next highest rank position
                    If lTmp > lRankPos Then
                        If lTmp - lRankPos < lCurrBest Then
                            oResult = oRanks(X)
                            lCurrBest = lTmp - lRankPos
                        End If
                    End If
                End If
            End If
        Next X

        Return oResult
    End Function
	Public Function VoteAlreadyInProgress(ByVal yTypeOfVote As Byte, ByVal lSelectedItem As Int32) As Boolean
		For X As Int32 = 0 To lVoteUB
			If oVotes(X) Is Nothing = False Then
				If oVotes(X).yTypeOfVote = yTypeOfVote AndAlso oVotes(X).lSelectedItem = lSelectedItem Then
					If oVotes(X).dtVoteStarts.ToLocalTime.AddHours(oVotes(X).lVoteDuration) > Now Then
						Return True
					End If
				End If
			End If
		Next X
		Return False
	End Function
	Public Sub DeleteEvent(ByVal lEventID As Int32)
		For X As Int32 = 0 To lEventUB
			If oEvents(X) Is Nothing = False AndAlso oEvents(X).EventID = lEventID Then
				oEvents(X).DeleteMe()
				oEvents(X) = Nothing
			End If
		Next X
	End Sub

	Public Function GetGuildAssetsMsg() As Byte()
		'If oGuildHall Is Nothing Then Return Nothing

		If oGuildHall Is Nothing Then
			Dim yTmp(5) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eGetGuildAssets).CopyTo(yTmp, 0)
			System.BitConverter.GetBytes(-1I).CopyTo(yTmp, 2)
			Return yTmp
		End If

		Dim lComponents As Int32 = 0
		Dim lMinerals As Int32 = 0
		'Dim lUnits As Int32 = 0

		'components and minerals
		For X As Int32 = 0 To oGuildHall.lCargoUB
			If oGuildHall.lCargoIdx(X) <> -1 Then
				If oGuildHall.oCargoContents(X).ObjTypeID = ObjectType.eComponentCache Then
					lComponents += 1
				ElseIf oGuildHall.oCargoContents(X).ObjTypeID = ObjectType.eMineralCache Then
					lMinerals += 1
				End If
			End If
		Next X
		''units 
		'For X As Int32 = 0 To oGuildHall.lHangarUB
		'	If oGuildHall.lHangarIdx(X) <> -1 Then
		'		If oGuildHall.oHangarContents(X).ObjTypeID = ObjectType.eUnit Then lUnits += 1
		'	End If
		'Next X

		Dim yResp(9 + (lComponents * 30) + (lMinerals * 26)) As Byte ' + (lUnits * 19)) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eGetGuildAssets).CopyTo(yResp, lPos) : lPos += 2
		System.BitConverter.GetBytes(oGuildHall.ObjectID).CopyTo(yResp, lPos) : lPos += 4
		System.BitConverter.GetBytes(lComponents + lMinerals).CopyTo(yResp, lPos) : lPos += 4

		'components and minerals
		For X As Int32 = 0 To oGuildHall.lCargoUB
			If oGuildHall.lCargoIdx(X) <> -1 Then
				If oGuildHall.oCargoContents(X).ObjTypeID = ObjectType.eComponentCache Then
					With CType(oGuildHall.oCargoContents(X), ComponentCache)
						System.BitConverter.GetBytes(.ObjectID).CopyTo(yResp, lPos) : lPos += 4
						System.BitConverter.GetBytes(CShort(.ComponentTypeID * -1S)).CopyTo(yResp, lPos) : lPos += 2
						System.BitConverter.GetBytes(.ComponentID).CopyTo(yResp, lPos) : lPos += 4
						System.BitConverter.GetBytes(.ComponentOwnerID).CopyTo(yResp, lPos) : lPos += 4
						System.BitConverter.GetBytes(.Quantity).CopyTo(yResp, lPos) : lPos += 4

						Dim oSO As SellOrder = goGTC.GetSellOrderItem(oGuildHall.ObjectID, .ComponentID, CShort(.ComponentTypeID * -1S))
						If oSO Is Nothing = False Then
							System.BitConverter.GetBytes(CInt(oSO.blQuantity)).CopyTo(yResp, lPos) : lPos += 4
							System.BitConverter.GetBytes(oSO.blPrice).CopyTo(yResp, lPos) : lPos += 8
						Else
							System.BitConverter.GetBytes(0I).CopyTo(yResp, lPos) : lPos += 4
							System.BitConverter.GetBytes(0L).CopyTo(yResp, lPos) : lPos += 8
						End If
					End With
				ElseIf oGuildHall.oCargoContents(X).ObjTypeID = ObjectType.eMineralCache Then
					With CType(oGuildHall.oCargoContents(X), MineralCache)
						.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6
						System.BitConverter.GetBytes(.oMineral.ObjectID).CopyTo(yResp, lPos) : lPos += 4
						System.BitConverter.GetBytes(.Quantity).CopyTo(yResp, lPos) : lPos += 4

						Dim oSO As SellOrder = goGTC.GetSellOrderItem(oGuildHall.ObjectID, .ObjectID, .ObjTypeID)
						If oSO Is Nothing = False Then
							System.BitConverter.GetBytes(CInt(oSO.blQuantity)).CopyTo(yResp, lPos) : lPos += 4
							System.BitConverter.GetBytes(oSO.blPrice).CopyTo(yResp, lPos) : lPos += 8
						Else
							System.BitConverter.GetBytes(0I).CopyTo(yResp, lPos) : lPos += 4
							System.BitConverter.GetBytes(0L).CopyTo(yResp, lPos) : lPos += 8
						End If
					End With
				End If
			End If
		Next X
		''units 
		'For X As Int32 = 0 To oTP.lHangarUB
		'	If oTP.lHangarIdx(X) <> -1 Then
		'		If oTP.oHangarContents(X).ObjTypeID = ObjectType.eUnit Then
		'			With CType(oTP.oHangarContents(X), Unit)
		'				.GetGUIDAsString.CopyTo(yResp, lPos) : lPos += 6

		'				Dim oSO As SellOrder = goGTC.GetSellOrderItem(oTP.ObjectID, .ObjectID, .ObjTypeID)
		'				If oSO Is Nothing = False Then
		'					System.BitConverter.GetBytes(CInt(oSO.blQuantity)).CopyTo(yResp, lPos) : lPos += 4
		'					System.BitConverter.GetBytes(oSO.blPrice).CopyTo(yResp, lPos) : lPos += 8
		'				Else
		'					System.BitConverter.GetBytes(0I).CopyTo(yResp, lPos) : lPos += 4
		'					System.BitConverter.GetBytes(0L).CopyTo(yResp, lPos) : lPos += 8
		'				End If

		'				yResp(lPos) = .EntityDef.yChassisType : lPos += 1
		'			End With
		'		End If
		'	End If
		'Next X

		'goGTC.FillWithIntelItemsBeingsold(yResp, lPos, lTradePostID)

		Return yResp
	End Function
	Public Function ProcessGuildTransaction(ByVal yTransType As Byte, ByVal bHighSec As Boolean, ByVal lItemID As Int32, ByVal iTypeID As Int16, ByVal blQty As Int64, ByRef oTradepost As Facility) As Boolean
		If oGuildHall Is Nothing Then Return False
		If oTradepost Is Nothing Then Return False

		'Credits: 1, 38, -1
		'Agents: AgentID, AgentTypeID
		'Components: ComponentCacheID, -ComponentTypeID
		'Minerals: MinCacheID, CacheTypeID, MineralID
		'Player Intel: PlayerID, PlayerIntelTypeID, PlayerTypeID?
		'Playeritemintel: TechID, PTKTypeID, TechTypeID
		'ColonyLocIntel: colonyid, PII TypeID, Colonytypeid
		'Units: ObjectID, ObjTypeID

		Dim iOrigTypeID As Int16 = iTypeID
		If iTypeID < 0 Then iTypeID = ObjectType.eComponentCache

		'ok, determine what we are doing
		If yTransType = 1 Then		'deposit
			'take items from the tradepost and place into the guild hall
			Select Case iTypeID
				Case ObjectType.eCredits
					Dim blActual As Int64 = Math.Min(blQty, oTradepost.Owner.blCredits)
					If blActual < 1 Then Return False
					oTradepost.Owner.blCredits -= blActual
					blTreasury += blActual
					SendGuildTreasuryUpdate()
				Case ObjectType.eComponentCache
					For X As Int32 = 0 To oTradepost.lCargoUB
						If oTradepost.lCargoIdx(X) <> -1 Then
							If oTradepost.oCargoContents(X).ObjTypeID = ObjectType.eComponentCache Then
								With CType(oTradepost.oCargoContents(X), ComponentCache)
									If .ObjectID = lItemID AndAlso .ComponentTypeID = Math.Abs(iOrigTypeID) Then
										Dim blActual As Int64 = Math.Min(blQty, .Quantity)
										If blActual < 1 Then Return False
										If .Quantity = blActual Then
											.ParentObject = oGuildHall
											oGuildHall.AddCargoRef(oTradepost.oCargoContents(X))
											oTradepost.lCargoIdx(X) = -1
											oTradepost.oCargoContents(X) = Nothing
											Exit For
										Else
											.Quantity -= CInt(blActual)
											oGuildHall.AddComponentCacheToCargo(.ComponentID, .ComponentTypeID, CInt(blActual), .ComponentOwnerID)
											Exit For
										End If
									End If
								End With
							End If
						End If
					Next X
				Case ObjectType.eMineralCache
					For X As Int32 = 0 To oTradepost.lCargoUB
						If oTradepost.lCargoIdx(X) <> -1 Then
							If oTradepost.oCargoContents(X).ObjTypeID = ObjectType.eMineralCache Then
								With CType(oTradepost.oCargoContents(X), MineralCache)
									If .ObjectID = lItemID Then
										Dim blActual As Int64 = Math.Min(blQty, .Quantity)
										If blActual < 1 Then Return False
										If .Quantity = blActual Then
											.ParentObject = oGuildHall
											oGuildHall.AddCargoRef(oTradepost.oCargoContents(X))
											oTradepost.lCargoIdx(X) = -1
											oTradepost.oCargoContents(X) = Nothing
											Exit For
										Else
											.Quantity -= CInt(blActual)
											oGuildHall.AddMineralCacheToCargo(.oMineral.ObjectID, CInt(blActual))
											Exit For
										End If
									End If
								End With
							End If
						End If
					Next X
			End Select
		ElseIf yTransType = 2 Then	'withdrawal
			'take items from the guild hall and place into the tradepost
			Select Case iTypeID
				Case ObjectType.eCredits
					Dim blActual As Int64 = Math.Min(blQty, blTreasury)
					If blActual < 1 Then Return False
					blTreasury -= blActual
					oTradepost.Owner.blCredits += blActual
					SendGuildTreasuryUpdate()
				Case ObjectType.eComponentCache
					For X As Int32 = 0 To oGuildHall.lCargoUB
						If oGuildHall.lCargoIdx(X) <> -1 Then
							If oGuildHall.oCargoContents(X).ObjTypeID = ObjectType.eComponentCache Then
								With CType(oGuildHall.oCargoContents(X), ComponentCache)
									If .ObjectID = lItemID AndAlso .ComponentTypeID = Math.Abs(iOrigTypeID) Then
										Dim blActual As Int64 = Math.Min(blQty, .Quantity)
										If blActual < 1 Then Return False
										If .Quantity = blActual Then
											.ParentObject = oTradepost
											oTradepost.AddCargoRef(oGuildHall.oCargoContents(X))
											oGuildHall.lCargoIdx(X) = -1
											oGuildHall.oCargoContents(X) = Nothing
											Exit For
										Else
											.Quantity -= CInt(blActual)
											oTradepost.AddComponentCacheToCargo(.ComponentID, .ComponentTypeID, CInt(blActual), .ComponentOwnerID)
											Exit For
										End If
									End If
								End With
							End If
						End If
					Next X
				Case ObjectType.eMineralCache
					For X As Int32 = 0 To oGuildHall.lCargoUB
						If oGuildHall.lCargoIdx(X) <> -1 Then
							If oGuildHall.oCargoContents(X).ObjTypeID = ObjectType.eMineralCache Then
								With CType(oGuildHall.oCargoContents(X), MineralCache)
									If .ObjectID = lItemID Then
										Dim blActual As Int64 = Math.Min(blQty, .Quantity)
										If blActual < 1 Then Return False
										If .Quantity = blActual Then
											.ParentObject = oTradepost
											oTradepost.AddCargoRef(oGuildHall.oCargoContents(X))
											oGuildHall.lCargoIdx(X) = -1
											oGuildHall.oCargoContents(X) = Nothing
											Exit For
										Else
											.Quantity -= CInt(blActual)
											oTradepost.AddMineralCacheToCargo(.oMineral.ObjectID, CInt(blActual))
											Exit For
										End If
									End If
								End With
							End If
						End If
					Next X
			End Select
		Else : Return False
		End If
		Return True
	End Function
	Public Sub RemoveMember(ByVal lMemberID As Int32)
		'delete the record too?
		For X As Int32 = 0 To lMemberUB
            If oMembers(X) Is Nothing = False AndAlso oMembers(X).lMemberID = lMemberID Then

                If (oMembers(X).yMemberState And GuildMemberState.Approved) <> 0 Then
                    'Now, send the message to the email server
                    Dim yEmail(10) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eSetGuildRel).CopyTo(yEmail, 0)
                    System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yEmail, 2)
                    System.BitConverter.GetBytes(lMemberID).CopyTo(yEmail, 6)
                    yEmail(10) = 0
                    goMsgSys.SendToEmailSrvr(yEmail)
                End If

                oMembers(X).DeleteMe(Me.ObjectID)
                oMembers(X) = Nothing

                Return
            End If
		Next X
	End Sub
	Public Sub MemberJoined(ByVal lPlayerID As Int32, ByVal bIgnoreRankSetup As Boolean)
		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing Then Return

		Dim oMember As GuildMember = GetMember(lPlayerID)
		Dim bAdded As Boolean = False
		If oMember Is Nothing Then
			oMember = New GuildMember()
			bAdded = True
		End If

		oMember.oMember.lJoinedGuildOn = GetDateAsNumber(Now.ToUniversalTime)
		If bAdded = True OrElse bIgnoreRankSetup = False Then
			oMember.lMemberID = lPlayerID
			oMember.yMemberState = GuildMemberState.Approved
			oMember.oMember.lGuildID = Me.ObjectID
			oMember.oMember.lGuildRankID = -1
			oMember.oMember.lInvitedToJoinUB = -1
			Dim lMaxPos As Int32 = -1
			Dim lMaxPosRankID As Int32 = -1
			For X As Int32 = 0 To lRankUB
				If oRanks(X) Is Nothing = False AndAlso oRanks(X).yPosition > lMaxPos Then
					lMaxPos = oRanks(X).yPosition
					lMaxPosRankID = oRanks(X).lRankID
				End If
			Next X
			oMember.oMember.lGuildRankID = lMaxPosRankID
		End If
		If bAdded = True Then Me.AddMember(oMember)


		'check for require peace
		If (lBaseGuildFlags And elGuildFlags.RequirePeaceBetweenMembers) <> 0 Then
			For X As Int32 = 0 To oPlayer.PlayerRelUB
				Dim oRel As PlayerRel = oPlayer.GetPlayerRelByIndex(X)
				If oRel Is Nothing = False Then
					If oRel.WithThisScore <= elRelTypes.eWar Then
						If oRel.oThisPlayer.lGuildID = Me.ObjectID Then
							Dim yTmp(10) As Byte
							System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerRel).CopyTo(yTmp, 0)
							System.BitConverter.GetBytes(oRel.oPlayerRegards.ObjectID).CopyTo(yTmp, 2)
							System.BitConverter.GetBytes(oRel.oThisPlayer.ObjectID).CopyTo(yTmp, 6)
							yTmp(10) = elRelTypes.eNeutral

							goMsgSys.HandleSetPlayerRel(yTmp, -1)
							System.BitConverter.GetBytes(oRel.oPlayerRegards.ObjectID).CopyTo(yTmp, 6)
							System.BitConverter.GetBytes(oRel.oThisPlayer.ObjectID).CopyTo(yTmp, 2)
							goMsgSys.HandleSetPlayerRel(yTmp, -1)
						End If
					End If
				End If
			Next X
		End If

        ''ok, find any rels that the player has that the guild does not have
        'For X As Int32 = 0 To oPlayer.PlayerRelUB
        '	Dim oRel As PlayerRel = oPlayer.GetPlayerRelByIndex(X)
        '	If oRel Is Nothing = False Then
        '		Dim lOtherPlayerID As Int32 = -1
        '		If oRel.oPlayerRegards.ObjectID = lPlayerID Then
        '			lOtherPlayerID = oRel.oThisPlayer.ObjectID
        '		Else : lOtherPlayerID = oRel.oPlayerRegards.ObjectID
        '		End If
        '		If lOtherPlayerID <> -1 Then
        '			Dim oGuildRel As GuildRel = GetRel(lOtherPlayerID, ObjectType.ePlayer)
        '			If oGuildRel Is Nothing Then
        '				'ok, don't have this one
        '				oGuildRel = New GuildRel()
        '				With oGuildRel
        '					.dtWhenFirstContactMade = Now.ToUniversalTime
        '					.iEntityTypeID = ObjectType.ePlayer
        '					.iLocationTypeIDOfFC = -1
        '					.lEntityID = lOtherPlayerID
        '					.lLocationIDOfFC = -1
        '					.lLocXOfFC = 0
        '					.lLocZOfFC = 0
        '					.lWhoFirstContactWasMadeWith = lOtherPlayerID
        '					.lWhoMadeFirstContact = lPlayerID
        '					.yNote = StringToBytes("Acquired when " & oPlayer.sPlayerNameProper & " joined.")
        '					.yRelTowardsThem = elRelTypes.eNeutral
        '					.yRelTowardsUs = elRelTypes.eNeutral
        '				End With
        '				AddRel(oGuildRel)

        '				Dim yForward(15) As Byte
        '				System.BitConverter.GetBytes(GlobalMessageCode.eSetGuildRel).CopyTo(yForward, 0)
        '				System.BitConverter.GetBytes(oGuildRel.lEntityID).CopyTo(yForward, 2)
        '				System.BitConverter.GetBytes(oGuildRel.iEntityTypeID).CopyTo(yForward, 6)
        '				yForward(8) = oGuildRel.yRelTowardsThem
        '				yForward(9) = oGuildRel.yRelTowardsUs
        '				Me.GetGUIDAsString.CopyTo(yForward, 10)

        '				Me.SendMsgToGuildMembers(yForward)

        '				If oRel.oPlayerRegards.ObjectID = lPlayerID Then
        '					oRel.oThisPlayer.SendPlayerMessage(yForward, False, 0)
        '				Else : oRel.oPlayerRegards.SendPlayerMessage(yForward, False, 0)
        '				End If
        '			End If
        '		End If
        '	End If
        'Next X

        'Now, send the message to the email server
        Dim yEmail(10) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eSetGuildRel).CopyTo(yEmail, 0)
        System.BitConverter.GetBytes(Me.ObjectID).CopyTo(yEmail, 2)
        System.BitConverter.GetBytes(oPlayer.ObjectID).CopyTo(yEmail, 6)
        yEmail(10) = 1
        goMsgSys.SendToEmailSrvr(yEmail)


		'Now, send off a msg to everyone
		Me.SendMsgToGuildMembers(MsgSystem.CreateChatMsg(-1, oPlayer.sPlayerNameProper & " has joined " & BytesToString(Me.yGuildName) & "!", ChatMessageType.eSysAdminMessage))

		'Finally, send an addobject of the guild to the player
		Me.SendMsgToGuildMembers(goMsgSys.GetAddObjectMessage(Me, GlobalMessageCode.eAddObjectCommand))
        'oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(Me, GlobalMessageCode.eAddObjectCommand), False, 0)
    End Sub
	Public Sub PlayerApplied(ByVal lPlayerID As Int32)
		'send emails to those with rights to recruit
		Dim oPlayer As Player = GetEpicaPlayer(lPlayerID)
		If oPlayer Is Nothing Then Return

		Dim sBody As String = oPlayer.sPlayerNameProper & " has applied for membership with " & BytesToString(Me.yGuildName) & ". To view the application, open the guild window and go to the membership tab."

		For X As Int32 = 0 To lMemberUB
			If oMembers(X) Is Nothing = False AndAlso oMembers(X).oMember Is Nothing = False Then
				If (lBaseGuildFlags And elGuildFlags.AcceptMemberToGuild_RA) = 0 OrElse Me.RankHasPermission(oMembers(X).oMember.lGuildRankID, RankPermissions.AcceptApplicant) = True Then
					If (oMembers(X).oMember.iInternalEmailSettings And eEmailSettings.eGuildMembershipNotices) <> 0 Then
                        Dim oPC As PlayerComm = oMembers(X).oMember.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, sBody, "Guild Application", lPlayerID, GetDateAsNumber(Now), False, oMembers(X).oMember.sPlayerNameProper, Nothing)
						If oPC Is Nothing = False Then
							oMembers(X).oMember.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
						End If
						oPC = Nothing
					End If
				End If
			End If
		Next X
	End Sub
	Public Sub MoveRank(ByVal lRankID As Int32, ByVal yMoveDir As Byte)
		'0 for nothing, 1 for up, 2 for down, 255 for remove
		If yMoveDir = 1 Then
			'move up... find next lowest rank number
			Dim oRank As GuildRank = GetRank(lRankID)
			If oRank Is Nothing Then Return
			Dim oNextRank As GuildRank = GetNextRankPosition(-1, lRankID)
			If oNextRank Is Nothing Then Return
			Dim lTmp As Int32 = oNextRank.yPosition
			oNextRank.yPosition = oRank.yPosition
			oRank.yPosition = CByte(lTmp)
		Else
			'move down
			Dim oRank As GuildRank = GetRank(lRankID)
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

		For X As Int32 = 0 To lMemberUB
			If oMembers(X) Is Nothing = False Then
				If oMembers(X).oMember Is Nothing = False AndAlso oMembers(X).oMember.lGuildRankID = lRankID Then
					oMembers(X).oMember.lGuildRankID = oNextRank.lRankID
				End If
			End If
		Next X

		For X As Int32 = 0 To lRankUB
			If oRanks(X) Is Nothing = False AndAlso oRanks(X).lRankID = lRankID Then
				oRanks(X).DeleteMe(Me.ObjectID)
				oRanks(X) = Nothing
			End If
		Next X
	End Sub

	Private Shared mlPreviousGuildCheck As Int32 = -1
	Public Shared Sub GuildCheck()
		Dim dtNow As Date = Now
		Dim lNowMin As Int32 = dtNow.Minute \ 5
		If lNowMin <> mlPreviousGuildCheck Then
			mlPreviousGuildCheck = lNowMin

			'process our guilds...
			Dim lCurUB As Int32 = -1
			If glGuildIdx Is Nothing = False Then lCurUB = Math.Min(glGuildUB, glGuildIdx.GetUpperBound(0))
			For X As Int32 = 0 To lCurUB
				If glGuildIdx(X) <> -1 Then
					Dim oGuild As Guild = goGuild(X)
					If oGuild Is Nothing = False Then
						For Y As Int32 = 0 To oGuild.lVoteUB
							If oGuild.oVotes(Y) Is Nothing = False AndAlso oGuild.oVotes(Y).yVoteState = eyVoteState.eInProgress Then
								If oGuild.oVotes(Y).dtVoteStarts.AddHours(oGuild.oVotes(Y).lVoteDuration) < dtNow Then
									oGuild.oVotes(Y).yVoteState = eyVoteState.eProcessed
									'ok, let's process it
									oGuild.oVotes(Y).ProcessVote(oGuild)
								End If
							End If
                        Next Y

                        Dim bTax As Boolean = False
                        With oGuild
                            Select Case oGuild.yGuildTaxRateInterval
                                Case eyGuildInterval.Daily
                                    If .dtLastTaxInterval = Date.MinValue OrElse .dtLastTaxInterval.Day <> dtNow.Day Then
                                        bTax = True
                                        .dtLastTaxInterval = dtNow
                                    End If
                                Case eyGuildInterval.Weekly
                                    If .dtLastTaxInterval = Date.MinValue OrElse .dtLastTaxInterval.AddDays(5) < dtNow Then
                                        If dtNow.DayOfWeek >= .yGuildTaxBaseDay Then
                                            bTax = True
                                            .dtLastTaxInterval = dtNow
                                        End If
                                    End If
                                Case eyGuildInterval.SemiMonthly
                                    Dim lDueDay As Int32 = .dtLastTaxInterval.Day
                                    Dim lDueMonth As Int32 = .dtLastTaxInterval.Month
                                    Dim lDueYear As Int32 = .dtLastTaxInterval.Year
                                    If lDueDay = 1 Then
                                        lDueDay = 15
                                    Else
                                        lDueDay = 1
                                        lDueMonth += 1
                                        If lDueMonth > 12 Then
                                            lDueMonth -= 12
                                            lDueYear += 1
                                        End If
                                    End If
                                    Dim dtDueDate As Date = New Date(lDueYear, lDueMonth, Math.Min(Date.DaysInMonth(lDueYear, lDueMonth), lDueDay))
                                    If dtDueDate < dtNow AndAlso (.dtLastTaxInterval = Date.MinValue OrElse .dtLastTaxInterval <= dtDueDate) Then
                                        'yes, did we already pay the taxes?
                                        If .dtLastTaxInterval = Date.MinValue OrElse .dtLastTaxInterval.Year < dtNow.Year OrElse .dtLastTaxInterval.Month < dtDueDate.Month OrElse .dtLastTaxInterval.Day < dtDueDate.Day Then
                                            bTax = True
                                            .dtLastTaxInterval = dtNow
                                        End If
                                    End If
                                Case eyGuildInterval.Monthly
                                    'ok, check if the day has past
                                    If .yGuildTaxBaseDay <= dtNow.Day Then
                                        'yes, it has past
                                        If .dtLastTaxInterval = Date.MinValue OrElse .dtLastTaxInterval.Month < dtNow.Month OrElse (.dtLastTaxInterval.Month = 12 AndAlso dtNow.Month = 1) Then
                                            bTax = True
                                            .dtLastTaxInterval = dtNow
                                        End If
                                    End If
                                Case eyGuildInterval.EveryTwoMonths
                                    Dim lDueMonth As Int32 = .dtLastTaxInterval.Month
                                    lDueMonth += 2
                                    Dim lDueYear As Int32 = .dtLastTaxInterval.Year
                                    If lDueMonth > 12 Then
                                        lDueMonth -= 12
                                        lDueYear += 1
                                    End If
                                    Dim dtDueDate As Date = New Date(lDueYear, lDueMonth, Math.Min(Date.DaysInMonth(lDueYear, lDueMonth), .yGuildTaxBaseDay))
                                    If dtDueDate < dtNow AndAlso (.dtLastTaxInterval = Date.MinValue OrElse .dtLastTaxInterval <= dtDueDate) Then
                                        'yes, did we already pay the taxes?
                                        If .dtLastTaxInterval = Date.MinValue OrElse .dtLastTaxInterval.Year < dtNow.Year OrElse .dtLastTaxInterval.Month < dtDueDate.Month OrElse .dtLastTaxInterval.Day < dtDueDate.Day Then
                                            bTax = True
                                            .dtLastTaxInterval = dtNow
                                        End If
                                    End If

                                Case eyGuildInterval.Quarterly
                                    Dim lDueMonth As Int32 = .dtLastTaxInterval.Month
                                    lDueMonth += 3
                                    Dim lDueYear As Int32 = .dtLastTaxInterval.Year
                                    If lDueMonth > 12 Then
                                        lDueMonth -= 12
                                        lDueYear += 1
                                    End If
                                    Dim dtDueDate As Date = New Date(lDueYear, lDueMonth, Math.Min(Date.DaysInMonth(lDueYear, lDueMonth), .yGuildTaxBaseDay))
                                    If dtDueDate < dtNow AndAlso (.dtLastTaxInterval = Date.MinValue OrElse .dtLastTaxInterval <= dtDueDate) Then
                                        'yes, did we already pay the taxes?
                                        If .dtLastTaxInterval = Date.MinValue OrElse .dtLastTaxInterval.Year < dtNow.Year OrElse .dtLastTaxInterval.Month < dtDueDate.Month OrElse .dtLastTaxInterval.Day < dtDueDate.Day Then
                                            bTax = True
                                            .dtLastTaxInterval = dtNow
                                        End If
                                    End If
                                Case eyGuildInterval.SemiAnnually
                                    Dim lDueMonth As Int32 = .dtLastTaxInterval.Month
                                    lDueMonth += 6
                                    Dim lDueYear As Int32 = .dtLastTaxInterval.Year
                                    If lDueMonth > 12 Then
                                        lDueMonth -= 12
                                        lDueYear += 1
                                    End If
                                    Dim dtDueDate As Date = New Date(lDueYear, lDueMonth, Math.Min(Date.DaysInMonth(lDueYear, lDueMonth), .yGuildTaxBaseDay))
                                    If dtDueDate < dtNow AndAlso (.dtLastTaxInterval = Date.MinValue OrElse .dtLastTaxInterval <= dtDueDate) Then
                                        'yes, did we already pay the taxes?
                                        If .dtLastTaxInterval = Date.MinValue OrElse .dtLastTaxInterval.Year < dtNow.Year OrElse .dtLastTaxInterval.Month < dtDueDate.Month OrElse .dtLastTaxInterval.Day < dtDueDate.Day Then
                                            bTax = True
                                            .dtLastTaxInterval = dtNow
                                        End If
                                    End If
                                Case eyGuildInterval.Annually
                                    Dim dtDueDate As Date = New Date(.dtLastTaxInterval.Year + 1, .yGuildTaxBaseMonth, Math.Min(Date.DaysInMonth(.dtLastTaxInterval.Year + 1, .yGuildTaxBaseMonth), .yGuildTaxBaseDay))
                                    If dtDueDate < dtNow AndAlso (.dtLastTaxInterval = Date.MinValue OrElse .dtLastTaxInterval <= dtDueDate) Then
                                        'yes, did we already pay the taxes?
                                        If .dtLastTaxInterval = Date.MinValue OrElse .dtLastTaxInterval.Year < dtNow.Year OrElse .dtLastTaxInterval.Month < dtDueDate.Month OrElse .dtLastTaxInterval.Day < dtDueDate.Day Then
                                            bTax = True
                                            .dtLastTaxInterval = dtNow
                                        End If
                                    End If
                            End Select

                            If bTax = True Then
                                Dim oSB As New System.Text.StringBuilder()
                                Dim oSBDetail As New System.Text.StringBuilder()

                                Dim blTotalDue As Int64 = 0
                                Dim blTotalPaid As Int64 = 0
                                oSB.AppendLine("Guild Tax Report for guild taxes collected on " & Now.ToString("MM/dd/yyyy"))

                                oSBDetail.AppendLine("Collections Per Member:")

                                Dim lSortedIdx() As Int32 = Nothing
                                Dim lSortedUB As Int32 = -1
                                For Y As Int32 = 0 To .lMemberUB
                                    If .oMembers(Y) Is Nothing = False Then
                                        If (.oMembers(Y).yMemberState And GuildMemberState.Approved) <> 0 Then
                                            'ok, do it
                                            Dim oPlayer As Player = .oMembers(Y).oMember
                                            If oPlayer Is Nothing = False Then
                                                Dim sNewValue As String = oPlayer.sPlayerName
                                                Dim lIdx As Int32 = -1
                                                For Z As Int32 = 0 To lSortedUB
                                                    If .oMembers(lSortedIdx(Z)).oMember.sPlayerName > sNewValue Then
                                                        lIdx = Z
                                                        Exit For
                                                    End If
                                                Next Z

                                                lSortedUB += 1
                                                ReDim Preserve lSortedIdx(lSortedUB)
                                                If lIdx = -1 Then
                                                    lSortedIdx(lSortedUB) = Y
                                                Else
                                                    For Z As Int32 = lSortedUB To lIdx + 1 Step -1
                                                        lSortedIdx(Z) = lSortedIdx(Z - 1)
                                                    Next Z
                                                    lSortedIdx(lIdx) = Y
                                                End If

                                            End If
                                        End If
                                    End If
                                Next Y

                                For Y As Int32 = 0 To lSortedUB ' .lMemberUB
                                    'by rank
                                    Dim lIdx As Int32 = lSortedIdx(Y)
                                    If .oMembers(lIdx) Is Nothing = False Then
                                        If (.oMembers(lIdx).yMemberState And GuildMemberState.Approved) <> 0 Then
                                            'ok, do it
                                            Dim oPlayer As Player = .oMembers(lIdx).oMember
                                            If oPlayer Is Nothing = False Then
                                                Dim oRank As GuildRank = .GetRank(oPlayer.lGuildRankID)
                                                If oRank Is Nothing = False Then
                                                    Dim blTax As Int64 = oRank.GetPlayerTaxCost(oPlayer)
                                                    blTotalDue += blTax
                                                    If blTax <= 0 Then
                                                        oSBDetail.AppendLine("  " & oPlayer.sPlayerNameProper & " did not owe taxes.")
                                                        'ElseIf blTax < 0 Then
                                                        '    If .blTreasury > Math.Abs(blTax) Then
                                                        '        .blTreasury -= Math.Abs(blTax)
                                                        '        oPlayer.blCredits += Math.Abs(blTax)
                                                        '        oSBDetail.AppendLine("  " & oPlayer.sPlayerNameProper & " received " & Math.Abs(blTax).ToString("#,##0") & " credits from the guild due to the tax settings.")
                                                        '    End If
                                                    Else
                                                        If oPlayer.blCredits < blTax Then
                                                            If oPlayer.blCredits < 0 Then
                                                                oSBDetail.AppendLine("  " & oPlayer.sPlayerNameProper & " could not pay " & blTax.ToString("#,##0") & ". The entire balance is outstanding.")
                                                            Else
                                                                Dim blDiff As Int64 = blTax - oPlayer.blCredits
                                                                blTotalPaid += oPlayer.blCredits
                                                                oPlayer.blCredits = 0
                                                                oSBDetail.AppendLine("  " & oPlayer.sPlayerNameProper & " could not pay " & blTax.ToString("#,##0") & ". Outstanding balance: " & blDiff.ToString("#,##0"))
                                                            End If
                                                        Else
                                                            oPlayer.blCredits -= blTax
                                                            blTotalPaid += blTax
                                                            oSBDetail.AppendLine("  " & oPlayer.sPlayerNameProper & " paid the total balance of " & blTax.ToString("#,##0"))
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                Next Y

                                oSB.AppendLine("Total Taxes Due: " & blTotalDue.ToString("#,##0"))
                                oSB.AppendLine("Total Taxes Paid: " & blTotalPaid.ToString("#,##0"))

                                oSB.AppendLine()
                                oSB.Append(oSBDetail.ToString)

                                .blLastTaxIncome = blTotalPaid
                                Try
                                    .blTreasury += blTotalPaid
                                Catch
                                End Try

                                For Y As Int32 = 0 To .lMemberUB
                                    If .oMembers(Y) Is Nothing = False Then
                                        If (.oMembers(Y).yMemberState And GuildMemberState.Approved) <> 0 Then
                                            'ok, do it
                                            Dim oPlayer As Player = .oMembers(Y).oMember
                                            If oPlayer Is Nothing = False Then
                                                Dim oPC As PlayerComm = oPlayer.AddEmailMsg(-1, PlayerCommFolder.ePCF_ID_HARDCODES.eInbox_PCF, 0, oSB.ToString, "Guild Tax Report", oPlayer.ObjectID, GetDateAsNumber(Now), False, oPlayer.sPlayerNameProper, Nothing)
                                                If oPC Is Nothing = False Then
                                                    oPlayer.SendPlayerMessage(goMsgSys.GetAddObjectMessage(oPC, GlobalMessageCode.eAddObjectCommand), False, AliasingRights.eViewEmail)
                                                End If
                                            End If
                                        End If
                                    End If
                                Next Y

                                .SaveObject()

                            End If
                        End With
                        
					End If
				End If
			Next X
		End If
	End Sub

	Public Function GetLowestRank() As GuildRank
		If oRanks Is Nothing = False Then
			Dim yLowest As Byte = 0
			Dim oRank As GuildRank = Nothing
			For X As Int32 = 0 To oRanks.GetUpperBound(0)
				If oRanks(X) Is Nothing = False AndAlso oRanks(X).yPosition >= yLowest Then
					yLowest = oRanks(X).yPosition
					oRank = oRanks(X)
				End If
			Next X
			Return oRank
		End If
		Return Nothing
    End Function

    Public Function GetHighestPlayerTitle() As Byte
        Dim yMaxVal As Byte = 0
        For X As Int32 = 0 To lMemberUB
            If oMembers(X) Is Nothing = False Then
                If (oMembers(X).yMemberState And GuildMemberState.Approved) <> 0 Then
                    Dim yTemp As Byte = oMembers(X).oMember.yPlayerTitle
                    If (yTemp And Player.PlayerRank.ExRankShift) <> 0 Then yTemp = yTemp Xor Player.PlayerRank.ExRankShift
                    yMaxVal = Math.Max(yMaxVal, yTemp)
                End If
            End If
        Next X
        Return yMaxVal
    End Function

    Public Function GetNextHighestSystemByColonyCount(ByVal lPrevSysID As Int32, ByVal bAllowSpawnSystem As Boolean) As Int32
        Dim lSystems(-1) As Int32
        Dim lColonyCnt(-1) As Int32
        Dim lSystemUB As Int32 = -1


        Dim lPrevSysCnt As Int32 = Int32.MaxValue

        For X As Int32 = 0 To lMemberUB
            If oMembers(X) Is Nothing = False AndAlso oMembers(X).oMember Is Nothing = False AndAlso (oMembers(X).yMemberState And GuildMemberState.Approved) <> 0 Then
                With oMembers(X).oMember
                    For Y As Int32 = 0 To .mlColonyUB
                        If .mlColonyIdx(Y) > -1 Then
                            If glColonyIdx(.mlColonyIdx(Y)) = .mlColonyID(Y) Then
                                Dim oColony As Colony = goColony(.mlColonyIdx(Y))
                                If oColony Is Nothing = False Then
                                    If oColony.ParentObject Is Nothing = False Then
                                        Dim oNewParent As Epica_GUID = CType(oColony.ParentObject, Epica_GUID)
                                        Dim iTemp As Int16 = oNewParent.ObjTypeID

                                        Dim lSystemID As Int32 = -1

                                        If iTemp = ObjectType.eSolarSystem Then
                                            If CType(oNewParent, SolarSystem).SystemType = 0 AndAlso bAllowSpawnSystem = False Then Continue For
                                            lSystemID = oNewParent.ObjectID
                                        ElseIf iTemp = ObjectType.ePlanet Then
                                            With CType(oNewParent, Planet)
                                                If .ParentSystem Is Nothing = False Then
                                                    If .ParentSystem.SystemType = 0 Then Continue For
                                                    lSystemID = .ParentSystem.ObjectID
                                                End If
                                            End With
                                        ElseIf iTemp = ObjectType.eFacility Then
                                            oNewParent = CType(CType(oNewParent, Facility).ParentObject, Epica_GUID)
                                            If oNewParent Is Nothing = False AndAlso oNewParent.ObjTypeID = ObjectType.eSolarSystem Then
                                                If CType(oNewParent, SolarSystem).SystemType = 0 Then Continue For
                                                lSystemID = oNewParent.ObjectID
                                            End If
                                        End If

                                        If lSystemID = -1 Then Continue For

                                        Dim bFound As Boolean = False
                                        For Z As Int32 = 0 To lSystemUB
                                            If lSystems(Z) = lSystemID Then
                                                bFound = True
                                                lColonyCnt(Z) += 1

                                                If lSystemID = lPrevSysID Then lPrevSysCnt = lColonyCnt(Z)

                                                Exit For
                                            End If
                                        Next Z

                                        If bFound = False Then
                                            lSystemUB += 1
                                            ReDim Preserve lSystems(lSystemUB)
                                            ReDim Preserve lColonyCnt(lSystemUB)

                                            lSystems(lSystemUB) = lSystemID
                                            lColonyCnt(lSystemUB) = 1

                                            If lSystemID = lPrevSysID Then lPrevSysCnt = lColonyCnt(lSystemUB)
                                        End If


                                    End If
                                End If
                            End If
                        End If
                    Next Y
                End With
            End If
        Next X

        'ok, we now effectively have a list of systems the guild belongs to
        Dim lMaxCnt As Int32 = 0
        Dim lMaxID As Int32 = -1
        For X As Int32 = 0 To lSystemUB
            If lColonyCnt(X) < lPrevSysCnt Then
                If lColonyCnt(X) > lMaxCnt Then
                    lMaxCnt = lColonyCnt(X)
                    lMaxID = lSystems(X)
                End If
            End If
        Next X

        Return lMaxID

    End Function

End Class

'Public Class Guild
'    Inherits Epica_GUID

'    Public Enum ForeignPolicyShare As Byte
'        eNoSharing = 0
'        eLooseSharing = 1
'        eFullSharing = 2
'    End Enum
'    Public Enum ForeignPolicyVote As Byte
'        eNoVoting = 0
'        eDemocratic = 1
'        eDelegate = 2
'        eOfficers = 3
'    End Enum
'    Public Enum VotingWeight As Byte
'        eOneVote = 0
'        ePopulation = 1
'        eRankWeighted = 2
'    End Enum
'    Public Enum ImpeachmentRule As Byte
'        eNoImpeachment = 0
'        eAnyMember = 1
'        eHighestRank = 2
'    End Enum
'    Public Enum ReoccurIntervalType As Byte
'        eDaily = 0          'once per day
'        eWeekly = 1         'once per week
'        eSemiMonthly = 2    'twice per month
'        eMonthly = 3        'once per month
'        eBiMonthly = 4      'once per 2 months
'        eQuarterly = 5      'once per 3 months
'        eSemiAnnually = 6   'once per 6 months
'        eAnnually = 7       'once per 12 months
'    End Enum
'    Public Enum MemberRights As Integer
'        eNoRights = 0
'        eCastVote = 1
'        eViewFinancials = 2
'        eInitiateVote = 4
'    End Enum

'    Public yGuildName(49) As Byte
'    Public lDelegatePlayerID As Int32 = -1

'    Public bRequirePeaceBetweenMembers As Boolean
'    Public yForeignPolicyShare As Byte
'    Public yForeignPolicyVote As Byte
'    Public bAutomaticTradeAgreements As Boolean
'    Public yVotingWeight As Byte
'    Public bMemberOpenBorders As Boolean
'    Public yImpeachmentRule As Byte

'    Public TaxPercentage As Byte
'    Public TaxFlatRate As Int64
'    Public TaxInterval As Byte      'uses the ReoccurIntervalType
'    Public DelegateRevotePeriod As Byte     'uses the ReoccurIntervalType

'    Public MemberCount As Int32 = 1     'TODO: Need to set this somewhere

'    'Sorted by Pecking Order. Pecking Order of 0 is top dog.
'    Private moRanks() As MemberRank
'    Private mlRankUB As Int32 = -1

'    Private moDelegatePlayer As Player = Nothing
'    Public ReadOnly Property oGuildDelegate() As Player
'        Get
'            If moDelegatePlayer Is Nothing OrElse moDelegatePlayer.ObjectID <> lDelegatePlayerID Then
'                moDelegatePlayer = Nothing
'                If lDelegatePlayerID <> -1 Then
'                    moDelegatePlayer = GetEpicaPlayer(lDelegatePlayerID)
'                End If
'            End If
'            Return moDelegatePlayer
'        End Get
'    End Property

'    Public Function GetObjAsString() As Byte()
'        'TODO: Finish me
'    End Function

'    Public Function SaveObject() As Boolean
'        Dim bResult As Boolean = False
'        Dim sSQL As String
'        Dim oComm As OleDb.OleDbCommand

'        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

'        Try
'            If ObjectID = -1 Then
'                'INSERT
'                sSQL = "INSERT INTO tblGuild (GuildName, DelegateID, RequirePeaceBetweenMembers, AutomaticTrade, " & _
'                  "MemberOpenBorders, ForeignPolicyShare, ForeignPolicyVote, VotingWeight, ImpeachmentRule, TaxPercentage, " & _
'                  "TaxFlatRate, TaxInterval, DelegateRevotePeriod) VALUES ('" & MakeDBStr(BytesToString(yGuildName)) & _
'                  "', " & lDelegatePlayerID & ", "
'                If Me.bRequirePeaceBetweenMembers = True Then sSQL &= "1, " Else sSQL &= "0, "
'                If Me.bAutomaticTradeAgreements = True Then sSQL &= "1, " Else sSQL &= "0, "
'                If Me.bMemberOpenBorders = True Then sSQL &= "1, " Else sSQL &= "0, "
'                sSQL &= yForeignPolicyShare & ", " & yForeignPolicyVote & ", " & yVotingWeight & ", " & yImpeachmentRule & _
'                  ", " & TaxPercentage & ", " & TaxFlatRate & ", " & TaxInterval & ", " & DelegateRevotePeriod & ")"
'            Else
'                'UPDATE
'                sSQL = "UPDATE tblGuild SET GuildName = '" & MakeDBStr(BytesToString(yGuildName)) & "', DelegateID = " & _
'                  lDelegatePlayerID & ", "
'                If bRequirePeaceBetweenMembers = True Then sSQL &= "RequirePeaceBetweenMembers = 1, " Else sSQL &= "RequirePeaceBetweenMembers = 0, "
'                If bAutomaticTradeAgreements = True Then sSQL &= "AutomaticTrade = 1, " Else sSQL &= "AutomaticTrade = 0, "
'                If bMemberOpenBorders = True Then sSQL &= "MemberOpenBorders = 1, " Else sSQL &= "MemberOpenBorders = 0, "
'                sSQL &= "ForeignPolicyShare = " & yForeignPolicyShare & ", ForeignPolicyVote = " & yForeignPolicyVote & _
'                  ", VotingWeight = " & yVotingWeight & ", ImpeachmentRule = " & yImpeachmentRule & ", TaxPercentage = " & _
'                  TaxPercentage & ", TaxFlatRate = " & TaxFlatRate & ", TaxInterval = " & TaxInterval & _
'                  ", DelegateRevotePeriod = " & DelegateRevotePeriod & " WHERE GuildID = " & Me.ObjectID
'            End If
'            oComm = New OleDb.OleDbCommand(sSQL, goCN)
'            If oComm.ExecuteNonQuery() = 0 Then
'                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
'            End If
'            If ObjectID = -1 Then
'                Dim oData As OleDb.OleDbDataReader
'                oComm = Nothing
'                sSQL = "SELECT MAX(GuildID) FROM tblGuild WHERE GuildName = '" & MakeDBStr(BytesToString(yGuildName)) & "'"
'                oComm = New OleDb.OleDbCommand(sSQL, goCN)
'                oData = oComm.ExecuteReader(CommandBehavior.Default)
'                If oData.Read Then
'                    ObjectID = CInt(oData(0))
'                End If
'                oData.Close()
'                oData = Nothing
'            End If

'            'Now, save our ranks
'            For X As Int32 = 0 To mlRankUB
'                moRanks(X).SaveObject(Me.ObjectID)
'            Next X

'            bResult = True
'        Catch
'            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
'        Finally
'            oComm = Nothing
'        End Try
'        Return bResult
'    End Function

'    Public Sub AddMemberRank(ByRef oRank As MemberRank)
'        Dim lIdx As Int32 = -1

'        For X As Int32 = 0 To mlRankUB
'            If moRanks(X).yPeckingOrder > oRank.yPeckingOrder Then
'                lIdx = X
'                Exit For
'            End If
'        Next X

'        'Ok, if lidx = -1
'        If lIdx = -1 Then
'            'Just put this rank at the end
'            mlRankUB += 1
'            ReDim Preserve moRanks(mlRankUB)
'            moRanks(mlRankUB) = oRank
'        Else
'            'Ok, make room
'            ReDim Preserve moRanks(mlRankUB + 1)
'            For X As Int32 = mlRankUB To lIdx Step -1
'                moRanks(X + 1) = moRanks(X)
'            Next X
'            moRanks(lIdx) = oRank
'            mlRankUB += 1
'        End If

'    End Sub

'    Public Sub DeleteRank(ByVal lRankID As Int32)
'        'Need to scoot the ranks after this rank up to fill the gap
'        'If the last rank is being deleted, players of that rank are promoted, otherwise, players of the rank are demoted
'    End Sub

'End Class