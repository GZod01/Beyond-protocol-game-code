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
    'GuildChatChannel = 8192
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

	Public ParentGuild As Guild = Nothing

	Public lRankID As Int32 = -1
	Public yRankName(19) As Byte
	Public lRankPermissions As Int32
	Public lVoteStrength As Int32
	Public yPosition As Byte

	Public TaxRatePercentage As Byte = 0			'per interval in Guild
	Public TaxRatePercType As eyGuildTaxPercType = eyGuildTaxPercType.CashFlow
	Public TaxRateFlat As Int32 = 0					'per interval in Guild

	Public Function SaveObject(ByVal lParentGuildID As Int32) As Boolean
		Dim bResult As Boolean = False
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

		Try
			If lRankID = -1 Then
				'INSERT
				sSQL = "INSERT INTO tblRank (GuildID, RankName, RankRights, VotingWeight, PeckingOrder, TaxPerc, TaxPercType, TaxFlat) VALUES (" & _
				  lParentGuildID & ", '" & MakeDBStr(BytesToString(yRankName)) & "', " & lRankPermissions & ", " & _
				  lVoteStrength & ", " & yPosition & ", " & TaxRatePercentage & ", " & CInt(TaxRatePercType) & ", " & TaxRateFlat & ")"
			Else
				'UPDATE
				sSQL = "UPDATE tblRank SET GuildID = " & lParentGuildID & ", RankName = '" & _
				  MakeDBStr(BytesToString(yRankName)) & "', RankRights = " & lRankPermissions & ", VotingWeight = " & _
				  lVoteStrength & ", PeckingOrder = " & yPosition & ", TaxPerc = " & TaxRatePercentage & ", TaxPercType = " & _
				  CInt(TaxRatePercType) & ", TaxFlat = " & TaxRateFlat & " WHERE RankID = " & lRankID
			End If
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			If oComm.ExecuteNonQuery() = 0 Then
				Err.Raise(-1, "SaveObject", "No records affected when saving object!")
			End If
			If lRankID < 1 Then
				Dim oData As OleDb.OleDbDataReader
				oComm = Nothing
				sSQL = "SELECT MAX(RankID) FROM tblRank WHERE RankName = '" & MakeDBStr(BytesToString(yRankName)) & "'"
				oComm = New OleDb.OleDbCommand(sSQL, goCN)
				oData = oComm.ExecuteReader(CommandBehavior.Default)
				If oData.Read Then
					lRankID = CInt(oData(0))
				End If
				oData.Close()
				oData = Nothing
			End If

			bResult = True
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & lRankID & " of type Rank. Reason: " & Err.Description)
		Finally
			oComm = Nothing
		End Try
		Return bResult
	End Function

	Public Function GetObjectAsString() As Byte()
		Dim yMsg(42) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(lRankID).CopyTo(yMsg, lPos) : lPos += 4
		yRankName.CopyTo(yMsg, lPos) : lPos += 20
		System.BitConverter.GetBytes(lRankPermissions).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lVoteStrength).CopyTo(yMsg, lPos) : lPos += 4
		yMsg(lPos) = yPosition : lPos += 1
		yMsg(lPos) = TaxRatePercentage : lPos += 1
		yMsg(lPos) = TaxRatePercType : lPos += 1
		System.BitConverter.GetBytes(TaxRateFlat).CopyTo(yMsg, lPos) : lPos += 4

		Dim lCnt As Int32 = 0
		If ParentGuild Is Nothing = False Then
			lCnt = ParentGuild.GetMemberCountOfRank(Me.lRankID)
		End If
		System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4

		Return yMsg
	End Function

	Public Sub DeleteMe(ByVal lGuildID As Int32)
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		Try
			sSQL = "DELETE FROM tblRank WHERE GuildID = " & lGuildID & " AND RankID = " & Me.lRankID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to delete Rank. Reason: " & Err.Description)
		Finally
			oComm = Nothing
		End Try

	End Sub

	Public Function GetGuildDetailsMsg() As Byte()
		Dim yTmp() As Byte = GetObjectAsString()
		Dim yMsg(2 + yTmp.Length) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eGuildRequestDetails).CopyTo(yMsg, lPos) : lPos += 2
		yMsg(lPos) = eyGuildRequestDetailsType.GuildRank : lPos += 1
		yTmp.CopyTo(yMsg, lPos)
		Return yMsg
    End Function

    Public Function GetPlayerTaxCost(ByVal oPlayer As Player) As Int64
        'orank.TaxRateFlat
        Dim blTotal As Int64 = TaxRateFlat

        If TaxRatePercentage > 0 Then
            Select Case TaxRatePercType
                Case eyGuildTaxPercType.CashFlow
                    blTotal += CLng(oPlayer.oBudget.GetCashFlow() * (TaxRatePercentage / 100))
                Case eyGuildTaxPercType.TotalIncome
                    blTotal += CLng(oPlayer.oBudget.GetTotalIncome * (TaxRatePercentage / 100))
                Case eyGuildTaxPercType.Treasury
                    If oPlayer.blCredits > 0 Then
                        blTotal += CLng(oPlayer.blCredits * (TaxRatePercentage / 100))
                    End If
            End Select
        End If
        
        Return blTotal

    End Function
End Class