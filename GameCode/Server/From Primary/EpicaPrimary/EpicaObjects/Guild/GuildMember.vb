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
	Public lMemberID As Int32 = -1		'playerid of this member
	Public yMemberState As GuildMemberState = GuildMemberState.Unaffiliated

	Public lCreateRankID As Int32		'ONLY USE WHEN THE GUILD IS BEING CREATED!!!

	Private moMember As Player = Nothing
	Public ReadOnly Property oMember() As Player
		Get
			If moMember Is Nothing Then
				moMember = GetEpicaPlayer(lMemberID)
			End If
			Return moMember
		End Get
	End Property

	Public Function GetObjectSmallString() As Byte()
		If oMember Is Nothing Then Return Nothing

		Dim yMsg(18) As Byte			'NOTE: Changing this len means changing the length in guild
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(lMemberID).CopyTo(yMsg, lPos) : lPos += 4
		yMsg(lPos) = yMemberState : lPos += 1
		System.BitConverter.GetBytes(oMember.lGuildRankID).CopyTo(yMsg, lPos) : lPos += 4
		yMsg(lPos) = oMember.yPlayerTitle : lPos += 1
		yMsg(lPos) = oMember.yGender : lPos += 1
		System.BitConverter.GetBytes(GetDateAsNumber(oMember.LastLogin.ToUniversalTime)).CopyTo(yMsg, lPos) : lPos += 4
		Dim lVal As Int32 = 0
		If (yMemberState And GuildMemberState.Approved) <> 0 Then
			lVal = oMember.lJoinedGuildOn
		End If
		System.BitConverter.GetBytes(lVal).CopyTo(yMsg, lPos) : lPos += 4
		Return yMsg
	End Function

	Public Function GetObjectAsString() As Byte()
		'Send back the details msg, which is compiled here???
		If oMember Is Nothing Then Return Nothing

		Dim yMsg() As Byte
		Dim lPos As Int32 = 0

		If (yMemberState And GuildMemberState.Applied) <> 0 Then
			ReDim yMsg(64)
		Else
			ReDim yMsg(8)
		End If

		System.BitConverter.GetBytes(lMemberID).CopyTo(yMsg, lPos) : lPos += 4
		yMsg(lPos) = yMemberState : lPos += 1

		With oMember
			If (yMemberState And GuildMemberState.Applied) <> 0 Then
                'applied members expose everything...
                'If .lLastGlobalRequestTurnIn = 0 OrElse glCurrentCycle - .lLastGlobalRequestTurnIn > 9000 Then
                '    goMsgSys.SendRequestGlobalPlayerScores(.ObjectID, 64, .ObjectID)
                '    If .lLastGlobalRequestTurnIn = 0 Then
                .lLGTechScore = .TechnologyScore
                .lLGDiplomacyScore = .DiplomacyScore
                .lLGWealthScore = .WealthScore
                .lLGProductionScore = .ProductionScore
                .lLGPopulationScore = .PopulationScore
                .lLGMilitaryScore = .lMilitaryScore
                .lLGTotalScore = .TotalScore
                '    End If
                'End If
                System.BitConverter.GetBytes(.lLGTechScore).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lLGDiplomacyScore).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lLGMilitaryScore).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lLGPopulationScore).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lLGProductionScore).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lLGWealthScore).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lLGTotalScore).CopyTo(yMsg, lPos) : lPos += 4

				System.BitConverter.GetBytes(.GetColonyCount).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(.lStartedEnvirID).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(.iStartedEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(.TutorialPhaseWaves).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(.PlayedTimeInTutorialOne).CopyTo(yMsg, lPos) : lPos += 4

				Dim lTimeInWaves As Int32 = 0
				If .PlayedTimeWhenFirstWave <> Int32.MinValue AndAlso .PlayedTimeAtEndOfWaves <> Int32.MinValue Then
					lTimeInWaves = .PlayedTimeAtEndOfWaves - .PlayedTimeWhenFirstWave
				End If
				System.BitConverter.GetBytes(lTimeInWaves).CopyTo(yMsg, lPos) : lPos += 4

				System.BitConverter.GetBytes(.blMaxPopulation).CopyTo(yMsg, lPos) : lPos += 8
			Else
                'non-applied members do not
                Dim lMax As Int32 = Math.Max(.TotalScore, .lLGTotalScore)
                System.BitConverter.GetBytes(lMax).CopyTo(yMsg, lPos) : lPos += 4
			End If
		End With

		Return yMsg
	End Function

	Public Function SaveObject(ByVal lGuildID As Int32) As Boolean
		Dim bResult As Boolean = False
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		Try
			sSQL = "UPDATE tblGuildMember SET MemberState = " & CByte(yMemberState) & " WHERE MemberID = " & lMemberID & " and GuildID = " & lGuildID
			oComm = Nothing
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			If oComm.ExecuteNonQuery() = 0 Then
				sSQL = "INSERT INTO tblGuildMember (MemberID, GuildID, MemberState) VALUES (" & lMemberID & ", " & lGuildID & ", " & CByte(yMemberState) & ")"
				oComm = Nothing
				oComm = New OleDb.OleDbCommand(sSQL, goCN)
				If oComm.ExecuteNonQuery() = 0 Then
					Err.Raise(-1, "SaveObject", "No records affected when saving object!")
				End If
			End If
			oComm = Nothing

			bResult = True
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object GuildMember. Reason: " & Err.Description)
		Finally
			oComm = Nothing
		End Try
		Return bResult
	End Function 

	Public Sub DeleteMe(ByVal lGuildID As Int32)
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		Try
			sSQL = "DELETE FROM tblGuildMember WHERE MemberID = " & Me.lMemberID & " and GuildID = " & lGuildID
			oComm = Nothing
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm = Nothing
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to delete GuildMember. Reason: " & Err.Description)
		Finally
			oComm = Nothing
		End Try

	End Sub
End Class
