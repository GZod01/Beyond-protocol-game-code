Option Strict On

Public MustInherit Class EventAdvancedConfig
	Public lEventID As Int32
	Public MustOverride Function SaveObject() As Boolean
	Public Sub DeleteMe()
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		Try
			sSQL = "DELETE FROM tblEventAdvancedConfigItem WHERE EventID = " & Me.lEventID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()

			sSQL = "DELETE FROM tblEventAdvancedConfig WHERE EventID = " & Me.lEventID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
		Catch
		Finally
			oComm = Nothing
		End Try
	End Sub
	Public MustOverride Function GetMsg() As Byte() 

End Class

Public Class RaceConfig
	Inherits EventAdvancedConfig

	Public blEntryFee As Int64
	Public lMinHull As Int32
	Public lMaxHull As Int32
	Public yMinRacers As Byte
	Public yFirstPlace As Byte
	Public ySecondPlace As Byte
	Public yThirdPlace As Byte
	Public yGuildTake As Byte
	Public yLaps As Byte

	Public yGroundOnly As Byte

	Public Structure RaceWP
		Public EnvirID As Int32
		Public EnvirTypeID As Int16
		Public lX As Int32
		Public lZ As Int32
	End Structure
	Public lWPUB As Int32 = -1
	Public uWP() As RaceWP

	Public Function VerifyData() As Boolean
		Dim lTotal As Int32 = CInt(yFirstPlace) + CInt(ySecondPlace) + CInt(yThirdPlace) + CInt(yGuildTake)
		If blEntryFee <> 0 Then
			If lTotal <> 100 Then Return False
		End If
		If blEntryFee < 0 Then Return False
		If lMinHull < 0 Then Return False
		If lMinHull > lMaxHull Then Return False
		If lMaxHull < 0 Then Return False
		If yLaps < 1 Then Return False

		If yGroundOnly <> 0 Then
			For X As Int32 = 0 To lWPUB
				If uWP(X).EnvirTypeID = ObjectType.eSolarSystem Then Return False
			Next X
		End If

		Return True
	End Function

	Public Overrides Function GetMsg() As Byte()
		Dim lWPCnt As Int32 = lWPUB + 1

		Dim yMsg(33 + (lWPCnt * 14)) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eAdvancedEventConfig).CopyTo(yMsg, lPos) : lPos += 2
		yMsg(lPos) = 1 : lPos += 1		'indicating we are adding a race config
		System.BitConverter.GetBytes(lEventID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(blEntryFee).CopyTo(yMsg, lPos) : lPos += 8
		System.BitConverter.GetBytes(lMinHull).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lMaxHull).CopyTo(yMsg, lPos) : lPos += 4
		yMsg(lPos) = yMinRacers : lPos += 1
		yMsg(lPos) = yFirstPlace : lPos += 1
		yMsg(lPos) = ySecondPlace : lPos += 1
		yMsg(lPos) = yThirdPlace : lPos += 1
		yMsg(lPos) = yGuildTake : lPos += 1
		yMsg(lPos) = yLaps : lPos += 1
		yMsg(lPos) = yGroundOnly : lPos += 1

		System.BitConverter.GetBytes(lWPCnt).CopyTo(yMsg, lPos) : lPos += 4
		For X As Int32 = 0 To lWPUB
			With uWP(X)
				System.BitConverter.GetBytes(.EnvirID).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(.EnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(.lX).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(.lZ).CopyTo(yMsg, lPos) : lPos += 4
			End With
		Next X
		Return yMsg
	End Function

	Public Overrides Function SaveObject() As Boolean
		Dim bResult As Boolean = False
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		Try
			sSQL = "UPDATE tblEventAdvancedConfig SET EntryFee = " & blEntryFee & ", MinHull = " & lMinHull & ", MaxHull = " & _
			 lMaxHull & ", MinRacers = " & yMinRacers & ", FirstPlace = " & yFirstPlace & ", SecondPlace = " & ySecondPlace & _
			 ", ThirdPlace = " & yThirdPlace & ", GuildTake = " & yGuildTake & ", Laps = " & yLaps & ", GroundOnly = " & _
			 yGroundOnly & " WHERE EventID = " & lEventID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			If oComm.ExecuteNonQuery() = 0 Then
				oComm = Nothing
				sSQL = "INSERT tblEventAdvancedConfig (EntryFee, MinHull, MaxHull, MinRacers, FirstPlace, SecondPlace, ThirdPlace, " & _
				 "GuildTake, Laps, GroundOnly, EventID) VALUES (" & blEntryFee & ", " & lMinHull & ", " & lMaxHull & ", " & yMinRacers & _
				 ", " & yFirstPlace & ", " & ySecondPlace & ", " & yThirdPlace & ", " & yGuildTake & ", " & yLaps & ", " & yGroundOnly & _
				 ", " & lEventID & ")"
				oComm = New OleDb.OleDbCommand(sSQL, goCN)
				If oComm.ExecuteNonQuery() = 0 Then
					Err.Raise(-1, "SaveObject", "No records affected when saving object!")
				End If
			End If

			sSQL = "DELETE FROM tblEventAdvancedConfigItem WHERE EventID = " & Me.lEventID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm = Nothing

			For X As Int32 = 0 To lWPUB
				With uWP(X)
					sSQL = "INSERT INTO tblEventAdvancedConfigItem (EventID, ItemPos, EnvirID, EnvirTypeID, XPos, ZPos) VALUES(" & _
					 Me.lEventID & ", " & (X + 1) & ", " & .EnvirID & ", " & .EnvirTypeID & ", " & .lX & ", " & .lZ & ")"
					oComm = New OleDb.OleDbCommand(sSQL, goCN)
					oComm.ExecuteNonQuery()
					oComm = Nothing
				End With
			Next X

			bResult = True
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object tblEventAdvancedConfig. Reason: " & Err.Description)
		Finally
			oComm = Nothing
		End Try
		Return bResult
	End Function 
End Class

Public Class TournamentConfig
	Inherits EventAdvancedConfig

	Public blEntryFee As Int64
	Public lMapID As Int32
	Public yMaxPlayers As Byte
	Public yMaxUnits As Byte
	Public yMaxGround As Byte
	Public yMaxAir As Byte
	Public yGuildTake As Byte

	Public Overrides Function GetMsg() As Byte()
		Dim yMsg(23) As Byte
		Dim lPos As Int32 = 0

		System.BitConverter.GetBytes(GlobalMessageCode.eAdvancedEventConfig).CopyTo(yMsg, lPos) : lPos += 2
		yMsg(lPos) = 2 : lPos += 1		'indicating we are adding a race config
		System.BitConverter.GetBytes(lEventID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(blEntryFee).CopyTo(yMsg, lPos) : lPos += 8
		System.BitConverter.GetBytes(lMapID).CopyTo(yMsg, lPos) : lPos += 4
		yMsg(lPos) = yMaxPlayers : lPos += 1
		yMsg(lPos) = yMaxUnits : lPos += 1
		yMsg(lPos) = yMaxGround : lPos += 1
		yMsg(lPos) = yMaxAir : lPos += 1
		yMsg(lPos) = yGuildTake : lPos += 1

		Return yMsg
	End Function

	Public Overrides Function SaveObject() As Boolean
		Dim bResult As Boolean = False
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		Try
			sSQL = "UPDATE tblEventAdvancedConfig SET EntryFee = " & blEntryFee & ", MapID = " & lMapID & ", MaxPlayers = " & _
			 yMaxPlayers & ", MaxUnits = " & yMaxUnits & ", MaxGround = " & yMaxGround & ", MaxAir = " & yMaxAir & _
			 ", GuildTake = " & yGuildTake & " WHERE EventID = " & lEventID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			If oComm.ExecuteNonQuery() = 0 Then
				oComm = Nothing
				sSQL = "INSERT tblEventAdvancedConfig (EntryFee, MapID, MaxPlayers, MaxUnits, MaxGround, MaxAir, GuildTake, EventID) VALUES ("
				sSQL &= blEntryFee & ", " & lMapID & ", " & yMaxPlayers & ", " & yMaxUnits & ", " & yMaxGround & ", " & yMaxAir & _
				 ", " & yGuildTake & ", " & lEventID & ")"
				oComm = New OleDb.OleDbCommand(sSQL, goCN)
				If oComm.ExecuteNonQuery() = 0 Then
					Err.Raise(-1, "SaveObject", "No records affected when saving object!")
				End If
			End If

			bResult = True
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object tblEventAdvancedConfig. Reason: " & Err.Description)
		Finally
			oComm = Nothing
		End Try
		Return bResult
	End Function
 
End Class