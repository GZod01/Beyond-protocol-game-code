Option Strict On

Public Class GuildEvent
	Public EventID As Int32 = -1
	Public ParentGuild As Guild
	Public yTitle(49) As Byte
	Public lPostedBy As Int32 = -1
	Public dtPostedOn As Date = Date.MinValue		'in gmt
	Public dtStartsAt As Date = Date.MinValue		'in gmt
	Public lDuration As Int32 = 1					'in minutes

	Public yEventType As Byte = 0
	Public ySendAlerts As Byte = 0
	Public yEventIcon As Byte = 0
	Public yMembersCanAccept As Byte = 0

	Public yRecurrence As Byte = 0	'no recurrence

	Public yDetails() As Byte = Nothing		'variable length

	Public Attachments() As SketchPad
	Public AttachmentUB As Int32 = -1

	Public oAdvancedConfig As EventAdvancedConfig = Nothing

	Private Structure uAcceptance
		Public lPlayerID As Int32
		Public yAcceptance As Byte
	End Structure
	Private muAcceptance() As uAcceptance
	Private mlAcceptanceUB As Int32 = -1

	Public Function GetSmallString() As Byte()
		Dim yData(8) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(EventID).CopyTo(yData, lPos) : lPos += 4
		yData(lPos) = yEventIcon : lPos += 1

		Dim lVal As Int32 = GetDateAsNumber(dtStartsAt)
		System.BitConverter.GetBytes(lVal).CopyTo(yData, lPos) : lPos += 4

		Return yData
	End Function

	Public Function GetFullObjectString() As Byte()

		Dim lDetailLen As Int32 = 0
		If yDetails Is Nothing = False Then lDetailLen = yDetails.Length
		Dim yMsg(76 + lDetailLen) As Byte
		Dim lPos As Int32 = 0

		System.BitConverter.GetBytes(EventID).CopyTo(yMsg, lPos) : lPos += 4
		yTitle.CopyTo(yMsg, lPos) : lPos += 50
		System.BitConverter.GetBytes(lPostedBy).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(GetDateAsNumber(dtPostedOn)).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(GetDateAsNumber(dtStartsAt)).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lDuration).CopyTo(yMsg, lPos) : lPos += 4
		yMsg(lPos) = yEventIcon : lPos += 1

		'first 4 bits are the eventtype, 5th bit is members can accept, last 3 bits are send alerts
		Dim yTmpVal As Byte = yEventType
		'cccbaaaa
		If yMembersCanAccept <> 0 Then
			yTmpVal = CByte(yTmpVal Or 16)
		End If
		yTmpVal = CByte(yTmpVal Or (ySendAlerts * 32))
		yMsg(lPos) = yTmpVal : lPos += 1

		yMsg(lPos) = yRecurrence : lPos += 1

		If lDetailLen <> 0 Then
			System.BitConverter.GetBytes(lDetailLen).CopyTo(yMsg, lPos) : lPos += 4
			yDetails.CopyTo(yMsg, lPos) : lPos += yDetails.Length
		Else
			System.BitConverter.GetBytes(lDetailLen).CopyTo(yMsg, lPos) : lPos += 4
		End If

		Return yMsg
	End Function
 
	Public Function SaveEvent(ByVal lGuildID As Int32) As Boolean
		Dim bResult As Boolean = False
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

		Try
			Dim lPostedOn As Int32 = 0
			Dim lStartAt As Int32 = 0
			If dtPostedOn <> Date.MinValue Then lPostedOn = GetDateAsNumber(dtPostedOn)
			If dtStartsAt <> Date.MinValue Then lStartAt = GetDateAsNumber(dtStartsAt)

			If EventID = -1 Then
				'INSERT
				sSQL = "INSERT INTO tblGuildEvent (GuildID, PostedByID, PostedOn, StartsAt, Duration, EventType, SendAlerts, EventIcon, " & _
				  "MembersCanAccept, Recurrence, Details, EventTitle) VALUES (" & lGuildID & ", " & lPostedBy & ", " & lPostedOn & ", " & lStartAt & _
				 ", " & lDuration & ", " & CInt(yEventType) & ", " & CInt(ySendAlerts) & ", " & CInt(yEventIcon) & ", " & _
				 CInt(yMembersCanAccept) & ", " & CInt(yRecurrence) & ", '"
				If yDetails Is Nothing = False Then sSQL &= MakeDBStr(BytesToString(yDetails))
				sSQL &= "', '" & MakeDBStr(BytesToString(yTitle)) & "')"
			Else
				'UPDATE
				sSQL = "UPDATE tblGuildEvent SET StartsAt = " & lStartAt & ", Duration = " & lDuration & ", EventType = " & CInt(yEventType) & _
				  ", SendAlerts = " & CInt(ySendAlerts) & ", EventIcon = " & CInt(yEventIcon) & ", MembersCanAccept = " & _
				  CInt(yMembersCanAccept) & ", Recurrence = " & CInt(yRecurrence) & ", Details = '"
				If yDetails Is Nothing = False Then sSQL &= MakeDBStr(BytesToString(yDetails))
				sSQL &= "', EventTitle = '" & MakeDBStr(BytesToString(yTitle)) & "' WHERE EventID = " & Me.EventID
			End If
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			If oComm.ExecuteNonQuery() = 0 Then
				Err.Raise(-1, "SaveObject", "No records affected when saving object!")
			End If
			If Me.EventID = -1 Then
				Dim oData As OleDb.OleDbDataReader
				oComm = Nothing
				sSQL = "SELECT MAX(EventID) FROM tblGuildEvent WHERE GuildID = " & lGuildID
				oComm = New OleDb.OleDbCommand(sSQL, goCN)
				oData = oComm.ExecuteReader(CommandBehavior.Default)
				If oData.Read Then
					Me.EventID = CInt(oData(0))
				End If
				oData.Close()
				oData = Nothing
			End If

			sSQL = "DELETE FROM tblGuildEventAcceptance WHERE EventID = " & Me.EventID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm = Nothing

			For X As Int32 = 0 To mlAcceptanceUB
				sSQL = "INSERT INTO tblGuildEventAcceptance (EventID, PlayerID, Acceptance) VALUES (" & Me.EventID & ", " & muAcceptance(X).lPlayerID & ", " & muAcceptance(X).yAcceptance & ")"
				oComm = New OleDb.OleDbCommand(sSQL, goCN)
				oComm.ExecuteNonQuery()
				oComm = Nothing
			Next X

			If oAdvancedConfig Is Nothing = False Then
				oAdvancedConfig.lEventID = Me.EventID
				oAdvancedConfig.SaveObject()
			End If

			bResult = True
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object GuildEvent. Reason: " & Err.Description)
		Finally
			oComm = Nothing
		End Try
		Return bResult
	End Function

	Public Sub AddOrSetAttachment(ByRef oAttach As SketchPad)

		Dim lIdx As Int32 = -1
		For X As Int32 = 0 To AttachmentUB
			If Attachments(X) Is Nothing Then
				If lIdx <> -1 Then lIdx = X
			ElseIf Attachments(X).lID = oAttach.lID Then
				Attachments(X) = oAttach
				Return
			End If
		Next X
		If lIdx = -1 Then
			ReDim Preserve Attachments(AttachmentUB + 1)
			AttachmentUB += 1
			lIdx = AttachmentUB
		End If
		Attachments(lIdx) = oAttach
	End Sub

	Public Sub DeleteAttachment(ByVal lID As Int32)
		For X As Int32 = 0 To AttachmentUB
			If Attachments(X) Is Nothing = False Then
				If Attachments(X).lID = lID Then
					Attachments(X).DeleteMe()
					Return
				End If
			End If
		Next X
	End Sub

	Public Sub SetPlayerAcceptance(ByVal lPlayerID As Int32, ByVal yAcceptance As Byte)
		If muAcceptance Is Nothing Then ReDim muAcceptance(-1)
		For X As Int32 = 0 To mlAcceptanceUB
			If muAcceptance(X).lPlayerID = lPlayerID Then
				muAcceptance(X).yAcceptance = yAcceptance
				Return
			End If
		Next X

		SyncLock muAcceptance
			ReDim Preserve muAcceptance(mlAcceptanceUB + 1)
			mlAcceptanceUB += 1
			muAcceptance(mlAcceptanceUB).lPlayerID = lPlayerID
			muAcceptance(mlAcceptanceUB).yAcceptance = yAcceptance
		End SyncLock
	End Sub

	Public Sub DeleteMe()
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		Try

			If oAdvancedConfig Is Nothing = False Then oAdvancedConfig.DeleteMe()

			sSQL = "DELETE FROM tblSketchPadItem WHERE SketchPadID IN (Select SketchPadID FROM tblSketchPad WHERE EventID = " & Me.EventID & ")"
			oComm = Nothing
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm = Nothing

			sSQL = "DELETE FROM tblSketchPad WHERE EventID = " & Me.EventID
			oComm = Nothing
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm = Nothing

			sSQL = "DELETE FROM tblGuildEventAcceptance WHERE EventID = " & Me.EventID
			oComm = Nothing
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm = Nothing

			sSQL = "DELETE FROM tblGuildEvent WHERE EventID = " & Me.EventID
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

	Public Function GetAttachment(ByVal lID As Int32) As SketchPad
		For X As Int32 = 0 To AttachmentUB
			If Attachments(X) Is Nothing = False AndAlso Attachments(X).lID = lID Then Return Attachments(X)
		Next X
		Return Nothing
	End Function
End Class

