Option Strict On

Public Class SketchPad
	Public Enum eySketchShapes As Byte
		Box = 0
		Circle = 1
		Line = 2
		Arrow = 3
		Cross = 4
		Text = 5
	End Enum

	Public lID As Int32 = -1
	Public yName(19) As Byte
	Public lEnvirID As Int32
	Public iEnvirTypeID As Int16
	Public ViewID As Int32
	Public CameraX As Int32
	Public CameraY As Int32
	Public CameraZ As Int32
	Public CameraAtX As Int32
	Public CameraAtY As Int32
	Public CameraAtZ As Int32

	Private Structure SketchPadItem
		Public yType As Byte
		Public fPtA_X As Single
		Public fPtA_Y As Single
		Public fPtB_X As Single
		Public fPtB_Y As Single
		Public yClrVal As Byte
		Public sText As String		'for textboxes
	End Structure

	Private muItems() As SketchPadItem

	Public Sub AddSketchPadItem(ByVal yType As Byte, ByVal fPtA_X As Single, ByVal fPtA_Y As Single, ByVal fPtB_X As Single, ByVal fPtB_Y As Single, ByVal yClrVal As Byte, ByVal sText As String)
		Dim uItem As SketchPadItem
		With uItem
			.fPtA_X = fPtA_X
			.fPtA_Y = fPtA_Y
			.fPtB_X = fPtB_X
			.fPtB_Y = fPtB_Y
			.sText = sText
			.yClrVal = yClrVal
			.yType = yType
		End With
		If muItems Is Nothing Then ReDim muItems(-1)
		SyncLock muItems
			ReDim Preserve muItems(muItems.GetUpperBound(0) + 1)
			muItems(muItems.GetUpperBound(0)) = uItem
		End SyncLock
	End Sub

	Public Function SaveObject(ByVal lEventID As Int32) As Boolean
		Dim bResult As Boolean = False
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		Try
			If lID < 1 Then
				'insert
				sSQL = "INSERT INTO tblSketchPad (EventID, SketchPadName, EnvirID, EnvirTypeID, ViewID, CameraX, CameraY, CameraZ, CameraAtX, CameraAtY, CameraAtZ) VALUES ("
				sSQL &= lEventID.ToString & ", '" & MakeDBStr(BytesToString(yName)) & "', " & lEnvirID & ", " & iEnvirTypeID & ", " & _
				 ViewID & ", " & CameraX & ", " & CameraY & ", " & CameraZ & ", " & CameraAtX & ", " & CameraAtY & ", " & CameraAtZ & ")"
			Else
				'update
				sSQL = "UPDATE tblSketchPad SET SketchPadName = '" & MakeDBStr(BytesToString(yName)) & "', EnvirID = " & lEnvirID & ", EnvirTypeID = " & _
				 iEnvirTypeID & ", ViewID = " & ViewID & ", CameraX = " & CameraX & ", CameraY = " & CameraY & _
				 ", CameraZ = " & CameraZ & ", CameraAtX = " & CameraAtX & ", CameraAtY = " & CameraAtY & ", CameraAtZ = " & CameraAtZ & _
				 " WHERE SketchPadID = " & lID
			End If

			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			If oComm.ExecuteNonQuery() = 0 Then
				Err.Raise(-1, "SaveObject", "No records affected when saving object!")
			End If
			If lID = -1 Then
				Dim oData As OleDb.OleDbDataReader
				oComm = Nothing
				sSQL = "SELECT MAX(SketchPadID) FROM tblSketchPad WHERE EventID = " & lEventID
				oComm = New OleDb.OleDbCommand(sSQL, goCN)
				oData = oComm.ExecuteReader(CommandBehavior.Default)
				If oData.Read Then
					lID = CInt(oData(0))
				End If
				oData.Close()
				oData = Nothing
			End If
			oComm = Nothing

			sSQL = "DELETE FROM tblSketchPadItem WHERE SketchPadID = " & Me.lID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm = Nothing

			If muItems Is Nothing = False Then
				For X As Int32 = 0 To muItems.GetUpperBound(0)
					With muItems(X)
						sSQL = "INSERT INTO tblSketchPadItem (SketchPadID, ItemType, PtAX, PtAY, PtBX, PtBY, ClrVal, ItemText) VALUES ("
						sSQL &= Me.lID & ", " & .yType & ", " & .fPtA_X & ", " & .fPtA_Y & ", " & .fPtB_X & ", " & .fPtB_Y & ", " & _
						 .yClrVal & ", '"
						If .sText Is Nothing = False Then sSQL &= MakeDBStr(.sText)
						sSQL &= "')"

						oComm = New OleDb.OleDbCommand(sSQL, goCN)
						oComm.ExecuteNonQuery()
						oComm = Nothing
					End With
				Next X
			End If

			bResult = True
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object SketchPad. Reason: " & Err.Description)
		Finally
			oComm = Nothing
		End Try
		Return bResult
	End Function

	Public Sub DeleteMe()
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		Try

			sSQL = "DELETE FROM tblSketchPadItem WHERE SketchPadID = " & Me.lID
			oComm = Nothing
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm = Nothing

			sSQL = "DELETE FROM tblSketchPad WHERE SketchPadID = " & Me.lID
			oComm = Nothing
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm = Nothing

		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to delete SketchPad. Reason: " & Err.Description)
		Finally
			oComm = Nothing
		End Try
	End Sub
End Class
