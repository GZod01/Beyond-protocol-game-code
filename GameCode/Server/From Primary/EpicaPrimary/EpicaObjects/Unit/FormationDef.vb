Option Strict On

Public Enum CriteriaType As Byte
	eHullSize = 0
	eWeaponSlots
	eMaxRadarRange
	eMostFrontArmorHP
	eMostShieldHP
	eManeuver
	eSpeed
	eCombatRating
	eCargoBayCap
	eHangarCap
	eEntityName
End Enum

Public Class FormationDef
	Public Const ml_GRID_SIZE_WH As Int32 = 25
	Public FormationID As Int32 = -1
	Public mlSlots(,) As Int32
	Public yName(19) As Byte
	Public yDefault As Byte
	Public yCriteria As CriteriaType
	Public lOwnerID As Int32
	Public lCellSize As Int32

    Public Function FillFromMsg(ByRef yData() As Byte, ByVal bFromOperator As Boolean) As Boolean
        Dim lPos As Int32 = 2   'for msgcode

        FormationID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        'Dim bIsDefault As Boolean = yDefault <> 0
        yDefault = yData(lPos) : lPos += 1
        'If bIsDefault = True AndAlso yDefault = 0 Then
        '    yDefault = 1
        'ElseIf bIsDefault = False AndAlso yDefault <> 0 Then
        '    Try
        '        For X As Int32 = 0 To glFormationDefUB
        '            If glFormationDefIdx(X) <> -1 Then
        '                If goFormationDefs(X).lOwnerID = Me.lOwnerID AndAlso goFormationDefs(X).yDefault <> 0 AndAlso goFormationDefs(X).FormationID <> Me.FormationID Then goFormationDefs(X).yDefault = 0
        '            End If
        '        Next X
        '    Catch
        '    End Try
        'End If

        yCriteria = CType(yData(lPos), CriteriaType) : lPos += 1

        ReDim yName(19)
        Array.Copy(yData, lPos, yName, 0, 20) : lPos += 20

        If bFromOperator = True Then
            lOwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        End If

        lCellSize = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        If lCnt < 2 Then Return False

        ReDim mlSlots(ml_GRID_SIZE_WH - 1, ml_GRID_SIZE_WH - 1)

        For lIdx As Int32 = 0 To lCnt - 1
            Dim lSlotIdx As Int32 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
            Dim iSlotVal As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            Dim lSlotY As Int32 = lSlotIdx \ ml_GRID_SIZE_WH
            Dim lSlotX As Int32 = lSlotIdx - (lSlotY * ml_GRID_SIZE_WH)

            If lSlotX < ml_GRID_SIZE_WH AndAlso lSlotY < ml_GRID_SIZE_WH AndAlso lSlotX > -1 AndAlso lSlotY > -1 Then
                mlSlots(lSlotX, lSlotY) = iSlotVal
            End If
        Next lIdx

        Return True
    End Function

	Public Function SaveObject() As Boolean
		Dim bResult As Boolean = False
		Dim sSQL As String
		Dim oComm As OleDb.OleDbCommand

		'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

		Try
			If FormationID < 1 Then
				'INSERT
				sSQL = "INSERT INTO tblFormationDef (FormationName, yDefault, yCriteria, OwnerID, CellSize) VALUES ('" & _
				  MakeDBStr(BytesToString(yName)) & "', " & yDefault & ", " & CByte(yCriteria) & ", " & lOwnerID & "," & lCellSize & ")"
			Else
				'UPDATE
				sSQL = "UPDATE tblFormationDef SET FormationName = '" & MakeDBStr(BytesToString(yName)) & "', yDefault = " & _
				  yDefault & ", yCriteria = " & yCriteria & ", OwnerID = " & lOwnerID & ", CellSize = " & lCellSize & " WHERE FormationDefID = " & Me.FormationID
			End If
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			If oComm.ExecuteNonQuery() = 0 Then
				Err.Raise(-1, "SaveObject", "No records affected when saving object!")
			End If
			If FormationID < 1 Then
				Dim oData As OleDb.OleDbDataReader
				oComm = Nothing
                sSQL = "SELECT MAX(FormationDefID) FROM tblFormationDef WHERE FormationName = '" & MakeDBStr(BytesToString(yName)) & "'"
				oComm = New OleDb.OleDbCommand(sSQL, goCN)
				oData = oComm.ExecuteReader(CommandBehavior.Default)
				If oData.Read Then
					FormationID = CInt(oData(0))
				End If
				oData.Close()
				oData = Nothing
			End If
			oComm = Nothing

			'now, our slots
			oComm = New OleDb.OleDbCommand("DELETE FROM tblFormationDefSlot WHERE FormationDefID = " & Me.FormationID, goCN)
			oComm.ExecuteNonQuery()
			oComm = Nothing

			For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
				For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
					If mlSlots(X, Y) > 0 Then
						sSQL = "INSERT INTO tblFormationDefSlot (FormationDefID, SlotX, SlotY, SlotValue) VALUES (" & _
						  Me.FormationID & ", " & X & ", " & Y & ", " & mlSlots(X, Y) & ")"
						oComm = New OleDb.OleDbCommand(sSQL, goCN)
						If oComm.ExecuteNonQuery() = 0 Then
							LogEvent(LogEventType.CriticalError, "Formation Save Slot did not update any records!")
						End If
						oComm = Nothing
					End If
				Next Y
			Next X

			bResult = True
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & FormationID & " of formationdef. Reason: " & Err.Description)
		Finally
			oComm = Nothing
		End Try
		Return bResult
	End Function

	Private Function GetSlotPortionOfMsg() As Byte()
		Dim lCnt As Int32 = 0
		For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
				If mlSlots(X, Y) > 0 Then lCnt += 1
			Next Y
		Next X

		Dim yResp((lCnt * 4) + 3) As Byte
		Dim lPos As Int32 = 0

		System.BitConverter.GetBytes(lCnt).CopyTo(yResp, lPos) : lPos += 4

		For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
			For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
				If mlSlots(X, Y) > 0 Then
					Dim iSlotIdx As Int16 = CShort((Y * ml_GRID_SIZE_WH) + X)
					Dim iSlotValue As Int16 = CShort(mlSlots(X, Y))

					System.BitConverter.GetBytes(iSlotIdx).CopyTo(yResp, lPos) : lPos += 2
					System.BitConverter.GetBytes(iSlotValue).CopyTo(yResp, lPos) : lPos += 2
				End If
			Next Y
		Next X

		Return yResp
	End Function

	Public Function GetAsAddMsg() As Byte()
		'Ok, let's build our msg
		Dim yGridData() As Byte = GetSlotPortionOfMsg()
		If yGridData Is Nothing Then Return Nothing

		Dim yMsg(35 + yGridData.Length) As Byte
		Dim lPos As Int32 = 0

		System.BitConverter.GetBytes(GlobalMessageCode.eAddFormation).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(FormationID).CopyTo(yMsg, lPos) : lPos += 4

		yMsg(lPos) = yDefault : lPos += 1
		yMsg(lPos) = CByte(yCriteria) : lPos += 1
		yName.CopyTo(yMsg, lPos) : lPos += 20

		System.BitConverter.GetBytes(lOwnerID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lCellSize).CopyTo(yMsg, lPos) : lPos += 4

		yGridData.CopyTo(yMsg, lPos) : lPos += yGridData.Length

		Return yMsg
	End Function

	Public Sub DeleteMe()
		Try
			If Me.FormationID < 1 Then Return
			Dim sSQL As String = "DELETE FROM tblFormationDefSlot WHERE FormationDefID = " & Me.FormationID
			Dim oComm As OleDb.OleDbCommand = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm = Nothing
			sSQL = "DELETE FROM tblFormationDef WHERE FormationDefID = " & Me.FormationID
			oComm = New OleDb.OleDbCommand(sSQL, goCN)
			oComm.ExecuteNonQuery()
			oComm = Nothing
		Catch ex As Exception
			LogEvent(LogEventType.CriticalError, "Delete Formation: " & ex.Message)
		End Try
	End Sub
End Class