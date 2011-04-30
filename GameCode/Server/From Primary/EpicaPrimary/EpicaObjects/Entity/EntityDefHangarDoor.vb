Option Strict On

Public Class EntityDefHangarDoor
    Public ED_HD_ID As Int32
    Public lEntityDefID As Int32
    Public iEntityDefTypeID As Int16
    Public DoorSize As Int32
    Public ySideArc As Byte

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ED_HD_ID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblEntityDefHangarDoor (EntityDefID, EntityDefTypeID, DoorSize, SideArc) VALUES (" & _
                   Me.lEntityDefID & ", " & Me.iEntityDefTypeID & ", " & Me.DoorSize & ", " & Me.ySideArc & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblEntityDefHangarDoor SET EntityDefID = " & Me.lEntityDefID & ", EntityDefTypeID = " & _
                    Me.iEntityDefTypeID & ", DoorSize = " & Me.DoorSize & ", SideArc = " & Me.ySideArc & " WHERE ED_HD_ID = " & Me.ED_HD_ID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ED_HD_ID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(ED_HD_ID) FROM tblEntityDefHangarDoor WHERE EntityDefID = " & Me.lEntityDefID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ED_HD_ID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save EntityDefHangarDoor. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
	End Function

	Public Function GetSaveObjectText() As String
		Dim sSQL As String

		Try
			If ED_HD_ID = -1 Then
				'INSERT
				sSQL = "INSERT INTO tblEntityDefHangarDoor (EntityDefID, EntityDefTypeID, DoorSize, SideArc) VALUES (" & _
				   Me.lEntityDefID & ", " & Me.iEntityDefTypeID & ", " & Me.DoorSize & ", " & Me.ySideArc & ")"
			Else
				'UPDATE
				sSQL = "UPDATE tblEntityDefHangarDoor SET EntityDefID = " & Me.lEntityDefID & ", EntityDefTypeID = " & _
				 Me.iEntityDefTypeID & ", DoorSize = " & Me.DoorSize & ", SideArc = " & Me.ySideArc & " WHERE ED_HD_ID = " & Me.ED_HD_ID
			End If
			Return sSQL
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save EntityDefHangarDoor. Reason: " & Err.Description)
		End Try
		Return ""
	End Function
End Class
