Public Class ProductionCostItem
    Public PCM_ID As Int32          'unique ID
    Public PC_ID As Int32 = -1           'FK to ProductionCost
    Public ItemID As Int32 = -1         'FK to object ID
    Public ItemTypeID As Int16 = -1     'object type id

    Public QuantityNeeded As Int32  'number of this item required

    'helpers...
    Public oProdCost As ProductionCost

    Public Sub New()
        '
    End Sub

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If PCM_ID < 1 Then
                'INSERT
                sSQL = "INSERT INTO tblProductionCostItem (PC_ID, ItemID, ItemTypeID, Quantity) VALUES (" & _
                  Me.PC_ID & ", " & Me.ItemID & ", " & Me.ItemTypeID & ", " & Me.QuantityNeeded & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblProductionCostItem SET PC_ID = " & Me.PC_ID & _
                  ", ItemID = " & Me.ItemID & ", ItemTypeID = " & Me.ItemTypeID & ",Quantity = " & Me.QuantityNeeded & _
                  " WHERE PCM_ID = " & Me.PCM_ID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If PCM_ID < 1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(PCM_ID) FROM tblProductionCostItem WHERE PC_ID = " & Me.PC_ID & " AND ItemID = " & Me.ItemID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    PCM_ID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            oComm = Nothing

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & PCM_ID & " of type ProductionCostItem. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
	End Function
	Public Function GetSaveObjectText() As String
		Dim sSQL As String

		Try
			If PCM_ID < 1 Then
				'INSERT
				sSQL = "INSERT INTO tblProductionCostItem (PC_ID, ItemID, ItemTypeID, Quantity) VALUES (" & _
				  Me.PC_ID & ", " & Me.ItemID & ", " & Me.ItemTypeID & ", " & Me.QuantityNeeded & ")"
			Else
				'UPDATE
				sSQL = "UPDATE tblProductionCostItem SET PC_ID = " & Me.PC_ID & _
				  ", ItemID = " & Me.ItemID & ", ItemTypeID = " & Me.ItemTypeID & ",Quantity = " & Me.QuantityNeeded & _
				  " WHERE PCM_ID = " & Me.PCM_ID
			End If
			Return sSQL
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & PCM_ID & " of type ProductionCostItem. Reason: " & Err.Description)
		End Try
		Return ""
	End Function
End Class