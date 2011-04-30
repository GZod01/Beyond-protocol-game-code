Public Class MineralPropertyValue
    Public MineralPropertyValueID As Int32 = -1  'PK

    Public lPropertyID As Int32             'for quick lookup
    Public oProperty As MineralProperty     'reference to the mineral property
    Public lValue As Int32                  'value

    Public lParentID As Int32               'for quick lookup of Parent Mineral/PlayerMineral

    ''' <summary>
    ''' Saves the Mineral Property to the appropriate table
    ''' </summary>
    ''' <param name="lPropValType">Indicates table to save to: 0 is tblMineralPropertyValue, 1 is tblPlayerMineralPropertyValue</param>
    ''' <returns>Boolean</returns>
    ''' <remarks></remarks>
    Public Function SaveObject(ByVal lPropValType As Int32) As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Dim sTableName As String
        Dim sKeyName As String
        Dim sParentName As String

        If lPropValType = 0 Then
            sTableName = "tblMineralPropertyValue"
            sParentName = "MineralID"
            sKeyName = "MineralPropertyValueID"
        Else
            sTableName = "tblPlayerMineralPropertyValue"
            sParentName = "PlayerMineralID"
            sKeyName = "PlayerMineralPropertyValueID"
        End If

        Try
            If MineralPropertyValueID = -1 Then
                'INSERT
                sSQL = "INSERT INTO " & sTableName & " (" & sParentName & ", MineralPropertyID, PropertyValue) VALUES (" & _
                  lParentID & ", " & lPropertyID & ", " & lValue & ")"
            Else
                'UPDATE
                sSQL = "UPDATE " & sTableName & " SET " & sParentName & " = " & lParentID & _
                  ", MineralPropertyID = " & lPropertyID & ", PropertyValue = " & lValue & " WHERE " & _
                  sKeyName & " = " & MineralPropertyValueID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If MineralPropertyValueID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(" & sKeyName & ") FROM " & sTableName & " WHERE " & sParentName & " = " & lParentID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    MineralPropertyValueID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & MineralPropertyValueID & " of MineralPropertyValue. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
	End Function

	Public Function GetSaveObjectText(ByVal lPropValType As Int32) As String
		Dim sSQL As String
		Dim sTableName As String
		Dim sKeyName As String
		Dim sParentName As String

		If lPropValType = 0 Then
			sTableName = "tblMineralPropertyValue"
			sParentName = "MineralID"
			sKeyName = "MineralPropertyValueID"
		Else
			sTableName = "tblPlayerMineralPropertyValue"
			sParentName = "PlayerMineralID"
			sKeyName = "PlayerMineralPropertyValueID"
		End If

		Try
			If MineralPropertyValueID = -1 Then
				'INSERT
				sSQL = "INSERT INTO " & sTableName & " (" & sParentName & ", MineralPropertyID, PropertyValue) VALUES (" & _
				  lParentID & ", " & lPropertyID & ", " & lValue & ")"
			Else
				'UPDATE
				sSQL = "UPDATE " & sTableName & " SET " & sParentName & " = " & lParentID & _
				  ", MineralPropertyID = " & lPropertyID & ", PropertyValue = " & lValue & " WHERE " & _
				  sKeyName & " = " & MineralPropertyValueID
			End If
			Return sSQL
		Catch
			LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & MineralPropertyValueID & " of MineralPropertyValue. Reason: " & Err.Description)
 
		End Try
		Return ""
	End Function

End Class
