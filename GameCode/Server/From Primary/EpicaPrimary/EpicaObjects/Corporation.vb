Public Class Corporation
    Inherits Epica_GUID

    Public CorporationName(19) As Byte
    Public HeadQuarters As Facility         'facility that is the HQ of the corporation
    Public CurrentStockValue As Int32
    Public MinimumValue As Int32
    Public MaximumValue As Int32

    Private mySendString() As Byte
    
    Public Function GetObjAsString() As Byte()
        'here we will return the entire object as a string
        If mbStringReady = False Then
            ReDim mySendString(41)  '0 to 41 = 42 bytes

            GetGUIDAsString.CopyTo(mySendString, 0)
            CorporationName.CopyTo(mySendString, 6)
            If HeadQuarters Is Nothing = False Then
                System.BitConverter.GetBytes(HeadQuarters.ObjectID).CopyTo(mySendString, 26)
            Else : System.BitConverter.GetBytes(CInt(-1)).CopyTo(mySendString, 26)
            End If
            System.BitConverter.GetBytes(CurrentStockValue).CopyTo(mySendString, 30)
            System.BitConverter.GetBytes(MinimumValue).CopyTo(mySendString, 34)
            System.BitConverter.GetBytes(MaximumValue).CopyTo(mySendString, 38)
            mbStringReady = True
        End If
        Return mySendString
    End Function

    Public Function GetHistory(ByVal lCycleCnt As Int32, ByVal lInterval As Int32) As String
        'returns a msg-ready string to get the history of a corporation's stock values (stored in
        '  CorporationStock_Arch).
        'lCycleCnt is how far to check back
        'lInterval is the interval to use (how many cycles per interval...)
        '  for example, if a player wants to look at the monthly values over the past year...
        '  and cycles were 1 day (realtime), then lCycleCnt = 356, and lInterval = 30

        Return ""
    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblCorporation (CorporationName, HQ_ID, CurrentStockValue, MinimumValue, " & _
                  "MaximumValue) VALUES ('" & MakeDBStr(BytesToString(CorporationName)) & ", " & HeadQuarters.ObjectID & ", " & _
                  CurrentStockValue & ", " & MinimumValue & ", " & MaximumValue & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblCorporation SET CorporationName = '" & MakeDBStr(BytesToString(CorporationName)) & "', HQ_ID = " & _
                  HeadQuarters.ObjectID & ", CurrentStockValue = " & CurrentStockValue & ", MinimumValue = " & _
                  MinimumValue & ", MaximumValue = " & MaximumValue & " WHERE CorporationID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(CorporationID) FROM tblCorporation WHERE CorporationName = '" & MakeDBStr(BytesToString(CorporationName)) & "'"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                oData = oComm.ExecuteReader(CommandBehavior.Default)
                If oData.Read Then
                    ObjectID = CInt(oData(0))
                End If
                oData.Close()
                oData = Nothing
            End If
            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object " & ObjectID & " of type " & ObjTypeID & ". Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function
End Class
