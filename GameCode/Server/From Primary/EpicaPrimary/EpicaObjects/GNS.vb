Public Class GNS
    Inherits Epica_GUID

    Public GNS_Name(19) As Byte
    Public OwningCorporation As Corporation
    Public OwningPlayer As Player

    Private mySendString() As Byte

    Public Function GetObjAsString() As Byte()
        'here we will return the entire object as a string
        If mbStringReady = False Then
            ReDim mySendString(33)  '0 to 33 = 34 bytes

            GetGUIDAsString.CopyTo(mySendString, 0)
            GNS_Name.CopyTo(mySendString, 6)
            System.BitConverter.GetBytes(OwningCorporation.ObjectID).CopyTo(mySendString, 26)
            System.BitConverter.GetBytes(OwningPlayer.ObjectID).CopyTo(mySendString, 30)
            mbStringReady = True
        End If
        Return mySendString
    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblGNS (GNS_Name, OwningCorpID, OwnerID) VALUES ('" & MakeDBStr(BytesToString(GNS_Name)) & "', " & _
                  OwningCorporation.ObjectID & ", " & OwningPlayer.ObjectID & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblGNS SET GNS_Name = '" & MakeDBStr(BytesToString(GNS_Name)) & "', OwningCorpID = " & OwningCorporation.ObjectID & _
                  ", OwnerID = " & OwningPlayer.ObjectID & " WHERE GNS_ID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(GNS_ID) FROM tblGNS WHERE GNS_Name = '" & MakeDBStr(BytesToString(GNS_Name)) & "'"
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
