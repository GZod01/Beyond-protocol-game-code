Public Class Galaxy
    Inherits Epica_GUID

    Public GalaxyName(19) As Byte   '20 byte string
    Public CycleNumber As Int32

	Private mySendString() As Byte

    Public Function GetObjAsString() As Byte()
        'here we will return the object populated as a a string

        If mbStringReady = False Then
            'getGuid = 6
            '20 name
            'cyclenumber = 4
            ReDim mySendString(29)      '0 to 29 = 30 bytes

            GetGUIDAsString.CopyTo(mySendString, 0)
            GalaxyName.CopyTo(mySendString, 6)
            System.BitConverter.GetBytes(CycleNumber).CopyTo(mySendString, 26)

            mbStringReady = True
        End If

        Return mySendString
    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand
        Dim oSB As New System.Text.StringBuilder()


        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

        Try
            If ObjectID = -1 Then
                'INSERT
                oSB.Append("INSERT INTO tblGalaxy (GalaxyName, CycleNumber) VALUES ('")
                oSB.Append(MakeDBStr(BytesToString(GalaxyName)))
                oSB.Append("', ")
                oSB.Append(CycleNumber)
                oSB.Append(")")
                sSQL = oSB.ToString
            Else
                'UPDATE
                oSB.Append("UPDATE tblGalaxy SET GalaxyName = '")
                oSB.Append(MakeDBStr(BytesToString(GalaxyName)))
                oSB.Append("', CycleNumber = ")
                oSB.Append(CycleNumber)
                oSB.Append(" WHERE GalaxyID = ")
                oSB.Append(ObjectID)
                sSQL = oSB.ToString
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(GalaxyID) FROM tblGalaxy WHERE GalaxyName = '" & MakeDBStr(BytesToString(GalaxyName)) & "'"
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
            oSB = Nothing
            oComm = Nothing
        End Try
        Return bResult
    End Function
End Class
