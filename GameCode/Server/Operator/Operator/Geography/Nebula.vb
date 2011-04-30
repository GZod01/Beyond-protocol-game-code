Public Class Nebula
    Inherits Epica_GUID

    Public ParentGalaxy As Galaxy
    Public NebulaName(19) As Byte
    Public NebulaDesc As String     'TODO: store this as a byte array

    'TODO: if you change the Galaxy SizeX and SizeY, you will also need to change the
    '  LocX, LocY, Width and Height values in here
    Public LocX As Byte
    Public LocY As Byte
    Public Width As Byte
    Public Height As Byte
    Public Effects As Int32

    Private mySendString() As Byte
    
    Public Function GetObjAsString() As Byte()
        'here we will return the object populated as a string
        If mbStringReady = False Then
            'guid-6, parentid-4, name-20, locx-1, locy-1, width-1,height-1, effects-4
            ReDim mySendString(37)     '0 to 37 = 38

            GetGUIDAsString.CopyTo(mySendString, 0)
            System.BitConverter.GetBytes(ParentGalaxy.ObjectID).CopyTo(mySendString, 6)
            NebulaName.CopyTo(mySendString, 10)
            mySendString(30) = LocX
            mySendString(31) = LocY
            mySendString(32) = Width
            mySendString(33) = Height
            System.BitConverter.GetBytes(Effects).CopyTo(mySendString, 34)
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
                sSQL = "INSERT INTO tblNebula (GalaxyID, NebulaName, NebulaDesc, LocX, LocY, Width, Height, Effects)" & _
                  " VALUES (" & ParentGalaxy.ObjectID & ", '" & MakeDBStr(BytesToString(NebulaName)) & "','" & MakeDBStr(NebulaDesc) & "', " & LocX & _
                  ", " & LocY & ", " & Width & ", " & Height & ", " & Effects & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblNebula SET GalaxyID = " & ParentGalaxy.ObjectID & ", NebulaName='" & MakeDBStr(BytesToString(NebulaName)) & _
                "', NebulaDesc='" & MakeDBStr(NebulaDesc) & "', LocX = " & LocX & ", LocY = " & LocY & ", Width = " & Width & _
                ", Height = " & Height & ", Effects = " & Effects & " WHERE NebulaID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(NebulaID) FROM tblNebula WHERE NebulaName = '" & MakeDBStr(BytesToString(NebulaName)) & "'"
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
