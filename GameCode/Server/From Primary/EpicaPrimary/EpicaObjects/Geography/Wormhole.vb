Public Enum elWormholeFlag As int32
    eSystem1Detectable = 1      'system 1 detectable indicates that system 1 can detect it, however, system 2 detectable
    eSystem2Detectable = 2      ' should not be set until system 1 jumps through to system 2
    eSystem1Jumpable = 4
    eSystem2Jumpable = 8
    eSystem1Colonized = 16
End Enum

Public Class Wormhole
    Inherits Epica_GUID

    Public System1 As SolarSystem
    Public System2 As SolarSystem
    Public LocX1 As Int32
    Public LocY1 As Int32
    Public LocX2 As Int32
    Public LocY2 As Int32

    Public StartCycle As Int32
    Public WormholeFlags As Int32

    Private mySendString() As Byte

    Public Function GetObjAsString() As Byte()
        'here we will return the object populated as a string
        If mbStringReady = False Then
            'guid-6
            'sys1id-4
            'sys2id-4
            'locx1-4
            'locy1-4
            'locx2-4
            'locy2-4
            'start-4
            'end-4
            ReDim mySendString(37)  '0 to 37 = 38 bytes

            GetGUIDAsString.CopyTo(mySendString, 0)
            System.BitConverter.GetBytes(System1.ObjectID).CopyTo(mySendString, 6)
            System.BitConverter.GetBytes(System2.ObjectID).CopyTo(mySendString, 10)
            System.BitConverter.GetBytes(LocX1).CopyTo(mySendString, 14)
            System.BitConverter.GetBytes(LocY1).CopyTo(mySendString, 18)
            System.BitConverter.GetBytes(LocX2).CopyTo(mySendString, 22)
            System.BitConverter.GetBytes(LocY2).CopyTo(mySendString, 26)
            System.BitConverter.GetBytes(StartCycle).CopyTo(mySendString, 30)
            System.BitConverter.GetBytes(WormholeFlags).CopyTo(mySendString, 34)
            mbStringReady = True
        End If
        Return mySendString
    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!
        If Me.InMyDomain = False Then Return True

        Try
            If ObjectID = -1 Then
                'INSERT
                sSQL = "INSERT INTO tblWormhole (System1ID, System2ID, LocX1, LocY1, LocX2, LocY2, StartCycle, " & _
                  "WormholeFlag) VALUES (" & System1.ObjectID & ", " & System2.ObjectID & ", " & LocX1 & ", " & LocY1 & _
                  ", " & LocX2 & ", " & LocY2 & ", " & StartCycle & ", " & WormholeFlags & ")"
            Else
                'UPDATE
                sSQL = "UPDATE tblWormhole SET System1ID = " & System1.ObjectID & ", System2ID = " & System2.ObjectID & _
                  ", LocX1 = " & LocX1 & ", LocX2 = " & LocX2 & ", LocY1 = " & LocY1 & ", LocY2 = " & LocY2 & _
                  ", StartCycle = " & StartCycle & ", WormholeFlag = " & WormholeFlags & " WHERE WormholeID = " & ObjectID
            End If
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If
            If ObjectID = -1 Then
                Dim oData As OleDb.OleDbDataReader
                oComm = Nothing
                sSQL = "SELECT MAX(WormholeID) FROM tblWormhole"
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
