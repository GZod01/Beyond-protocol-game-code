'Public Class MissionGoal
'    Public oMission As Mission
'    Public oGoal As Goal
'    Public OrderNumber As Byte
'    Public StartCycle As Int32
'    Public EndCycle As Int32
'    Public PointsAccumulated As Short
'    Public Modifier As Short
'    Public ExtraAttrs As Byte

'    Private mbStringReady As Boolean = False
'    Private mySendString() As Byte
'    Public Sub DataChanged()    'call this sub when changing properties of this object
'        mbStringReady = False
'    End Sub

'    Public Function GetObjAsString() As Byte()
'        'here we will return the entire object as a string
'        If mbStringReady = False Then
'            ReDim mySendString(21)  '0 to 21 = 22 bytes

'            System.BitConverter.GetBytes(oMission.ObjectID).CopyTo(mySendString, 0)
'            System.BitConverter.GetBytes(oGoal.ObjectID).CopyTo(mySendString, 4)
'            mySendString(8) = OrderNumber
'            System.BitConverter.GetBytes(StartCycle).CopyTo(mySendString, 9)
'            System.BitConverter.GetBytes(EndCycle).CopyTo(mySendString, 13)
'            System.BitConverter.GetBytes(PointsAccumulated).CopyTo(mySendString, 17)
'            System.BitConverter.GetBytes(Modifier).CopyTo(mySendString, 19)
'            mySendString(21) = ExtraAttrs
'            mbStringReady = True
'        End If
'        Return mySendString
'    End Function

'    Public Function SaveObject() As Boolean
'        Dim bResult As Boolean = False
'        Dim sSQL As String
'        Dim oComm As OleDb.OleDbCommand

'        'if the object's OBJECT ID is -1, then we know that this is the first time saving this object!

'        Try
'            'TODO: This may be bad if we are changing the order number of the mission goal...
'            sSQL = "DELETE FROM tblMissionGoal WHERE MissionID = " & oMission.ObjectID & " AND GoalID = " & _
'              oGoal.ObjectID & " AND OrderNumber = " & OrderNumber
'            oComm = New OleDb.OleDbCommand(sSQL, goCN)
'            oComm.ExecuteNonQuery()
'            oComm = Nothing

'            sSQL = "INSERT INTO tblMissionGoal (MissionID, GoalID, OrderNumber, StartCycle, EndCycle, " & _
'              "PointsAccumulated, Modifier, ExtraAttrs) VALUES (" & oMission.ObjectID & ", " & oGoal.ObjectID & _
'              ", " & OrderNumber & ", " & StartCycle & ", " & EndCycle & ", " & PointsAccumulated & ", " & _
'              Modifier & ", " & ExtraAttrs & ")"
'            If oComm.ExecuteNonQuery() = 0 Then
'                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
'            End If
'            bResult = True
'        Catch
'            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object MissionGoal. Reason: " & Err.Description)
'        Finally
'            oComm = Nothing
'        End Try
'        Return bResult
'    End Function
'End Class
