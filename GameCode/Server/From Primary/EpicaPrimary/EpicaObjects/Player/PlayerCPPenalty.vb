Option Strict On

Public Class PlayerCPPenalty
    Public oPlayer As Player
    Public oDecPlayer As Player
    Public yPenalty As Byte

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False

        Dim oComm As OleDb.OleDbCommand = Nothing
        Try

            Dim sSQL As String = "UPDATE tblPlayerCPPenalty SET Penalty = " & yPenalty & " WHERE PlayerID = " & oPlayer.ObjectID & " AND DecPlayerID = " & oDecPlayer.ObjectID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                If oComm Is Nothing = False Then oComm.Dispose()
                oComm = Nothing

                sSQL = "INSERT INTO tblPlayerCPPenalty (PlayerID, DecPlayerID, Penalty) VALUES (" & oPlayer.ObjectID & ", " & oDecPlayer.ObjectID & ", " & yPenalty & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Throw New Exception("No records saved when inserted!")
                End If
            End If

            bResult = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "PlayerCPPenalty.SaveObject: " & ex.Message)
        Finally
            If oComm Is Nothing = False Then oComm.Dispose()
            oComm = Nothing
        End Try

        Return bResult
    End Function
End Class
