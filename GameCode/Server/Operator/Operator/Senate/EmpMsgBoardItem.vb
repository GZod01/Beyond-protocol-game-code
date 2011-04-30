Option Strict On

Public Class EmpMsgBoardItem
    Public lPosterID As Int32
    Public lPostedOn As Int32       'GMT Date Time as Int32
    Public yMsg() As Byte

    Public Function SaveObject() As Boolean
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim bResult As Boolean = False
        Try
            Dim sSQL As String = "UPDATE tblChamberMsgBoardItem SET MsgBody = '" & MakeDBStr(BytesToString(yMsg)) & _
                      "' WHERE PosterID = " & lPosterID & " AND PostedOn = " & lPostedOn
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                sSQL = "INSERT tblChamberMsgBoardItem (PosterID, PostedOn, MsgBody) VALUES (" & lPosterID & ", " & _
                    lPostedOn & ", '" & MakeDBStr(BytesToString(yMsg)) & "')"
                oComm.Dispose()
                oComm = Nothing
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Throw New Exception("No records effected with save!")
                End If
            End If
            bResult = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "EmpMsgBoardItem.SaveObject error: " & ex.Message)
        End Try
        If oComm Is Nothing = False Then oComm.Dispose()
        oComm = Nothing

        Return bResult
    End Function
End Class
