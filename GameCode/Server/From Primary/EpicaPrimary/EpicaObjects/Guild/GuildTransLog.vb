Option Strict On

Public Class GuildTransLog
    Public lGuildID As Int32
    Public lPlayerID As Int32
    Public blAmount As Int64
    Public blBalance As Int64
    Public lTransDate As Int32

    Public Shared Function CreateGuildTransLog(ByVal lGID As Int32, ByVal lPID As Int32, ByVal blAmt As Int64, ByVal blBal As Int64, ByVal yType As Byte) As GuildTransLog

        Dim blBalSheetAmt As Int64 = Math.Abs(blAmt)
        If yType = 0 Then
            'deposit
            LogEvent(LogEventType.ExtensiveLogging, "GuildDeposit: " & lPID & " deposited " & blBalSheetAmt.ToString)
        Else
            'withdraw
            LogEvent(LogEventType.ExtensiveLogging, "GuildWithdraw: " & lPID & " withdrew " & blBalSheetAmt.ToString)
            blBalSheetAmt = -blBalSheetAmt
        End If

        Dim oNew As New GuildTransLog
        With oNew
            .blAmount = blBalSheetAmt
            .blBalance = blBal
            .lGuildID = lGID
            .lPlayerID = lPID
            .lTransDate = GetDateAsNumber(Date.SpecifyKind(Now, DateTimeKind.Local).ToUniversalTime)
            If .SaveObject() = False Then Return Nothing
        End With
        Return oNew
    End Function

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim oComm As OleDb.OleDbCommand = Nothing
        Try
            Dim sSQL As String = "INSERT INTO tblGuildTransLog (GuildID, PlayerID, TransAmount, BalanceAmount, TransDate, TDMilli) VALUES ("
            sSQL &= lGuildID.ToString & ", " & lPlayerID.ToString & ", " & blAmount.ToString & ", " & blBalance.ToString & ", " & lTransDate.ToString & _
            ", " & Now.Millisecond.ToString & ")"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then Throw New Exception("No records affected!")
            bResult = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Unable to save GuildTransLog item: " & ex.Message)
        End Try
        If oComm Is Nothing = False Then oComm.Dispose()
        oComm = Nothing

        Return bResult

    End Function
End Class
