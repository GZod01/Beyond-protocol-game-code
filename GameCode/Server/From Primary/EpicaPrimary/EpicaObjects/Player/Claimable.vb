Option Strict On

Public Enum eyClaimFlags As Byte
    eClaimed = 1
End Enum

Public Class Claimable
    Public lID As Int32
    Public iTypeID As Int16
    Public lOfferCode As Int32
    Public yItemName(19) As Byte
    Public lPlayerID As Int32
    Public yClaimFlag As Byte

    Public blQuantity As Int64      'not sent down, used in research production

    Public Function SaveObject() As Boolean
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim bResult As Boolean = False
        Try
            Dim sSQL As String = "UPDATE tblClaimable SET ClaimFlag = " & yClaimFlag & ", ClaimQuantity = " & blQuantity & " WHERE ObjectID = " & lID & " AND ObjTypeID = " & iTypeID & _
                " AND OfferCode = " & lOfferCode & " AND PlayerID = " & lPlayerID
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                oComm.Dispose()
                oComm = Nothing
                sSQL = "INSERT INTO tblClaimable (ObjectID, ObjTypeID, OfferCode, PlayerID, ItemName, ClaimFlag, ClaimQuantity) VALUES (" & lID & ", " & iTypeID & _
                    ", " & lOfferCode & ", " & lPlayerID & ", '" & BytesToString(yItemName) & "', " & yClaimFlag & ", " & blQuantity & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then Throw New Exception("No records saved!")
            End If
            bResult = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "Claimable.SaveObject: " & Me.lID & "," & Me.iTypeID & "," & Me.lOfferCode & "," & Me.yClaimFlag & "," & Me.lPlayerID & ". Reason: " & ex.Message)
        Finally
            If oComm Is Nothing = False Then oComm.Dispose()
            oComm = Nothing
        End Try
        Return bResult
    End Function
End Class
