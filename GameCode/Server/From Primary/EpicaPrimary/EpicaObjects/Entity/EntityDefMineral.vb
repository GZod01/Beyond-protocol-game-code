Public Class EntityDefMineral
    Public oEntityDef As Epica_Entity_Def
    Public iEntityDefTypeID As Int16
    Public oMineral As Mineral
    Public lQuantity As Int32

    Public Function SaveObject() As Boolean
        Dim oComm As OleDb.OleDbCommand = Nothing
        Dim bResult As Boolean = False
        Try
            Dim sSQL As String = "INSERT INTO tblEntityDefMineral (EntityDefID, EntityDefTypeID, MineralID, Quantity) VALUES (" & _
                oEntityDef.ObjectID & ", " & oEntityDef.ObjTypeID & ", " & oMineral.ObjectID & ", " & lQuantity & ")"
            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            oComm.ExecuteNonQuery()

            bResult = True
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "EntityDefMineral.SaveObject()")
        Finally
            If oComm Is Nothing = False Then oComm.Dispose()
            oComm = Nothing
        End Try
        Return bResult
    End Function
End Class
