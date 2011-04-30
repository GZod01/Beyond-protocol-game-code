Option Strict On

Public Class TransportCargo
    Public oParentTransport As Transport
    Public CargoID As Int32
    Public CargoTypeID As Int16
    Public OwnerID As Int32
    Public Quantity As Int32

    'ONLY USED WHEN THIS CARGO IS A UNIT!!!
    Private mlLastUnitSetCycle As Int32 = Int32.MinValue
    Private moUnit As Unit = Nothing
    Public ReadOnly Property oUnit() As Unit
        Get
            If CargoTypeID <> ObjectType.eUnit Then Return Nothing
            If moUnit Is Nothing OrElse mlLastUnitSetCycle = Int32.MinValue OrElse glCurrentCycle - mlLastUnitSetCycle > 300 Then
                mlLastUnitSetCycle = glCurrentCycle
                moUnit = GetEpicaUnit(CargoID)
                If moUnit Is Nothing Then CargoTypeID = -1
            End If
            Return moUnit
        End Get
    End Property

    Public Function SaveObject(ByVal bInsertOnly As Boolean) As Boolean
        Dim sSQL As String = ""
        Dim oComm As OleDb.OleDbCommand = Nothing
        Try
            If bInsertOnly = True Then
                sSQL = "INSERT INTO tblTransportCargo (TransportID, CargoID, CargoTypeID, Quantity, OwnerID, TransOwnerID) VALUES (" & oParentTransport.TransportID & ", " & _
                    CargoID & ", " & CargoTypeID & ", " & Quantity & ", " & OwnerID & ", " & oParentTransport.OwnerID & ")"
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Throw New Exception("No records affected!")
                End If
                oComm.Dispose()
                oComm = Nothing
            Else
                sSQL = "UPDATE tblTransportCargo SET Quantity = " & Quantity & " WHERE CargoID = " & CargoID & " and cargotypeid = " & CargoTypeID & _
                    " and OwnerId = " & OwnerID
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    oComm.Dispose()
                    oComm = Nothing
                    sSQL = "INSERT INTO tblTransportCargo (TransportID, CargoID, CargoTypeID, Quantity, OwnerID) VALUES (" & oParentTransport.TransportID & ", " & _
                        CargoID & ", " & CargoTypeID & ", " & Quantity & ", " & OwnerID & ")"
                    oComm = New OleDb.OleDbCommand(sSQL, goCN)
                    If oComm.ExecuteNonQuery() = 0 Then
                        Throw New Exception("No records affected!")
                    End If
                End If
                oComm.Dispose()
                oComm = Nothing
            End If
        Catch ex As Exception
            LogEvent(LogEventType.CriticalError, "TransportCargo Save Error: " & ex.Message & vbCrLf & "  SQL: " & sSQL)
            Return False
        End Try
        Return True
    End Function

    'SCENARIOS:
    '==========
    'Component
    '  CargoID = ComponentID (techid)
    '  CargoTypeID = ComponentTypeID (type of tech, objecttype.eenginetech for example)
    '  OwnerID = owner of the original design
    '
    'Mineral
    '  CargoID = MineralID of the mineral
    '  CargoTypeID = ObjectType.eMineral
    '  OwnerID = Transport.Owner
    '
    'Enlisted/Officer/Colonists
    '  CargoID = 1
    '  CargoTypeID = ObjectType.eEnlisted/eOfficer/eColonists - whichever may be the case
    '  OwnerID = Transport.Owner
    '
    'Unit
    '  CargoID = UnitID
    '  CargoTypeID = ObjectType.eUnit
    '  OwnerID = Unit.Owner
End Class
