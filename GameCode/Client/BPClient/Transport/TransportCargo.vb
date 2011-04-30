Option Strict On

Public Class TransportCargo
    Public oParentTransport As Transport
    Public CargoID As Int32
    Public CargoTypeID As Int16
    Public OwnerID As Int32
    Public Quantity As Int32

    Private msDisplay As String = ""
    Private mlDisplayID As Int32 = -1
    Private miDisplayTypeID As Int16 = -1
    Private mlDisplayQty As Int32 = -1
    Public Function GetDisplay() As String
        If msDisplay = "" OrElse mlDisplayID <> CargoID OrElse miDisplayTypeID <> CargoTypeID OrElse mlDisplayQty <> Quantity Then
            Dim sName As String = ""
            Dim bFound As Boolean = False

            If miDisplayTypeID = ObjectType.eMineral Then
                sName = "Unknown Mineral"
                For X As Int32 = 0 To glMineralUB
                    If glMineralIdx(X) > -1 Then
                        If goMinerals(X).ObjectID = CargoID Then
                            sName = goMinerals(X).MineralName
                            bFound = True
                            Exit For
                        End If
                    End If
                Next X
            Else
                sName = GetCacheObjectValueCheckUnknowns(CargoID, CargoTypeID, bFound)
                If sName = "Unknown" Then bFound = False Else bFound = True
            End If
            If bFound = False Then
                mlDisplayID = -1
            Else
                msDisplay = sName
                mlDisplayID = CargoID
                miDisplayTypeID = CargoTypeID
                mlDisplayQty = Quantity
            End If
            msDisplay = sName.PadRight(21, " "c) & Quantity.ToString("#,##0").PadLeft(8, " "c)
        End If
        Return msDisplay
    End Function
End Class
