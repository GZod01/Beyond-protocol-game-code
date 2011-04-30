Option Strict On

Public Class TransportRouteAction
    Public Enum TransportRouteActionType As Byte
        eUnload = 1
        eLoad = 2
        eQty = 4            'if this is set, extended3 is treated as an absolute quantity. -1 indicates all available.
        ePercentage = 8     'if this is set, extended3 is treated as a percentage (0 to 100). -1 or anything outside 0 to 100 indicates all available.
    End Enum

    Public oParentRoute As TransportRoute
    Public ActionOrderNum As Byte
    Public ActionTypeID As TransportRouteActionType

    ''' <summary>
    ''' A specific ObjectID or -1. When -1, ObjectID is not used in comparisons.
    ''' </summary>
    ''' <remarks></remarks>
    Public Extended1 As Int32
    ''' <summary>
    ''' A specific ObjectType ID.
    ''' </summary>
    ''' <remarks></remarks>
    Public Extended2 As Int16
    ''' <summary>
    ''' A quantity identifier or -1. If Quantity is considered to be ALL or 100%.
    ''' </summary>
    ''' <remarks></remarks>
    Public Extended3 As Int32

    Public Function GetDisplay() As String

        Dim sName As String = ""
        Select Case Extended2
            Case ObjectType.eColonists
                sName = "Colonists"
            Case ObjectType.eEnlisted
                sName = "Enlisted"
            Case ObjectType.eOfficers
                sName = "Officers"
            Case ObjectType.eMineral
                If Extended1 < 1 Then
                    sName = "Any/All Minerals"
                Else
                    sName = "Unknown Mineral"
                    For X As Int32 = 0 To glMineralUB
                        If glMineralIdx(X) = Extended1 Then
                            Dim oMineral As Mineral = goMinerals(X)
                            If oMineral Is Nothing = False Then
                                sName = oMineral.MineralName
                            End If
                            Exit For
                        End If
                    Next X
                End If
            Case ObjectType.eComponentCache
                sName = "Any/All Components"
            Case -1
                sName = "Any/All Cargo"
            Case Else
                If Extended1 < 1 Then
                    sName = "Any/All " & Base_Tech.GetComponentTypeName(Extended2)
                Else
                    Dim oTech As Base_Tech = goCurrentPlayer.GetTech(Extended1, Extended2)
                    If oTech Is Nothing = False Then sName = oTech.GetComponentName()
                End If
        End Select

        Dim sFinal As String = ""
        If (ActionTypeID And TransportRouteActionType.eLoad) <> 0 Then
            sFinal = "Load "
        Else
            sFinal = "Unload "
        End If
        If (ActionTypeID And TransportRouteActionType.eQty) <> 0 Then
            sFinal &= Extended3.ToString()
        Else
            If Extended3 < 1 OrElse Extended3 > 100 Then
                sFinal &= "100%"
            Else
                sFinal &= Extended3.ToString & "%"
            End If
        End If
        sFinal &= " " & sName

        Return sFinal
    End Function
End Class
