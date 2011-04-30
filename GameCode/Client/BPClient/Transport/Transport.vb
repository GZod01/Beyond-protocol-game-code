Option Strict On

Public Class Transport
    'THIS ENUM IS BITWISE
    Public Enum elTransportFlags As Byte
        eIdle = 1           'sitting on its thumb
        eEnRoute = 2        'moving from point A to point B
        eInTransfer = 4     'Transferring cargo - unit has arrived and is transferring - do not tell it to arrive again
        eBroken = 8         'Transport arrived at a destination but no colony was found, it then returned (
        ePaused = 16        'paused at the current location - pauses cannot happen mid flight
        eLoop = 32          'indicates whether to loop once at the end
    End Enum

    Public TransportID As Int32
    'Public sName As String     use frmTransportManagement.GetTransportName()

    'Is the Dest when transport is moving... is the loc when the transport is idle
    Public LocationID As Int32
    Public LocationTypeID As Int16

    Public ETA As DateTime
    Public TransFlags As Byte
    Public CurrentWaypoint As Byte

    Public UnitDefID As Int32
    Public ModelID As Int16 = 0
    Public MaxSpeed As Byte
    Public Maneuver As Byte
    Public CargoCapAvail As Int32
    Public CargoCapTotal As Int32

    Public oRoute() As TransportRoute
    Public lRouteUB As Int32 = -1

    Public oCargo() As TransportCargo
    Public lCargoUB As Int32 = -1

    Public lLastSetStatusMsg As Int32 = -1

    Private mbDetailsRequested As Boolean = False
    Private mlLastReRequest As Int32
    Public Sub RequestDetails()

        If ETA <> DateTime.MinValue AndAlso ETA < Now Then
            If glCurrentCycle - mlLastReRequest > 300 Then
                mlLastReRequest = glCurrentCycle
                mbDetailsRequested = False
            End If
        ElseIf ETA = DateTime.MinValue AndAlso (TransFlags And Transport.elTransportFlags.eEnRoute) <> 0 Then
            If glCurrentCycle - mlLastReRequest > 300 Then
                mlLastReRequest = glCurrentCycle
                mbDetailsRequested = False
            End If
        ElseIf (TransFlags And Transport.elTransportFlags.eInTransfer) <> 0 Then
            If glCurrentCycle - mlLastReRequest > 300 Then
                mlLastReRequest = glCurrentCycle
                mbDetailsRequested = False
            End If
        End If


        If mbDetailsRequested = False Then
            mbDetailsRequested = True
            'goUILib.AddNotification("Requesting for " & Me.TransportID, Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

            'request the details
            Dim yMsg(5) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestTransportDetails).copyto(yMsg, 0)
            System.BitConverter.GetBytes(TransportID).CopyTo(yMsg, 2)
            goUILib.SendMsgToPrimary(yMsg)
        End If
    End Sub

    Public Sub HandleTransportDetails(ByVal yData() As Byte)
        Dim lPos As Int32 = 6       'for msgcode and transprotid

        UnitDefID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        MaxSpeed = yData(lPos) : lPos += 1
        Maneuver = yData(lPos) : lPos += 1
        CargoCapAvail = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        CargoCapTotal = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        ModelID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        Dim lSecs As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        If lSecs < 1 Then
            ETA = DateTime.MinValue
        Else
            ETA = Now.AddSeconds(lSecs)
        End If

        LocationID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        LocationTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        TransFlags = yData(lPos) : lPos += 1

        Dim lUB As Int32 = yData(lPos) : lPos += 1
        lUB -= 1       'to make it a ub
        Dim oTemp(lUB) As TransportRoute
        For X As Int32 = 0 To lUB
            oTemp(X) = New TransportRoute
            With oTemp(X)
                .DestinationID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .DestinationTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .lActionUB = yData(lPos) : lPos += 1
                .lActionUB -= 1     'to make it a ub
                .OrderNum = CByte(X)
                .oTransport = Me
                ReDim .oActions(.lActionUB)
            End With

            For Y As Int32 = 0 To oTemp(X).lActionUB
                oTemp(X).oActions(Y) = New TransportRouteAction
                With oTemp(X).oActions(Y)
                    .oParentRoute = oTemp(X)
                    .ActionOrderNum = CByte(Y)
                    .ActionTypeID = CType(yData(lPos), TransportRouteAction.TransportRouteActionType) : lPos += 1
                    .Extended1 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                    .Extended2 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                    .Extended3 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                End With
            Next Y
        Next X
        lRouteUB = -1
        oRoute = oTemp
        lRouteUB = lUB

        'Now, our cargo
        lUB = -1
        lUB = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        lUB -= 1    'to make it a ub
        Dim oTC(lUB) As TransportCargo
        For X As Int32 = 0 To lUB
            oTC(X) = New TransportCargo()
            With oTC(X)
                .CargoID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .CargoTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
                .OwnerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .Quantity = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            End With
        Next X
        lCargoUB = -1
        oCargo = oTC
        lCargoUB = lUB
    End Sub

    Public Function GetStatusText() As String
        Dim sLoc As String = GetCacheObjectValue(LocationID, LocationTypeID)
        If (TransFlags And elTransportFlags.ePaused) <> 0 Then
            Return "Paused at " & sLoc
        End If
        If (TransFlags And elTransportFlags.eIdle) <> 0 Then
            Return "Idle at " & sLoc
        End If
        If (TransFlags And elTransportFlags.eBroken) <> 0 Then
            Return "Returning to " & sLoc
        End If
        If (TransFlags And elTransportFlags.eEnRoute) <> 0 Then
            Return "Moving to " & sLoc
        End If
        If (TransFlags And elTransportFlags.eInTransfer) <> 0 Then
            Return "Arrived at " & sLoc
        End If
        Return "Idle"
    End Function

    Public Sub RemoveCargo(ByVal lID As Int32, ByVal iTypeID As Int16, ByVal lOwnerID As Int32)
        Dim lUB As Int32 = -1
        If oCargo Is Nothing = False Then lUB = Math.Min(lCargoUB, oCargo.GetUpperBound(0))

        For X As Int32 = 0 To lUB
            If oCargo(X) Is Nothing = False Then
                If oCargo(X).CargoID = lID AndAlso oCargo(X).CargoTypeID = iTypeID AndAlso oCargo(X).OwnerID = lOwnerID Then
                    For Y As Int32 = X To lUB - 1
                        oCargo(Y) = oCargo(Y + 1)
                    Next Y
                    lCargoUB -= 1
                    Exit For
                End If
            End If
        Next X
    End Sub
End Class
