Option Strict On

Public Class TransportRoute
    Public oTransport As Transport = Nothing
    Public OrderNum As Byte = 0         'Order for which this waypoint is to be executed

    Public DestinationID As Int32
    Public DestinationTypeID As Int16

    Public WaypointFlags As Byte = 0    'currently unused

    Public oActions() As TransportRouteAction
    Public lActionUB As Int32 = -1
End Class
