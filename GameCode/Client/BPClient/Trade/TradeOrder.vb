Option Strict On

Public MustInherit Class TradeOrder
    Public lTradePostID As Int32
    Public lObjectID As Int32           'BuyOrderID for Buy Orders
    Public iObjTypeID As Int16 = -1         'not used for buy orders

    Public yRouteType As Byte           '0 = planetary, 1 = intrasystem, 2 = system to system

    Public lDeliveryTime As Int32 = -1  'seconds to deliver
    Public lSellerID As Int32 = -1      'playerID of the seller

    Public blQty As Int64
    Public blPrice As Int64

    Protected mbDetailsRequested As Boolean = False
    Public MustOverride Sub RequestDetails(ByVal lFromTPID As Int32)
    Public MustOverride Sub SetFromMsg(ByRef yData() As Byte, ByVal lPos As Int32)
    Public MustOverride Sub SetFromDetailMsg(ByRef yData() As Byte, ByVal lPos As Int32)
    Public MustOverride Function GetCurrentListText() As String
End Class