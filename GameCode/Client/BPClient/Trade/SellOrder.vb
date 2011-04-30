Option Strict On

Public Class SellOrder
    Inherits TradeOrder

	Public sItemName As String

	Public iExtTypeID As Int16

    Public ItemScore As Int32

    Public Function GetPurchaseMessage(ByVal blPurchaseQty As Int64, ByVal lBuyerTradepostID As Int32) As Byte()
        If blPurchaseQty > blQty Then
            goUILib.AddNotification("Not enough quantity from seller to purchase that amount.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return Nothing
        End If

        If blPurchaseQty * blPrice * goCurrentPlayer.GetPlayerTradeCost() > goCurrentPlayer.blCredits Then
            goUILib.AddNotification("You do not have enough credits to purchase that quantity.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return Nothing
        End If

        Dim yMsg(23) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.ePurchaseSellOrder).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lBuyerTradepostID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lTradePostID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lObjectID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(iObjTypeID).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(blPurchaseQty).CopyTo(yMsg, lPos) : lPos += 8

        Return yMsg
    End Function

    Public Overrides Function GetCurrentListText() As String
        '"Item Name            Price                Quantity"
        Dim sResult As String = sItemName.PadRight(21, " "c)
        sResult &= blPrice.ToString("#,##0").PadRight(21, " "c)
        sResult &= blQty.ToString("#,##0")
        Return sResult
    End Function

    Public Overrides Sub RequestDetails(ByVal lFromTPID As Integer)
        If mbDetailsRequested = True Then Return
        mbDetailsRequested = True

        Dim yMsg(16) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eGetOrderSpecifics).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(lTradePostID).CopyTo(yMsg, 2)
        yMsg(6) = 0     'sell order
        System.BitConverter.GetBytes(lObjectID).CopyTo(yMsg, 7)
        System.BitConverter.GetBytes(iObjTypeID).CopyTo(yMsg, 11)
        System.BitConverter.GetBytes(lFromTPID).CopyTo(yMsg, 13)
        goUILib.SendMsgToPrimary(yMsg)
    End Sub

    Public Overrides Sub SetFromDetailMsg(ByRef yData() As Byte, ByVal lPos As Integer)
        Me.lDeliveryTime = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Me.lSellerID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    End Sub

    Public Overrides Sub SetFromMsg(ByRef yData() As Byte, ByVal lPos As Integer)
        yRouteType = yData(lPos) : lPos += 1
        lTradePostID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        lObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
        blQty = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
        blPrice = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
        ItemScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		sItemName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
		iExtTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
    End Sub
End Class

'Public Class SellOrder
'    Public lTradePostID As Int32
'    Public lObjectID As Int32
'    Public iObjTypeID As Int16

'    Public sItemName As String

'    Public blPrice As Int64
'    Public blQty As Int64

'    Public ItemScore As Int32

'    Public yRouteType As Byte           '0 = planetary, 1 = intrasystem, 2 = system to system

'    Public lDeliveryTime As Int32 = -1  'seconds to deliver
'    Public lSellerID As Int32 = -1      'playerID of the seller

'    Private mbDetailsRequested As Boolean = False
'    Public Sub RequestDetails(ByVal lFromTPID As Int32)
'        If mbDetailsRequested = True Then Return
'        mbDetailsRequested = True

'        Dim yMsg(16) As Byte
'        System.BitConverter.GetBytes(EpicaMessageCode.eGetOrderSpecifics).CopyTo(yMsg, 0)
'        System.BitConverter.GetBytes(lTradePostID).CopyTo(yMsg, 2)
'        yMsg(6) = 0     'sell order
'        System.BitConverter.GetBytes(lObjectID).CopyTo(yMsg, 7)
'        System.BitConverter.GetBytes(iObjTypeID).CopyTo(yMsg, 11)
'        System.BitConverter.GetBytes(lFromTPID).CopyTo(yMsg, 13)
'        goUILib.SendMsgToPrimary(yMsg)
'    End Sub

'    Public Sub SetFromMsg(ByRef yData() As Byte, ByVal lPos As Int32)
'        yRouteType = yData(lPos) : lPos += 1
'        lTradePostID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'        lObjectID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
'        iObjTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
'        blQty = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
'        blPrice = System.BitConverter.ToInt64(yData, lPos) : lPos += 8
'        ItemScore = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

'        sItemName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
'    End Sub

'    Public Function GetCurrentListText() As String
'        '"Item Name            Price                Quantity"
'        Dim sResult As String = sItemName.PadRight(21, " "c)
'        sResult &= blPrice.ToString("#,##0").PadRight(21, " "c)
'        sResult &= blQty.ToString("#,##0")
'        Return sResult
'    End Function

'    Public Function GetPurchaseMessage(ByVal blPurchaseQty As Int64, ByVal lBuyerTradepostID As Int32) As Byte()
'        If blPurchaseQty > blQty Then
'            goUILib.AddNotification("Not enough quantity from seller to purchase that amount.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, EpicaSound.SoundUsage.eUserInterface, Nothing, Nothing)
'            Return Nothing
'        End If

'        If blPurchaseQty * blPrice * goCurrentPlayer.GetPlayerTradeCost() > goCurrentPlayer.blCredits Then
'            goUILib.AddNotification("You do not have enough credits to purchase that quantity.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, EpicaSound.SoundUsage.eUserInterface, Nothing, Nothing)
'            Return Nothing
'        End If

'        Dim yMsg(23) As Byte
'        Dim lPos As Int32 = 0
'        System.BitConverter.GetBytes(EpicaMessageCode.ePurchaseSellOrder).CopyTo(yMsg, lPos) : lPos += 2
'        System.BitConverter.GetBytes(lBuyerTradepostID).CopyTo(yMsg, lPos) : lPos += 4
'        System.BitConverter.GetBytes(lTradePostID).CopyTo(yMsg, lPos) : lPos += 4
'        System.BitConverter.GetBytes(lObjectID).CopyTo(yMsg, lPos) : lPos += 4
'        System.BitConverter.GetBytes(iObjTypeID).CopyTo(yMsg, lPos) : lPos += 2
'        System.BitConverter.GetBytes(blPurchaseQty).CopyTo(yMsg, lPos) : lPos += 8

'        Return yMsg
'    End Function
'End Class