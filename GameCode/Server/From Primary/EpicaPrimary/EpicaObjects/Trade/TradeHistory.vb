Option Strict On

Public Class TradeHistory
    Public Enum TradeResult As Byte
        eCompletedSuccessfully = 0
        eUnknownFailure = 1
        eNoBuyOrderAcceptor = 2
        eBuyOrderAcceptorFailed = 3
        eBuyOrderPlaced = 4
        eBuyOrderEscrow = 5
        eBuyOrderFinished = 6
    End Enum
    Public Enum TradeEventType As Byte
        eDirectTrade = 0
        eSellOrder = 1
        eBuyOrder = 2

        eBuyer = 64
        eSeller = 128
    End Enum
    Public Enum TradeHistoryItemType As Byte
        eBuyerPurchased = 0
        eSellerPurchased = 1
        eBuyerSold = 2
        eSellerSold = 3
    End Enum
    Private Structure TradeHistoryItem
        Public yItemName() As Byte      '20 bytes
        Public blQuantity As Int64
        Public yTradeHistoryItemType As Byte
    End Structure

    'PK
    Public lPlayerID As Int32
    Public lTransactionDate As Int32
    Public lOtherPlayerID As Int32
    'END OF PK
    Public yTradeResult As TradeResult
    Public yTradeEventType As TradeEventType
    Public blTradeAmt As Int64                  'amount of credits exchanged from lPlayerID

    Public lDeliveryTime As Int32 = 0           'in seconds

    Private muItems() As TradeHistoryItem
	Private mlItemUB As Int32 = -1

    Private mySendString() As Byte = Nothing
    Public Function GetObjAsString() As Byte()
        'If mySendString Is Nothing Then
        ReDim mySendString(22 + ((mlItemUB + 1) * 29)) 'As Byte
        Dim lPos As Int32 = 0
        'System.BitConverter.GetBytes(lPlayerID).CopyTo(mySendString, lPos) : lPos += 4

        Dim lSeconds As Int32 = CInt(Now.Subtract(GetDateFromNumber(lTransactionDate)).TotalSeconds)
        System.BitConverter.GetBytes(lSeconds).CopyTo(mySendString, lPos) : lPos += 4

        System.BitConverter.GetBytes(lOtherPlayerID).CopyTo(mySendString, lPos) : lPos += 4
        mySendString(lPos) = yTradeResult : lPos += 1
        mySendString(lPos) = yTradeEventType : lPos += 1
        System.BitConverter.GetBytes(blTradeAmt).CopyTo(mySendString, lPos) : lPos += 8
        System.BitConverter.GetBytes(lDeliveryTime).CopyTo(mySendString, lPos) : lPos += 4

        mySendString(lPos) = CByte(mlItemUB + 1) : lPos += 1
        For X As Int32 = 0 To mlItemUB
            With muItems(X)
                .yItemName.CopyTo(mySendString, lPos) : lPos += 20
                System.BitConverter.GetBytes(.blQuantity).CopyTo(mySendString, lPos) : lPos += 8
                mySendString(lPos) = .yTradeHistoryItemType : lPos += 1
            End With
        Next X
        'End If
        Return mySendString
    End Function

    Public Sub ResetMsg()
        mySendString = Nothing
    End Sub

    Public Function SaveObject() As Boolean
        Dim bResult As Boolean = False
        Dim sSQL As String
        Dim oComm As OleDb.OleDbCommand

        Try
            sSQL = "INSERT INTO tblTradeHistory (PlayerID, TransactionDate, OtherPlayerID, TradeResult, TradeEventType, TradeAmt, DeliveryTime) VALUES (" & _
              lPlayerID & ", " & lTransactionDate & ", " & lOtherPlayerID & ", " & CByte(yTradeResult) & ", " & CByte(yTradeEventType) & _
              ", " & blTradeAmt & ", " & lDeliveryTime & ")"

            oComm = New OleDb.OleDbCommand(sSQL, goCN)
            If oComm.ExecuteNonQuery() = 0 Then
                Err.Raise(-1, "SaveObject", "No records affected when saving object!")
            End If 
            oComm = Nothing

            'Save the items
            For X As Int32 = 0 To mlItemUB
                With muItems(X)
                    sSQL = "INSERT INTO tblTradeHistoryItem (PlayerID, TransactionDate, OtherPlayerID, ItemName, Quantity, " & _
                      "TradeHistoryItemType) VALUES (" & lPlayerID & ", " & lTransactionDate & ", " & lOtherPlayerID & ", '" & _
                      MakeDBStr(BytesToString(.yItemName)) & "', " & .blQuantity & ", " & CByte(.yTradeHistoryItemType) & ")"
                End With
                oComm = New OleDb.OleDbCommand(sSQL, goCN)
                If oComm.ExecuteNonQuery() = 0 Then
                    Err.Raise(-1, "SaveObject", "No records affected when saving TradeHistoryItem!")
                End If
                oComm = Nothing
            Next X

            bResult = True
        Catch
            LogEvent(GlobalVars.LogEventType.CriticalError, "Unable to save object TradeHistory. Reason: " & Err.Description)
        Finally
            oComm = Nothing
        End Try
        Return bResult
    End Function

    Public Sub AddTradeItem(ByVal yItemName() As Byte, ByVal blQty As Int64, ByVal yType As TradeHistoryItemType)
        mlItemUB += 1
        ReDim Preserve muItems(mlItemUB)
        With muItems(mlItemUB)
            ReDim .yItemName(19)
            Array.Copy(yItemName, 0, .yItemName, 0, Math.Min(20, yItemName.Length))
            'yItemName.CopyTo(.yItemName, 0)
            .blQuantity = blQty
            .yTradeHistoryItemType = yType
        End With
    End Sub

    Public Shared Sub CreateAndSaveSellOrderPurchase(ByRef oSellOrder As SellOrder, ByVal lBuyerID As Int32, ByVal blQty As Int64, ByVal blBuyerCost As Int64, ByVal lTimeToDeliver As Int32)
        Dim oNew As New TradeHistory()

        'create the trade history of the seller
        With oNew
            .lPlayerID = oSellOrder.TradePost.Owner.ObjectID
            .lTransactionDate = GetDateAsNumber(Now)
            .lOtherPlayerID = lBuyerID
            .blTradeAmt = blQty * oSellOrder.blPrice
            .yTradeEventType = TradeEventType.eSeller Or TradeEventType.eSellOrder
            .yTradeResult = TradeResult.eCompletedSuccessfully
            .lDeliveryTime = lTimeToDeliver
            .AddTradeItem(oSellOrder.yItemName, -blQty, TradeHistoryItemType.eSellerSold)
        End With
        oNew.SaveObject()
        oSellOrder.TradePost.Owner.AddTradeHistory(oNew)
        oNew = Nothing

        oNew = New TradeHistory()
        'Create the trade history of the buyer
        With oNew
            .lPlayerID = lBuyerID
            .blTradeAmt = -blBuyerCost
            If oSellOrder.iTradeTypeID = ObjectType.ePlayerIntel OrElse oSellOrder.iTradeTypeID = ObjectType.ePlayerItemIntel OrElse oSellOrder.iTradeTypeID = ObjectType.ePlayerTechKnowledge Then
                .lOtherPlayerID = -oSellOrder.TradePost.Owner.ObjectID
            Else
                .lOtherPlayerID = oSellOrder.TradePost.Owner.ObjectID
            End If
            .lTransactionDate = GetDateAsNumber(Now)
            .yTradeEventType = TradeEventType.eSellOrder Or TradeEventType.eBuyer
            .yTradeResult = TradeResult.eCompletedSuccessfully
            .lDeliveryTime = lTimeToDeliver
            .AddTradeItem(oSellOrder.yItemName, blQty, TradeHistoryItemType.eBuyerPurchased)
        End With
        oNew.SaveObject()
        Dim oPlayer As Player = GetEpicaPlayer(lBuyerID)
        If oPlayer Is Nothing = False Then oPlayer.AddTradeHistory(oNew)
        oNew = Nothing
    End Sub
    Public Shared Sub CreateAndSaveBuyOrderDelivery(ByRef oBuyOrder As BuyOrder, ByVal lDelivererID As Int32, ByRef yItemName() As Byte, ByVal lTimeToDeliver As Int32)
        Dim oNew As New TradeHistory()

        With oNew
            .lPlayerID = lDelivererID
            .blTradeAmt = oBuyOrder.blPaymentAmt + oBuyOrder.blEscrow
            .lOtherPlayerID = oBuyOrder.TradePost.Owner.ObjectID
            .lTransactionDate = GetDateAsNumber(Now)
            .yTradeEventType = TradeEventType.eBuyOrder Or TradeEventType.eSeller
            .yTradeResult = TradeResult.eBuyOrderFinished
            .AddTradeItem(yItemName, -oBuyOrder.blQuantity, TradeHistoryItemType.eSellerSold)
            .lDeliveryTime = lTimeToDeliver
        End With
        oNew.SaveObject()
        Dim oPlayer As Player = GetEpicaPlayer(lDelivererID)
        If oPlayer Is Nothing = False Then oPlayer.AddTradeHistory(oNew)
        oNew = Nothing
    End Sub
End Class
