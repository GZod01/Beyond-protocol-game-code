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
        Public sItemName As String 'yItemName() As Byte      '20 bytes
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

    Public lDeliveryTime As Int32               'in seconds

    Private muItems() As TradeHistoryItem
	Private mlItemUB As Int32 = -1

    Public Sub ClearItems()
        mlItemUB = -1
        ReDim muItems(-1)
    End Sub

    Public Sub AddTradeItem(ByVal psItemName As String, ByVal blQty As Int64, ByVal yType As TradeHistoryItemType)
        mlItemUB += 1
        ReDim Preserve muItems(mlItemUB)
        With muItems(mlItemUB)
            .sItemName = psItemName
            .blQuantity = blQty
            .yTradeHistoryItemType = yType
        End With
    End Sub

    Public Function GetListBoxText() As String
        Dim dtValue As Date = GetDateFromNumber(lTransactionDate)
        Dim sName As String '= GetCacheObjectValue(lOtherPlayerID, ObjectType.ePlayer)
        If lOtherPlayerID < 0 Then sName = "Anonymous" Else sName = GetCacheObjectValue(lOtherPlayerID, ObjectType.ePlayer)
        Dim sType As String = GetTradeTypeText(yTradeEventType)
        Dim sResult As String = GetResultText(yTradeResult)

        'DATE    PLAYER                TYPE                  RESULT                AMOUNT
        Dim sFinal As String = dtValue.Month.ToString.PadLeft(2, "0"c) & "/" & dtValue.Day.ToString.PadLeft(2, "0"c)
        sFinal = sFinal.PadRight(8, " "c)
        sFinal &= sName.PadRight(22, " "c)
        sFinal &= sType.PadRight(22, " "c)
        sFinal &= sResult.PadRight(22, " "c)

        If blTradeAmt = 0 Then
            'determine where credits might be...
            Dim yGiveType As TradeHistoryItemType
            Dim yGotType As TradeHistoryItemType
            If (yTradeEventType And TradeEventType.eBuyer) <> 0 Then
                'I'm the buyer...
                yGiveType = TradeHistoryItemType.eBuyerSold
                yGotType = TradeHistoryItemType.eBuyerPurchased
            Else
                'I'm the seller
                yGiveType = TradeHistoryItemType.eSellerSold
                yGotType = TradeHistoryItemType.eSellerPurchased
            End If

            For X As Int32 = 0 To mlItemUB
                If muItems(X).sItemName Is Nothing = False AndAlso muItems(X).sItemName.ToUpper = "CREDITS" Then
                    If muItems(X).yTradeHistoryItemType = yGotType Then
                        blTradeAmt += muItems(X).blQuantity
                    ElseIf muItems(X).yTradeHistoryItemType = yGiveType Then
                        blTradeAmt -= muItems(X).blQuantity
                    End If
                End If
            Next X

        End If
        sFinal &= blTradeAmt.ToString("#,##0").PadLeft(15, " "c)
        Return sFinal
    End Function

    Public Shared Function GetTradeTypeText(ByVal pyType As Byte) As String
        Dim sType As String = "Trade"

        Dim yTemp As Byte = pyType ' Xor (TradeEventType.eBuyer Or TradeEventType.eSeller)
        If (yTemp And TradeEventType.eBuyer) <> 0 Then yTemp = CByte(yTemp - TradeEventType.eBuyer)
        If (yTemp And TradeEventType.eSeller) <> 0 Then yTemp = CByte(yTemp - TradeEventType.eSeller)

        Select Case yTemp
            Case TradeEventType.eBuyOrder
                If (pyType And TradeEventType.eBuyer) <> 0 Then sType = "Work Order" Else sType = "Buy Order"
            Case TradeEventType.eDirectTrade
                sType = "Direct Trade"
            Case TradeEventType.eSellOrder
                If (pyType And TradeEventType.eBuyer) <> 0 Then sType = "Purchase" Else sType = "Sell Order"
        End Select
        Return sType
    End Function

    Public Shared Function GetResultText(ByVal pyResult As Byte) As String
        Dim sResult As String = "Completed"
        Select Case pyResult
            Case TradeResult.eCompletedSuccessfully
                sResult = "Completed"
            Case TradeResult.eUnknownFailure
                sResult = "Failed"
            Case TradeResult.eBuyOrderAcceptorFailed
                sResult = "Undelivered"
            Case TradeResult.eNoBuyOrderAcceptor
                sResult = "Not Accepted"
            Case TradeResult.eBuyOrderPlaced
                sResult = "Buy Order Placed"
            Case TradeResult.eBuyOrderEscrow
                sResult = "Escrow Deposited"
            Case TradeResult.eBuyOrderFinished
                sResult = "Buy Order Fulfilled"
        End Select
        Return sResult
    End Function

    Public Sub FillItemListbox(ByRef lst As UIListBox, ByVal bGiven As Boolean)
        lst.Clear()

        Dim yType As TradeHistoryItemType
        If bGiven = True Then
            If (yTradeEventType And TradeEventType.eBuyer) <> 0 Then
                'I'm the buyer...
                yType = TradeHistoryItemType.eBuyerSold
            Else
                'I'm the seller
                yType = TradeHistoryItemType.eSellerSold
            End If
        Else
            If (yTradeEventType And TradeEventType.eBuyer) <> 0 Then
                'I'm the buyer...
                yType = TradeHistoryItemType.eBuyerPurchased
            Else
                'I'm the seller
                yType = TradeHistoryItemType.eSellerPurchased
            End If
        End If

        For X As Int32 = 0 To mlItemUB
            If muItems(X).yTradeHistoryItemType = yType Then
                lst.AddItem(muItems(X).sItemName & " (" & muItems(X).blQuantity.ToString & ")", False)
            End If
        Next X
    End Sub

End Class
