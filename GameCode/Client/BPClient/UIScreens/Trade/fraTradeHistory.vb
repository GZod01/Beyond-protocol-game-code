Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class fraTradeHistory
    Inherits UIWindow

    Private Enum eListSortType As Byte
        eDate = 0
        ePlayer = 1
        eType = 2
        eResult = 3
        eAmount = 4

        eDescendOrder = 128
    End Enum

    Private WithEvents lstHistory As UIListBox
    Private mfraDetail As fraTradeHistoryDetail

    Private mlLastUpdate As Int32 = -1
    Private Shared miLastListIndex As Int32 = -1
    Private mlSortNum As eListSortType = eListSortType.eDate Or eListSortType.eDescendOrder

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'fraTradeHistory initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eTradeHistory
            .ControlName = "fraTradeHistory"
            .Left = 78
            .Top = 110
            .Width = 790
            .Height = 543
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
            .mbAcceptReprocessEvents = True
        End With

        'lstHistory initial props
        lstHistory = New UIListBox(oUILib)
        With lstHistory
            .ControlName = "lstHistory"
            .Left = 5
            .Top = 5
            .Width = 780
            .Height = 350
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .sHeaderRow = "DATE    PLAYER                TYPE                  RESULT                AMOUNT"
            .SetFont(New System.Drawing.Font("Courier New", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(lstHistory, UIControl))

        mfraDetail = New fraTradeHistoryDetail(oUILib)
        With mfraDetail
            .ControlName = "fraTradeHistoryDetail"
            .Left = 5
            .Top = lstHistory.Top + lstHistory.Height + 5
            .Width = 780
            .Height = 170
            .Enabled = True
            .Visible = True
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(mfraDetail, UIControl))
        AddHandler mfraDetail.DeleteSelectedItem, AddressOf frame_DeleteSelectedItem

        FillHistoryList()
    End Sub

    Public Sub fraTradeHistory_OnNewFrame() Handles Me.OnNewFrame
        If lstHistory Is Nothing Then Return
        If goCurrentPlayer Is Nothing Then Return

        If glCurrentCycle - mlLastUpdate > 30 Then      'one second

            Try
                With goCurrentPlayer
                    'First check for changes to existing items
                    For X As Int32 = 0 To lstHistory.ListCount - 1
                        Dim sTemp As String = .moTradeHistory(lstHistory.ItemData(X)).GetListBoxText()
                        If sTemp <> lstHistory.List(X) Then lstHistory.List(X) = sTemp
                    Next X

                    'Now, check for items not in the list
                    For X As Int32 = 0 To .mlTradeHistoryUB
                        Dim bFound As Boolean = False
                        For Y As Int32 = 0 To lstHistory.ListCount - 1
                            If lstHistory.ItemData(Y) = X Then
                                bFound = True
                                Exit For
                            End If
                        Next Y
                        If bFound = False Then
                            FillHistoryList()
                            Me.IsDirty = True
                            Exit Try
                        End If
                    Next X
                End With
            Catch
                'do nothing
            End Try
        End If
        miLastListIndex = lstHistory.ScrollBarValue

    End Sub

    Private Sub FillHistoryList()
        Dim lID As Int32 = -1

        If lstHistory.ListIndex <> -1 Then lID = lstHistory.ItemData(lstHistory.ListIndex)
        lstHistory.Clear()

        With goCurrentPlayer
            Dim lSorted() As Int32 = GetSortedList()
            If lSorted Is Nothing = False Then
                For X As Int32 = 0 To lSorted.GetUpperBound(0)
                    lstHistory.AddItem(.moTradeHistory(lSorted(X)).GetListBoxText())
                    lstHistory.ItemData(lstHistory.NewIndex) = lSorted(X)
                Next X
            End If
        End With

        For X As Int32 = 0 To lstHistory.ListCount - 1
            If lstHistory.ItemData(X) = lID Then
                lstHistory.ListIndex = X
                Exit For
            End If
        Next X
        lstHistory.ScrollBarValue = miLastListIndex
    End Sub

    Private Sub lstHistory_ItemClick(ByVal lIndex As Integer) Handles lstHistory.ItemClick
        If lIndex > -1 Then
            Dim lIdx As Int32 = lstHistory.ItemData(lIndex)
            Try
                mfraDetail.SetFromHistoryItem(goCurrentPlayer.moTradeHistory(lIdx))
            Catch
            End Try
        End If
    End Sub

    Private Sub lstHistory_HeaderRowClick(ByVal lX As Integer) Handles lstHistory.HeaderRowClick
        Dim lWordClick As Int32 = 0
        Dim bInSpace As Boolean = False

        Using oTmpFont As New Font(MyBase.moUILib.oDevice, New System.Drawing.Font(lstHistory.GetFont, FontStyle.Bold))
            For X As Int32 = 1 To lstHistory.sHeaderRow.Length - 1

                If lstHistory.sHeaderRow.Substring(X, 1) = " " Then
                    bInSpace = True
                ElseIf bInSpace = True Then
                    lWordClick += 1
                End If

                Dim rcTemp As Rectangle = oTmpFont.MeasureString(Nothing, lstHistory.sHeaderRow.Substring(0, X), DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, Color.White)
                If lX < rcTemp.Width Then
                    Exit For
                End If
            Next X
        End Using

        Dim lCur As Int32 = mlSortNum
        Dim bDescend As Boolean = False
        If (lCur And eListSortType.eDescendOrder) <> 0 Then
            bDescend = True
            lCur -= eListSortType.eDescendOrder
        End If

        If lCur = lWordClick Then
            If bDescend = True Then
                mlSortNum = CType(lCur, eListSortType)
            Else
                mlSortNum = CType(lCur Or eListSortType.eDescendOrder, eListSortType)
            End If
        Else
            mlSortNum = CType(lCur, eListSortType)
        End If

        FillHistoryList()

    End Sub

    Private Function GetSortedList() As Int32()
        Dim lSorted() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1

        Dim lSortBy As Int32 = mlSortNum
        Dim bDescend As Boolean = False
        If (lSortBy And eListSortType.eDescendOrder) <> 0 Then
            lSortBy -= eListSortType.eDescendOrder
            bDescend = True
        End If

        With goCurrentPlayer

            For X As Int32 = 0 To .mlTradeHistoryUB
                Dim lIdx As Int32 = -1

                For Y As Int32 = 0 To lSortedUB

                    Select Case lSortBy
                        Case eListSortType.eAmount, eListSortType.eDate

                            Dim blSortedVal As Int64
                            Dim blCurrentVal As Int64

                            If lSortBy = eListSortType.eAmount Then
                                blSortedVal = .moTradeHistory(lSorted(Y)).blTradeAmt
                                blCurrentVal = .moTradeHistory(X).blTradeAmt
                            Else
                                blSortedVal = .moTradeHistory(lSorted(Y)).lTransactionDate
                                blCurrentVal = .moTradeHistory(X).lTransactionDate
                            End If

                            If bDescend = True Then
                                If blSortedVal < blCurrentVal Then
                                    lIdx = Y
                                    Exit For
                                End If
                            Else
                                If blSortedVal > blCurrentVal Then
                                    lIdx = Y
                                    Exit For
                                End If
                            End If
                        Case Else
                            Dim sSortedVal As String
                            Dim sCurrVal As String

                            If lSortBy = eListSortType.ePlayer Then
                                'sSortedVal = GetCacheObjectValue(.moTradeHistory(lSorted(Y)).lOtherPlayerID, ObjectType.ePlayer)
                                If .moTradeHistory(lSorted(Y)).lOtherPlayerID < 0 Then sSortedVal = "Anonymous" Else sSortedVal = GetCacheObjectValue(.moTradeHistory(lSorted(Y)).lOtherPlayerID, ObjectType.ePlayer)
                                'sCurrVal = GetCacheObjectValue(.moTradeHistory(X).lOtherPlayerID, ObjectType.ePlayer)
                                If .moTradeHistory(X).lOtherPlayerID < 0 Then sCurrVal = "Anonymous" Else sCurrVal = GetCacheObjectValue(.moTradeHistory(X).lOtherPlayerID, ObjectType.ePlayer)
                            ElseIf lSortBy = eListSortType.eResult Then
                                sSortedVal = TradeHistory.GetResultText(.moTradeHistory(lSorted(Y)).yTradeResult)
                                sCurrVal = TradeHistory.GetResultText(.moTradeHistory(X).yTradeResult)
                            Else
                                sSortedVal = TradeHistory.GetTradeTypeText(.moTradeHistory(lSorted(Y)).yTradeEventType)
                                sCurrVal = TradeHistory.GetTradeTypeText(.moTradeHistory(X).yTradeEventType)
                            End If

                            If bDescend = True Then
                                If sSortedVal < sCurrVal Then
                                    lIdx = Y
                                    Exit For
                                End If
                            ElseIf sSortedVal > sCurrVal Then
                                lIdx = Y
                                Exit For
                            End If
                    End Select
                Next Y

                lSortedUB += 1
                ReDim Preserve lSorted(lSortedUB)
                If lIdx = -1 Then
                    lSorted(lSortedUB) = X
                Else
                    For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                        lSorted(Y) = lSorted(Y - 1)
                    Next Y
                    lSorted(lIdx) = X
                End If
            Next X
        End With

        Return lSorted
    End Function

    Private Sub frame_DeleteSelectedItem()
        If lstHistory.ListIndex > -1 Then
            lstHistory.RemoveItem(lstHistory.ListIndex)
            mfraDetail.Visible = False
            lstHistory.ListIndex = -1
        End If
    End Sub

    'Interface created from Interface Builder
    Private Class fraTradeHistoryDetail
        Inherits UIWindow

        Private lblTransType As UILabel
        Private lblOutcome As UILabel
        Private lstReceived As UIListBox
        Private lblReceived As UILabel
        Private lblGiven As UILabel
        Private lstGiven As UIListBox
        Private lblOtherPLayer As UILabel
        Private lblTransDate As UILabel
        Private lblDeliveryTime As UILabel
        Private WithEvents btnDelete As UIButton

        Public Event DeleteSelectedItem()

        Private moItem As TradeHistory = Nothing

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraTradeHistoryDetail initial props
            With Me
                .ControlName = "fraTradeHistoryDetail"
                .Left = 141
                .Top = 230
                .Width = 780
                .Height = 170
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
                .BorderLineWidth = 2
                .mbAcceptReprocessEvents = True
            End With

            'lblTransType initial props
            lblTransType = New UILabel(oUILib)
            With lblTransType
                .ControlName = "lblTransType"
                .Left = 5
                .Top = 5
                .Width = 275
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Transaction Type:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTransType, UIControl))

            'lstReceived initial props
            lstReceived = New UIListBox(oUILib)
            With lstReceived
                .ControlName = "lstReceived"
                .Left = 290
                .Top = 24
                .Width = 220
                .Height = 140
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(lstReceived, UIControl))

            'lblReceived initial props
            lblReceived = New UILabel(oUILib)
            With lblReceived
                .ControlName = "lblReceived"
                .Left = 290
                .Top = 5
                .Width = 100
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Items Received:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblReceived, UIControl))

            'lblGiven initial props
            lblGiven = New UILabel(oUILib)
            With lblGiven
                .ControlName = "lblGiven"
                .Left = 555
                .Top = 5
                .Width = 100
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Items Given:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblGiven, UIControl))

            'lstGiven initial props
            lstGiven = New UIListBox(oUILib)
            With lstGiven
                .ControlName = "lstGiven"
                .Left = 555
                .Top = 25
                .Width = 220
                .Height = 140
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(lstGiven, UIControl))

            'lblOtherPLayer initial props
            lblOtherPLayer = New UILabel(oUILib)
            With lblOtherPLayer
                .ControlName = "lblOtherPLayer"
                .Left = 5
                .Top = 35
                .Width = 275
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Other Player:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblOtherPLayer, UIControl))

            'lblTransDate initial props
            lblTransDate = New UILabel(oUILib)
            With lblTransDate
                .ControlName = "lblTransDate"
                .Left = 5
                .Top = 55
                .Width = 275
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Transaction Date:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblTransDate, UIControl))

            'lblDeliveryTime initial props
            lblDeliveryTime = New UILabel(oUILib)
            With lblDeliveryTime
                .ControlName = "lblDeliveryTime"
                .Left = 5
                .Top = 75
                .Width = 275
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Time to Deliver:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblDeliveryTime, UIControl))

            'lblOutcome initial props
            lblOutcome = New UILabel(oUILib)
            With lblOutcome
                .ControlName = "lblOutcome"
                .Left = 5
                .Top = 95
                .Width = 275
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Transaction Result: "
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblOutcome, UIControl))

            'btnDelete initial props
            btnDelete = New UIButton(oUILib)
            With btnDelete
                .ControlName = "btnDelete"
                .Left = 85
                .Top = 135
                .Width = 120
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Delete History"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnDelete, UIControl))
        End Sub

        Public Sub SetFromHistoryItem(ByRef oHistory As TradeHistory)
            moItem = oHistory
            lstGiven.Clear()
            lstReceived.Clear()
            If oHistory Is Nothing = False Then
                With oHistory
                    Dim sTemp As String = "Transaction Result: "
                    Select Case .yTradeResult
                        Case TradeHistory.TradeResult.eBuyOrderAcceptorFailed
                            sTemp &= "Acceptor Failed"
                        Case TradeHistory.TradeResult.eBuyOrderEscrow
                            sTemp &= "Escrow Deposit"
                        Case TradeHistory.TradeResult.eBuyOrderFinished
                            sTemp &= "Acceptor Finished"
                        Case TradeHistory.TradeResult.eBuyOrderPlaced
                            sTemp &= "Buy Order Placed"
                        Case TradeHistory.TradeResult.eCompletedSuccessfully
                            sTemp &= "Completed Successfully"
                        Case TradeHistory.TradeResult.eNoBuyOrderAcceptor
                            sTemp &= "Expired, No Acceptor"
                        Case Else
                            sTemp &= "Unknown"
                    End Select
                    lblOutcome.Caption = sTemp

                    sTemp = "Transaction Type: "
                    Dim bBuyer As Boolean = (.yTradeEventType And TradeHistory.TradeEventType.eBuyer) <> 0
                    Dim bSeller As Boolean = (.yTradeEventType And TradeHistory.TradeEventType.eSeller) <> 0
                    Dim yTemp As TradeHistory.TradeEventType = .yTradeEventType
                    If bBuyer = True Then yTemp = yTemp Xor TradeHistory.TradeEventType.eBuyer
                    If bSeller = True Then yTemp = yTemp Xor TradeHistory.TradeEventType.eSeller
                    Select Case yTemp
                        Case TradeHistory.TradeEventType.eBuyOrder
                            sTemp &= "Buy Order"
                        Case TradeHistory.TradeEventType.eDirectTrade
                            sTemp &= "Direct Trade"
                        Case TradeHistory.TradeEventType.eSellOrder
                            sTemp &= "Sell Order"
                        Case Else
                            sTemp &= "Unknown"
                    End Select
                    lblTransType.Caption = sTemp

                    If bBuyer = True AndAlso bSeller = True Then
                        lblOtherPLayer.Caption = "Other Player: Yourself"
                    Else
                        If .lOtherPlayerID < 0 Then lblOtherPLayer.Caption = "Other Player: Anonymous" Else lblOtherPLayer.Caption = "Other Player: " & GetCacheObjectValue(.lOtherPlayerID, ObjectType.ePlayer)
                    End If

                    Dim dtTemp As Date = GetDateFromNumber(.lTransactionDate)
                    lblTransDate.Caption = "Transaction Date: " & dtTemp.ToShortDateString & " " & dtTemp.ToShortTimeString

                    lblDeliveryTime.Caption = "Delivery Time: " & GetDurationFromSeconds(.lDeliveryTime, True)

                    .FillItemListbox(lstReceived, False)
                    .FillItemListbox(lstGiven, True)
                End With
            Else
                lblTransType.Caption = "Transaction Type: Select an item"
                lblOtherPLayer.Caption = "Other Player: Select an item"
                lblTransDate.Caption = "Transaction Date: Select an item"
                lblDeliveryTime.Caption = "Delivery Time: Select an item"
                lblOutcome.Caption = "Transaction Result: Select an item"
            End If
        End Sub

        Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
            'delete me
            If moItem Is Nothing = False Then
                Dim yMsg(15) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eDeleteTradeHistoryItem).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(moItem.lPlayerID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(moItem.lTransactionDate).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(Math.Abs(moItem.lOtherPlayerID)).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = moItem.yTradeResult : lPos += 1
                yMsg(lPos) = moItem.yTradeEventType : lPos += 1
                MyBase.moUILib.SendMsgToPrimary(yMsg)

                RaiseEvent DeleteSelectedItem()
            End If
        End Sub
    End Class

End Class