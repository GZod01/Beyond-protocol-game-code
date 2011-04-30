Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class fraTradeInProgress
    Inherits UIWindow

    Private lblDirectTrades As UILabel
    Private WithEvents lstDirectTrades As UIListBox
    Private lblTradeDelivery As UILabel
    Private WithEvents lstTradeDelivery As UIListBox
    Private WithEvents btnViewDirectTrade As UIButton
    Private lnDiv1 As UILine
    Private Shared miLastListIndex As Int32

    Private mfraTradeDeliveryDetail As fraTradeDeliveryDetail

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'fraTradeInProgress initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eTradeInProgress
            .ControlName = "fraTradeInProgress"
            .Left = 88
            .Top = 54
            .Width = 790
            .Height = 543
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With

        'lblDirectTrades initial props
        lblDirectTrades = New UILabel(oUILib)
        With lblDirectTrades
            .ControlName = "lblDirectTrades"
            .Left = 5
            .Top = 5
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Direct Trades"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDirectTrades, UIControl))

        'lstDirectTrades initial props
        lstDirectTrades = New UIListBox(oUILib)
        With lstDirectTrades
            .ControlName = "lstDirectTrades"
            .Left = 5
            .Top = 25
            .Width = 780
            .Height = 150
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .sHeaderRow = "With Player".PadRight(22, " "c) & "Status"
        End With
        Me.AddChild(CType(lstDirectTrades, UIControl))

        'lblTradeDelivery initial props
        lblTradeDelivery = New UILabel(oUILib)
        With lblTradeDelivery
            .ControlName = "lblTradeDelivery"
            .Left = 5
            .Top = 220
            .Width = 186
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "GTC Trade Delivery Status"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTradeDelivery, UIControl))

        'lstTradeDelivery initial props
        lstTradeDelivery = New UIListBox(oUILib)
        With lstTradeDelivery
            .ControlName = "lstTradeDelivery"
            .Left = 5
            .Top = 240
            .Width = 390
            .Height = 298
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .sHeaderRow = "From Player".PadRight(22, " "c) & "ETA"
        End With
        Me.AddChild(CType(lstTradeDelivery, UIControl))

        'btnViewDirectTrade initial props
        btnViewDirectTrade = New UIButton(oUILib)
        With btnViewDirectTrade
            .ControlName = "btnViewDirectTrade"
            .Left = 685
            .Top = 180
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "View"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnViewDirectTrade, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 1
            .Top = 210
            .Width = 789
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'mfraTradeDeliveryDetail initial props
        mfraTradeDeliveryDetail = New fraTradeDeliveryDetail(oUILib)
        With mfraTradeDeliveryDetail
            .ControlName = "mfraTradeDeliveryDetail"
            .Left = 405
            .Top = 240
            .Width = 380
            .Height = 298
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With
        Me.AddChild(CType(mfraTradeDeliveryDetail, UIControl))
    End Sub

    Public Sub fraTradeInProgress_OnNewFrame() Handles Me.OnNewFrame

        Try
            If goCurrentPlayer Is Nothing = False Then
                Dim bListItemFound(lstDirectTrades.ListCount - 1) As Boolean
                For X As Int32 = 0 To lstDirectTrades.ListCount - 1
                    bListItemFound(X) = False
                Next X
                For X As Int32 = 0 To goCurrentPlayer.mlTradeUB
                    If goCurrentPlayer.mlTradeIdx(X) <> -1 Then
                        Dim oTrade As Trade = goCurrentPlayer.moTrades(X)
                        If oTrade Is Nothing = False AndAlso goCurrentPlayer.moTrades(X).yTradeState <> Trade.eTradeStateValues.TradeRejected AndAlso (goCurrentPlayer.moTrades(X).yTradeState And Trade.eTradeStateValues.TradeCompleted) = 0 Then
                            Dim bFound As Boolean = False
                            Dim lID As Int32 = goCurrentPlayer.mlTradeIdx(X)
                            Dim sTemp As String = goCurrentPlayer.moTrades(X).GetTradeMainListText()

                            For Y As Int32 = 0 To lstDirectTrades.ListCount - 1
                                If lstDirectTrades.ItemData(Y) = lID Then
                                    If lstDirectTrades.List(Y) <> sTemp Then lstDirectTrades.List(Y) = sTemp
                                    bListItemFound(Y) = True
                                    bFound = True
                                    Exit For
                                End If
                            Next Y
                            If bFound = False Then
                                lstDirectTrades.AddItem(sTemp)
                                lstDirectTrades.ItemData(lstDirectTrades.NewIndex) = lID
                                lstDirectTrades.ScrollBarValue = miLastListIndex
                            End If
                        End If
                    End If
                Next X
                For X As Int32 = 0 To lstDirectTrades.ListCount - 1
                    If bListItemFound(X) = False Then
                        lstDirectTrades.RemoveItem(X)
                        Exit For
                    End If
                Next X
            End If

            Dim lCnt As Int32 = 0
            For X As Int32 = 0 To TradeDeliveryPackage.lTradeDeliveryUB
                If TradeDeliveryPackage.yTradeDeliveryUsed(X) <> 0 Then
                    With TradeDeliveryPackage.oTradeDeliveries(X)
                        Dim sTemp As String = .GetListBoxText()
                        Dim bFound As Boolean = False

                        lCnt += 1

                        For Y As Int32 = 0 To lstTradeDelivery.ListCount - 1
                            If lstTradeDelivery.ItemData(Y) = X Then
                                bFound = True
                                If lstTradeDelivery.List(Y) <> sTemp Then
                                    lstTradeDelivery.List(Y) = sTemp
                                End If
                                Exit For
                            End If
                        Next Y
                        If bFound = False Then
                            lstTradeDelivery.AddItem(sTemp)
                            lstTradeDelivery.ItemData(lstTradeDelivery.NewIndex) = X
                        End If
                    End With
                End If
            Next X
            If lstTradeDelivery.ListCount <> lCnt Then
                lstTradeDelivery.Clear()
                Return
            End If
            miLastListIndex = lstDirectTrades.ScrollBarValue
            mfraTradeDeliveryDetail.fraTradeDeliveryDetail_OnNewFrame()

        Catch
        End Try
    End Sub

	Private Sub btnViewDirectTrade_Click(ByVal sName As String) Handles btnViewDirectTrade.Click
		If lstDirectTrades.ListIndex > -1 Then
			Dim lTradeID As Int32 = lstDirectTrades.ItemData(lstDirectTrades.ListIndex)
			Dim oTrade As Trade = Nothing
			For X As Int32 = 0 To goCurrentPlayer.mlTradeUB
				If goCurrentPlayer.mlTradeIdx(X) = lTradeID Then
					oTrade = goCurrentPlayer.moTrades(X)
					Exit For
				End If
			Next X
			If oTrade Is Nothing = False Then
				CType(Me.ParentControl, frmTradeMain).ViewDirectTradeDetails(oTrade)
			End If
		End If
	End Sub

    Private Sub lstDirectTrades_ItemDblClick(ByVal lIndex As Integer) Handles lstDirectTrades.ItemDblClick
		btnViewDirectTrade_Click(btnViewDirectTrade.ControlName)
    End Sub

    Private Sub lstTradeDelivery_ItemClick(ByVal lIndex As Integer) Handles lstTradeDelivery.ItemClick
        If lIndex <> -1 Then
            Dim lIdx As Int32 = lstTradeDelivery.ItemData(lIndex)
            If lIdx > -1 AndAlso lIdx <= TradeDeliveryPackage.lTradeDeliveryUB AndAlso TradeDeliveryPackage.yTradeDeliveryUsed(lIdx) <> 0 Then
                Dim oTDP As TradeDeliveryPackage = TradeDeliveryPackage.oTradeDeliveries(lIdx)
                If oTDP Is Nothing = False Then
                    mfraTradeDeliveryDetail.SetTradeDelivery(oTDP)
                End If
            End If
        End If
    End Sub

#Region " fraTradeDeliveryDetail "
    'Interface created from Interface Builder
    Private Class fraTradeDeliveryDetail
        Inherits UIWindow

        Private lblPlayer As UILabel
        Private lblETA As UILabel
        'Private lblOriginalETA As UILabel
        Private lblDestination As UILabel
        Private lblGoodShipped As UILabel
        Private WithEvents lstShipped As UIListBox
        Private txtItemDetails As UITextBox

        Private moTDP As TradeDeliveryPackage = Nothing

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraTradeDeliveryDetail initial props
            With Me
                .ControlName = "fraTradeDeliveryDetail"
                .Left = 85
                .Top = 81
                .Width = 380
                .Height = 298
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
                .BorderLineWidth = 2
            End With

            'lblPlayer initial props
            lblPlayer = New UILabel(oUILib)
            With lblPlayer
                .ControlName = "lblPlayer"
                .Left = 10
                .Top = 10
                .Width = 360
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "From Player: "
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblPlayer, UIControl))

            'lblETA initial props
            lblETA = New UILabel(oUILib)
            With lblETA
                .ControlName = "lblETA"
                .Left = 10
                .Top = 30
                .Width = 360
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Estimated Time of Arrival: "
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblETA, UIControl))

            ''lblOriginalETA initial props
            'lblOriginalETA = New UILabel(oUILib)
            'With lblOriginalETA
            '    .ControlName = "lblOriginalETA"
            '    .Left = 10
            '    .Top = 50
            '    .Width = 360
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Original Estimated Time of Arrival: "
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            'End With
            'Me.AddChild(CType(lblOriginalETA, UIControl))

            'lblDestination initial props
            lblDestination = New UILabel(oUILib)
            With lblDestination
                .ControlName = "lblDestination"
                .Left = 10
                .Top = 50
                .Width = 360
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Destination: "
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblDestination, UIControl))

            'lblGoodShipped initial props
            lblGoodShipped = New UILabel(oUILib)
            With lblGoodShipped
                .ControlName = "lblGoodShipped"
                .Left = 10
                .Top = 70
                .Width = 110
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Goods Shipped"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblGoodShipped, UIControl))

            'lstShipped initial props
            lstShipped = New UIListBox(oUILib)
            With lstShipped
                .ControlName = "lstShipped"
                .Left = 10
                .Top = 90
                .Width = 200
                .Height = 198
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            End With
            Me.AddChild(CType(lstShipped, UIControl))

            'txtItemDetails initial props
            txtItemDetails = New UITextBox(oUILib)
            With txtItemDetails
                .ControlName = "txtItemDetails"
                .Left = 215
                .Top = 90
                .Width = 155
                .Height = 198
                .Enabled = True
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(0, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 0
                .BorderColor = muSettings.InterfaceBorderColor
                .MultiLine = True
                .Locked = True
            End With
            Me.AddChild(CType(txtItemDetails, UIControl))
        End Sub

        Public Sub SetTradeDelivery(ByRef oTradeDelivery As TradeDeliveryPackage)
            moTDP = oTradeDelivery
            lstShipped.Clear()
        End Sub

        Public Sub fraTradeDeliveryDetail_OnNewFrame() Handles Me.OnNewFrame
            If moTDP Is Nothing = False Then
                Dim sTemp As String '= "From Player: " '& GetCacheObjectValue(moTDP.lSourcePlayerID, ObjectType.ePlayer)
                If moTDP.ItemContainsIntel() = True Then
                    sTemp = "From Player: Anonymous"
                Else
                    sTemp = "From Player: " & GetCacheObjectValue(moTDP.lSourcePlayerID, ObjectType.ePlayer)
                End If
                If lblPlayer.Caption <> sTemp Then lblPlayer.Caption = sTemp

                sTemp = "Estimated Time of Arrival: " & moTDP.GetDeliveryText()
                If lblETA.Caption <> sTemp Then lblETA.Caption = sTemp

                sTemp = "Destination: " & moTDP.GetDestinationText()
                If lblDestination.Caption <> sTemp Then lblDestination.Caption = sTemp

                moTDP.SmartFillListBox(lstShipped)

                If lstShipped.ListIndex > -1 Then
                    sTemp = GetNonOwnerItemData(lstShipped.ItemData(lstShipped.ListIndex), CShort(lstShipped.ItemData2(lstShipped.ListIndex)))
                    If txtItemDetails.Caption <> sTemp Then txtItemDetails.Caption = sTemp
                End If
            End If
        End Sub
    End Class
#End Region
End Class