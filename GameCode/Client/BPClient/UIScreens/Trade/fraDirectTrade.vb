Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class fraDirectTrade
    Inherits UIWindow

    Private lblSellToPlayer As UILabel
    Private lblSellables As UILabel
    Private lblQuantity As UILabel
    Private lblYourOffer As UILabel
    Private lblTheirOffer As UILabel
    Private txtNotesToMe As UITextBox
    Private lblMsgToYou As UILabel
    Private lblMsgToThem As UILabel
    Private txtNotesToThem As UITextBox
    Private chkAccept As UICheckBox
    Private lblItemDetails As UILabel
    Private txtItemDetails As UITextBox
    Private lnDiv1 As UILine
    Private lblTax As UILabel

    Private WithEvents chkTheirAccept As UICheckBox
    Private WithEvents cboSellToPlayer As UIComboBox
    Private WithEvents lstTradeables As UIListBox
    Private WithEvents txtQuantity As UITextBox
    Private WithEvents btnAdd As UIButton
    Private WithEvents lstOffer As UIListBox
    Private WithEvents btnRemove As UIButton
    Private WithEvents lstGetting As UIListBox
    Private WithEvents btnSubmit As UIButton
    Private WithEvents btnReject As UIButton

    Private moTrade As Trade = Nothing
    Private mblQtys() As Int64
    Private mlTradePostID As Int32 = -1
    Private mlLastUpdate As Int32 = -1

    Private mlTPCLast As Int32 = -1

    Private mlCurrentItemDetail As Int32 = -1
	Private miCurrentItemDetailType As Int16 = -1
	Private mlCurrentItemExtended As Int32 = -1
	Private mlCurrentItemOwnerID As Int32 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'fraDirectTrade initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eDirectTrade
            .ControlName = "fraDirectTrade"
            .Left = 93
            .Top = 83
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

        'lblSellToPlayer initial props
        lblSellToPlayer = New UILabel(oUILib)
        With lblSellToPlayer
            .ControlName = "lblSellToPlayer"
            .Left = 5
            .Top = 5
            .Width = 125
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Trade with Player:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSellToPlayer, UIControl))

        'lblSellables initial props
        lblSellables = New UILabel(oUILib)
        With lblSellables
            .ControlName = "lblSellables"
            .Left = 5
            .Top = 35
            .Width = 80
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Tradeables"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSellables, UIControl))

        'lstSellables initial props
        lstTradeables = New UIListBox(oUILib)
        With lstTradeables
            .ControlName = "lstTradeables"
            .Left = 5
            .Top = 55
            .Width = 250
            .Height = 450
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(lstTradeables, UIControl))

        'lblQuantity initial props
        lblQuantity = New UILabel(oUILib)
        With lblQuantity
            .ControlName = "lblQuantity"
            .Left = 5
            .Top = 515
            .Width = 63
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Quantity:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblQuantity, UIControl))

        'txtQuantity initial props
        txtQuantity = New UITextBox(oUILib)
        With txtQuantity
            .ControlName = "txtQuantity"
            .Left = 70
            .Top = 515
            .Width = 128
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "0"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(txtQuantity, UIControl))

        'btnAdd initial props
        btnAdd = New UIButton(oUILib)
        With btnAdd
            .ControlName = "btnAdd"
            .Left = 200
            .Top = 513
            .Width = 55
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Add"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAdd, UIControl))

        'lblYourOffer initial props
        lblYourOffer = New UILabel(oUILib)
        With lblYourOffer
            .ControlName = "lblYourOffer"
            .Left = 270
            .Top = 35
            .Width = 120
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "You Are Offering:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblYourOffer, UIControl))

        'lstYourOffer initial props
        lstOffer = New UIListBox(oUILib)
        With lstOffer
            .ControlName = "lstOffer"
            .Left = 270
            .Top = 55
            .Width = 250
            .Height = 200
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(lstOffer, UIControl))

        'btnRemove initial props
        btnRemove = New UIButton(oUILib)
        With btnRemove
            .ControlName = "btnRemove"
            .Left = 420
            .Top = 260
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Remove"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnRemove, UIControl))

        'lblTheirOffer initial props
        lblTheirOffer = New UILabel(oUILib)
        With lblTheirOffer
            .ControlName = "lblTheirOffer"
            .Left = 270
            .Top = 285
            .Width = 127
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "They Are Offering:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTheirOffer, UIControl))

        'lstTheirOffer initial props
        lstGetting = New UIListBox(oUILib)
        With lstGetting
            .ControlName = "lstGetting"
            .Left = 270
            .Top = 305
            .Width = 250
            .Height = 200
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(lstGetting, UIControl))

        'txtMsgToYou initial props
        txtNotesToMe = New UITextBox(oUILib)
        With txtNotesToMe
            .ControlName = "txtNotesToMe"
            .Left = 535
            .Top = 275
            .Width = 240
            .Height = 100
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .Locked = True
            .MultiLine = True

        End With
        Me.AddChild(CType(txtNotesToMe, UIControl))

        'lblMsgToYou initial props
        lblMsgToYou = New UILabel(oUILib)
        With lblMsgToYou
            .ControlName = "lblMsgToYou"
            .Left = 535
            .Top = 255
            .Width = 139
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Messages For You:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblMsgToYou, UIControl))

        'lblMsgToThem initial props
        lblMsgToThem = New UILabel(oUILib)
        With lblMsgToThem
            .ControlName = "lblMsgToThem"
            .Left = 535
            .Top = 385
            .Width = 150
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Messages For Them:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblMsgToThem, UIControl))

        'txtMsgToThem initial props
        txtNotesToThem = New UITextBox(oUILib)
        With txtNotesToThem
            .ControlName = "txtNotesToThem"
            .Left = 535
            .Top = 405
            .Width = 240
            .Height = 100
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
        End With
        Me.AddChild(CType(txtNotesToThem, UIControl))

        'chkAccept initial props
        chkAccept = New UICheckBox(oUILib)
        With chkAccept
            .ControlName = "chkAccept"
            .Left = 365
            .Top = 5
            .Width = 161
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Accept This Agreement"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkAccept, UIControl))

        'lblItemDetails initial props
        lblItemDetails = New UILabel(oUILib)
        With lblItemDetails
            .ControlName = "lblItemDetails"
            .Left = 535
            .Top = 35
            .Width = 139
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Item Details:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblItemDetails, UIControl))

        'txtItemDetails initial props
        txtItemDetails = New UITextBox(oUILib)
        With txtItemDetails
            .ControlName = "txtItemDetails"
            .Left = 535
            .Top = 55
            .Width = 240
            .Height = 190
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

        'btnSubmit initial props
        btnSubmit = New UIButton(oUILib)
        With btnSubmit
            .ControlName = "btnSubmit"
            .Left = 545
            .Top = 515
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Submit"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSubmit, UIControl))

        'btnReject initial props
        btnReject = New UIButton(oUILib)
        With btnReject
            .ControlName = "btnReject"
            .Left = 675
            .Top = 515
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Reject"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnReject, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = Me.BorderLineWidth \ 2
            .Top = 30
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'chkTheirAccept initial props
        chkTheirAccept = New UICheckBox(oUILib)
        With chkTheirAccept
            .ControlName = "chkTheirAccept"
            .Left = 580
            .Top = 5
            .Width = 161
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Their Acceptance"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkTheirAccept, UIControl))

        'lblTax initial props
        lblTax = New UILabel(oUILib)
        With lblTax
            .ControlName = "lblTax"
            .Left = lblTheirOffer.Left
            .Top = lblQuantity.Top
            .Width = lstGetting.Width
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "GTC Tax: 10,000,000"
            .ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Both parties will be charged this amount by the GTC." & vbCrLf & "This amount is based on the credits exchanged in the trade."
        End With
        Me.AddChild(CType(lblTax, UIControl))

        'cboSellToPlayer initial props
        cboSellToPlayer = New UIComboBox(oUILib)
        With cboSellToPlayer
            .ControlName = "cboSellToPlayer"
            .Left = 135
            .Top = 5
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .l_ListBoxHeight = 300
        End With
        Me.AddChild(CType(cboSellToPlayer, UIControl))

        FillPlayerList()
    End Sub

    Public Sub SetTradePost(ByVal lTP_ID As Int32)
        mlTradePostID = lTP_ID
    End Sub

    Private Sub FillPlayerList()
		cboSellToPlayer.Clear()

		Dim lSorted() As Int32 = GetSortedPlayerRelIdxArray()
		If lSorted Is Nothing = False Then
			For X As Int32 = 0 To lSorted.GetUpperBound(0)
				Dim oTmpRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(lSorted(X))
				If oTmpRel Is Nothing = False Then
					Dim lPlayerID As Int32 = -1
					If oTmpRel.lPlayerRegards = glPlayerID Then
						lPlayerID = oTmpRel.lThisPlayer
					Else : lPlayerID = oTmpRel.lPlayerRegards
					End If

					If lPlayerID <> -1 Then
						Dim sName As String = GetCacheObjectValue(lPlayerID, ObjectType.ePlayer)
						cboSellToPlayer.AddItem(sName)
						cboSellToPlayer.ItemData(cboSellToPlayer.NewIndex) = lPlayerID
					End If
				End If
			Next X
		End If
	End Sub

    Public Sub SetFromTradeObject(ByRef oTrade As Trade)
        moTrade = oTrade
        chkAccept.Enabled = True

        If moTrade.yTradeState = Trade.eTradeStateValues.Proposal Then
            Dim sOther As String
            If moTrade.lPlayer1ID = glPlayerID Then
                sOther = GetCacheObjectValue(moTrade.lPlayer2ID, ObjectType.ePlayer)
                cboSellToPlayer.FindComboItemData(moTrade.lPlayer2ID)
            Else
                sOther = GetCacheObjectValue(moTrade.lPlayer1ID, ObjectType.ePlayer)
                cboSellToPlayer.FindComboItemData(moTrade.lPlayer1ID)
            End If

            chkTheirAccept.Caption = sOther & "'s Acceptance"
            Dim rcTemp As Rectangle = chkTheirAccept.GetTextDimensions()
            chkTheirAccept.Width = rcTemp.Width + 20
        Else
            Dim sOther As String
            If moTrade.lPlayer1ID = glPlayerID Then
                sOther = GetCacheObjectValue(moTrade.lPlayer2ID, ObjectType.ePlayer)
                cboSellToPlayer.FindComboItemData(moTrade.lPlayer2ID)
            Else
                sOther = GetCacheObjectValue(moTrade.lPlayer1ID, ObjectType.ePlayer)
                cboSellToPlayer.FindComboItemData(moTrade.lPlayer1ID)
            End If

            chkTheirAccept.Caption = sOther & "'s Acceptance"
            Dim rcTemp As Rectangle = chkTheirAccept.GetTextDimensions()
            chkTheirAccept.Width = rcTemp.Width + 20
        End If

        lstOffer.Clear()
        lstGetting.Clear()
        Dim iTotalQty As Int64 = 0
        With moTrade 
            If .lPlayer1ID = glPlayerID Then
                'Fill our offer list
                iTotalQty = 0
                ReDim mblQtys(.mlPlayer1ItemUB)
                For X As Int32 = 0 To .mlPlayer1ItemUB
                    Dim sValue As String = GetOfferDescription(.muPlayer1Items(X).ObjectID, .muPlayer1Items(X).ObjTypeID, .muPlayer1Items(X).Quantity, .muPlayer1Items(X).lExtendedID, .lPlayer1ID)
                    lstOffer.AddItem(sValue)
                    lstOffer.ItemData(lstOffer.NewIndex) = .muPlayer1Items(X).ObjectID
                    lstOffer.ItemData2(lstOffer.NewIndex) = .muPlayer1Items(X).ObjTypeID
                    lstOffer.ItemData3(lstOffer.NewIndex) = .muPlayer1Items(X).lExtendedID
                    mblQtys(lstOffer.NewIndex) = .muPlayer1Items(X).Quantity
                    iTotalQty += .muPlayer1Items(X).Quantity
                Next X
                lstOffer.sHeaderRow = "Items: " & (.mlPlayer1ItemUB + 1).ToString("#,##0") & " Qty: " & iTotalQty.ToString("#,##0")

                'Fill our getting list
                iTotalQty = 0
                For X As Int32 = 0 To .mlPlayer2ItemUB
                    Dim sValue As String = GetTheirOfferDescription(.muPlayer2Items(X).ObjectID, .muPlayer2Items(X).ObjTypeID, .muPlayer2Items(X).Quantity, .muPlayer2Items(X).lExtendedID, .lPlayer2ID)
                    lstGetting.AddItem(sValue)
                    lstGetting.ItemData(lstGetting.NewIndex) = .muPlayer2Items(X).ObjectID
                    lstGetting.ItemData2(lstGetting.NewIndex) = .muPlayer2Items(X).ObjTypeID
                    lstGetting.ItemData3(lstGetting.NewIndex) = .muPlayer2Items(X).lExtendedID
                    iTotalQty += .muPlayer2Items(X).Quantity
                Next X
                lstGetting.sHeaderRow = "Items: " & (.mlPlayer2ItemUB + 1).ToString("#,##0") & " Qty: " & iTotalQty.ToString("#,##0")

                txtNotesToMe.Caption = .sP2Notes
                txtNotesToThem.Caption = .sP1Notes

                If (.yTradeState And Trade.eTradeStateValues.Player1Accepted) <> 0 Then
                    chkAccept.Value = True
                Else : chkAccept.Value = False
                End If
                If (.yTradeState And Trade.eTradeStateValues.Player2Accepted) <> 0 Then
                    chkTheirAccept.Value = True
                Else : chkTheirAccept.Value = False
                End If
            Else
                'cboSource.FindComboItemData(.lP2SourceID)
                'cboDest.FindComboItemData(.lP2DestID)

                'Fill our offer list
                iTotalQty = 0
                ReDim mblQtys(.mlPlayer2ItemUB)
                For X As Int32 = 0 To .mlPlayer2ItemUB
                    Dim sValue As String = GetOfferDescription(.muPlayer2Items(X).ObjectID, .muPlayer2Items(X).ObjTypeID, .muPlayer2Items(X).Quantity, .muPlayer2Items(X).lExtendedID, .lPlayer2ID)
                    lstOffer.AddItem(sValue)
                    lstOffer.ItemData(lstOffer.NewIndex) = .muPlayer2Items(X).ObjectID
                    lstOffer.ItemData2(lstOffer.NewIndex) = .muPlayer2Items(X).ObjTypeID
                    lstOffer.ItemData3(lstOffer.NewIndex) = .muPlayer2Items(X).lExtendedID
                    mblQtys(lstOffer.NewIndex) = .muPlayer2Items(X).Quantity
                    iTotalQty += .muPlayer2Items(X).Quantity
                Next X
                lstOffer.sHeaderRow = "Items: " & (.mlPlayer2ItemUB + 1).ToString("#,##0") & " Qty: " & iTotalQty.ToString("#,##0")

                'Fill our getting list
                iTotalQty = 0
                For X As Int32 = 0 To .mlPlayer1ItemUB
                    Dim sValue As String = GetTheirOfferDescription(.muPlayer1Items(X).ObjectID, .muPlayer1Items(X).ObjTypeID, .muPlayer1Items(X).Quantity, .muPlayer1Items(X).lExtendedID, .lPlayer1ID)
                    lstGetting.AddItem(sValue)
                    lstGetting.ItemData(lstGetting.NewIndex) = .muPlayer1Items(X).ObjectID
                    lstGetting.ItemData2(lstGetting.NewIndex) = .muPlayer1Items(X).ObjTypeID
                    lstGetting.ItemData3(lstGetting.NewIndex) = .muPlayer1Items(X).lExtendedID
                    iTotalQty += .muPlayer1Items(X).Quantity
                Next X
                lstGetting.sHeaderRow = "Items: " & (.mlPlayer1ItemUB + 1).ToString("#,##0") & " Qty: " & iTotalQty.ToString("#,##0")

                txtNotesToMe.Caption = .sP1Notes
                txtNotesToThem.Caption = .sP2Notes

                If (.yTradeState And Trade.eTradeStateValues.Player1Accepted) <> 0 Then
                    chkTheirAccept.Value = True
                Else : chkTheirAccept.Value = False
                End If
                If (.yTradeState And Trade.eTradeStateValues.Player2Accepted) <> 0 Then
                    chkAccept.Value = True
                Else : chkAccept.Value = False
                End If
            End If

            If chkAccept.Value = True AndAlso chkTheirAccept.Value = True AndAlso (.yFailureReason = Trade.eFailureReason.NoFailureReason) Then
                btnReject.Enabled = False
                btnSubmit.Enabled = False
                btnAdd.Enabled = False
                txtQuantity.Enabled = False
                txtNotesToMe.Enabled = False
                txtNotesToThem.Enabled = False
                btnRemove.Enabled = False
                chkAccept.Enabled = False
                chkTheirAccept.Enabled = False
            End If
        End With


    End Sub

    Private Sub UpdateTaxLabel()

        Dim blCredInTrade As Int64 = 0
        Dim blTemp As Int64 = 0

        If moTrade Is Nothing = False Then
            With moTrade
                blTemp = 0
                For X As Int32 = 0 To .mlPlayer1ItemUB
                    If .muPlayer1Items(X).ObjTypeID = ObjectType.eCredits Then
                        blTemp += .muPlayer1Items(X).Quantity
                    End If
                Next X
                blCredInTrade = Math.Max(blCredInTrade, blTemp)

                blTemp = 0
                For X As Int32 = 0 To .mlPlayer2ItemUB
                    If .muPlayer2Items(X).ObjTypeID = ObjectType.eCredits Then
                        blTemp += .muPlayer2Items(X).Quantity
                    End If
                Next X
                blCredInTrade = Math.Max(blCredInTrade, blTemp)
            End With
        End If

        blTemp = 0
        For X As Int32 = 0 To lstOffer.ListCount - 1
            If lstOffer.ItemData2(X) = ObjectType.eCredits Then
                blTemp += mblQtys(X)
            End If
        Next X
        blCredInTrade = Math.Max(blCredInTrade, blTemp)

        blCredInTrade \= 10
        blCredInTrade *= 3
        blCredInTrade = Math.Max(blCredInTrade, 10000000)
        Dim sTax As String = "GTC Tax: " & blCredInTrade.ToString("#,##0")
        If lblTax.Caption <> sTax Then lblTax.Caption = sTax
    End Sub

    Public Sub CreateNewTrade()
        If cboSellToPlayer.ListIndex = -1 Then Return

        Dim lOtherPlayerID As Int32 = cboSellToPlayer.ItemData(cboSellToPlayer.ListIndex)
        moTrade = Nothing
        chkAccept.Enabled = False
        lstGetting.Clear()
        lstOffer.Clear()

        Dim sOther As String = GetCacheObjectValue(lOtherPlayerID, ObjectType.ePlayer)
        chkTheirAccept.Caption = sOther & "'s Acceptance"
        Dim rcTemp As Rectangle = chkTheirAccept.GetTextDimensions()
        chkTheirAccept.Width = rcTemp.Width + 20
    End Sub

    Private Function ValidateData() As Boolean
        If cboSellToPlayer.ListIndex = -1 Then
            MyBase.moUILib.AddNotification("You must select someone to trade with!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If
        Return True
    End Function

    Private Sub CreateAndSendSubmitTradeMessage(ByVal bReject As Boolean)
        If bReject = False AndAlso ValidateData() = False Then Return

        'Ok, do all the work now
        Dim lPos As Int32 = 0
        Dim sNotes As String = txtNotesToThem.Caption.Trim
		Dim yMsg(20 + sNotes.Length + (lstOffer.ListCount * 18)) As Byte

        Dim lOtherPlayerID As Int32 = cboSellToPlayer.ItemData(cboSellToPlayer.ListIndex)

        System.BitConverter.GetBytes(GlobalMessageCode.eSubmitTrade).CopyTo(yMsg, lPos) : lPos += 2
        If moTrade Is Nothing = False Then
            moTrade.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6

            'Trade state can only be set if the trade is initiated... and then, the player can only set THEIR Accepted/Rejected
            If bReject = True Then
                yMsg(lPos) = 255
            ElseIf chkAccept.Value = True Then
                If moTrade.lPlayer1ID = glPlayerID Then
                    yMsg(lPos) = Trade.eTradeStateValues.Player1Accepted
                Else : yMsg(lPos) = Trade.eTradeStateValues.Player2Accepted
                End If
            End If
            lPos += 1
        Else
            System.BitConverter.GetBytes(-Math.Abs(lOtherPlayerID)).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(ObjectType.eTrade).CopyTo(yMsg, lPos) : lPos += 2
            yMsg(lPos) = 0
            lPos += 1
        End If

        'Next is my dest ID, sourceid
        System.BitConverter.GetBytes(mlTradePostID).CopyTo(yMsg, lPos) : lPos += 4

        'Now, my notes
        System.BitConverter.GetBytes(sNotes.Length).CopyTo(yMsg, lPos) : lPos += 4
        If sNotes.Length > 0 Then
            System.Text.ASCIIEncoding.ASCII.GetBytes(sNotes).CopyTo(yMsg, lPos)
            lPos += sNotes.Length
        End If

        'Next is the count
        If bReject = True Then
            System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
        Else
            System.BitConverter.GetBytes(lstOffer.ListCount).CopyTo(yMsg, lPos) : lPos += 4
            'And now the items...
            For X As Int32 = 0 To lstOffer.ListCount - 1
                System.BitConverter.GetBytes(lstOffer.ItemData(X)).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(CShort(lstOffer.ItemData2(X))).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(mblQtys(X)).CopyTo(yMsg, lPos) : lPos += 8
				System.BitConverter.GetBytes(lstOffer.ItemData3(X)).CopyTo(yMsg, lPos) : lPos += 4
            Next X
        End If

        MyBase.moUILib.SendMsgToPrimary(yMsg)

        goUILib.AddNotification("New trade proposal details submitted.", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

		If moTrade Is Nothing Then
			CType(Me.ParentControl, frmTradeMain).GotoNewView(5)
		End If
    End Sub

	Private Sub btnAdd_Click(ByVal sName As String) Handles btnAdd.Click
		If HasAliasedRights(AliasingRights.eAlterTrades) = False Then
			MyBase.moUILib.AddNotification("You lack rights to alter/place trades.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If
		If lstTradeables.ListIndex > -1 Then
			Dim lID As Int32 = lstTradeables.ItemData(lstTradeables.ListIndex)
			Dim iTypeID As Int16 = CShort(lstTradeables.ItemData2(lstTradeables.ListIndex))
			Dim lExtID As Int32 = lstTradeables.ItemData3(lstTradeables.ListIndex)

			'Eventually, everything will be tradeable, but for now, I need to indicate what isn't implemented
			Select Case iTypeID
				Case ObjectType.eColony
					goUILib.AddNotification("Trading colonies is not implemented yet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				Case ObjectType.eFood
					goUILib.AddNotification("Trading food is not implemented yet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				Case ObjectType.eStock
					goUILib.AddNotification("Trading stock is not implemented yet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eSpecialTech, ObjectType.eWeaponTech
					goUILib.AddNotification("Trading component designs is not implemented yet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
					'Case ObjectType.eFacility
					'    goUILib.AddNotification("Trading facilities is not implemented yet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					'    Return
			End Select

			If IsNumeric(txtQuantity.Caption) = True Then
				Dim blQty As Int64 = CLng(Val(txtQuantity.Caption))
				If blQty > 0 Then
					Dim lIdx As Int32 = -1

					For X As Int32 = 0 To lstOffer.ListCount - 1
						If lstOffer.ItemData(X) = lID AndAlso lstOffer.ItemData2(X) = iTypeID AndAlso lstOffer.ItemData3(X) = lExtID Then
							lIdx = X
							Exit For
						End If
					Next X

					If lIdx = -1 Then
						lstOffer.AddItem(GetOfferDescription(lID, iTypeID, blQty, lExtID, glPlayerID))
						lIdx = lstOffer.NewIndex
						lstOffer.ItemData(lIdx) = lID
						lstOffer.ItemData2(lIdx) = iTypeID
						lstOffer.ItemData3(lIdx) = lExtID
						ReDim Preserve mblQtys(lstOffer.ListCount - 1)
                        mblQtys(lIdx) = 0
                        lstTradeables.ItemBold(lstTradeables.ListIndex) = True
					Else
						'Ensure we have enough qty
						If ItemStackable(CShort(lstTradeables.ItemData2(lstTradeables.ListIndex))) = False Then
							goUILib.AddNotification("That item is already in the offer list.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
							Return
						End If
					End If

					mblQtys(lIdx) += blQty
					Dim sValue As String = GetOfferDescription(lID, iTypeID, mblQtys(lIdx), lExtID, glPlayerID)
					If lstOffer.List(lIdx) <> sValue Then lstOffer.List(lIdx) = sValue
				Else : goUILib.AddNotification("Please enter a numeric value greater than 0 and try again.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				End If
			Else : goUILib.AddNotification("Please enter a numeric value greater than 0 and try again.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			End If
		Else : goUILib.AddNotification("Please select an item in the tradeables list and try again.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		End If
	End Sub

	Private Sub btnReject_Click(ByVal sName As String) Handles btnReject.Click
		If HasAliasedRights(AliasingRights.eAlterTrades) = False Then
			MyBase.moUILib.AddNotification("You lack rights to alter/place trades.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If
		If moTrade Is Nothing Then
			MyBase.moUILib.RemoveWindow(Me.ControlName)
			Return
		End If

		If btnReject.Caption.ToLower = "reject" Then
			btnReject.Caption = "Confirm"
			goUILib.AddNotification("Press Confirm to Reject the Trade. This action cannot be undone.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Alert.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
		Else
			'reject the trade
			CreateAndSendSubmitTradeMessage(True)
			btnReject.Enabled = False
			btnSubmit.Enabled = False
			btnAdd.Enabled = False
			txtQuantity.Enabled = False
			txtNotesToMe.Enabled = False
			txtNotesToThem.Enabled = False
			btnRemove.Enabled = False
			chkAccept.Enabled = False
			chkTheirAccept.Enabled = False
		End If
	End Sub

	Private Sub btnRemove_Click(ByVal sName As String) Handles btnRemove.Click
		If HasAliasedRights(AliasingRights.eAlterTrades) = False Then
			MyBase.moUILib.AddNotification("You lack rights to alter/place trades.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If
        If lstOffer.ListCount > 0 AndAlso lstOffer.ListIndex > -1 Then
            For X As Int32 = lstOffer.ListIndex To mblQtys.GetUpperBound(0) - 1
                mblQtys(X) = mblQtys(X + 1)
            Next X
            ReDim Preserve mblQtys(mblQtys.GetUpperBound(0) - 1)
            lstOffer.RemoveItem(lstOffer.ListIndex)
            If lstOffer.ListCount = 0 Then
                lstOffer.ListIndex = -1
            ElseIf lstOffer.ListIndex >= lstOffer.ListCount Then
                lstOffer.ListIndex = lstOffer.ListIndex - 1
            End If
        Else
            goUILib.AddNotification("Please select an item in the list of items you are offering to remove and try again.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

	Private Sub btnSubmit_Click(ByVal sName As String) Handles btnSubmit.Click
		If HasAliasedRights(AliasingRights.eAlterTrades) = False Then
			MyBase.moUILib.AddNotification("You lack rights to alter/place trades.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If
		CreateAndSendSubmitTradeMessage(False)
	End Sub

    Private Sub lstGetting_ItemClick(ByVal lIndex As Integer) Handles lstGetting.ItemClick
        With lstGetting
			If .ListIndex > -1 Then
				If moTrade Is Nothing = False Then
					If moTrade.lPlayer1ID = glPlayerID Then
						ShowItemDetails(.ItemData(.ListIndex), CShort(.ItemData2(.ListIndex)), .ItemData3(.ListIndex), moTrade.lPlayer2ID)
					Else
						ShowItemDetails(.ItemData(.ListIndex), CShort(.ItemData2(.ListIndex)), .ItemData3(.ListIndex), moTrade.lPlayer1ID)
					End If
				End If
			End If
        End With
    End Sub

    Private Sub lstOffer_ItemClick(ByVal lIndex As Integer) Handles lstOffer.ItemClick
        With lstOffer
            If .ListIndex > -1 Then
				ShowItemDetails(.ItemData(.ListIndex), CShort(.ItemData2(.ListIndex)), .ItemData3(.ListIndex), glPlayerID)
            End If
        End With
    End Sub

    Private Sub lstTradables_ItemClick(ByVal lIndex As Integer) Handles lstTradeables.ItemClick
        With lstTradeables
            If .ListIndex > -1 Then
				ShowItemDetails(.ItemData(.ListIndex), CShort(.ItemData2(.ListIndex)), .ItemData3(.ListIndex), glPlayerID)
				If .ItemData2(.ListIndex) = ObjectType.eUnit OrElse .ItemData2(.ListIndex) = ObjectType.eFacility OrElse .ItemData2(.ListIndex) = ObjectType.eAgent OrElse .ItemData2(.ListIndex) = ObjectType.ePlayerIntel OrElse .ItemData2(.ListIndex) = ObjectType.ePlayerItemIntel OrElse .ItemData2(.ListIndex) = ObjectType.ePlayerTechKnowledge Then
					txtQuantity.Enabled = False
					txtQuantity.Caption = "1"
				Else
                    txtQuantity.Enabled = True
                    If .ItemBold(.ListIndex) = False Then
                        Dim sParts() As String = Split(.List(.ListIndex), "(")
                        If sParts Is Nothing = False AndAlso sParts.GetUpperBound(0) >= 1 Then
                            Dim sQty() As String = Split(sParts(1), ")")
                            If sQty Is Nothing = False AndAlso sQty.GetUpperBound(0) >= 1 Then
                                txtQuantity.Caption = sQty(0).Replace(",", "")
                            End If
                        End If
                    Else
                        txtQuantity.Caption = "0"
                    End If
                End If
            End If
        End With
    End Sub

    Private Sub lstTradables_ItemDblClick(ByVal lIndex As Integer) Handles lstTradeables.ItemDblClick
		btnAdd_Click(btnAdd.ControlName)
    End Sub

    Private Sub lstTradeables_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles lstTradeables.OnKeyDown
        If e.KeyCode = Keys.Enter Then
            btnAdd_Click(btnAdd.ControlName)
        End If
    End Sub

	Private Function GetOfferDescription(ByVal lObjID As Int32, ByVal iObjTypeID As Int16, ByVal blQty As Int64, ByVal lExtID As Int32, ByVal lPlayerID As Int32) As String
		Dim sValue As String = ""
		Select Case Math.Abs(iObjTypeID)
			Case ObjectType.eColony
				sValue = "Entire Colony"
			Case ObjectType.eColonists
				sValue = "Colonists" ' (" & lQty & ")"
			Case ObjectType.eEnlisted
				sValue = "Enlisted"	' (" & lQty & ")"
			Case ObjectType.eOfficers
				sValue = "Officers"	' (" & lQty & ")"
			Case ObjectType.eAgent
				For X As Int32 = 0 To goCurrentPlayer.AgentUB
					If goCurrentPlayer.AgentIdx(X) = lObjID Then
						sValue = goCurrentPlayer.Agents(X).sAgentName
						Exit For
					End If
				Next X
			Case ObjectType.eCredits
				sValue = "Credits" ' (" & lQty.ToString("#,###") & ")"
			Case ObjectType.eFood
				sValue = "Food"	' (" & lQty & ")"
			Case ObjectType.eAmmunition
				Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lObjID, ObjectType.eWeaponTech)
				Dim sTemp As String = ""
				If oTech Is Nothing = False Then
					sTemp = oTech.GetComponentName
				Else : sTemp = GetCacheObjectValue(lObjID, ObjectType.eWeaponTech)
				End If
				sValue = "Ammunition (" & sTemp & ")"
			Case ObjectType.eStock
				'TODO: Define this
			Case ObjectType.ePlayerIntel
				sValue = GetCacheObjectValue(lObjID, ObjectType.ePlayer) & " Scores"
			Case ObjectType.ePlayerItemIntel, ObjectType.ePlayerTechKnowledge
				sValue = GetCacheObjectValue(lObjID, CShort(lExtID)) & " Intel"
			Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech
				If iObjTypeID < 0 Then
					Dim oTPC As TradePostContents = Nothing
					For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
						If TradePostContents.lTradePostContentsIdx(X) = mlTradePostID Then
							oTPC = TradePostContents.oTradePostContents(X)
							Exit For
						End If
					Next X
					sValue = "Unknown Components"
					If oTPC Is Nothing = False Then
						Dim lCompID As Int32 = -1
						Dim lOwnerID As Int32 = -1
						oTPC.PopulateComponentCacheProperties(lObjID, lCompID, iObjTypeID, lOwnerID)
						If lOwnerID <> glPlayerID Then
							sValue = GetCacheObjectValue(lCompID, iObjTypeID) & " (" & Base_Tech.GetComponentTypeName(iObjTypeID) & ")"
						Else
							Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lCompID, iObjTypeID)
							If oTech Is Nothing = False Then
								sValue = oTech.GetComponentName & " (" & Base_Tech.GetComponentTypeName(iObjTypeID) & ")"
							End If
						End If
					End If
				Else
					Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lObjID, Math.Abs(iObjTypeID))
					If oTech Is Nothing = False Then
						sValue = oTech.GetComponentName & " (" & Base_Tech.GetComponentTypeName(Math.Abs(iObjTypeID)) & ")"
					Else
						If iObjTypeID < 0 Then
							sValue = "Unknown Components"
						Else : sValue = "Unknown " & Base_Tech.GetComponentTypeName(iObjTypeID) & " Design"
						End If
					End If
				End If
			Case ObjectType.eMineral ', ObjectType.eMineralCache
				sValue = "Unknown Minerals"
				For X As Int32 = 0 To glMineralUB
					If glMineralIdx(X) = lObjID Then
						sValue = goMinerals(X).MineralName
						Exit For
					End If
				Next X
				'sValue &= " (" & lQty & ")"
			Case ObjectType.eMineralCache
				If lExtID > -1 Then
					sValue = "Unknown Minerals"
					For X As Int32 = 0 To glMineralUB
						If glMineralIdx(X) = lExtID Then
							sValue = goMinerals(X).MineralName
							Exit For
						End If
					Next X
				End If
			Case Else
				sValue = GetCacheObjectValue(lObjID, iObjTypeID)
		End Select

		If blQty > 1 Then sValue &= " (" & blQty.ToString("#,###") & ")"
		Return sValue
	End Function

	Private Function GetTheirOfferDescription(ByVal lObjID As Int32, ByVal iObjTypeID As Int16, ByVal blQty As Int64, ByVal lExtID As Int32, ByVal lPlayerID As Int32) As String
		Dim sValue As String = ""
		Select Case Math.Abs(iObjTypeID)
			Case ObjectType.eColony
				sValue = "Entire Colony"
				sValue = GetCacheObjectValue(lObjID, iObjTypeID)
			Case ObjectType.eColonists
				sValue = "Colonists" ' (" & lQty & ")"
			Case ObjectType.eEnlisted
				sValue = "Enlisted"	' (" & lQty & ")"
			Case ObjectType.eOfficers
				sValue = "Officers"	' (" & lQty & ")"
			Case ObjectType.eAgent
				sValue = GetCacheObjectValue(lObjID, iObjTypeID)
			Case ObjectType.eCredits
				sValue = "Credits" ' (" & lQty.ToString("#,###") & ")"
			Case ObjectType.eFood
				sValue = "Food"	' (" & lQty & ")"
			Case ObjectType.eAmmunition
				Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lObjID, ObjectType.eWeaponTech)
				Dim sTemp As String = ""
				If oTech Is Nothing = False Then
					sTemp = oTech.GetComponentName
				Else : sTemp = GetCacheObjectValue(lObjID, ObjectType.eWeaponTech)
				End If
				sValue = "Ammunition (" & sTemp & ")"
			Case ObjectType.eStock
				'TODO: Define this
			Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech
				If iObjTypeID < 0 Then
					sValue = GetCacheObjectValue(lObjID, ObjectType.eComponentCache)
				Else : sValue = GetCacheObjectValue(lObjID, Math.Abs(iObjTypeID))
				End If
			Case ObjectType.ePlayerIntel, ObjectType.ePlayerItemIntel, ObjectType.ePlayerTechKnowledge
				sValue = GetNonOwnerIntelItemData(-lPlayerID, lObjID, iObjTypeID, CShort(lExtID))
			Case ObjectType.eMineral ', ObjectType.eMineralCache
				sValue = "Unknown Minerals"
				For X As Int32 = 0 To glMineralUB
					If glMineralIdx(X) = lObjID Then
						sValue = goMinerals(X).MineralName
						Exit For
					End If
				Next X
				'sValue &= " (" & lQty & ")"
			Case ObjectType.ePlayerIntel
				sValue = GetCacheObjectValue(lObjID, ObjectType.ePlayer) & " Scores"
			Case ObjectType.ePlayerItemIntel, ObjectType.ePlayerTechKnowledge
				sValue = GetCacheObjectValue(lObjID, CShort(lExtID)) & " Intel"
			Case ObjectType.eMineralCache
				If lExtID > -1 Then
					sValue = "Unknown Minerals"
					For X As Int32 = 0 To glMineralUB
						If glMineralIdx(X) = lExtID Then
							sValue = goMinerals(X).MineralName
							Exit For
						End If
					Next X
				End If
			Case Else
				sValue = GetCacheObjectValue(lObjID, iObjTypeID)
		End Select

		If blQty > 1 Then sValue &= " (" & blQty.ToString("#,###") & ")"
		Return sValue
	End Function

	'Private Sub EnsureTradeablesContains(ByVal lObjID As Int32, ByVal iObjTypeID As Int16, ByVal blQty As Int64)
	'    Dim sValue As String

	'    Dim blUsedQty As Int64 = 0

	'    For X As Int32 = 0 To lstOffer.ListCount - 1
	'        If lstOffer.ItemData(X) = lObjID AndAlso lstOffer.ItemData2(X) = iObjTypeID Then
	'            blUsedQty = mblQtys(X)
	'            Exit For
	'        End If
	'    Next X

	'    Select Case Math.Abs(iObjTypeID)
	'        Case ObjectType.eColony
	'            sValue = "Entire Colony"
	'        Case ObjectType.eColonists
	'            sValue = "Colonists (" & (blQty - blUsedQty) & ")"
	'        Case ObjectType.eEnlisted
	'            sValue = "Enlisted (" & (blQty - blUsedQty) & ")"
	'        Case ObjectType.eOfficers
	'            sValue = "Officers (" & (blQty - blUsedQty) & ")"
	'        Case ObjectType.eAgent
	'            sValue = "Unknown Agent"
	'            sValue = GetCacheObjectValue(lObjID, iObjTypeID)
	'        Case ObjectType.eCredits
	'            sValue = "Credits (" & (goCurrentPlayer.blCredits - blUsedQty).ToString("#,###") & ")"
	'        Case ObjectType.eFood
	'            sValue = "Food (" & (blQty - blUsedQty) & ")"
	'        Case ObjectType.eAmmunition
	'            sValue = "Ammunition (" & (blQty - blUsedQty) & ")"
	'        Case ObjectType.eStock
	'            sValue = "Unknown Stock"
	'            sValue = GetCacheObjectValue(lObjID, iObjTypeID)
	'        Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech
	'            If iObjTypeID < 0 Then
	'                'represents a component cache
	'                sValue = "Unknown Components"
	'                sValue = GetCacheObjectValue(lObjID, Math.Abs(iObjTypeID)) & " (" & (blQty - blUsedQty) & ")"
	'            Else
	'                'represents the tech design
	'                sValue = "Unknown " & Base_Tech.GetComponentTypeName(iObjTypeID) & " Design"
	'                For X As Int32 = 0 To goCurrentPlayer.mlTechUB
	'                    If goCurrentPlayer.moTechs(X).ObjectID = lObjID AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = iObjTypeID Then
	'                        sValue = goCurrentPlayer.moTechs(X).GetComponentName() & " (" & Base_Tech.GetComponentTypeName(iObjTypeID) & " Design)"
	'                        Exit For
	'                    End If
	'                Next X
	'            End If
	'        Case ObjectType.eMineral, ObjectType.eMineralCache
	'            'represents a mineral cache... ID should be mineral
	'            sValue = "Unknown Minerals"
	'            For X As Int32 = 0 To glMineralUB
	'                If glMineralIdx(X) = lObjID Then
	'                    sValue = goMinerals(X).MineralName
	'                    Exit For
	'                End If
	'            Next X
	'            sValue &= " (" & (blQty - blUsedQty) & ")"
	'        Case Else
	'            sValue = GetCacheObjectValue(lObjID, iObjTypeID)
	'    End Select


	'    For X As Int32 = 0 To lstTradeables.ListCount - 1
	'        If lstTradeables.ItemData(X) = lObjID AndAlso lstTradeables.ItemData2(X) = iObjTypeID Then
	'            If lstTradeables.List(X) <> sValue Then lstTradeables.List(X) = sValue
	'            Return
	'        End If
	'    Next X

	'    'If we are here, it does not contain it...
	'    lstTradeables.AddItem(sValue)
	'    lstTradeables.ItemData(lstTradeables.NewIndex) = lObjID
	'    lstTradeables.ItemData2(lstTradeables.NewIndex) = iObjTypeID
	'End Sub

    Private Sub frmTrade_OnNewFrame() Handles Me.OnNewFrame
        If glCurrentCycle - mlLastUpdate > 15 Then
            mlLastUpdate = glCurrentCycle

            UpdateTaxLabel()

            For X As Int32 = 0 To cboSellToPlayer.ListCount - 1
                Dim sTemp As String = GetCacheObjectValue(cboSellToPlayer.ItemData(X), ObjectType.ePlayer)
                If cboSellToPlayer.List(X) <> sTemp Then cboSellToPlayer.List(X) = sTemp
            Next X

            If moTrade Is Nothing = False Then
                If moTrade.lPlayer1ID = glPlayerID Then
                    If txtNotesToMe.Caption <> moTrade.sP2Notes Then txtNotesToMe.Caption = moTrade.sP2Notes
                Else
                    If txtNotesToMe.Caption <> moTrade.sP1Notes Then txtNotesToMe.Caption = moTrade.sP1Notes
                End If
            End If

            'Check Tradeables listbox (including qtys)
            If mlTradePostID > -1 Then
                'Yes, ok, find our TradePostContents
                Dim oTPC As TradePostContents = Nothing
                For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
                    If TradePostContents.lTradePostContentsIdx(X) = mlTradePostID Then
                        oTPC = TradePostContents.oTradePostContents(X)
                        Exit For
                    End If
                Next X

                If oTPC Is Nothing = False Then
                    'Ok... ensure the list has everything
                    If oTPC.lLastUpdate <> mlTPCLast Then
                        oTPC.PopulateList(lstTradeables, 1)
                        mlTPCLast = oTPC.lLastUpdate
                    Else : oTPC.SmartPopulateList(lstTradeables, 1)
                    End If
                Else
                    If lstTradeables.ListCount > 0 Then
                        lstTradeables.Clear()
                        lstOffer.Clear()
                    End If
                End If
                'If moTrade Is Nothing = False Then
                For x As Int32 = 0 To lstTradeables.ListCount - 1
                    'If lstTradeables.ItemBold(x) = True Then
                    Dim lIdx As Int32 = -1
                    Dim lID As Int32 = lstTradeables.ItemData(x)
                    Dim iTypeID As Int16 = CShort(lstTradeables.ItemData2(x))
                    Dim lExtID As Int32 = lstTradeables.ItemData3(x)

                    For y As Int32 = 0 To lstOffer.ListCount - 1
                        If lstOffer.ItemData(y) = lID AndAlso lstOffer.ItemData2(y) = iTypeID AndAlso lstOffer.ItemData3(y) = lExtID Then
                            lIdx = x
                            lstTradeables.ItemBold(x) = True
                            Exit For
                        End If
                    Next y
                    If lIdx = -1 Then
                        lstTradeables.ItemBold(x) = False
                    End If
                    'End If
                Next x
                'End If
            ElseIf lstTradeables.ListCount > 0 Then
                lstTradeables.Clear()
                lstOffer.Clear()
            End If

            If mlCurrentItemDetail <> -1 AndAlso miCurrentItemDetailType <> -1 Then
                Dim lTempID As Int32 = mlCurrentItemDetail
                Dim iTypeID As Int16 = miCurrentItemDetailType
                If miCurrentItemDetailType < 0 Then
                    Dim oTPC As TradePostContents = Nothing
                    For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
                        If TradePostContents.lTradePostContentsIdx(X) = mlTradePostID Then
                            oTPC = TradePostContents.oTradePostContents(X)
                            Exit For
                        End If
                    Next X
                    If oTPC Is Nothing = False Then
                        Dim lCompOwnerID As Int32 = -1
                        oTPC.PopulateComponentCacheProperties(mlCurrentItemDetail, lTempID, iTypeID, lCompOwnerID)
                    End If
                End If
                Dim sTemp As String = ""
                If iTypeID < 0 Then
                    sTemp = GetNonOwnerItemData(lTempID, ObjectType.eComponentCache)
                ElseIf iTypeID = ObjectType.ePlayerIntel OrElse iTypeID = ObjectType.ePlayerItemIntel OrElse iTypeID = ObjectType.ePlayerTechKnowledge Then
                    sTemp = GetNonOwnerIntelItemData(-mlCurrentItemOwnerID, mlCurrentItemDetail, miCurrentItemDetailType, CShort(mlCurrentItemExtended))
                Else
                    sTemp = GetNonOwnerItemData(lTempID, iTypeID)
                End If
                If txtItemDetails.Caption <> sTemp Then txtItemDetails.Caption = sTemp
            ElseIf txtItemDetails.Caption <> "" Then
                txtItemDetails.Caption = ""
            End If

            'Check lstGetting
            If moTrade Is Nothing = False Then
                With moTrade

                    Dim bNewVal As Boolean = (.yTradeState And Trade.eTradeStateValues.Player1Accepted) <> 0
                    If .lPlayer1ID <> glPlayerID AndAlso chkTheirAccept.Value <> bNewVal Then
                        chkTheirAccept.Value = bNewVal
                    End If
                    bNewVal = (.yTradeState And Trade.eTradeStateValues.Player2Accepted) <> 0
                    If .lPlayer2ID <> glPlayerID AndAlso chkTheirAccept.Value <> bNewVal Then
                        chkTheirAccept.Value = bNewVal
                    End If

                    If .lPlayer1ID = glPlayerID Then
                        For X As Int32 = 0 To .mlPlayer2ItemUB
                            Dim lID As Int32 = .muPlayer2Items(X).ObjectID
                            Dim iTypeID As Int16 = .muPlayer2Items(X).ObjTypeID
                            Dim blQty As Int64 = .muPlayer2Items(X).Quantity
                            Dim lExtID As Int32 = .muPlayer2Items(X).lExtendedID

                            'Dim sValue As String =  GetCacheObjectValue(lID, iTypeID)
                            'If lQty > 1 Then sValue &= " (" & lQty & ")"
                            Dim sValue As String = GetTheirOfferDescription(lID, iTypeID, blQty, lExtID, .lPlayer2ID)
                            If sValue Is Nothing = False Then
                                Dim lTmpIdx As Int32 = sValue.IndexOf(vbCrLf)
                                If lTmpIdx > -1 Then sValue = sValue.Substring(0, lTmpIdx)
                            End If

                            Dim bGettingFound As Boolean = False
                            For Y As Int32 = 0 To lstGetting.ListCount - 1
                                If lstGetting.ItemData(Y) = lID AndAlso lstGetting.ItemData2(Y) = iTypeID AndAlso lstGetting.ItemData3(Y) = lExtID Then
                                    If lstGetting.List(Y) <> sValue Then lstGetting.List(Y) = sValue
                                    bGettingFound = True
                                    Exit For
                                End If
                            Next Y
                            If bGettingFound = False Then
                                lstGetting.AddItem(sValue)
                                lstGetting.ItemData(lstGetting.NewIndex) = lID
                                lstGetting.ItemData2(lstGetting.NewIndex) = iTypeID
                                lstGetting.ItemData3(lstGetting.NewIndex) = lExtID
                            End If
                        Next X
                    Else
                        For X As Int32 = 0 To .mlPlayer1ItemUB
                            Dim lID As Int32 = .muPlayer1Items(X).ObjectID
                            Dim iTypeID As Int16 = .muPlayer1Items(X).ObjTypeID
                            Dim blQty As Int64 = .muPlayer1Items(X).Quantity
                            Dim lExtID As Int32 = .muPlayer1Items(X).lExtendedID

                            'Dim sValue As String = GetCacheObjectValue(lID, iTypeID)
                            'If lQty > 1 Then sValue &= " (" & lQty & ")"
                            Dim sValue As String = GetTheirOfferDescription(lID, iTypeID, blQty, lExtID, .lPlayer1ID)
                            If sValue Is Nothing = False Then
                                Dim lTmpIdx As Int32 = sValue.IndexOf(vbCrLf)
                                If lTmpIdx > -1 Then sValue = sValue.Substring(0, lTmpIdx)
                            End If

                            Dim bGettingFound As Boolean = False
                            For Y As Int32 = 0 To lstGetting.ListCount - 1
                                If lstGetting.ItemData(Y) = lID AndAlso lstGetting.ItemData2(Y) = iTypeID AndAlso lstGetting.ItemData3(Y) = lExtID Then
                                    If lstGetting.List(Y) <> sValue Then lstGetting.List(Y) = sValue
                                    bGettingFound = True
                                    Exit For
                                End If
                            Next Y
                            If bGettingFound = False Then
                                lstGetting.AddItem(sValue)
                                lstGetting.ItemData(lstGetting.NewIndex) = lID
                                lstGetting.ItemData2(lstGetting.NewIndex) = iTypeID
                                lstGetting.ItemData3(lstGetting.NewIndex) = lExtID
                            End If
                        Next X
                    End If
                End With
            End If
        End If
    End Sub

	Private Sub ShowItemDetails(ByVal lObjectID As Int32, ByVal iObjTypeID As Int16, ByVal lExtID As Int32, ByVal lOwnerID As Int32)
		Select Case Math.Abs(iObjTypeID)
			Case ObjectType.eColonists, ObjectType.eEnlisted, ObjectType.eOfficers, ObjectType.eCredits, ObjectType.eFood, ObjectType.eAmmunition, ObjectType.eColony
				'No details to display, so remove the window
				MyBase.moUILib.RemoveWindow("frmNonOwnerItemDetails")
				MyBase.moUILib.RemoveWindow("frmMinDetail")
			Case ObjectType.eMineral ', ObjectType.eMineralCache
				'If the player knows of the mineral, display the mineral details
				For X As Int32 = 0 To glMineralUB
					If glMineralIdx(X) = lObjectID Then
						If goMinerals(X).bDiscovered = True Then
							Dim ofrmMinDet As New frmMinDetail(goUILib)
							ofrmMinDet.ShowMineralDetail(MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 200, Me.Top, 500, lObjectID)
							ofrmMinDet = Nothing
						End If
						Exit For
					End If
				Next X
			Case ObjectType.ePlayerIntel, ObjectType.ePlayerItemIntel, ObjectType.ePlayerTechKnowledge
				mlCurrentItemDetail = lObjectID
				miCurrentItemDetailType = iObjTypeID
				mlCurrentItemExtended = lExtID
				mlCurrentItemOwnerID = lOwnerID
				txtItemDetails.Caption = GetNonOwnerIntelItemData(-lOwnerID, lObjectID, iObjTypeID, CShort(lExtID))
			Case Else
				'Everything else calls the window
				'Dim ofrmDetails As New frmNonOwnerItemDetails(goUILib)
				'ofrmDetails.SetFromItem(lObjectID, Math.Abs(iObjTypeID))
				'ofrmDetails = Nothing
				mlCurrentItemDetail = lObjectID
				miCurrentItemDetailType = iObjTypeID

				Dim lTempID As Int32 = mlCurrentItemDetail
				Dim iTypeID As Int16 = miCurrentItemDetailType
				If miCurrentItemDetailType < 0 Then
					Dim oTPC As TradePostContents = Nothing
					For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
						If TradePostContents.lTradePostContentsIdx(X) = mlTradePostID Then
							oTPC = TradePostContents.oTradePostContents(X)
							Exit For
						End If
					Next X
					If oTPC Is Nothing = False Then
						Dim lCompOwnerID As Int32 = -1
						oTPC.PopulateComponentCacheProperties(mlCurrentItemDetail, lTempID, iTypeID, lCompOwnerID)
					End If
				End If
				If iTypeID < 0 Then
					txtItemDetails.Caption = GetNonOwnerItemData(lTempID, ObjectType.eComponentCache)
				Else : txtItemDetails.Caption = GetNonOwnerItemData(lTempID, iTypeID)
				End If
		End Select

	End Sub

    Private Sub chkTheirAccept_Click() Handles chkTheirAccept.Click
        If moTrade Is Nothing = False Then
            If moTrade.lPlayer1ID = glPlayerID Then
                chkTheirAccept.Value = (moTrade.yTradeState And Trade.eTradeStateValues.Player2Accepted) <> 0
            Else : chkTheirAccept.Value = (moTrade.yTradeState And Trade.eTradeStateValues.Player1Accepted) <> 0
            End If
        Else : chkTheirAccept.Value = False
        End If
    End Sub

    Private Function ItemStackable(ByVal iTypeID As Int16) As Boolean
        Select Case iTypeID
			Case ObjectType.eAgent, ObjectType.eColony, ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, _
			  ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, _
			  ObjectType.eUnit, ObjectType.eFacility, ObjectType.ePlayerIntel, ObjectType.ePlayerItemIntel, ObjectType.ePlayerTechKnowledge
				Return False
            Case Else
                Return True
        End Select
    End Function

    Private Sub cboSellToPlayer_ItemSelected(ByVal lItemIndex As Integer) Handles cboSellToPlayer.ItemSelected
        If lItemIndex <> -1 Then
            Dim sTemp As String = cboSellToPlayer.List(lItemIndex) & "'s Acceptance"
            chkTheirAccept.Caption = sTemp
            Dim rcTemp As Rectangle = chkTheirAccept.GetTextDimensions()
            chkTheirAccept.Width = rcTemp.Width + 20
        End If
    End Sub
End Class