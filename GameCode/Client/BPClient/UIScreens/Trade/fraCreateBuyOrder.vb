Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class fraCreateBuyOrder
    Inherits UIWindow

    Private lblItemType As UILabel
    Private lblRequires As UILabel
    Private lblDetails As UILabel
    Private lnDiv1 As UILine
    Private lblEscrow As UILabel
    Private lblPayment As UILabel
    Private lblQuantity As UILabel
    Private lblDeadline As UILabel
    Private lblDays As UILabel
    Private lblHours As UILabel
    Private lnDiv2 As UILine
    Private lblSpecificMineral As UILabel

    Private WithEvents lstItemType As UIListBox
    Private WithEvents txtEscrow As UITextBox
    Private WithEvents txtPayment As UITextBox
    Private WithEvents txtQuantity As UITextBox
    Private WithEvents txtDays As UITextBox
    Private WithEvents txtHours As UITextBox
    Private WithEvents btnClose As UIButton
    Private WithEvents btnPlaceOrder As UIButton

    Private fraRequires As fraBuyOrderRequires

    Private cboSpecificMineral As UIComboBox

    Private mlTradePostID As Int32 = -1

    Private mbLoading As Boolean = True

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'fraCreateBuyOrder initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eCreateBuyOrder
            .ControlName = "fraCreateBuyOrder"
            .Left = 53
            .Top = 69
            .Width = 790
            .Height = 543
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
            .mbAcceptReprocessEvents = True
            .BorderLineWidth = 2
        End With

        'lblItemType initial props
        lblItemType = New UILabel(oUILib)
        With lblItemType
            .ControlName = "lblItemType"
            .Left = 5
            .Top = 5
            .Width = 76
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Item Type"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblItemType, UIControl))

        'lstItemType initial props
        lstItemType = New UIListBox(oUILib)
        With lstItemType
            .ControlName = "lstItemType"
            .Left = 5
            .Top = 25
            .Width = 250
            .Height = 512
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(lstItemType, UIControl))

        'lblRequires initial props
        lblRequires = New UILabel(oUILib)
        With lblRequires
            .ControlName = "lblRequires"
            .Left = 530
            .Top = 5
            .Width = 169
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Buy Order Requirements"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblRequires, UIControl))

        'lblDetails initial props
        lblDetails = New UILabel(oUILib)
        With lblDetails
            .ControlName = "lblDetails"
            .Left = 265
            .Top = 5
            .Width = 123
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Buy Order Details"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDetails, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 265
            .Top = 25
            .Width = 255
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lblEscrow initial props
        lblEscrow = New UILabel(oUILib)
        With lblEscrow
            .ControlName = "lblEscrow"
            .Left = 270
            .Top = 30
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Escrow:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblEscrow, UIControl))

        'txtEscrow initial props
        txtEscrow = New UITextBox(oUILib)
        With txtEscrow
            .ControlName = "txtEscrow"
            .Left = 370
            .Top = 30
            .Width = 145
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
            .MaxLength = 18
			.BorderColor = muSettings.InterfaceBorderColor
			.bNumericOnly = True
			.ToolTipText = "The amount of credits the Acceptor of this buy" & vbCrLf & _
			  "order places up front when they accept the work." & vbCrLf & _
			  "This amount is refunded to the Acceptor upon" & vbCrLf & _
			  "successful completion of this order or it is" & vbCrLf & _
			  "paid to you upon failure to complete this order."
        End With
        Me.AddChild(CType(txtEscrow, UIControl))

        'lblPayment initial props
        lblPayment = New UILabel(oUILib)
        With lblPayment
            .ControlName = "lblPayment"
            .Left = 270
            .Top = 55
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Payment:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblPayment, UIControl))

        'txtPayment initial props
        txtPayment = New UITextBox(oUILib)
        With txtPayment
            .ControlName = "txtPayment"
            .Left = 370
            .Top = 55
            .Width = 145
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
            .MaxLength = 18
			.BorderColor = muSettings.InterfaceBorderColor
			.bNumericOnly = True
			.ToolTipText = "The amount of credits you will pay for this buy order." & vbCrLf & _
			 "This amount is deducted from your account when you place" & vbCrLf & _
			 "this buy order. The amount is transferred to the Acceptor" & vbCrLf & _
			 "when they successfully complete the order or it is refunded" & vbCrLf & _
			 "to you if they fail to complete the order."
        End With
        Me.AddChild(CType(txtPayment, UIControl))

        'lblQuantity initial props
        lblQuantity = New UILabel(oUILib)
        With lblQuantity
            .ControlName = "lblQuantity"
            .Left = 270
            .Top = 80
            .Width = 100
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
            .Left = 370
            .Top = 80
            .Width = 145
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
            .MaxLength = 18
			.BorderColor = muSettings.InterfaceBorderColor
			.bNumericOnly = True
			.ToolTipText = "The number of items required to fill the Buy Order."
        End With
        Me.AddChild(CType(txtQuantity, UIControl))

        'lblDeadline initial props
        lblDeadline = New UILabel(oUILib)
        With lblDeadline
            .ControlName = "lblDeadline"
            .Left = 270
            .Top = 105
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Deadline:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDeadline, UIControl))

        'lblDays initial props
        lblDays = New UILabel(oUILib)
        With lblDays
            .ControlName = "lblDays"
            .Left = 420
            .Top = 105
            .Width = 40
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Days"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDays, UIControl))

        'lblHours initial props
        lblHours = New UILabel(oUILib)
        With lblHours
            .ControlName = "lblHours"
            .Left = 420
            .Top = 130
            .Width = 43
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hours"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblHours, UIControl))

        'txtDays initial props
        txtDays = New UITextBox(oUILib)
        With txtDays
            .ControlName = "txtDays"
            .Left = 370
            .Top = 105
            .Width = 45
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
            .MaxLength = 3
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtDays, UIControl))

        'txtHours initial props
        txtHours = New UITextBox(oUILib)
        With txtHours
            .ControlName = "txtHours"
            .Left = 370
            .Top = 130
            .Width = 45
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
            .MaxLength = 3
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtHours, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 335
            .Top = 510
            .Width = 110
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Close"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'btnPlaceOrder initial props
        btnPlaceOrder = New UIButton(oUILib)
        With btnPlaceOrder
            .ControlName = "btnPlaceOrder"
            .Left = 335
            .Top = 480
            .Width = 110
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Place Order"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnPlaceOrder, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = 265
            .Top = 160
            .Width = 255
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        'fraRequires initial props
        fraRequires = New fraBuyOrderRequires(oUILib)
        With fraRequires
            .Left = 530
            .Top = 28
            .Width = 256
            .Height = 512
            .Moveable = False
            .BorderLineWidth = 2
        End With
        Me.AddChild(CType(fraRequires, UIControl))

        'lblHours initial props
        lblSpecificMineral = New UILabel(oUILib)
        With lblSpecificMineral
            .ControlName = "lblSpecificMineral"
            .Left = fraRequires.Left
            .Top = fraRequires.Top
            .Width = 80
            .Height = 18
            .Enabled = True
            .Visible = False
            .Caption = "Mineral:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSpecificMineral, UIControl))

        cboSpecificMineral = New UIComboBox(oUILib)
        With cboSpecificMineral
            .ControlName = "cboSpecificMineral"
            .Left = lblSpecificMineral.Left + lblSpecificMineral.Width + 5
            .Top = lblSpecificMineral.Top
            .Width = 155
            .Height = 18
            .Enabled = True
            .Visible = False
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .ToolTipText = "Select the specific mineral you wish to buy."
        End With
        Me.AddChild(CType(cboSpecificMineral, UIControl))

        FillItemTypeList()

        mbLoading = False

        For X As Int32 = 0 To lstItemType.ListCount - 1
            If lstItemType.ItemData(X) = MarketListType.eArmorComponent Then
                lstItemType.ListIndex = X
                Exit For
            End If
        Next X

    End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		CType(Me.ParentControl, frmTradeMain).ReturnToCurrentView()
	End Sub

	Private Sub btnPlaceOrder_Click(ByVal sName As String) Handles btnPlaceOrder.Click
		If HasAliasedRights(AliasingRights.eAlterTrades) = False Then
			MyBase.moUILib.AddNotification("You lack rights to alter/place trades.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If
		Try

			Dim oTPC As TradePostContents = Nothing
			For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
				If TradePostContents.lTradePostContentsIdx(X) = mlTradePostID Then
					oTPC = TradePostContents.oTradePostContents(X)
					Exit For
				End If
			Next X
			If oTPC Is Nothing = False AndAlso oTPC.yBuySlotsUsed >= goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBuyOrderSlots) Then
				MyBase.moUILib.AddNotification("You have no Buy Slots available to place a new buy order.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If

			If lstItemType.ListIndex = -1 Then
				MyBase.moUILib.AddNotification("Select an Item Type first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
				Return
			End If
			Dim yItemType As Byte = CByte(lstItemType.ItemData(lstItemType.ListIndex))
			Dim lDays As Int32 = CInt(Val(txtDays.Caption))
			Dim lHours As Int32 = CInt(Val(txtHours.Caption))
			Dim blEscrow As Int64 = CLng(Val(txtEscrow.Caption))
			Dim blPayment As Int64 = CLng(Val(txtPayment.Caption))
			Dim blQty As Int64 = CLng(Val(txtQuantity.Caption))

			If lDays = 0 AndAlso lHours = 0 Then
				MyBase.moUILib.AddNotification("Enter a deadline for this buy order.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
				Return
			End If
			If blEscrow = 0 Then
				MyBase.moUILib.AddNotification("Escrow amount of 0 indicates no liability on the Acceptor.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
				Return
			End If
			If blPayment = 0 Then
				MyBase.moUILib.AddNotification("Enter a payment amount for this buy order.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
				Return
			End If
			If blQty = 0 Then
				MyBase.moUILib.AddNotification("Enter a quantity for this buy order.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
				Return
			End If

            If yItemType = MarketListType.eBuyOrderSpecificMineral Then
                If cboSpecificMineral.ListIndex < 0 Then
                    MyBase.moUILib.AddNotification("You must select the specific mineral for the buy order.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                    Return
                End If
            Else
                If fraRequires.ValidateData() = False Then
                    MyBase.moUILib.AddNotification("Please enter valid requirements for this buy order.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                    Return
                End If
            End If

			If btnPlaceOrder.Caption.ToLower = "confirm" Then
				'Place the order
                Dim yProps() As Byte = Nothing
                If yItemType = MarketListType.eBuyOrderSpecificMineral Then
                    yProps = System.BitConverter.GetBytes(cboSpecificMineral.ItemData(cboSpecificMineral.ListIndex))
                Else
                    yProps = fraRequires.GetPropertyList()
                End If
				Dim yMsg(38 + yProps.Length) As Byte
				Dim lPos As Int32 = 0
				System.BitConverter.GetBytes(GlobalMessageCode.eSubmitBuyOrder).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(mlTradePostID).CopyTo(yMsg, lPos) : lPos += 4
				yMsg(lPos) = yItemType : lPos += 1
				System.BitConverter.GetBytes(lDays).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(lHours).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(blEscrow).CopyTo(yMsg, lPos) : lPos += 8
				System.BitConverter.GetBytes(blPayment).CopyTo(yMsg, lPos) : lPos += 8
				System.BitConverter.GetBytes(blQty).CopyTo(yMsg, lPos) : lPos += 8
				yProps.CopyTo(yMsg, lPos)
				MyBase.moUILib.SendMsgToPrimary(yMsg)
				'TODO: Now what?
			Else
				btnPlaceOrder.Caption = "Confirm"
				MyBase.moUILib.AddNotification("Confirm Place Order Command.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Alert.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
			End If

		Catch ex As Exception
			MyBase.moUILib.AddNotification("Unable to place order, please check all fields and try again.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		End Try
	End Sub

    Private Sub lstItemType_ItemClick(ByVal lIndex As Integer) Handles lstItemType.ItemClick
        If fraRequires Is Nothing Then Return

        If lIndex <> -1 Then

            If lstItemType.ItemData(lIndex) = MarketListType.eBuyOrderSpecificMineral Then
                fraRequires.Visible = False
                lblSpecificMineral.Visible = True
                cboSpecificMineral.Visible = True

                cboSpecificMineral.Clear()
                Dim lSorted() As Int32 = GetSortedMineralIdxArray(True)
                If lSorted Is Nothing = False Then
                    For X As Int32 = 0 To lSorted.GetUpperBound(0)
                        If lSorted(X) > -1 Then
                            Dim oMineral As Mineral = goMinerals(lSorted(X))
                            If oMineral Is Nothing = False Then
                                cboSpecificMineral.AddItem(oMineral.MineralName)
                                cboSpecificMineral.ItemData(cboSpecificMineral.NewIndex) = oMineral.ObjectID
                            End If
                        End If
                    Next X
                End If
            Else
                cboSpecificMineral.Visible = False
                lblSpecificMineral.Visible = False
                fraRequires.Visible = True
                fraRequires.SetFromItemType(CType(lstItemType.ItemData(lIndex), MarketListType))
            End If
        Else : fraRequires.ClearList()
        End If
    End Sub

    Private Sub TextBoxItem_TextChanged() Handles txtDays.TextChanged, txtEscrow.TextChanged, txtHours.TextChanged, txtPayment.TextChanged, txtQuantity.TextChanged
        If mbLoading = True Then Return
        btnPlaceOrder.Caption = "Place Order"
		If txtDays.Caption <> "" Then
			Dim lTemp As Int32 = CInt(Val(txtDays.Caption))
			If lTemp < 0 Then lTemp = 0
			If lTemp.ToString <> txtDays.Caption Then
				txtDays.Caption = lTemp.ToString
			End If
		End If
        If txtEscrow.Caption <> "" Then
			Dim blTemp As Int64 = CLng((txtEscrow.Caption))
            If blTemp < 0 Then blTemp = 0
            If blTemp.ToString <> txtEscrow.Caption Then
                txtEscrow.Caption = blTemp.ToString
            End If
        End If
        If txtHours.Caption <> "" Then
			Dim lTemp As Int32 = CInt(Val(txtHours.Caption))
			If lTemp < 0 Then lTemp = 0

			If lTemp > 23 Then
				Dim lAddDays As Int32 = lTemp \ 24
				lTemp -= (lAddDays * 24)
				Dim lDays As Int32 = lAddDays + CInt(Val(txtDays.Caption))
				mbLoading = True
				txtDays.Caption = lDays.ToString
				mbLoading = False
			End If

            If lTemp.ToString <> txtHours.Caption Then
                txtHours.Caption = lTemp.ToString
            End If
        End If
        If txtPayment.Caption <> "" Then
			Dim blTemp As Int64 = CLng((txtPayment.Caption))
            If blTemp < 0 Then blTemp = 0
            If blTemp.ToString <> txtPayment.Caption Then
                txtPayment.Caption = blTemp.ToString
            End If
        End If
        If txtQuantity.Caption <> "" Then
			Dim blTemp As Int64 = CLng((txtQuantity.Caption))
            If blTemp < 0 Then blTemp = 0
            If blTemp.ToString <> txtQuantity.Caption Then
                txtQuantity.Caption = blTemp.ToString
            End If
        End If
    End Sub

    Private Sub FillItemTypeList()
        lstItemType.Clear()
        lstItemType.AddItem("COMPONENTS", True) : lstItemType.ItemData(lstItemType.NewIndex) = -1 : lstItemType.ItemLocked(lstItemType.NewIndex) = True
        lstItemType.AddItem("  Armor", False) : lstItemType.ItemData(lstItemType.NewIndex) = MarketListType.eArmorComponent
        lstItemType.AddItem("  Beam (Pulse)", False) : lstItemType.ItemData(lstItemType.NewIndex) = MarketListType.eBeamPulseComponent
        lstItemType.AddItem("  Beam (Solid)", False) : lstItemType.ItemData(lstItemType.NewIndex) = MarketListType.eBeamSolidComponent
        lstItemType.AddItem("  Engine", False) : lstItemType.ItemData(lstItemType.NewIndex) = MarketListType.eEngineComponent
        lstItemType.AddItem("  Missile", False) : lstItemType.ItemData(lstItemType.NewIndex) = MarketListType.eMissileComponent
        lstItemType.AddItem("  Projectile", False) : lstItemType.ItemData(lstItemType.NewIndex) = MarketListType.eProjectileComponent
        lstItemType.AddItem("  Radar", False) : lstItemType.ItemData(lstItemType.NewIndex) = MarketListType.eRadarComponent
        lstItemType.AddItem("  Shield", False) : lstItemType.ItemData(lstItemType.NewIndex) = MarketListType.eShieldComponent
        'lstItemType.AddItem("  Specials", False) : lstItemType.ItemData(lstItemType.NewIndex) = MarketListType.eSpecialComponent
        lstItemType.AddItem("MATERIALS", True) : lstItemType.ItemData(lstItemType.NewIndex) = -1 : lstItemType.ItemLocked(lstItemType.NewIndex) = True
        lstItemType.AddItem("  Minerals", False) : lstItemType.ItemData(lstItemType.NewIndex) = MarketListType.eMinerals
        lstItemType.AddItem("  Alloys", False) : lstItemType.ItemData(lstItemType.NewIndex) = MarketListType.eAlloys
        lstItemType.AddItem("  Specific Mineral", False) : lstItemType.ItemData(lstItemType.NewIndex) = MarketListType.eBuyOrderSpecificMineral
    End Sub

    Public Sub SetTradePostID(ByVal lTPID As Int32)
        mlTradePostID = lTPID
    End Sub
End Class