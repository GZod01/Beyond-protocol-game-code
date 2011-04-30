Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmGuildMain
    'Private Class fra_Assets
    '	Inherits guildframe

    '	Private lblFacLoc As UILabel
    '	Private lblTreasury As UILabel
    '	Private lblAssets As UILabel
    '	Private lblYourAssets As UILabel
    '	Private txtWithdrawQty As UITextBox
    '	Private txtDepositQty As UITextBox
    '	Private lstBankAssets As UIListBox
    '	Private lstYourAssets As UIListBox

    '	Private WithEvents btnGotoFacLoc As UIButton
    '	Private WithEvents btnPlaceFac As UIButton
    '	Private WithEvents btnDeposit As UIButton
    '	Private WithEvents btnWithdraw As UIButton
    '	Private WithEvents btnViewLog As UIButton
    '	Private WithEvents cboTradepost As UIComboBox

    '	Private mbViewLog As Boolean = False
    '	Private mbViewHighSec As Boolean = False
    '	Private mlTPCLast As Int32
    '	Private mlGuildTPCLast As Int32

    '	Public Sub New(ByRef oUILib As UILib)
    '		MyBase.New(oUILib)

    '		With Me
    '			.ControlName = "fra_Assets"
    '			.Left = 15
    '			.Top = 5
    '			.Width = 128
    '			.Height = 128
    '			.Enabled = True
    '			.Visible = True
    '			.BorderColor = muSettings.InterfaceBorderColor
    '			.FillColor = muSettings.InterfaceFillColor
    '			.FullScreen = False
    '			.BorderLineWidth = 2
    '			.Moveable = False
    '		End With

    '		'lblFacLoc initial props
    '		lblFacLoc = New UILabel(MyBase.moUILib)
    '		With lblFacLoc
    '			.ControlName = "lblFacLoc"
    '			.Left = 15
    '			.Top = 5
    '			.Width = 450
    '			.Height = 24
    '			.Enabled = True
    '			.Visible = False
    '			If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
    '				.Visible = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ViewGuildBase)
    '			End If
    '			.Caption = "Guild Facility Location: Facility Not Built/Placed"
    '			.ForeColor = muSettings.InterfaceBorderColor
    '			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '			.DrawBackImage = False
    '			.FontFormat = CType(4, DrawTextFormat)
    '		End With
    '		Me.AddChild(CType(lblFacLoc, UIControl))

    '		'lblFacLoc initial props
    '		lblTreasury = New UILabel(MyBase.moUILib)
    '		With lblTreasury
    '			.ControlName = "lblTreasury"
    '			.Left = 15
    '			.Top = lblFacLoc.Top + lblFacLoc.Height + 5
    '			.Width = 450
    '			.Height = 24
    '			.Enabled = True
    '			.Visible = False
    '			.Caption = ""
    '			If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
    '				.Visible = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ViewContentsLowSec)
    '				.Caption = "Guild Treasury: " & goCurrentPlayer.oGuild.blTreasury.ToString("#,##0")
    '			End If

    '			.ForeColor = muSettings.InterfaceBorderColor
    '			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '			.DrawBackImage = False
    '			.FontFormat = CType(4, DrawTextFormat)
    '		End With
    '		Me.AddChild(CType(lblTreasury, UIControl))

    '		'btnGotoFacLoc initial props
    '		btnGotoFacLoc = New UIButton(MyBase.moUILib)
    '		With btnGotoFacLoc
    '			.ControlName = "btnGotoFacLoc"
    '			.Left = 470
    '			.Top = 5
    '			.Width = 100
    '			.Height = 24
    '			.Enabled = True
    '			.Visible = lblFacLoc.Visible
    '			.Caption = "Goto"
    '			.ForeColor = muSettings.InterfaceBorderColor
    '			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '			.DrawBackImage = True
    '			.FontFormat = CType(5, DrawTextFormat)
    '			.ControlImageRect = New Rectangle(0, 0, 120, 32)
    '		End With
    '		Me.AddChild(CType(btnGotoFacLoc, UIControl))

    '		'btnPlaceFac initial props
    '		btnPlaceFac = New UIButton(MyBase.moUILib)
    '		With btnPlaceFac
    '			.ControlName = "btnPlaceFac"
    '			.Left = 580
    '			.Top = 5
    '			.Width = 100
    '			.Height = 24
    '			.Enabled = True
    '			If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
    '				.Enabled = goCurrentPlayer.oGuild.lGuildHallLocID < 1 OrElse goCurrentPlayer.oGuild.iGuildHallLocTypeID < 1
    '			End If
    '			.Visible = lblFacLoc.Visible
    '			.Caption = "Build"
    '			.ForeColor = muSettings.InterfaceBorderColor
    '			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '			.DrawBackImage = True
    '			.FontFormat = CType(5, DrawTextFormat)
    '			.ControlImageRect = New Rectangle(0, 0, 120, 32)
    '		End With
    '		Me.AddChild(CType(btnPlaceFac, UIControl))

    '		'lblAssets initial props
    '		lblAssets = New UILabel(MyBase.moUILib)
    '		With lblAssets
    '			.ControlName = "lblAssets"
    '			.Left = 15
    '			.Top = 55
    '			.Width = 100
    '			.Height = 24
    '			.Enabled = True
    '			.Visible = True
    '			.Caption = "Bank Assets"
    '			.ForeColor = muSettings.InterfaceBorderColor
    '			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '			.DrawBackImage = False
    '			.FontFormat = CType(4, DrawTextFormat)
    '		End With
    '		Me.AddChild(CType(lblAssets, UIControl))

    '		'lstBankAssets initial props
    '		lstBankAssets = New UIListBox(MyBase.moUILib)
    '		With lstBankAssets
    '			.ControlName = "lstBankAssets"
    '			.Left = 15
    '			.Top = 75
    '			.Width = 360
    '			.Height = 410
    '			.Enabled = True
    '			.Visible = True
    '			.BorderColor = muSettings.InterfaceBorderColor
    '			.FillColor = muSettings.InterfaceFillColor
    '			.ForeColor = muSettings.InterfaceBorderColor
    '			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
    '			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
    '		End With
    '		Me.AddChild(CType(lstBankAssets, UIControl))

    '		'lstYourAssets initial props
    '		lstYourAssets = New UIListBox(MyBase.moUILib)
    '		With lstYourAssets
    '			.ControlName = "lstYourAssets"
    '			.Left = 425
    '			.Top = 74
    '			.Width = 360
    '			.Height = 410
    '			.Enabled = True
    '			.Visible = True
    '			.BorderColor = muSettings.InterfaceBorderColor
    '			.FillColor = muSettings.InterfaceFillColor
    '			.ForeColor = muSettings.InterfaceBorderColor
    '			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
    '			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
    '		End With
    '		Me.AddChild(CType(lstYourAssets, UIControl))

    '		'btnViewLog initial props
    '		btnViewLog = New UIButton(MyBase.moUILib)
    '		With btnViewLog
    '			.ControlName = "btnViewLog"
    '			.Left = 275
    '			.Top = 51
    '			.Width = 100
    '			.Height = 24
    '			.Enabled = True
    '			.Visible = False
    '			If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
    '				.Visible = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ViewBankLog)
    '			End If
    '			.Caption = "View Log"
    '			.ForeColor = muSettings.InterfaceBorderColor
    '			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '			.DrawBackImage = True
    '			.FontFormat = CType(5, DrawTextFormat)
    '			.ControlImageRect = New Rectangle(0, 0, 120, 32)
    '		End With
    '		Me.AddChild(CType(btnViewLog, UIControl))

    '		'lblYourAssets initial props
    '		lblYourAssets = New UILabel(MyBase.moUILib)
    '		With lblYourAssets
    '			.ControlName = "lblYourAssets"
    '			.Left = 425
    '			.Top = 55
    '			.Width = 100
    '			.Height = 24
    '			.Enabled = True
    '			.Visible = True
    '			.Caption = "Your Assets"	' at <colony name>"
    '			.ForeColor = muSettings.InterfaceBorderColor
    '			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '			.DrawBackImage = False
    '			.FontFormat = CType(4, DrawTextFormat)
    '		End With
    '		Me.AddChild(CType(lblYourAssets, UIControl))

    '		'txtWithdrawQty initial props
    '		txtWithdrawQty = New UITextBox(MyBase.moUILib)
    '		With txtWithdrawQty
    '			.ControlName = "txtWithdrawQty"
    '			.Left = 170
    '			.Top = 495
    '			.Width = 100
    '			.Height = 20
    '			.Enabled = True
    '			.Visible = True
    '			.Caption = "0"
    '			.ForeColor = muSettings.InterfaceTextBoxForeColor
    '			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
    '			.DrawBackImage = False
    '			.FontFormat = CType(4, DrawTextFormat)
    '			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
    '			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
    '			.MaxLength = 9
    '			.BorderColor = muSettings.InterfaceBorderColor
    '			.bNumericOnly = True
    '		End With
    '		Me.AddChild(CType(txtWithdrawQty, UIControl))

    '		'btnWithdraw initial props
    '		btnWithdraw = New UIButton(MyBase.moUILib)
    '		With btnWithdraw
    '			.ControlName = "btnWithdraw"
    '			.Left = 276
    '			.Top = 495
    '			.Width = 100
    '			.Height = 24
    '			.Enabled = False
    '			If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
    '				.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.WithdrawLowSec) Or goCurrentPlayer.oGuild.HasPermission(RankPermissions.WithdrawHiSec)
    '			End If
    '			.Visible = True
    '			.Caption = "Withdraw"
    '			.ForeColor = muSettings.InterfaceBorderColor
    '			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '			.DrawBackImage = True
    '			.FontFormat = CType(5, DrawTextFormat)
    '			.ControlImageRect = New Rectangle(0, 0, 120, 32)
    '		End With
    '		Me.AddChild(CType(btnWithdraw, UIControl))

    '		'btnDeposit initial props
    '		btnDeposit = New UIButton(MyBase.moUILib)
    '		With btnDeposit
    '			.ControlName = "btnDeposit"
    '			.Left = 686
    '			.Top = 495
    '			.Width = 100
    '			.Height = 24
    '			.Enabled = False
    '			If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
    '				.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ViewContentsLowSec) Or goCurrentPlayer.oGuild.HasPermission(RankPermissions.ViewContentsHiSec)
    '			End If
    '			.Visible = True
    '			.Caption = "Deposit"
    '			.ForeColor = muSettings.InterfaceBorderColor
    '			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
    '			.DrawBackImage = True
    '			.FontFormat = CType(5, DrawTextFormat)
    '			.ControlImageRect = New Rectangle(0, 0, 120, 32)
    '		End With
    '		Me.AddChild(CType(btnDeposit, UIControl))

    '		'txtDepositQty initial props
    '		txtDepositQty = New UITextBox(MyBase.moUILib)
    '		With txtDepositQty
    '			.ControlName = "txtDepositQty"
    '			.Left = 580
    '			.Top = 495
    '			.Width = 100
    '			.Height = 20
    '			.Enabled = True
    '			.Visible = True
    '			.Caption = "0"
    '			.ForeColor = muSettings.InterfaceTextBoxForeColor
    '			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
    '			.DrawBackImage = False
    '			.FontFormat = CType(4, DrawTextFormat)
    '			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
    '			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
    '			.MaxLength = 9
    '			.BorderColor = muSettings.InterfaceBorderColor
    '			.bNumericOnly = True
    '		End With
    '		Me.AddChild(CType(txtDepositQty, UIControl))

    '		'cboTradePost initial props
    '		cboTradepost = New UIComboBox(oUILib)
    '		With cboTradepost
    '			.ControlName = "cboTradePost"
    '			.Left = lblYourAssets.Left + lblYourAssets.Width + 5
    '			.Top = lblYourAssets.Top
    '			.Width = lstYourAssets.Left + lstYourAssets.Width - .Left
    '			.Height = 18
    '			.Enabled = True
    '			.Visible = True
    '			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
    '			.BorderColor = muSettings.InterfaceBorderColor
    '               .FillColor = muSettings.InterfaceTextBoxFillColor
    '               .ForeColor = muSettings.InterfaceBorderColor
    '			.HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
    '			.mbAcceptReprocessEvents = True
    '		End With
    '		Me.AddChild(CType(cboTradepost, UIControl))

    '		Dim yMsg(1) As Byte
    '		System.BitConverter.GetBytes(GlobalMessageCode.eGetTradePostList).CopyTo(yMsg, 0)
    '		MyBase.moUILib.SendMsgToPrimary(yMsg)

    '		If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
    '			If goCurrentPlayer.oGuild.lGuildHallID > -1 Then
    '				ReDim yMsg(5)
    '				System.BitConverter.GetBytes(GlobalMessageCode.eGetGuildAssets).CopyTo(yMsg, 0)
    '				System.BitConverter.GetBytes(goCurrentPlayer.oGuild.lGuildHallID).CopyTo(yMsg, 2)
    '				MyBase.moUILib.SendMsgToPrimary(yMsg)
    '			End If
    '		End If

    '	End Sub

    '	Private Sub btnDeposit_Click(ByVal sName As String) Handles btnDeposit.Click
    '		If lstYourAssets Is Nothing = False Then
    '			If lstYourAssets.ListIndex < 0 Then
    '				MyBase.moUILib.AddNotification("Select an item to deposit first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '				Return
    '			End If
    '			If cboTradepost.ListIndex < 0 Then
    '				MyBase.moUILib.AddNotification("Select a tradepost to send the goods.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '				Return
    '			End If
    '			Dim blQty As Int64 = 1
    '			If txtDepositQty Is Nothing = False Then
    '				If txtDepositQty.Caption = "" Then txtDepositQty.Caption = "1"
    '				blQty = CLng(txtDepositQty.Caption)
    '			End If

    '			Dim lID As Int32 = lstYourAssets.ItemData(lstYourAssets.ListIndex)
    '			Dim iTypeID As Int16 = CShort(lstYourAssets.ItemData2(lstYourAssets.ListIndex))
    '			GuildBankTrans(1, lID, iTypeID, blQty)
    '		End If
    '	End Sub

    '	Private Sub btnGotoFacLoc_Click(ByVal sName As String) Handles btnGotoFacLoc.Click
    '		If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
    '			If goCurrentPlayer.oGuild.lGuildHallLocID > 0 AndAlso goCurrentPlayer.oGuild.iGuildHallLocTypeID > 0 Then
    '				Dim utmp As PlayerComm.WPAttachment
    '				With utmp
    '					.EnvirID = goCurrentPlayer.oGuild.lGuildHallLocID
    '					.EnvirTypeID = goCurrentPlayer.oGuild.iGuildHallLocTypeID
    '					.LocX = goCurrentPlayer.oGuild.lGuildHallLocX
    '					.LocZ = goCurrentPlayer.oGuild.lGuildHallLocZ
    '					.sWPName = ""
    '					.AttachNumber = 0
    '				End With
    '				utmp.JumpToAttachment()
    '			End If
    '		End If
    '	End Sub

    '	Private Sub btnPlaceFac_Click(ByVal sName As String) Handles btnPlaceFac.Click
    '		If goCurrentEnvir Is Nothing = False Then
    '			If goCurrentEnvir.ObjTypeID <> ObjectType.ePlanet Then
    '				MyBase.moUILib.AddNotification("Guild Facilities can only be placed on planets.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '				Return
    '			End If
    '			MyBase.moUILib.BuildGhost = goResMgr.GetMesh(42)
    '			MyBase.moUILib.vecBuildGhostLoc = New Vector3(0, 0, 0)
    '			MyBase.moUILib.BuildGhostAngle = 0
    '			MyBase.moUILib.BuildGhostID = 84
    '			MyBase.moUILib.BuildGhostTypeID = ObjectType.eFacilityDef
    '               MyBase.moUILib.lUISelectState = UILib.eSelectState.ePlaceGuildFacility
    '               MyBase.moUILib.bBuildGhostNaval = False
    '		End If
    '	End Sub

    '	Private Sub btnViewLog_Click(ByVal sName As String) Handles btnViewLog.Click
    '		mbViewLog = Not mbViewLog
    '		If mbViewLog = True Then btnViewLog.Caption = "View Bank" Else btnViewLog.Caption = "View Log"
    '	End Sub

    '	Private Sub btnWithdraw_Click(ByVal sName As String) Handles btnWithdraw.Click
    '		If lstBankAssets Is Nothing = False Then
    '			If lstBankAssets.ListIndex < 0 Then
    '				MyBase.moUILib.AddNotification("Select an item to withdraw first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '				Return
    '			End If
    '			If cboTradepost.ListIndex < 0 Then
    '				MyBase.moUILib.AddNotification("Select a tradepost to receive the goods.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '				Return
    '			End If
    '			Dim blQty As Int64 = 1
    '			If txtWithdrawQty Is Nothing = False Then
    '				If txtWithdrawQty.Caption = "" Then txtWithdrawQty.Caption = "1"
    '				blQty = CLng(txtWithdrawQty.Caption)
    '			End If

    '			Dim lID As Int32 = lstBankAssets.ItemData(lstBankAssets.ListIndex)
    '			Dim iTypeID As Int16 = CShort(lstBankAssets.ItemData2(lstBankAssets.ListIndex))
    '			GuildBankTrans(2, lID, iTypeID, blQty)

    '		End If
    '	End Sub

    '	Private Sub GuildBankTrans(ByVal yType As Byte, ByVal lID As Int32, ByVal iTypeID As Int16, ByVal blQty As Int64)
    '		Dim yMsg(20) As Byte
    '		Dim lPos As Int32 = 0
    '		System.BitConverter.GetBytes(GlobalMessageCode.eGuildBankTransaction).CopyTo(yMsg, lPos) : lPos += 2

    '		If mbViewHighSec = True Then
    '			yType = CByte(yType Or 128)
    '		End If

    '		yMsg(lPos) = yType : lPos += 1			'1 indicates deposit
    '		System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
    '		System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, lPos) : lPos += 2
    '		System.BitConverter.GetBytes(blQty).CopyTo(yMsg, lPos) : lPos += 8
    '		Dim lTPID As Int32 = cboTradepost.ItemData(cboTradepost.ListIndex)
    '		System.BitConverter.GetBytes(lTPID).CopyTo(yMsg, lPos) : lPos += 4

    '		If yType = 0 Then
    '			MyBase.moUILib.AddNotification("Request sent to deposit items into guild bank.", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '		Else
    '			MyBase.moUILib.AddNotification("Request sent to withdraw items into guild bank.", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    '		End If

    '		MyBase.moUILib.SendMsgToPrimary(yMsg)
    '	End Sub

    '	Public Overrides Sub NewFrame()
    '		Dim oTPC As TradePostContents = Nothing

    '		If cboTradepost Is Nothing = False AndAlso cboTradepost.ListIndex > -1 Then
    '			Dim lTradePostID As Int32 = cboTradepost.ItemData(cboTradepost.ListIndex)
    '			oTPC = Nothing
    '			For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
    '				If TradePostContents.lTradePostContentsIdx(X) = lTradePostID Then
    '					oTPC = TradePostContents.oTradePostContents(X)
    '					Exit For
    '				End If
    '			Next X

    '			If oTPC Is Nothing = False Then
    '				'Ok... ensure the list has everything
    '				If mlTPCLast <> oTPC.lLastUpdate Then
    '					mlTPCLast = oTPC.lLastUpdate
    '					oTPC.PopulateList(lstYourAssets, 255)
    '				Else : oTPC.SmartPopulateList(lstYourAssets, 255)
    '				End If
    '			Else
    '				If lstYourAssets.ListCount > 0 Then lstYourAssets.Clear()
    '			End If
    '		End If

    '		If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
    '			oTPC = Nothing
    '			Dim lGuildTP As Int32 = goCurrentPlayer.oGuild.lGuildHallID
    '			If lGuildTP > -1 Then
    '				For X As Int32 = 0 To TradePostContents.lTradePostContentsUB
    '					If TradePostContents.lTradePostContentsIdx(X) = lGuildTP Then
    '						oTPC = TradePostContents.oTradePostContents(X)
    '						Exit For
    '					End If
    '				Next X
    '				If oTPC Is Nothing = False Then
    '					If mlGuildTPCLast <> oTPC.lLastUpdate Then
    '						mlTPCLast = oTPC.lLastUpdate
    '						oTPC.PopulateList(lstBankAssets, 255)
    '					Else : oTPC.SmartPopulateList(lstBankAssets, 255)
    '					End If
    '				ElseIf lstBankAssets.ListCount > 0 Then
    '					lstBankAssets.Clear()
    '				End If
    '			End If

    '			Dim sText As String = "Guild Treasury: " & goCurrentPlayer.oGuild.blTreasury.ToString("#,##0")
    '			If lblTreasury.Caption <> sText Then lblTreasury.Caption = sText
    '			With goCurrentPlayer.oGuild
    '				If .lGuildHallLocID > -1 AndAlso .iGuildHallLocTypeID > -1 Then
    '					sText = "Guild Facility Location: " & GetCacheObjectValue(.lGuildHallLocID, .iGuildHallLocTypeID)
    '				Else
    '					sText = "Guild Facility Location: Facility Not Built/Placed"
    '				End If
    '				If lblFacLoc.Caption <> sText Then lblFacLoc.Caption = sText
    '			End With
    '		End If



    '		'If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
    '		'	If lstBankAssets.ListCount = 0 Then
    '		'		lstBankAssets.AddItem("Credits: " & goCurrentPlayer.oGuild.blTreasury.ToString("#,##0"), False)
    '		'		lstBankAssets.ItemData(lstBankAssets.NewIndex) = 1 : lstBankAssets.ItemData2(lstBankAssets.NewIndex) = ObjectType.eCredits
    '		'	End If
    '		'End If

    '	End Sub

    '	Public Overrides Sub RenderEnd()

    '	End Sub

    '	Private Sub cboTradepost_ItemSelected(ByVal lItemIndex As Integer) Handles cboTradepost.ItemSelected
    '		If lItemIndex < 0 Then Return
    '		Dim lTradePostID As Int32 = cboTradepost.ItemData(cboTradepost.ListIndex)
    '		Dim yMsg(5) As Byte
    '		System.BitConverter.GetBytes(GlobalMessageCode.eGetTradePostTradeables).CopyTo(yMsg, 0)
    '		System.BitConverter.GetBytes(lTradePostID).CopyTo(yMsg, 2)
    '		MyBase.moUILib.SendMsgToPrimary(yMsg)
    '	End Sub

    '	Public Sub HandleGetTradePostList(ByRef yData() As Byte)
    '		Dim lPos As Int32 = 2		'for msgcode
    '		Dim lUB As Int32 = System.BitConverter.ToInt16(yData, lPos) - 1 : lPos += 2
    '		Dim lPrevID As Int32 = -1
    '		If cboTradePost.ListIndex <> -1 Then lPrevID = cboTradePost.ItemData(cboTradePost.ListIndex)

    '		cboTradePost.Clear()

    '		For X As Int32 = 0 To lUB
    '			Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
    '			Dim sName As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
    '			Dim sColony As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20

    '			cboTradePost.AddItem(sName & " (" & sColony & ")")
    '			cboTradePost.ItemData(cboTradePost.NewIndex) = lID
    '		Next X 
    '	End Sub

    'End Class
End Class
