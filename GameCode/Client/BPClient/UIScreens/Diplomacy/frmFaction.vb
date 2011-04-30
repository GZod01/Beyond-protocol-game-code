Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmFaction
	Inherits UIWindow

	Private Enum eySlotState As Byte
		Unaccepted = 0
		Accepted = 1
		WarNotJoined = 2
		RankTooHigh = 4
		InsufficientFactionSlots = 8
        FactionAtWar = 16
        ExRankMember = 32
		ForceRemove = 255
	End Enum

	Private lblTitle As UILabel
	Private lnDiv1 As UILine
	Private lblCurrentTitle As UILabel
	Private lblFactionSlots As UILabel
	Private lblTheirRank As UILabel
	Private lblTheirTitle1 As UILabel
	Private lblTheirTitle2 As UILabel
	Private lblTheirTitle3 As UILabel
	Private lblTheirTitle4 As UILabel
	Private lblTheirTitle5 As UILabel
	Private lblStatus As UILabel
	Private lblStatus1 As UILabel
	Private lblStatus2 As UILabel
	Private lblStatus3 As UILabel
	Private lblStatus4 As UILabel
	Private lblStatus5 As UILabel
	Private lblFactionBonus As UILabel
	Private lblOtherBonus As UILabel
	Private lblFactionCnt As UILabel
	Private lnDiv2 As UILine
	Private lblFactionStatus1 As UILabel
	Private lblFactionStatus2 As UILabel
	Private lblFactionStatus3 As UILabel
	Private txtFaction1 As UITextBox
	Private txtFaction2 As UITextBox
	Private txtFaction3 As UITextBox
    Private lblResTimeReduct As UILabel
    Private lblResTimeMaxReduct As UILabel
	Private lnDiv3 As UILine
	Private btnRemoveFaction1 As UIButton
	Private btnRemoveFaction2 As UIButton
	Private btnRemoveFaction3 As UIButton
	Private btnAcceptFaction1 As UIButton
	Private btnAcceptFaction2 As UIButton
	Private btnAcceptFaction3 As UIButton
	Private btnRemove1 As UIButton
	Private btnRemove2 As UIButton
	Private btnRemove3 As UIButton
	Private btnRemove4 As UIButton
	Private btnRemove5 As UIButton

	Private WithEvents cboSlot1 As UIComboBox
	Private WithEvents cboSlot2 As UIComboBox
	Private WithEvents cboSlot3 As UIComboBox
	Private WithEvents cboSlot4 As UIComboBox
	Private WithEvents cboSlot5 As UIComboBox

	Private WithEvents btnClose As UIButton
	Private WithEvents btnHelp As UIButton

	Private mbLoading As Boolean = True

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

        'If gbAliased = True Then Return

		'frmFaction initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eFaction
            .ControlName = "frmFaction"
            .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 256
            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 188
            .Width = 512
            .Height = 385
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 2
        End With

		'lblTitle initial props
		lblTitle = New UILabel(oUILib)
		With lblTitle
			.ControlName = "lblTitle"
			.Left = 5
			.Top = 5
			.Width = 164
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "Faction Management"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 11.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTitle, UIControl))

		'btnClose initial props
		btnClose = New UIButton(oUILib)
		With btnClose
			.ControlName = "btnClose"
			.Left = 487
			.Top = 2
			.Width = 24
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "X"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnClose, UIControl))

		'lnDiv1 initial props
		lnDiv1 = New UILine(oUILib)
		With lnDiv1
			.ControlName = "lnDiv1"
			.Left = 1
			.Top = 30
			.Width = 511
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv1, UIControl))

		'lblCurrentTitle initial props
		lblCurrentTitle = New UILabel(oUILib)
		With lblCurrentTitle
			.ControlName = "lblCurrentTitle"
			.Left = 5
			.Top = 35
			.Width = 218
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Your Current Title: Magistrate"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblCurrentTitle, UIControl))

		'lblFactionSlots initial props
		lblFactionSlots = New UILabel(oUILib)
		With lblFactionSlots
			.ControlName = "lblFactionSlots"
			.Left = 5
			.Top = 60
			.Width = 91
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Faction Slots"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblFactionSlots, UIControl))

		'lblTheirRank initial props
		lblTheirRank = New UILabel(oUILib)
		With lblTheirRank
			.ControlName = "lblTheirRank"
			.Left = 185
			.Top = 60
			.Width = 91
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Their Title"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTheirRank, UIControl))

		'lblTheirTitle1 initial props
		lblTheirTitle1 = New UILabel(oUILib)
		With lblTheirTitle1
			.ControlName = "lblTheirTitle1"
			.Left = 190
			.Top = 80
			.Width = 70
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Magistrate"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTheirTitle1, UIControl))

		'lblTheirTitle2 initial props
		lblTheirTitle2 = New UILabel(oUILib)
		With lblTheirTitle2
			.ControlName = "lblTheirTitle2"
			.Left = 190
			.Top = 105
			.Width = 70
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Magistrate"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTheirTitle2, UIControl))

		'lblTheirTitle3 initial props
		lblTheirTitle3 = New UILabel(oUILib)
		With lblTheirTitle3
			.ControlName = "lblTheirTitle3"
			.Left = 190
			.Top = 130
			.Width = 70
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Magistrate"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTheirTitle3, UIControl))

		'lblTheirTitle4 initial props
		lblTheirTitle4 = New UILabel(oUILib)
		With lblTheirTitle4
			.ControlName = "lblTheirTitle4"
			.Left = 190
			.Top = 155
			.Width = 70
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Magistrate"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTheirTitle4, UIControl))

		'lblTheirTitle5 initial props
		lblTheirTitle5 = New UILabel(oUILib)
		With lblTheirTitle5
			.ControlName = "lblTheirTitle5"
			.Left = 190
			.Top = 180
			.Width = 70
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Magistrate"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTheirTitle5, UIControl))

		'lblStatus1 initial props
		lblStatus1 = New UILabel(oUILib)
		With lblStatus1
			.ControlName = "lblStatus1"
			.Left = 280
			.Top = 80
			.Width = 150
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Waiting for Acceptance"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblStatus1, UIControl))

		'lblStatus initial props
		lblStatus = New UILabel(oUILib)
		With lblStatus
			.ControlName = "lblStatus"
			.Left = 275
			.Top = 60
			.Width = 91
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Status"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblStatus, UIControl))

		'btnRemove1 initial props
		btnRemove1 = New UIButton(oUILib)
		With btnRemove1
			.ControlName = "btnRemove1"
			.Left = 430
			.Top = 80
			.Width = 80
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Remove"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnRemove1, UIControl))

		'btnRemove2 initial props
		btnRemove2 = New UIButton(oUILib)
		With btnRemove2
			.ControlName = "btnRemove2"
			.Left = 430
			.Top = 105
			.Width = 80
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Remove"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnRemove2, UIControl))

		'btnRemove3 initial props
		btnRemove3 = New UIButton(oUILib)
		With btnRemove3
			.ControlName = "btnRemove3"
			.Left = 430
			.Top = 130
			.Width = 80
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Remove"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnRemove3, UIControl))

		'btnRemove4 initial props
		btnRemove4 = New UIButton(oUILib)
		With btnRemove4
			.ControlName = "btnRemove4"
			.Left = 430
			.Top = 155
			.Width = 80
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Remove"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnRemove4, UIControl))

		'btnRemove5 initial props
		btnRemove5 = New UIButton(oUILib)
		With btnRemove5
			.ControlName = "btnRemove5"
			.Left = 430
			.Top = 180
			.Width = 80
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Remove"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnRemove5, UIControl))

		'lblStatus2 initial props
		lblStatus2 = New UILabel(oUILib)
		With lblStatus2
			.ControlName = "lblStatus2"
			.Left = 280
			.Top = 105
			.Width = 150
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Waiting for Acceptance"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblStatus2, UIControl))

		'lblStatus3 initial props
		lblStatus3 = New UILabel(oUILib)
		With lblStatus3
			.ControlName = "lblStatus3"
			.Left = 280
			.Top = 130
			.Width = 150
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Waiting for Acceptance"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblStatus3, UIControl))

		'lblStatus4 initial props
		lblStatus4 = New UILabel(oUILib)
		With lblStatus4
			.ControlName = "lblStatus4"
			.Left = 280
			.Top = 155
			.Width = 150
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Waiting for Acceptance"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblStatus4, UIControl))

		'lblStatus5 initial props
		lblStatus5 = New UILabel(oUILib)
		With lblStatus5
			.ControlName = "lblStatus5"
			.Left = 280
			.Top = 180
			.Width = 150
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Waiting for Acceptance"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblStatus5, UIControl))

		'lblFactionBonus initial props
		lblFactionBonus = New UILabel(oUILib)
		With lblFactionBonus
			.ControlName = "lblFactionBonus"
			.Left = 30
			.Top = 205
			.Width = 218
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Your Faction Bonus: 0%"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblFactionBonus, UIControl))

		'lblFactionCnt initial props
		lblFactionCnt = New UILabel(oUILib)
		With lblFactionCnt
			.ControlName = "lblFactionCnt"
			.Left = 5
			.Top = 240
			.Width = 207
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "You are currently in 0 factions:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblFactionCnt, UIControl))

		'lnDiv2 initial props
		lnDiv2 = New UILine(oUILib)
		With lnDiv2
			.ControlName = "lnDiv2"
			.Left = 1
			.Top = 230
			.Width = 511
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv2, UIControl))

		'txtFaction1 initial props
		txtFaction1 = New UITextBox(oUILib)
		With txtFaction1
			.ControlName = "txtFaction1"
			.Left = 5
			.Top = 260
			.Width = 170
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
			.Locked = True
		End With
		Me.AddChild(CType(txtFaction1, UIControl))

		''lblFactionStatus initial props
		'lblFactionStatus = New UILabel(oUILib)
		'With lblFactionStatus
		'	.ControlName = "lblFactionStatus"
		'	.Left = 245
		'	.Top = 240
		'	.Width = 91
		'	.Height = 18
		'	.Enabled = True
		'	.Visible = True
		'	.Caption = "Status"
		'	.ForeColor = muSettings.InterfaceBorderColor
		'	.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
		'	.DrawBackImage = False
		'	.FontFormat = CType(4, DrawTextFormat)
		'End With
		'Me.AddChild(CType(lblFactionStatus, UIControl))

		'lblFactionStatus1 initial props
		lblFactionStatus1 = New UILabel(oUILib)
		With lblFactionStatus1
			.ControlName = "lblFactionStatus1"
			.Left = 180
			.Top = 260
			.Width = 170
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Waiting for Acceptance"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblFactionStatus1, UIControl))

		'lblFactionStatus2 initial props
		lblFactionStatus2 = New UILabel(oUILib)
		With lblFactionStatus2
			.ControlName = "lblFactionStatus2"
			.Left = 180
			.Top = 285
			.Width = 170
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Waiting for Acceptance"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblFactionStatus2, UIControl))

		'lblFactionStatus3 initial props
		lblFactionStatus3 = New UILabel(oUILib)
		With lblFactionStatus3
			.ControlName = "lblFactionStatus3"
			.Left = 180
			.Top = 310
			.Width = 170
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Waiting for Acceptance"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblFactionStatus3, UIControl))

		'txtFaction2 initial props
		txtFaction2 = New UITextBox(oUILib)
		With txtFaction2
			.ControlName = "txtFaction2"
			.Left = 5
			.Top = 285
			.Width = 170
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
			.Locked = True
		End With
		Me.AddChild(CType(txtFaction2, UIControl))

		'txtFaction3 initial props
		txtFaction3 = New UITextBox(oUILib)
		With txtFaction3
			.ControlName = "txtFaction3"
			.Left = 5
			.Top = 310
			.Width = 170
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
			.Locked = True
		End With
		Me.AddChild(CType(txtFaction3, UIControl))

		'btnRemoveFaction1 initial props
		btnRemoveFaction1 = New UIButton(oUILib)
		With btnRemoveFaction1
			.ControlName = "btnRemoveFaction1"
			.Left = 430
			.Top = 260
			.Width = 80
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Remove"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnRemoveFaction1, UIControl))

		'btnRemoveFaction2 initial props
		btnRemoveFaction2 = New UIButton(oUILib)
		With btnRemoveFaction2
			.ControlName = "btnRemoveFaction2"
			.Left = 430
			.Top = 285
			.Width = 80
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Remove"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnRemoveFaction2, UIControl))

		'btnRemoveFaction3 initial props
		btnRemoveFaction3 = New UIButton(oUILib)
		With btnRemoveFaction3
			.ControlName = "btnRemoveFaction3"
			.Left = 430
			.Top = 310
			.Width = 80
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Remove"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnRemoveFaction3, UIControl))

		'btnAcceptFaction1 initial props
		btnAcceptFaction1 = New UIButton(oUILib)
		With btnAcceptFaction1
			.ControlName = "btnAcceptFaction1"
			.Left = 350
			.Top = 260
			.Width = 80
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Accept"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnAcceptFaction1, UIControl))

		'btnAcceptFaction2 initial props
		btnAcceptFaction2 = New UIButton(oUILib)
		With btnAcceptFaction2
			.ControlName = "btnAcceptFaction2"
			.Left = 350
			.Top = 285
			.Width = 80
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Accept"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnAcceptFaction2, UIControl))

		'btnAcceptFaction3 initial props
		btnAcceptFaction3 = New UIButton(oUILib)
		With btnAcceptFaction3
			.ControlName = "btnAcceptFaction3"
			.Left = 350
			.Top = 310
			.Width = 80
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Accept"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnAcceptFaction3, UIControl))

		'lblResTimeReduct initial props
		lblResTimeReduct = New UILabel(oUILib)
		With lblResTimeReduct
			.ControlName = "lblResTimeReduct"
			.Left = 1
			.Top = 360
            .Width = 285 '511
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Current Research Time Reduction: 0%"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
		End With
        Me.AddChild(CType(lblResTimeReduct, UIControl))

        'lblResTimeReduct initial props
        lblResTimeMaxReduct = New UILabel(oUILib)
        With lblResTimeMaxReduct
            .ControlName = "lblResTimeMaxReduct"
            .Left = lblResTimeReduct.Left + lblResTimeReduct.Width + 5
            .Top = lblResTimeReduct.Top
            .Width = Me.Width - lblResTimeReduct.Width - 10
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Max Reduction: 0%"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblResTimeMaxReduct, UIControl))

		'lblOtherBonus initial props
		lblOtherBonus = New UILabel(oUILib)
		With lblOtherBonus
			.ControlName = "lblOtherBonus"
			.Left = 30
			.Top = 335
			.Width = 268
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Bonuses from other factions: 0%"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblOtherBonus, UIControl))

		'lnDiv3 initial props
		lnDiv3 = New UILine(oUILib)
		With lnDiv3
			.ControlName = "lnDiv3"
			.Left = 1
			.Top = 355
			.Width = 511
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv3, UIControl))

		'btnHelp initial props
		btnHelp = New UIButton(oUILib)
		With btnHelp
			.ControlName = "btnHelp"
			.Left = 462
			.Top = 2
			.Width = 24
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "?"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnHelp, UIControl))

		'cboSlot5 initial props
		cboSlot5 = New UIComboBox(oUILib)
		With cboSlot5
			.ControlName = "cboSlot5"
			.Left = 5
			.Top = 180
			.Width = 170
			.Height = 18
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(cboSlot5, UIControl))

        'cboSlot4 initial props
        cboSlot4 = New UIComboBox(oUILib)
        With cboSlot4
            .ControlName = "cboSlot4"
            .Left = 5
            .Top = 155
            .Width = 170
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(cboSlot4, UIControl))

        'cboSlot3 initial props
        cboSlot3 = New UIComboBox(oUILib)
        With cboSlot3
            .ControlName = "cboSlot3"
            .Left = 5
            .Top = 130
            .Width = 170
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(cboSlot3, UIControl))

        'cboSlot2 initial props
        cboSlot2 = New UIComboBox(oUILib)
        With cboSlot2
            .ControlName = "cboSlot2"
            .Left = 5
            .Top = 105
            .Width = 170
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(cboSlot2, UIControl))

        'cboSlot1 initial props
        cboSlot1 = New UIComboBox(oUILib)
        With cboSlot1
            .ControlName = "cboSlot1"
            .Left = 5
            .Top = 80
            .Width = 170
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
		Me.AddChild(CType(cboSlot1, UIControl))

		'Our event handlers
		AddHandler btnAcceptFaction1.Click, AddressOf btnAcceptFaction_Click
		AddHandler btnAcceptFaction2.Click, AddressOf btnAcceptFaction_Click
		AddHandler btnAcceptFaction3.Click, AddressOf btnAcceptFaction_Click
		AddHandler btnRemoveFaction1.Click, AddressOf btnRemoveFaction_Click
		AddHandler btnRemoveFaction2.Click, AddressOf btnRemoveFaction_Click
		AddHandler btnRemoveFaction3.Click, AddressOf btnRemoveFaction_Click
		AddHandler btnRemove1.Click, AddressOf btnRemoveSlot_Click
		AddHandler btnRemove2.Click, AddressOf btnRemoveSlot_Click
		AddHandler btnRemove3.Click, AddressOf btnRemoveSlot_Click
		AddHandler btnRemove4.Click, AddressOf btnRemoveSlot_Click
		AddHandler btnRemove5.Click, AddressOf btnRemoveSlot_Click
		'end of event handlers

		SendSetStateMsg(0, -1, 0)	'request our slot states

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)

		FillSlots()

		mbLoading = False
	End Sub

	Private Sub FillSlots()
		cboSlot1.Clear() : cboSlot2.Clear() : cboSlot3.Clear() : cboSlot4.Clear() : cboSlot5.Clear()

		If goCurrentPlayer Is Nothing Then Return

        Dim yMyRealTitle As Byte = goCurrentPlayer.yPlayerTitle
        If (yMyRealTitle And Player.PlayerRank.ExRankShift) <> 0 Then yMyRealTitle = yMyRealTitle Xor Player.PlayerRank.ExRankShift

		Dim lSorted() As Int32 = GetSortedPlayerRelIdxArray()
		If lSorted Is Nothing = False Then
			For X As Int32 = 0 To lSorted.GetUpperBound(0)
				Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(lSorted(X))
				If oRel Is Nothing = False Then
					Dim lPID As Int32 = oRel.lPlayerRegards
                    If lPID = goCurrentPlayer.ObjectID Then lPID = oRel.lThisPlayer

                    If oRel.oPlayerIntel Is Nothing = False Then
                        If oRel.oPlayerIntel.yPlayerTitle < yMyRealTitle Then

                            Dim sName As String = GetCacheObjectValue(lPID, ObjectType.ePlayer)

                            cboSlot1.AddItem(sName) : cboSlot1.ItemData(cboSlot1.NewIndex) = lPID
                            cboSlot2.AddItem(sName) : cboSlot2.ItemData(cboSlot2.NewIndex) = lPID
                            cboSlot3.AddItem(sName) : cboSlot3.ItemData(cboSlot3.NewIndex) = lPID
                            cboSlot4.AddItem(sName) : cboSlot4.ItemData(cboSlot4.NewIndex) = lPID
                            cboSlot5.AddItem(sName) : cboSlot5.ItemData(cboSlot5.NewIndex) = lPID

                        End If
                    End If
                End If
            Next X
        End If
    End Sub

	Private Sub frmFaction_OnNewFrame() Handles Me.OnNewFrame
		If goCurrentPlayer Is Nothing = False Then
            'Now, check our cbo list...
            Dim yMyRealTitle As Byte = goCurrentPlayer.yPlayerTitle
            If (yMyRealTitle And Player.PlayerRank.ExRankShift) <> 0 Then yMyRealTitle = yMyRealTitle Xor Player.PlayerRank.ExRankShift

			Dim lSorted() As Int32 = GetSortedPlayerRelIdxArray()
            If lSorted Is Nothing = False Then
                Dim lTmpIdx As Int32 = -1
                For X As Int32 = 0 To lSorted.GetUpperBound(0)
                    If cboSlot1.ListCount > lTmpIdx Then
                        Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(lSorted(X))
                        If oRel Is Nothing = False Then
                            Dim lPID As Int32 = oRel.lPlayerRegards
                            If lPID = goCurrentPlayer.ObjectID Then lPID = oRel.lThisPlayer

                            If oRel.oPlayerIntel Is Nothing = False Then
                                If oRel.oPlayerIntel.yPlayerTitle < yMyRealTitle Then
                                    lTmpIdx += 1
                                    If cboSlot1.ItemData(lTmpIdx) <> lPID Then
                                        FillSlots()
                                        Exit For
                                    End If

                                End If

                            End If

                        End If
                    Else
                        FillSlots()
                        Exit For
                    End If
                Next
            End If

			'ok, let's check our stuff
			Try
				'Our 5 slots
				SetSlotLine(goCurrentPlayer.lSlotID(0), goCurrentPlayer.ySlotState(0), cboSlot1, lblStatus1, btnRemove1, lblTheirTitle1)
				SetSlotLine(goCurrentPlayer.lSlotID(1), goCurrentPlayer.ySlotState(1), cboSlot2, lblStatus2, btnRemove2, lblTheirTitle2)
				SetSlotLine(goCurrentPlayer.lSlotID(2), goCurrentPlayer.ySlotState(2), cboSlot3, lblStatus3, btnRemove3, lblTheirTitle3)
				SetSlotLine(goCurrentPlayer.lSlotID(3), goCurrentPlayer.ySlotState(3), cboSlot4, lblStatus4, btnRemove4, lblTheirTitle4)
				SetSlotLine(goCurrentPlayer.lSlotID(4), goCurrentPlayer.ySlotState(4), cboSlot5, lblStatus5, btnRemove5, lblTheirTitle5)

                'our 3 factions
                Dim bRank1Correct As Boolean = goCurrentPlayer.yPlayerTitle <= Player.PlayerRank.Duke
                Dim bRank2Correct As Boolean = goCurrentPlayer.yPlayerTitle <= Player.PlayerRank.Governor
                Dim bRank3Correct As Boolean = goCurrentPlayer.yPlayerTitle = Player.PlayerRank.Magistrate
                SetFactionLine(goCurrentPlayer.lFactionID(0), goCurrentPlayer.yFactionState(0), txtFaction1, lblFactionStatus1, btnAcceptFaction1, btnRemoveFaction1, bRank1Correct)
                SetFactionLine(goCurrentPlayer.lFactionID(1), goCurrentPlayer.yFactionState(1), txtFaction2, lblFactionStatus2, btnAcceptFaction2, btnRemoveFaction2, bRank2Correct)
                SetFactionLine(goCurrentPlayer.lFactionID(2), goCurrentPlayer.yFactionState(2), txtFaction3, lblFactionStatus3, btnAcceptFaction3, btnRemoveFaction3, bRank3Correct)

				Dim fFactionBonus As Single = GetFactionBonus()
				Dim fOtherBonus As Single = GetOtherFactionBonus()
				Dim lFactionCnt As Int32 = 0
				For X As Int32 = 0 To 2
					If goCurrentPlayer.lFactionID(X) > 0 Then lFactionCnt += 1
				Next X

				Dim lMaxCnt As Int32 = 0
				If goCurrentPlayer.yPlayerTitle = Player.PlayerRank.Magistrate Then
					lMaxCnt = 3
				ElseIf goCurrentPlayer.yPlayerTitle = Player.PlayerRank.Governor Then
					lMaxCnt = 2
				ElseIf goCurrentPlayer.yPlayerTitle = Player.PlayerRank.Overseer OrElse goCurrentPlayer.yPlayerTitle = Player.PlayerRank.Duke Then
					lMaxCnt = 1
				End If

				Dim sText As String = "You are currently in " & lFactionCnt & " factions: (MAX " & lMaxCnt & ")"
				If lblFactionCnt.Caption <> sText Then lblFactionCnt.Caption = sText
				sText = "Your Faction Bonus: " & ((1.0F - fFactionBonus) * 100).ToString("#0.#") & "%"
				If lblFactionBonus.Caption <> sText Then lblFactionBonus.Caption = sText
				sText = "Bonuses from other factions: " & ((1.0F - fOtherBonus) * 100).ToString("#0.#") & "%"
				If lblOtherBonus.Caption <> sText Then lblOtherBonus.Caption = sText
				sText = "Current Research Time Reduction: " & ((1.0F - (fFactionBonus * fOtherBonus)) * 100).ToString("#0.#") & "%"
				If lblResTimeReduct.Caption <> sText Then lblResTimeReduct.Caption = sText
				sText = "Your Current Title: " & Player.GetPlayerTitle(goCurrentPlayer.yPlayerTitle, goCurrentPlayer.bIsMale)
                If lblCurrentTitle.Caption <> sText Then lblCurrentTitle.Caption = sText

                sText = "Max Reduction: " & ((1.0F - (GetMaxFactionBonus() * GetMaxOtherFactionBonus())) * 100).ToString("#0.#") & "%"
                If lblResTimeMaxReduct.Caption <> sText Then lblResTimeMaxReduct.Caption = sText


			Catch
				'Do nothing, we'll get it next pass
			End Try
		End If
	End Sub

	Private Sub SetSlotLine(ByVal lSlotID As Int32, ByVal ySlotState As Byte, ByRef cboData As UIComboBox, ByRef lblStat As UILabel, ByRef btn As UIButton, ByRef lblSlotTitle As UILabel)
		If lSlotID > 0 Then
			If cboData.ListIndex = -1 OrElse cboData.ItemData(cboData.ListIndex) <> lSlotID Then
				Dim bResetLoad As Boolean = Not mbLoading
				mbLoading = True
				cboData.FindComboItemData(lSlotID)
				Me.IsDirty = True
				If bResetLoad = True Then mbLoading = False
			End If
			If btn.Visible = False Then btn.Visible = True
			Dim sText As String = GetSlotStateText(ySlotState)
			If lblStat.Caption <> sText Then lblStat.Caption = sText
			Dim clrVal As System.Drawing.Color = GetSlotStateColor(ySlotState)
			If lblStat.ForeColor <> clrVal Then lblStat.ForeColor = clrVal

			If lSlotID <> -1 Then
				Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRel(lSlotID)
				If oRel Is Nothing = False Then
					If oRel.oPlayerIntel Is Nothing = False Then
						sText = Player.GetPlayerTitle(oRel.oPlayerIntel.yPlayerTitle, oRel.oPlayerIntel.bIsMale)
					End If
				End If
			Else : sText = ""
			End If
			If lblSlotTitle.Caption <> sText Then lblSlotTitle.Caption = sText
		Else
			If cboData.ListIndex <> -1 Then cboData.ListIndex = -1
			If lblStat.Caption <> "Unassigned/Unused" Then lblStat.Caption = "Unassigned/Unused"
			If lblStat.ForeColor <> System.Drawing.Color.FromArgb(255, 128, 128, 128) Then lblStat.ForeColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
			If lblSlotTitle.Caption <> "" Then lblSlotTitle.Caption = ""
			If btn.Visible = True Then btn.Visible = False
		End If
	End Sub
	Private Function GetSlotStateText(ByVal yValue As Byte) As String
		If yValue = 0 Then Return "Unaccepted"

		Dim sValue As String = "Accepted"

        If (yValue And eySlotState.WarNotJoined) <> 0 Then sValue = "War Not Joined"
		If (yValue And eySlotState.RankTooHigh) <> 0 Then sValue = "Rank Too High"
		If (yValue And eySlotState.FactionAtWar) <> 0 Then sValue = "Factions At War"
        If (yValue And eySlotState.InsufficientFactionSlots) <> 0 Then sValue = "Exceeds Faction Slots"
        If (yValue And eySlotState.ExRankMember) <> 0 Then sValue = "Ex-Ranked"
		Return sValue
	End Function
	Private Function GetSlotStateColor(ByVal yValue As Byte) As Color
		If yValue = 0 Then Return System.Drawing.Color.FromArgb(255, 255, 255, 0)
		If yValue = 1 Then Return System.Drawing.Color.FromArgb(255, 0, 255, 0)
		Return System.Drawing.Color.FromArgb(255, 255, 0, 0)
	End Function
    Private Sub SetFactionLine(ByVal lPlayerID As Int32, ByVal yState As Byte, ByRef txtBox As UITextBox, ByRef lblItem As UILabel, ByRef btnAcc As UIButton, ByRef btnRem As UIButton, ByVal bIsRankCorrect As Boolean)
        If lPlayerID > 0 Then
            Dim sText As String = GetCacheObjectValue(lPlayerID, ObjectType.ePlayer)
            If txtBox.Caption <> sText Then txtBox.Caption = sText
            sText = GetSlotStateText(yState)
            If lblItem.Caption <> sText Then lblItem.Caption = sText
            Dim clrVal As System.Drawing.Color = GetSlotStateColor(yState)
            If lblItem.ForeColor <> clrVal Then lblItem.ForeColor = clrVal

            'now, our buttons...
            If yState = eySlotState.Unaccepted Then
                If btnAcc.Visible <> True Then btnAcc.Visible = True
            ElseIf btnAcc.Visible = True Then
                btnAcc.Visible = False
            End If
            If btnRem.Visible = False Then btnRem.Visible = True
        Else
            If txtBox.Caption <> "" Then txtBox.Caption = ""
            If lblItem.Caption <> "" Then lblItem.Caption = ""
            If btnAcc.Visible = True Then btnAcc.Visible = False
            If btnRem.Visible = True Then btnRem.Visible = False

            If txtBox.Visible <> bIsRankCorrect Then txtBox.Visible = bIsRankCorrect
            If lblItem.Visible <> bIsRankCorrect Then lblItem.Visible = bIsRankCorrect

        End If
    End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
		MyBase.moUILib.AddNotification("Tutorial not yet implemented.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
	End Sub

    Private Sub btnAcceptFaction_Click(ByVal sName As String)

        If goCurrentPlayer Is Nothing = False AndAlso (goCurrentPlayer.yPlayerTitle And Player.PlayerRank.ExRankShift) <> 0 Then
            MyBase.moUILib.AddNotification("Unable to accept faction while you are ex-ranked.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        Dim lID As Int32 = -1
        If sName.ToUpper = "BTNACCEPTFACTION1" Then
            lID = goCurrentPlayer.lFactionID(0)
        ElseIf sName.ToUpper = "BTNACCEPTFACTION2" Then
            lID = goCurrentPlayer.lFactionID(1)
        Else : lID = goCurrentPlayer.lFactionID(2)
        End If
        If lID < 1 Then Return

        SendSetStateMsg(1, lID, eySlotState.Accepted)
    End Sub

	Private Sub btnRemoveFaction_Click(ByVal sName As String)
		Dim lID As Int32 = -1
		If sName.ToUpper = "BTNREMOVEFACTION1" Then
			lID = goCurrentPlayer.lFactionID(0)
		ElseIf sName.ToUpper = "BTNREMOVEFACTION2" Then
			lID = goCurrentPlayer.lFactionID(1)
		Else : lID = goCurrentPlayer.lFactionID(2)
		End If
		If lID < 1 Then Return
		SendSetStateMsg(1, lID, eySlotState.ForceRemove)
	End Sub

	Private Sub btnRemoveSlot_Click(ByVal sName As String)
		Dim lID As Int32 = -1
		Select Case sName.ToUpper
			Case "BTNREMOVE1"
				lID = goCurrentPlayer.lSlotID(0)
			Case "BTNREMOVE2"
				lID = goCurrentPlayer.lSlotID(1)
			Case "BTNREMOVE3"
				lID = goCurrentPlayer.lSlotID(2)
			Case "BTNREMOVE4"
				lID = goCurrentPlayer.lSlotID(3)
			Case Else
				lID = goCurrentPlayer.lSlotID(4)
		End Select
		If lID < 1 Then Return

		SendSetStateMsg(0, lID, eySlotState.ForceRemove)
	End Sub

	Private Sub cboSlot1_ItemSelected(ByVal lItemIndex As Integer) Handles cboSlot1.ItemSelected
		If mbLoading = True Then Return
		Dim lID As Int32 = goCurrentPlayer.lSlotID(0)
		Dim lNewID As Int32 = -1
		If cboSlot1.ListIndex <> -1 Then
			lNewID = cboSlot1.ItemData(cboSlot1.ListIndex)
		End If
		If VerifySlotSelection(lNewID) = False Then Return
		For X As Int32 = 0 To 4
			If goCurrentPlayer.lSlotID(X) = lNewID Then
				MyBase.moUILib.AddNotification("That player is already in a slot.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If
		Next X
		If lID <> lNewID AndAlso lID > 0 Then
			'ok, remove the current one
			SendSetStateMsg(0, lID, eySlotState.ForceRemove)
		End If
		SendSetStateMsg(0, lNewID, eySlotState.Unaccepted)
		cboSlot1.ListIndex = -1
	End Sub

	Private Sub cboSlot2_ItemSelected(ByVal lItemIndex As Integer) Handles cboSlot2.ItemSelected
		If mbLoading = True Then Return
		Dim lID As Int32 = goCurrentPlayer.lSlotID(1)
		Dim lNewID As Int32 = -1
		If cboSlot2.ListIndex <> -1 Then
			lNewID = cboSlot2.ItemData(cboSlot2.ListIndex)
		End If
		If VerifySlotSelection(lNewID) = False Then Return
		For X As Int32 = 0 To 4
			If goCurrentPlayer.lSlotID(X) = lNewID Then
				MyBase.moUILib.AddNotification("That player is already in a slot.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If
		Next X
		If lID <> lNewID AndAlso lID > 0 Then
			'ok, remove the current one
			SendSetStateMsg(0, lID, eySlotState.ForceRemove)
		End If
		SendSetStateMsg(0, lNewID, eySlotState.Unaccepted)
		cboSlot2.ListIndex = -1
	End Sub

	Private Sub cboSlot3_ItemSelected(ByVal lItemIndex As Integer) Handles cboSlot3.ItemSelected
		If mbLoading = True Then Return
		Dim lID As Int32 = goCurrentPlayer.lSlotID(2)
		Dim lNewID As Int32 = -1
		If cboSlot3.ListIndex <> -1 Then
			lNewID = cboSlot3.ItemData(cboSlot3.ListIndex)
		End If
		If VerifySlotSelection(lNewID) = False Then Return
		For X As Int32 = 0 To 4
			If goCurrentPlayer.lSlotID(X) = lNewID Then
				MyBase.moUILib.AddNotification("That player is already in a slot.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If
		Next X
		If lID <> lNewID AndAlso lID > 0 Then
			'ok, remove the current one
			SendSetStateMsg(0, lID, eySlotState.ForceRemove)
		End If
		SendSetStateMsg(0, lNewID, eySlotState.Unaccepted)
		cboSlot3.ListIndex = -1
	End Sub

	Private Sub cboSlot4_ItemSelected(ByVal lItemIndex As Integer) Handles cboSlot4.ItemSelected
		If mbLoading = True Then Return
		Dim lID As Int32 = goCurrentPlayer.lSlotID(3)
		Dim lNewID As Int32 = -1
		If cboSlot4.ListIndex <> -1 Then
			lNewID = cboSlot4.ItemData(cboSlot4.ListIndex)
		End If
		If VerifySlotSelection(lNewID) = False Then Return
		For X As Int32 = 0 To 4
			If goCurrentPlayer.lSlotID(X) = lNewID Then
				MyBase.moUILib.AddNotification("That player is already in a slot.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If
		Next X
		If lID <> lNewID AndAlso lID > 0 Then
			'ok, remove the current one
			SendSetStateMsg(0, lID, eySlotState.ForceRemove)
		End If
		SendSetStateMsg(0, lNewID, eySlotState.Unaccepted)
		cboSlot4.ListIndex = -1
	End Sub

	Private Sub cboSlot5_ItemSelected(ByVal lItemIndex As Integer) Handles cboSlot5.ItemSelected
		If mbLoading = True Then Return
		Dim lID As Int32 = goCurrentPlayer.lSlotID(4)
		Dim lNewID As Int32 = -1
		If cboSlot5.ListIndex <> -1 Then
			lNewID = cboSlot5.ItemData(cboSlot5.ListIndex)
		End If
		If VerifySlotSelection(lNewID) = False Then Return
		For X As Int32 = 0 To 4
			If goCurrentPlayer.lSlotID(X) = lNewID Then
				MyBase.moUILib.AddNotification("That player is already in a slot.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If
		Next X
		If lID <> lNewID AndAlso lID > 0 Then
			'ok, remove the current one
			SendSetStateMsg(0, lID, eySlotState.ForceRemove)
		End If
		SendSetStateMsg(0, lNewID, eySlotState.Unaccepted)
		cboSlot5.ListIndex = -1
	End Sub

	Private Function VerifySlotSelection(ByVal lPlayerID As Int32) As Boolean
		If goCurrentPlayer Is Nothing Then Return True

		Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRel(lPlayerID)
		If oRel Is Nothing Then Return True
		If oRel.oPlayerIntel Is Nothing Then Return True

		If oRel.oPlayerIntel.yPlayerTitle < goCurrentPlayer.yPlayerTitle Then
			Return True
		Else
			MyBase.moUILib.AddNotification("That player cannot be in your faction because their rank is too high", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return False
		End If
	End Function

	Private Sub SendSetStateMsg(ByVal yType As Byte, ByVal lID As Int32, ByVal yState As Byte)
		Dim yMsg(7) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eUpdateSlotStates).CopyTo(yMsg, lPos) : lPos += 2
		yMsg(lPos) = yType : lPos += 1
		System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
		yMsg(lPos) = yState : lPos += 1

		MyBase.moUILib.SendMsgToPrimary(yMsg)
	End Sub

	Protected Overrides Sub Finalize()
		Try
			RemoveHandler btnAcceptFaction1.Click, AddressOf btnAcceptFaction_Click
			RemoveHandler btnAcceptFaction2.Click, AddressOf btnAcceptFaction_Click
			RemoveHandler btnAcceptFaction3.Click, AddressOf btnAcceptFaction_Click
			RemoveHandler btnRemoveFaction1.Click, AddressOf btnRemoveFaction_Click
			RemoveHandler btnRemoveFaction2.Click, AddressOf btnRemoveFaction_Click
			RemoveHandler btnRemoveFaction3.Click, AddressOf btnRemoveFaction_Click
			RemoveHandler btnRemove1.Click, AddressOf btnRemoveSlot_Click
			RemoveHandler btnRemove2.Click, AddressOf btnRemoveSlot_Click
			RemoveHandler btnRemove3.Click, AddressOf btnRemoveSlot_Click
			RemoveHandler btnRemove4.Click, AddressOf btnRemoveSlot_Click
			RemoveHandler btnRemove5.Click, AddressOf btnRemoveSlot_Click
		Catch
		End Try

		MyBase.Finalize()
	End Sub

    Private Function GetFactionBonus() As Single
        Dim fValue As Single = 1.0F

        Dim fMult As Single = 1.0F
        Select Case goCurrentPlayer.yPlayerTitle
            Case Player.PlayerRank.Emperor
                fMult = 0.85F
            Case Player.PlayerRank.King
                fMult = 0.76F
            Case Player.PlayerRank.Baron
                fMult = 0.81F
            Case Player.PlayerRank.Duke
                fMult = 0.96F
            Case Player.PlayerRank.Overseer
                fMult = 0.99F
            Case Player.PlayerRank.Governor
                fMult = 0.99F
        End Select

        If goCurrentPlayer Is Nothing = False Then
            For X As Int32 = 0 To 4
                If goCurrentPlayer.lSlotID(X) > 0 Then
                    Dim ySlotVal As Byte = goCurrentPlayer.ySlotState(X)
                    If (ySlotVal And eySlotState.Accepted) <> 0 AndAlso ((eySlotState.RankTooHigh Or eySlotState.InsufficientFactionSlots Or eySlotState.WarNotJoined Or eySlotState.ExRankMember) And ySlotVal) = 0 Then
                        fValue *= fMult
                    End If
                End If
            Next X
        End If
        Return fValue
    End Function
	Private Function GetOtherFactionBonus() As Single
		Dim fValue As Single = 1.0F
		If goCurrentPlayer Is Nothing = False Then
			For X As Int32 = 0 To 2
				If goCurrentPlayer.lFactionID(X) > 0 Then
					Dim yValue As Byte = goCurrentPlayer.yFactionState(X)
					If (yValue And eySlotState.Accepted) <> 0 AndAlso ((eySlotState.FactionAtWar Or eySlotState.RankTooHigh Or eySlotState.InsufficientFactionSlots) And yValue) = 0 Then
						fValue *= 0.5F
					End If
				End If
			Next X
		End If
		Return fValue
	End Function

    Private Function GetMaxFactionBonus() As Single
        Dim fValue As Single = 1.0F

        Dim fMult As Single = 1.0F
        Select Case goCurrentPlayer.yPlayerTitle
            Case Player.PlayerRank.Emperor
                fMult = 0.85F
            Case Player.PlayerRank.King
                fMult = 0.76F
            Case Player.PlayerRank.Baron
                fMult = 0.81F
            Case Player.PlayerRank.Duke
                fMult = 0.96F
            Case Player.PlayerRank.Overseer
                fMult = 0.99F
            Case Player.PlayerRank.Governor
                fMult = 0.99F
        End Select
        For X As Int32 = 0 To 4
            fValue *= fMult
        Next X
        Return fValue
    End Function
    Private Function GetMaxOtherFactionBonus() As Single
        Dim fValue As Single = 1.0F
        Dim lMaxCnt As Int32 = 0
        If goCurrentPlayer.yPlayerTitle = Player.PlayerRank.Magistrate Then
            lMaxCnt = 3
        ElseIf goCurrentPlayer.yPlayerTitle = Player.PlayerRank.Governor Then
            lMaxCnt = 2
        ElseIf goCurrentPlayer.yPlayerTitle = Player.PlayerRank.Overseer OrElse goCurrentPlayer.yPlayerTitle = Player.PlayerRank.Duke Then
            lMaxCnt = 1
        End If
        For X As Int32 = 0 To lMaxCnt - 1
            fValue *= 0.5F
        Next X
        Return fValue
    End Function
End Class