Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmAliasing
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private lblAliasUN As UILabel
    Private txtAliasUN As UITextBox
    Private lblAliasPW As UILabel
    Private txtAliasPW As UITextBox
    Private lblAliasPlayer As UILabel
    Private txtAliasPlayer As UITextBox
    Private lblRights As UILabel
    Private WithEvents chkViewBG As UICheckBox
    Private WithEvents chkCreateBG As UICheckBox
    Private WithEvents chkModifyBG As UICheckBox
    Private chkBudget As UICheckBox
    Private WithEvents chkViewDipl As UICheckBox
    Private WithEvents chkAlterDipl As UICheckBox
    Private WithEvents chkViewAgents As UICheckBox
    Private WithEvents chkAlterAgents As UICheckBox
    Private WithEvents chkViewColonyStats As UICheckBox
    Private WithEvents chkViewEmail As UICheckBox
    Private chkViewMining As UICheckBox
	Private WithEvents chkViewTechs As UICheckBox
    Private WithEvents chkViewTreasury As UICheckBox
    Private WithEvents chkViewTrades As UICheckBox
    Private WithEvents chkViewUnits As UICheckBox
    Private WithEvents chkAddProd As UICheckBox
    Private WithEvents chkAlterLaunchPower As UICheckBox
    Private WithEvents chkChangeOrders As UICheckBox
    Private WithEvents chkTransferCargo As UICheckBox
    Private WithEvents chkMoveUnits As UICheckBox
    Private WithEvents chkDismantleUnits As UICheckBox
    Private WithEvents chkDockUndock As UICheckBox
    Private WithEvents chkChangEnvir As UICheckBox
    Private WithEvents chkViewResearch As UICheckBox
    Private WithEvents chkAddResearch As UICheckBox
    Private WithEvents chkAlterColonyStats As UICheckBox
    Private WithEvents chkAlterEmail As UICheckBox
    Private WithEvents chkAlterTrades As UICheckBox
    Private WithEvents chkCancelProd As UICheckBox
	Private WithEvents chkCancelResearch As UICheckBox
	Private WithEvents chkDesignTechs As UICheckBox
    Private lnDiv3 As UILine
    Private lnDiv4 As UILine
    Private lnDiv5 As UILine
    Private WithEvents lstAliases As UIListBox
    Private WithEvents btnCreateNew As UIButton
    Private WithEvents btnRemove As UIButton
    Private WithEvents btnClose As UIButton
    Private WithEvents btnSave As UIButton
    Private WithEvents btnReset As UIButton
    Private WithEvents btnHelp As UIButton

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmAliasing initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eAliasingWindow
            .ControlName = "frmAliasing"
            .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 256
            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 256
            .Width = 512
            .Height = 512
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 1
            .Width = 187
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Player Alias Configuration"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 26
            .Top = 2
            .Width = 24
            .Height = 22
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
            .Left = 2
            .Top = 25
            .Width = Me.Width - 4
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lblAliasUN initial props
        lblAliasUN = New UILabel(oUILib)
        With lblAliasUN
            .ControlName = "lblAliasUN"
            .Left = 255
            .Top = 165
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = False 'True
            .Caption = "Alias Username:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblAliasUN, UIControl))

        'txtAliasUN initial props
        txtAliasUN = New UITextBox(oUILib)
        With txtAliasUN
            .ControlName = "txtAliasUN"
            .Left = 370
            .Top = 165
            .Width = 130
            .Height = 18
            .Enabled = True
            .Visible = False ' True
            .Caption = ""
            If goCurrentPlayer Is Nothing = False Then .Caption = goCurrentPlayer.PlayerName
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 20
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtAliasUN, UIControl))

        'lblAliasPW initial props
        lblAliasPW = New UILabel(oUILib)
        With lblAliasPW
            .ControlName = "lblAliasPW"
            .Left = 255
            .Top = 190
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = False
            .Caption = "Alias Password:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblAliasPW, UIControl))

        'txtAliasPW initial props
        txtAliasPW = New UITextBox(oUILib)
        With txtAliasPW
            .ControlName = "txtAliasPW"
            .Left = 370
            .Top = 190
            .Width = 130
            .Height = 18
            .Enabled = True
            .Visible = False
            .Caption = "PASSWORD"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 20
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtAliasPW, UIControl))

        'lblAliasPlayer initial props
        lblAliasPlayer = New UILabel(oUILib)
        With lblAliasPlayer
            .ControlName = "lblAliasPlayer"
            .Left = 5
            .Top = 165
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Aliased Player:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat) 
        End With
        Me.AddChild(CType(lblAliasPlayer, UIControl))

        'txtAliasPlayer initial props
        txtAliasPlayer = New UITextBox(oUILib)
        With txtAliasPlayer
            .ControlName = "txtAliasPlayer"
            .Left = 110
            .Top = 165
            .Width = 130
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
            .MaxLength = 20
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtAliasPlayer, UIControl))

        'lblRights initial props
        lblRights = New UILabel(oUILib)
        With lblRights
            .ControlName = "lblRights"
            .Left = 5
            .Top = 200
            .Width = 178
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Alias Rights Configuration"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblRights, UIControl))

        'chkViewBG initial props
        chkViewBG = New UICheckBox(oUILib)
        With chkViewBG
            .ControlName = "chkViewBG"
            .Left = 15
            .Top = 225
            .Width = 125
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "View Battlegroups"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkViewBG, UIControl))

        'chkCreateBG initial props
        chkCreateBG = New UICheckBox(oUILib)
        With chkCreateBG
            .ControlName = "chkCreateBG"
            .Left = 30
            .Top = 240
            .Width = 137
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Create Battlegroups"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkCreateBG, UIControl))

        'chkModifyBG initial props
        chkModifyBG = New UICheckBox(oUILib)
        With chkModifyBG
            .ControlName = "chkModifyBG"
            .Left = 30
            .Top = 255
            .Width = 138
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Modify Battlegroups"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkModifyBG, UIControl))

        'chkBudget initial props
        chkBudget = New UICheckBox(oUILib)
        With chkBudget
            .ControlName = "chkBudget"
            .Left = 15
            .Top = 270'275
            .Width = 144
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "View Budget Window"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkBudget, UIControl))

        'chkViewDipl initial props
        chkViewDipl = New UICheckBox(oUILib)
        With chkViewDipl
            .ControlName = "chkViewDipl"
            .Left = 15
			.Top = 285 '295
            .Width = 116
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "View Diplomacy"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkViewDipl, UIControl))

        'chkAlterDipl initial props
        chkAlterDipl = New UICheckBox(oUILib)
        With chkAlterDipl
            .ControlName = "chkAlterDipl"
            .Left = 30
			.Top = 300 '310
            .Width = 128
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Modify Diplomacy"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkAlterDipl, UIControl))

        'chkViewAgents initial props
        chkViewAgents = New UICheckBox(oUILib)
        With chkViewAgents
            .ControlName = "chkViewAgents"
            .Left = 30
			.Top = 315 '325
            .Width = 95
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "View Agents"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkViewAgents, UIControl))

        'chkAlterAgents initial props
        chkAlterAgents = New UICheckBox(oUILib)
        With chkAlterAgents
            .ControlName = "chkAlterAgents"
            .Left = 50
			.Top = 330 '340
            .Width = 105
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Modify Agents"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkAlterAgents, UIControl))

        'chkViewColonyStats initial props
        chkViewColonyStats = New UICheckBox(oUILib)
        With chkViewColonyStats
            .ControlName = "chkViewColonyStats"
            .Left = 15
			.Top = 345 '360
            .Width = 125
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "View Colony Stats"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkViewColonyStats, UIControl))

        'chkViewEmail initial props
        chkViewEmail = New UICheckBox(oUILib)
        With chkViewEmail
            .ControlName = "chkViewEmail"
            .Left = 15
			.Top = 375 '395
            .Width = 84
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "View Email"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkViewEmail, UIControl))

        'chkViewMining initial props
        chkViewMining = New UICheckBox(oUILib)
        With chkViewMining
            .ControlName = "chkViewMining"
            .Left = 15
			.Top = 405 '430
            .Width = 140
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "View Mining Window"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkViewMining, UIControl))

        'chkViewTechs initial props
        chkViewTechs = New UICheckBox(oUILib)
        With chkViewTechs
            .ControlName = "chkViewTechs"
            .Left = 15
			.Top = 420 '450
            .Width = 134
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "View Tech Designs"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
		Me.AddChild(CType(chkViewTechs, UIControl))

		'chkDesignTechs initial props
		chkDesignTechs = New UICheckBox(oUILib)
		With chkDesignTechs
			.ControlName = "chkDesignTechs"
			.Left = 30
			.Top = 435
			.Width = 150
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Create Tech Designs"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		Me.AddChild(CType(chkDesignTechs, UIControl))

        'chkViewTreasury initial props
        chkViewTreasury = New UICheckBox(oUILib)
        With chkViewTreasury
            .ControlName = "chkViewTreasury"
            .Left = 285
            .Top = 225
            .Width = 105
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "View Treasury"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkViewTreasury, UIControl))

        'chkViewTrades initial props
        chkViewTrades = New UICheckBox(oUILib)
        With chkViewTrades
            .ControlName = "chkViewTrades"
            .Left = 300
            .Top = 240
            .Width = 94
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "View Trades"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkViewTrades, UIControl))

        'chkViewUnits initial props
        chkViewUnits = New UICheckBox(oUILib)
        With chkViewUnits
            .ControlName = "chkViewUnits"
            .Left = 285
            .Top = 275
            .Width = 163
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "View Units and Facilities"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkViewUnits, UIControl))

        'chkAddProd initial props
        chkAddProd = New UICheckBox(oUILib)
        With chkAddProd
            .ControlName = "chkAddProd"
            .Left = 300
            .Top = 290
            .Width = 111
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Add Production"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkAddProd, UIControl))

        'chkAlterLaunchPower initial props
        chkAlterLaunchPower = New UICheckBox(oUILib)
        With chkAlterLaunchPower
            .ControlName = "chkAlterLaunchPower"
            .Left = 300
            .Top = 320
            .Width = 182
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Alter Autolaunch and Power"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkAlterLaunchPower, UIControl))

        'chkChangeOrders initial props
        chkChangeOrders = New UICheckBox(oUILib)
        With chkChangeOrders
            .ControlName = "chkChangeOrders"
            .Left = 300
            .Top = 335
            .Width = 153
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Change Order Settings"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkChangeOrders, UIControl))

        'chkTransferCargo initial props
        chkTransferCargo = New UICheckBox(oUILib)
        With chkTransferCargo
            .ControlName = "chkTransferCargo"
            .Left = 300
            .Top = 350
            .Width = 109
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Transfer Cargo"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkTransferCargo, UIControl))

        'chkMoveUnits initial props
        chkMoveUnits = New UICheckBox(oUILib)
        With chkMoveUnits
            .ControlName = "chkMoveUnits"
            .Left = 300
            .Top = 365
            .Width = 86
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Move Units"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkMoveUnits, UIControl))

        'chkDockUndock initial props
        chkDockUndock = New UICheckBox(oUILib)
        With chkDockUndock
            .ControlName = "chkDockUndock"
            .Left = 315
            .Top = 380
            .Width = 135
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Dock/Undock Units"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkDockUndock, UIControl))

        'chkChangEnvir initial props
        chkChangEnvir = New UICheckBox(oUILib)
        With chkChangEnvir
            .ControlName = "chkChangEnvir"
            .Left = 315
            .Top = 395
            .Width = 168
            .Height = 17
            .Enabled = True
            .Visible = True
            .Caption = "Change Unit Environment"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkChangEnvir, UIControl))

        'chkViewResearch initial props
        chkViewResearch = New UICheckBox(oUILib)
        With chkViewResearch
            .ControlName = "chkViewResearch"
            .Left = 300
            .Top = 410
            .Width = 178
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "View Research Production"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkViewResearch, UIControl))

        'chkAddResearch initial props
        chkAddResearch = New UICheckBox(oUILib)
        With chkAddResearch
            .ControlName = "chkAddResearch"
            .Left = 315
            .Top = 425
            .Width = 105
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Add Research"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkAddResearch, UIControl))

        'chkAlterColonyStats initial props
        chkAlterColonyStats = New UICheckBox(oUILib)
        With chkAlterColonyStats
            .ControlName = "chkAlterColonyStats"
            .Left = 30
			.Top = 360 '375
            .Width = 138
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Modify Colony Stats"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkAlterColonyStats, UIControl))

        'chkAlterEmail initial props
        chkAlterEmail = New UICheckBox(oUILib)
        With chkAlterEmail
            .ControlName = "chkAlterEmail"
            .Left = 30
			.Top = 390 '410
            .Width = 97
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Modify Email"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkAlterEmail, UIControl))

        'chkAlterTrades initial props
        chkAlterTrades = New UICheckBox(oUILib)
        With chkAlterTrades
            .ControlName = "chkAlterTrades"
            .Left = 315
            .Top = 255
            .Width = 106
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Modify Trades"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkAlterTrades, UIControl))

        'chkCancelProd initial props
        chkCancelProd = New UICheckBox(oUILib)
        With chkCancelProd
            .ControlName = "chkCancelProd"
            .Left = 315
            .Top = 305
            .Width = 127
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Cancel Production"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkCancelProd, UIControl))

        'chkCancelResearch initial props
        chkCancelResearch = New UICheckBox(oUILib)
        With chkCancelResearch
            .ControlName = "chkCancelResearch"
            .Left = 315
            .Top = 440
            .Width = 58
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Cancel"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkCancelResearch, UIControl))

        'chkDismantleUnits initial props
        chkDismantleUnits = New UICheckBox(oUILib)
        With chkDismantleUnits
            .ControlName = "chkDismantleUnits"
            .Left = 300
            .Top = 455
            .Width = 110
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Dismantle Units"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkDismantleUnits, UIControl))

        'lnDiv3 initial props
        lnDiv3 = New UILine(oUILib)
        With lnDiv3
            .ControlName = "lnDiv3"
            .Left = 2
            .Top = 220
            .Width = Me.Width - 4
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv3, UIControl))

        'lstAliases initial props
        lstAliases = New UIListBox(oUILib)
        With lstAliases
            .ControlName = "lstAliases"
            .Left = 5
            .Top = 30
            .Width = 500
            .Height = 100
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstAliases, UIControl))

        'btnCreateNew initial props
        btnCreateNew = New UIButton(oUILib)
        With btnCreateNew
            .ControlName = "btnCreateNew"
            .Left = 300
            .Top = 135
            .Width = 100
            .Height = 22
            .Enabled = True
            .Visible = False
            .Caption = "Create New"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnCreateNew, UIControl))

        'btnRemove initial props
        btnRemove = New UIButton(oUILib)
        With btnRemove
            .ControlName = "btnRemove"
            .Left = 405
            .Top = 135
            .Width = 100
            .Height = 22
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

        'lnDiv5 initial props
        lnDiv5 = New UILine(oUILib)
        With lnDiv5
            .ControlName = "lnDiv5"
            .Left = 2
            .Top = 160
            .Width = Me.Width - 4
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv5, UIControl))

        'btnSave initial props
        btnSave = New UIButton(oUILib)
        With btnSave
            .ControlName = "btnSave"
            .Left = 110
            .Top = 480
            .Width = 120
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "Save Settings"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSave, UIControl))

        'btnReset initial props
        btnReset = New UIButton(oUILib)
        With btnReset
            .ControlName = "btnReset"
            .Left = 275
            .Top = 480
            .Width = 120
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "Reset"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnReset, UIControl))

        'lnDiv4 initial props
        lnDiv4 = New UILine(oUILib)
        With lnDiv4
            .ControlName = "lnDiv4"
            .Left = 2
            .Top = 475
            .Width = Me.Width - 4
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv4, UIControl))

        'btnHelp initial props
        btnHelp = New UIButton(oUILib)
        With btnHelp
            .ControlName = "btnHelp"
            .Left = btnClose.Left - 25
            .Top = btnClose.Top
            .Width = 24
            .Height = 22
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

        Dim oFrm As New frmAliasLogins(oUILib)
        oFrm.Visible = True
        oFrm.Left = Me.Left - oFrm.Width
        oFrm.Top = Me.Top

        Dim yMsg(1) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestAliasConfigs).CopyTo(yMsg, 0)
        MyBase.moUILib.SendMsgToPrimary(yMsg)

        If gbAliased = True Then
            '    MyBase.moUILib.AddNotification("You cannot alter the alias conigurations of an aliased player.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Dim oofrm As UIWindow = MyBase.moUILib.GetWindow("frmOptions")
            If oofrm Is Nothing = False Then
                oFrm.Left = oofrm.Left - oFrm.Width
                oFrm.Top = oofrm.Top
            End If
            ofrm = Nothing

            Return
        End If

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)


    End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.RemoveWindow("frmAliasLogins")
	End Sub

	Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
		'TODO: Do tutorial
		MyBase.moUILib.AddNotification("Tutorial not implemented yet.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
	End Sub

	Private Sub btnReset_Click(ByVal sName As String) Handles btnReset.Click
		If lstAliases.ListIndex < 0 Then
			For X As Int32 = 0 To Me.ChildrenUB
				If Me.moChildren(X) Is Nothing = False AndAlso moChildren(X).ControlName Is Nothing = False AndAlso moChildren(X).ControlName.ToUpper.StartsWith("CHK") Then
					CType(moChildren(X), UICheckBox).Value = False
				End If
			Next X
		Else
			Dim lIdx As Int32 = lstAliases.ItemData(lstAliases.ListIndex)
			If goCurrentPlayer Is Nothing = False Then
				If goCurrentPlayer.mlAliasUB >= lIdx Then
					If goCurrentPlayer.muAliases(lIdx).sPlayerName <> "" Then
						With goCurrentPlayer.muAliases(lIdx)
							SetChecksFromRights(CType(.lRights, AliasingRights))
						End With
					End If
				End If
			End If
		End If
	End Sub

	Private Sub btnSave_Click(ByVal sName As String) Handles btnSave.Click
		CreateAndSendMessage(1)	'indicates add/update
	End Sub

	Private Sub CreateAndSendMessage(ByVal yType As Byte)
		If txtAliasPlayer.Caption.Trim = "" Then
			MyBase.moUILib.AddNotification("Please enter a valid player name to send the alias to.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		If txtAliasUN.Caption.Trim.Length < 3 Then
			MyBase.moUILib.AddNotification("Username must be three characters or longer.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		If txtAliasPW.Caption.Trim.Length < 3 Then
			MyBase.moUILib.AddNotification("Password must be three characters or longer.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		Dim lRights As Int32 = CompileRights()
		If lRights = 0 Then
			MyBase.moUILib.AddNotification("Please select the rights for this alias configuration.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		Dim yMsg(66) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.ePlayerAliasConfig).CopyTo(yMsg, lPos) : lPos += 2
		yMsg(lPos) = yType : lPos += 1
		System.Text.ASCIIEncoding.ASCII.GetBytes(txtAliasPlayer.Caption.ToUpper).CopyTo(yMsg, lPos) : lPos += 20
		System.Text.ASCIIEncoding.ASCII.GetBytes(txtAliasUN.Caption.ToUpper).CopyTo(yMsg, lPos) : lPos += 20
		System.Text.ASCIIEncoding.ASCII.GetBytes(txtAliasPW.Caption.ToUpper).CopyTo(yMsg, lPos) : lPos += 20
		System.BitConverter.GetBytes(lRights).CopyTo(yMsg, lPos) : lPos += 4
		MyBase.moUILib.AddNotification("Alias Configuration Submitted to Server...", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		MyBase.moUILib.SendMsgToPrimary(yMsg)

		btnCreateNew.Caption = "Create New"
		lstAliases.Enabled = True
	End Sub

	Private Function CompileRights() As Int32
		Dim lResult As Int32 = 0

		If chkAddProd.Value = True Then lResult = lResult Or AliasingRights.eAddProduction
		If chkAddResearch.Value = True Then lResult = lResult Or AliasingRights.eAddResearch
		If chkAlterAgents.Value = True Then lResult = lResult Or AliasingRights.eAlterAgents
		If chkAlterColonyStats.Value = True Then lResult = lResult Or AliasingRights.eAlterColonyStats
		If chkAlterDipl.Value = True Then lResult = lResult Or AliasingRights.eAlterDiplomacy
		If chkAlterEmail.Value = True Then lResult = lResult Or AliasingRights.eAlterEmail
		If chkAlterLaunchPower.Value = True Then lResult = lResult Or AliasingRights.eAlterAutoLaunchPower
		If chkAlterTrades.Value = True Then lResult = lResult Or AliasingRights.eAlterTrades
		If chkBudget.Value = True Then lResult = lResult Or AliasingRights.eViewBudget
		If chkCancelProd.Value = True Then lResult = lResult Or AliasingRights.eCancelProduction
		If chkCancelResearch.Value = True Then lResult = lResult Or AliasingRights.eCancelResearch
		If chkChangEnvir.Value = True Then lResult = lResult Or AliasingRights.eChangeEnvironment
		If chkChangeOrders.Value = True Then lResult = lResult Or AliasingRights.eChangeBehavior
		If chkCreateBG.Value = True Then lResult = lResult Or AliasingRights.eCreateBattleGroups
		If chkDesignTechs.Value = True Then lResult = lResult Or AliasingRights.eCreateDesigns
		If chkDockUndock.Value = True Then lResult = lResult Or AliasingRights.eDockUndockUnits
		If chkModifyBG.Value = True Then lResult = lResult Or AliasingRights.eModifyBattleGroups
		If chkMoveUnits.Value = True Then lResult = lResult Or AliasingRights.eMoveUnits
		If chkTransferCargo.Value = True Then lResult = lResult Or AliasingRights.eTransferCargo
		If chkViewAgents.Value = True Then lResult = lResult Or AliasingRights.eViewAgents
		If chkViewBG.Value = True Then lResult = lResult Or AliasingRights.eViewBattleGroups
		If chkViewColonyStats.Value = True Then lResult = lResult Or AliasingRights.eViewColonyStats
		If chkViewDipl.Value = True Then lResult = lResult Or AliasingRights.eViewDiplomacy
		If chkViewEmail.Value = True Then lResult = lResult Or AliasingRights.eViewEmail
		If chkViewMining.Value = True Then lResult = lResult Or AliasingRights.eViewMining
		If chkViewResearch.Value = True Then lResult = lResult Or AliasingRights.eViewResearch
		If chkViewTechs.Value = True Then lResult = lResult Or AliasingRights.eViewTechDesigns
		If chkViewTrades.Value = True Then lResult = lResult Or AliasingRights.eViewTrades
		If chkViewTreasury.Value = True Then lResult = lResult Or AliasingRights.eViewTreasury
        If chkViewUnits.Value = True Then lResult = lResult Or AliasingRights.eViewUnitsAndFacilities
        If chkDismantleUnits.Value = True Then lResult = lResult Or AliasingRights.eDismantle

		Return lResult
	End Function

	Private Sub lstAliases_ItemClick(ByVal lIndex As Integer) Handles lstAliases.ItemClick
		If lIndex < 0 Then Return
		Dim lIdx As Int32 = lstAliases.ItemData(lIndex)
		If goCurrentPlayer Is Nothing = False Then
			If goCurrentPlayer.mlAliasUB >= lIdx Then
                If goCurrentPlayer.muAliases(lIdx).sPlayerName <> "" Then
                    SetFromItem(goCurrentPlayer.muAliases(lIdx), False)
                End If
			End If
		End If
    End Sub

    Public Sub SetFromItem(ByVal uItem As PlayerAlias, ByVal bClearIndex As Boolean)
        With uItem
            txtAliasPlayer.Caption = .sPlayerName
            txtAliasPW.Caption = .sPassword
            txtAliasUN.Caption = .sUserName
            SetChecksFromRights(CType(.lRights, AliasingRights))
        End With
        If bClearIndex = True Then lstAliases.ListIndex = -1
        btnSave.Enabled = Not bClearIndex
    End Sub

	Private Sub btnRemove_Click(ByVal sName As String) Handles btnRemove.Click
		CreateAndSendMessage(0)	'indicates removal
	End Sub

#Region "  Checkbox Clicks  "
	Private Sub chkAddProd_Click() Handles chkAddProd.Click
		If chkAddProd.Value = False Then chkCancelProd.Value = False Else chkViewUnits.Value = True
	End Sub

	Private Sub chkAddResearch_Click() Handles chkAddResearch.Click
		If chkAddResearch.Value = True Then
			chkViewResearch.Value = True
			chkViewUnits.Value = True
		Else : chkCancelResearch.Value = False
		End If
	End Sub

	Private Sub chkAlterAgents_Click() Handles chkAlterAgents.Click
		If chkAlterAgents.Value = True Then
			chkViewAgents.Value = True
			chkViewDipl.Value = True
		End If
	End Sub

	Private Sub chkAlterColonyStats_Click() Handles chkAlterColonyStats.Click
		If chkAlterColonyStats.Value = True Then chkViewColonyStats.Value = True
	End Sub

	Private Sub chkAlterDipl_Click() Handles chkAlterDipl.Click
		If chkAlterDipl.Value = True Then chkViewDipl.Value = True
	End Sub

	Private Sub chkAlterEmail_Click() Handles chkAlterEmail.Click
		If chkAlterEmail.Value = True Then chkViewEmail.Value = True
	End Sub

	Private Sub chkAlterLaunchPower_Click() Handles chkAlterLaunchPower.Click
		If chkAlterLaunchPower.Value = True Then chkViewUnits.Value = True
	End Sub

	Private Sub chkAlterTrades_Click() Handles chkAlterTrades.Click
		If chkAlterTrades.Value = True Then
			chkViewTrades.Value = True
			chkViewTreasury.Value = True
		End If
	End Sub

	Private Sub chkCancelProd_Click() Handles chkCancelProd.Click
		If chkCancelProd.Value = True Then
			chkAddProd.Value = True
			chkViewUnits.Value = True
		End If
	End Sub

	Private Sub chkCancelResearch_Click() Handles chkCancelResearch.Click
		If chkCancelResearch.Value = True Then
			chkViewResearch.Value = True
			chkViewUnits.Value = True
		End If
	End Sub

	Private Sub chkChangEnvir_Click() Handles chkChangEnvir.Click
		If chkChangEnvir.Value = True Then
			chkMoveUnits.Value = True
			chkViewUnits.Value = True
		End If
	End Sub

	Private Sub chkChangeOrders_Click() Handles chkChangeOrders.Click
		If chkChangeOrders.Value = True Then chkViewUnits.Value = True
	End Sub

	Private Sub chkCreateBG_Click() Handles chkCreateBG.Click
		If chkCreateBG.Value = True Then chkViewBG.Value = True
	End Sub

	Private Sub chkDesignTechs_Click() Handles chkDesignTechs.Click
		If chkDesignTechs.Value = True Then
			chkViewTechs.Value = True
		End If
	End Sub

	Private Sub chkDockUndock_Click() Handles chkDockUndock.Click
		If chkDockUndock.Value = True Then
			chkMoveUnits.Value = True
			chkViewUnits.Value = True
		End If
	End Sub

	Private Sub chkModifyBG_Click() Handles chkModifyBG.Click
		If chkModifyBG.Value = True Then
			chkViewBG.Value = True
		End If
	End Sub

	Private Sub chkMoveUnits_Click() Handles chkMoveUnits.Click
		If chkMoveUnits.Value = True Then
			chkViewUnits.Value = True
		Else
			chkDockUndock.Value = False
			chkChangEnvir.Value = False
		End If
	End Sub

	Private Sub chkTransferCargo_Click() Handles chkTransferCargo.Click
		If chkTransferCargo.Value = True Then
			chkViewUnits.Value = True
		End If
	End Sub

	Private Sub chkViewAgents_Click() Handles chkViewAgents.Click
		If chkViewAgents.Value = True Then
			chkViewDipl.Value = True
		Else : chkAlterAgents.Value = False
		End If
	End Sub

	Private Sub chkViewBG_Click() Handles chkViewBG.Click
		If chkViewBG.Value = False Then
			chkModifyBG.Value = False
			chkCreateBG.Value = False
		End If
	End Sub

	Private Sub chkViewColonyStats_Click() Handles chkViewColonyStats.Click
		If chkViewColonyStats.Value = False Then chkAlterColonyStats.Value = False
	End Sub

	Private Sub chkViewDipl_Click() Handles chkViewDipl.Click
		If chkViewDipl.Value = False Then
			chkAlterDipl.Value = False
			chkViewAgents.Value = False
			chkAlterAgents.Value = False
		End If
	End Sub

	Private Sub chkViewEmail_Click() Handles chkViewEmail.Click
		If chkViewEmail.Value = False Then chkAlterEmail.Value = False
	End Sub

	Private Sub chkViewResearch_Click() Handles chkViewResearch.Click
		If chkViewResearch.Value = False Then
			chkAddResearch.Value = False
			chkCancelResearch.Value = False
		Else : chkViewUnits.Value = True
		End If
	End Sub

	Private Sub chkViewTechs_Click() Handles chkViewTechs.Click
		If chkViewTechs.Value = False Then
			chkDesignTechs.Value = False
		End If
	End Sub

	Private Sub chkViewTrades_Click() Handles chkViewTrades.Click
		If chkViewTrades.Value = True Then
			chkViewTreasury.Value = True
		Else : chkAlterTrades.Value = False
		End If
	End Sub

	Private Sub chkViewTreasury_Click() Handles chkViewTreasury.Click
		If chkViewTreasury.Value = False Then
			chkViewTrades.Value = False
			chkAlterTrades.Value = False
		End If
	End Sub

	Private Sub chkViewUnits_Click() Handles chkViewUnits.Click
		If chkViewUnits.Value = False Then
			chkAddProd.Value = False
			chkCancelProd.Value = False
			chkAlterLaunchPower.Value = False
			chkChangeOrders.Value = False
			chkTransferCargo.Value = False
			chkMoveUnits.Value = False
			chkDockUndock.Value = False
			chkChangEnvir.Value = False
			chkViewResearch.Value = False
			chkAddResearch.Value = False
            chkCancelResearch.Value = False
            chkDismantleUnits.Value = False
		End If
	End Sub
#End Region

	Private Sub frmAliasing_OnNewFrame() Handles Me.OnNewFrame
        Dim lCnt As Int32 = 0
        Dim bSortMe As Boolean = False
		If goCurrentPlayer Is Nothing = False Then
			For X As Int32 = 0 To goCurrentPlayer.mlAliasUB
				If goCurrentPlayer.muAliases(X).sPlayerName.Trim <> "" Then
					lCnt += 1

					Dim bFound As Boolean = False
					For Y As Int32 = 0 To lstAliases.ListCount - 1
						If lstAliases.ItemData(Y) = X Then
							If lstAliases.List(Y) <> goCurrentPlayer.muAliases(X).sPlayerName Then lstAliases.List(Y) = goCurrentPlayer.muAliases(X).sPlayerName
							bFound = True
							Exit For
						End If
					Next Y
					If bFound = False Then
						lstAliases.AddItem(goCurrentPlayer.muAliases(X).sPlayerName)
                        lstAliases.ItemData(lstAliases.NewIndex) = X
                        bSortMe = True
					End If
				End If
			Next X
        End If
        If bSortMe Then lstAliases.SortList(True, False)
		If lstAliases.ListCount <> lCnt Then lstAliases.Clear()
	End Sub

	Private Sub SetChecksFromRights(ByVal lRights As AliasingRights)
		chkAddProd.Value = (lRights And AliasingRights.eAddProduction) <> 0
		chkAddResearch.Value = (lRights And AliasingRights.eAddResearch) <> 0
		chkAlterAgents.Value = (lRights And AliasingRights.eAlterAgents) <> 0
		chkAlterColonyStats.Value = (lRights And AliasingRights.eAlterColonyStats) <> 0
		chkAlterDipl.Value = (lRights And AliasingRights.eAlterDiplomacy) <> 0
		chkAlterEmail.Value = (lRights And AliasingRights.eAlterEmail) <> 0
		chkAlterLaunchPower.Value = (lRights And AliasingRights.eAlterAutoLaunchPower) <> 0
		chkAlterTrades.Value = (lRights And AliasingRights.eAlterTrades) <> 0
		chkBudget.Value = (lRights And AliasingRights.eViewBudget) <> 0
		chkCancelProd.Value = (lRights And AliasingRights.eCancelProduction) <> 0
		chkCancelResearch.Value = (lRights And AliasingRights.eCancelResearch) <> 0
		chkChangEnvir.Value = (lRights And AliasingRights.eChangeEnvironment) <> 0
		chkChangeOrders.Value = (lRights And AliasingRights.eChangeBehavior) <> 0
		chkCreateBG.Value = (lRights And AliasingRights.eCreateBattleGroups) <> 0
		chkDesignTechs.Value = (lRights And AliasingRights.eCreateDesigns) <> 0
		chkDockUndock.Value = (lRights And AliasingRights.eDockUndockUnits) <> 0
		chkModifyBG.Value = (lRights And AliasingRights.eModifyBattleGroups) <> 0
		chkMoveUnits.Value = (lRights And AliasingRights.eMoveUnits) <> 0
		chkTransferCargo.Value = (lRights And AliasingRights.eTransferCargo) <> 0
		chkViewAgents.Value = (lRights And AliasingRights.eViewAgents) <> 0
		chkViewBG.Value = (lRights And AliasingRights.eViewBattleGroups) <> 0
		chkViewColonyStats.Value = (lRights And AliasingRights.eViewColonyStats) <> 0
		chkViewDipl.Value = (lRights And AliasingRights.eViewDiplomacy) <> 0
		chkViewEmail.Value = (lRights And AliasingRights.eViewEmail) <> 0
		chkViewMining.Value = (lRights And AliasingRights.eViewMining) <> 0
		chkViewResearch.Value = (lRights And AliasingRights.eViewResearch) <> 0
		chkViewTechs.Value = (lRights And AliasingRights.eViewTechDesigns) <> 0
		chkViewTrades.Value = (lRights And AliasingRights.eViewTrades) <> 0
		chkViewTreasury.Value = (lRights And AliasingRights.eViewTreasury) <> 0
        chkViewUnits.Value = (lRights And AliasingRights.eViewUnitsAndFacilities) <> 0
        chkDismantleUnits.Value = (lRights And AliasingRights.eDismantle) <> 0
    End Sub

	Private Sub btnCreateNew_Click(ByVal sName As String) Handles btnCreateNew.Click
		txtAliasPlayer.Caption = ""
		txtAliasPW.Caption = ""
		txtAliasUN.Caption = ""
		SetChecksFromRights(0)

		If btnCreateNew.Caption.ToUpper = "CANCEL" Then
			btnCreateNew.Caption = "Create New"
			lstAliases.Enabled = True
			If lstAliases.ListCount > 0 Then
				lstAliases.ListIndex = 0
			End If
		Else
			btnCreateNew.Caption = "Cancel"
			lstAliases.Enabled = False
		End If
	End Sub

    Private Sub frmAliasing_WindowMoved() Handles Me.WindowMoved
        Dim ofrm As UIWindow = MyBase.moUILib.GetWindow("frmAliasLogins")
        If ofrm Is Nothing = False Then
            ofrm.Left = Me.Left - ofrm.Width
            ofrm.Top = Me.Top
        End If
    End Sub
End Class