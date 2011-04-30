Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class frmGuildSetup
	Inherits UIWindow

	Private lblTitle As UILabel
	Private lnDiv1 As UILine
	Private lblGuildName As UILabel
	Private txtGuildName As UITextBox

	Private WithEvents btnSubmit As UIButton
	Private WithEvents btnCancel As UIButton
	Private WithEvents btnCheck As UIButton

	Private fraIconChooser As frmIconChooser

	Private fraInvite As UIWindow
	Private lblPlayerName As UILabel
	Private txtInvite1 As UITextBox
	Private txtInvite2 As UITextBox
	Private txtInvite3 As UITextBox
	Private txtInvite4 As UITextBox
	Private WithEvents chkAccept1 As UICheckBox
	Private WithEvents chkAccept2 As UICheckBox
	Private WithEvents chkAccept3 As UICheckBox
	Private WithEvents chkAccept4 As UICheckBox
	Private WithEvents btnInvite1 As UIButton
	Private WithEvents btnInvite2 As UIButton
	Private WithEvents btnInvite3 As UIButton
	Private WithEvents btnInvite4 As UIButton
	Private mlInviteIDs(3) As Int32

	'fraRanks....
	Private fraRanks As UIWindow
	Private txtRankName As UITextBox
	Private txtVoteRate As UITextBox
	Private lblVoteStr As UILabel
	Private lblRankName As UILabel
	Private lblTaxFlat As UILabel
	Private txtFlatRate As UITextBox
	Private lblTaxPerc As UILabel
	Private txtTaxPerc As UITextBox
	Private cboTaxPercType As UIComboBox
	Private WithEvents tvwRanks As UITreeView
	Private WithEvents btnAdd As UIButton
	Private WithEvents btnDelete As UIButton
	Private WithEvents btnUp As UIButton
	Private WithEvents btnDown As UIButton
	Private WithEvents btnUpdate As UIButton

	Private fraPermissions As UIWindow
	Private WithEvents vscrPerms As UIScrollBar
	Private chkRule() As UICheckBox

	'fraRules...
	Private fraRules As UIWindow
	Private chkRequirePeace As UICheckBox
	Private chkAutoTrade As UICheckBox
	Private chkShareVision As UICheckBox
    'Private chkGuildRels As UICheckBox
	Private cboVoteWeight As UIComboBox
	Private lblVoteWeight As UILabel
	Private lblTaxInterval As UILabel
	Private cboTaxInterval As UIComboBox
	Private lblTaxDay As UILabel
	Private txtTaxDay As UITextBox
	Private lblTaxMonth As UILabel
	Private txtTaxMonth As UITextBox
	Private chkDemoteRA As UICheckBox
	Private chkAcceptRA As UICheckBox
	Private chkPromoteRA As UICheckBox
	Private chkRemoveRA As UICheckBox
	Private chkCreateRankRA As UICheckBox
	Private chkDeleteRankRA As UICheckBox
	Private chkChangeVoteRA As UICheckBox
	Private WithEvents btnUpdateRules As UIButton

	Private mbLoading As Boolean = True

	Private moGuild As Guild = Nothing
	Private mlGuildFormer As Int32 = -1

	Private rcRules As Rectangle = New Rectangle(402, 35, 50, 21)
	Private rcRanks As Rectangle = New Rectangle(452, 35, 50, 21)

	Private myViewState As Byte = 0		'rules = 0, ranks = 1

	Private mlPermVals() As Int32
	Private msPermVals() As String
	Private Sub FillPermVals()
        ReDim mlPermVals(24)
        ReDim msPermVals(24)
        mlPermVals(0) = RankPermissions.AcceptApplicant : msPermVals(0) = "Accept Applicant"
        mlPermVals(1) = RankPermissions.AcceptEvents : msPermVals(1) = "Accept Events"
        'mlPermVals(2) = RankPermissions.BuildGuildBase : msPermVals(2) = "Build Guild Base"
        mlPermVals(2) = RankPermissions.ChangeMOTD : msPermVals(2) = "Change Message of the Day"
        mlPermVals(3) = RankPermissions.ChangeRankNames : msPermVals(3) = "Change Rank Names"
        mlPermVals(4) = RankPermissions.ChangeRankPermissions : msPermVals(4) = "Change Rank Permissions"
        mlPermVals(5) = RankPermissions.ChangeRankVotingWeight : msPermVals(5) = "Change Rank Voting Weight"
        mlPermVals(6) = RankPermissions.ChangeRecruitment : msPermVals(6) = "Change Recruitment"
        mlPermVals(7) = RankPermissions.CreateEvents : msPermVals(7) = "Create Events"
        mlPermVals(8) = RankPermissions.CreateRanks : msPermVals(8) = "Create/Move Ranks"
        mlPermVals(9) = RankPermissions.DeleteEvents : msPermVals(9) = "Delete Events"
        mlPermVals(10) = RankPermissions.DeleteRanks : msPermVals(10) = "Delete Ranks"
        mlPermVals(11) = RankPermissions.DemoteMember : msPermVals(11) = "Demote Member Rank"
        'mlPermVals(12) = RankPermissions.GuildChatChannel : msPermVals(12) = "Guild Chat Channel Access"
        mlPermVals(12) = RankPermissions.InviteMember : msPermVals(12) = "Invite Players To Join"
        mlPermVals(13) = RankPermissions.ModifyGuildRelation : msPermVals(13) = "Modify Guild Relationships"
        mlPermVals(14) = RankPermissions.PromoteMember : msPermVals(14) = "Promote Member Rank"
        mlPermVals(15) = RankPermissions.ProposeVotes : msPermVals(15) = "Propose New Votes"
        mlPermVals(16) = RankPermissions.RejectMember : msPermVals(16) = "Reject Applicant"
        mlPermVals(17) = RankPermissions.RemoveMember : msPermVals(17) = "Remove Member From Guild"
        mlPermVals(18) = RankPermissions.ViewBankLog : msPermVals(18) = "View Bank Log"
        mlPermVals(19) = RankPermissions.ViewEvents : msPermVals(19) = "View Events"
        mlPermVals(20) = RankPermissions.ViewEventAttachments : msPermVals(20) = "View Event Attachments"
        mlPermVals(21) = RankPermissions.ViewContentsLowSec : msPermVals(21) = "View Guild Treasury" 'Low Security Bank Contents"
        'mlPermVals(22) = RankPermissions.ViewContentsMedSec : msPermVals(22) = "View Medium Security Bank Contents"
        'mlPermVals(24) = RankPermissions.ViewContentsHiSec : msPermVals(24) = "View High Security Bank Contents"
        'mlPermVals(25) = RankPermissions.ViewGuildBase : msPermVals(25) = "View Guild Bank Facility"
        mlPermVals(22) = RankPermissions.ViewVotesHistory : msPermVals(22) = "View Vote History"
        mlPermVals(23) = RankPermissions.ViewVotesInProgress : msPermVals(23) = "View Votes In Progress"
        mlPermVals(24) = RankPermissions.WithdrawLowSec : msPermVals(24) = "Withdraw Guild Treasury" 'Low Security Bank Contents"
        'mlPermVals(28) = RankPermissions.WithdrawMedSec : msPermVals(28) = "Withdraw Medium Security Bank Contents"
        'mlPermVals(29) = RankPermissions.WithdrawHiSec : msPermVals(29) = "Withdraw High Security Bank Contents"

	End Sub

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'frmPlayerSetup initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eGuildSetup
            .ControlName = "frmGuildSetup"
            .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 245
            .Width = 512 '256
            .Height = 512 '410
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
        End With

		'lblTitle initial props
		lblTitle = New UILabel(oUILib)
		With lblTitle
			.ControlName = "lblTitle"
			.Left = 5
			.Top = 5
			.Width = 179
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Guild Initial Setup"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 14.25F, FontStyle.Bold Or FontStyle.Italic, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTitle, UIControl))

		'NewControl8 initial props
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

		'lblGuildName initial props
		lblGuildName = New UILabel(oUILib)
		With lblGuildName
			.ControlName = "lblGuildName"
			.Left = 10
			.Top = 35
			.Width = 85
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Guild Name:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.ToolTipText = "This is what other players will see."
		End With
		Me.AddChild(CType(lblGuildName, UIControl))

		'txtGuildName initial props
		txtGuildName = New UITextBox(oUILib)
		With txtGuildName
			.ControlName = "txtGuildName"
			.Left = 95
			.Top = 35
			.Width = 155
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
			.MaxLength = 50
			.BorderColor = muSettings.InterfaceBorderColor
			.ToolTipText = "This is what other players will see."
		End With
		Me.AddChild(CType(txtGuildName, UIControl))

		'btnCheck initial props
		btnCheck = New UIButton(oUILib)
		With btnCheck
			.ControlName = "btnCheck"
			.Left = Me.Width \ 4 - 75
			.Top = 60 '440
			.Width = 150
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Check Availability"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnCheck, UIControl))

		'fraIconChooser initial props
		fraIconChooser = New frmIconChooser(oUILib, False)
		With fraIconChooser
			.Left = 10
			.Top = btnCheck.Top + btnCheck.Height + 10
			.Width = 250
			.Height = 250
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
		End With
		Me.AddChild(CType(fraIconChooser, UIControl))

		fraInvite = New UIWindow(oUILib)
		With fraInvite
			.ControlName = "fraInvite"
			.Left = 10
			.Top = fraIconChooser.Top + fraIconChooser.Height + 10
			.Width = 245
			.Height = 125
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			.BorderLineWidth = 2
			.Caption = "Four Other Players Must Accept"
		End With
		Me.AddChild(CType(fraInvite, UIControl))

		'lblPlayerName initial props
		lblPlayerName = New UILabel(oUILib)
		With lblPlayerName
			.ControlName = "lblPlayerName"
			.Left = 5
			.Top = 5
			.Width = 90
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Player Name"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraInvite.AddChild(CType(lblPlayerName, UIControl))

		'txtInvite1 initial props
		txtInvite1 = New UITextBox(oUILib)
		With txtInvite1
			.ControlName = "txtInvite1"
			.Left = 5
			.Top = 25
			.Width = 145
			.Height = 20
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
		fraInvite.AddChild(CType(txtInvite1, UIControl))

		'txtInvite2 initial props
		txtInvite2 = New UITextBox(oUILib)
		With txtInvite2
			.ControlName = "txtInvite2"
			.Left = 5
			.Top = 50
			.Width = 145
			.Height = 20
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
		fraInvite.AddChild(CType(txtInvite2, UIControl))

		'txtInvite3 initial props
		txtInvite3 = New UITextBox(oUILib)
		With txtInvite3
			.ControlName = "txtInvite3"
			.Left = 5
			.Top = 75
			.Width = 145
			.Height = 20
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
		fraInvite.AddChild(CType(txtInvite3, UIControl))

		'txtInvite4 initial props
		txtInvite4 = New UITextBox(oUILib)
		With txtInvite4
			.ControlName = "txtInvite4"
			.Left = 5
			.Top = 100
			.Width = 145
			.Height = 20
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
		fraInvite.AddChild(CType(txtInvite4, UIControl))

		'chkAccept1 initial props
		chkAccept1 = New UICheckBox(oUILib)
		With chkAccept1
			.ControlName = "chkAccept1"
			.Left = 230	'195
			.Top = 28
			.Width = 18
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
			.Value = False
			.Locked = True
		End With
		fraInvite.AddChild(CType(chkAccept1, UIControl))

		'chkAccept2 initial props
		chkAccept2 = New UICheckBox(oUILib)
		With chkAccept2
			.ControlName = "chkAccept2"
			.Left = 230	'195
			.Top = 53
			.Width = 18
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
			.Value = False
			.Locked = True
		End With
		fraInvite.AddChild(CType(chkAccept2, UIControl))

		'chkAccept3 initial props
		chkAccept3 = New UICheckBox(oUILib)
		With chkAccept3
			.ControlName = "chkAccept3"
			.Left = 230	'195
			.Top = 78
			.Width = 18
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
			.Value = False
			.Locked = True
		End With
		fraInvite.AddChild(CType(chkAccept3, UIControl))

		'chkAccept4 initial props
		chkAccept4 = New UICheckBox(oUILib)
		With chkAccept4
			.ControlName = "chkAccept4"
			.Left = 230	'195
			.Top = 103
			.Width = 18
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
			.Value = False
			.Locked = True
		End With
		fraInvite.AddChild(CType(chkAccept4, UIControl))

		'btnInvite1 initial props
		btnInvite1 = New UIButton(oUILib)
		With btnInvite1
			.ControlName = "btnInvite1"
			.Left = 150
			.Top = 25
			.Width = 77
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "Invite"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraInvite.AddChild(CType(btnInvite1, UIControl))

		'btnInvite2 initial props
		btnInvite2 = New UIButton(oUILib)
		With btnInvite2
			.ControlName = "btnInvite2"
			.Left = 150
			.Top = 50
			.Width = 77
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "Invite"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraInvite.AddChild(CType(btnInvite2, UIControl))

		'btnInvite3 initial props
		btnInvite3 = New UIButton(oUILib)
		With btnInvite3
			.ControlName = "btnInvite3"
			.Left = 150
			.Top = 75
			.Width = 77
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "Invite"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraInvite.AddChild(CType(btnInvite3, UIControl))

		'btnInvite4 initial props
		btnInvite4 = New UIButton(oUILib)
		With btnInvite4
			.ControlName = "btnInvite4"
			.Left = 150
			.Top = 100
			.Width = 77
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "Invite"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraInvite.AddChild(CType(btnInvite4, UIControl))

		'btnSubmit initial props
		btnSubmit = New UIButton(oUILib)
		With btnSubmit
			.ControlName = "btnSubmit"
			.Left = 6
			.Top = fraInvite.Top + fraInvite.Height + 5
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

		'btnCancel initial props
		btnCancel = New UIButton(oUILib)
		With btnCancel
			.ControlName = "btnCancel"
			.Left = 150
			.Top = btnSubmit.Top
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Cancel"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnCancel, UIControl))


		'================= fraRanks ================

		fraRanks = New UIWindow(oUILib)
		With fraRanks
			.ControlName = "fraRanks"
			.Left = 262
			.Top = 65
			.Width = 245
			.Height = 415
			.Enabled = False
			.Visible = False
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			.BorderLineWidth = 2
			.Caption = "Rank Configuration"
		End With
		Me.AddChild(CType(fraRanks, UIControl))

		'btnAdd initial props
		btnAdd = New UIButton(oUILib)
		With btnAdd
			.ControlName = "btnAdd"
			.Left = 5
			.Top = 385
			.Width = 100
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
		fraRanks.AddChild(CType(btnAdd, UIControl))

		'btnDelete initial props
		btnDelete = New UIButton(oUILib)
		With btnDelete
			.ControlName = "btnDelete"
			.Left = 218
			.Top = 80
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
		fraRanks.AddChild(CType(btnDelete, UIControl))

		'btnUp initial props
		btnUp = New UIButton(oUILib)
		With btnUp
			.ControlName = "btnUp"
			.Left = 218
			.Top = 10
			.Width = 24
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Up"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraRanks.AddChild(CType(btnUp, UIControl))

		'btnDown initial props
		btnDown = New UIButton(oUILib)
		With btnDown
			.ControlName = "btnDown"
			.Left = 218
			.Top = 37
			.Width = 24
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Dn"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraRanks.AddChild(CType(btnDown, UIControl))

		'txtRankName initial props
		txtRankName = New UITextBox(oUILib)
		With txtRankName
			.ControlName = "txtRankName"
			.Left = 5
			.Top = 130
			.Width = 150
			.Height = 20
			.Enabled = True
			.Visible = True
			.Caption = "Rank Name"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		fraRanks.AddChild(CType(txtRankName, UIControl))

		'btnUpdate initial props
		btnUpdate = New UIButton(oUILib)
		With btnUpdate
			.ControlName = "btnUpdate"
			.Left = 142
			.Top = 385
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Update"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraRanks.AddChild(CType(btnUpdate, UIControl))

		'txtVoteRate initial props
		txtVoteRate = New UITextBox(oUILib)
		With txtVoteRate
			.ControlName = "txtVoteRate"
			.Left = 170
			.Top = 130
			.Width = 49
			.Height = 20
			.Enabled = True
			.Visible = True
			.Caption = "1"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 5
			.bNumericOnly = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		fraRanks.AddChild(CType(txtVoteRate, UIControl))

		'lblVoteStr initial props
		lblVoteStr = New UILabel(oUILib)
		With lblVoteStr
			.ControlName = "lblVoteStr"
			.Left = 170
			.Top = 105
			.Width = 60
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Vote Rate"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraRanks.AddChild(CType(lblVoteStr, UIControl))

		'lblRankName initial props
		lblRankName = New UILabel(oUILib)
		With lblRankName
			.ControlName = "lblRankName"
			.Left = 5
			.Top = 105
			.Width = 76
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Rank Name"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraRanks.AddChild(CType(lblRankName, UIControl))

		'lblTaxFlat initial props
		lblTaxFlat = New UILabel(oUILib)
		With lblTaxFlat
			.ControlName = "lblTaxFlat"
			.Left = 5
			.Top = 150
			.Width = 76
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Flat Rate"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraRanks.AddChild(CType(lblTaxFlat, UIControl))

		'txtFlatRate initial props
		txtFlatRate = New UITextBox(oUILib)
		With txtFlatRate
			.ControlName = "txtFlatRate"
			.Left = 5
			.Top = 175
			.Width = 80
			.Height = 20
			.Enabled = True
			.Visible = True
			.Caption = "0"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 10
			.bNumericOnly = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		fraRanks.AddChild(CType(txtFlatRate, UIControl))

		'lblTaxPerc initial props
		lblTaxPerc = New UILabel(oUILib)
		With lblTaxPerc
			.ControlName = "lblTaxPerc"
			.Left = 90
			.Top = 150
			.Width = 89
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Percentage Of"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraRanks.AddChild(CType(lblTaxPerc, UIControl))

		'txtTaxPerc initial props
		txtTaxPerc = New UITextBox(oUILib)
		With txtTaxPerc
			.ControlName = "txtTaxPerc"
			.Left = 90
			.Top = 175
			.Width = 30
			.Height = 20
			.Enabled = True
			.Visible = True
			.Caption = "100"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 3
			.bNumericOnly = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		fraRanks.AddChild(CType(txtTaxPerc, UIControl))

		'tvwRanks initial props
		tvwRanks = New UITreeView(oUILib)
		With tvwRanks
			.ControlName = "tvwRanks"
			.mbAcceptReprocessEvents = True
			.Left = 5
			.Top = 10
			.Width = 210
			.Height = 95
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
		End With
		fraRanks.AddChild(CType(tvwRanks, UIControl))

		'fraPermissions initial props
		fraPermissions = New UIWindow(oUILib)
		With fraPermissions
			.ControlName = "fraPermissions"
			.Left = 5
			.Top = 203
			.Width = 235
			.Height = 175
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			.BorderLineWidth = 2
		End With
		fraRanks.AddChild(CType(fraPermissions, UIControl))

		'vscrPerms initial props
		vscrPerms = New UIScrollBar(oUILib, True)
		With vscrPerms
			.ControlName = "vscrPerms"
			.Left = 210
			.Top = 4
			.Width = 24
			.Height = 170
			.Enabled = True
			.Visible = True
			.Value = 0
			.MaxValue = 100
			.MinValue = 0
			.SmallChange = 1
			.LargeChange = 1
			.ReverseDirection = True
		End With
		fraPermissions.AddChild(CType(vscrPerms, UIControl))

		'chkRule initial props
		Dim lRuleCnt As Int32 = CInt((fraPermissions.Height - 20) / 20.0F)
        vscrPerms.MaxValue = 24 - lRuleCnt
		ReDim chkRule(lRuleCnt - 1)
		For X As Int32 = 0 To lRuleCnt - 1
			chkRule(X) = New UICheckBox(oUILib)
			With chkRule(X)
				.ControlName = X.ToString
				.Left = 10
				.Top = 10 + (X * 20)
				.Width = 100
				.Height = 18
				.Enabled = True
				.Visible = False
				.Caption = "Check Caption"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraPermissions.AddChild(CType(chkRule(X), UIControl))
			AddHandler chkRule(X).Click, AddressOf chkRuleClick
		Next X

		'=========== fraRules =============
		fraRules = New UIWindow(oUILib)
		With fraRules
			.ControlName = "fraRules"
			.Left = 262
			.Top = 65
			.Width = 245
			.Height = 350
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			.BorderLineWidth = 2
			.Caption = "Initial Rules"
		End With
		Me.AddChild(CType(fraRules, UIControl))

		'chkRequirePeace initial props
		chkRequirePeace = New UICheckBox(oUILib)
		With chkRequirePeace
			.ControlName = "chkRequirePeace"
			.Left = 10
			.Top = 100
			.Width = 179
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Require Peace Between Members"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraRules.AddChild(CType(chkRequirePeace, UIControl))

		'chkAutoTrade initial props
		chkAutoTrade = New UICheckBox(oUILib)
		With chkAutoTrade
			.ControlName = "chkAutoTrade"
			.Left = 10
			.Top = 118
			.Width = 187
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Automatic Trade Between Members"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraRules.AddChild(CType(chkAutoTrade, UIControl))

		'chkShareVision initial props
		chkShareVision = New UICheckBox(oUILib)
		With chkShareVision
			.ControlName = "chkShareVision"
			.Left = 10
			.Top = 136
			.Width = 136
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Share Unit/Facility Vision"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraRules.AddChild(CType(chkShareVision, UIControl))

        ''chkGuildRels initial props
        'chkGuildRels = New UICheckBox(oUILib)
        'With chkGuildRels
        '	.ControlName = "chkGuildRels"
        '	.Left = 10
        '	.Top = 154
        '	.Width = 120
        '	.Height = 18
        '	.Enabled = True
        '	.Visible = True
        '	.Caption = "Unified Foreign Policy"
        '	.ForeColor = muSettings.InterfaceBorderColor
        '	.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '	.DrawBackImage = False
        '	.FontFormat = CType(6, DrawTextFormat)
        '	.Value = False
        'End With
        'fraRules.AddChild(CType(chkGuildRels, UIControl))

		'lblVoteWeight initial props
		lblVoteWeight = New UILabel(oUILib)
		With lblVoteWeight
			.ControlName = "lblVoteWeight"
			.Left = 10
			.Top = 10
			.Width = 111
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Vote Weight Based On:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraRules.AddChild(CType(lblVoteWeight, UIControl))

		'lblTaxInterval initial props
		lblTaxInterval = New UILabel(oUILib)
		With lblTaxInterval
			.ControlName = "lblTaxInterval"
			.Left = 10
			.Top = 40
			.Width = 60
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Tax Interval:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraRules.AddChild(CType(lblTaxInterval, UIControl))

		'lblTaxDay initial props
		lblTaxDay = New UILabel(oUILib)
		With lblTaxDay
			.ControlName = "lblTaxDay"
			.Left = 10
			.Top = 70
			.Width = 60
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Tax Day:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraRules.AddChild(CType(lblTaxDay, UIControl))

		'txtTaxDay initial props
		txtTaxDay = New UITextBox(oUILib)
		With txtTaxDay
			.ControlName = "txtTaxDay"
			.Left = 65
			.Top = 72
			.Width = 25
			.Height = 20
			.Enabled = True
			.Visible = True
			.Caption = "30"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 2
			.bNumericOnly = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		fraRules.AddChild(CType(txtTaxDay, UIControl))

		'lblTaxMonth initial props
		lblTaxMonth = New UILabel(oUILib)
		With lblTaxMonth
			.ControlName = "lblTaxMonth"
			.Left = 135
			.Top = 70
			.Width = 60
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Tax Month:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraRules.AddChild(CType(lblTaxMonth, UIControl))

		'txtTaxMonth initial props
		txtTaxMonth = New UITextBox(oUILib)
		With txtTaxMonth
			.ControlName = "txtTaxMonth"
			.Left = 200
			.Top = 72
			.Width = 25
			.Height = 20
			.Enabled = True
			.Visible = True
			.Caption = "1"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 2
			.bNumericOnly = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		fraRules.AddChild(CType(txtTaxMonth, UIControl))

		'chkDemoteRA initial props
		chkDemoteRA = New UICheckBox(oUILib)
		With chkDemoteRA
			.ControlName = "chkDemoteRA"
			.Left = 10
			.Top = 172
			.Width = 160
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Demote Member By Vote Only"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraRules.AddChild(CType(chkDemoteRA, UIControl))

		'chkAcceptRA initial props
		chkAcceptRA = New UICheckBox(oUILib)
		With chkAcceptRA
			.ControlName = "chkAcceptRA"
			.Left = 10
			.Top = 191
			.Width = 157
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Accept Member By Vote Only"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraRules.AddChild(CType(chkAcceptRA, UIControl))

		'chkPromoteRA initial props
		chkPromoteRA = New UICheckBox(oUILib)
		With chkPromoteRA
			.ControlName = "chkPromoteRA"
			.Left = 10
			.Top = 208
			.Width = 162
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Promote Member By Vote Only"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraRules.AddChild(CType(chkPromoteRA, UIControl))

		'chkRemoveRA initial props
		chkRemoveRA = New UICheckBox(oUILib)
		With chkRemoveRA
			.ControlName = "chkRemoveRA"
			.Left = 10
			.Top = 226
			.Width = 163
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Remove Member By Vote Only"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraRules.AddChild(CType(chkRemoveRA, UIControl))

		'chkCreateRankRA initial props
		chkCreateRankRA = New UICheckBox(oUILib)
		With chkCreateRankRA
			.ControlName = "chkCreateRankRA"
			.Left = 10
			.Top = 244
			.Width = 142
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Create Rank By Vote Only"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraRules.AddChild(CType(chkCreateRankRA, UIControl))

		'chkDeleteRankRA initial props
		chkDeleteRankRA = New UICheckBox(oUILib)
		With chkDeleteRankRA
			.ControlName = "chkDeleteRankRA"
			.Left = 10
			.Top = 262
			.Width = 142
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Delete Rank By Vote Only"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraRules.AddChild(CType(chkDeleteRankRA, UIControl))

		'chkChangeVoteRA initial props
		chkChangeVoteRA = New UICheckBox(oUILib)
		With chkChangeVoteRA
			.ControlName = "chkChangeVoteRA"
			.Left = 10
			.Top = 280
			.Width = 189
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Change Voting Weight By Vote Only"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraRules.AddChild(CType(chkChangeVoteRA, UIControl))

		'btnUpdateRules initial props
		btnUpdateRules = New UIButton(oUILib)
		With btnUpdateRules
			.ControlName = "btnUpdateRules"
			.Left = fraRules.Width \ 2 - 60
			.Top = 320
			.Width = 120
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Update Rules"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraRules.AddChild(CType(btnUpdateRules, UIControl))

		'cboTaxPercType initial props
		cboTaxPercType = New UIComboBox(oUILib)
		With cboTaxPercType
			.ControlName = "cboTaxPercType"
			.Left = 125
			.Top = 175
			.Width = 115
			.Height = 20
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
        fraRanks.AddChild(CType(cboTaxPercType, UIControl))

        'cboTaxInterval initial props
        cboTaxInterval = New UIComboBox(oUILib)
        With cboTaxInterval
            .ControlName = "cboTaxInterval"
            .Left = 130
            .Top = 42
            .Width = 110
            .Height = 20
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
        fraRules.AddChild(CType(cboTaxInterval, UIControl))

        'cboVoteWeight initial props
        cboVoteWeight = New UIComboBox(oUILib)
        With cboVoteWeight
            .ControlName = "cboVoteWeight"
            .Left = 130
            .Top = 12
            .Width = 110
            .Height = 20
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
        End With
		fraRules.AddChild(CType(cboVoteWeight, UIControl))

		FillStaticData()

		moGuild = New Guild
		mlGuildFormer = glPlayerID
		moGuild.yState = eyGuildState.Forming
		moGuild.sName = "Unknown"
		moGuild.lBaseGuildRules = 0
		moGuild.iRecruitFlags = 0
		Dim oRank As New GuildRank()
		With oRank
			.lRankID = -1
            .lRankPermissions = RankPermissions.AcceptApplicant Or RankPermissions.AcceptEvents Or RankPermissions.BuildGuildBase Or _
             RankPermissions.ChangeMOTD Or RankPermissions.ChangeRankNames Or RankPermissions.ChangeRankPermissions Or _
             RankPermissions.ChangeRankVotingWeight Or RankPermissions.ChangeRecruitment Or RankPermissions.CreateEvents Or _
             RankPermissions.CreateRanks Or RankPermissions.DeleteEvents Or RankPermissions.DeleteRanks Or RankPermissions.DemoteMember Or _
             RankPermissions.InviteMember Or RankPermissions.ModifyGuildRelation Or _
             RankPermissions.PromoteMember Or RankPermissions.ProposeVotes Or RankPermissions.RejectMember Or RankPermissions.RemoveMember Or _
             RankPermissions.ViewBankLog Or RankPermissions.ViewContentsHiSec Or RankPermissions.ViewContentsLowSec Or _
             RankPermissions.ViewEventAttachments Or RankPermissions.ViewEvents Or RankPermissions.ViewGuildBase Or _
             RankPermissions.ViewVotesHistory Or RankPermissions.ViewVotesInProgress Or RankPermissions.WithdrawHiSec Or _
             RankPermissions.WithdrawLowSec
			.lVoteStrength = 1
			.sRankName = "Founder"
			.TaxRateFlat = 0
			.TaxRatePercentage = 0
			.TaxRatePercType = eyGuildTaxPercType.CashFlow
			.yPosition = 0
		End With
		ReDim moGuild.moRanks(0)
		moGuild.moRanks(0) = oRank

		ReDim moGuild.moMembers(0)
		moGuild.moMembers(0) = New GuildMember
		With moGuild.moMembers(0)
			.lMemberID = glPlayerID
			.lRankID = oRank.lRankID
			.yMemberState = GuildMemberState.Invited Or GuildMemberState.AcceptedGuildFormInvite Or GuildMemberState.Approved
		End With


		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)

		mbLoading = False 
	End Sub

	Private Sub FillStaticData()
		With cboTaxPercType
			.Clear()
			.AddItem("Cash Flow") : .ItemData(.NewIndex) = eyGuildTaxPercType.CashFlow
			.AddItem("Total Income") : .ItemData(.NewIndex) = eyGuildTaxPercType.TotalIncome
			.AddItem("Treasury") : .ItemData(.NewIndex) = eyGuildTaxPercType.Treasury
		End With
		With cboVoteWeight
			.Clear()
			.AddItem("1 Per Member") : .ItemData(.NewIndex) = eyVoteWeightType.StaticValue
			.AddItem("Population") : .ItemData(.NewIndex) = eyVoteWeightType.PopulationBased
			.AddItem("Rank-Based") : .ItemData(.NewIndex) = eyVoteWeightType.RankBased
			.AddItem("Seniority") : .ItemData(.NewIndex) = eyVoteWeightType.AgeOfPlayer
		End With
		With cboTaxInterval
			.Clear()
			.AddItem("Annually") : .ItemData(.NewIndex) = eyGuildInterval.Annually
			.AddItem("Daily") : .ItemData(.NewIndex) = eyGuildInterval.Daily
			.AddItem("Every Two Months") : .ItemData(.NewIndex) = eyGuildInterval.EveryTwoMonths
			.AddItem("Monthly") : .ItemData(.NewIndex) = eyGuildInterval.Monthly
			.AddItem("Quarterly") : .ItemData(.NewIndex) = eyGuildInterval.Quarterly
			.AddItem("Semi-Annually") : .ItemData(.NewIndex) = eyGuildInterval.SemiAnnually
			.AddItem("Twice A Month") : .ItemData(.NewIndex) = eyGuildInterval.SemiMonthly
			.AddItem("Weekly") : .ItemData(.NewIndex) = eyGuildInterval.Weekly
		End With
	End Sub

	Private Sub frmGuildSetup_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
		Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
		Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y

		'Now, check for what the player is hovering over
		If rcRules.Contains(lX, lY) = True Then
			myViewState = 0
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
			Me.IsDirty = True
			fraRules.Visible = True
			fraRules.Enabled = True
			fraRanks.Visible = False
			fraRanks.Enabled = False
		ElseIf rcRanks.Contains(lX, lY) = True Then
			myViewState = 1
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
			Me.IsDirty = True
			fraRules.Visible = False
			fraRules.Enabled = False
			fraRanks.Visible = True
			fraRanks.Enabled = True
		End If

	End Sub

	Private Sub frmGuildSetup_OnNewFrame() Handles Me.OnNewFrame
		If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
			MyBase.moUILib.RemoveWindow(Me.ControlName)
			Dim ofrm As New frmGuildMain(MyBase.moUILib)
			Return
		End If

		If moGuild Is Nothing Then Return

		'ok, check our values versus the guild object
		If mlGuildFormer <> glPlayerID Then
			If txtGuildName.Caption <> moGuild.sName Then txtGuildName.Caption = moGuild.sName
			Dim lCurr As Int32 = fraIconChooser.GetStartingIcon
			If moGuild.lIcon <> lCurr Then fraIconChooser.SetStartingIcon(moGuild.lIcon)

			Dim bValue As Boolean = (moGuild.lBaseGuildRules And elGuildFlags.RequirePeaceBetweenMembers) <> 0
			If chkRequirePeace.Value <> bValue Then chkRequirePeace.Value = bValue
			bValue = (moGuild.lBaseGuildRules And elGuildFlags.AutomaticTradeAgreements) <> 0
			If chkAutoTrade.Value <> bValue Then chkAutoTrade.Value = bValue
			bValue = (moGuild.lBaseGuildRules And elGuildFlags.ShareUnitVision) <> 0
			If chkShareVision.Value <> bValue Then chkShareVision.Value = bValue
            'bValue = (moGuild.lBaseGuildRules And elGuildFlags.UnifiedForeignPolicy) <> 0
            'If chkGuildRels.Value <> bValue Then chkGuildRels.Value = bValue

			If cboVoteWeight.ListIndex > -1 Then
				If cboVoteWeight.ItemData(cboVoteWeight.ListIndex) <> moGuild.yVoteWeightType Then cboVoteWeight.FindComboItemData(moGuild.yVoteWeightType)
			Else
				cboVoteWeight.FindComboItemData(moGuild.yVoteWeightType)
			End If
			If cboTaxInterval.ListIndex > -1 Then
				If cboTaxInterval.ItemData(cboTaxInterval.ListIndex) <> moGuild.yGuildTaxRateInterval Then cboTaxInterval.FindComboItemData(moGuild.yGuildTaxRateInterval)
			Else
				cboTaxInterval.FindComboItemData(moGuild.yGuildTaxRateInterval)
			End If
			If txtTaxDay.Caption <> moGuild.yGuildTaxBaseDay.ToString Then txtTaxDay.Caption = moGuild.yGuildTaxBaseDay.ToString
			If txtTaxMonth.Caption <> moGuild.yGuildTaxBaseMonth.ToString Then txtTaxMonth.Caption = moGuild.yGuildTaxBaseMonth.ToString

			bValue = (moGuild.lBaseGuildRules And elGuildFlags.DemoteMember_RA) <> 0
			If chkDemoteRA.Value <> bValue Then chkDemoteRA.Value = bValue
			bValue = (moGuild.lBaseGuildRules And elGuildFlags.AcceptMemberToGuild_RA) <> 0
			If chkAcceptRA.Value <> bValue Then chkAcceptRA.Value = bValue
			bValue = (moGuild.lBaseGuildRules And elGuildFlags.PromoteGuildMember_RA) <> 0
			If chkPromoteRA.Value <> bValue Then chkPromoteRA.Value = bValue
			bValue = (moGuild.lBaseGuildRules And elGuildFlags.RemoveGuildMember_RA) <> 0
			If chkRemoveRA.Value <> bValue Then chkRemoveRA.Value = bValue
			bValue = (moGuild.lBaseGuildRules And elGuildFlags.CreateRank_RA) <> 0
			If chkCreateRankRA.Value <> bValue Then chkCreateRankRA.Value = bValue
			bValue = (moGuild.lBaseGuildRules And elGuildFlags.DeleteRank_RA) <> 0
			If chkDeleteRankRA.Value <> bValue Then chkDeleteRankRA.Value = bValue
			bValue = (moGuild.lBaseGuildRules And elGuildFlags.ChangeVotingWeight_RA) <> 0
			If chkChangeVoteRA.Value <> bValue Then chkChangeVoteRA.Value = bValue
		End If

		If moGuild Is Nothing = False Then
			If moGuild.moMembers Is Nothing = False Then
				For X As Int32 = 0 To moGuild.moMembers.GetUpperBound(0)
					If moGuild.moMembers(X) Is Nothing = False Then
						Dim oNode As UITreeView.UITreeViewItem = tvwRanks.GetNodeByItemData2(moGuild.moMembers(X).lMemberID, ObjectType.ePlayer)
						If oNode Is Nothing Then Continue For
						Dim sName As String = GetCacheObjectValue(moGuild.moMembers(X).lMemberID, ObjectType.ePlayer)
						If oNode.sItem <> sName Then
							oNode.sItem = sName
							Me.IsDirty = True
						End If

						If mlInviteIDs Is Nothing = False Then
							For Y As Int32 = 0 To 3
								If mlInviteIDs(Y) = moGuild.moMembers(X).lMemberID Then
									Select Case Y
										Case 0
											If txtInvite1.Caption <> sName Then txtInvite1.Caption = sName
										Case 1
											If txtInvite2.Caption <> sName Then txtInvite2.Caption = sName
										Case 2
											If txtInvite3.Caption <> sName Then txtInvite3.Caption = sName
										Case 3
											If txtInvite4.Caption <> sName Then txtInvite4.Caption = sName
									End Select
									Exit For
								End If
							Next Y
						End If

					End If
				Next X
			End If
		End If

	End Sub

	Private Sub FillRankNodes()
		tvwRanks.Clear()
		If moGuild Is Nothing = False Then
			If moGuild.moRanks Is Nothing = False Then
				For X As Int32 = 0 To moGuild.moRanks.GetUpperBound(0)
					If moGuild.moRanks(X) Is Nothing = False Then
						tvwRanks.AddNode(moGuild.moRanks(X).sRankName, moGuild.moRanks(X).lRankID, -1, -1, Nothing, Nothing)
					End If
				Next X

				'now, our members
				If moGuild.moMembers Is Nothing = False Then
					For X As Int32 = 0 To moGuild.moMembers.GetUpperBound(0)
						If moGuild.moMembers(X) Is Nothing = False Then
							Dim oParent As UITreeView.UITreeViewItem = tvwRanks.GetNodeByItemData2(moGuild.moMembers(X).lRankID, -1)
							moGuild.moMembers(X).sPlayerName = GetCacheObjectValue(moGuild.moMembers(X).lMemberID, ObjectType.ePlayer)
							tvwRanks.AddNode(moGuild.moMembers(X).sPlayerName, moGuild.moMembers(X).lMemberID, ObjectType.ePlayer, -1, oParent, Nothing)
						End If
					Next X
				End If
			End If
		End If
	End Sub

	Private Sub frmGuildSetup_OnRenderEnd() Handles Me.OnRenderEnd
		fraIconChooser.frmIconChooser_OnRenderEnd()

		If myViewState = 0 Then
			Dim oSelColor As System.Drawing.Color
			oSelColor = System.Drawing.Color.FromArgb(255, 128, 128, 160)
			MyBase.moUILib.DoAlphaBlendColorFill(rcRules, oSelColor, rcRules.Location)
		Else
			Dim oSelColor As System.Drawing.Color
			oSelColor = System.Drawing.Color.FromArgb(255, 128, 128, 160)
			MyBase.moUILib.DoAlphaBlendColorFill(rcRanks, oSelColor, rcRanks.Location)
		End If

		MyBase.RenderRoundedBorder(rcRules, 1, muSettings.InterfaceBorderColor)
		MyBase.RenderRoundedBorder(rcRanks, 1, muSettings.InterfaceBorderColor)

        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return

		'Now, render our text...
		Using oFont As Font = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
				oTextSpr.Begin(SpriteFlags.AlphaBlend)
				Try
					oFont.DrawText(oTextSpr, "Rules", rcRules, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
					oFont.DrawText(oTextSpr, "Ranks", rcRanks, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
				Catch
				End Try
				oTextSpr.End()
				oTextSpr.Dispose()
			End Using
			oFont.Dispose()
		End Using
	End Sub

	Private Function ValidateData() As Boolean
		If fraIconChooser.GetStartingIcon() = 0 Then
			MyBase.moUILib.AddNotification("Please select a valid Icon.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return False
		End If

		If txtGuildName.Caption.Trim.Length > 50 Then
			MyBase.moUILib.AddNotification("Max length of guild name is 50 characters.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return False
		End If

		'Character checker
		Dim sTemp As String = txtGuildName.Caption
		For X As Int32 = 0 To sTemp.Length - 1
			Dim lChrVal As Int32 = Asc(Mid$(sTemp, X + 1, 1))
			If lChrVal <> 39 AndAlso lChrVal <> 44 AndAlso lChrVal <> 46 AndAlso lChrVal <> 96 AndAlso lChrVal <> 32 Then
				If (lChrVal > 64 AndAlso lChrVal < 91) = False AndAlso (lChrVal > 96 AndAlso lChrVal < 123) = False AndAlso (lChrVal > 47 AndAlso lChrVal < 58) = False Then
					MyBase.moUILib.AddNotification("Guild Name contains invalid characters. Alphabetical characters only.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return False
				End If
			End If
		Next X

		Return True
	End Function

	Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
		'Cancel creation or i am bowing, either way, notify all users
		SendGuildCreationAcceptance(255)
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnCheck_Click(ByVal sName As String) Handles btnCheck.Click
		If txtGuildName.Caption.Trim.Length < 3 Then
			MyBase.moUILib.AddNotification("Name of guild must be at least 3 characters.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		Dim yMsg(51) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eCheckGuildName).CopyTo(yMsg, 0)
		System.Text.ASCIIEncoding.ASCII.GetBytes(txtGuildName.Caption.Trim).CopyTo(yMsg, 2)
		MyBase.moUILib.SendMsgToPrimary(yMsg)
	End Sub

	Private Sub btnSubmit_Click(ByVal sName As String) Handles btnSubmit.Click

		If chkAccept1.Value = False OrElse chkAccept2.Value = False OrElse chkAccept3.Value = False OrElse chkAccept4.Value = False Then
			MyBase.moUILib.AddNotification("You must have 4 other players willing to form this guild with you.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
        End If

        If txtGuildName.Caption.Contains("*") = True Then
            MyBase.moUILib.AddNotification("Please enter a valid guild name.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

		Dim yMsg(55) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eCreateGuild).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(fraIconChooser.GetStartingIcon()).CopyTo(yMsg, lPos) : lPos += 4
		System.Text.ASCIIEncoding.ASCII.GetBytes(txtGuildName.Caption.Trim).CopyTo(yMsg, lPos) : lPos += 50
		MyBase.moUILib.SendMsgToPrimary(yMsg)
	End Sub

	Private Sub RemoveInvitee(ByVal sName As String)
		If glPlayerID <> mlGuildFormer Then Return
		Dim yMsg(22) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eInviteFormGuild).CopyTo(yMsg, 0)
		yMsg(2) = 128	'indicates removing invite
		System.Text.ASCIIEncoding.ASCII.GetBytes(sName).CopyTo(yMsg, 3)
		MyBase.moUILib.SendMsgToPrimary(yMsg)
	End Sub

	Private Sub InviteButtonClick(ByRef txtbox As UITextBox, ByRef btn As UIButton, ByVal chkbox As UICheckBox)
		If btn.Caption.ToUpper = "INVITE" Then
			If txtbox.Caption.Trim = "" Then
				MyBase.moUILib.AddNotification("Enter a player to invite.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If
			txtbox.Locked = True
			SendGuildUpdate()
			If SendInviteMsg(txtbox.Caption.Trim) = False Then
				btn.Caption = "Invite"
				txtbox.Locked = False
			Else
				btn.Caption = "Remove"
				MyBase.moUILib.AddNotification("Invitation Sent!", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			End If
		Else
			'remove the invitee... or member, whichever it may be
			RemoveInvitee(txtbox.Caption)
			btn.Caption = "Invite"
			txtbox.Locked = False
		End If
	End Sub

	Private Sub btnInvite1_Click(ByVal sName As String) Handles btnInvite1.Click
		InviteButtonClick(txtInvite1, btnInvite1, chkAccept1)
	End Sub

	Private Sub btnInvite2_Click(ByVal sName As String) Handles btnInvite2.Click
		InviteButtonClick(txtInvite2, btnInvite2, chkAccept2)
	End Sub

	Private Sub btnInvite3_Click(ByVal sName As String) Handles btnInvite3.Click
		InviteButtonClick(txtInvite3, btnInvite3, chkAccept3)
	End Sub

	Private Sub btnInvite4_Click(ByVal sName As String) Handles btnInvite4.Click
		InviteButtonClick(txtInvite4, btnInvite4, chkAccept4)
	End Sub

	Private Function SendInviteMsg(ByVal sName As String) As Boolean

		If sName.ToUpper.Trim = goCurrentPlayer.PlayerName.ToUpper.Trim Then
			MyBase.moUILib.AddNotification("You must invite someone other than yourself.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return False
		End If

		Dim yMsg(22) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eInviteFormGuild).CopyTo(yMsg, 0)
		yMsg(2) = 1	'indicates inviting
		System.Text.ASCIIEncoding.ASCII.GetBytes(sName).CopyTo(yMsg, 3)
		MyBase.moUILib.SendMsgToPrimary(yMsg)

		Return True
	End Function

	Public Sub PlayerInviteResponse(ByVal yType As Byte, ByVal sPlayerName As String, ByVal lPlayerID As Int32)
		Dim sPNameUpper As String = sPlayerName.ToUpper.Trim
		If yType = 255 Then
			If txtInvite1.Caption.Trim.ToUpper = sPNameUpper Then
				txtInvite1.Locked = False
				btnInvite1.Caption = "Invite"
				chkAccept1.Value = False
			End If
			If txtInvite2.Caption.Trim.ToUpper = sPNameUpper Then
				txtInvite2.Locked = False
				btnInvite2.Caption = "Invite"
				chkAccept2.Value = False
			End If
			If txtInvite3.Caption.Trim.ToUpper = sPNameUpper Then
				txtInvite2.Locked = False
				btnInvite2.Caption = "Invite"
				chkAccept2.Value = False
			End If
			If txtInvite4.Caption.Trim.ToUpper = sPNameUpper Then
				txtInvite2.Locked = False
				btnInvite2.Caption = "Invite"
				chkAccept2.Value = False
			End If
		Else
			If txtInvite1.Caption.Trim.ToUpper = sPNameUpper Then
				chkAccept1.Value = False
				mlInviteIDs(0) = lPlayerID
			End If
			If txtInvite2.Caption.Trim.ToUpper = sPNameUpper Then
				chkAccept2.Value = False
				mlInviteIDs(1) = lPlayerID
			End If
			If txtInvite3.Caption.Trim.ToUpper = sPNameUpper Then
				chkAccept3.Value = False
				mlInviteIDs(2) = lPlayerID
			End If
			If txtInvite4.Caption.Trim.ToUpper = sPNameUpper Then
				chkAccept4.Value = False
				mlInviteIDs(3) = lPlayerID
			End If

			If moGuild.moMembers Is Nothing Then ReDim moGuild.moMembers(-1)
			Dim bFound As Boolean = False
			For X As Int32 = 0 To moGuild.moMembers.GetUpperBound(0)
				If moGuild.moMembers(X) Is Nothing = False AndAlso moGuild.moMembers(X).lMemberID = lPlayerID Then
					bFound = True
					moGuild.moMembers(X).yMemberState = GuildMemberState.Invited Or GuildMemberState.AcceptedGuildFormInvite
					Exit For
				End If
			Next X
			If bFound = False Then
				ReDim Preserve moGuild.moMembers(moGuild.moMembers.GetUpperBound(0) + 1)
				moGuild.moMembers(moGuild.moMembers.GetUpperBound(0)) = New GuildMember
				With moGuild.moMembers(moGuild.moMembers.GetUpperBound(0))
					.lMemberID = lPlayerID
					Dim oRank As GuildRank = moGuild.GetLowestRank()
					If oRank Is Nothing = False Then .lRankID = oRank.lRankID Else .lRankID = -1
					.yMemberState = GuildMemberState.Invited Or GuildMemberState.AcceptedGuildFormInvite
				End With
			End If

			SendGuildUpdate()

			FillRankNodes()
		End If

	End Sub

	Public Sub SetNewGuildObject(ByRef oGuild As Guild, ByVal lFormer As Int32)
		moGuild = oGuild
		mlGuildFormer = lFormer
		fraIconChooser.SetStartingIcon(oGuild.lIcon)
		txtGuildName.Caption = oGuild.sName
		FillRankNodes()
		chkAccept1.Value = False
		chkAccept2.Value = False
		chkAccept3.Value = False
		chkAccept4.Value = False

		If mlGuildFormer <> glPlayerID Then
			txtGuildName.Locked = True
			fraIconChooser.Enabled = False
			txtInvite1.Locked = True
			txtInvite2.Locked = True
			txtInvite3.Locked = True
			txtInvite4.Locked = True
			txtRankName.Locked = True
			txtVoteRate.Locked = True
			txtFlatRate.Locked = True
			txtTaxPerc.Locked = True
			cboTaxPercType.Enabled = False
			btnAdd.Enabled = False
			btnDelete.Enabled = False
			btnUp.Enabled = False
			btnDown.Enabled = False
			btnUpdate.Enabled = False
			chkRequirePeace.Locked = True
			chkAutoTrade.Locked = True
			chkShareVision.Locked = True
            'chkGuildRels.Locked = True
			cboVoteWeight.Enabled = False
			cboTaxInterval.Enabled = False
			txtTaxDay.Locked = True
			txtTaxMonth.Locked = True
			chkDemoteRA.Locked = True
			chkAcceptRA.Locked = True
			chkPromoteRA.Locked = True
			chkRemoveRA.Locked = True
			chkCreateRankRA.Locked = True
			chkDeleteRankRA.Locked = True
			chkChangeVoteRA.Locked = True
			btnUpdateRules.Enabled = False

			If chkRule Is Nothing = False Then
				For X As Int32 = 0 To chkRule.GetUpperBound(0)
					If chkRule(X) Is Nothing = False Then
						chkRule(X).Locked = True
					End If
				Next X
			End If

			txtInvite1.Caption = "" : txtInvite2.Caption = "" : txtInvite3.Caption = "" : txtInvite4.Caption = ""
			btnInvite1.Visible = False : btnInvite2.Visible = False : btnInvite3.Visible = False : btnInvite4.Visible = False
			chkAccept1.Locked = True : chkAccept2.Locked = True : chkAccept3.Locked = True : chkAccept4.Locked = True
			chkAccept1.Value = False : chkAccept2.Value = False : chkAccept3.Value = False : chkAccept4.Value = False

			If moGuild.moMembers Is Nothing = False Then
				Dim lItemIdx As Int32 = 0
				For X As Int32 = 0 To moGuild.moMembers.GetUpperBound(0)
					If moGuild.moMembers(X) Is Nothing = False AndAlso moGuild.moMembers(X).lMemberID <> mlGuildFormer Then
						lItemIdx += 1
						Dim sTxt As String = GetCacheObjectValue(moGuild.moMembers(X).lMemberID, ObjectType.ePlayer)
                        Dim bChk As Boolean = False '(moGuild.moMembers(X).yMemberState And GuildMemberState.Approved) <> 0
						Dim bBtn As Boolean = moGuild.moMembers(X).lMemberID <> glPlayerID
						Select Case lItemIdx
							Case 1
								txtInvite1.Caption = sTxt : chkAccept1.Value = bChk : chkAccept1.Locked = bBtn
								mlInviteIDs(0) = moGuild.moMembers(X).lMemberID
							Case 2
								txtInvite2.Caption = sTxt : chkAccept2.Value = bChk : chkAccept2.Locked = bBtn
								mlInviteIDs(1) = moGuild.moMembers(X).lMemberID
							Case 3
								txtInvite3.Caption = sTxt : chkAccept3.Value = bChk : chkAccept3.Locked = bBtn
								mlInviteIDs(2) = moGuild.moMembers(X).lMemberID
							Case 4
								txtInvite4.Caption = sTxt : chkAccept4.Value = bChk : chkAccept4.Locked = bBtn
								mlInviteIDs(3) = moGuild.moMembers(X).lMemberID
							Case Else
								Exit For
						End Select
					End If
				Next X
			End If

			
		End If
	End Sub

	Private Sub tvwRanks_NodeSelected(ByRef oNode As UITreeView.UITreeViewItem) Handles tvwRanks.NodeSelected
		'selected a node
		If oNode Is Nothing = False Then
			Dim lID As Int32 = tvwRanks.oSelectedNode.lItemData
			Dim lTypeID As Int32 = tvwRanks.oSelectedNode.lItemData2

			If lTypeID = -1 Then
				'guild rank
				If mlPermVals Is Nothing Then FillPermVals()
				Dim oRank As GuildRank = moGuild.GetRankByID(lID)
				If oRank Is Nothing Then Return
				FillRankRuleList(oRank)
				txtRankName.Caption = oRank.sRankName
				txtVoteRate.Caption = oRank.lVoteStrength.ToString
				txtFlatRate.Caption = oRank.TaxRateFlat.ToString
				txtTaxPerc.Caption = oRank.TaxRatePercentage.ToString
				cboTaxPercType.FindComboItemData(oRank.TaxRatePercType)
			Else
				'player
				If mlPermVals Is Nothing Then FillPermVals()
				For X As Int32 = 0 To chkRule.GetUpperBound(0)
					chkRule(X).Visible = False
				Next X
			End If
		End If
	End Sub

	Private Sub btnAdd_Click(ByVal sName As String) Handles btnAdd.Click
		'add rank
		If moGuild.moRanks Is Nothing Then ReDim moGuild.moRanks(-1)
		ReDim Preserve moGuild.moRanks(moGuild.moRanks.GetUpperBound(0) + 1)
		moGuild.moRanks(moGuild.moRanks.GetUpperBound(0)) = New GuildRank
		Dim lRankID As Int32 = moGuild.moRanks.GetUpperBound(0) + 2
		With moGuild.moRanks(moGuild.moRanks.GetUpperBound(0))
			.sRankName = "New Rank"
			.yPosition = CByte(moGuild.moRanks.GetUpperBound(0) + 2)
			.lRankID = -lRankID
		End With
		tvwRanks.AddNode("New Rank", -lRankID, -1, -1, Nothing, Nothing)
		SendGuildUpdate()
	End Sub

	Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
		'delete rank
		If tvwRanks.oSelectedNode Is Nothing = False Then
			Dim lID As Int32 = tvwRanks.oSelectedNode.lItemData
			Dim lTypeID As Int32 = tvwRanks.oSelectedNode.lItemData2
			If lTypeID = -1 Then
				'rank
				For X As Int32 = 0 To moGuild.moRanks.GetUpperBound(0)
					If moGuild.moRanks(X).lRankID = lID Then
						moGuild.moRanks(X) = Nothing
						SendGuildUpdate()
						Exit For
					End If
				Next X
				FillRankNodes()
			End If
		End If
	End Sub

	Private Sub btnDown_Click(ByVal sName As String) Handles btnDown.Click
		'move rank down
		If tvwRanks.oSelectedNode Is Nothing = False Then
			Dim lID As Int32 = tvwRanks.oSelectedNode.lItemData
			Dim lTypeID As Int32 = tvwRanks.oSelectedNode.lItemData2

			If lTypeID = -1 Then
				'guild rank - move the rank up
				moGuild.MoveRank(lID, 2)
				SendGuildUpdate()
			Else
				'player - promote the player
				Dim oMember As GuildMember = moGuild.GetMember(lID)
				If oMember Is Nothing = False Then
					Dim oNextRank As GuildRank = moGuild.GetNextRankPosition(1, oMember.lRankID)
					If oNextRank Is Nothing = False Then
						oMember.lRankID = oNextRank.lRankID
						SendGuildUpdate()
					End If
				End If
			End If
			FillRankNodes()
		End If
	End Sub

	Private Sub btnUp_Click(ByVal sName As String) Handles btnUp.Click
		'move rank up
		If tvwRanks.oSelectedNode Is Nothing = False Then
			Dim lID As Int32 = tvwRanks.oSelectedNode.lItemData
			Dim lTypeID As Int32 = tvwRanks.oSelectedNode.lItemData2

			If lTypeID = -1 Then
				'guild rank - move the rank up
				moGuild.MoveRank(lID, 1)
				SendGuildUpdate()
			Else
				'player - promote the player
				Dim oMember As GuildMember = moGuild.GetMember(lID)
				If oMember Is Nothing = False Then
					Dim oNextRank As GuildRank = moGuild.GetNextRankPosition(-1, oMember.lRankID)
					If oNextRank Is Nothing = False Then
						oMember.lRankID = oNextRank.lRankID
						SendGuildUpdate()
					End If
				End If
			End If
			FillRankNodes()
		End If
	End Sub

	Private Sub btnUpdate_Click(ByVal sName As String) Handles btnUpdate.Click
		'update ranks
		If tvwRanks.oSelectedNode Is Nothing = False Then
			If tvwRanks.oSelectedNode.lItemData2 = -1 Then
				'rank... so we're good
				Dim oRank As GuildRank = moGuild.GetRankByID(tvwRanks.oSelectedNode.lItemData)
				If oRank Is Nothing Then Return

				If txtFlatRate.Caption.Contains(".") = True Then
					MyBase.moUILib.AddNotification("Tax Flat Rate must be a whole number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				End If
				If IsNumeric(txtFlatRate.Caption) = False Then
					MyBase.moUILib.AddNotification("Tax Flat Rate must be a numeric value!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				End If
				Dim blTest As Int64 = CLng(txtFlatRate.Caption)
				If blTest < 0 Then
					MyBase.moUILib.AddNotification("Tax Flat Rate must be a non-negative number!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				End If
				If blTest > Int32.MaxValue Then
					MyBase.moUILib.AddNotification("Tax Flat Rate is an invalid value.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				End If
				Dim lTaxFlat As Int32 = CInt(blTest)
				If txtTaxPerc.Caption = "" Then txtTaxPerc.Caption = "0"
				If IsNumeric(txtTaxPerc.Caption) = False OrElse txtTaxPerc.Caption.Contains("-") = True OrElse txtTaxPerc.Caption.Contains(".") = True OrElse CLng(txtTaxPerc.Caption) > 100 Then
					MyBase.moUILib.AddNotification("Tax Percentage must be a numeric whole number value between 0 and 100!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				End If
				Dim yTaxPerc As Byte = CByte(txtTaxPerc.Caption)

				If txtVoteRate.Caption.Contains(".") = True Then
					MyBase.moUILib.AddNotification("Vote Rate must be a whole number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				End If
				If IsNumeric(txtVoteRate.Caption) = False Then
					MyBase.moUILib.AddNotification("Vote Rate must be a numeric value!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				End If
				blTest = CLng(txtVoteRate.Caption)
				If blTest < 0 Then
					MyBase.moUILib.AddNotification("Vote Rate must be a non-negative number!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				End If
				If blTest > Int32.MaxValue Then
					MyBase.moUILib.AddNotification("Vote Rate is an invalid value.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				End If
				Dim lVoteRate As Int32 = CInt(blTest)

				If txtRankName.Caption.Trim.Length < 3 Then
					MyBase.moUILib.AddNotification("Rank Name must be at least 3 characters in length.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				End If

				oRank.sRankName = txtRankName.Caption.Trim
				tvwRanks.oSelectedNode.sItem = oRank.sRankName
				oRank.lVoteStrength = lVoteRate
				oRank.TaxRateFlat = lTaxFlat
				oRank.TaxRatePercentage = yTaxPerc
				oRank.TaxRatePercType = eyGuildTaxPercType.CashFlow
				If cboTaxPercType.ListIndex > -1 Then oRank.TaxRatePercType = CType(cboTaxPercType.ItemData(cboTaxPercType.ListIndex), eyGuildTaxPercType)

				SendGuildUpdate()
			End If
		End If
	End Sub

	Private Sub btnUpdateRules_Click(ByVal sName As String) Handles btnUpdateRules.Click
		'update rules
		SendGuildUpdate()
	End Sub

	Private Sub SendGuildUpdate()
		If mlGuildFormer <> glPlayerID Then Return
		If moGuild Is Nothing = False Then
			With moGuild
				'Properties not used at this stage:
				'	blTreasury, dtFormed, dtJoined, dtLastGuildRuleChange, iGuildHallLocTypeID, iRecruitFlags
				'	lGuildHallID, lGuildHallLocID, lGuildHallLocX, lGuildHallLocZ, lCurrentRankID, moEvents
				'	moRels, moVotes, sBillboard, sMOTD, ObjectID, ObjTypeID

				Dim lMemberCnt As Int32 = 0
				If .moMembers Is Nothing = False Then
					For X As Int32 = 0 To .moMembers.GetUpperBound(0)
						If .moMembers(X) Is Nothing = False Then lMemberCnt += 1
					Next X
				End If
				Dim lRankCnt As Int32 = 0
				If .moRanks Is Nothing = False Then
					For X As Int32 = 0 To .moRanks.GetUpperBound(0)
						If .moRanks(X) Is Nothing = False Then lRankCnt += 1
					Next X
				End If

				'So, that leaves us with:
				Dim yMsg(70 + (lRankCnt * 39) + (lMemberCnt * 9)) As Byte
				Dim lPos As Int32 = 0

				System.BitConverter.GetBytes(GlobalMessageCode.eGuildCreationUpdate).CopyTo(yMsg, lPos) : lPos += 2

				System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, lPos) : lPos += 4

				'lBaseGuildRules
				.lBaseGuildRules = 0
                If chkAcceptRA.Value = False Then .lBaseGuildRules = .lBaseGuildRules Or elGuildFlags.AcceptMemberToGuild_RA
                If chkAutoTrade.Value = True Then .lBaseGuildRules = .lBaseGuildRules Or elGuildFlags.AutomaticTradeAgreements
                If chkChangeVoteRA.Value = False Then .lBaseGuildRules = .lBaseGuildRules Or elGuildFlags.ChangeVotingWeight_RA
                If chkCreateRankRA.Value = False Then .lBaseGuildRules = .lBaseGuildRules Or elGuildFlags.CreateRank_RA
                If chkDeleteRankRA.Value = False Then .lBaseGuildRules = .lBaseGuildRules Or elGuildFlags.DeleteRank_RA
                If chkDemoteRA.Value = False Then .lBaseGuildRules = .lBaseGuildRules Or elGuildFlags.DemoteMember_RA
                'If chkGuildRels.Value = True Then .lBaseGuildRules = .lBaseGuildRules Or elGuildFlags.UnifiedForeignPolicy
                If chkPromoteRA.Value = False Then .lBaseGuildRules = .lBaseGuildRules Or elGuildFlags.PromoteGuildMember_RA
                If chkRemoveRA.Value = False Then .lBaseGuildRules = .lBaseGuildRules Or elGuildFlags.RemoveGuildMember_RA
                If chkRequirePeace.Value = True Then .lBaseGuildRules = .lBaseGuildRules Or elGuildFlags.RequirePeaceBetweenMembers
                If chkShareVision.Value = True Then .lBaseGuildRules = .lBaseGuildRules Or elGuildFlags.ShareUnitVision
				System.BitConverter.GetBytes(.lBaseGuildRules).CopyTo(yMsg, lPos) : lPos += 4

				'lIcon
				.lIcon = fraIconChooser.GetStartingIcon()
				System.BitConverter.GetBytes(.lIcon).CopyTo(yMsg, lPos) : lPos += 4

				'sName
				.sName = txtGuildName.Caption.Trim
				If .sName.Length > 50 Then .sName = Mid$(.sName, 1, 50)
				System.Text.ASCIIEncoding.ASCII.GetBytes(.sName).CopyTo(yMsg, lPos) : lPos += 50

				Dim lTemp As Int32 = CInt(Val(txtTaxMonth.Caption.Trim))
				If lTemp < 1 Then lTemp = 1
				If lTemp > 12 Then lTemp = 12
				.yGuildTaxBaseMonth = CByte(lTemp)
				If .yGuildTaxBaseMonth.ToString <> txtTaxMonth.Caption Then txtTaxMonth.Caption = .yGuildTaxBaseMonth.ToString
				yMsg(lPos) = .yGuildTaxBaseMonth : lPos += 1

				lTemp = CInt(Val(txtTaxDay.Caption))
				If lTemp < 1 Then lTemp = 1
				Dim lMaxDay As Int32 = Date.DaysInMonth(Now.Year, .yGuildTaxBaseMonth)
				If lTemp > lMaxDay Then lTemp = lMaxDay
				.yGuildTaxBaseDay = CByte(lTemp)
				If .yGuildTaxBaseDay.ToString <> txtTaxDay.Caption Then txtTaxDay.Caption = .yGuildTaxBaseDay.ToString
				yMsg(lPos) = .yGuildTaxBaseDay : lPos += 1

				If cboTaxInterval.ListIndex > -1 Then
					.yGuildTaxRateInterval = CType(cboTaxInterval.ItemData(cboTaxInterval.ListIndex), eyGuildInterval)
				Else
					.yGuildTaxRateInterval = eyGuildInterval.Daily
					cboTaxInterval.FindComboItemData(.yGuildTaxRateInterval)
				End If
				yMsg(lPos) = .yGuildTaxRateInterval : lPos += 1

				'yState
				.yState = eyGuildState.Forming
				yMsg(lPos) = .yState : lPos += 1

				'yVoteWeightType
				If cboVoteWeight.ListIndex > -1 Then
					.yVoteWeightType = CType(cboVoteWeight.ItemData(cboVoteWeight.ListIndex), eyVoteWeightType)
				Else
					.yVoteWeightType = eyVoteWeightType.RankBased
					cboVoteWeight.FindComboItemData(.yVoteWeightType)
				End If
				yMsg(lPos) = .yVoteWeightType : lPos += 1

				yMsg(lPos) = CByte(lMemberCnt) : lPos += 1
				If .moMembers Is Nothing = False Then
					Dim lMembersUsed As Int32 = 0
					For X As Int32 = 0 To .moMembers.GetUpperBound(0)
						If .moMembers(X) Is Nothing = False Then
							lMembersUsed += 1
							System.BitConverter.GetBytes(.moMembers(X).lMemberID).CopyTo(yMsg, lPos) : lPos += 4
							System.BitConverter.GetBytes(.moMembers(X).lRankID).CopyTo(yMsg, lPos) : lPos += 4
							yMsg(lPos) = .moMembers(X).yMemberState : lPos += 1
							If lMembersUsed = lMemberCnt Then Exit For
						End If
					Next X
				End If

				yMsg(lPos) = CByte(lRankCnt) : lPos += 1
				If .moRanks Is Nothing = False Then
					Dim lRanksUsed As Int32 = 0
					For X As Int32 = 0 To .moRanks.GetUpperBound(0)
						If .moRanks(X) Is Nothing = False Then
							lRanksUsed += 1
							System.BitConverter.GetBytes(.moRanks(X).lRankID).CopyTo(yMsg, lPos) : lPos += 4
							System.BitConverter.GetBytes(.moRanks(X).lRankPermissions).CopyTo(yMsg, lPos) : lPos += 4
							System.BitConverter.GetBytes(.moRanks(X).lVoteStrength).CopyTo(yMsg, lPos) : lPos += 4
							System.Text.ASCIIEncoding.ASCII.GetBytes(.moRanks(X).sRankName).CopyTo(yMsg, lPos) : lPos += 20
							System.BitConverter.GetBytes(.moRanks(X).TaxRateFlat).CopyTo(yMsg, lPos) : lPos += 4
							yMsg(lPos) = .moRanks(X).TaxRatePercentage : lPos += 1
							yMsg(lPos) = .moRanks(X).TaxRatePercType : lPos += 1
							yMsg(lPos) = .moRanks(X).yPosition : lPos += 1
							If lRanksUsed = lRankCnt Then Exit For
						End If
					Next X
				End If

				MyBase.moUILib.SendMsgToPrimary(yMsg)
			End With
		End If

		chkAccept1.Value = False
		chkAccept2.Value = False
		chkAccept3.Value = False
		chkAccept4.Value = False
	End Sub

	Private Sub chkRuleClick()
		If glPlayerID <> mlGuildFormer Then Return
		If tvwRanks.oSelectedNode Is Nothing = False Then
			Dim lID As Int32 = tvwRanks.oSelectedNode.lItemData
			Dim lTypeID As Int32 = tvwRanks.oSelectedNode.lItemData2

			If lTypeID = -1 Then
				'guild rank
				If mlPermVals Is Nothing Then FillPermVals()
				Dim oRank As GuildRank = moGuild.GetRankByID(lID)
				If oRank Is Nothing Then Return
				Dim lRuleIdx As Int32 = vscrPerms.Value
				For X As Int32 = 0 To chkRule.GetUpperBound(0)
					If chkRule(X).Value = True Then
						If (oRank.lRankPermissions And mlPermVals(lRuleIdx + X)) = 0 Then
							oRank.lRankPermissions = CType(oRank.lRankPermissions Or mlPermVals(lRuleIdx + X), RankPermissions)
						End If
					Else
						If (oRank.lRankPermissions And mlPermVals(lRuleIdx + X)) <> 0 Then
							oRank.lRankPermissions = CType(oRank.lRankPermissions Xor mlPermVals(lRuleIdx + X), RankPermissions)
						End If
					End If
				Next X
				SendGuildUpdate()
			End If
		End If
	End Sub

	Private Sub chkAccept_Click() Handles chkAccept1.Click, chkAccept2.Click, chkAccept3.Click, chkAccept4.Click
		If glPlayerID <> mlGuildFormer Then
			
			Dim bValue As Boolean = True
            'If chkAccept1.Enabled = True Then bValue = chkAccept1.Value
            'If chkAccept2.Enabled = True Then bValue = chkAccept2.Value
            'If chkAccept3.Enabled = True Then bValue = chkAccept3.Value
            'If chkAccept4.Enabled = True Then bValue = chkAccept4.Value
            For X As Int32 = 0 To mlInviteIDs.GetUpperBound(0)
                If mlInviteIDs(X) = glPlayerID Then
                    Select Case X
                        Case 0
                            bValue = chkAccept1.Value
                        Case 1
                            bValue = chkAccept2.Value
                        Case 2
                            bValue = chkAccept3.Value
                        Case 3
                            bValue = chkAccept4.Value
                    End Select
                    Exit For
                End If
            Next X
			Dim yAcceptValue As Byte = 1
			If bValue = True Then yAcceptValue = 1 Else yAcceptValue = 0
			SendGuildCreationAcceptance(yAcceptValue)
		End If
	End Sub

	Private Sub SendGuildCreationAcceptance(ByVal yAcceptValue As Byte)
		Dim yMsg(10) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eGuildCreationAcceptance).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(mlGuildFormer).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, lPos) : lPos += 4
		yMsg(lPos) = yAcceptValue : lPos += 1
		MyBase.moUILib.SendMsgToPrimary(yMsg)
	End Sub

	Public Sub HandleGuildCreationAcceptance(ByRef yData() As Byte)
		Dim lPos As Int32 = 2 'for msgcode
		Dim lFormer As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		Dim yValue As Byte = yData(lPos) : lPos += 1

		If yValue = 255 Then
			'someone is leaving, is it the former?
			If lFormer = lPlayerID Then
				'yes
				MyBase.moUILib.AddNotification("Guild Creator left the creation session.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				MyBase.moUILib.RemoveWindow(Me.ControlName)
			Else
				If moGuild Is Nothing = False AndAlso moGuild.moMembers Is Nothing = False Then
					For X As Int32 = 0 To moGuild.moMembers.GetUpperBound(0)
						If moGuild.moMembers(X) Is Nothing = False AndAlso moGuild.moMembers(X).lMemberID = lPlayerID Then
							moGuild.moMembers(X) = Nothing
							Exit For
						End If
					Next X
				End If

				For X As Int32 = 0 To 3
					If mlInviteIDs(X) = lPlayerID Then
						Select Case X
							Case 0
								txtInvite1.Caption = "" : txtInvite1.Locked = False : btnInvite1.Caption = "Invite" : chkAccept1.Value = False
							Case 1
								txtInvite2.Caption = "" : txtInvite2.Locked = False : btnInvite2.Caption = "Invite" : chkAccept2.Value = False
							Case 2
								txtInvite3.Caption = "" : txtInvite3.Locked = False : btnInvite3.Caption = "Invite" : chkAccept3.Value = False
							Case 3
								txtInvite4.Caption = "" : txtInvite4.Locked = False : btnInvite4.Caption = "Invite" : chkAccept4.Value = False
						End Select
					End If
				Next X

				FillRankNodes()
			End If
		Else
            'ok, update someone's status
            Dim bResult As Boolean = False
			If moGuild Is Nothing = False AndAlso moGuild.moMembers Is Nothing = False Then
				For X As Int32 = 0 To moGuild.moMembers.GetUpperBound(0)
					If moGuild.moMembers(X) Is Nothing = False AndAlso moGuild.moMembers(X).lMemberID = lPlayerID Then
						If yValue = 0 Then
							If (moGuild.moMembers(X).yMemberState And GuildMemberState.Approved) <> 0 Then
								moGuild.moMembers(X).yMemberState = moGuild.moMembers(X).yMemberState Xor GuildMemberState.Approved
							End If
						Else
                            moGuild.moMembers(X).yMemberState = moGuild.moMembers(X).yMemberState Or GuildMemberState.Approved
                            bResult = True
						End If
						Exit For
					End If
				Next X
			End If

			For X As Int32 = 0 To 3
				If mlInviteIDs(X) = lPlayerID Then
					Select Case X
						Case 0
                            chkAccept1.Value = bResult
						Case 1
                            chkAccept2.Value = bResult
						Case 2
                            chkAccept3.Value = bResult
						Case 3
                            chkAccept4.Value = bResult
					End Select
				End If
			Next X

		End If
	End Sub

	Private Sub vscrPerms_ValueChange() Handles vscrPerms.ValueChange
		If tvwRanks.oSelectedNode Is Nothing = False Then
			Dim lID As Int32 = tvwRanks.oSelectedNode.lItemData
			Dim lTypeID As Int32 = tvwRanks.oSelectedNode.lItemData2
			If mlPermVals Is Nothing Then FillPermVals()
			Dim oRank As GuildRank = moGuild.GetRankByID(lID)
			If oRank Is Nothing Then Return
			FillRankRuleList(oRank)
		End If
	End Sub

	Private Sub FillRankRuleList(ByRef oRank As GuildRank)
		Dim lRuleIdx As Int32 = vscrPerms.Value
		For X As Int32 = 0 To chkRule.GetUpperBound(0)
			chkRule(X).Caption = msPermVals(lRuleIdx + X)
			chkRule(X).Visible = True
			chkRule(X).Value = (oRank.lRankPermissions And mlPermVals(lRuleIdx + X)) <> 0
			Dim rcTemp As Rectangle = chkRule(X).GetTextDimensions(chkRule(X).Caption)
			chkRule(X).Width = rcTemp.Width + 15
		Next X
	End Sub
End Class
