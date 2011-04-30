Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmGuildMain
	Private Class fra_Membership
		Inherits guildframe

		Private fraMemberList As UIWindow
		Private lblInvite As UILabel
		Private txtInvite As UITextBox
		Private lblDetails As UILabel
		Private WithEvents chkDiplomacy As UICheckBox
		Private WithEvents chkEspionage As UICheckBox
		Private WithEvents chkMilitary As UICheckBox
		Private WithEvents chkProduction As UICheckBox
		Private txtDetails As UITextBox
		Private fraRecruitment As UIWindow
		Private WithEvents chkWillTrain As UICheckBox
		Private WithEvents chkTrade As UICheckBox
		Private WithEvents chkResearch As UICheckBox
		Private WithEvents chkRecruitTutorial As UICheckBox
		Private WithEvents chkRecruiting As UICheckBox
		Private lblLFP As UILabel
		Private lblBillboard As UILabel
		Private WithEvents lstMembers As UIListBox

		Private WithEvents btnAccept As UIButton
		Private WithEvents btnInvite As UIButton
		Private WithEvents btnInviteBrowse As UIButton
		Private WithEvents btnRemove As UIButton
		Private WithEvents btnDemote As UIButton
		Private WithEvents btnPromote As UIButton
        Private WithEvents btnUpdate As UIButton
        Private WithEvents btnBillboard As UIButton
		Private WithEvents txtBillboard As UITextBox

		Private mbEditBillboard As Boolean = False

		Private mbEditDiplomacy As Boolean = False
		Private mbEditEspionage As Boolean = False
		Private mbEditMilitary As Boolean = False
		Private mbEditProduction As Boolean = False
		Private mbEditTrade As Boolean = False
		Private mbEditResearch As Boolean = False
		Private mbEditRecruitTutorial As Boolean = False
		Private mbEditRecruiting As Boolean = False
		Private mbEditWillTrain As Boolean = False

		Private myMemberSort As Byte = 0		'0 member name asc, 1 member name desc, 2 rank asc, 2 rank desc

		Public Sub New(ByRef oUILib As UILib)
			MyBase.New(oUILib)

            With Me
                .lWindowMetricID = BPMetrics.eWindow.eGuildMain_Membership
                .ControlName = "fra_Membership"
                .Left = 15
                .Top = ml_CONTENTS_TOP
                .Width = 128
                .Height = 128
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 2
                .Moveable = False
            End With

			'fraMemberList initial props
			fraMemberList = New UIWindow(MyBase.moUILib)
			With fraMemberList
				.ControlName = "fraMemberList"
				.Left = 15
				.Top = 5
				.Width = 510
				.Height = 515
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.Moveable = False
				.BorderLineWidth = 2
				.Caption = "Memberlist"
			End With
			Me.AddChild(CType(fraMemberList, UIControl))

			lblInvite = New UILabel(MyBase.moUILib)
			With lblInvite
				.ControlName = "lblInvite"
				.Left = 10
				.Top = 10
				.Width = 120
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Invite a Player:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
			End With
			fraMemberList.AddChild(CType(lblInvite, UIControl))

			'btnInvite initial props
			btnInvite = New UIButton(MyBase.moUILib)
			With btnInvite
				.ControlName = "btnInvite"
				.Left = fraMemberList.Width - 110
				.Top = lblInvite.Top
				.Width = 100
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.InviteMember)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to invite players to join."
				End If
				.Visible = True
				.Caption = "Invite"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraMemberList.AddChild(CType(btnInvite, UIControl))

			'btnInviteBrowse initial props
			btnInviteBrowse = New UIButton(MyBase.moUILib)
			With btnInviteBrowse
				.ControlName = "btnInviteBrowse"
				.Left = btnInvite.Left - 38
				.Top = lblInvite.Top
				.Width = 32
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.InviteMember)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to invite players to join."
				End If
				.Visible = True
				.Caption = "..."
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraMemberList.AddChild(CType(btnInviteBrowse, UIControl))

			'txtDetails initial props
			txtInvite = New UITextBox(MyBase.moUILib)
			With txtInvite
				.ControlName = "txtInvite"
				.Left = lblInvite.Left + lblInvite.Width + 5
				.Top = lblInvite.Top + 1
				.Width = btnInviteBrowse.Left - .Left - 5
				.Height = 22 '120
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.InviteMember)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to invite players to join."
				End If
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
			End With
			fraMemberList.AddChild(CType(txtInvite, UIControl))

			'lstMembers initial props
			lstMembers = New UIListBox(MyBase.moUILib)
			With lstMembers
				.ControlName = "lstMembers"
				.Left = 10
				.Top = 40			'15
				.Width = 490
				.Height = 310
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Courier New", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
			End With
			fraMemberList.AddChild(CType(lstMembers, UIControl))

			'btnRemove initial props
			btnRemove = New UIButton(MyBase.moUILib)
			With btnRemove
				.ControlName = "btnRemove"
				.Left = 10
				.Top = 485
				.Width = 100
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.RemoveMember)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to remove members."
				End If
				.Visible = True
				.Caption = "Remove"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraMemberList.AddChild(CType(btnRemove, UIControl))

			'btnAccept initial props
			btnAccept = New UIButton(MyBase.moUILib)
			With btnAccept
				.ControlName = "btnAccept"
				.Left = 400
				.Top = btnRemove.Top
				.Width = 100
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.AcceptApplicant)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to accept applications."
				End If
				.Visible = True
				.Caption = "Accept"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraMemberList.AddChild(CType(btnAccept, UIControl))

			'btnPromote initial props
			btnPromote = New UIButton(MyBase.moUILib)
			With btnPromote
				.ControlName = "btnPromote"
				.Left = 270
				.Top = btnRemove.Top
				.Width = 100
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.PromoteMember)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to promote members."
				End If
				.Visible = True
				.Caption = "Promote"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraMemberList.AddChild(CType(btnPromote, UIControl))

			'btnDemote initial props
			btnDemote = New UIButton(MyBase.moUILib)
			With btnDemote
				.ControlName = "btnDemote"
				.Left = 140
				.Top = btnRemove.Top
				.Width = 100
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.DemoteMember)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to demote members."
				End If
				.Visible = True
				.Caption = "Demote"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraMemberList.AddChild(CType(btnDemote, UIControl))

			'lblDetails initial props
			lblDetails = New UILabel(MyBase.moUILib)
			With lblDetails
				.ControlName = "lblDetails"
				.Left = 10
				.Top = 350
				.Width = 100
				.Height = 32
				.Enabled = True
				.Visible = True
				.Caption = "Details"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraMemberList.AddChild(CType(lblDetails, UIControl))

			'txtDetails initial props
			txtDetails = New UITextBox(MyBase.moUILib)
			With txtDetails
				.ControlName = "txtDetails"
				.Left = 10
				.Top = 380
				.Width = 490
				.Height = 95 '120
				.Enabled = True
				.Visible = True
				.Caption = "Text Caption"
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
			fraMemberList.AddChild(CType(txtDetails, UIControl))

			'=== Recruitment
			'fraRecruitment initial props
			fraRecruitment = New UIWindow(MyBase.moUILib)
			With fraRecruitment
				.ControlName = "fraRecruitment"
				.Left = 530
				.Top = 5
				.Width = 255
				.Height = 515
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.BorderLineWidth = 2
				.Moveable = False
				.Caption = "Recruitment"
			End With
			Me.AddChild(CType(fraRecruitment, UIControl))

			'chkRecruiting initial props
			chkRecruiting = New UICheckBox(MyBase.moUILib)
			With chkRecruiting
				.ControlName = "chkRecruiting"
				.Left = 10
				.Top = 10
				.Width = 151
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRecruitment)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to change the recruitment settings."
				End If
				.Visible = True
				.Caption = "Actively Recruiting"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraRecruitment.AddChild(CType(chkRecruiting, UIControl))

			'chkRecruitTutorial initial props
			chkRecruitTutorial = New UICheckBox(MyBase.moUILib)
			With chkRecruitTutorial
				.ControlName = "chkRecruitTutorial"
				.Left = 10
				.Top = 30
				.Width = 216
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRecruitment)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to change the recruitment settings."
				End If
				.Visible = True
				.Caption = "Recruit from Tutorial System"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraRecruitment.AddChild(CType(chkRecruitTutorial, UIControl))

			'chkWillTrain initial props
			chkWillTrain = New UICheckBox(MyBase.moUILib)
			With chkWillTrain
				.ControlName = "chkWillTrain"
				.Left = 10
				.Top = 50
				.Width = 216
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRecruitment)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to change the recruitment settings."
				End If
				.Visible = True
				.Caption = "Willing to Train New Players"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraRecruitment.AddChild(CType(chkWillTrain, UIControl))

			'lblLFP initial props
			lblLFP = New UILabel(MyBase.moUILib)
			With lblLFP
				.ControlName = "lblLFP"
				.Left = 10
				.Top = 75
				.Width = 232
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Looking for Players Who Specialize In:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraRecruitment.AddChild(CType(lblLFP, UIControl))

			'chkProduction initial props
			chkProduction = New UICheckBox(MyBase.moUILib)
			With chkProduction
				.ControlName = "chkProduction"
				.Left = 30
				.Top = 100
				.Width = 86
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRecruitment)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to change the recruitment settings."
				End If
				.Visible = True
				.Caption = "Production"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraRecruitment.AddChild(CType(chkProduction, UIControl))

			'chkResearch initial props
			chkResearch = New UICheckBox(MyBase.moUILib)
			With chkResearch
				.ControlName = "chkResearch"
				.Left = 30
				.Top = 120
				.Width = 80
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRecruitment)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to change the recruitment settings."
				End If
				.Visible = True
				.Caption = "Research"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraRecruitment.AddChild(CType(chkResearch, UIControl))

			'chkDiplomacy initial props
			chkDiplomacy = New UICheckBox(MyBase.moUILib)
			With chkDiplomacy
				.ControlName = "chkDiplomacy"
				.Left = 30
				.Top = 140
				.Width = 86
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRecruitment)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to change the recruitment settings."
				End If
				.Visible = True
				.Caption = "Diplomacy"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraRecruitment.AddChild(CType(chkDiplomacy, UIControl))

			'chkMilitary initial props
			chkMilitary = New UICheckBox(MyBase.moUILib)
			With chkMilitary
				.ControlName = "chkMilitary"
				.Left = 30
				.Top = 160
				.Width = 114
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRecruitment)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to change the recruitment settings."
				End If
				.Visible = True
				.Caption = "Military/Combat"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraRecruitment.AddChild(CType(chkMilitary, UIControl))

			'chkTrade initial props
			chkTrade = New UICheckBox(MyBase.moUILib)
			With chkTrade
				.ControlName = "chkTrade"
				.Left = 30
				.Top = 180
				.Width = 58
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRecruitment)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to change the recruitment settings."
				End If
				.Visible = True
				.Caption = "Trade"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraRecruitment.AddChild(CType(chkTrade, UIControl))

			'chkEspionage initial props
			chkEspionage = New UICheckBox(MyBase.moUILib)
			With chkEspionage
				.ControlName = "chkEspionage"
				.Left = 30
				.Top = 200
				.Width = 86
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRecruitment)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to change the recruitment settings."
				End If
				.Visible = True
				.Caption = "Espionage"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraRecruitment.AddChild(CType(chkEspionage, UIControl))

			'lblBillboard initial props
			lblBillboard = New UILabel(MyBase.moUILib)
			With lblBillboard
				.ControlName = "lblBillboard"
				.Left = 10
				.Top = 235
				.Width = 232
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Billboard Description:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraRecruitment.AddChild(CType(lblBillboard, UIControl))

			'txtBillboard initial props
			txtBillboard = New UITextBox(MyBase.moUILib)
			With txtBillboard
				.ControlName = "txtBillboard"
				.Left = 10
				.Top = 260
				.Width = 235
				.Height = 220
				.Enabled = True
				.Visible = True
				.Caption = "Looking for a few good empires!"
				.ForeColor = muSettings.InterfaceTextBoxForeColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(0, DrawTextFormat)
				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 200
				.BorderColor = muSettings.InterfaceBorderColor
				.MultiLine = True
				.Locked = True
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Locked = Not goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRecruitment)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to change the recruitment settings."
				End If
			End With
			fraRecruitment.AddChild(CType(txtBillboard, UIControl))

			'btnUpdate initial props
			btnUpdate = New UIButton(MyBase.moUILib)
			With btnUpdate
				.ControlName = "btnUpdate"
				.Left = txtBillboard.Left + txtBillboard.Width - 100
				.Top = txtBillboard.Top + txtBillboard.Height + 5
				.Width = 100
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRecruitment)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to change the recruitment settings."
				End If
				.Visible = True
				.Caption = "Update"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
            fraRecruitment.AddChild(CType(btnUpdate, UIControl))

            'btnBillboard initial props
            btnBillboard = New UIButton(MyBase.moUILib)
            With btnBillboard
                .ControlName = "btnBillboard"
                .Left = txtBillboard.Left '+ txtBillboard.Width - 100
                .Top = btnUpdate.Top
                .Width = 100
                .Height = 24
                .Enabled = False
                If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
                    .Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRecruitment)
                End If
                If .Enabled = False Then
                    .ToolTipText = "You lack the rights to change the recruitment settings."
                End If
                .Visible = True
                .Caption = "Placement"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            fraRecruitment.AddChild(CType(btnBillboard, UIControl))

			SetDetails()
			mbEditBillboard = False
		End Sub

		Private Sub SetDetails()
			If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
				With goCurrentPlayer.oGuild
					chkDiplomacy.Value = (.iRecruitFlags And eiRecruitmentFlags.RecruitDiplomacy) <> 0
					chkEspionage.Value = (.iRecruitFlags And eiRecruitmentFlags.RecruitEspionage) <> 0
					chkMilitary.Value = (.iRecruitFlags And eiRecruitmentFlags.RecruitMilitary) <> 0
					chkProduction.Value = (.iRecruitFlags And eiRecruitmentFlags.RecruitProduction) <> 0
					chkWillTrain.Value = (.iRecruitFlags And eiRecruitmentFlags.WillTrain) <> 0
					chkTrade.Value = (.iRecruitFlags And eiRecruitmentFlags.RecruitTrade) <> 0
					chkResearch.Value = (.iRecruitFlags And eiRecruitmentFlags.RecruitResearch) <> 0
					chkRecruitTutorial.Value = (.iRecruitFlags And eiRecruitmentFlags.RecruitingTutorial) <> 0
					chkRecruiting.Value = (.iRecruitFlags And eiRecruitmentFlags.GuildRecruiting) <> 0
                End With
			End If
		End Sub

		Private Sub FillMembers()
			lstMembers.Clear()
			If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
				With goCurrentPlayer.oGuild
					lstMembers.sHeaderRow = "Member Name and Title".PadRight(30, " "c) & "Member Rank"

					Dim lSortType As GetSortedIndexArrayType
					Dim bDesc As Boolean = False
					Select Case myMemberSort
						Case 1
							lSortType = GetSortedIndexArrayType.eGuildMemberListName
							bDesc = True
						Case 2
							lSortType = GetSortedIndexArrayType.eGuildMemberRankName
						Case 3
							lSortType = GetSortedIndexArrayType.eGuildMemberRankName
							bDesc = True
						Case Else
							lSortType = GetSortedIndexArrayType.eGuildMemberListName
					End Select

					Dim lSorted() As Int32 = GetSortedIndexArrayNoIdxArray(.moMembers, lSortType, bDesc)
					If lSorted Is Nothing = False Then
						For X As Int32 = 0 To lSorted.GetUpperBound(0)
							If lSorted(X) <> -1 Then
								Dim oMember As GuildMember = .moMembers(lSorted(X))
								If oMember Is Nothing = False Then
									lstMembers.AddItem(oMember.MemberListText, False)
									lstMembers.ItemData(lstMembers.NewIndex) = oMember.lMemberID
								End If
							End If
						Next X
					End If
				End With
			End If

		End Sub

		Private Sub SendGuildMemberStatusMsg(ByVal lMemberID As Int32, ByVal lStatusUpdate As Int32)
			Dim yMsg(13) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eGuildMemberStatus).CopyTo(yMsg, lPos) : lPos += 2
			If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
				System.BitConverter.GetBytes(goCurrentPlayer.oGuild.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
			Else : Return
			End If
			System.BitConverter.GetBytes(lMemberID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(lStatusUpdate).CopyTo(yMsg, lPos) : lPos += 4
			MyBase.moUILib.SendMsgToPrimary(yMsg)
		End Sub

		Private Sub btnAccept_Click(ByVal sName As String) Handles btnAccept.Click
			If lstMembers Is Nothing = False Then
				If lstMembers.ListIndex > -1 Then
					Dim lID As Int32 = lstMembers.ItemData(lstMembers.ListIndex)
					If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return
					Dim oMember As GuildMember = goCurrentPlayer.oGuild.GetMember(lID)
					If oMember Is Nothing Then Return
                    If (oMember.yMemberState And GuildMemberState.Applied) <> 0 AndAlso (oMember.yMemberState And GuildMemberState.Approved) = 0 Then
                        'ok, do it
                        SendGuildMemberStatusMsg(lID, GuildMemberState.Approved)
                    Else : MyBase.moUILib.AddNotification("Select an applicant in the list that has applied.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    End If
				Else : MyBase.moUILib.AddNotification("Select an applicant in the list.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				End If
			End If
		End Sub

		Private Sub btnDemote_Click(ByVal sName As String) Handles btnDemote.Click
			If lstMembers Is Nothing = False Then
				If lstMembers.ListIndex > -1 Then
					Dim lID As Int32 = lstMembers.ItemData(lstMembers.ListIndex)
					If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return
					Dim oMember As GuildMember = goCurrentPlayer.oGuild.GetMember(lID)
					If oMember Is Nothing Then Return
					Dim oRank As GuildRank = goCurrentPlayer.oGuild.GetRankByID(oMember.lRankID)
					Dim oMyRank As GuildRank = goCurrentPlayer.oGuild.GetRankByID(goCurrentPlayer.oGuild.lCurrentRankID)
					If oRank Is Nothing OrElse oMyRank Is Nothing Then Return
                    If oRank.yPosition > oMyRank.yPosition Then
                        If (goCurrentPlayer.oGuild.lBaseGuildRules And elGuildFlags.DemoteMember_RA) = 0 Then
                            MyBase.moUILib.AddNotification("Demotion in this guild is only allowed through vote!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                            Return
                        End If

                        SendGuildMemberStatusMsg(lID, -1)
                    Else
                        MyBase.moUILib.AddNotification("You can only demote members of less rank than yourself.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    End If
				Else : MyBase.moUILib.AddNotification("Select an applicant in the list.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				End If
			End If
		End Sub

		Private Sub btnInvite_Click(ByVal sName As String) Handles btnInvite.Click
			Dim sValue As String = txtInvite.Caption
			If sValue.Length < 3 Then Return

            Dim yMsg(21) As Byte
			Dim lPos As Int32 = 0

			System.BitConverter.GetBytes(GlobalMessageCode.eInvitePlayerToGuild).CopyTo(yMsg, lPos) : lPos += 2
			If sValue.Length > 20 Then sValue = sValue.Substring(0, 20)
			System.Text.ASCIIEncoding.ASCII.GetBytes(sValue).CopyTo(yMsg, lPos) : lPos += 20

			MyBase.moUILib.SendMsgToPrimary(yMsg)
			MyBase.moUILib.AddNotification("Invitation Request Sent...", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		End Sub

		Private Sub btnInviteBrowse_Click(ByVal sName As String) Handles btnInviteBrowse.Click
			Dim ofrm As frmAddressBook = New frmAddressBook(MyBase.moUILib)
			ofrm.Visible = True
			ofrm.Left = Me.Left + (Me.Width \ 2) - (ofrm.Width \ 2)
			ofrm.Top = Me.Top + (Me.Height \ 2) - (ofrm.Height \ 2)
			AddHandler ofrm.ContactSelected, AddressOf frmAddressBook_ContactSelected
		End Sub

		Private Sub frmAddressBook_ContactSelected(ByVal sName As String)
			txtInvite.Caption = sName
		End Sub

		Private Sub btnPromote_Click(ByVal sName As String) Handles btnPromote.Click
			If lstMembers Is Nothing = False Then
				If lstMembers.ListIndex > -1 Then
					Dim lID As Int32 = lstMembers.ItemData(lstMembers.ListIndex)
					If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return
					Dim oMember As GuildMember = goCurrentPlayer.oGuild.GetMember(lID)
					If oMember Is Nothing Then Return
					Dim oRank As GuildRank = goCurrentPlayer.oGuild.GetRankByID(oMember.lRankID)
					Dim oMyRank As GuildRank = goCurrentPlayer.oGuild.GetRankByID(goCurrentPlayer.oGuild.lCurrentRankID)
					If oRank Is Nothing OrElse oMyRank Is Nothing Then Return
                    If oRank.yPosition > oMyRank.yPosition Then

                        If (goCurrentPlayer.oGuild.lBaseGuildRules And elGuildFlags.PromoteGuildMember_RA) = 0 Then
                            MyBase.moUILib.AddNotification("Promotion in this guild is only allowed through vote!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                            Return
                        End If

                        SendGuildMemberStatusMsg(lID, -2)
                    Else
                        MyBase.moUILib.AddNotification("You can only promote members of less rank than yourself.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    End If
				Else : MyBase.moUILib.AddNotification("Select an applicant in the list.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				End If
			End If
		End Sub

        Private Sub RemoveMyselfConfirm(ByVal lResult As MsgBoxResult)
            If lResult = MsgBoxResult.Yes Then
                If lstMembers Is Nothing = False Then
                    If lstMembers.ListIndex > -1 Then
                        Dim lID As Int32 = lstMembers.ItemData(lstMembers.ListIndex)
                        If lID < 1 Then Return
                        SendGuildMemberStatusMsg(lID, Int32.MinValue)
                    End If
                End If
            End If
        End Sub

		Private Sub btnRemove_Click(ByVal sName As String) Handles btnRemove.Click
			If lstMembers Is Nothing = False Then
                If lstMembers.ListIndex > -1 Then

                    Dim lID As Int32 = lstMembers.ItemData(lstMembers.ListIndex)
                    Dim oMember As GuildMember = goCurrentPlayer.oGuild.GetMember(lID)
                    If oMember Is Nothing Then Return
                    If (oMember.yMemberState And GuildMemberState.Applied) <> 0 Then
                        Dim oMsgBox As New frmMsgBox(goUILib, "Are you sure you wish to reject this applicant?", MsgBoxStyle.YesNo, "Confirm Application Reject")
                        oMsgBox.Visible = True
                        AddHandler oMsgBox.DialogClosed, AddressOf RemoveMyselfConfirm
                        Return
                    Else
                        If lID = glPlayerID Then
                            Dim oMsgBox As New frmMsgBox(goUILib, "You will be unable to join another guild for 3 days after being kicked or leaving a guild." & vbCrLf & vbCrLf & "Do you wish to leave the guild?", MsgBoxStyle.YesNo, "Confirm Guild Removal")
                            oMsgBox.Visible = True
                            AddHandler oMsgBox.DialogClosed, AddressOf RemoveMyselfConfirm
                            Return
                        Else
                            Dim oMsgBox As New frmMsgBox(goUILib, "The player will be unable to join another guild for 3 days after being kicked or leaving a guild." & vbCrLf & vbCrLf & "Do you wish to leave the guild?", MsgBoxStyle.YesNo, "Confirm Guild Removal")
                            oMsgBox.Visible = True
                            AddHandler oMsgBox.DialogClosed, AddressOf RemoveMyselfConfirm
                            Return
                        End If
                    End If
                    If lID < 1 Then Return
                    SendGuildMemberStatusMsg(lID, Int32.MinValue)
                Else : MyBase.moUILib.AddNotification("Select an applicant / guild member in the list.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
            End If
		End Sub

		Private Sub btnUpdate_Click(ByVal sName As String) Handles btnUpdate.Click
			'update the recruitment settings
			Dim iFlags As eiRecruitmentFlags = 0
			If chkRecruiting.Value = True Then iFlags = iFlags Or eiRecruitmentFlags.GuildRecruiting
			If chkRecruitTutorial.Value = True Then iFlags = iFlags Or eiRecruitmentFlags.RecruitingTutorial
			If chkDiplomacy.Value = True Then iFlags = iFlags Or eiRecruitmentFlags.RecruitDiplomacy
			If chkEspionage.Value = True Then iFlags = iFlags Or eiRecruitmentFlags.RecruitEspionage
			If chkMilitary.Value = True Then iFlags = iFlags Or eiRecruitmentFlags.RecruitMilitary
			If chkProduction.Value = True Then iFlags = iFlags Or eiRecruitmentFlags.RecruitProduction
			If chkResearch.Value = True Then iFlags = iFlags Or eiRecruitmentFlags.RecruitResearch
			If chkTrade.Value = True Then iFlags = iFlags Or eiRecruitmentFlags.RecruitTrade
			If chkWillTrain.Value = True Then iFlags = iFlags Or eiRecruitmentFlags.WillTrain

			Dim lLen As Int32 = 0
			If txtBillboard.Caption Is Nothing Then txtBillboard.Caption = ""
			lLen = txtBillboard.Caption.Trim.Length
			Dim yMsg(7 + lLen) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eUpdateGuildRecruitment).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(iFlags).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(lLen).CopyTo(yMsg, lPos) : lPos += 4
			System.Text.ASCIIEncoding.ASCII.GetBytes(txtBillboard.Caption.Trim).CopyTo(yMsg, lPos) : lPos += lLen

			MyBase.moUILib.SendMsgToPrimary(yMsg)
			MyBase.moUILib.AddNotification("Recruitment Settings Submitted.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

			mbEditBillboard = False
			mbEditDiplomacy = False
			mbEditEspionage = False
			mbEditMilitary = False
			mbEditProduction = False
			mbEditRecruiting = False
			mbEditRecruitTutorial = False
			mbEditResearch = False
			mbEditTrade = False
			mbEditWillTrain = False
		End Sub

        Public Overrides Sub NewFrame()
            Try
                If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
                    If lstMembers Is Nothing = False Then

                        Dim lMemberCnt As Int32 = 0
                        If goCurrentPlayer.oGuild.moMembers Is Nothing = False Then
                            With goCurrentPlayer.oGuild
                                For X As Int32 = 0 To .moMembers.GetUpperBound(0)
                                    If .moMembers(X) Is Nothing = False Then lMemberCnt += 1
                                Next X
                            End With
                        End If
                        If lstMembers.ListCount <> lMemberCnt Then FillMembers()

                        If lstMembers.ListIndex > -1 Then
                            Dim lID As Int32 = lstMembers.ItemData(lstMembers.ListIndex)
                            Dim oMember As GuildMember = goCurrentPlayer.oGuild.GetMember(lID)
                            If oMember Is Nothing = False Then
                                If oMember.DetailsText <> txtDetails.Caption Then txtDetails.Caption = oMember.sDetails
                            End If
                        End If
                        For X As Int32 = 0 To lstMembers.ListCount - 1
                            Dim lID As Int32 = lstMembers.ItemData(X)
                            Dim oMember As GuildMember = goCurrentPlayer.oGuild.GetMember(lID)
                            If oMember Is Nothing = False Then
                                If lstMembers.List(X) <> oMember.MemberListText Then lstMembers.List(X) = oMember.MemberListText
                            End If
                        Next X
                    End If

                    If mbEditBillboard = False Then
                        If txtBillboard.Caption <> goCurrentPlayer.oGuild.sBillboard Then
                            txtBillboard.Caption = goCurrentPlayer.oGuild.sBillboard
                            mbEditBillboard = False
                        End If
                    End If
                    If mbEditDiplomacy = False Then
                        Dim bValue As Boolean = ((goCurrentPlayer.oGuild.iRecruitFlags And eiRecruitmentFlags.RecruitDiplomacy) <> 0)
                        If chkDiplomacy.Value <> bValue Then chkDiplomacy.Value = bValue
                    End If
                    If mbEditEspionage = False Then
                        Dim bValue As Boolean = ((goCurrentPlayer.oGuild.iRecruitFlags And eiRecruitmentFlags.RecruitEspionage) <> 0)
                        If chkEspionage.Value <> bValue Then chkEspionage.Value = bValue
                    End If
                    If mbEditMilitary = False Then
                        Dim bValue As Boolean = ((goCurrentPlayer.oGuild.iRecruitFlags And eiRecruitmentFlags.RecruitMilitary) <> 0)
                        If chkMilitary.Value <> bValue Then chkMilitary.Value = bValue
                    End If
                    If mbEditProduction = False Then
                        Dim bValue As Boolean = ((goCurrentPlayer.oGuild.iRecruitFlags And eiRecruitmentFlags.RecruitProduction) <> 0)
                        If chkProduction.Value <> bValue Then chkProduction.Value = bValue
                    End If
                    If mbEditRecruiting = False Then
                        Dim bValue As Boolean = ((goCurrentPlayer.oGuild.iRecruitFlags And eiRecruitmentFlags.GuildRecruiting) <> 0)
                        If chkRecruiting.Value <> bValue Then chkRecruiting.Value = bValue
                    End If
                    If mbEditRecruitTutorial = False Then
                        Dim bValue As Boolean = ((goCurrentPlayer.oGuild.iRecruitFlags And eiRecruitmentFlags.RecruitingTutorial) <> 0)
                        If chkRecruitTutorial.Value <> bValue Then chkRecruitTutorial.Value = bValue
                    End If
                    If mbEditResearch = False Then
                        Dim bValue As Boolean = ((goCurrentPlayer.oGuild.iRecruitFlags And eiRecruitmentFlags.RecruitResearch) <> 0)
                        If chkResearch.Value <> bValue Then chkResearch.Value = bValue
                    End If
                    If mbEditTrade = False Then
                        Dim bValue As Boolean = ((goCurrentPlayer.oGuild.iRecruitFlags And eiRecruitmentFlags.RecruitTrade) <> 0)
                        If chkTrade.Value <> bValue Then chkTrade.Value = bValue
                    End If
                    If mbEditWillTrain = False Then
                        Dim bValue As Boolean = ((goCurrentPlayer.oGuild.iRecruitFlags And eiRecruitmentFlags.WillTrain) <> 0)
                        If chkWillTrain.Value <> bValue Then chkWillTrain.Value = bValue
                    End If
                End If
            Catch
            End Try
        End Sub

		Public Overrides Sub RenderEnd()

		End Sub

		Private Sub txtBillboard_TextChanged() Handles txtBillboard.TextChanged
			mbEditBillboard = True
		End Sub

		Private Sub chkDiplomacy_Click() Handles chkDiplomacy.Click
			mbEditDiplomacy = True
		End Sub

		Private Sub chkEspionage_Click() Handles chkEspionage.Click
			mbEditEspionage = True
		End Sub

		Private Sub chkMilitary_Click() Handles chkMilitary.Click
			mbEditMilitary = True
		End Sub

		Private Sub chkProduction_Click() Handles chkProduction.Click
			mbEditProduction = True
		End Sub

		Private Sub chkRecruiting_Click() Handles chkRecruiting.Click
			mbEditRecruiting = True
		End Sub

		Private Sub chkRecruitTutorial_Click() Handles chkRecruitTutorial.Click
			mbEditRecruitTutorial = True
		End Sub

		Private Sub chkResearch_Click() Handles chkResearch.Click
			mbEditResearch = True
		End Sub

		Private Sub chkTrade_Click() Handles chkTrade.Click
			mbEditTrade = True
		End Sub

		Private Sub chkWillTrain_Click() Handles chkWillTrain.Click
			mbEditWillTrain = True
		End Sub

		Private Sub lstMembers_HeaderRowClick(ByVal lX As Integer) Handles lstMembers.HeaderRowClick
			If lX > 245 Then
				'ok, sort by rank
				If myMemberSort = 2 Then myMemberSort = 3 Else myMemberSort = 2
			Else
				'sort by name
				If myMemberSort = 0 Then myMemberSort = 1 Else myMemberSort = 0
			End If
			FillMembers()
		End Sub

		Private Sub lstMembers_ItemClick(ByVal lIndex As Integer) Handles lstMembers.ItemClick
			If lstMembers.ListIndex > -1 Then
				Dim bEnabled As Boolean = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					bEnabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.RemoveMember)
				End If
				bEnabled = bEnabled OrElse (lstMembers.ItemData(lstMembers.ListIndex) = glPlayerID)
				btnRemove.Enabled = bEnabled
			End If
		End Sub

        Private Sub btnBillboard_Click(ByVal sName As String) Handles btnBillboard.Click
            Dim ofrm As New frmGuildBillboardBid(goUILib)
            ofrm.Visible = True
        End Sub
    End Class
End Class
