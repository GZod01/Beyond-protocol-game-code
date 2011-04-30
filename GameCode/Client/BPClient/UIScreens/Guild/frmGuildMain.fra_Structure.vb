Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmGuildMain

	Private Class fra_Structure
		Inherits guildframe

		Private fraRanks As UIWindow
		Private fraPermissions As UIWindow
		Private fraRankLog As UIWindow
		Private lblRankName As UILabel
		Private lblVoteRate As UILabel
		Private txtRankName As UITextBox
		Private txtVoteRate As UITextBox
		Private lstRankLog As UIListBox

		Private WithEvents vscrList As UIScrollBar
		Private WithEvents lstRanks As UIListBox
		Private WithEvents btnMoveRankDown As UIButton
		Private WithEvents btnMoveRankUp As UIButton
		Private WithEvents btnRemoveRank As UIButton
		Private WithEvents btnAddRank As UIButton
        Private WithEvents btnUpdateRank As UIButton

        Private lblTaxPercType As UILabel
        Private lblTaxFlat As UILabel
        Private lblTaxPerc As UILabel
        Private txtTaxFlat As UITextBox
        Private txtTaxPerc As UITextBox
        Private cboTaxPercType As UIComboBox

		Private chkRules() As UICheckBox

		Private mlPermVals() As Int32
		Private msPermVals() As String
		Private Sub FillPermVals()
			'Ok, now, let's set up our permissions...
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

			FillPermVals()

            With Me
                .lWindowMetricID = BPMetrics.eWindow.eGuildMain_Structure
                .ControlName = "fra_Structure"
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
                .mbAcceptReprocessEvents = True
            End With

			'fraRanks initial props
			fraRanks = New UIWindow(MyBase.moUILib)
			With fraRanks
				.ControlName = "fraRanks"
				.Left = 15
				.Top = 5
				.Width = 305
				.Height = 370
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.BorderLineWidth = 2
				.Moveable = False
				.Caption = "Ranks"
			End With
			Me.AddChild(CType(fraRanks, UIControl))

			'fraPermissions initial props
			fraPermissions = New UIWindow(MyBase.moUILib)
			With fraPermissions
				.ControlName = "fraPermissions"
				.Left = 340
				.Top = 5
				.Width = 450
				.Height = 370
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.BorderLineWidth = 2
				.Moveable = False
				.Caption = "Rank Permissions"
				.mbAcceptReprocessEvents = True
			End With
			Me.AddChild(CType(fraPermissions, UIControl))

			'fraRankLog initial props
			fraRankLog = New UIWindow(MyBase.moUILib)
			With fraRankLog
				.ControlName = "fraRankLog"
				.Left = 15
				.Top = 385
				.Width = 775
				.Height = 140
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.BorderLineWidth = 2
				.Moveable = False
				.Caption = "Rank Change Log"
			End With
			Me.AddChild(CType(fraRankLog, UIControl))

			'==== fraRanks ====
			'lstRanks initial props
			lstRanks = New UIListBox(MyBase.moUILib)
			With lstRanks
				.ControlName = "lstRanks"
				.Left = 10
				.Top = 15
				.Width = 285
                .Height = 180 '230
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Courier New", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
				.sHeaderRow = "Rank Name".PadRight(24, " "c) & "Members"
			End With
			fraRanks.AddChild(CType(lstRanks, UIControl))

			'btnMoveRankDown initial props
			btnMoveRankDown = New UIButton(MyBase.moUILib)
			With btnMoveRankDown
				.ControlName = "btnMoveRankDown"
				.Left = 10
				.Top = 340
				.Width = 100
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.CreateRanks)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to alter ranks."
				End If
				.Visible = True
				.Caption = "Move Down"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraRanks.AddChild(CType(btnMoveRankDown, UIControl))

			'btnMoveRankUp initial props
			btnMoveRankUp = New UIButton(MyBase.moUILib)
			With btnMoveRankUp
				.ControlName = "btnMoveRankUp"
				.Left = 10
				.Top = 310
				.Width = 100
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.CreateRanks)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to alter ranks."
				End If
				.Visible = True
				.Caption = "Move Up"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraRanks.AddChild(CType(btnMoveRankUp, UIControl))

			'btnRemoveRank initial props
			btnRemoveRank = New UIButton(MyBase.moUILib)
			With btnRemoveRank
				.ControlName = "btnRemoveRank"
				.Left = 200
				.Top = 340
				.Width = 100
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.DeleteRanks)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to delete ranks."
				End If
				.Visible = True
				.Caption = "Delete Rank"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraRanks.AddChild(CType(btnRemoveRank, UIControl))

			'btnAddRank initial props
			btnAddRank = New UIButton(MyBase.moUILib)
			With btnAddRank
				.ControlName = "btnAddRank"
				.Left = 200
				.Top = 310
				.Width = 100
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.CreateRanks)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to create ranks."
				End If
				.Visible = True
				.Caption = "Add Rank"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraRanks.AddChild(CType(btnAddRank, UIControl))

			'btnUpdateRank initial props
			btnUpdateRank = New UIButton(MyBase.moUILib)
			With btnUpdateRank
				.ControlName = "btnUpdateRank"
				.Left = 105
				.Top = 283
				.Width = 100
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRankNames Or RankPermissions.ChangeRankPermissions Or RankPermissions.ChangeRankVotingWeight)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to alter ranks."
				End If
				.Visible = True
				.Caption = "Update Rank"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraRanks.AddChild(CType(btnUpdateRank, UIControl))

			'lblRankName initial props
			lblRankName = New UILabel(MyBase.moUILib)
			With lblRankName
				.ControlName = "lblRankName"
				.Left = 10
                .Top = 205
				.Width = 45
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Name:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraRanks.AddChild(CType(lblRankName, UIControl))

            'lblTaxFlat initial props
            lblTaxFlat = New UILabel(MyBase.moUILib)
            With lblTaxFlat
                .ControlName = "lblTaxFlat"
                .Left = 10
                .Top = 230
                .Width = 60
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Flat Tax:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            fraRanks.AddChild(CType(lblTaxFlat, UIControl))

            'txtTaxFlat initial props
            txtTaxFlat = New UITextBox(MyBase.moUILib)
            With txtTaxFlat
                .ControlName = "txtTaxFlat"
                .Left = lblTaxFlat.Left + lblTaxFlat.Width + 5
                .Top = lblTaxFlat.Top + 2
                .Width = 100
                .Height = 20
                .Enabled = False
                If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
                    .Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRankPermissions)
                End If
                If .Enabled = False Then
                    .ToolTipText = "You lack the rights to alter rank permissions."
                End If
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 9
                .BorderColor = muSettings.InterfaceBorderColor
                .bNumericOnly = True
            End With
            fraRanks.AddChild(CType(txtTaxFlat, UIControl))

            'lblTaxPerc initial props
            lblTaxPerc = New UILabel(MyBase.moUILib)
            With lblTaxPerc
                .ControlName = "lblTaxPerc"
                .Left = txtTaxFlat.Width + txtTaxFlat.Left + 5
                .Top = lblTaxFlat.Top
                .Width = 73
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Rate Tax:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            fraRanks.AddChild(CType(lblTaxPerc, UIControl))

            'txtTaxPerc initial props
            txtTaxPerc = New UITextBox(MyBase.moUILib)
            With txtTaxPerc
                .ControlName = "txtTaxPerc"
                .Left = 260 'lblTaxPerc.Left + lblTaxPerc.Width + 5
                .Top = lblTaxPerc.Top + 2
                .Width = 35
                .Height = 20
                .Enabled = False
                If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
                    .Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRankNames)
                End If
                If .Enabled = False Then
                    .ToolTipText = "You lack the rights to alter rank names."
                End If
                .Visible = True
                .Caption = ""
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                .MaxLength = 3
                .BorderColor = muSettings.InterfaceBorderColor
                .bNumericOnly = True
            End With
            fraRanks.AddChild(CType(txtTaxPerc, UIControl))

			'txtRankName initial props
			txtRankName = New UITextBox(MyBase.moUILib)
			With txtRankName
				.ControlName = "txtRankName"
				.Left = 60
                .Top = 207
				.Width = 115
				.Height = 20
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRankNames)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to alter rank names."
				End If
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
			fraRanks.AddChild(CType(txtRankName, UIControl))

            'lblTaxPercType initial props
            lblTaxPercType = New UILabel(MyBase.moUILib)
            With lblTaxPercType
                .ControlName = "lblTaxPercType"
                .Left = 10
                .Top = 255
                .Width = 73
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Based On:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            fraRanks.AddChild(CType(lblTaxPercType, UIControl))

			'lblVoteRate initial props
			lblVoteRate = New UILabel(MyBase.moUILib)
			With lblVoteRate
				.ControlName = "lblVoteRate"
				.Left = 180
                .Top = 205
				.Width = 73
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Vote Rate:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraRanks.AddChild(CType(lblVoteRate, UIControl))

			'txtVoteRate initial props
			txtVoteRate = New UITextBox(MyBase.moUILib)
			With txtVoteRate
				.ControlName = "txtVoteRate"
				.Left = 260
                .Top = 207
				.Width = 35
				.Height = 20
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRankVotingWeight)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to alter rank voting weight."
				End If
				.Visible = True
				.Caption = "1"
				.ForeColor = muSettings.InterfaceTextBoxForeColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
				.MaxLength = 4
				.BorderColor = muSettings.InterfaceBorderColor
				.bNumericOnly = True
			End With
			fraRanks.AddChild(CType(txtVoteRate, UIControl))

			'==== fraPermissions ====
			'vscrList initial props
			vscrList = New UIScrollBar(MyBase.moUILib, True)
			With vscrList
				.ControlName = "vscrList"
				.Left = 424
				.Top = 2
				.Width = 24
				.Height = 366
				.Enabled = True
				.Visible = True
				.Value = 0
				.MaxValue = 100
				.MinValue = 0
				.SmallChange = 1
				.LargeChange = 1
				.ReverseDirection = True
				.mbAcceptReprocessEvents = True
			End With
			fraPermissions.AddChild(CType(vscrList, UIControl))

			'==== fraRankLog ====
			'lstRankLog initial props
			lstRankLog = New UIListBox(MyBase.moUILib)
			With lstRankLog
				.ControlName = "lstRankLog"
				.Left = 5
				.Top = 10
				.Width = 765
				.Height = 125
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
			End With
            fraRankLog.AddChild(CType(lstRankLog, UIControl))

            cboTaxPercType = New UIComboBox(MyBase.moUILib)
            With cboTaxPercType
                .ControlName = "cboTaxPercType"
                .Left = lblTaxPercType.Left + lblTaxPercType.Width + 5
                .Top = lblTaxPercType.Top + 2
                .Width = 155
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
            End With
            fraRanks.AddChild(CType(cboTaxPercType, UIControl))

            With cboTaxPercType
                .Clear()
                .AddItem("Cash Flow") : .ItemData(.NewIndex) = eyGuildTaxPercType.CashFlow
                .AddItem("Total Income") : .ItemData(.NewIndex) = eyGuildTaxPercType.TotalIncome
                .AddItem("Treasury") : .ItemData(.NewIndex) = eyGuildTaxPercType.Treasury
            End With
		End Sub

		Private Sub btnClick_Click(ByVal sName As String) Handles btnAddRank.Click, btnMoveRankDown.Click, btnMoveRankUp.Click, btnRemoveRank.Click, btnUpdateRank.Click
			Dim yMoveVal As Byte = 0
			Dim sNewName As String = ""
			Dim lNewVoteStr As Int32 = Int32.MinValue
            Dim bRequiresRank As Boolean = True
            Dim lTaxFlat As Int32 = 0
            Dim yTaxPerc As Byte = 0
            Dim yTaxPercType As Byte = 0

			Select Case sName.ToUpper
				Case "BTNMOVERANKDOWN"
					yMoveVal = 2
				Case "BTNMOVERANKUP"
					yMoveVal = 1
				Case "BTNREMOVERANK"
					yMoveVal = 255
				Case "BTNADDRANK"
					yMoveVal = 128
					sNewName = txtRankName.Caption
                    bRequiresRank = False
                    lTaxFlat = CInt(Val(txtTaxFlat.Caption))

                    Try
                        Dim lTmp As Int32 = Math.Min(Math.Max(CInt(Val(txtTaxPerc.Caption)), 0), 100)

                        yTaxPerc = CByte(lTmp)

                        If cboTaxPercType.ListIndex > -1 Then yTaxPercType = CByte(cboTaxPercType.ItemData(cboTaxPercType.ListIndex))
                    Catch
                    End Try
				Case "BTNUPDATERANK"
					bRequiresRank = True
					sNewName = txtRankName.Caption
                    lNewVoteStr = CInt(Val(txtVoteRate.Caption))
                    lTaxFlat = CInt(Val(txtTaxFlat.Caption))

                    Try
                        Dim lTmp As Int32 = Math.Min(Math.Max(CInt(Val(txtTaxPerc.Caption)), 0), 100)
                        yTaxPerc = CByte(lTmp)

                        If cboTaxPercType.ListIndex > -1 Then yTaxPercType = CByte(cboTaxPercType.ItemData(cboTaxPercType.ListIndex))
                    Catch
                    End Try
			End Select

			Dim lRankID As Int32 = -1
			If bRequiresRank = True Then
				If lstRanks.ListIndex < 0 Then Return
                lRankID = lstRanks.ItemData(lstRanks.ListIndex)
                If yMoveVal = 255 Then
                    If lstRanks.ListCount < 2 Then
                        MyBase.moUILib.AddNotification("Unable to delete the last rank in the guild.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                        Return
                    End If
                End If
				Dim oRank As GuildRank = Nothing
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then oRank = goCurrentPlayer.oGuild.GetRankByID(lRankID)
				If oRank Is Nothing Then
					MyBase.moUILib.AddNotification("Select a rank to update first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				End If
			End If

            StructureSendRankUpdate(lRankID, yMoveVal, sNewName, lNewVoteStr, lTaxFlat, yTaxPerc, yTaxPercType)
		End Sub

		Private Sub lstRanks_ItemClick(ByVal lIndex As Integer) Handles lstRanks.ItemClick

			'Clear all children except vscrList
			fraPermissions.RemoveAllChildren()
			If vscrList Is Nothing = False Then fraPermissions.AddChild(CType(vscrList, UIControl))

			If lIndex > -1 Then
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False AndAlso goCurrentPlayer.oGuild.moRanks Is Nothing = False Then
					Dim oRank As GuildRank = Nothing

					Dim lID As Int32 = lstRanks.ItemData(lIndex)
					oRank = goCurrentPlayer.oGuild.GetRankByID(lID)
					If oRank Is Nothing = False Then

						txtRankName.Caption = oRank.sRankName
                        txtVoteRate.Caption = oRank.lVoteStrength.ToString

                        txtTaxFlat.Caption = oRank.TaxRateFlat.ToString("#,##0")
                        txtTaxPerc.Caption = oRank.TaxRatePercentage.ToString("##0")
                        cboTaxPercType.FindComboItemData(CInt(oRank.TaxRatePercType))


						FillCheckBoxes(oRank)
					End If
				End If

			End If
		End Sub

		Private Sub vscrList_ValueChange() Handles vscrList.ValueChange
			If lstRanks Is Nothing = False AndAlso lstRanks.ListIndex > -1 Then
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False AndAlso goCurrentPlayer.oGuild.moRanks Is Nothing = False Then
					Dim oRank As GuildRank = Nothing

					Dim lID As Int32 = lstRanks.ItemData(lstRanks.ListIndex)
					oRank = goCurrentPlayer.oGuild.GetRankByID(lID)
					If oRank Is Nothing = False Then
						'FillCheckBoxes(oRank)
						SetCheckBoxValues(oRank)
						Me.IsDirty = True
					End If
				End If

			End If
		End Sub

		Private Sub StructureCheckClick()
			If lstRanks.ListIndex > -1 Then
				Dim lID As Int32 = lstRanks.ItemData(lstRanks.ListIndex)
				Dim oRank As GuildRank = goCurrentPlayer.oGuild.GetRankByID(lID)

				If goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRankPermissions) = False Then
					MyBase.moUILib.AddNotification("You do not have permission to change rank rights.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				End If

				If oRank Is Nothing = False AndAlso chkRules Is Nothing = False Then
					'ok, check our values vs. the values that the rank currently has
					Dim lVals() As Int32 = Nothing
					Dim lValUB As Int32 = -1

					For X As Int32 = 0 To 16
						Dim lRuleIdx As Int32 = X + vscrList.Value

						With chkRules(X)
                            Dim lVal As Int32 = mlPermVals(lRuleIdx)
							If .Value = True Then
								If (oRank.lRankPermissions And lVal) = 0 Then
									'Ok, need to alter the permission
									lValUB += 1
									ReDim Preserve lVals(lValUB)
									lVals(lValUB) = lVal
								End If
							Else
								If (oRank.lRankPermissions And lVal) <> 0 Then
									'Ok, need to alter the permission
									lValUB += 1
									ReDim Preserve lVals(lValUB)
									lVals(lValUB) = -lVal
								End If
							End If
						End With
					Next X 

					If lValUB = -1 Then Return

					Dim yMsg(9 + ((lValUB + 1) * 4)) As Byte
					Dim lPos As Int32 = 0
					System.BitConverter.GetBytes(GlobalMessageCode.eUpdateRankPermission).CopyTo(yMsg, lPos) : lPos += 2
					System.BitConverter.GetBytes(oRank.lRankID).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(lValUB + 1).CopyTo(yMsg, lPos) : lPos += 4
					For X As Int32 = 0 To lValUB
						System.BitConverter.GetBytes(lVals(X)).CopyTo(yMsg, lPos) : lPos += 4
					Next X
					MyBase.moUILib.SendMsgToPrimary(yMsg)
				End If
			End If
		End Sub

		Public Overrides Sub NewFrame()
			'Rank Name - number of players in rank
			If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
				If goCurrentPlayer.oGuild.moRanks Is Nothing Then Return

				If lstRanks Is Nothing Then Return

				Dim lCnt As Int32 = 0
				For X As Int32 = 0 To goCurrentPlayer.oGuild.moRanks.GetUpperBound(0)
					If goCurrentPlayer.oGuild.moRanks(X) Is Nothing = False Then lCnt += 1
				Next X

				If lstRanks.ListCount <> lCnt Then
					lstRanks.Clear()
					With goCurrentPlayer.oGuild
						Dim bDone As Boolean = False
						Dim lLastPosition As Int32 = -1
						While bDone = False
							Dim lCurrIdx As Int32 = -1
							Dim lCurrPosition As Int32 = Int32.MaxValue
							For X As Int32 = 0 To .moRanks.GetUpperBound(0)
								If .moRanks(X) Is Nothing = False Then
									If .moRanks(X).yPosition < lCurrPosition AndAlso .moRanks(X).yPosition > lLastPosition Then
										lCurrPosition = .moRanks(X).yPosition
										lCurrIdx = X
									End If
								End If
							Next X
							If lCurrIdx = -1 Then
								bDone = True
							Else
								lstRanks.AddItem(.moRanks(lCurrIdx).ListBoxDisplay, False)
								lstRanks.ItemData(lstRanks.NewIndex) = .moRanks(lCurrIdx).lRankID
								lLastPosition = lCurrPosition
							End If
						End While
					End With

				Else
					'Now, go through all of our items
					Dim lLastPosition As Int32 = -1
					Dim lCurrPosition As Int32 = Int32.MaxValue
					Dim lCurrIdx As Int32 = -1
					Dim lListIdx As Int32 = 0

					If lstRanks.ListCount = 0 Then Return

					With goCurrentPlayer.oGuild
						Dim bDone As Boolean = False
						While bDone = False
							lCurrIdx = -1
							lCurrPosition = Int32.MaxValue
							For X As Int32 = 0 To .moRanks.GetUpperBound(0)
								If .moRanks(X) Is Nothing = False Then
									If .moRanks(X).yPosition < lCurrPosition AndAlso .moRanks(X).yPosition > lLastPosition Then
										lCurrPosition = .moRanks(X).yPosition
										lCurrIdx = X
									End If
								End If
							Next X
							If lCurrIdx = -1 Then
								bDone = True
							Else
								Dim oRank As GuildRank = .moRanks(lCurrIdx)
								lLastPosition = oRank.yPosition

								'now, check our list
								If lstRanks.ItemData(lListIdx) <> .moRanks(lCurrIdx).lRankID Then
									lstRanks.Clear()
									Return
								End If
								Dim sTemp As String = oRank.ListBoxDisplay()
								If lstRanks.List(lListIdx) <> sTemp Then lstRanks.List(lListIdx) = sTemp

								lListIdx += 1
							End If
						End While
					End With
				End If

				If lstRanks.ListIndex > -1 Then
					Dim lRankID As Int32 = lstRanks.ItemData(lstRanks.ListIndex)
					Dim oRank As GuildRank = goCurrentPlayer.oGuild.GetRankByID(lRankID)
					If oRank Is Nothing = False AndAlso chkRules Is Nothing = False Then
						Try
							If mlPermVals Is Nothing Then FillPermVals()
							For X As Int32 = 0 To chkRules.GetUpperBound(0)
								Dim lRuleIdx As Int32 = X + vscrList.Value
								Dim lVal As Int32 = mlPermVals(lRuleIdx)
								Dim bValue As Boolean = (oRank.lRankPermissions And lVal) <> 0
								If chkRules(X).Value <> bValue Then chkRules(X).Value = bValue
							Next X
						Catch
						End Try
					End If
				End If

			End If
		End Sub

        Private Sub StructureSendRankUpdate(ByVal lRankID As Int32, ByVal yMove As Byte, ByVal sNewName As String, ByVal lNewVoteStr As Int32, ByVal lNewTaxFlat As Int32, ByVal yNewTaxPerc As Byte, ByVal yNewTaxPercType As Byte)
            'ymove = 0 for nothing, 1 for up, 2 for down, 255 for remove
            'sNewName = "" if no change, otherwise, the actual value
            'lNewVoteStr = int32.minvalue if no change, otherwise, the value
            Dim yChanges As Byte = 0
            Dim oRank As GuildRank = Nothing
            If goCurrentPlayer.oGuild Is Nothing = False Then
                oRank = goCurrentPlayer.oGuild.GetRankByID(lRankID)
            End If
            If yMove = 0 Then
                If oRank Is Nothing = False Then
                    If oRank.TaxRateFlat <> lNewTaxFlat OrElse oRank.TaxRatePercentage <> yNewTaxPerc OrElse oRank.TaxRatePercType <> yNewTaxPercType Then
                        yChanges = CByte(yChanges Or 8)
                    End If
                End If
            ElseIf yMove = 128 Then
                yChanges = CByte(yChanges Or 8)
            End If

            If yMove <> 0 Then yChanges = CByte(yChanges Or 1)
            If sNewName Is Nothing = False AndAlso sNewName <> "" AndAlso (oRank Is Nothing = True OrElse oRank.sRankName <> sNewName) Then yChanges = CByte(yChanges Or 2)
            If lNewVoteStr <> Int32.MinValue AndAlso (oRank Is Nothing = False AndAlso oRank.lVoteStrength <> lNewVoteStr) Then yChanges = CByte(yChanges Or 4)


            Dim lChangeLen As Int32 = 0
            If (yChanges And 1) <> 0 Then lChangeLen += 1
            If (yChanges And 2) <> 0 Then lChangeLen += 20
            If (yChanges And 4) <> 0 Then lChangeLen += 4
            If (yChanges And 8) <> 0 Then lChangeLen += 6 'tax info
            Dim yMsg(6 + lChangeLen) As Byte
            Dim lPos As Int32 = 0

            System.BitConverter.GetBytes(GlobalMessageCode.eUpdateGuildRank).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lRankID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = yChanges : lPos += 1
            If (yChanges And 1) <> 0 Then
                yMsg(lPos) = yMove : lPos += 1
            End If
            If (yChanges And 2) <> 0 Then
                If sNewName.Length > 20 Then sNewName = sNewName.Substring(0, 20)
                System.Text.ASCIIEncoding.ASCII.GetBytes(sNewName).CopyTo(yMsg, lPos) : lPos += 20
            End If
            If (yChanges And 4) <> 0 Then
                System.BitConverter.GetBytes(lNewVoteStr).CopyTo(yMsg, lPos) : lPos += 4
            End If
            If (yChanges And 8) <> 0 Then
                System.BitConverter.GetBytes(lNewTaxFlat).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = yNewTaxPerc : lPos += 1
                yMsg(lPos) = yNewTaxPercType : lPos += 1
            End If

            MyBase.moUILib.SendMsgToPrimary(yMsg)

            MyBase.moUILib.AddNotification("Rank Changes Submitted.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

        End Sub

		Public Overrides Sub RenderEnd()

		End Sub

		Private Sub SetCheckBoxValues(ByRef oRank As GuildRank)
			If mlPermVals Is Nothing OrElse chkRules Is Nothing Then
				FillCheckBoxes(oRank)
				Return
			End If

			Dim bEnabled As Boolean = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRankPermissions)

			For X As Int32 = 0 To 16

				Dim lRuleIdx As Int32 = X + vscrList.Value

				Dim lTop As Int32 = 10 + (X * 20)
				If lTop + 24 > fraPermissions.Height Then
					Exit For
				End If
				'chkRule initial props
				With chkRules(X)
					.Enabled = bEnabled
					.Caption = msPermVals(lRuleIdx)
					.Value = (oRank.lRankPermissions And mlPermVals(lRuleIdx)) <> 0
					Dim rcTemp As Rectangle = .GetTextDimensions(.Caption)
					.Width = rcTemp.Width + 32
				End With
			Next X
		End Sub

		Private Sub FillCheckBoxes(ByRef oRank As GuildRank)
			Try
				Dim bDone As Boolean = False
				While bDone = False
					bDone = True
					For X As Int32 = 0 To fraPermissions.ChildrenUB
						If fraPermissions.moChildren(X) Is Nothing = False AndAlso fraPermissions.moChildren(X).ControlName Is Nothing = False Then
							If fraPermissions.moChildren(X).ControlName.ToUpper.StartsWith("CHKRULE") = True Then
								fraPermissions.RemoveChild(X)
								bDone = False
								Exit For
							End If
						End If
					Next X
				End While


				If mlPermVals Is Nothing Then FillPermVals()
				Dim lScrollMax As Int32 = mlPermVals.Length - 17
				Dim bEnabled As Boolean = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ChangeRankPermissions)

				ReDim chkRules(16)

				For X As Int32 = 0 To 16 'mlPermVals.GetUpperBound(0)

					Dim lRuleIdx As Int32 = X + vscrList.Value

					Dim lTop As Int32 = 10 + (X * 20)
					If lTop + 24 > fraPermissions.Height Then
						lScrollMax += 1
						Continue For
					End If
					'chkRule initial props
					chkRules(X) = New UICheckBox(MyBase.moUILib)
					With chkRules(X)
						.ControlName = "chkRule(" & X & ")"	'mlPermVals(lRuleIdx) & ")"
						.Left = 15
						.Top = 10 + (X * 20)
						.Width = 100
						.Height = 24
						.Enabled = bEnabled
						.Visible = True
						.Caption = msPermVals(lRuleIdx)
						.ForeColor = muSettings.InterfaceBorderColor
						.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
						.DrawBackImage = False
						.FontFormat = CType(6, DrawTextFormat)
						.Value = (oRank.lRankPermissions And mlPermVals(lRuleIdx)) <> 0

						Dim rcTemp As Rectangle = .GetTextDimensions(.Caption)
						.Width = rcTemp.Width + 32
					End With
					fraPermissions.AddChild(CType(chkRules(X), UIControl))
					AddHandler chkRules(X).Click, AddressOf StructureCheckClick
				Next X

				CType(vscrList, UIScrollBar).MaxValue = lScrollMax
			Catch
			End Try
		End Sub
	End Class

End Class
