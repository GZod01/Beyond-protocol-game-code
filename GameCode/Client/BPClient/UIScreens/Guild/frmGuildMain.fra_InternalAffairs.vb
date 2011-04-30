Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmGuildMain

	Private Class fra_InternalAffairs
		Inherits guildframe

		Private lblCharter As UILabel
		Private txtCharter As UITextBox
		Private fraProposalDetails As UIWindow
		Private lblProposedBy As UILabel
		Private lblDescription As UILabel
		Private lblVotingStarts As UILabel
		Private lblVotingEnds As UILabel
		Private lblNewValue As UILabel
		Private lblSelectedItem As UILabel
		Private lblTypeOfVote As UILabel
		Private txtProposedBy As UITextBox
		Private txtVoteStarts As UITextBox
		Private txtVoteEnds As UITextBox
		Private txtSummary As UITextBox
		Private lblYourVote As UILabel

		Private txtNewValue As UITextBox
		Private cboNewValue As UIComboBox

		Private WithEvents fraProposals As UIWindow
		Private WithEvents btnNewProposal As UIButton
		Private WithEvents cboSelectedItem As UIComboBox
		Private WithEvents cboTypeOfVote As UIComboBox
		Private WithEvents lstProposals As UIListBox
		Private WithEvents optVoteYes As UIOption
		Private WithEvents optVoteNo As UIOption
		Private WithEvents optVoteAbstain As UIOption
		Private WithEvents btnPropose As UIButton

		Public Sub New(ByRef oUILib As UILib)
			MyBase.New(oUILib)

            With Me
                .lWindowMetricID = BPMetrics.eWindow.eGuildMain_Internal
                .ControlName = "fra_InternalAffairs"
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

			'txtCharter initial props
			txtCharter = New UITextBox(MyBase.moUILib)
			With txtCharter
				.ControlName = "txtCharter"
				.Left = 15
				.Top = 30
				.Width = 400
				.Height = 490
				.Enabled = True
				.Visible = True
				.Caption = "Charter goes here..."
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
				.mbAcceptReprocessEvents = True
			End With
			Me.AddChild(CType(txtCharter, UIControl))

			'lblCharter initial props
			lblCharter = New UILabel(MyBase.moUILib)
			With lblCharter
				.ControlName = "lblCharter"
				.Left = 15
				.Top = 5
				.Width = 55
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Charter"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblCharter, UIControl))

			fraProposals = New UIWindow(MyBase.moUILib)
			With fraProposals
				.ControlName = "fraProposals"
				.Left = txtCharter.Left + txtCharter.Width + 10
				.Top = 5 'txtCharter.Top - 10
				.Width = 360
				.Height = 120
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.BorderLineWidth = 2
				.Moveable = False
			End With
			Me.AddChild(CType(fraProposals, UIControl))

			'lstProposals initial props
			lstProposals = New UIListBox(MyBase.moUILib)
			With lstProposals
				.ControlName = "lstProposals"
				.Left = 5
				.Top = 15
				.Width = 350
				.Height = 100
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
				.mbAcceptReprocessEvents = True
			End With
			fraProposals.AddChild(CType(lstProposals, UIControl))

			'btnNewProposal initial props
			btnNewProposal = New UIButton(MyBase.moUILib)
			With btnNewProposal
				.ControlName = "btnNewProposal"
				.Left = 205
				.Top = -13
				.Width = 100
				.Height = 24
				.Enabled = False
				.Visible = True
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ProposeVotes)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to propose new votes."
				End If
				.Caption = "Propose New"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraProposals.AddChild(CType(btnNewProposal, UIControl))

			'===== end of proposals

			'fraProposalDetails initial props
			fraProposalDetails = New UIWindow(MyBase.moUILib)
			With fraProposalDetails
				.ControlName = "fraProposalDetails"
				.Left = fraProposals.Left
				.Top = fraProposals.Top + fraProposals.Height + 10
				.Width = 360
				.Height = 385
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.BorderLineWidth = 2
				.Moveable = False
				.Caption = "Proposal Details"
			End With
			Me.AddChild(CType(fraProposalDetails, UIControl))

			'lblProposedBy initial props
			lblProposedBy = New UILabel(MyBase.moUILib)
			With lblProposedBy
				.ControlName = "lblProposedBy"
				.Left = 10
				.Top = 10
				.Width = 100
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Proposed By:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraProposalDetails.AddChild(CType(lblProposedBy, UIControl))

			'lblVotingStarts initial props
			lblVotingStarts = New UILabel(MyBase.moUILib)
			With lblVotingStarts
				.ControlName = "lblVotingStarts"
				.Left = 10
				.Top = 40
				.Width = 100
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Voting Starts On:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraProposalDetails.AddChild(CType(lblVotingStarts, UIControl))

			'lblVotingEnds initial props
			lblVotingEnds = New UILabel(MyBase.moUILib)
			With lblVotingEnds
				.ControlName = "lblVotingEnds"
				.Left = 10
				.Top = 70
				.Width = 100
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Voting Ends On:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraProposalDetails.AddChild(CType(lblVotingEnds, UIControl))

			'lblTypeOfVote initial props
			lblTypeOfVote = New UILabel(MyBase.moUILib)
			With lblTypeOfVote
				.ControlName = "lblTypeOfVote"
				.Left = 10
				.Top = 110
				.Width = 100
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Type of Vote:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraProposalDetails.AddChild(CType(lblTypeOfVote, UIControl))

			'txtProposedBy initial props
			txtProposedBy = New UITextBox(MyBase.moUILib)
			With txtProposedBy
				.ControlName = "txtProposedBy"
				.Left = 130
				.Top = 12
				.Width = 220
				.Height = 22
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
			fraProposalDetails.AddChild(CType(txtProposedBy, UIControl))

			'txtVoteStarts initial props
			txtVoteStarts = New UITextBox(MyBase.moUILib)
			With txtVoteStarts
				.ControlName = "txtVoteStarts"
				.Left = 130
				.Top = 42
				.Width = 220
				.Height = 22
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
			End With
			fraProposalDetails.AddChild(CType(txtVoteStarts, UIControl))

			'txtVoteEnds initial props
			txtVoteEnds = New UITextBox(MyBase.moUILib)
			With txtVoteEnds
				.ControlName = "txtVoteEnds"
				.Left = 130
				.Top = 72
				.Width = 220
				.Height = 22
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
				.bNumericOnly = True
			End With
			fraProposalDetails.AddChild(CType(txtVoteEnds, UIControl))

			'lblSelectedItem initial props
			lblSelectedItem = New UILabel(MyBase.moUILib)
			With lblSelectedItem
				.ControlName = "lblSelectedItem"
				.Left = 10
				.Top = 140
				.Width = 100
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Selected Item:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraProposalDetails.AddChild(CType(lblSelectedItem, UIControl))

			'lblNewValue initial props
			lblNewValue = New UILabel(MyBase.moUILib)
			With lblNewValue
				.ControlName = "lblNewValue"
				.Left = 10
				.Top = 170
				.Width = 100
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "New Value:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraProposalDetails.AddChild(CType(lblNewValue, UIControl))

			'lblDescription initial props
			lblDescription = New UILabel(MyBase.moUILib)
			With lblDescription
				.ControlName = "lblDescription"
				.Left = 10
				.Top = 200
				.Width = 120
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Proposal Summary:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraProposalDetails.AddChild(CType(lblDescription, UIControl))

			'txtSummary initial props
			txtSummary = New UITextBox(MyBase.moUILib)
			With txtSummary
				.ControlName = "txtSummary"
				.Left = 10
				.Top = 225
				.Width = 340
				.Height = 100
				.Enabled = True
				.Visible = True
				.Caption = "Proposal Summary"
				.ForeColor = muSettings.InterfaceTextBoxForeColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(0, DrawTextFormat)
				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
				.MaxLength = 255
				.BorderColor = muSettings.InterfaceBorderColor
				.MultiLine = True
			End With
			fraProposalDetails.AddChild(CType(txtSummary, UIControl))

			'txtNewValue initial props
			txtNewValue = New UITextBox(MyBase.moUILib)
			With txtNewValue
				.ControlName = "txtNewValue"
				.Left = 130
				.Top = 172
				.Width = 220
				.Height = 22
				.Enabled = True
				.Visible = False
				.ForeColor = muSettings.InterfaceTextBoxForeColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(0, DrawTextFormat)
				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
				.MaxLength = 0
				.BorderColor = muSettings.InterfaceBorderColor
			End With
			fraProposalDetails.AddChild(CType(txtNewValue, UIControl))

			'lblYourVote initial props
			lblYourVote = New UILabel(MyBase.moUILib)
			With lblYourVote
				.ControlName = "lblYourVote"
				.Left = 10
				.Top = 330
				.Width = 164
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Your Vote On this Proposal:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraProposalDetails.AddChild(CType(lblYourVote, UIControl))

			'optVoteYes initial props
			optVoteYes = New UIOption(MyBase.moUILib)
			With optVoteYes
				.ControlName = "optVoteYes"
				.Left = 30
				.Top = 355
				.Width = 45
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Yes"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraProposalDetails.AddChild(CType(optVoteYes, UIControl))

			'btnPropose initial props
			btnPropose = New UIButton(MyBase.moUILib)
			With btnPropose
				.ControlName = "btnPropose"
				.Left = fraProposalDetails.Width - 110
				.Top = lblYourVote.Top
				.Width = 100
				.Height = 24
				.Enabled = True
				.Visible = False
				.Caption = "Propose"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraProposalDetails.AddChild(CType(btnPropose, UIControl))

			'optVoteNo initial props
			optVoteNo = New UIOption(MyBase.moUILib)
			With optVoteNo
				.ControlName = "optVoteNo"
				.Left = 310
				.Top = 355
				.Width = 39
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "No"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraProposalDetails.AddChild(CType(optVoteNo, UIControl))

			'optVoteAbstain initial props
			optVoteAbstain = New UIOption(MyBase.moUILib)
			With optVoteAbstain
				.ControlName = "optVoteAbstain"
				.Left = 155
				.Top = 355
				.Width = 66
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Abstain"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = True
			End With
			fraProposalDetails.AddChild(CType(optVoteAbstain, UIControl))

			'cboNewValue initial props
			cboNewValue = New UIComboBox(MyBase.moUILib)
			With cboNewValue
				.ControlName = "cboNewValue"
				.Left = 130
				.Top = 172
				.Width = 220
				.Height = 22
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
            fraProposalDetails.AddChild(CType(cboNewValue, UIControl))

            'cboSelectedItem initial props
            cboSelectedItem = New UIComboBox(MyBase.moUILib)
            With cboSelectedItem
                .ControlName = "cboSelectedItem"
                .Left = 130
                .Top = 142
                .Width = 220
                .Height = 22
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
            fraProposalDetails.AddChild(CType(cboSelectedItem, UIControl))

            'cboTypeOfVote initial props
            cboTypeOfVote = New UIComboBox(MyBase.moUILib)
            With cboTypeOfVote
                .ControlName = "cboTypeOfVote"
                .Left = 130
                .Top = 112
                .Width = 220
                .Height = 22
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
			fraProposalDetails.AddChild(CType(cboTypeOfVote, UIControl))

			FillTypeOfVotes()
		End Sub

		Private Sub FillTypeOfVotes()
			cboTypeOfVote.Clear()
			cboTypeOfVote.AddItem("Charter")
			cboTypeOfVote.ItemData(cboTypeOfVote.NewIndex) = eyTypeOfVote.Charter
			cboTypeOfVote.AddItem("Membership")
			cboTypeOfVote.ItemData(cboTypeOfVote.NewIndex) = eyTypeOfVote.Membership
			cboTypeOfVote.AddItem("Change Voting Weight")
			cboTypeOfVote.ItemData(cboTypeOfVote.NewIndex) = eyTypeOfVote.VotingWeights
			cboTypeOfVote.AddItem("Add Rank Permission")
			cboTypeOfVote.ItemData(cboTypeOfVote.NewIndex) = eyTypeOfVote.AddRankPermissions
			cboTypeOfVote.AddItem("Remove Rank Permission")
			cboTypeOfVote.ItemData(cboTypeOfVote.NewIndex) = eyTypeOfVote.RemoveRankPermission
			cboTypeOfVote.AddItem("Free-form")
            cboTypeOfVote.ItemData(cboTypeOfVote.NewIndex) = eyTypeOfVote.FreeFormVote

            Dim yMsg(5) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestGuildVoteProposals).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(ml_IntAffSelectedTab).CopyTo(yMsg, 2)
            MyBase.moUILib.SendMsgToPrimary(yMsg)
        End Sub

		Private Sub btnNewProposal_Click(ByVal sName As String) Handles btnNewProposal.Click
			SetInternalAffairsFromVote(Nothing)
		End Sub

		Private Sub cboSelectedItem_ItemSelected(ByVal lItemIndex As Integer) Handles cboSelectedItem.ItemSelected
			Dim lType As Int32 = 0
			Dim lItem As Int32 = 0

			cboNewValue.Clear()
			cboNewValue.Visible = True
			lblNewValue.Visible = True
			txtNewValue.Visible = False
			txtNewValue.bNumericOnly = False

			If goCurrentPlayer Is Nothing OrElse goCurrentPlayer.oGuild Is Nothing Then Return

			If lItemIndex > -1 Then
				If cboTypeOfVote.ListIndex > -1 Then lType = cboTypeOfVote.ItemData(cboTypeOfVote.ListIndex) Else Return
				lItem = cboSelectedItem.ItemData(cboSelectedItem.ListIndex)

				Select Case lType
					Case eyTypeOfVote.AddRankPermissions
						Dim oRank As GuildRank = goCurrentPlayer.oGuild.GetRankByID(lItem)
						If oRank Is Nothing Then Return
						txtSummary.Caption = "Add the permission indicated to " & oRank.sRankName

						'Permission List...
						With cboNewValue
							'we're adding, so only add permissions that the rank currently does not have
							If (oRank.lRankPermissions And RankPermissions.AcceptApplicant) = 0 Then
								.AddItem("Accept Applicant") : .ItemData(.NewIndex) = RankPermissions.AcceptApplicant
							End If
							If (oRank.lRankPermissions And RankPermissions.AcceptEvents) = 0 Then
								.AddItem("Accept Events") : .ItemData(.NewIndex) = RankPermissions.AcceptEvents
							End If
							If (oRank.lRankPermissions And RankPermissions.BuildGuildBase) = 0 Then
								.AddItem("Build Guild Base") : .ItemData(.NewIndex) = RankPermissions.BuildGuildBase
							End If
							If (oRank.lRankPermissions And RankPermissions.ChangeMOTD) = 0 Then
								.AddItem("Change MOTD") : .ItemData(.NewIndex) = RankPermissions.ChangeMOTD
							End If
							If (oRank.lRankPermissions And RankPermissions.ChangeRankNames) = 0 Then
								.AddItem("Change Rank Names") : .ItemData(.NewIndex) = RankPermissions.ChangeRankNames
							End If
							If (oRank.lRankPermissions And RankPermissions.ChangeRankPermissions) = 0 Then
								.AddItem("Change Rank Permissions") : .ItemData(.NewIndex) = RankPermissions.ChangeRankPermissions
							End If
							If (oRank.lRankPermissions And RankPermissions.ChangeRankVotingWeight) = 0 Then
								.AddItem("Change Rank Voting Weight") : .ItemData(.NewIndex) = RankPermissions.ChangeRankVotingWeight
							End If
							If (oRank.lRankPermissions And RankPermissions.ChangeRecruitment) = 0 Then
								.AddItem("Change Recruitment Settings") : .ItemData(.NewIndex) = RankPermissions.ChangeRecruitment
							End If
							If (oRank.lRankPermissions And RankPermissions.CreateEvents) = 0 Then
								.AddItem("Create Events") : .ItemData(.NewIndex) = RankPermissions.CreateEvents
							End If
							If (oRank.lRankPermissions And RankPermissions.CreateRanks) = 0 Then
								.AddItem("Create/Move Ranks") : .ItemData(.NewIndex) = RankPermissions.CreateRanks
							End If
							If (oRank.lRankPermissions And RankPermissions.DeleteEvents) = 0 Then
								.AddItem("Delete Events") : .ItemData(.NewIndex) = RankPermissions.DeleteEvents
							End If
							If (oRank.lRankPermissions And RankPermissions.DeleteRanks) = 0 Then
								.AddItem("Delete Ranks") : .ItemData(.NewIndex) = RankPermissions.DeleteRanks
							End If
							If (oRank.lRankPermissions And RankPermissions.DemoteMember) = 0 Then
								.AddItem("Demote Member Rank") : .ItemData(.NewIndex) = RankPermissions.DemoteMember
							End If
                            'If (oRank.lRankPermissions And RankPermissions.GuildChatChannel) = 0 Then
                            '	.AddItem("Guild Chat Channel Access") : .ItemData(.NewIndex) = RankPermissions.GuildChatChannel
                            'End If
							If (oRank.lRankPermissions And RankPermissions.PromoteMember) = 0 Then
								.AddItem("Promote Member Rank") : .ItemData(.NewIndex) = RankPermissions.PromoteMember
							End If
							If (oRank.lRankPermissions And RankPermissions.ProposeVotes) = 0 Then
								.AddItem("Propose New Votes") : .ItemData(.NewIndex) = RankPermissions.ProposeVotes
							End If
							If (oRank.lRankPermissions And RankPermissions.RejectMember) = 0 Then
								.AddItem("Reject Applicants") : .ItemData(.NewIndex) = RankPermissions.RejectMember
							End If
							If (oRank.lRankPermissions And RankPermissions.RemoveMember) = 0 Then
								.AddItem("Remove Member") : .ItemData(.NewIndex) = RankPermissions.RemoveMember
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewBankLog) = 0 Then
								.AddItem("View Bank Log") : .ItemData(.NewIndex) = RankPermissions.ViewBankLog
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewEvents) = 0 Then
								.AddItem("View Events") : .ItemData(.NewIndex) = RankPermissions.ViewEvents
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewEventAttachments) = 0 Then
								.AddItem("View Event Attachments") : .ItemData(.NewIndex) = RankPermissions.ViewEventAttachments
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewContentsLowSec) = 0 Then
								.AddItem("View Low Security Bank Contents") : .ItemData(.NewIndex) = RankPermissions.ViewContentsLowSec
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewContentsHiSec) = 0 Then
								.AddItem("View High Security Bank Contents") : .ItemData(.NewIndex) = RankPermissions.ViewContentsHiSec
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewGuildBase) = 0 Then
								.AddItem("View Bank Facility") : .ItemData(.NewIndex) = RankPermissions.ViewGuildBase
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewVotesHistory) = 0 Then
								.AddItem("View Vote History") : .ItemData(.NewIndex) = RankPermissions.ViewVotesHistory
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewVotesInProgress) = 0 Then
								.AddItem("View Votes In Progress") : .ItemData(.NewIndex) = RankPermissions.ViewVotesInProgress
							End If
							If (oRank.lRankPermissions And RankPermissions.WithdrawLowSec) = 0 Then
								.AddItem("Withdraw Low Security Bank") : .ItemData(.NewIndex) = RankPermissions.WithdrawLowSec
							End If
							If (oRank.lRankPermissions And RankPermissions.WithdrawHiSec) = 0 Then
								.AddItem("Withdraw High Security Bank") : .ItemData(.NewIndex) = RankPermissions.WithdrawHiSec
							End If

						End With
					Case eyTypeOfVote.RemoveRankPermission
						Dim oRank As GuildRank = goCurrentPlayer.oGuild.GetRankByID(lItem)
						If oRank Is Nothing Then Return
						txtSummary.Caption = "Remove the permission indicated to " & oRank.sRankName

						'Permission List...
						With cboNewValue
							If (oRank.lRankPermissions And RankPermissions.AcceptApplicant) <> 0 Then
								.AddItem("Accept Applicant") : .ItemData(.NewIndex) = RankPermissions.AcceptApplicant
							End If
							If (oRank.lRankPermissions And RankPermissions.AcceptEvents) <> 0 Then
								.AddItem("Accept Events") : .ItemData(.NewIndex) = RankPermissions.AcceptEvents
							End If
							If (oRank.lRankPermissions And RankPermissions.BuildGuildBase) <> 0 Then
								.AddItem("Build Guild Base") : .ItemData(.NewIndex) = RankPermissions.BuildGuildBase
							End If
							If (oRank.lRankPermissions And RankPermissions.ChangeMOTD) <> 0 Then
								.AddItem("Change MOTD") : .ItemData(.NewIndex) = RankPermissions.ChangeMOTD
							End If
							If (oRank.lRankPermissions And RankPermissions.ChangeRankNames) <> 0 Then
								.AddItem("Change Rank Names") : .ItemData(.NewIndex) = RankPermissions.ChangeRankNames
							End If
							If (oRank.lRankPermissions And RankPermissions.ChangeRankPermissions) <> 0 Then
								.AddItem("Change Rank Permissions") : .ItemData(.NewIndex) = RankPermissions.ChangeRankPermissions
							End If
							If (oRank.lRankPermissions And RankPermissions.ChangeRankVotingWeight) <> 0 Then
								.AddItem("Change Rank Voting Weight") : .ItemData(.NewIndex) = RankPermissions.ChangeRankVotingWeight
							End If
							If (oRank.lRankPermissions And RankPermissions.ChangeRecruitment) <> 0 Then
								.AddItem("Change Recruitment Settings") : .ItemData(.NewIndex) = RankPermissions.ChangeRecruitment
							End If
							If (oRank.lRankPermissions And RankPermissions.CreateEvents) <> 0 Then
								.AddItem("Create Events") : .ItemData(.NewIndex) = RankPermissions.CreateEvents
							End If
							If (oRank.lRankPermissions And RankPermissions.CreateRanks) <> 0 Then
								.AddItem("Create/Move Ranks") : .ItemData(.NewIndex) = RankPermissions.CreateRanks
							End If
							If (oRank.lRankPermissions And RankPermissions.DeleteEvents) <> 0 Then
								.AddItem("Delete Events") : .ItemData(.NewIndex) = RankPermissions.DeleteEvents
							End If
							If (oRank.lRankPermissions And RankPermissions.DeleteRanks) <> 0 Then
								.AddItem("Delete Ranks") : .ItemData(.NewIndex) = RankPermissions.DeleteRanks
							End If
							If (oRank.lRankPermissions And RankPermissions.DemoteMember) <> 0 Then
								.AddItem("Demote Member Rank") : .ItemData(.NewIndex) = RankPermissions.DemoteMember
							End If
                            'If (oRank.lRankPermissions And RankPermissions.GuildChatChannel) <> 0 Then
                            '	.AddItem("Guild Chat Channel Access") : .ItemData(.NewIndex) = RankPermissions.GuildChatChannel
                            'End If
							If (oRank.lRankPermissions And RankPermissions.PromoteMember) <> 0 Then
								.AddItem("Promote Member Rank") : .ItemData(.NewIndex) = RankPermissions.PromoteMember
							End If
							If (oRank.lRankPermissions And RankPermissions.ProposeVotes) <> 0 Then
								.AddItem("Propose New Votes") : .ItemData(.NewIndex) = RankPermissions.ProposeVotes
							End If
							If (oRank.lRankPermissions And RankPermissions.RejectMember) <> 0 Then
								.AddItem("Reject Applicants") : .ItemData(.NewIndex) = RankPermissions.RejectMember
							End If
							If (oRank.lRankPermissions And RankPermissions.RemoveMember) <> 0 Then
								.AddItem("Remove Member") : .ItemData(.NewIndex) = RankPermissions.RemoveMember
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewBankLog) <> 0 Then
								.AddItem("View Bank Log") : .ItemData(.NewIndex) = RankPermissions.ViewBankLog
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewEvents) <> 0 Then
								.AddItem("View Events") : .ItemData(.NewIndex) = RankPermissions.ViewEvents
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewEventAttachments) <> 0 Then
								.AddItem("View Event Attachments") : .ItemData(.NewIndex) = RankPermissions.ViewEventAttachments
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewContentsLowSec) <> 0 Then
								.AddItem("View Low Security Bank Contents") : .ItemData(.NewIndex) = RankPermissions.ViewContentsLowSec
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewContentsHiSec) <> 0 Then
								.AddItem("View High Security Bank Contents") : .ItemData(.NewIndex) = RankPermissions.ViewContentsHiSec
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewGuildBase) <> 0 Then
								.AddItem("View Bank Facility") : .ItemData(.NewIndex) = RankPermissions.ViewGuildBase
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewVotesHistory) <> 0 Then
								.AddItem("View Vote History") : .ItemData(.NewIndex) = RankPermissions.ViewVotesHistory
							End If
							If (oRank.lRankPermissions And RankPermissions.ViewVotesInProgress) <> 0 Then
								.AddItem("View Votes In Progress") : .ItemData(.NewIndex) = RankPermissions.ViewVotesInProgress
							End If
							If (oRank.lRankPermissions And RankPermissions.WithdrawLowSec) <> 0 Then
								.AddItem("Withdraw Low Security Bank") : .ItemData(.NewIndex) = RankPermissions.WithdrawLowSec
							End If
							If (oRank.lRankPermissions And RankPermissions.WithdrawHiSec) <> 0 Then
								.AddItem("Withdraw High Security Bank") : .ItemData(.NewIndex) = RankPermissions.WithdrawHiSec
							End If
						End With

					Case eyTypeOfVote.Charter
						Select Case lItem
							Case -1
								txtSummary.Caption = "The entire group is disbanded and ceases to exist. This action is irreversable."
								cboNewValue.Visible = False : lblNewValue.Visible = False
							Case elGuildFlags.AutomaticTradeAgreements
								txtSummary.Caption = "Indicates whether trade routes are automatically opened between members. This increases cashflow at the cost of security."
								cboNewValue.AddItem("Yes") : cboNewValue.ItemData(cboNewValue.NewIndex) = 1
								cboNewValue.AddItem("No") : cboNewValue.ItemData(cboNewValue.NewIndex) = 0
							Case elGuildFlags.RequirePeaceBetweenMembers
								txtSummary.Caption = "Indicates whether all members are assumed to be at peace with other members."
								cboNewValue.AddItem("Yes") : cboNewValue.ItemData(cboNewValue.NewIndex) = 1
								cboNewValue.AddItem("No") : cboNewValue.ItemData(cboNewValue.NewIndex) = 0
                                'Case elGuildFlags.UnifiedForeignPolicy
                                '	txtSummary.Caption = "Indicates whether the all members share a relationship pool. If yes, members cannot adjust their personal diplomacy settings with other players."
                                '	cboNewValue.AddItem("Yes") : cboNewValue.ItemData(cboNewValue.NewIndex) = 1
                                '	cboNewValue.AddItem("No") : cboNewValue.ItemData(cboNewValue.NewIndex) = 0
							Case elGuildFlags.ShareUnitVision
								txtSummary.Caption = "Indicates whether members automatically share visibility with other members."
								cboNewValue.AddItem("Yes") : cboNewValue.ItemData(cboNewValue.NewIndex) = 1
								cboNewValue.AddItem("No") : cboNewValue.ItemData(cboNewValue.NewIndex) = 0
							Case -2		'interval list
								txtSummary.Caption = "Dictates when taxes are due."
								With cboNewValue
									.AddItem("Annually") : .ItemData(.NewIndex) = eyGuildInterval.Annually
									.AddItem("Bi-Monthly") : .ItemData(.NewIndex) = eyGuildInterval.EveryTwoMonths
									.AddItem("Daily") : .ItemData(.NewIndex) = eyGuildInterval.Daily
									.AddItem("Monthly") : .ItemData(.NewIndex) = eyGuildInterval.Monthly
									.AddItem("Quarterly") : .ItemData(.NewIndex) = eyGuildInterval.Quarterly
									.AddItem("Semi-Annually") : .ItemData(.NewIndex) = eyGuildInterval.SemiAnnually
									.AddItem("Semi-Monthly") : .ItemData(.NewIndex) = eyGuildInterval.SemiMonthly
									.AddItem("Weekly") : .ItemData(.NewIndex) = eyGuildInterval.Weekly
								End With
							Case -3		'vote weight type list
								txtSummary.Caption = "Dictates how voting weight is determined. Seniority means how long the member was with the guild in days. Population is on a per populace vote. Rank means it is determined through rank vote strengths. Static Value means that every vote counts as 1 vote."
								With cboNewValue
									.AddItem("Membership Seniority") : .ItemData(.NewIndex) = eyVoteWeightType.AgeOfPlayer
									.AddItem("Population") : .ItemData(.NewIndex) = eyVoteWeightType.PopulationBased
									.AddItem("Rank") : .ItemData(.NewIndex) = eyVoteWeightType.RankBased
									.AddItem("Static Value") : .ItemData(.NewIndex) = eyVoteWeightType.StaticValue
								End With
							Case Else
								txtSummary.Caption = "Determines how the action indicated is permitted."
								cboNewValue.AddItem("Rank-Based") : cboNewValue.ItemData(cboNewValue.NewIndex) = 1
								cboNewValue.AddItem("By Vote Only") : cboNewValue.ItemData(cboNewValue.NewIndex) = 0
						End Select
					Case eyTypeOfVote.FreeFormVote
						cboNewValue.Visible = False
                        lblNewValue.Visible = False
                    Case eyTypeOfVote.Membership

                        Select Case lItem
                            Case 1      'accept member
                                txtSummary.Caption = "Accept an applicant's application for membership."
                                With goCurrentPlayer.oGuild
                                    If .moMembers Is Nothing = False Then
                                        For X As Int32 = 0 To .moMembers.GetUpperBound(0)
                                            If .moMembers(X) Is Nothing = False Then
                                                If .moMembers(X).yMemberState = GuildMemberState.Applied Then
                                                    cboNewValue.AddItem(.moMembers(X).sPlayerName)
                                                    cboNewValue.ItemData(cboNewValue.ListIndex) = .moMembers(X).lMemberID
                                                End If
                                            End If
                                        Next X
                                    End If
                                End With
                            Case 5      'create rank
                                txtSummary.Caption = "Create a new rank with the specified name."
                                cboNewValue.Visible = False
                                txtNewValue.Visible = True
                                txtNewValue.Caption = " "
                                txtNewValue.Caption = ""
                            Case 6      'delete rank
                                txtSummary.Caption = "Delete a rank. All members belonging to the rank are automatically assigned to the next lowest rank possible."
                                With goCurrentPlayer.oGuild
                                    If .moRanks Is Nothing = False Then
                                        For X As Int32 = 0 To .moRanks.GetUpperBound(0)
                                            If .moRanks(X) Is Nothing = False Then
                                                cboNewValue.AddItem(.moRanks(X).sRankName)
                                                cboNewValue.ItemData(cboNewValue.NewIndex) = .moRanks(X).lRankID
                                            End If
                                        Next X
                                    End If
                                End With
                            Case 7      'guild recruiting
                                txtSummary.Caption = "Affects whether active recruitment is done."
                                cboNewValue.AddItem("On") : cboNewValue.ItemData(cboNewValue.NewIndex) = 1
                                cboNewValue.AddItem("Off") : cboNewValue.ItemData(cboNewValue.NewIndex) = 0
                            Case Else
                                'demote member
                                'promote member
                                'remove member
                                txtSummary.Caption = "Affects the selected member."
                                With goCurrentPlayer.oGuild
                                    If .moMembers Is Nothing = False Then
                                        For X As Int32 = 0 To .moMembers.GetUpperBound(0)
                                            If .moMembers(X) Is Nothing = False Then
                                                If (.moMembers(X).yMemberState And GuildMemberState.Approved) <> 0 Then
                                                    cboNewValue.AddItem(.moMembers(X).sPlayerName)
                                                    cboNewValue.ItemData(cboNewValue.NewIndex) = .moMembers(X).lMemberID
                                                End If
                                            End If
                                        Next X
                                    End If
                                End With
                        End Select

					Case eyTypeOfVote.VotingWeights
						cboNewValue.Visible = False
						txtNewValue.Visible = True
						txtNewValue.Caption = " "
						txtNewValue.Caption = ""
						txtNewValue.bNumericOnly = True
						Dim sRankName As String = goCurrentPlayer.oGuild.GetRankName(lItem)
						txtSummary.Caption = "Change the voting weight of " & sRankName & " to the new value specified."
				End Select
			End If

		End Sub

		Private Sub cboTypeOfVote_ItemSelected(ByVal lItemIndex As Integer) Handles cboTypeOfVote.ItemSelected
			If lItemIndex > -1 Then
				Dim lType As Int32 = cboTypeOfVote.ItemData(lItemIndex)

				cboSelectedItem.Visible = True
				lblSelectedItem.Visible = True
				cboNewValue.Visible = True
				lblNewValue.Visible = True

				Select Case lType
					Case eyTypeOfVote.AddRankPermissions, eyTypeOfVote.RemoveRankPermission, eyTypeOfVote.VotingWeights
						cboSelectedItem.Clear()
						If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
							With goCurrentPlayer.oGuild
								If .moRanks Is Nothing = False Then
									For X As Int32 = 0 To .moRanks.GetUpperBound(0)
										If .moRanks(X) Is Nothing = False Then
											cboSelectedItem.AddItem(.moRanks(X).sRankName)
											cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = .moRanks(X).lRankID
										End If
									Next X
								End If
							End With
						End If
					Case eyTypeOfVote.Charter
						cboSelectedItem.Clear()
						cboSelectedItem.AddItem("Accept Member Rights") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = elGuildFlags.AcceptMemberToGuild_RA
						cboSelectedItem.AddItem("Automatic Trade Agreements") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = elGuildFlags.AutomaticTradeAgreements
						cboSelectedItem.AddItem("Change Vote Weight Rights") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = elGuildFlags.ChangeVotingWeight_RA
						cboSelectedItem.AddItem("Create Rank Rights") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = elGuildFlags.CreateRank_RA
						cboSelectedItem.AddItem("Delete Rank Rights") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = elGuildFlags.DeleteRank_RA
						cboSelectedItem.AddItem("Demote Member Rights") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = elGuildFlags.DemoteMember_RA
						cboSelectedItem.AddItem("Disband") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = -1
						cboSelectedItem.AddItem("Share Unit/Facility Visibility") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = elGuildFlags.ShareUnitVision
						cboSelectedItem.AddItem("Require Member Peace") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = elGuildFlags.RequirePeaceBetweenMembers
						cboSelectedItem.AddItem("Promote Member Rights") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = elGuildFlags.PromoteGuildMember_RA
						cboSelectedItem.AddItem("Remove Member Rights") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = elGuildFlags.RemoveGuildMember_RA
						cboSelectedItem.AddItem("Tax Rate Interval") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = -2
                        'cboSelectedItem.AddItem("Unified Foreign Policy") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = elGuildFlags.UnifiedForeignPolicy
						cboSelectedItem.AddItem("Vote Weight Type") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = -3
					Case eyTypeOfVote.Membership
						cboSelectedItem.Clear()
						cboSelectedItem.AddItem("Accept Member") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = 1
						cboSelectedItem.AddItem("Demote Member") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = 2
						cboSelectedItem.AddItem("Promote Member") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = 3
						cboSelectedItem.AddItem("Remove Member") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = 4
						cboSelectedItem.AddItem("Create Rank") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = 5
						cboSelectedItem.AddItem("Delete Rank") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = 6
                        cboSelectedItem.AddItem("Actively Recruiting") : cboSelectedItem.ItemData(cboSelectedItem.NewIndex) = 7
                    Case Else
                        'freeform
                        cboSelectedItem.Clear()
                        cboSelectedItem.Visible = False
                        lblSelectedItem.Visible = False
                        cboNewValue.Visible = False
                        lblNewValue.Visible = False
                End Select
			End If
		End Sub

		Private Sub lstProposals_ItemClick(ByVal lIndex As Integer) Handles lstProposals.ItemClick
			If lIndex > -1 Then
				Dim lID As Int32 = lstProposals.ItemData(lIndex)

				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					Dim oVote As GuildVote = goCurrentPlayer.oGuild.GetVote(lID)
					SetInternalAffairsFromVote(oVote)
				End If
			End If
		End Sub

		Private Sub optVoteAbstain_Click() Handles optVoteAbstain.Click
			optVoteYes.Value = False
			optVoteNo.Value = False
			optVoteAbstain.Value = True
			SendVoteToServer()
		End Sub

		Private Sub optVoteNo_Click() Handles optVoteNo.Click
			optVoteYes.Value = False
			optVoteNo.Value = True
			optVoteAbstain.Value = False
			SendVoteToServer()
		End Sub

		Private Sub optVoteYes_Click() Handles optVoteYes.Click
			optVoteYes.Value = True
			optVoteNo.Value = False
			optVoteAbstain.Value = False
			SendVoteToServer()
		End Sub

		Private Sub SendVoteToServer()
			'Send Msg to update the player's vote
			If lstProposals.ListIndex < 0 Then Return

			Dim lProposalID As Int32 = lstProposals.ItemData(lstProposals.ListIndex)

			Dim oVote As GuildVote = goCurrentPlayer.oGuild.GetVote(lProposalID)
			If oVote Is Nothing Then Return
			If oVote.yVoteState <> eyVoteState.eInProgress Then
				MyBase.moUILib.AddNotification("Unable to change vote as the voting process is over.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If

			Dim yVoteVal As Byte = 0		'0 = abstain
			If optVoteAbstain.Value = True Then
				yVoteVal = eyVoteValue.AbstainVote
			ElseIf optVoteNo.Value = True Then
				yVoteVal = eyVoteValue.NoVote
			Else : yVoteVal = eyVoteValue.YesVote
			End If

			Dim yMsg(7) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerVote).CopyTo(yMsg, lPos) : lPos += 2
			yMsg(lPos) = 0 : lPos += 1		 '0 indicates guild
			System.BitConverter.GetBytes(lProposalID).CopyTo(yMsg, lPos) : lPos += 4
			yMsg(lPos) = yVoteVal : lPos += 1
			MyBase.moUILib.SendMsgToPrimary(yMsg)
			MyBase.moUILib.AddNotification("Vote has been submitted.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		End Sub

		Private Sub SetInternalAffairsFromVote(ByRef oVote As GuildVote)
			If oVote Is Nothing = False Then
				txtProposedBy.Caption = GetCacheObjectValue(oVote.ProposedByID, ObjectType.ePlayer)
				txtVoteStarts.Locked = True : txtVoteStarts.Caption = oVote.dtVoteStarts.ToLocalTime.ToShortDateString & " " & oVote.dtVoteStarts.ToLocalTime.ToShortTimeString
				txtVoteEnds.Locked = True : txtVoteEnds.Caption = oVote.dtVoteEnds.ToLocalTime.ToShortDateString & " " & oVote.dtVoteEnds.ToLocalTime.ToShortTimeString
				cboTypeOfVote.Enabled = glPlayerID = oVote.ProposedByID : cboTypeOfVote.FindComboItemData(oVote.yTypeOfVote)
				cboSelectedItem.Enabled = glPlayerID = oVote.ProposedByID : cboSelectedItem.FindComboItemData(oVote.lSelectedItem)
				cboNewValue.Enabled = glPlayerID = oVote.ProposedByID : cboNewValue.FindComboItemData(oVote.lNewValue)
				txtSummary.Locked = glPlayerID <> oVote.ProposedByID : txtSummary.Caption = oVote.sSummary
				optVoteYes.Value = oVote.yPlayerVote = eyVoteValue.YesVote
				optVoteNo.Value = oVote.yPlayerVote = eyVoteValue.NoVote
				optVoteAbstain.Value = oVote.yPlayerVote = eyVoteValue.AbstainVote

				lblVotingEnds.Caption = "Voting Ends On:"
				btnPropose.Visible = False

				Dim yMsg(6) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eGetMyVoteValue).CopyTo(yMsg, 0)
				yMsg(2) = 0	'for guild
				System.BitConverter.GetBytes(oVote.VoteID).CopyTo(yMsg, 3)
				MyBase.moUILib.SendMsgToPrimary(yMsg)
			Else
				lstProposals.ListIndex = -1
				btnPropose.Visible = True

				If goCurrentPlayer Is Nothing = False Then txtProposedBy.Caption = goCurrentPlayer.PlayerName
				txtVoteStarts.Locked = True : txtVoteStarts.Caption = "Upon Proposal"
				txtVoteEnds.Locked = False : txtVoteEnds.Caption = "24"
				cboTypeOfVote.Enabled = True
				cboSelectedItem.Enabled = True
				cboNewValue.Enabled = True
				txtSummary.Locked = False : txtSummary.Caption = "Enter Summary..."
				optVoteYes.Value = False
				optVoteNo.Value = False
				optVoteAbstain.Value = False

				lblVotingEnds.Caption = "Duration (hours):"
			End If
		End Sub

		Private Sub FillProposalsList()
			lstProposals.Clear()

			If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False AndAlso goCurrentPlayer.oGuild.moVotes Is Nothing = False Then

				With goCurrentPlayer.oGuild

					If ml_IntAffSelectedTab = 0 Then
						If .HasPermission(RankPermissions.ViewVotesInProgress) = False Then Return
					ElseIf .HasPermission(RankPermissions.ViewVotesHistory) = False Then
						Return
					End If

					For X As Int32 = 0 To .moVotes.GetUpperBound(0)
						Dim oVote As GuildVote = .moVotes(X)
						If oVote Is Nothing = False Then
							If ml_IntAffSelectedTab = 0 Then
								'in progress
								If oVote.dtVoteEnds.ToLocalTime > Now AndAlso oVote.yVoteState = eyVoteState.eInProgress Then
									lstProposals.AddItem(oVote.ListBoxText, False)
									lstProposals.ItemData(lstProposals.NewIndex) = oVote.VoteID
								End If
							Else
								'history
								If oVote.dtVoteEnds.ToLocalTime < Now AndAlso oVote.yVoteState <> eyVoteState.eInProgress Then
									lstProposals.AddItem(oVote.ListBoxText, False)
									lstProposals.ItemData(lstProposals.NewIndex) = oVote.VoteID
								End If
							End If
						End If
					Next X
				End With
			End If
		End Sub

		Public Overrides Sub NewFrame()
			If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
				With goCurrentPlayer.oGuild
					Dim sText As String = .GetCharterText
					If txtCharter.Caption <> sText Then txtCharter.Caption = sText

					'check our proposals list
					Dim lCnt As Int32 = 0
					If .moVotes Is Nothing = False Then
						For X As Int32 = 0 To .moVotes.GetUpperBound(0)
							If .moVotes(X) Is Nothing = False Then
								If ml_IntAffSelectedTab = 0 Then
									'in progress
									If .moVotes(X).dtVoteEnds.ToLocalTime > Now Then
										lCnt += 1
									End If
								Else
									'historical
									If .moVotes(X).dtVoteEnds.ToLocalTime < Now Then
										lCnt += 1
									End If
								End If
							End If
						Next X
					End If
					If lstProposals.ListCount <> lCnt Then
						'ok, fill our proposal list
						FillProposalsList()
					End If

					If lstProposals.ListIndex > -1 Then
						Dim lID As Int32 = lstProposals.ItemData(lstProposals.ListIndex)
						Dim oProposal As GuildVote = .GetVote(lID)
						If oProposal Is Nothing = False Then
							sText = GetCacheObjectValue(oProposal.ProposedByID, ObjectType.ePlayer)
							If txtProposedBy.Caption <> sText Then txtProposedBy.Caption = sText
							sText = oProposal.dtVoteStarts.ToLocalTime.ToShortDateString & " " & oProposal.dtVoteStarts.ToLocalTime.ToShortTimeString
							If txtVoteStarts.Caption <> sText Then txtVoteStarts.Caption = sText

							If oProposal.VoteID <> -1 Then
								sText = oProposal.dtVoteEnds.ToLocalTime.ToShortDateString & " " & oProposal.dtVoteEnds.ToLocalTime.ToShortTimeString
								If txtVoteEnds.Caption <> sText Then txtVoteEnds.Caption = sText
							End If

							sText = oProposal.sSummary
							If oProposal.yVoteState = eyVoteState.eVoteFailed Then
								sText &= " - VOTE DID NOT PASS"
							ElseIf oProposal.yVoteState = eyVoteState.eVotePassed Then
								sText &= " - VOTE PASSED"
							End If
							If txtSummary.Caption <> sText Then txtSummary.Caption = sText

							Dim bVal As Boolean
							bVal = oProposal.yPlayerVote = eyVoteValue.YesVote
							If optVoteYes.Value <> bVal Then optVoteYes.Value = bVal
							bVal = oProposal.yPlayerVote = eyVoteValue.NoVote
							If optVoteNo.Value <> bVal Then optVoteNo.Value = bVal
							bVal = oProposal.yPlayerVote = eyVoteValue.AbstainVote
							If optVoteAbstain.Value <> bVal Then optVoteAbstain.Value = bVal

							If cboTypeOfVote.ListIndex > -1 Then
								If oProposal.yTypeOfVote <> cboTypeOfVote.ItemData(cboTypeOfVote.ListIndex) Then cboTypeOfVote.FindComboItemData(oProposal.yTypeOfVote)
							End If
							If cboSelectedItem.ListIndex > -1 Then
								If oProposal.lSelectedItem <> cboSelectedItem.ItemData(cboSelectedItem.ListIndex) Then cboSelectedItem.FindComboItemData(oProposal.lSelectedItem)
							End If
							If cboNewValue.ListIndex > -1 Then
								If oProposal.lNewValue <> cboNewValue.ItemData(cboNewValue.ListIndex) Then cboNewValue.FindComboItemData(oProposal.lNewValue)
							End If



						End If
					End If
				End With
			End If
		End Sub

		Public Overrides Sub RenderEnd()

		End Sub

		Private Sub fraProposals_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles fraProposals.OnMouseDown

			Dim pt As Point = Me.GetAbsolutePosition()
			lMouseX -= pt.X - Me.Left
			lMouseY -= pt.Y - Me.Top

            If mrc_INT_AFF_HISTORY.Contains(lMouseX, lMouseY) = True Then
                If ml_IntAffSelectedTab = 0 Then
                    Dim yMsg(5) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eRequestGuildVoteProposals).CopyTo(yMsg, 0)
                    System.BitConverter.GetBytes(1I).CopyTo(yMsg, 2)
                    MyBase.moUILib.SendMsgToPrimary(yMsg)
                End If
                ml_IntAffSelectedTab = 1
                Me.IsDirty = True
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            ElseIf mrc_INT_AFF_INPROGRESS.Contains(lMouseX, lMouseY) = True Then
                If ml_IntAffSelectedTab = 1 Then
                    Dim yMsg(5) As Byte
                    System.BitConverter.GetBytes(GlobalMessageCode.eRequestGuildVoteProposals).CopyTo(yMsg, 0)
                    System.BitConverter.GetBytes(0I).CopyTo(yMsg, 2)
                    MyBase.moUILib.SendMsgToPrimary(yMsg)
                End If
                ml_IntAffSelectedTab = 0
                Me.IsDirty = True
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            End If

		End Sub

		Private Sub btnPropose_Click(ByVal sName As String) Handles btnPropose.Click
			'submit the new proposal... double-check the proposal...
			If ValidateNewProposal() = False Then Return

			Dim yTypeOfVote As Byte = eyTypeOfVote.FreeFormVote
			If cboTypeOfVote.ListIndex > -1 Then
				yTypeOfVote = CByte(cboTypeOfVote.ItemData(cboTypeOfVote.ListIndex))
			End If
			Dim lSelectedItem As Int32 = -1
			If cboSelectedItem.ListIndex > -1 Then
				lSelectedItem = cboSelectedItem.ItemData(cboSelectedItem.ListIndex)
			End If
			Dim lNewValue As Int32 = 0
			Dim sNewValueText As String = ""
			If txtNewValue.Visible = True Then
				sNewValueText = txtNewValue.Caption
				If sNewValueText.Length > 20 Then sNewValueText = sNewValueText.Substring(0, 20)
			Else
				If cboNewValue.ListIndex > -1 Then
					lNewValue = cboNewValue.ItemData(cboNewValue.ListIndex)
				End If
			End If
            If yTypeOfVote = eyTypeOfVote.VotingWeights Then
                If IsNumeric(sNewValueText) = False OrElse sNewValueText.Length > 9 OrElse CInt(sNewValueText) < -1 Then
                    MyBase.moUILib.AddNotification("You must enter a positive, numeric value for new value.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If
            End If
			Dim sSummary As String = txtSummary.Caption
			If sSummary Is Nothing Then sSummary = ""
			If sSummary.Length > 255 Then sSummary = sSummary.Substring(0, 255)

			Dim lDuration As Int32 = CInt(Val(txtVoteEnds.Caption))


			Dim yMsg(289) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eProposeGuildVote).CopyTo(yMsg, lPos) : lPos += 2
			yMsg(lPos) = yTypeOfVote : lPos += 1
			System.BitConverter.GetBytes(lSelectedItem).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(lNewValue).CopyTo(yMsg, lPos) : lPos += 4
			System.Text.ASCIIEncoding.ASCII.GetBytes(sNewValueText).CopyTo(yMsg, lPos) : lPos += 20
			System.BitConverter.GetBytes(lDuration).CopyTo(yMsg, lPos) : lPos += 4
			System.Text.ASCIIEncoding.ASCII.GetBytes(sSummary).CopyTo(yMsg, lPos) : lPos += 255

			MyBase.moUILib.SendMsgToPrimary(yMsg)
			MyBase.moUILib.AddNotification("Proposal Submitted...", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

		End Sub

		Private Function ValidateNewProposal() As Boolean
			If cboTypeOfVote.ListIndex > -1 Then
				If cboTypeOfVote.ItemData(cboTypeOfVote.ListIndex) <> eyTypeOfVote.FreeFormVote Then
					If cboSelectedItem.ListIndex > -1 Then
						If cboNewValue.ListIndex > -1 OrElse (txtNewValue.Visible = True AndAlso txtNewValue.Caption.Trim <> "") Then
							Dim lDuration As Int32 = CInt(Val(txtVoteEnds.Caption))
							If lDuration < 8 Then
								MyBase.moUILib.AddNotification("Votes must last at least 8 hours.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
								Return False
							End If
							If txtSummary.Caption.Length < 3 Then
								MyBase.moUILib.AddNotification("Please enter a valid Summary for this proposal.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
								Return False
							End If
						Else
							MyBase.moUILib.AddNotification("Select a value to set this item to.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
							Return False
						End If
					Else
						MyBase.moUILib.AddNotification("Select an item to vote on.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
						Return False
					End If
				End If
			Else
				MyBase.moUILib.AddNotification("Select a Type of Vote.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return False
			End If
			Return True
		End Function

		Private Sub fra_InternalAffairs_OnNewFrame() Handles Me.OnNewFrame
			If lstProposals Is Nothing = False AndAlso goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
				If lstProposals.ListIndex > -1 Then
					Dim lVoteID As Int32 = lstProposals.ItemData(lstProposals.ListIndex)
					Dim oVote As GuildVote = goCurrentPlayer.oGuild.GetVote(lVoteID)
					If oVote Is Nothing = False Then
						Dim bOptYes As Boolean = oVote.yPlayerVote = eyVoteValue.YesVote
						Dim bOptNo As Boolean = oVote.yPlayerVote = eyVoteValue.NoVote
						Dim bOptAbs As Boolean = oVote.yPlayerVote = eyVoteValue.AbstainVote

						If optVoteAbstain.Value <> bOptAbs Then optVoteAbstain.Value = bOptAbs
						If optVoteYes.Value <> bOptYes Then optVoteYes.Value = bOptYes
						If optVoteNo.Value <> bOptNo Then optVoteNo.Value = bOptNo

						If txtSummary.Caption <> oVote.sSummary Then txtSummary.Caption = oVote.sSummary
					End If
				End If
			End If
		End Sub
	End Class

End Class
