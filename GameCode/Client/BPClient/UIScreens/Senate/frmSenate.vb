Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmSenate
	Inherits UIWindow

	Private rcFloor As Rectangle = New Rectangle(3, 28, 150, 24)			'Senate Floor
    Private rcEmperor As Rectangle = New Rectangle(153, 28, 150, 24)        'Emperor's Chamber 
    Private rcHistory As Rectangle = New Rectangle(303, 28, 150, 24)        'Historical

	Private lblWindowTitle As UILabel
	Private lnDiv1 As UILine
	Private lblProposed As UILabel
	Private lnDiv2 As UILine
    Private WithEvents btnProposeNew As UIButton
    Private WithEvents btnDeleteProposal As UIButton
	Private WithEvents btnClose As UIButton
    Private WithEvents lstProposals As UIListBox
    Private WithEvents btnEmpMsgBrd As UIButton

	'==== fraDetails
	Private fraDetails As UIWindow
	Private lblProposedBy As UILabel
	Private lblProposedOn As UILabel
	Private lblTitle As UILabel
    Private WithEvents txtTitle As UITextBox
	Private lblDesc As UILabel
    Private WithEvents txtDescription As UITextBox
	Private lblPropVotes As UILabel
	Private lstVotes As UIListBox
	Private tvwVotes As UITreeView
	Private txtFind As UITextBox
	Private txtMessage As UITextBox
	Private lblMessages As UILabel
	Private lblYourVote As UILabel
    Private WithEvents lstMessages As UIListBox

	Private WithEvents btnFindVote As UIButton
	Private WithEvents btnAddMessage As UIButton
	Private WithEvents btnVoteYes As UIButton
	Private WithEvents btnVoteNo As UIButton

    Private WithEvents lblPriority As UILabel
    Private WithEvents opt1 As UIOption
    Private WithEvents opt2 As UIOption
    Private WithEvents opt3 As UIOption
    Private WithEvents opt4 As UIOption
    Private WithEvents opt5 As UIOption

    Private moProposals(-1) As SenateProposal

    Private mbEditting As Boolean = False

	Private myViewState As Byte = 0 'senatefloor = 0, emperor chamber = 1
	Private Property ViewState() As Byte
		Get
			Return myViewState
		End Get
		Set(ByVal value As Byte)
            If value = 1 Then
                If goCurrentPlayer Is Nothing = False Then
                    If (goCurrentPlayer.yPlayerTitle <> Player.PlayerRank.Emperor OrElse gbAliased = True) AndAlso glPlayerID <> 1 AndAlso glPlayerID <> 6 Then
                        MyBase.moUILib.AddNotification("You do not have authorization to enter the Emperor's Chamber.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                        Return
                    End If
                End If
            End If
			If myViewState <> value Then

				myViewState = value

				If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
				Me.IsDirty = True

                lstProposals.sHeaderRow = "Title".PadRight(55, " "c) & "Votes".PadRight(15, " "c) & "Statements".PadRight(15, " "c) & "Ends"

				If myViewState = 0 Then
                    'senate floor
                    btnEmpMsgBrd.Visible = False
					btnProposeNew.Visible = False
					txtTitle.Locked = True
					txtDescription.Locked = True
					tvwVotes.Visible = True
					tvwVotes.Enabled = True
					lstVotes.Visible = False
					lstVotes.Enabled = False
                    lblPropVotes.Caption = "System Vote Results"
                    lblPriority.Visible = False
                    opt1.Visible = False
                    opt2.Visible = False
                    opt3.Visible = False
                    opt4.Visible = False
                    opt5.Visible = False
                    txtDescription.Height = (txtFind.Top + txtFind.Height) - txtDescription.Top
                    btnDeleteProposal.Visible = False
                    btnVoteNo.Caption = "Vote No"
                    btnVoteNo.Enabled = True
                    btnVoteYes.Enabled = True
                ElseIf myViewState = 1 Then
                    btnEmpMsgBrd.Visible = True
                    lstProposals.sHeaderRow = "Title".PadRight(55, " "c) & "Votes".PadRight(15, " "c) & "Statements".PadRight(15, " "c) & "Priority"
                    btnProposeNew.Visible = True
                    txtTitle.Locked = False
                    txtDescription.Locked = False
                    tvwVotes.Visible = False
                    tvwVotes.Enabled = False
                    lstVotes.Visible = True
                    lstVotes.Enabled = True
                    lblPropVotes.Caption = "Proposal Endorsements"
                    txtDescription.Height = 110
                    lblPriority.Visible = True
                    opt1.Visible = True
                    opt2.Visible = True
                    opt3.Visible = True
                    opt4.Visible = True
                    opt5.Visible = True
                    btnDeleteProposal.Visible = True
                    btnDeleteProposal.Enabled = False
                    btnVoteNo.Caption = "Veto"
                    btnVoteNo.Enabled = True
                    btnVoteYes.Enabled = True
                ElseIf myViewState = 2 Then
                    btnEmpMsgBrd.Visible = False
                    btnProposeNew.Visible = False
                    txtTitle.Locked = True
                    txtDescription.Locked = True
                    tvwVotes.Visible = True
                    tvwVotes.Enabled = True
                    lstVotes.Visible = False
                    lstVotes.Enabled = False
                    lblPropVotes.Caption = "System Vote Results"
                    lblPriority.Visible = False
                    opt1.Visible = False
                    opt2.Visible = False
                    opt3.Visible = False
                    opt4.Visible = False
                    opt5.Visible = False
                    txtDescription.Height = (txtFind.Top + txtFind.Height) - txtDescription.Top
                    btnDeleteProposal.Visible = False
                    btnVoteNo.Enabled = False
                    btnVoteYes.Enabled = False
                End If

                lstProposals.Clear()
                lstProposals.ListIndex = -1
			End If
		End Set
	End Property

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

        'TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenSenateWindow)

		'frmSenate initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eSenate
            .ControlName = "frmSenate"
            .Left = 129
            .Top = 65
            .Width = 800
            .Height = 600
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
        End With

		'lblTitle initial props
		lblWindowTitle = New UILabel(oUILib)
		With lblWindowTitle
			.ControlName = "lblTitle"
			.Left = 5
			.Top = 2
			.Width = 130
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Galactic Senate"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblWindowTitle, UIControl))

		'btnClose initial props
		btnClose = New UIButton(oUILib)
		With btnClose
			.ControlName = "btnClose"
            .Left = Me.Width - 24 - Me.BorderLineWidth
            .Top = Me.BorderLineWidth
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
            .Left = Me.BorderLineWidth \ 2
			.Top = 25
            .Width = Me.Width - (Me.BorderLineWidth)
			.Height = 0
			.Enabled = True
            .Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv1, UIControl))

		'lblProposed initial props
		lblProposed = New UILabel(oUILib)
		With lblProposed
			.ControlName = "lblProposed"
			.Left = 10
            .Top = 55
			.Width = 130
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "Proposed Legislation"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblProposed, UIControl))

		'lstProposals initial props
		lstProposals = New UIListBox(oUILib)
		With lstProposals
			.ControlName = "lstProposals"
			.Left = 10
            .Top = 75
			.Width = 780
			.Height = 195
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Courier New", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .sHeaderRow = "Title".PadRight(55, " "c) & "Votes".PadRight(15, " "c) & "Statements".PadRight(15, " "c) & "Ends"
		End With
		Me.AddChild(CType(lstProposals, UIControl))

		'lnDiv2 initial props
		lnDiv2 = New UILine(oUILib)
		With lnDiv2
			.ControlName = "lnDiv2"
            .Left = Me.BorderLineWidth \ 2
			.Top = 53
            .Width = Me.Width - (Me.BorderLineWidth)
			.Height = 0
			.Enabled = True
            .Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv2, UIControl))

		'btnProposeNew initial props
		btnProposeNew = New UIButton(oUILib)
		With btnProposeNew
			.ControlName = "btnProposeNew"
			.Left = 645
			.Top = 28
			.Width = 150
			.Height = 24
			.Enabled = True
			.Visible = False
			.Caption = "Create Proposal"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
        Me.AddChild(CType(btnProposeNew, UIControl))

        'btnDeleteProposal initial props
        btnDeleteProposal = New UIButton(oUILib)
        With btnDeleteProposal
            .ControlName = "btnDeleteProposal"
            .Left = btnProposeNew.Left - 155
            .Top = 28
            .Width = 150
            .Height = 24
            .Enabled = True
            .Visible = False
            .Caption = "Delete Proposal"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnDeleteProposal, UIControl))

        'btnEmpMsgBrd initial props
        btnEmpMsgBrd = New UIButton(oUILib)
        With btnEmpMsgBrd
            .ControlName = "btnEmpMsgBrd"
            .Left = btnProposeNew.Left
            .Top = lnDiv2.Top
            .Width = 150
            .Height = 24
            .Enabled = True
            .Visible = False
            .Caption = "Message Board"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnEmpMsgBrd, UIControl))

		'fraDetails initial props
		fraDetails = New UIWindow(oUILib)
		With fraDetails
			.ControlName = "fraDetails"
			.Left = 10
            .Top = 280
			.Width = 780
            .Height = 310
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			.BorderLineWidth = 2
			.Caption = "Proposal Details"
		End With
		Me.AddChild(CType(fraDetails, UIControl))

		'lblProposedBy initial props
		lblProposedBy = New UILabel(oUILib)
		With lblProposedBy
			.ControlName = "lblProposedBy"
			.Left = 15
			.Top = 10
			.Width = 270
			.Height = 22
			.Enabled = True
			.Visible = True
            .Caption = "Proposed By: "
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraDetails.AddChild(CType(lblProposedBy, UIControl))

		'lblProposedOn initial props
		lblProposedOn = New UILabel(oUILib)
		With lblProposedOn
			.ControlName = "lblProposedOn"
			.Left = 15
			.Top = 30
			.Width = 270
			.Height = 22
			.Enabled = True
			.Visible = True
            .Caption = "Proposed On: "
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraDetails.AddChild(CType(lblProposedOn, UIControl))

		'lblTitle initial props
		lblTitle = New UILabel(oUILib)
		With lblTitle
			.ControlName = "lblTitle"
			.Left = 15
            .Top = 50
			.Width = 29
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "Title:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraDetails.AddChild(CType(lblTitle, UIControl))

		'txtTitle initial props
		txtTitle = New UITextBox(oUILib)
		With txtTitle
			.ControlName = "txtTitle"
			.Left = 14
            .Top = 70
			.Width = 270
			.Height = 45
			.Enabled = True
			.Visible = True
			.Caption = "Title"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(0, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 255
			.MultiLine = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		fraDetails.AddChild(CType(txtTitle, UIControl))

		'lblDesc initial props
		lblDesc = New UILabel(oUILib)
		With lblDesc
			.ControlName = "lblDesc"
			.Left = 15
            .Top = 120
			.Width = 70
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "Description:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraDetails.AddChild(CType(lblDesc, UIControl))

		'txtDescription initial props
		txtDescription = New UITextBox(oUILib)
		With txtDescription
			.ControlName = "txtDescription"
			.Left = 15
            .Top = 140
			.Width = 270
            .Height = 110 '117
			.Enabled = True
			.Visible = True
			.Caption = "Description"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(0, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 1000
			.MultiLine = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		fraDetails.AddChild(CType(txtDescription, UIControl))

		'lblPropVotes initial props
		lblPropVotes = New UILabel(oUILib)
		With lblPropVotes
			.ControlName = "lblPropVotes"
			.Left = 300
			.Top = 10
			.Width = 270
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "Proposal Endorsements"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraDetails.AddChild(CType(lblPropVotes, UIControl))

		'lstVotes initial props
		lstVotes = New UIListBox(oUILib)
		With lstVotes
			.ControlName = "lstVotes"
			.Left = 300
			.Top = 30
			.Width = 200
			.Height = 210
			.Enabled = False
			.Visible = False
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
		End With
		fraDetails.AddChild(CType(lstVotes, UIControl))

		tvwVotes = New UITreeView(oUILib)
		With tvwVotes
			.ControlName = "tvwVotes"
			.mbAcceptReprocessEvents = True
			.Left = 300
			.Top = 30
			.Width = 200
			.Height = 210
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Courier New", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
		End With
		fraDetails.AddChild(CType(tvwVotes, UIControl))

		'txtFind initial props
		txtFind = New UITextBox(oUILib)
		With txtFind
			.ControlName = "txtFind"
			.Left = 300
			.Top = 245
			.Width = 145
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "Find Voter"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 20
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		fraDetails.AddChild(CType(txtFind, UIControl))

		'btnFindVote initial props
		btnFindVote = New UIButton(oUILib)
		With btnFindVote
			.ControlName = "btnFindVote"
			.Left = 450
			.Top = 245
			.Width = 54
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Find"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraDetails.AddChild(CType(btnFindVote, UIControl))

		'lblMessages initial props
		lblMessages = New UILabel(oUILib)
		With lblMessages
			.ControlName = "lblMessages"
			.Left = 510
			.Top = 10
			.Width = 270
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "Messages"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraDetails.AddChild(CType(lblMessages, UIControl))

		'lstMessages initial props
		lstMessages = New UIListBox(oUILib)
		With lstMessages
			.ControlName = "lstMessages"
			.Left = 510
			.Top = 32
			.Width = 259
			.Height = 103
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
		End With
		fraDetails.AddChild(CType(lstMessages, UIControl))

		'txtMessage initial props
		txtMessage = New UITextBox(oUILib)
		With txtMessage
			.ControlName = "txtMessage"
			.Left = 510
			.Top = 140
			.Width = 260
			.Height = 125
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(0, DrawTextFormat)
            '.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorEnabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 500
            .MultiLine = True
            .Locked = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		fraDetails.AddChild(CType(txtMessage, UIControl))

		'btnAddMessage initial props
		btnAddMessage = New UIButton(oUILib)
		With btnAddMessage
			.ControlName = "btnAddMessage"
			.Left = 590
            .Top = 280
			.Width = 110
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Add Message"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraDetails.AddChild(CType(btnAddMessage, UIControl))

		'btnVoteYes initial props
		btnVoteYes = New UIButton(oUILib)
		With btnVoteYes
			.ControlName = "btnVoteYes"
			.Left = 15
            .Top = 280
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Endorse"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraDetails.AddChild(CType(btnVoteYes, UIControl))

		'btnVoteNo initial props
		btnVoteNo = New UIButton(oUILib)
		With btnVoteNo
			.ControlName = "btnVoteNo"
			.Left = 185
            .Top = 280
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Unendorse"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraDetails.AddChild(CType(btnVoteNo, UIControl))

		'lblYourVote initial props
		lblYourVote = New UILabel(oUILib)
		With lblYourVote
			.ControlName = "lblYourVote"
			.Left = 300
            .Top = 280
			.Width = 200
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "You have voted FOR"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
		End With
		fraDetails.AddChild(CType(lblYourVote, UIControl))


        'New Control initial props
        lblPriority = New UILabel(oUILib)
        With lblPriority
            .ControlName = "lblPriority"
            .Left = 15
            .Top = txtDescription.Top + txtDescription.Height + 5
            .Width = 94
            .Height = 22
            .Enabled = True
            .Visible = True
            .Caption = "Priority:       Low"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraDetails.AddChild(CType(lblPriority, UIControl))

        'opt1 initial props
        opt1 = New UIOption(oUILib)
        With opt1
            .ControlName = "opt1"
            .Left = 115
            .Top = lblPriority.Top + 3
            .Width = 16
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
            .Value = False
        End With
        fraDetails.AddChild(CType(opt1, UIControl))

        'opt2 initial props
        opt2 = New UIOption(oUILib)
        With opt2
            .ControlName = "opt2"
            .Left = 146
            .Top = opt1.Top
            .Width = 16
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
            .Value = False
        End With
        fraDetails.AddChild(CType(opt2, UIControl))

        'opt3 initial props
        opt3 = New UIOption(oUILib)
        With opt3
            .ControlName = "opt3"
            .Left = 178
            .Top = opt1.Top
            .Width = 16
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
            .Value = False
        End With
        fraDetails.AddChild(CType(opt3, UIControl))

        'opt4 initial props
        opt4 = New UIOption(oUILib)
        With opt4
            .ControlName = "opt4"
            .Left = 209
            .Top = opt1.Top
            .Width = 16
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
            .Value = False
        End With
        fraDetails.AddChild(CType(opt4, UIControl))

        'opt5 initial props
        opt5 = New UIOption(oUILib)
        With opt5
            .ControlName = "opt5"
            .Left = 240
            .Top = opt1.Top
            .Width = 49
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "High"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        fraDetails.AddChild(CType(opt5, UIControl))

        myViewState = 5
		ViewState = 0

        '      Dim yMsg(6) As Byte
        '      Dim lPos As Int32 = 0
        '      System.BitConverter.GetBytes(GlobalMessageCode.eGetSenateObjectDetails).CopyTo(yMsg, lPos) : lPos += 2
        '      lPos += 4       'for player id
        '      yMsg(lPos) = eySenateRequestDetailsType.SenateObject : lPos += 1
        'MyBase.moUILib.SendMsgToPrimary(yMsg)
        RequestSummaryDetails()

		'SetupTest()

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)
	End Sub
 
	Private Sub frmSenate_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown

		Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
		Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y

		'Now, check for what the player is hovering over
		If rcFloor.Contains(lX, lY) = True Then
			ViewState = 0
		ElseIf rcEmperor.Contains(lX, lY) = True Then
            ViewState = 1
        ElseIf rcHistory.Contains(lX, lY) = True Then
            ViewState = 2
        End If

	End Sub

	Private Sub frmSenate_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
		MyBase.moUILib.SetToolTip(False)

		Dim lX As Int32 = lMouseX - GetAbsolutePosition.X
		Dim lY As Int32 = lMouseY - GetAbsolutePosition.Y

		'Now, check for what the player is hovering over
		If rcFloor.Contains(lX, lY) = True Then
			MyBase.moUILib.SetToolTip("Click to see the Senate Floor.", lMouseX, lMouseY)
		ElseIf rcEmperor.Contains(lX, lY) = True Then
            MyBase.moUILib.SetToolTip("Click to see the Emperor's Chamber." & vbCrLf & "Only Emperor's can enter the chamber.", lMouseX, lMouseY)
        ElseIf rcHistory.Contains(lX, lY) = True Then
            MyBase.moUILib.SetToolTip("Click to see historical senate proposals and results.", lMouseX, lMouseY)
        End If
	End Sub

	Private Sub frmSenate_OnNewFrame() Handles Me.OnNewFrame
		'validate our list and all details
		'check our proposals list
		If moProposals Is Nothing = False Then
			For X As Int32 = 0 To moProposals.GetUpperBound(0)
				If moProposals(X) Is Nothing = False Then
                    If myViewState = 0 Then
                        If lstProposals.ListIndex > -1 AndAlso lstProposals.ItemData(lstProposals.ListIndex) = moProposals(X).ObjectID Then
                            Dim sPropVotes As String = "Proposal Votes (Requires: " & moProposals(X).lRequiredVoteCnt.ToString & ")"
                            If lblPropVotes.Caption <> sPropVotes Then lblPropVotes.Caption = sPropVotes
                        End If
                        If (moProposals(X).yProposalState And eyProposalState.OnSenateFloor) <> 0 AndAlso (moProposals(X).yProposalState And eyProposalState.Archived) = 0 Then
                            Dim bFound As Boolean = False
                            Dim sText As String = moProposals(X).GetListBoxText
                            For Y As Int32 = 0 To lstProposals.ListCount - 1
                                If lstProposals.ItemData(Y) = moProposals(X).ObjectID Then
                                    bFound = True
                                    If lstProposals.List(Y) <> sText Then lstProposals.List(Y) = sText
                                    Exit For
                                End If
                            Next Y

                            If bFound = False Then
                                lstProposals.AddItem(sText, False)
                                lstProposals.ItemData(lstProposals.NewIndex) = moProposals(X).ObjectID
                            End If
                        End If
                    ElseIf myViewState = 1 AndAlso (moProposals(X).yProposalState And eyProposalState.OnSenateFloor) = 0 Then
                        If lstProposals.ListIndex > -1 AndAlso lstProposals.ItemData(lstProposals.ListIndex) = moProposals(X).ObjectID Then
                            Dim sPropVotes As String = "Endorsements (Requires: " & moProposals(X).lRequiredVoteCnt.ToString & ")"
                            If lblPropVotes.Caption <> sPropVotes Then lblPropVotes.Caption = sPropVotes
                        End If

                        Dim bFound As Boolean = False
                        Dim sText As String = moProposals(X).GetListBoxText
                        For Y As Int32 = 0 To lstProposals.ListCount - 1
                            If lstProposals.ItemData(Y) = moProposals(X).ObjectID Then
                                bFound = True
                                If lstProposals.List(Y) <> sText Then lstProposals.List(Y) = sText
                                Exit For
                            End If
                        Next Y
                        If bFound = False Then
                            lstProposals.AddItem(sText, False)
                            lstProposals.ItemData(lstProposals.NewIndex) = moProposals(X).ObjectID
                        End If
                    ElseIf myViewState = 2 AndAlso (moProposals(X).yProposalState And eyProposalState.Archived) <> 0 Then
                        If lstProposals.ListIndex > -1 AndAlso lstProposals.ItemData(lstProposals.ListIndex) = moProposals(X).ObjectID Then
                            Dim sPropVotes As String = "Proposal Votes (Requires: " & moProposals(X).lRequiredVoteCnt.ToString & ")"
                            If lblPropVotes.Caption <> sPropVotes Then lblPropVotes.Caption = sPropVotes
                        End If

                        Dim bFound As Boolean = False
                        Dim sText As String = moProposals(X).GetListBoxText
                        For Y As Int32 = 0 To lstProposals.ListCount - 1
                            If lstProposals.ItemData(Y) = moProposals(X).ObjectID Then
                                bFound = True
                                If lstProposals.List(Y) <> sText Then lstProposals.List(Y) = sText
                                Exit For
                            End If
                        Next Y
                        If bFound = False Then
                            lstProposals.AddItem(sText, False)
                            lstProposals.ItemData(lstProposals.NewIndex) = moProposals(X).ObjectID
                        End If
                    End If
                End If
            Next X
		End If

		If lstProposals Is Nothing = False Then
			If lstProposals.ListIndex > -1 Then
				Dim oProposal As SenateProposal = GetProposal(lstProposals.ItemData(lstProposals.ListIndex))
				If oProposal Is Nothing = False Then
					'Ok, check our details
                    With oProposal
                        Dim sText As String
                        If (.yProposalState And eyProposalState.OnSenateFloor) <> 0 Then
                            sText = "Voting Started: " & Date.SpecifyKind(GetDateFromNumber(.lVotingStartDate), DateTimeKind.Utc).ToLocalTime.ToString
                        Else
                            sText = "Proposed By: " & .sProposedByName

                            mbNoSendOptClick = True
                            Select Case .yPlayerPriority
                                Case 1
                                    If opt1.Value <> True Then
                                        SelectOption(1)
                                    End If
                                Case 2
                                    If opt2.Value <> True Then
                                        SelectOption(2)
                                    End If
                                Case 3
                                    If opt3.Value <> True Then
                                        SelectOption(3)
                                    End If
                                Case 4
                                    If opt4.Value <> True Then
                                        SelectOption(4)
                                    End If
                                Case 5
                                    If opt5.Value <> True Then
                                        SelectOption(5)
                                    End If
                                Case Else
                                    If (opt1.Value Or opt2.Value Or opt3.Value Or opt4.Value Or opt5.Value) = True Then SelectOption(0)
                            End Select
                            mbNoSendOptClick = False
                        End If
                        If lblProposedBy.Caption <> sText Then lblProposedBy.Caption = sText

                        If mbEditting = False Then
                            If (.yProposalState And eyProposalState.OnSenateFloor) <> 0 Then
                                sText = "Voting Ends: " & Date.SpecifyKind(GetDateFromNumber(.lVotingEndDate), DateTimeKind.Utc).ToLocalTime.ToString
                            Else
                                sText = "Proposed On: " & Date.SpecifyKind(GetDateFromNumber(.lProposedOn), DateTimeKind.Utc).ToLocalTime.ToString
                            End If
                            If lblProposedOn.Caption <> sText Then lblProposedOn.Caption = sText
                            If txtTitle.Caption <> .sTitle Then txtTitle.Caption = .sTitle


                            sText = ""
                            If .lDeliveryEstimate > 0 Then
                                Dim dtTemp As Date = Date.SpecifyKind(GetDateFromNumber(.lDeliveryEstimate), DateTimeKind.Utc).ToLocalTime
                                sText = "Should the proposal pass the senate floor, it is estimated to be delivered by: " & vbCrLf & "  " & dtTemp.ToString("MM/dd/yyyy")
                            End If
                            If (.yProposalState And eyProposalState.OnSenateFloor) <> 0 Then
                                If sText <> "" Then sText &= vbCrLf
                                If .yDefaultVote = eyVoteValue.NoVote Then
                                    sText &= "Default Vote: No (Against)"
                                Else
                                    sText &= "Default Vote: Yes (For)"
                                End If
                            End If
                            If sText <> "" Then sText &= vbCrLf & vbCrLf
                            sText &= .sDescription
                            If txtDescription.Caption <> sText Then txtDescription.Caption = sText
                        End If

                        sText = ""
                        Select Case .yPlayerVote
                            Case eyVoteValue.AbstainVote
                                sText = "You have not voted"
                            Case eyVoteValue.NoVote
                                sText = "You have voted AGAINST"
                            Case eyVoteValue.YesVote
                                sText = "You have voted FOR"
                        End Select
                        If lblYourVote.Caption <> sText Then lblYourVote.Caption = sText

                        'Now, check our votes list - lstVotes
                        If myViewState = 0 OrElse myViewState = 2 Then
                            'senate floor - by system
                            If .moSystemVotes Is Nothing = False Then
                                SmartPopulateSystemVotes(oProposal)
                            Else : tvwVotes.Clear()
                            End If
                        Else
                            'emperors chamber - by voter
                            If .muVotes Is Nothing = False Then
                                Dim lCnt As Int32 = 0
                                For X As Int32 = 0 To .muVotes.GetUpperBound(0)
                                    If .muVotes(X).yVote = eyVoteValue.YesVote Then lCnt += 1
                                Next X
                                If lCnt <> lstVotes.ListCount Then
                                    FillChamberVotes(oProposal)
                                End If
                            Else : lstVotes.Clear()
                            End If
                        End If

                        'Now, check our message list
                        If lstMessages.ListCount <> .lMsgs Then
                            FillMessageList(oProposal)
                        ElseIf Not .muMsgs Is Nothing Then
                            For X As Int32 = 0 To .lMsgs - 1
                                If lstMessages.ItemData(X) <> .muMsgs(X).lPosterID OrElse lstMessages.ItemData2(X) <> .muMsgs(X).lPostedOn Then
                                    FillMessageList(oProposal)
                                    Exit For
                                Else
                                    If lstMessages.List(X) <> .muMsgs(X).sPosterName Then lstMessages.List(X) = .muMsgs(X).sPosterName
                                End If
                            Next X
                        End If

                        If lstMessages.ListIndex > -1 Then
                            Dim lPosterID As Int32 = lstMessages.ItemData(lstMessages.ListIndex)
                            Dim lPostedOn As Int32 = lstMessages.ItemData2(lstMessages.ListIndex)
                            For X As Int32 = 0 To .lMsgs - 1
                                If .muMsgs(X).lPosterID = lPosterID AndAlso .muMsgs(X).lPostedOn = lPostedOn Then
                                    .muMsgs(X).RequestDetails(.ObjectID)
                                    If txtMessage.Caption <> .muMsgs(X).sMsgData Then txtMessage.Caption = .muMsgs(X).sMsgData
                                    If txtMessage.Locked <> True Then txtMessage.Locked = True
                                    Exit For
                                End If
                            Next X
                        End If
                    End With
				End If
			End If
		End If
	End Sub 

    Private Sub frmSenate_OnRenderEnd() Handles Me.OnRenderEnd
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return

        Dim oSelColor As System.Drawing.Color
        oSelColor = System.Drawing.Color.FromArgb(255, 128, 128, 160)

        If ViewState = 0 Then
            MyBase.moUILib.DoAlphaBlendColorFill(rcFloor, oSelColor, rcFloor.Location)
        ElseIf ViewState = 1 Then
            MyBase.moUILib.DoAlphaBlendColorFill(rcEmperor, oSelColor, rcEmperor.Location)
        ElseIf ViewState = 2 Then
            MyBase.moUILib.DoAlphaBlendColorFill(rcHistory, oSelColor, rcHistory.Location)
        End If

        'Now, draw our borders around the buttons 
        MyBase.RenderRoundedBorder(rcFloor, 1, muSettings.InterfaceBorderColor)
        Dim bRenderDisabled As Boolean = False
        If goCurrentPlayer Is Nothing = False Then
            bRenderDisabled = (goCurrentPlayer.yPlayerTitle <> Player.PlayerRank.Emperor OrElse gbAliased = True) AndAlso glPlayerID <> 1 AndAlso glPlayerID <> 6
        End If
        Dim clrEmpChmbr As System.Drawing.Color = muSettings.InterfaceBorderColor
        If bRenderDisabled = True Then
            clrEmpChmbr = System.Drawing.Color.FromArgb(clrEmpChmbr.A \ 2, clrEmpChmbr.R, clrEmpChmbr.G, clrEmpChmbr.B)
        End If
        MyBase.RenderRoundedBorder(rcHistory, 1, muSettings.InterfaceBorderColor)
        MyBase.RenderRoundedBorder(rcEmperor, 1, clrEmpChmbr)

        'Now, render our text...
        Using oFont As Font = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                oTextSpr.Begin(SpriteFlags.AlphaBlend)
                Try
                    oFont.DrawText(oTextSpr, "Senate Floor", rcFloor, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                    oFont.DrawText(oTextSpr, "Emperor's Chamber", rcEmperor, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, clrEmpChmbr)
                    oFont.DrawText(oTextSpr, "Historical", rcHistory, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                Catch
                End Try
                oTextSpr.End()
                oTextSpr.Dispose()
            End Using
            oFont.Dispose()
        End Using

    End Sub

	Private Sub btnProposeNew_Click(ByVal sName As String) Handles btnProposeNew.Click
		lstProposals.ListIndex = -1

		btnVoteNo.Caption = "Cancel"
		btnVoteYes.Caption = "Submit"

		'clear our fields
		lblProposedBy.Caption = "Proposed By: Not Yet Proposed"
		lblProposedOn.Caption = "Proposed On: Not Yet Proposed"
		txtTitle.Caption = ""
		txtTitle.Locked = False
		txtDescription.Caption = ""
		txtDescription.Locked = False

		lstVotes.Clear()
		tvwVotes.Clear()
		lstMessages.Clear()
		txtMessage.Caption = "" : txtMessage.Locked = True
        btnAddMessage.Caption = "Add Message"

        If goUILib.FocusedControl Is Nothing = False Then goUILib.FocusedControl.HasFocus = False
        goUILib.FocusedControl = txtTitle
        txtTitle.HasFocus = True
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub lstProposals_ItemClick(ByVal lIndex As Integer) Handles lstProposals.ItemClick
		If lIndex = -1 Then Return
        If myViewState = 0 OrElse myViewState = 2 Then
            btnVoteNo.Caption = "Vote No"
            btnVoteYes.Caption = "Vote Yes"
        Else
            btnVoteNo.Caption = "Veto"
            btnVoteYes.Caption = "Endorse"
        End If

        btnAddMessage.Caption = "Add Message"
        txtMessage.Caption = "" : txtMessage.Locked = True
        txtMessage.BackColorEnabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)

		Dim lProposalID As Int32 = lstProposals.ItemData(lIndex)
		Dim oProposal As SenateProposal = GetProposal(lProposalID)
		If oProposal Is Nothing = False Then
			'fill our details
			With oProposal
                .RequestDetails()
                btnDeleteProposal.Enabled = (.lProposedBy = glPlayerID) AndAlso myViewState = 1 AndAlso btnDeleteProposal.Visible = True

                lblProposedBy.Caption = "Proposed By: " & .sProposedByName
				lblProposedOn.Caption = "Proposed On: " & Date.SpecifyKind(GetDateFromNumber(.lProposedOn), DateTimeKind.Utc).ToLocalTime.ToString
				txtTitle.Caption = .sTitle
				txtTitle.Locked = True
				txtDescription.Caption = .sDescription
                txtDescription.Locked = True

                txtTitle.Locked = Not btnDeleteProposal.Enabled
                txtDescription.Locked = txtTitle.Locked

				lstVotes.Clear()
				tvwVotes.Clear()
				lstMessages.Clear()

			End With
		End If
	End Sub

	Private Sub btnAddMessage_Click(ByVal sName As String) Handles btnAddMessage.Click
		If lstProposals.ListIndex > -1 Then
			If btnAddMessage.Caption.ToUpper = "SUBMIT" Then
				'submit a msg
				btnAddMessage.Enabled = False
				If txtMessage.Caption.Trim.Length > 3 Then
					Dim lMsgLen As Int32 = txtMessage.Caption.Trim.Length
                    Dim yMsg(13 + lMsgLen) As Byte
					Dim lPos As Int32 = 0
                    System.BitConverter.GetBytes(GlobalMessageCode.eAddSenateProposalMessage).CopyTo(yMsg, lPos) : lPos += 2
                    lPos += 4       'for playerid
					System.BitConverter.GetBytes(lstProposals.ItemData(lstProposals.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
					System.BitConverter.GetBytes(lMsgLen).CopyTo(yMsg, lPos) : lPos += 4
					System.Text.ASCIIEncoding.ASCII.GetBytes(txtMessage.Caption.Trim).CopyTo(yMsg, lPos) : lPos += lMsgLen
					MyBase.moUILib.SendMsgToPrimary(yMsg)

					'Now, add our msg
					Dim oProposal As SenateProposal = GetProposal(lstProposals.ItemData(lstProposals.ListIndex))
					If oProposal Is Nothing = False Then
						ReDim Preserve oProposal.muMsgs(oProposal.lMsgs)
						With oProposal.muMsgs(oProposal.lMsgs)
							.lPostedOn = GetDateAsNumber(Now.ToUniversalTime)
							.lPosterID = glPlayerID
							.sMsgData = txtMessage.Caption.Trim
							If goCurrentPlayer Is Nothing = False Then .sPosterName = goCurrentPlayer.PlayerName Else .sPosterName = "Unknown"
						End With
						oProposal.lMsgs += 1
					End If

					MyBase.moUILib.AddNotification("Message added to proposal.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Else
					MyBase.moUILib.AddNotification("Must enter in a message at least 3 characters to post it.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				End If
                btnAddMessage.Enabled = True
                txtMessage.Locked = True
                txtMessage.BackColorEnabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
                btnAddMessage.Caption = "Add Message"
            Else
                txtMessage.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                btnAddMessage.Caption = "Submit"
                lstMessages.ListIndex = -1
                txtMessage.Caption = ""
                txtMessage.Locked = False
			End If
		End If

	End Sub

	Private Sub btnFindVote_Click(ByVal sName As String) Handles btnFindVote.Click
		If lstProposals.ListIndex > -1 Then

			Dim sSearch As String = txtFind.Caption.Trim
			If sSearch.Length < 3 Then
				MyBase.moUILib.AddNotification("You must enter at least 3 characters to search for.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If

			Dim oProposal As SenateProposal = GetProposal(lstProposals.ItemData(lstProposals.ListIndex))
			If oProposal Is Nothing = False Then
                If myViewState = 0 OrElse myViewState = 2 Then
                    'senate floor
                    Dim lCurrentNode As Int32 = -1
                    If tvwVotes.oSelectedNode Is Nothing = False Then lCurrentNode = tvwVotes.oSelectedNode.lTreeViewIndex
                    Dim oNode As UITreeView.UITreeViewItem = tvwVotes.GetNodeByItemText(sSearch, False, False, lCurrentNode)
                    If oNode Is Nothing = False Then
                        tvwVotes.oSelectedNode = oNode
                        tvwVotes.oSelectedNode.bExpanded = True
                        tvwVotes.oSelectedNode.ExpandToRoot()
                    Else
                        MyBase.moUILib.AddNotification("That item was not found, click Find again to search from the beginning of the list.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        tvwVotes.oSelectedNode = Nothing
                    End If
                Else
                    'emperors chamber
                    Dim bFound As Boolean = False
                    Dim lStartIdx As Int32 = lstVotes.ListIndex + 1
                    For X As Int32 = lStartIdx To lstVotes.ListCount - 1
                        If lstVotes.List(X).ToUpper.Contains(sSearch) = True Then
                            lstVotes.ListIndex = X
                            lstVotes.EnsureItemVisible(X)
                            bFound = True
                            Exit For
                        End If
                    Next X
                    If bFound = False Then
                        MyBase.moUILib.AddNotification("That item was not found, click Find again to search from the beginning of the list.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        lstVotes.ListIndex = -1
                    End If
                End If
			End If
		End If
	End Sub

	Private Sub btnVoteNo_Click(ByVal sName As String) Handles btnVoteNo.Click
		If myViewState = 0 Then
			If lstProposals.ListIndex > -1 Then
                Dim lProposalID As Int32 = lstProposals.ItemData(lstProposals.ListIndex)

                Dim oProposal As SenateProposal = GetProposal(lProposalID)
                If oProposal Is Nothing = False Then
                    If (oProposal.yProposalState And (eyProposalState.PassedFloor Or eyProposalState.FailedFloor)) <> 0 Then
                        MyBase.moUILib.AddNotification("Voting for that proposal has already ended.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If
                End If

                Dim yMsg(12) As Byte
				Dim lPos As Int32 = 0
				System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerVote).CopyTo(yMsg, lPos) : lPos += 2
				yMsg(lPos) = 1 : lPos += 1
				System.BitConverter.GetBytes(lProposalID).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = eyVoteValue.NoVote : lPos += 1
                lPos += 4    'for playerid
                yMsg(lPos) = 0 : lPos += 1     'priority means nothing on the senate floor
                MyBase.moUILib.SendMsgToPrimary(yMsg)

                If oProposal Is Nothing = False Then
                    oProposal.ClearRequestedDetails()
                    oProposal.RequestDetails()
                End If
            End If
		ElseIf myViewState = 1 Then
			If lstProposals.ListIndex > -1 Then
                'Removing endorsement
                If btnVoteNo.Caption = "Cancel" Then
                    mbEditting = False
                    btnVoteNo.Caption = "Veto"
                    btnVoteYes.Caption = "Endorse"
                    Return
                End If
                Dim lProposalID As Int32 = lstProposals.ItemData(lstProposals.ListIndex)

                Dim oProposal As SenateProposal = GetProposal(lProposalID)
                If oProposal Is Nothing = False Then
                    If (oProposal.yProposalState And (eyProposalState.PassedFloor Or eyProposalState.FailedFloor)) <> 0 Then
                        MyBase.moUILib.AddNotification("Voting for that proposal has already ended.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If
                End If

                Dim yPriority As Byte = 0
                If opt1.Value = True Then yPriority = 1
                If opt2.Value = True Then yPriority = 2
                If opt3.Value = True Then yPriority = 3
                If opt4.Value = True Then yPriority = 4
                If opt5.Value = True Then yPriority = 5

                Dim yMsg(12) As Byte
				Dim lPos As Int32 = 0
				System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerVote).CopyTo(yMsg, lPos) : lPos += 2
				yMsg(lPos) = 1 : lPos += 1
				System.BitConverter.GetBytes(lProposalID).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = eyVoteValue.NoVote : lPos += 1
                lPos += 4    'for playerid
                yMsg(lPos) = yPriority : lPos += 1
                MyBase.moUILib.SendMsgToPrimary(yMsg)

                If oProposal Is Nothing = False Then
                    oProposal.ClearRequestedDetails()
                    oProposal.RequestDetails()
                End If
			Else
				'Cancelling new proposal...
				lstProposals.ListIndex = 0
			End If
		End If
        RequestSummaryDetails()
	End Sub

	Private Sub btnVoteYes_Click(ByVal sName As String) Handles btnVoteYes.Click
		If myViewState = 0 Then
			If lstProposals.ListIndex > -1 Then
                Dim lProposalID As Int32 = lstProposals.ItemData(lstProposals.ListIndex)

                Dim oProposal As SenateProposal = GetProposal(lProposalID)
                If oProposal Is Nothing = False Then
                    If (oProposal.yProposalState And (eyProposalState.PassedFloor Or eyProposalState.FailedFloor)) <> 0 Then
                        MyBase.moUILib.AddNotification("Voting for that proposal has already ended.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If
                End If

                Dim yMsg(12) As Byte
				Dim lPos As Int32 = 0
				System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerVote).CopyTo(yMsg, lPos) : lPos += 2
				yMsg(lPos) = 1 : lPos += 1
				System.BitConverter.GetBytes(lProposalID).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = eyVoteValue.YesVote : lPos += 1
                lPos += 4    'for playerid
                yMsg(lPos) = 0 : lPos += 1       'priority is unused on senate floor
                MyBase.moUILib.SendMsgToPrimary(yMsg)

                If oProposal Is Nothing = False Then
                    oProposal.ClearRequestedDetails()
                    oProposal.RequestDetails()
                End If
			End If
        ElseIf myViewState = 1 Then
            Dim yPriority As Byte = 0
            If opt1.Value = True Then yPriority = 1
            If opt2.Value = True Then yPriority = 2
            If opt3.Value = True Then yPriority = 3
            If opt4.Value = True Then yPriority = 4
            If opt5.Value = True Then yPriority = 5

            If lstProposals.ListIndex > -1 Then
                Dim lProposalID As Int32 = lstProposals.ItemData(lstProposals.ListIndex)

                Dim oProposal As SenateProposal = GetProposal(lProposalID)
                If oProposal Is Nothing = False Then
                    If (oProposal.yProposalState And (eyProposalState.PassedFloor Or eyProposalState.FailedFloor)) <> 0 Then
                        MyBase.moUILib.AddNotification("Voting for that proposal has already ended.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If
                End If

                Dim yMsg(12) As Byte
                Dim lPos As Int32 = 0

                If btnVoteYes.Caption = "Save" Then
                    'Ok, save our changes to the server
                    If oProposal.lProposedBy <> glPlayerID Then
                        MyBase.moUILib.AddNotification("You do not own that proposal.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If

                    Dim sTitle As String = txtTitle.Caption.Trim
                    Dim sDesc As String = txtDescription.Caption.Trim

                    Dim lTitleLen As Int32 = System.Text.ASCIIEncoding.ASCII.GetBytes(sTitle).Length
                    Dim lDescLen As Int32 = System.Text.ASCIIEncoding.ASCII.GetBytes(sDesc).Length
                    If lTitleLen < 3 OrElse lTitleLen > 255 Then
                        MyBase.moUILib.AddNotification("The title must be at least 3 characters and no more than 255 characters.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If
                    If lDescLen < 3 OrElse lDescLen > 1000 Then
                        MyBase.moUILib.AddNotification("The description must be at least 3 characters and no more than 1000 characters.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If

                    ReDim yMsg(21 + lTitleLen + lDescLen)
                    lPos = 0

                    System.BitConverter.GetBytes(GlobalMessageCode.eAddSenateProposal).CopyTo(yMsg, lPos) : lPos += 2
                    lPos += 4       'for playerid
                    System.BitConverter.GetBytes(-2I).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(oProposal.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                    System.BitConverter.GetBytes(lTitleLen).CopyTo(yMsg, lPos) : lPos += 4
                    System.Text.ASCIIEncoding.ASCII.GetBytes(sTitle).CopyTo(yMsg, lPos) : lPos += lTitleLen
                    System.BitConverter.GetBytes(lDescLen).CopyTo(yMsg, lPos) : lPos += 4
                    System.Text.ASCIIEncoding.ASCII.GetBytes(sDesc).CopyTo(yMsg, lPos) : lPos += lDescLen
                    MyBase.moUILib.SendMsgToPrimary(yMsg)

                    If oProposal Is Nothing = False Then
                        oProposal.ClearRequestedDetails()
                        oProposal.RequestDetails()
                    End If

                    MyBase.moUILib.AddNotification("Changes submitted.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If

                System.BitConverter.GetBytes(GlobalMessageCode.eUpdatePlayerVote).CopyTo(yMsg, lPos) : lPos += 2
                yMsg(lPos) = 1 : lPos += 1
                System.BitConverter.GetBytes(lProposalID).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = eyVoteValue.YesVote : lPos += 1
                lPos += 4    'for playerid
                yMsg(lPos) = yPriority : lPos += 1
                MyBase.moUILib.SendMsgToPrimary(yMsg)

                If oProposal Is Nothing = False Then
                    oProposal.ClearRequestedDetails()
                    oProposal.RequestDetails()
                End If
            Else
                'let's double-check the proposals of the player
                If moProposals Is Nothing = False Then
                    Dim lCnt As Int32 = 0
                    For X As Int32 = 0 To moProposals.GetUpperBound(0)
                        If moProposals(X) Is Nothing = False AndAlso moProposals(X).lProposedBy = glPlayerID AndAlso (moProposals(X).yProposalState And eyProposalState.OnSenateFloor) = 0 Then
                            lCnt += 1
                        End If
                    Next X
                    If lCnt > 9 Then
                        MyBase.moUILib.AddNotification("An emperor is only allowed ten proposals.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If
                End If

                'Submitting New Proposal...
                Dim sTitle As String = txtTitle.Caption.Trim
                Dim sDesc As String = txtDescription.Caption.Trim

                Dim lTitleLen As Int32 = System.Text.ASCIIEncoding.ASCII.GetBytes(sTitle).Length
                Dim lDescLen As Int32 = System.Text.ASCIIEncoding.ASCII.GetBytes(sDesc).Length
                If lTitleLen < 3 OrElse lTitleLen > 255 Then
                    MyBase.moUILib.AddNotification("The title must be at least 3 characters and no more than 255 characters.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If
                If lDescLen < 3 OrElse lDescLen > 1000 Then
                    MyBase.moUILib.AddNotification("The description must be at least 3 characters and no more than 1000 characters.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If

                Dim yMsg(13 + lTitleLen + lDescLen) As Byte
                Dim lPos As Int32 = 0

                System.BitConverter.GetBytes(GlobalMessageCode.eAddSenateProposal).CopyTo(yMsg, lPos) : lPos += 2
                lPos += 4       'for playerid
                System.BitConverter.GetBytes(lTitleLen).CopyTo(yMsg, lPos) : lPos += 4
                System.Text.ASCIIEncoding.ASCII.GetBytes(sTitle).CopyTo(yMsg, lPos) : lPos += lTitleLen
                System.BitConverter.GetBytes(lDescLen).CopyTo(yMsg, lPos) : lPos += 4
                System.Text.ASCIIEncoding.ASCII.GetBytes(sDesc).CopyTo(yMsg, lPos) : lPos += lDescLen
                MyBase.moUILib.SendMsgToPrimary(yMsg)

                MyBase.moUILib.AddNotification("Proposal Submitted.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                If lstProposals.ListCount > 0 Then lstProposals.ListIndex = 0
            End If
        End If
        RequestSummaryDetails()
	End Sub

	Public Sub AddSenateProposal(ByRef oProposal As SenateProposal)
		For X As Int32 = 0 To moProposals.GetUpperBound(0)
            If moProposals(X) Is Nothing = False AndAlso moProposals(X).ObjectID = oProposal.ObjectID Then

                With moProposals(X)
                    If oProposal.lVotesAgainst = 0 Then oProposal.lVotesAgainst = .lVotesAgainst
                    If oProposal.lVotesFor = 0 Then oProposal.lVotesFor = .lVotesFor
                End With
                moProposals(X) = oProposal


                Return
            End If
		Next X
		ReDim Preserve moProposals(moProposals.GetUpperBound(0) + 1)
		moProposals(moProposals.GetUpperBound(0)) = oProposal
	End Sub

	Private Function GetProposal(ByVal lProposalID As Int32) As SenateProposal
		If moProposals Is Nothing Then Return Nothing
		For X As Int32 = 0 To moProposals.GetUpperBound(0)
			If moProposals(X) Is Nothing = False AndAlso moProposals(X).ObjectID = lProposalID Then
				Return moProposals(X)
			End If
		Next X
		Return Nothing
	End Function

	Private Sub FillMessageList(ByRef oProposal As SenateProposal)
		lstMessages.Clear()
		Try
            For X As Int32 = 0 To oProposal.lMsgs - 1
                If oProposal.muMsgs Is Nothing = False Then
                    lstMessages.AddItem(oProposal.muMsgs(X).sPosterName, False)
                    lstMessages.ItemData(lstMessages.NewIndex) = oProposal.muMsgs(X).lPosterID
                    lstMessages.ItemData2(lstMessages.NewIndex) = oProposal.muMsgs(X).lPostedOn
                End If
            Next X
		Catch
		End Try
	End Sub

	Private Sub FillChamberVotes(ByRef oProposal As SenateProposal)
		lstVotes.Clear()
		If oProposal.muVotes Is Nothing = False Then
            For X As Int32 = 0 To oProposal.muVotes.GetUpperBound(0)
                If oProposal.muVotes(X).yVote = eyVoteValue.YesVote Then
                    lstVotes.AddItem(oProposal.muVotes(X).VoterName, False)
                    'Because chamber votes can only be endorsements.... we're done
                End If
            Next X
		End If
	End Sub

	Private Sub SmartPopulateSystemVotes(ByRef oProposal As SenateProposal)
		'ok, let's go through our systemvotes which should be our parent nodes...
		Dim clrVoteYes As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 0)
        Dim clrVoteNo As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 0, 0)

		If oProposal.moSystemVotes Is Nothing = False Then
			For X As Int32 = 0 To oProposal.moSystemVotes.GetUpperBound(0)
				Dim oSysVote As SystemVote = oProposal.moSystemVotes(X)
				If oSysVote Is Nothing = False Then
					'ok, get our node
                    Dim sSysName As String = goGalaxy.GetSystemName(oSysVote.SystemID)
					Dim oNode As UITreeView.UITreeViewItem = tvwVotes.GetNodeByItemData2(oSysVote.SystemID, ObjectType.eSolarSystem)
					If oNode Is Nothing Then
						oNode = tvwVotes.AddNode(sSysName, oSysVote.SystemID, ObjectType.eSolarSystem, -1, Nothing, Nothing)
					End If
                    If oNode Is Nothing = False Then
                        Dim lFor As Int32 = 0
                        Dim lAgainst As Int32 = 0
                        'now, go thru those votes
                        If oSysVote.oVotes Is Nothing Then
                            oSysVote.RequestDetails(oProposal.ObjectID)
                        Else
                            For Y As Int32 = 0 To oSysVote.oVotes.GetUpperBound(0)
                                If oSysVote.oVotes(Y).yVote = eyVoteValue.YesVote Then lFor += CInt(oSysVote.oVotes(Y).yRating) Else lAgainst += CInt(oSysVote.oVotes(Y).yRating)

                                Dim sNodeText As String = oSysVote.oVotes(Y).GetNodeText
                                Dim oTmpNode As UITreeView.UITreeViewItem = tvwVotes.GetNodeByItemData3(oSysVote.oVotes(Y).VoterID, oSysVote.SystemID, ObjectType.ePlayer)
                                If oTmpNode Is Nothing Then
                                    oTmpNode = tvwVotes.AddNode(sNodeText, oSysVote.oVotes(Y).VoterID, oSysVote.SystemID, ObjectType.ePlayer, oNode, Nothing)
                                End If
                                If oTmpNode Is Nothing = False Then
                                    If oTmpNode.sItem <> sNodeText Then
                                        oTmpNode.sItem = sNodeText
                                        tvwVotes.IsDirty = True
                                    End If
                                    If oTmpNode.bUseItemColor <> True Then
                                        oTmpNode.bUseItemColor = True
                                        tvwVotes.IsDirty = True
                                    End If
                                    If oSysVote.oVotes(Y).yVote = eyVoteValue.YesVote Then
                                        If oTmpNode.clrItemColor <> clrVoteYes Then
                                            oTmpNode.clrItemColor = clrVoteYes
                                            tvwVotes.IsDirty = True
                                        End If
                                    ElseIf oTmpNode.clrItemColor <> clrVoteNo Then
                                        If oTmpNode.clrItemColor <> clrVoteNo Then
                                            oTmpNode.clrItemColor = clrVoteNo
                                            tvwVotes.IsDirty = True
                                        End If
                                    End If
                                End If
                            Next Y
                        End If

                        sSysName &= " (" & lFor & "/" & lAgainst & ")"
                        If oNode.sItem <> sSysName Then
                            oNode.sItem = sSysName
                            tvwVotes.IsDirty = True
                        End If
                        If oNode.bUseItemColor = False Then
                            oNode.bUseItemColor = True
                            tvwVotes.IsDirty = True
                        End If
                        If oSysVote.ySystemVote = eyVoteValue.YesVote Then
                            If oNode.clrItemColor <> clrVoteYes Then
                                oNode.clrItemColor = clrVoteYes
                                tvwVotes.IsDirty = True
                            End If
                        ElseIf oNode.clrItemColor <> clrVoteNo Then
                            If oNode.clrItemColor <> clrVoteNo Then
                                oNode.clrItemColor = clrVoteNo
                                tvwVotes.IsDirty = True
                            End If
                        End If

                    End If
				End If
			Next X
		Else : tvwVotes.Clear()
        End If
    End Sub

    'Public Sub SetupTest()
    '    ReDim moProposals(2)
    '    moProposals(0) = New SenateProposal()
    '    With moProposals(0)
    '        .lProposedBy = 1
    '        .lProposedOn = GetDateAsNumber(Now.Subtract(New TimeSpan(2, 0, 0, 0)).ToUniversalTime)
    '        .lVotesAgainst = 0 ' 15
    '        .lVotesFor = 25
    '        .moSystemVotes = Nothing
    '        ReDim .muVotes(.lVotesAgainst + .lVotesFor - 1)
    '        For X As Int32 = 0 To .lVotesFor - 1
    '            .muVotes(X) = New SenateVote()
    '            .muVotes(X).VoterName = "Vote For " & X
    '            .muVotes(X).yVote = eyVoteValue.YesVote
    '        Next X
    '        For X As Int32 = .lVotesFor To .lVotesFor + .lVotesAgainst - 1
    '            .muVotes(X) = New SenateVote()
    '            .muVotes(X).VoterName = "Vote Against " & X
    '            .muVotes(X).yVote = eyVoteValue.NoVote
    '        Next X
    '        .lMsgs = 3
    '        ReDim .muMsgs(.lMsgs - 1)
    '        For X As Int32 = 0 To .lMsgs - 1
    '            .muMsgs(X) = New SenateProposalMessage()
    '            .muMsgs(X).lPostedOn = GetDateAsNumber(Now.Subtract(New TimeSpan(2, 0, 0, 0)).ToUniversalTime)
    '            .muMsgs(X).lPosterID = X + 1
    '            .muMsgs(X).sPosterName = "Msger " & (X + 1)
    '            .muMsgs(X).sMsgData = "This is a message " & X
    '        Next X
    '        .ObjectID = 2
    '        .ObjTypeID = ObjectType.eSenateLaw
    '        .sDescription = "Maximum of 10 Pulse Weapons on a unit."
    '        .sProposedByName = "Enoch Dagor"
    '        .sTitle = "Pulse Weapon Limit"
    '        .yProposalState = eyProposalState.EmperorsChamber
    '    End With
    '    moProposals(1) = New SenateProposal
    '    With moProposals(1)
    '        .lProposedBy = 1
    '        .lProposedOn = GetDateAsNumber(Now.Subtract(New TimeSpan(1, 12, 0, 0)).ToUniversalTime)
    '        .lVotesAgainst = 0 '20
    '        .lVotesFor = 10
    '        .moSystemVotes = Nothing
    '        ReDim .muVotes(.lVotesAgainst + .lVotesFor - 1)
    '        For X As Int32 = 0 To .lVotesFor - 1
    '            .muVotes(X) = New SenateVote()
    '            .muVotes(X).VoterName = "Vote For " & X
    '            .muVotes(X).yVote = eyVoteValue.YesVote
    '        Next X
    '        For X As Int32 = .lVotesFor To .lVotesFor + .lVotesAgainst - 1
    '            .muVotes(X) = New SenateVote()
    '            .muVotes(X).VoterName = "Vote Against " & X
    '            .muVotes(X).yVote = eyVoteValue.NoVote
    '        Next X
    '        .lMsgs = 3
    '        ReDim .muMsgs(.lMsgs - 1)
    '        For X As Int32 = 0 To .lMsgs - 1
    '            .muMsgs(X) = New SenateProposalMessage()
    '            .muMsgs(X).lPostedOn = GetDateAsNumber(Now.Subtract(New TimeSpan(2, 0, 0, 0)).ToUniversalTime)
    '            .muMsgs(X).lPosterID = X + 1
    '            .muMsgs(X).sPosterName = "Msger " & (X + 1)
    '            .muMsgs(X).sMsgData = "This is a message " & X
    '        Next X
    '        .ObjectID = 3
    '        .ObjTypeID = ObjectType.eSenateLaw
    '        .sDescription = "Space Station placement closer to planets"
    '        .sProposedByName = "Enoch Dagor"
    '        .sTitle = "Place stations closer"
    '        .yProposalState = eyProposalState.EmperorsChamber
    '    End With
    '    moProposals(2) = New SenateProposal
    '    With moProposals(2)
    '        .lProposedBy = 2
    '        .lProposedOn = GetDateAsNumber(Now.Subtract(New TimeSpan(2, 0, 0, 0)).ToUniversalTime)
    '        .lVotesAgainst = 1923
    '        .lVotesFor = 10042
    '        .muVotes = Nothing

    '        Dim lFor As Int32 = .lVotesFor
    '        Dim lAgainst As Int32 = .lVotesAgainst
    '        ReDim .moSystemVotes(129)
    '        For X As Int32 = 0 To .moSystemVotes.GetUpperBound(0)
    '            .moSystemVotes(X) = New SystemVote
    '            With .moSystemVotes(X)
    '                .SystemID = X + 1
    '                .yVoteRating = CByte(Rnd() * 15 + 1)
    '                ReDim .oVotes(.yVoteRating)
    '                Dim lSysFor As Int32 = 0
    '                Dim lSysAgainst As Int32 = 0
    '                For Y As Int32 = 0 To .yVoteRating
    '                    .oVotes(Y).VoterName = "S" & X & "V" & Y
    '                    .oVotes(Y).VoterID = Y + 1
    '                    .oVotes(Y).yRating = 1
    '                    If lFor > 0 AndAlso lAgainst > 0 Then
    '                        If CInt(Rnd() * 100) > 50 Then
    '                            .oVotes(Y).yVote = eyVoteValue.YesVote
    '                        Else
    '                            .oVotes(Y).yVote = eyVoteValue.NoVote
    '                        End If
    '                    ElseIf lFor > 0 Then
    '                        .oVotes(Y).yVote = eyVoteValue.YesVote
    '                    ElseIf lAgainst > 0 Then
    '                        .oVotes(Y).yVote = eyVoteValue.NoVote
    '                    End If
    '                    If .oVotes(Y).yVote = eyVoteValue.NoVote Then
    '                        lAgainst -= 1
    '                        lSysAgainst += 1
    '                    End If
    '                    If .oVotes(Y).yVote = eyVoteValue.YesVote Then
    '                        lFor -= 1
    '                        lSysFor += 1
    '                    End If
    '                Next Y
    '                If lSysFor > lSysAgainst Then
    '                    .ySystemVote = eyVoteValue.YesVote
    '                Else : .ySystemVote = eyVoteValue.NoVote
    '                End If
    '            End With
    '        Next X
    '        'ReDim .muVotes(.lVotesAgainst + .lVotesFor - 1)
    '        'For X As Int32 = 0 To .lVotesFor - 1
    '        '	.muVotes(X) = New SenateVote()
    '        '	.muVotes(X).VoterName = "Vote For " & X
    '        '	.muVotes(X).yVote = eyVoteValue.YesVote
    '        'Next X
    '        'For X As Int32 = .lVotesFor To .lVotesFor + .lVotesAgainst - 1
    '        '	.muVotes(X) = New SenateVote()
    '        '	.muVotes(X).VoterName = "Vote Against " & X
    '        '	.muVotes(X).yVote = eyVoteValue.NoVote
    '        'Next X
    '        .lMsgs = 3
    '        ReDim .muMsgs(.lMsgs - 1)
    '        For X As Int32 = 0 To .lMsgs - 1
    '            .muMsgs(X) = New SenateProposalMessage()
    '            .muMsgs(X).lPostedOn = GetDateAsNumber(Now.Subtract(New TimeSpan(2, 0, 0, 0)).ToUniversalTime)
    '            .muMsgs(X).lPosterID = X + 1
    '            .muMsgs(X).sPosterName = "Msger " & (X + 1)
    '            .muMsgs(X).sMsgData = "This is a message " & X
    '        Next X
    '        .ObjectID = 4
    '        .ObjTypeID = ObjectType.eSenateLaw
    '        .sDescription = "Remove Orbital Bombardment. It is nothing but a curse."
    '        .sProposedByName = "Csaj Schnutic"
    '        .sTitle = "Remove Orbital Bombardment"
    '        .yProposalState = eyProposalState.AwaitingApproval Or eyProposalState.OnSenateFloor
    '    End With

    '    For X As Int32 = 0 To moProposals.GetUpperBound(0)
    '        moProposals(X).ReSortData()
    '    Next X
    'End Sub

	Public Sub HandleGetSenateObjectDetails(ByRef yData() As Byte)
        Dim lPos As Int32 = 2 'for smgcode
        lPos += 4        'for playerid ph
		Dim yType As Byte = yData(lPos) : lPos += 1

		Select Case yType
			Case eySenateRequestDetailsType.ProposalMsg
				Dim lProposalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				Dim oProposal As SenateProposal = GetProposal(lProposalID)
				If oProposal Is Nothing = False AndAlso oProposal.muMsgs Is Nothing = False Then
					Dim lPosterID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					Dim lPostedOn As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					Dim sPostedBy As String = GetStringFromBytes(yData, lPos, 20) : lPos += 20
					Dim lMsgLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
					Dim sMsg As String = GetStringFromBytes(yData, lPos, lMsgLen) : lPos += lMsgLen

					For X As Int32 = 0 To oProposal.muMsgs.GetUpperBound(0)
						If oProposal.muMsgs(X).lPosterID = lPosterID AndAlso oProposal.muMsgs(X).lPostedOn = lPostedOn Then
							oProposal.muMsgs(X).sPosterName = sPostedBy
							oProposal.muMsgs(X).sMsgData = sMsg
							Exit For
						End If
					Next X
				End If
			Case eySenateRequestDetailsType.SystemVote
				Dim lSystemID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				Dim lProposalID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				Dim oProposal As SenateProposal = GetProposal(lProposalID)
				If oProposal Is Nothing = False Then
					If oProposal.moSystemVotes Is Nothing = False Then
						For X As Int32 = 0 To oProposal.moSystemVotes.GetUpperBound(0)
							If oProposal.moSystemVotes(X) Is Nothing = False AndAlso oProposal.moSystemVotes(X).SystemID = lSystemID Then
								oProposal.moSystemVotes(X).HandleDetailsMsg(yData)
								Exit For
							End If
						Next X
					End If
                End If
            Case eySenateRequestDetailsType.EmpChmbrMsg
                Dim oFrm As frmEmpMsgBrd = CType(goUILib.GetWindow("frmEmpMsgBrd"), frmEmpMsgBrd)
                If oFrm Is Nothing = False Then
                    oFrm.HandleEmpChmbrMsg(yData)
                End If
            Case eySenateRequestDetailsType.EmpChmbrMsgList
                Dim oFrm As frmEmpMsgBrd = CType(goUILib.GetWindow("frmEmpMsgBrd"), frmEmpMsgBrd)
                If oFrm Is Nothing = False Then
                    oFrm.HandleEmpChmbrList(yData)
                End If
        End Select
    End Sub

    Public Sub HandleAddSenateProposal(ByVal yData() As Byte)
        Dim oSenateProposal As New SenateProposal()
        Dim yProposalState As Byte = yData(12)
        If (yProposalState And eyProposalState.SmallMsgBitShift) <> 0 Then
            oSenateProposal.FillFromSmallMsg(yData)
        Else : oSenateProposal.FillFromBigMsg(yData)
        End If
        AddSenateProposal(oSenateProposal)
    End Sub

    Private Sub lstMessages_ItemClick(ByVal lIndex As Integer) Handles lstMessages.ItemClick
        btnAddMessage.Caption = "Add Message"
        txtMessage.Caption = "" : txtMessage.Locked = True
        txtMessage.BackColorEnabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
    End Sub

    Private mbNoSendOptClick As Boolean = False
    Private mbIgnoreOptClick As Boolean = False
    Private Sub opt1_Click() Handles opt1.Click
        SelectOption(1)
    End Sub

    Private Sub opt2_Click() Handles opt2.Click
        SelectOption(2)
    End Sub

    Private Sub opt3_Click() Handles opt3.Click
        SelectOption(3)
    End Sub

    Private Sub opt4_Click() Handles opt4.Click
        SelectOption(4)
    End Sub

    Private Sub opt5_Click() Handles opt5.Click
        SelectOption(5)
    End Sub

    Private Sub SelectOption(ByVal lVal As Int32)
        If mbIgnoreOptClick = True Then Return
        mbIgnoreOptClick = True

        opt1.Value = (lVal = 1)
        opt2.Value = (lVal = 2)
        opt3.Value = (lVal = 3)
        opt4.Value = (lVal = 4)
        opt5.Value = (lVal = 5)

        'Now, submit the vote option update
        If mbNoSendOptClick = False AndAlso lstProposals Is Nothing = False Then
            If lstProposals.ListIndex > -1 Then
                Dim oProposal As SenateProposal = GetProposal(lstProposals.ItemData(lstProposals.ListIndex))
                If oProposal Is Nothing = False Then
                    If oProposal.yPlayerVote = eyVoteValue.YesVote Then
                        btnVoteYes_Click(btnVoteYes.Caption)
                    ElseIf oProposal.yPlayerVote = eyVoteValue.NoVote Then
                        btnVoteNo_Click(btnVoteNo.Caption)
                    End If
                End If
            End If
        End If

        mbIgnoreOptClick = False
    End Sub

    Private Sub btnDeleteProposal_Click(ByVal sName As String) Handles btnDeleteProposal.Click

        If lstProposals.ListIndex > -1 Then
            Dim oProposal As SenateProposal = GetProposal(lstProposals.ItemData(lstProposals.ListIndex))
            If oProposal Is Nothing Then Return
            If oProposal.lProposedBy <> glPlayerID Then
                MyBase.moUILib.AddNotification("You can only delete proposals you have proposed.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            btnDeleteProposal.Enabled = False

            Dim yMsg(13) As Byte
            Dim lPos As Int32 = 0

            System.BitConverter.GetBytes(GlobalMessageCode.eAddSenateProposal).CopyTo(yMsg, lPos) : lPos += 2
            lPos += 4       'for playerid
            System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(oProposal.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
            MyBase.moUILib.SendMsgToPrimary(yMsg)


            'ReDim yMsg(6)
            'lPos = 0
            'System.BitConverter.GetBytes(GlobalMessageCode.eGetSenateObjectDetails).CopyTo(yMsg, lPos) : lPos += 2
            'lPos += 4       'for player id
            'yMsg(lPos) = eySenateRequestDetailsType.SenateObject : lPos += 1
            'MyBase.moUILib.SendMsgToPrimary(yMsg)
            RequestSummaryDetails()

            btnDeleteProposal.Enabled = True
        Else
            MyBase.moUILib.AddNotification("Select a proposal to delete first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If

        
    End Sub

    Private Sub txtDescription_OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtDescription.OnKeyPress
        If txtDescription.Locked = False Then
            mbEditting = True
            If btnVoteYes.Caption <> "Save" Then btnVoteYes.Caption = "Save"
            If btnVoteNo.Caption <> "Cancel" Then btnVoteNo.Caption = "Cancel"
        End If
    End Sub

    Private Sub txtTitle_OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtTitle.OnKeyPress
        If txtTitle.Locked = False Then
            mbEditting = True
            If btnVoteYes.Caption <> "Save" Then btnVoteYes.Caption = "Save"
            If btnVoteNo.Caption <> "Cancel" Then btnVoteNo.Caption = "Cancel"
        End If
    End Sub

    Private Sub btnEmpMsgBrd_Click(ByVal sName As String) Handles btnEmpMsgBrd.Click
        If myViewState = 1 Then
            If goCurrentPlayer Is Nothing = False Then
                If (goCurrentPlayer.yPlayerTitle <> Player.PlayerRank.Emperor OrElse gbAliased = True) AndAlso glPlayerID <> 1 AndAlso glPlayerID <> 6 Then
                    MyBase.moUILib.AddNotification("You do not have authorization to enter the Emperor's Chamber.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
                    Return
                End If
            End If
        End If

        Dim oFrm As New frmEmpMsgBrd(goUILib)
        oFrm.Visible = True
        oFrm = Nothing
    End Sub

    Private Sub RequestSummaryDetails()

        'Re-Request details to fill the top list.
        Dim yMsg(6) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eGetSenateObjectDetails).CopyTo(yMsg, lPos) : lPos += 2
        lPos += 4       'for player id
        yMsg(lPos) = eySenateRequestDetailsType.SenateObject : lPos += 1
        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub
End Class