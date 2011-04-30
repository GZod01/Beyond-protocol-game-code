Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmGuildMain
	Private Class fra_Events
		Inherits guildframe

		Private lblEvents As UILabel
		Private lblEventTitle As UILabel
		Private fraRecurrence As UIWindow
		Private txtEventTitle As UITextBox
		Private lblPostedBy As UILabel
		Private lblDailyDetail As UILabel
		Private lblAnnuallyDetail As UILabel
		Private lblDuration As UILabel
		Private lblStartsAt As UILabel
		Private lblDetails As UILabel
		Private lblEventIcon As UILabel
		Private lblEventType As UILabel
		Private lblMonthyDetail As UILabel
		Private txtPostedBy As UITextBox
		Private lblPostedOn As UILabel
		Private lblWeeklyDetail As UILabel
		Private lblSendAlerts As UILabel
		Private txtPostedOn As UITextBox
		Private txtDetails As UITextBox
		Private chkMembersCanAccept As UICheckBox
		Private cboSendAlerts As UIComboBox
		Private cboEventType As UIComboBox
		Private txtDuration As UITextBox
		Private cboStartsAt As UIComboBox

		Private WithEvents fraEventDetails As UIWindow
		Private WithEvents chkRecurs As UICheckBox
		Private WithEvents moCalendar As ctlCalendar
		Private WithEvents btnAccept As UIButton
		Private WithEvents btnAttachments As UIButton
		Private WithEvents btnCancelEvent As UIButton
		Private WithEvents btnCreateEvent As UIButton
		Private WithEvents btnDecline As UIButton
		Private WithEvents btnAdvanced As UIButton
		Private WithEvents lstEvents As UIListBox
		Private WithEvents optRecursDaily As UIOption
		Private WithEvents optRecursWeekly As UIOption
		Private WithEvents optRecursMonthly As UIOption
		Private WithEvents optRecursAnnually As UIOption

		Private myEventIcon As Byte = 0

		Private mbInNewEvent As Boolean = False

		Public Sub New(ByRef oUILib As UILib)
			MyBase.New(oUILib)

            With Me
                .lWindowMetricID = BPMetrics.eWindow.eGuildMain_Events
                .ControlName = "fra_Events"
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

			If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
				If goCurrentPlayer.oGuild.HasPermission(RankPermissions.ViewEvents) = False Then Return
			End If

			'calendar initial props
			moCalendar = New ctlCalendar(MyBase.moUILib)
			With moCalendar
				.ControlName = "moCalendar"
				.Left = 15
				.Top = 5
				.Visible = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Visible = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ViewEvents)
				End If
				.Enabled = True
				.Moveable = False
			End With
			Me.AddChild(CType(moCalendar, UIControl))

			'lblEvents initial props
			lblEvents = New UILabel(MyBase.moUILib)
			With lblEvents
				.ControlName = "lblEvents"
				.Left = 15
				.Top = 275
				.Width = 255
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Events Scheduled for " & Now.Month.ToString("00") & "/" & Now.Day.ToString("00") & "/" & Now.Year.ToString("00")
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			Me.AddChild(CType(lblEvents, UIControl))

			'lstEvents initial props
			lstEvents = New UIListBox(MyBase.moUILib)
			With lstEvents
				.ControlName = "lstEvents"
				.Left = 15
				.Top = 300
				.Width = 256
				.Height = 180
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
				.oIconTexture = goResMgr.GetTexture("GuildIcons.dds", GFXResourceManager.eGetTextureType.UserInterface, "gi.pak")
				.RenderIcons = True
			End With
			Me.AddChild(CType(lstEvents, UIControl))

			'btnCreateEvent initial props
			btnCreateEvent = New UIButton(MyBase.moUILib)
			With btnCreateEvent
				.ControlName = "btnCreateEvent"
				.Left = 15
				.Top = 495
				.Width = 100
				.Height = 24
				.Enabled = False
				.Visible = True
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.CreateEvents)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to create events."
				End If
				.Caption = "Create Event"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			Me.AddChild(CType(btnCreateEvent, UIControl))

			'btnCancelEvent initial props
			btnCancelEvent = New UIButton(MyBase.moUILib)
			With btnCancelEvent
				.ControlName = "btnCancelEvent"
				.Left = 175
				.Top = 495
				.Width = 100
				.Height = 24
				.Enabled = False
				.Visible = True
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.DeleteEvents)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to delete events."
				End If
				.Caption = "Cancel Event"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			Me.AddChild(CType(btnCancelEvent, UIControl))

			'fraEventDetails initial props
			fraEventDetails = New UIWindow(MyBase.moUILib)
			With fraEventDetails
				.ControlName = "fraEventDetails"
				.Left = 290
				.Top = 5
				.Width = 495
				.Height = 515
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.Moveable = False
				.BorderLineWidth = 2
			End With
			Me.AddChild(CType(fraEventDetails, UIControl))

			'=== begin of fraEventDetails ===
			'lblEventTitle initial props
			lblEventTitle = New UILabel(MyBase.moUILib)
			With lblEventTitle
				.ControlName = "lblEventTitle"
				.Left = 15
				.Top = 15
				.Width = 85
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Event Title:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraEventDetails.AddChild(CType(lblEventTitle, UIControl))

			'txtEventTitle initial props
			txtEventTitle = New UITextBox(MyBase.moUILib)
			With txtEventTitle
				.ControlName = "txtEventTitle"
				.Left = 120
				.Top = 17
				.Width = 360
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
				.MaxLength = 0
				.BorderColor = muSettings.InterfaceBorderColor
			End With
			fraEventDetails.AddChild(CType(txtEventTitle, UIControl))

			'lblPostedBy initial props
			lblPostedBy = New UILabel(MyBase.moUILib)
			With lblPostedBy
				.ControlName = "lblPostedBy"
				.Left = 15
				.Top = 50
				.Width = 81
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Posted By:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraEventDetails.AddChild(CType(lblPostedBy, UIControl))

			'txtPostedBy initial props
			txtPostedBy = New UITextBox(MyBase.moUILib)
			With txtPostedBy
				.ControlName = "txtPostedBy"
				.Left = 120
				.Top = 52
				.Width = 124
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
				.MaxLength = 0
				.BorderColor = muSettings.InterfaceBorderColor
				.Locked = True
			End With
			fraEventDetails.AddChild(CType(txtPostedBy, UIControl))

			'lblPostedOn initial props
			lblPostedOn = New UILabel(MyBase.moUILib)
			With lblPostedOn
				.ControlName = "lblPostedOn"
				.Left = 265
				.Top = 50
				.Width = 86
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Posted On:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraEventDetails.AddChild(CType(lblPostedOn, UIControl))

			'txtPostedOn initial props
			txtPostedOn = New UITextBox(MyBase.moUILib)
			With txtPostedOn
				.ControlName = "txtPostedOn"
				.Left = 355
				.Top = 52
				.Width = 124
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
				.MaxLength = 0
				.BorderColor = muSettings.InterfaceBorderColor
				.Locked = True
			End With
			fraEventDetails.AddChild(CType(txtPostedOn, UIControl))

			'lblStartsAt initial props
			lblStartsAt = New UILabel(MyBase.moUILib)
			With lblStartsAt
				.ControlName = "lblStartsAt"
				.Left = 15
				.Top = 85
				.Width = 81
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Starts At:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraEventDetails.AddChild(CType(lblStartsAt, UIControl))

			'lblEndsAt initial props
			lblDuration = New UILabel(MyBase.moUILib)
			With lblDuration
				.ControlName = "lblDuration"
				.Left = 265
				.Top = 85
				.Width = 140
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Duration (in mins):"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraEventDetails.AddChild(CType(lblDuration, UIControl))

			'lblEventType initial props
			lblEventType = New UILabel(MyBase.moUILib)
			With lblEventType
				.ControlName = "lblEventType"
				.Left = 15
				.Top = 120
				.Width = 81
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Event Type:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraEventDetails.AddChild(CType(lblEventType, UIControl))

			'lblSendAlerts initial props
			lblSendAlerts = New UILabel(MyBase.moUILib)
			With lblSendAlerts
				.ControlName = "lblSendAlerts"
				.Left = 265
				.Top = 120
				.Width = 85
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Send Alerts:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraEventDetails.AddChild(CType(lblSendAlerts, UIControl))

			'lblEventIcon initial props
			lblEventIcon = New UILabel(MyBase.moUILib)
			With lblEventIcon
				.ControlName = "lblEventIcon"
				.Left = 15
				.Top = 159
				.Width = 81
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Event Icon:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraEventDetails.AddChild(CType(lblEventIcon, UIControl))

			'chkMembersCanAccept initial props
			chkMembersCanAccept = New UICheckBox(MyBase.moUILib)
			With chkMembersCanAccept
				.ControlName = "chkMembersCanAccept"
				.Left = 15
				.Top = 190
				.Width = 210
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Members Can Accept Event"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraEventDetails.AddChild(CType(chkMembersCanAccept, UIControl))

			'lblDetails initial props
			lblDetails = New UILabel(MyBase.moUILib)
			With lblDetails
				.ControlName = "lblDetails"
				.Left = 15
				.Top = 345
				.Width = 81
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Details:"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraEventDetails.AddChild(CType(lblDetails, UIControl))

			'txtDetails initial props
			txtDetails = New UITextBox(MyBase.moUILib)
			With txtDetails
				.ControlName = "txtDetails"
				.Left = 15
				.Top = 365
				.Width = 465
				.Height = 100
				.Enabled = True
				.Visible = True
				.Caption = "Details are typed here"
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
			fraEventDetails.AddChild(CType(txtDetails, UIControl))

			'btnAttachments initial props
			btnAttachments = New UIButton(MyBase.moUILib)
			With btnAttachments
				.ControlName = "btnAttachments"
				.Left = 15
				.Top = 480
				.Width = 100
				.Height = 24
				.Enabled = False
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.ViewEventAttachments)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to view event attachments."
				End If
				.Visible = True
				.Caption = "Attachments"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraEventDetails.AddChild(CType(btnAttachments, UIControl))

			'btnAdvanced initial props
			btnAdvanced = New UIButton(MyBase.moUILib)
			With btnAdvanced
				.ControlName = "btnAdvanced"
				.Left = btnAttachments.Left + btnAttachments.Width + 15
				.Top = 480
				.Width = 100
				.Height = 24
				.Enabled = False
				.Visible = True
				.Caption = "Advanced"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraEventDetails.AddChild(CType(btnAdvanced, UIControl))

			'btnDecline initial props
			btnDecline = New UIButton(MyBase.moUILib)
			With btnDecline
				.ControlName = "btnDecline"
				.Left = btnAdvanced.Left + btnAdvanced.Width + 15
				.Top = 480
				.Width = 100
				.Height = 24
				.Enabled = False
				.Visible = True
				.Caption = "Decline"
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.AcceptEvents)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to accept or decline events."
				End If
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraEventDetails.AddChild(CType(btnDecline, UIControl))

			'btnAccept initial props
			btnAccept = New UIButton(MyBase.moUILib)
			With btnAccept
				.ControlName = "btnAccept"
				.Left = btnDecline.Left + btnDecline.Width + 15
				.Top = 480
				.Width = 100
				.Height = 24
				.Enabled = False
				.Visible = True
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.AcceptEvents)
				End If
				If .Enabled = False Then
					.ToolTipText = "You lack the rights to accept or decline events."
				End If
				.Caption = "Accept/Attend"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			fraEventDetails.AddChild(CType(btnAccept, UIControl))

			txtDuration = New UITextBox(oUILib)
			With txtDuration
				.ControlName = "txtDuration"
				.Left = 415
				.Top = 87
				.Width = 60
				.Height = 22
				.Enabled = True
				.Visible = True
				.Caption = "1"
				.ForeColor = muSettings.InterfaceTextBoxForeColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(0, DrawTextFormat)
				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
				.MaxLength = 8
				.BorderColor = muSettings.InterfaceBorderColor
			End With
			fraEventDetails.AddChild(CType(txtDuration, UIControl))

			'fraRecurrence initial props
			fraRecurrence = New UIWindow(MyBase.moUILib)
			With fraRecurrence
				.ControlName = "fraRecurrence"
				.Left = 15
				.Top = 225
				.Width = 465
				.Height = 115
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.BorderLineWidth = 2
				.Moveable = False
			End With
			fraEventDetails.AddChild(CType(fraRecurrence, UIControl))

			'=== fraRecurrence Begins Here ===
			'chkRecurs initial props
			chkRecurs = New UICheckBox(MyBase.moUILib)
			With chkRecurs
				.ControlName = "chkRecurs"
				.Left = 15
				.Top = 10
				.Width = 186
				.Height = 20
				.Enabled = True
				.Visible = True
				.Caption = "This is a recurring event"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraRecurrence.AddChild(CType(chkRecurs, UIControl))

			'optRecursDaily initial props
			optRecursDaily = New UIOption(MyBase.moUILib)
			With optRecursDaily
				.ControlName = "optRecursDaily"
				.Left = 30
				.Top = 25
				.Width = 50
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Daily"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraRecurrence.AddChild(CType(optRecursDaily, UIControl))

			'optRecursWeekly initial props
			optRecursWeekly = New UIOption(MyBase.moUILib)
			With optRecursWeekly
				.ControlName = "optRecursWeekly"
				.Left = 30
				.Top = 45
				.Width = 65
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Weekly"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraRecurrence.AddChild(CType(optRecursWeekly, UIControl))

			'optRecursMonthly initial props
			optRecursMonthly = New UIOption(MyBase.moUILib)
			With optRecursMonthly
				.ControlName = "optRecursMonthly"
				.Left = 30
				.Top = 65
				.Width = 65
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Monthly"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraRecurrence.AddChild(CType(optRecursMonthly, UIControl))

			'optRecursAnnually initial props
			optRecursAnnually = New UIOption(MyBase.moUILib)
			With optRecursAnnually
				.ControlName = "optRecursAnnually"
				.Left = 30
				.Top = 85
				.Width = 71
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Annually"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(6, DrawTextFormat)
				.Value = False
			End With
			fraRecurrence.AddChild(CType(optRecursAnnually, UIControl))

			'lblDailyDetail initial props
			lblDailyDetail = New UILabel(MyBase.moUILib)
			With lblDailyDetail
				.ControlName = "lblDailyDetail"
				.Left = 130
				.Top = 25
				.Width = 242
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "- Every day at the scheduled times"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraRecurrence.AddChild(CType(lblDailyDetail, UIControl))

			'lblWeeklyDetail initial props
			lblWeeklyDetail = New UILabel(MyBase.moUILib)
			With lblWeeklyDetail
				.ControlName = "lblWeeklyDetail"
				.Left = 130
				.Top = 45
				.Width = 283
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "- Every week on the day of the scheduled times"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraRecurrence.AddChild(CType(lblWeeklyDetail, UIControl))

			'lblMonthyDetail initial props
			lblMonthyDetail = New UILabel(MyBase.moUILib)
			With lblMonthyDetail
				.ControlName = "lblMonthyDetail"
				.Left = 130
				.Top = 65
				.Width = 242
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "- Every month on the date specified"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraRecurrence.AddChild(CType(lblMonthyDetail, UIControl))

			'lblAnnuallyDetail initial props
			lblAnnuallyDetail = New UILabel(MyBase.moUILib)
			With lblAnnuallyDetail
				.ControlName = "lblAnnuallyDetail"
				.Left = 130
				.Top = 85
				.Width = 242
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "- Every year on the date specified"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
			End With
			fraRecurrence.AddChild(CType(lblAnnuallyDetail, UIControl))

			'cboSendAlerts initial props
			cboSendAlerts = New UIComboBox(MyBase.moUILib)
			With cboSendAlerts
				.ControlName = "cboSendAlerts"
				.Left = 355
				.Top = 122
				.Width = 123
				.Height = 22
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
            End With
            fraEventDetails.AddChild(CType(cboSendAlerts, UIControl))

            'cboEventType initial props
            cboEventType = New UIComboBox(MyBase.moUILib)
            With cboEventType
                .ControlName = "cboEventType"
                .Left = 120
                .Top = 122
                .Width = 123
                .Height = 22
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
            End With
            fraEventDetails.AddChild(CType(cboEventType, UIControl))

            'cboStartsAt initial props
            cboStartsAt = New UIComboBox(MyBase.moUILib)
            With cboStartsAt
                .ControlName = "cboStartsAt"
                .Left = 120
                .Top = 87
                .Width = 123
                .Height = 22
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
            End With
			fraEventDetails.AddChild(CType(cboStartsAt, UIControl))

			moCalendar.SetSelectDate(Now.Day, Now.Month, Now.Year)

			FillValues()
		End Sub

		Private Sub FillValues()
			cboStartsAt.Clear()
			For X As Int32 = 0 To 95
				If X = 0 Then
					cboStartsAt.AddItem("Midnight")
					cboStartsAt.ItemData(cboStartsAt.NewIndex) = 0
				Else
					Dim lHr As Int32 = X \ 4
					Dim lMin As Int32 = (X Mod 4) * 15
					cboStartsAt.AddItem(lHr.ToString("00") & ":" & lMin.ToString("00"))
					cboStartsAt.ItemData(cboStartsAt.NewIndex) = CInt(lHr.ToString("00") & lMin.ToString("00"))
				End If
				
			Next X

			cboEventType.Clear()
			cboEventType.AddItem("Meeting") : cboEventType.ItemData(cboEventType.NewIndex) = 1
			cboEventType.AddItem("Race") : cboEventType.ItemData(cboEventType.NewIndex) = 2
			cboEventType.AddItem("Taxes") : cboEventType.ItemData(cboEventType.NewIndex) = 3
			cboEventType.AddItem("Tournament") : cboEventType.ItemData(cboEventType.NewIndex) = 4

			cboSendAlerts.Clear()
			cboSendAlerts.AddItem("Do Not Send") : cboSendAlerts.ItemData(cboSendAlerts.NewIndex) = 0
			cboSendAlerts.AddItem("Send To Acceptors") : cboSendAlerts.ItemData(cboSendAlerts.NewIndex) = 1
			cboSendAlerts.AddItem("Send To All") : cboSendAlerts.ItemData(cboSendAlerts.NewIndex) = 2
		End Sub

		Private Sub UpdateGuildEvent(ByVal lEventID As Int32, ByVal yAcceptance As Byte)
			Dim yMsg(6) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eUpdateGuildEvent).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(lEventID).CopyTo(yMsg, lPos) : lPos += 4
			yMsg(lPos) = yAcceptance : lPos += 1
			MyBase.moUILib.SendMsgToPrimary(yMsg)
		End Sub

		Private Sub SubmitGuildEvent()
			'submit new event

			If txtEventTitle.Caption.Trim = "" Then
				MyBase.moUILib.AddNotification("Enter an event title for this event.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If
			If cboStartsAt.ListIndex < 0 Then
				MyBase.moUILib.AddNotification("Select a valid Start Time for this event.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If
			If txtDuration.Caption.Trim = "" OrElse IsNumeric(txtDuration.Caption) = False Then
				MyBase.moUILib.AddNotification("Enter a valid duration for this event.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If
			If cboEventType.ListIndex < 0 Then
				MyBase.moUILib.AddNotification("Select an event type for this event.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If
			If cboSendAlerts.ListIndex < 0 Then
				cboSendAlerts.ListIndex = 0
			End If

			Dim lEventID As Int32 = -1
			If lstEvents.ListIndex > -1 Then
				lEventID = lstEvents.ItemData(lstEvents.ListIndex)
			End If
			Dim sDetails As String = txtDetails.Caption
			Dim lLen As Int32 = sDetails.Length

			Dim yMsg(72 + lLen) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eUpdateGuildEvent).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(lEventID).CopyTo(yMsg, lPos) : lPos += 4

			Dim sTitle As String = txtEventTitle.Caption
			If sTitle.Length > 50 Then sTitle = sTitle.Substring(0, 50)
			System.Text.ASCIIEncoding.ASCII.GetBytes(sTitle).CopyTo(yMsg, lPos) : lPos += 50

			Dim lDay As Int32 = 0
			Dim lMonth As Int32 = 0
			Dim lYear As Int32 = 0
			moCalendar.GetSelectedDate(lDay, lMonth, lYear)

			Dim dtTemp As Date = Date.SpecifyKind(GetLocaleSpecificDT(lMonth.ToString("00") & "/" & lDay.ToString("00") & "/" & lYear.ToString("00") & " " & cboStartsAt.ItemData(cboStartsAt.ListIndex).ToString("00:00")), DateTimeKind.Local)
			dtTemp = dtTemp.ToUniversalTime
			Dim lStarts As Int32 = GetDateAsNumber(dtTemp)
			System.BitConverter.GetBytes(lStarts).CopyTo(yMsg, lPos) : lPos += 4
			Dim lDuration As Int32 = CInt(txtDuration.Caption)
			System.BitConverter.GetBytes(lDuration).CopyTo(yMsg, lPos) : lPos += 4

			yMsg(lPos) = CByte(cboEventType.ItemData(cboEventType.ListIndex)) : lPos += 1
			yMsg(lPos) = CByte(cboSendAlerts.ItemData(cboSendAlerts.ListIndex)) : lPos += 1
			yMsg(lPos) = myEventIcon : lPos += 1
			If chkMembersCanAccept.Value = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 0
			lPos += 1

			Dim yRecurrence As Byte = 0
			If chkRecurs.Value = True Then
				If optRecursDaily.Value = True Then
					yRecurrence = 1
				ElseIf optRecursWeekly.Value = True Then
					yRecurrence = 2
				ElseIf optRecursMonthly.Value = True Then
					yRecurrence = 3
				ElseIf optRecursAnnually.Value = True Then
					yRecurrence = 4
				End If
			End If
			yMsg(lPos) = yRecurrence : lPos += 1

			System.BitConverter.GetBytes(lLen).CopyTo(yMsg, lPos) : lPos += 4
			System.Text.ASCIIEncoding.ASCII.GetBytes(sDetails).CopyTo(yMsg, lPos) : lPos += lLen

			MyBase.moUILib.SendMsgToPrimary(yMsg)

			MyBase.moUILib.AddNotification("Submitting New Guild Event...", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		End Sub

		Private Sub btnAccept_Click(ByVal sName As String) Handles btnAccept.Click
			If mbInNewEvent = True Then
				'ok, creating a new event
				SubmitGuildEvent()
			Else
				'ok, saving the event if i am the owner or accepting the event if i am not the owner
				If lstEvents.ListIndex > -1 Then
					Dim lID As Int32 = lstEvents.ItemData(lstEvents.ListIndex)
					Dim oEvent As GuildEvent = goCurrentPlayer.oGuild.GetEvent(lID)
					If oEvent Is Nothing = False Then
						If oEvent.lPostedBy = glPlayerID Then
							'ok, saving changes
							SubmitGuildEvent()
						Else
							'ok, accepting event
							UpdateGuildEvent(lID, 1)
						End If
					End If
				End If
			End If
		End Sub

		Private Sub btnAttachments_Click(ByVal sName As String) Handles btnAttachments.Click
			If lstEvents.ListIndex > -1 Then

				Dim oEvent As GuildEvent = Nothing
				Dim lID As Int32 = lstEvents.ItemData(lstEvents.ListIndex)
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False AndAlso goCurrentPlayer.oGuild.moEvents Is Nothing = False Then
					With goCurrentPlayer.oGuild
						For X As Int32 = 0 To .moEvents.GetUpperBound(0)
							If .moEvents(X) Is Nothing = False AndAlso .moEvents(X).EventID = lID Then
								oEvent = .moEvents(X)
								Exit For
							End If
						Next X
					End With
				End If
				If oEvent Is Nothing = False Then
					Dim ofrm As frmAttachments = New frmAttachments(MyBase.moUILib)
					ofrm.SetFromEvent(oEvent)
					ofrm.Visible = True
				End If
			Else
				MyBase.moUILib.AddNotification("You must submit the event before adding attachments.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			End If
		End Sub

		Private Sub btnCancelEvent_Click(ByVal sName As String) Handles btnCancelEvent.Click
			If lstEvents.ListIndex > -1 Then
				Dim lID As Int32 = lstEvents.ItemData(lstEvents.ListIndex)

				If btnCancelEvent.Caption.ToUpper = "CONFIRM" Then
					'Ok, cancel the event
					UpdateGuildEvent(lID, 255)
				Else
					btnCancelEvent.Caption = "Confirm"
				End If
			End If
		End Sub

		Private Sub btnCreateEvent_Click(ByVal sName As String) Handles btnCreateEvent.Click
			lstEvents.ListIndex = -1
			SetFromEvent(Nothing)
			mbInNewEvent = True
		End Sub

		Private Sub btnDecline_Click(ByVal sName As String) Handles btnDecline.Click
			If lstEvents.ListIndex > -1 Then
				Dim lID As Int32 = lstEvents.ItemData(lstEvents.ListIndex)
				Dim oEvent As GuildEvent = goCurrentPlayer.oGuild.GetEvent(lID)
				If oEvent Is Nothing = False Then
					'ok, accepting event
					UpdateGuildEvent(lID, 2)
				End If
			End If
		End Sub

		Private Sub chkRecurs_Click() Handles chkRecurs.Click
			Dim bChecked As Boolean = chkRecurs.Value

			optRecursAnnually.Enabled = bChecked
			optRecursDaily.Enabled = bChecked
			optRecursMonthly.Enabled = bChecked
			optRecursWeekly.Enabled = bChecked
		End Sub

		Private Sub lstEvents_ItemClick(ByVal lIndex As Integer) Handles lstEvents.ItemClick
			If lIndex > -1 Then
				mbInNewEvent = False
				Dim lEventID As Int32 = lstEvents.ItemData(lIndex)
				'ok, got an eventid
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					Dim oEvent As GuildEvent = goCurrentPlayer.oGuild.GetEvent(lEventID)
					SetFromEvent(oEvent)
				End If
			End If
		End Sub

		Private Sub moCalendar_DayClick(ByVal lDay As Integer, ByVal lMonth As Integer, ByVal lYear As Integer) Handles moCalendar.DayClick
			If lblEvents Is Nothing = False Then
				lblEvents.Caption = "Scheduled Events for " & lMonth.ToString("00") & "/" & lDay.ToString("00") & "/" & lYear.ToString("00")
			End If

			If lstEvents Is Nothing = False Then
				lstEvents.Clear()
				'Fill the List box with events
				If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
					If goCurrentPlayer.oGuild.moEvents Is Nothing = False Then
						With goCurrentPlayer.oGuild
							Dim oEvents() As GuildEvent = Nothing
							Dim lEventUB As Int32 = -1

							For X As Int32 = 0 To .moEvents.GetUpperBound(0)
								If .moEvents(X) Is Nothing = False Then
									Dim dt As Date = .moEvents(X).dtStartsAt.ToLocalTime
									If lDay = dt.Day AndAlso lMonth = dt.Month AndAlso lYear = dt.Year Then
										'Ok, find where to place this
										Dim lIdx As Int32 = -1
										For Y As Int32 = 0 To lEventUB
											If oEvents(Y).dtStartsAt.ToLocalTime.TimeOfDay.TotalSeconds > dt.TimeOfDay.TotalSeconds Then
												lIdx = Y
												Exit For
											End If
										Next Y
										'Ok, check our lidx
										lEventUB += 1
										ReDim Preserve oEvents(lEventUB)
										If lIdx = -1 Then
											lIdx = lEventUB
										Else
											For Y As Int32 = lEventUB To lIdx + 1 Step -1
												oEvents(Y) = oEvents(Y - 1)
											Next Y
										End If
										oEvents(lIdx) = .moEvents(X)
										Dim yTemp() As Byte = .moEvents(X).RequestDetails()
										If yTemp Is Nothing = False Then MyBase.moUILib.SendMsgToPrimary(yTemp)
									End If
								End If
							Next X

							'Now, go back through our list
							For X As Int32 = 0 To lEventUB
								Dim yTemp() As Byte = oEvents(X).RequestDetails()
								If yTemp Is Nothing = False Then MyBase.moUILib.SendMsgToPrimary(yTemp)

								lstEvents.AddItem(oEvents(X).ListboxString, False)
								lstEvents.ItemData(lstEvents.NewIndex) = oEvents(X).EventID
								lstEvents.rcIconRectangle(lstEvents.NewIndex) = oEvents(X).rcIconRect
								lstEvents.ApplyIconOffset(lstEvents.NewIndex) = True
								lstEvents.IconForeColor(lstEvents.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 255, 255)
								Dim dtLocal As Date = oEvents(X).dtStartsAt.ToLocalTime()
								Dim lKey As Int32 = CInt(dtLocal.Month.ToString("00") & dtLocal.Day.ToString("00") & dtLocal.Year.ToString("0000"))
								lstEvents.ItemData3(lstEvents.NewIndex) = lKey
							Next X

						End With
					End If
				End If
			End If
		End Sub

		Private Sub optRecursAnnually_Click() Handles optRecursAnnually.Click
			optRecursAnnually.Value = True
			optRecursDaily.Value = False
			optRecursMonthly.Value = False
			optRecursWeekly.Value = False
		End Sub

		Private Sub optRecursDaily_Click() Handles optRecursDaily.Click
			optRecursAnnually.Value = False
			optRecursDaily.Value = True
			optRecursMonthly.Value = False
			optRecursWeekly.Value = False
		End Sub

		Private Sub optRecursMonthly_Click() Handles optRecursMonthly.Click
			optRecursAnnually.Value = False
			optRecursDaily.Value = False
			optRecursMonthly.Value = True
			optRecursWeekly.Value = False
		End Sub

		Private Sub optRecursWeekly_Click() Handles optRecursWeekly.Click
			optRecursAnnually.Value = False
			optRecursDaily.Value = False
			optRecursMonthly.Value = False
			optRecursWeekly.Value = True
		End Sub

		Private Sub fraEventDetails_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles fraEventDetails.OnMouseDown
			Dim pt As Point = Me.GetAbsolutePosition()
			Dim lLeft As Int32 = 392 + pt.X
			Dim lTop As Int32 = 170 + pt.Y

			If lMouseY > lTop AndAlso lMouseY < lTop + 16 Then
				If lMouseX > lLeft Then
					Dim lIdx As Int32 = lMouseX - lLeft
					lIdx \= 16
					If lIdx < 11 Then
						myEventIcon = CByte(lIdx)
						Me.IsDirty = True
					End If
				End If
			End If
		End Sub

		Public Overrides Sub NewFrame()
			Dim lDay As Int32 = 0
			Dim lMonth As Int32 = 0
			Dim lYear As Int32 = 0

			If moCalendar Is Nothing Then Return
			If lstEvents Is Nothing Then Return

			moCalendar.GetSelectedDate(lDay, lMonth, lYear)

			'ok, construct our key number - mmddyyyy - we store the key num in the 3rd position (itemdata3)
			Dim lKeyNum As Int32 = CInt(lMonth.ToString("00") & lDay.ToString("00") & lYear.ToString("0000"))
			Dim bDone As Boolean = False
			While bDone = False
				bDone = True
				For X As Int32 = 0 To lstEvents.ListCount - 1
					If lstEvents.ItemData3(X) <> lKeyNum Then
						lstEvents.RemoveItem(X)
						bDone = False
						Exit For
					End If
				Next X
			End While

			'Ok, now, go thru and look forward
			If goCurrentPlayer Is Nothing = False Then
				Dim oGuild As Guild = goCurrentPlayer.oGuild
				If oGuild Is Nothing = False AndAlso oGuild.moEvents Is Nothing = False Then
					For X As Int32 = 0 To oGuild.moEvents.GetUpperBound(0)
						Dim oEvent As GuildEvent = oGuild.moEvents(X)
						If oEvent Is Nothing = False Then
							Dim dtLocal As Date = oEvent.dtStartsAt.ToLocalTime()
							If dtLocal.Day = lDay AndAlso dtLocal.Month = lMonth AndAlso dtLocal.Year = lYear Then
								'ok, found one, check our list...
								Dim bFound As Boolean = False
								For Y As Int32 = 0 To lstEvents.ListCount - 1
									If lstEvents.ItemData(Y) = oEvent.EventID Then
										Dim yTemp() As Byte = oEvent.RequestDetails()
										If yTemp Is Nothing = False Then MyBase.moUILib.SendMsgToPrimary(yTemp)
										Dim sTemp As String = oEvent.ListboxString
										If lstEvents.List(Y) <> sTemp Then lstEvents.List(Y) = sTemp
										bFound = True
										Exit For
									End If
								Next Y
								If bFound = False Then
									Dim yTemp() As Byte = oEvent.RequestDetails()
									If yTemp Is Nothing = False Then MyBase.moUILib.SendMsgToPrimary(yTemp)
									lstEvents.AddItem(oEvent.ListboxString, False)
									lstEvents.ItemData(lstEvents.NewIndex) = oEvent.EventID
									lstEvents.rcIconRectangle(lstEvents.NewIndex) = oEvent.rcIconRect
									lstEvents.ApplyIconOffset(lstEvents.NewIndex) = True
									lstEvents.IconForeColor(lstEvents.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 255, 255)
									Dim lKey As Int32 = CInt(dtLocal.Month.ToString("00") & dtLocal.Day.ToString("00") & dtLocal.Year.ToString("0000"))
									lstEvents.ItemData3(lstEvents.NewIndex) = lKey
								End If
							End If
						End If
					Next X
				End If
			End If

		End Sub

		Public Overrides Sub RenderEnd()

			If cboStartsAt Is Nothing OrElse cboEventType Is Nothing Then Return
            If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return

			If cboStartsAt.IsExpanded = False AndAlso cboEventType.IsExpanded = False Then
				Dim oTex As Texture = goResMgr.GetTexture("GuildIcons.dds", GFXResourceManager.eGetTextureType.UserInterface, "gi.pak")
				If oTex Is Nothing = False Then
					Using oSprite As New Sprite(MyBase.moUILib.oDevice)
						oSprite.Begin(SpriteFlags.AlphaBlend)
						'Presently, only 12 icons are available
						Dim lLeft As Int32 = 392 + Me.Left '410 + Me.Left
						Dim lTop As Int32 = 170 + Me.Top '232 + Me.Top

						oSprite.Draw2D(oTex, New Rectangle(0, 0, 64, 16), New Rectangle(lLeft, lTop, 64, 16), New Point(0, 0), 0, New Point(lLeft, lTop), Color.White)
						lLeft += 64
						oSprite.Draw2D(oTex, New Rectangle(0, 16, 64, 16), New Rectangle(lLeft, lTop, 64, 16), New Point(0, 0), 0, New Point(lLeft, lTop), Color.White)
						lLeft += 64
						oSprite.Draw2D(oTex, New Rectangle(0, 32, 64, 16), New Rectangle(lLeft, lTop, 64, 16), New Point(0, 0), 0, New Point(lLeft, lTop), Color.White)

						oSprite.End()
					End Using
				End If
				Dim rcDest As Rectangle = New Rectangle(392 + Me.Left + (16 * CInt(myEventIcon)) - 2, 170 + Me.Top - 2, 19, 19)
                RenderBox(rcDest, 3, muSettings.InterfaceBorderColor)
			End If

		End Sub

		Private Sub SetFromEvent(ByRef oEvent As GuildEvent)

			If oEvent Is Nothing = False Then
				txtEventTitle.Caption = oEvent.sTitle : txtEventTitle.Locked = oEvent.lPostedBy <> glPlayerID
				txtPostedBy.Caption = GetCacheObjectValue(oEvent.lPostedBy, ObjectType.ePlayer)
				txtPostedOn.Caption = oEvent.dtPostedOn.ToLocalTime.ToShortDateString & " " & oEvent.dtPostedOn.ToLocalTime.ToShortTimeString
				chkMembersCanAccept.Value = oEvent.yMembersCanAccept <> 0 : chkMembersCanAccept.Enabled = oEvent.lPostedBy = glPlayerID
				txtDetails.Locked = oEvent.lPostedBy <> glPlayerID : txtDetails.Caption = oEvent.sDetails
				chkRecurs.Enabled = oEvent.lPostedBy = glPlayerID : chkRecurs.Value = oEvent.yRecurrence <> 0
				optRecursDaily.Enabled = oEvent.lPostedBy = glPlayerID : optRecursDaily.Value = oEvent.yRecurrence = 1
				optRecursWeekly.Enabled = oEvent.lPostedBy = glPlayerID : optRecursWeekly.Value = oEvent.yRecurrence = 2
				optRecursMonthly.Enabled = oEvent.lPostedBy = glPlayerID : optRecursMonthly.Value = oEvent.yRecurrence = 3
				optRecursAnnually.Enabled = oEvent.lPostedBy = glPlayerID : optRecursAnnually.Value = oEvent.yRecurrence = 4
				cboSendAlerts.Enabled = oEvent.lPostedBy = glPlayerID : cboSendAlerts.FindComboItemData(oEvent.ySendAlerts)
				cboEventType.Enabled = oEvent.lPostedBy = glPlayerID : cboEventType.FindComboItemData(oEvent.yEventType)
				cboStartsAt.Enabled = oEvent.lPostedBy = glPlayerID
				Dim dt As Date = oEvent.dtStartsAt.ToLocalTime
				Dim lDataVal As Int32 = CInt(dt.Hour.ToString("00") & dt.Minute.ToString("00"))
				cboStartsAt.FindComboItemData(lDataVal)

				txtDuration.Caption = oEvent.lDuration.ToString
				txtDuration.Locked = oEvent.lPostedBy <> glPlayerID

				myEventIcon = oEvent.yEventIcon

				btnAccept.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.AcceptEvents)

				btnAdvanced.Enabled = oEvent.lPostedBy = glPlayerID
			Else
				txtEventTitle.Caption = "" : txtEventTitle.Locked = False
				If goCurrentPlayer Is Nothing = False Then txtPostedBy.Caption = goCurrentPlayer.PlayerName
				txtPostedOn.Caption = "Not Yet Posted"
				chkMembersCanAccept.Enabled = True : chkMembersCanAccept.Value = True
				txtDetails.Locked = False : txtDetails.Caption = "Enter Details Here"
				chkRecurs.Enabled = True : chkRecurs.Value = False
				optRecursDaily.Enabled = False : optRecursDaily.Value = False
				optRecursWeekly.Enabled = False : optRecursWeekly.Value = False
				optRecursMonthly.Enabled = False : optRecursMonthly.Value = False
				optRecursAnnually.Enabled = False : optRecursAnnually.Value = False
				cboSendAlerts.Enabled = True
				cboEventType.Enabled = True
				cboStartsAt.Enabled = True
				txtDuration.Enabled = True
				txtDuration.Locked = False

				btnAccept.Enabled = goCurrentPlayer.oGuild.HasPermission(RankPermissions.CreateEvents)
			End If
		End Sub

		Private Sub btnAdvanced_Click(ByVal sName As String) Handles btnAdvanced.Click
			If cboEventType.ListIndex > -1 Then
				If lstEvents.ListIndex < 0 Then
					MyBase.moUILib.AddNotification("You must submit the event before adding advanced configurations.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				End If

				Dim lEventID As Int32 = lstEvents.ItemData(lstEvents.NewIndex)
				If cboEventType.ItemData(cboEventType.ListIndex) = 4 Then
					'tournament
					Dim ofrm As New frmTournament(MyBase.moUILib, lEventID)
					ofrm.Visible = True
					ofrm.Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - ofrm.Width \ 2
					ofrm.Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - ofrm.Height \ 2
				ElseIf cboEventType.ItemData(cboEventType.ListIndex) = 2 Then
					'race
					Dim ofrm As New frmRaceConfig(MyBase.moUILib, lEventID)
					ofrm.Visible = True
					ofrm.Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - ofrm.Width \ 2
					ofrm.Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - ofrm.Height \ 2
				Else
					MyBase.moUILib.AddNotification("Advanced configuration is only available on race and tournament events.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				End If
			Else
				MyBase.moUILib.AddNotification("Select an event type first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			End If

		End Sub
	End Class
End Class
