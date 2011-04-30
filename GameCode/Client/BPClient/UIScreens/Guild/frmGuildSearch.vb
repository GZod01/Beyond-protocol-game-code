Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class frmGuildSearch
	Inherits UIWindow

	Private lblTitle As UILabel
	Private lnDiv1 As UILine
	Private WithEvents btnClose As UIButton
	Private WithEvents lstResults As UIListBox
	Private WithEvents btnCreateGuild As UIButton

	'fraSearchCriteria
	Private fraSearchCriteria As UIWindow
	Private lblName As UILabel
	Private txtName As UITextBox
	Private chkRecruiting As UICheckBox
	Private chkWillTrain As UICheckBox
	Private chkRecruitTutorial As UICheckBox
	Private chkProduction As UICheckBox
	Private lblSpecialities As UILabel
	Private chkResearch As UICheckBox
	Private chkDiplomacy As UICheckBox
	Private chkEspionage As UICheckBox
	Private chkCombat As UICheckBox
	Private chkTrade As UICheckBox
	Private WithEvents btnSearch As UIButton
	Private WithEvents btnReset As UIButton

	'fraInvites
	Private fraInvites As UIWindow
	Private WithEvents lstInvites As UIListBox
	Private WithEvents btnAccept As UIButton
	Private WithEvents btnReject As UIButton

	'fraSelectionDetails
	Private fraSelectionDetails As UIWindow
	Private fraIcon As UIWindow
	Private txtDetails As UITextBox
	Private WithEvents btnContact As UIButton
	Private WithEvents btnApply As UIButton

	Private mrcStatusIcons As Rectangle

	'for rendering the icon
	Private rcBack As Rectangle = Rectangle.Empty
	Private rcFore1 As Rectangle
	Private rcFore2 As Rectangle
	Private clrBack As System.Drawing.Color
	Private clrFore1 As System.Drawing.Color
	Private clrFore2 As System.Drawing.Color
	Private mlCurrIcon As Int32 = -1

	Private moSprite As Sprite
	Private moIconFore As Texture
	Private moIconBack As Texture

	Private Class GuildInvite
		Public lGuildID As Int32
		Public sName As String
		Public dtFormed As Date		'in gmt
		Public sBillboard As String
		Public iRecruitFlags As Int16
		Public lBaseGuildFlags As Int32
		Public yVoteWeightType As Byte
		Public yGuildTaxRateInterval As Byte
		Public lIcon As Int32

		Public Function GetDetailsText() As String
			Dim oSB As New System.Text.StringBuilder
			oSB.AppendLine(sName & " has invited you to join them.")
			oSB.AppendLine("To Accept this invitation, click the Accept button above.")
			oSB.AppendLine()

			If dtFormed <> Date.MinValue Then
				oSB.AppendLine(sName & " was formed on " & dtFormed.ToLocalTime.ToString)
			Else
				oSB.AppendLine(sName & " is in the forming stages.")
			End If
			oSB.AppendLine()

			If (iRecruitFlags And eiRecruitmentFlags.GuildRecruiting) <> 0 Then
				oSB.AppendLine("This guild is currently recruiting.")
				If (iRecruitFlags And (eiRecruitmentFlags.RecruitDiplomacy Or eiRecruitmentFlags.RecruitEspionage Or eiRecruitmentFlags.RecruitMilitary Or eiRecruitmentFlags.RecruitProduction Or eiRecruitmentFlags.RecruitResearch Or eiRecruitmentFlags.RecruitTrade)) <> 0 Then
					oSB.AppendLine("This guild is looking for players particularly good at:")
					If (iRecruitFlags And eiRecruitmentFlags.RecruitDiplomacy) <> 0 Then
						oSB.AppendLine("  Diplomacy")
					End If
					If (iRecruitFlags And eiRecruitmentFlags.RecruitEspionage) <> 0 Then
						oSB.AppendLine("  Espionage")
					End If
					If (iRecruitFlags And eiRecruitmentFlags.RecruitMilitary) <> 0 Then
						oSB.AppendLine("  Military/Combat")
					End If
					If (iRecruitFlags And eiRecruitmentFlags.RecruitProduction) <> 0 Then
						oSB.AppendLine("  Production")
					End If
					If (iRecruitFlags And eiRecruitmentFlags.RecruitResearch) <> 0 Then
						oSB.AppendLine("  Research")
					End If
					If (iRecruitFlags And eiRecruitmentFlags.RecruitTrade) <> 0 Then
						oSB.AppendLine("  Trade")
					End If
				End If
			End If

			oSB.AppendLine()
			If (lBaseGuildFlags And elGuildFlags.AutomaticTradeAgreements) <> 0 Then
				oSB.AppendLine("Members enjoy automatic trade agreements with other members" & vbCrLf)
			End If
			If (lBaseGuildFlags And elGuildFlags.RequirePeaceBetweenMembers) <> 0 Then
				oSB.AppendLine("Peace is upheld between members" & vbCrLf)
			End If
			If (lBaseGuildFlags And elGuildFlags.ShareUnitVision) <> 0 Then
				oSB.AppendLine("Members share unit and facility vision with other members" & vbCrLf)
			End If
            'If (lBaseGuildFlags And elGuildFlags.UnifiedForeignPolicy) <> 0 Then
            '	oSB.AppendLine("Utilizes a unified foreign policy" & vbCrLf)
            'End If
			Select Case yVoteWeightType
				Case eyVoteWeightType.AgeOfPlayer
					oSB.AppendLine("Voting is based off the time of membership")
				Case eyVoteWeightType.PopulationBased
					oSB.AppendLine("Voting is based off of Population")
				Case eyVoteWeightType.RankBased
					oSB.AppendLine("Voting is based off of Rank")
				Case eyVoteWeightType.StaticValue
					oSB.AppendLine("Voting is static and means 1 vote per 1 member")
			End Select

			Select Case yGuildTaxRateInterval
				Case eyGuildInterval.Annually
					oSB.AppendLine("Taxes are due annually.")
				Case eyGuildInterval.Daily
					oSB.AppendLine("Taxes are due daily.")
				Case eyGuildInterval.EveryTwoMonths
					oSB.AppendLine("Taxes are due every two months.")
				Case eyGuildInterval.Monthly
					oSB.AppendLine("Taxes are due monthly.")
				Case eyGuildInterval.Quarterly
					oSB.AppendLine("Taxes are due quarterly.")
				Case eyGuildInterval.SemiAnnually
					oSB.AppendLine("Taxes are due every six months.")
				Case eyGuildInterval.SemiMonthly
					oSB.AppendLine("Taxes are due twice per month.")
				Case eyGuildInterval.Weekly
					oSB.AppendLine("Taxes are due weekly.")
			End Select

			Return oSB.ToString
		End Function
	End Class
	Private moInvites() As GuildInvite
	Private mlInviteUB As Int32 = -1
	Private Class GuildSearchResult
		Public lGuildID As Int32
		Public sName As String
		Public dtFormed As Date
		Public sBillboard As String
		Public iRecruitFlags As Int16
		Public lIcon As Int32

		Public Function GetDetailsText() As String
			Dim oSB As New System.Text.StringBuilder
			oSB.AppendLine("Name: " & sName)
			oSB.AppendLine()

			If dtFormed <> Date.MinValue Then
				oSB.AppendLine(sName & " was formed on " & dtFormed.ToLocalTime.ToString)
			Else
				oSB.AppendLine(sName & " is in the forming stages.")
			End If
			oSB.AppendLine()

			If (iRecruitFlags And eiRecruitmentFlags.GuildRecruiting) <> 0 Then
				oSB.AppendLine("This guild is currently recruiting.")
				If (iRecruitFlags And (eiRecruitmentFlags.RecruitDiplomacy Or eiRecruitmentFlags.RecruitEspionage Or eiRecruitmentFlags.RecruitMilitary Or eiRecruitmentFlags.RecruitProduction Or eiRecruitmentFlags.RecruitResearch Or eiRecruitmentFlags.RecruitTrade)) <> 0 Then
					oSB.AppendLine("This guild is looking for players particularly good at:")
					If (iRecruitFlags And eiRecruitmentFlags.RecruitDiplomacy) <> 0 Then
						oSB.AppendLine("  Diplomacy")
					End If
					If (iRecruitFlags And eiRecruitmentFlags.RecruitEspionage) <> 0 Then
						oSB.AppendLine("  Espionage")
					End If
					If (iRecruitFlags And eiRecruitmentFlags.RecruitMilitary) <> 0 Then
						oSB.AppendLine("  Military/Combat")
					End If
					If (iRecruitFlags And eiRecruitmentFlags.RecruitProduction) <> 0 Then
						oSB.AppendLine("  Production")
					End If
					If (iRecruitFlags And eiRecruitmentFlags.RecruitResearch) <> 0 Then
						oSB.AppendLine("  Research")
					End If
					If (iRecruitFlags And eiRecruitmentFlags.RecruitTrade) <> 0 Then
						oSB.AppendLine("  Trade")
					End If
				End If
			End If
 

			Return oSB.ToString
		End Function
	End Class
	Private moResults() As GuildSearchResult
	Private mlResultUB As Int32 = -1

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.GuildSearchOpen)

		'frmSearchGuilds initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eGuildSearch
            .ControlName = "frmGuildSearch"
            .Left = 297
            .Top = 65
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
			.Top = 0
			.Width = 127
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Search For Guilds"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTitle, UIControl))

		'lnDiv1 initial props
		lnDiv1 = New UILine(oUILib)
		With lnDiv1
			.ControlName = "lnDiv1"
			.Left = 1
			.Top = 25
			.Width = 511
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv1, UIControl))

		'btnClose initial props
		btnClose = New UIButton(oUILib)
		With btnClose
			.ControlName = "btnClose"
			.Left = 487
			.Top = 1
			.Width = 24
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "X"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnClose, UIControl))

		'lstResults initial props
		lstResults = New UIListBox(oUILib)
		With lstResults
			.ControlName = "lstResults"
			.Left = 5
			.Top = 200
			.Width = 240
			.Height = 275
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
		End With
		Me.AddChild(CType(lstResults, UIControl))

		'btnCreateGuild initial props
		btnCreateGuild = New UIButton(oUILib)
		With btnCreateGuild
			.ControlName = "btnCreateGuild"
			.Left = 5
			.Top = 480
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Create Guild"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnCreateGuild, UIControl))


		'============ fraSearchCriteria ==============
		'fraSearchCriteria initial props
		fraSearchCriteria = New UIWindow(oUILib)
		With fraSearchCriteria
			.ControlName = "fraSearchCriteria"
			.Left = 5
			.Top = 35
			.Width = 240
			.Height = 160
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			.BorderLineWidth = 2
			.Caption = "Search Criteria"
		End With
		Me.AddChild(CType(fraSearchCriteria, UIControl))

		'lblName initial props
		lblName = New UILabel(oUILib)
		With lblName
			.ControlName = "lblName"
			.Left = 5
			.Top = 10
			.Width = 30
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Name:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraSearchCriteria.AddChild(CType(lblName, UIControl))

		'txtName initial props
		txtName = New UITextBox(oUILib)
		With txtName
			.ControlName = "txtName"
			.Left = 40
			.Top = 10
			.Width = 192
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 0
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		fraSearchCriteria.AddChild(CType(txtName, UIControl))

		'chkRecruiting initial props
		chkRecruiting = New UICheckBox(oUILib)
		With chkRecruiting
			.ControlName = "chkRecruiting"
			.Left = 20
			.Top = 35
			.Width = 67
			.Height = 13
			.Enabled = True
			.Visible = True
			.Caption = "Recruiting"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraSearchCriteria.AddChild(CType(chkRecruiting, UIControl))

		'chkWillTrain initial props
		chkWillTrain = New UICheckBox(oUILib)
		With chkWillTrain
			.ControlName = "chkWillTrain"
			.Left = 20
			.Top = 50
			.Width = 152
			.Height = 13
			.Enabled = True
			.Visible = True
			.Caption = "Willing to Train New Players"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraSearchCriteria.AddChild(CType(chkWillTrain, UIControl))

		'chkRecruitTutorial initial props
		chkRecruitTutorial = New UICheckBox(oUILib)
		With chkRecruitTutorial
			.ControlName = "chkRecruitTutorial"
			.Left = 125
			.Top = 35
			.Width = 105
			.Height = 13
			.Enabled = True
			.Visible = True
			.Caption = "Recruiting Tutorial"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraSearchCriteria.AddChild(CType(chkRecruitTutorial, UIControl))

		'chkProduction initial props
		chkProduction = New UICheckBox(oUILib)
		With chkProduction
			.ControlName = "chkProduction"
			.Left = 20
			.Top = 85
			.Width = 70
			.Height = 13
			.Enabled = True
			.Visible = True
			.Caption = "Production"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraSearchCriteria.AddChild(CType(chkProduction, UIControl))

		'lblSpecialities initial props
		lblSpecialities = New UILabel(oUILib)
		With lblSpecialities
			.ControlName = "lblSpecialities"
			.Left = 5
			.Top = 65
			.Width = 120
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Looking Players Good At:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		fraSearchCriteria.AddChild(CType(lblSpecialities, UIControl))

		'chkResearch initial props
		chkResearch = New UICheckBox(oUILib)
		With chkResearch
			.ControlName = "chkResearch"
			.Left = 20
			.Top = 100
			.Width = 65
			.Height = 13
			.Enabled = True
			.Visible = True
			.Caption = "Research"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraSearchCriteria.AddChild(CType(chkResearch, UIControl))

		'chkDiplomacy initial props
		chkDiplomacy = New UICheckBox(oUILib)
		With chkDiplomacy
			.ControlName = "chkDiplomacy"
			.Left = 20
			.Top = 115
			.Width = 68
			.Height = 13
			.Enabled = True
			.Visible = True
			.Caption = "Diplomacy"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraSearchCriteria.AddChild(CType(chkDiplomacy, UIControl))

		'chkEspionage initial props
		chkEspionage = New UICheckBox(oUILib)
		With chkEspionage
			.ControlName = "chkEspionage"
			.Left = 125
			.Top = 85
			.Width = 69
			.Height = 13
			.Enabled = True
			.Visible = True
			.Caption = "Espionage"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraSearchCriteria.AddChild(CType(chkEspionage, UIControl))

		'chkCombat initial props
		chkCombat = New UICheckBox(oUILib)
		With chkCombat
			.ControlName = "chkCombat"
			.Left = 125
			.Top = 100
			.Width = 55
			.Height = 13
			.Enabled = True
			.Visible = True
			.Caption = "Combat"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraSearchCriteria.AddChild(CType(chkCombat, UIControl))

		'chkTrade initial props
		chkTrade = New UICheckBox(oUILib)
		With chkTrade
			.ControlName = "chkTrade"
			.Left = 125
			.Top = 115
			.Width = 47
			.Height = 13
			.Enabled = True
			.Visible = True
			.Caption = "Trade"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		fraSearchCriteria.AddChild(CType(chkTrade, UIControl))

		'btnSearch initial props
		btnSearch = New UIButton(oUILib)
		With btnSearch
			.ControlName = "btnSearch"
			.Left = 5
			.Top = fraSearchCriteria.Height - 25
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Search"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraSearchCriteria.AddChild(CType(btnSearch, UIControl))

		'btnReset initial props
		btnReset = New UIButton(oUILib)
		With btnReset
			.ControlName = "btnReset"
			.Left = fraSearchCriteria.Width - 105
			.Top = btnSearch.Top
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Reset"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraSearchCriteria.AddChild(CType(btnReset, UIControl))


		'================== fraInvites ===================
		'fraInvites initial props
		fraInvites = New UIWindow(oUILib)
		With fraInvites
			.ControlName = "fraInvites"
			.Left = 250
			.Top = 35
			.Width = 255
			.Height = 160
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			.BorderLineWidth = 2
			.Caption = "Invitations to Join Guilds"
		End With
		Me.AddChild(CType(fraInvites, UIControl))

		'lstInvites initial props
		lstInvites = New UIListBox(oUILib)
		With lstInvites
			.ControlName = "lstInvites"
			.Left = 5
			.Top = 10
			.Width = 245
			.Height = 100
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
		End With
		fraInvites.AddChild(CType(lstInvites, UIControl))

		'btnReject initial props
		btnReject = New UIButton(oUILib)
		With btnReject
			.ControlName = "btnReject"
			.Left = 10
			.Top = 125
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Reject"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraInvites.AddChild(CType(btnReject, UIControl))

		'btnAccept initial props
		btnAccept = New UIButton(oUILib)
		With btnAccept
			.ControlName = "btnAccept"
			.Left = 145
			.Top = 125
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Accept"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraInvites.AddChild(CType(btnAccept, UIControl))


		'================ fraSelectionDetails ================
		'fraSelectionDetails initial props
		fraSelectionDetails = New UIWindow(oUILib)
		With fraSelectionDetails
			.ControlName = "fraSelectionDetails"
			.Left = 250
			.Top = 205
			.Width = 255
			.Height = 302
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			.BorderLineWidth = 2
			.Caption = "Selection Details"
		End With
		Me.AddChild(CType(fraSelectionDetails, UIControl))

		'fraIcon initial props
		fraIcon = New UIWindow(oUILib)
		With fraIcon
			.ControlName = "fraIcon"
			.Left = 5
			.Top = 10
			.Width = 67
			.Height = 67
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			.BorderLineWidth = 2
		End With
		fraSelectionDetails.AddChild(CType(fraIcon, UIControl))

		'txtDetails initial props
		txtDetails = New UITextBox(oUILib)
		With txtDetails
			.ControlName = "txtDetails"
			.Left = 5
			.Top = 100
			.Width = 245
			.Height = 195
			.Enabled = True
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
			.Locked = True
			.MultiLine = True
		End With
		fraSelectionDetails.AddChild(CType(txtDetails, UIControl))

		'btnContact initial props
		btnContact = New UIButton(oUILib)
		With btnContact
			.ControlName = "btnContact"
			.Left = 80
			.Top = 20
			.Width = 170
			.Height = 24
			.Enabled = False
			.Visible = True
			.Caption = "Contact An Officer"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraSelectionDetails.AddChild(CType(btnContact, UIControl))

		'btnApply initial props
		btnApply = New UIButton(oUILib)
		With btnApply
			.ControlName = "btnApply"
			.Left = 80
			.Top = 50
			.Width = 170
			.Height = 24
			.Enabled = False
			.Visible = True
			.Caption = "Apply for Membership"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		fraSelectionDetails.AddChild(CType(btnApply, UIControl))

		mrcStatusIcons = New Rectangle(fraSelectionDetails.Left + 5, fraSelectionDetails.Top + 75, 245, 32)

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)

		Dim yMsg(1) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eGetGuildInvites).CopyTo(yMsg, 0)
		MyBase.moUILib.SendMsgToPrimary(yMsg)

	End Sub

	Private Sub btnAccept_Click(ByVal sName As String) Handles btnAccept.Click
		If lstInvites.ListIndex > -1 Then
			If btnAccept.Caption.ToUpper = "ACCEPT" Then
				btnAccept.Caption = "Confirm"
			Else
				'confirming... send our acceptance
				Dim lID As Int32 = lstInvites.ItemData(lstInvites.ListIndex)
				SendGuildMemberStatusMsg(glPlayerID, GuildMemberState.Approved, lID)
			End If
		End If
	End Sub

	Private Sub btnApply_Click(ByVal sName As String) Handles btnApply.Click
		If lstResults.ListIndex > -1 Then
			Dim lID As Int32 = lstResults.ItemData(lstResults.ListIndex)

			Dim oResult As GuildSearchResult = Nothing
			For X As Int32 = 0 To mlResultUB
				If moResults(X) Is Nothing = False AndAlso moResults(X).lGuildID = lID Then
					oResult = moResults(X)
					Exit For
				End If
			Next X
			If oResult Is Nothing Then Return
			If (oResult.iRecruitFlags And eiRecruitmentFlags.GuildRecruiting) = 0 Then
				MyBase.moUILib.AddNotification(oResult.sName & " is not currently recruiting.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If

			If btnApply.Caption.ToUpper = "ACCEPT" Then
				btnApply.Caption = "Confirm"
			Else
                'confirming... send our acceptance
                btnApply.Enabled = False
				SendGuildMemberStatusMsg(glPlayerID, GuildMemberState.Applied, lID)
			End If
		End If
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnContact_Click(ByVal sName As String) Handles btnContact.Click

		Dim lID As Int32 = -1
		If lstInvites.ListIndex > -1 Then
			lID = lstInvites.ItemData(lstInvites.ListIndex)
		ElseIf lstResults.ListIndex > -1 Then
			lID = lstResults.ItemData(lstResults.ListIndex)
		Else : Return
		End If
		If lID = -1 Then Return

		Dim yMsg(5) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eRequestContactWithGuild).CopyTo(yMsg, 0)
		System.BitConverter.GetBytes(lID).CopyTo(yMsg, 2)
		MyBase.moUILib.SendMsgToPrimary(yMsg)
		MyBase.moUILib.AddNotification("Contact Request Submitted", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
	End Sub

	Private Sub btnCreateGuild_Click(ByVal sName As String) Handles btnCreateGuild.Click
		Dim ofrm As New frmGuildSetup(MyBase.moUILib)
		MyBase.moUILib.RemoveWindow(Me.ControlName)
		ofrm.Visible = True
	End Sub

	Private Sub btnReject_Click(ByVal sName As String) Handles btnReject.Click
		If lstInvites.ListIndex > -1 Then
			If btnReject.Caption.ToUpper = "ACCEPT" Then
				btnReject.Caption = "Confirm"
			Else
				'confirming... send our acceptance
				Dim lID As Int32 = lstInvites.ItemData(lstInvites.ListIndex)
				SendGuildMemberStatusMsg(glPlayerID, GuildMemberState.Rejected, lID)
			End If
		End If
	End Sub

	Private Sub btnReset_Click(ByVal sName As String) Handles btnReset.Click
		txtName.Caption = ""
		chkRecruiting.Value = False
		chkWillTrain.Value = False
		chkRecruitTutorial.Value = False
		chkProduction.Value = False
		chkResearch.Value = False
		chkDiplomacy.Value = False
		chkEspionage.Value = False
		chkCombat.Value = False
		chkTrade.Value = False
	End Sub

	Private Sub btnSearch_Click(ByVal sName As String) Handles btnSearch.Click
		btnSearch.Enabled = False
		btnApply.Enabled = False
		btnContact.Enabled = False
		lstResults.Clear()
		If btnSearch.Caption.ToUpper = "SEARCH" Then
			'validate our search criteria

			Dim sGuildName As String = txtName.Caption
			Dim lSearchCriteria As Int32 = 0

			If chkRecruiting.Value = True Then lSearchCriteria = lSearchCriteria Or eiRecruitmentFlags.GuildRecruiting
			If chkWillTrain.Value = True Then lSearchCriteria = lSearchCriteria Or eiRecruitmentFlags.WillTrain
			If chkRecruitTutorial.Value = True Then lSearchCriteria = lSearchCriteria Or eiRecruitmentFlags.RecruitingTutorial
			If chkProduction.Value = True Then lSearchCriteria = lSearchCriteria Or eiRecruitmentFlags.RecruitProduction
			If chkResearch.Value = True Then lSearchCriteria = lSearchCriteria Or eiRecruitmentFlags.RecruitResearch
			If chkDiplomacy.Value = True Then lSearchCriteria = lSearchCriteria Or eiRecruitmentFlags.RecruitDiplomacy
			If chkEspionage.Value = True Then lSearchCriteria = lSearchCriteria Or eiRecruitmentFlags.RecruitEspionage
			If chkCombat.Value = True Then lSearchCriteria = lSearchCriteria Or eiRecruitmentFlags.RecruitMilitary
			If chkTrade.Value = True Then lSearchCriteria = lSearchCriteria Or eiRecruitmentFlags.RecruitTrade

			If sGuildName = "" AndAlso lSearchCriteria = 0 Then
				MyBase.moUILib.AddNotification("You must enter a search criteria.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				btnSearch.Enabled = True
				Return
			End If
			If sGuildName.Length = 1 OrElse sGuildName.Length = 2 Then
				MyBase.moUILib.AddNotification("You must enter at least 3 characters to search by name.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				btnSearch.Enabled = True
				Return
			End If

			'search
			btnSearch.Caption = "Cancel"
			'Now, send off our search
			Dim yMsg(55) As Byte
			Dim lPos As Int32 = 0

			System.BitConverter.GetBytes(GlobalMessageCode.eSearchGuilds).CopyTo(yMsg, 0)
			System.BitConverter.GetBytes(lSearchCriteria).CopyTo(yMsg, 2)
			System.Text.ASCIIEncoding.ASCII.GetBytes(sGuildName).CopyTo(yMsg, 6)
			MyBase.moUILib.SendMsgToPrimary(yMsg)
		Else
			'cancel search
			btnSearch.Caption = "Search"
		End If
		btnSearch.Enabled = True
	End Sub

	Private Sub lstInvites_ItemClick(ByVal lIndex As Integer) Handles lstInvites.ItemClick
		If lIndex > -1 Then
			lstResults.ListIndex = -1

			btnAccept.Caption = "Accept"
			btnReject.Caption = "Reject"
			btnApply.Caption = "Apply"

			btnApply.Enabled = False
			btnContact.Enabled = True

			Dim lID As Int32 = lstInvites.ItemData(lstInvites.ListIndex)
			'ok, find our invite
			Dim oInvite As GuildInvite = Nothing
			For X As Int32 = 0 To mlInviteUB
				If moInvites(X) Is Nothing = False Then
					If moInvites(X).lGuildID = lID Then
						oInvite = moInvites(X)
						Exit For
					End If
				End If
			Next X
			If oInvite Is Nothing = False Then
				SetIcon(oInvite.lIcon)
				txtDetails.Caption = oInvite.GetDetailsText()
				btnApply.Enabled = False
				btnApply.ToolTipText = "You have been invited to join this guild. Click Accept or Reject above."
			End If
		End If
	End Sub

	Private Sub lstResults_ItemClick(ByVal lIndex As Integer) Handles lstResults.ItemClick
		If lIndex > -1 Then
			lstInvites.ListIndex = -1

			btnApply.Enabled = True
			btnContact.Enabled = True

			Dim lID As Int32 = lstResults.ItemData(lstResults.ListIndex)
			Dim oResult As GuildSearchResult = Nothing
			For X As Int32 = 0 To mlResultUB
				If moResults(X) Is Nothing = False Then
					If moResults(X).lGuildID = lID Then
						oResult = moResults(X)
						Exit For
					End If
				End If
			Next X
			If oResult Is Nothing = False Then
				SetIcon(oResult.lIcon)
				txtDetails.Caption = oResult.GetDetailsText()
				btnApply.Enabled = True
				btnApply.ToolTipText = "Apply to join the guild."
			End If
		End If
	End Sub

	Public Sub HandleSearchResults(ByRef yData() As Byte)
		Dim lPos As Int32 = 2	'for msgcode
		Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim lUB As Int32 = lCnt - 1
		Dim oResults(lUB) As GuildSearchResult

		For X As Int32 = 0 To lUB
			Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim sName As String = GetStringFromBytes(yData, lPos, 50) : lPos += 50
			Dim lFormed As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim iRecruit As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			Dim lIcon As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim lBillboardLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim sBillboard As String = ""
            If lBillboardLen <> 0 Then
                sBillboard = GetStringFromBytes(yData, lPos, lBillboardLen) : lPos += lBillboardLen
            End If

			oResults(X) = New GuildSearchResult
			With oResults(X)
				If lFormed <> 0 Then .dtFormed = Date.SpecifyKind(GetDateFromNumber(lFormed), DateTimeKind.Utc) Else .dtFormed = Date.MinValue
				.iRecruitFlags = iRecruit
				.lGuildID = lID
				.sBillboard = sBillboard
				.sName = sName
				.lIcon = lIcon
			End With
		Next X

		If btnSearch.Caption.ToUpper = "CANCEL" Then
			mlResultUB = -1
			moResults = oResults
			mlResultUB = lUB

			btnSearch.Caption = "Search"

			If lCnt = 0 Then
				MyBase.moUILib.AddNotification("Your search returned 0 results.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			End If

			lstResults.Clear()
			For X As Int32 = 0 To mlResultUB
				lstResults.AddItem(moResults(X).sName, False)
				lstResults.ItemData(lstResults.NewIndex) = moResults(X).lGuildID
			Next X

            mlForceRefresh = glCurrentCycle + 15
		End If
	End Sub

	Public Sub HandleInvitedGuildList(ByRef yData() As Byte)
		Dim lPos As Int32 = 6	'for msgcode and playerid
		Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

		Dim lUB As Int32 = lCnt - 1
		Dim oInvites(lUB) As GuildInvite

		For X As Int32 = 0 To lUB
			Dim lID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim sName As String = GetStringFromBytes(yData, lPos, 50) : lPos += 50
			Dim lFormed As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim iRecruit As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
			Dim lBaseGuildFlags As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			Dim yVoteWeightType As Byte = yData(lPos) : lPos += 1
			Dim yTaxInterval As Byte = yData(lPos) : lPos += 1
			Dim lIcon As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			'Dim lBillboardLen As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			'Dim sBillboard As String = GetStringFromBytes(yData, lPos, lBillboardLen) : lPos += lBillboardLen

			oInvites(X) = New GuildInvite
			With oInvites(X)
				If lFormed <> 0 Then .dtFormed = Date.SpecifyKind(GetDateFromNumber(lFormed), DateTimeKind.Utc) Else .dtFormed = Date.MinValue
				.iRecruitFlags = iRecruit
				.lBaseGuildFlags = lBaseGuildFlags
				.lGuildID = lID
				.sBillboard = "" 'sBillboard
				.sName = sName
				.yGuildTaxRateInterval = yTaxInterval
				.yVoteWeightType = yVoteWeightType
				.lIcon = lIcon
			End With
		Next X

		mlInviteUB = -1
		moInvites = oInvites
		mlInviteUB = lUB

		lstInvites.Clear()
		For X As Int32 = 0 To mlInviteUB
			If moInvites(X) Is Nothing = False Then
				lstInvites.AddItem(moInvites(X).sName, False)
				lstInvites.ItemData(lstInvites.NewIndex) = moInvites(X).lGuildID
			End If
		Next X
	End Sub

	Private Sub SetIcon(ByVal lIcon As Int32)
		Dim yBackImg As Byte
		Dim yBackClr As Byte
		Dim yFore1Img As Byte
		Dim yFore1Clr As Byte
		Dim yFore2Img As Byte
		Dim yFore2Clr As Byte

		PlayerIconManager.FillIconValues(lIcon, yBackImg, yBackClr, yFore1Img, yFore1Clr, yFore2Img, yFore2Clr)

		rcBack = PlayerIconManager.ReturnImageRectangle(yBackImg, True)
		rcFore1 = PlayerIconManager.ReturnImageRectangle(yFore1Img, False)
		rcFore2 = PlayerIconManager.ReturnImageRectangle(yFore2Img, False)

		clrBack = PlayerIconManager.GetColorValue(yBackClr)
		clrFore1 = PlayerIconManager.GetColorValue(yFore1Clr)
		clrFore2 = PlayerIconManager.GetColorValue(yFore2Clr)

		mlCurrIcon = lIcon
		Me.IsDirty = True
	End Sub

	Private Sub SendGuildMemberStatusMsg(ByVal lMemberID As Int32, ByVal lStatusUpdate As Int32, ByVal lGuildID As Int32)
		Dim yMsg(13) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eGuildMemberStatus).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(lGuildID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lMemberID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lStatusUpdate).CopyTo(yMsg, lPos) : lPos += 4
		MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub

    Private mlForceRefresh As Int32 = -1
    Private Sub frmGuildSearch_OnNewFrame() Handles Me.OnNewFrame
        If mlForceRefresh <> -1 AndAlso mlForceRefresh < glCurrentCycle Then
            mlForceRefresh = -1
            Me.IsDirty = True
        End If
    End Sub

	Private Sub frmGuildSearch_OnRenderEnd() Handles Me.OnRenderEnd
		If mlCurrIcon <> -1 Then
			If rcBack = Rectangle.Empty Then SetIcon(mlCurrIcon)

			If moSprite Is Nothing OrElse moSprite.Disposed = True Then
				Device.IsUsingEventHandlers = False
				moSprite = New Sprite(MyBase.moUILib.oDevice)
				Device.IsUsingEventHandlers = True
			End If
			If moIconFore Is Nothing OrElse moIconFore.Disposed = True Then
				Device.IsUsingEventHandlers = False
				moIconFore = goResMgr.GetTexture("DipPlayerFore.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
				Device.IsUsingEventHandlers = True
			End If
			If moIconBack Is Nothing OrElse moIconBack.Disposed = True Then
				Device.IsUsingEventHandlers = False
				moIconBack = goResMgr.GetTexture("DipPlayerBack.bmp", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
				Device.IsUsingEventHandlers = True
			End If

			moSprite.Begin(SpriteFlags.AlphaBlend)
			Try
				Dim rcDest As Rectangle = New Rectangle(0, 0, 64, 64)
				Dim ptDest As Point = fraIcon.GetAbsolutePosition()
				ptDest.X += 2
				ptDest.Y += 2

				moSprite.Draw2D(moIconBack, rcBack, rcDest, ptDest, clrBack)
				moSprite.Draw2D(moIconFore, rcFore1, rcDest, ptDest, clrFore1)
				moSprite.Draw2D(moIconFore, rcFore2, rcDest, ptDest, clrFore2)
			Catch
				'do nothing?
			End Try
			moSprite.End()
		End If
	End Sub

	Protected Overrides Sub Finalize()
		If moSprite Is Nothing = False Then moSprite.Dispose()
		moSprite = Nothing
		moIconFore = Nothing
		moIconBack = Nothing
		MyBase.Finalize()
	End Sub
End Class