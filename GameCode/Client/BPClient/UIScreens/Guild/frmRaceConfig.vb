Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmRaceConfig
	Inherits UIWindow

	Private lblTitle As UILabel
	Private lnDiv1 As UILine
	Private lblEntryFee As UILabel
	Private txtEntryFee As UITextBox
	Private lblHullSize As UILabel
	Private txtMinHull As UITextBox
	Private txtMaxHull As UITextBox
	Private lblMinRacers As UILabel
	Private txtMinRacers As UITextBox
	Private lblFirstPlace As UILabel
	Private txtFirstPlace As UITextBox
	Private lblSecondplace As UILabel
	Private lblThirdPlace As UILabel
	Private txtSecondplace As UITextBox
	Private txtThirdPlace As UITextBox
	Private lblGuildTake As UILabel
	Private txtGuildTake As UITextBox
	Private lblCourse As UILabel
	Private lblLaps As UILabel
	Private txtLaps As UITextBox
	Private lblPotSplit As UILabel

	Private WithEvents lstCourse As UIListBox
	Private WithEvents btnAddWP As UIButton
	Private WithEvents btnRemoveWP As UIButton
	Private WithEvents chkGroundOnly As UICheckBox
	Private WithEvents btnSave As UIButton
	Private WithEvents btnClose As UIButton

	Private Structure uWaypoint
		Public lX As Int32
		Public lZ As Int32
		Public lEnvirID As Int32
		Public iEnvirTypeID As Int16
	End Structure
	Private muWPs() As uWaypoint
	Private mlWPUB As Int32 = -1

	Private mlEventID As Int32

	Public Sub New(ByRef oUILib As UILib, ByVal lEventID As Int32)
		MyBase.New(oUILib)

		mlEventID = lEventID

		Dim oEvent As GuildEvent = Nothing
		If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
			oEvent = goCurrentPlayer.oGuild.GetEvent(lEventID)
		End If
		Dim bEnabled As Boolean = True
		If oEvent Is Nothing = False Then
			bEnabled = oEvent.lPostedBy = glPlayerID
		End If

		'frmRaceConfig initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eRaceConfig
            .ControlName = "frmRaceConfig"
            .Left = 347
            .Top = 193
            .Width = 512
            .Height = 256
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With

		'lblTitle initial props
		lblTitle = New UILabel(oUILib)
		With lblTitle
			.ControlName = "lblTitle"
			.Left = 5
			.Top = 5
			.Width = 155
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Race Configuration"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTitle, UIControl))

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

		'lblEntryFee initial props
		lblEntryFee = New UILabel(oUILib)
		With lblEntryFee
			.ControlName = "lblEntryFee"
			.Left = 10
			.Top = 40
			.Width = 70
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Entry Fee:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblEntryFee, UIControl))

		'txtEntryFee initial props
		txtEntryFee = New UITextBox(oUILib)
		With txtEntryFee
			.ControlName = "txtEntryFee"
			.Left = 100
			.Top = 42
			.Width = 135
			.Height = 22
			.Enabled = bEnabled
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 18
			.BorderColor = muSettings.InterfaceBorderColor
			.bNumericOnly = True
		End With
		Me.AddChild(CType(txtEntryFee, UIControl))

		'lblHullSize initial props
		lblHullSize = New UILabel(oUILib)
		With lblHullSize
			.ControlName = "lblHullSize"
			.Left = 10
			.Top = 70
			.Width = 60
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Hullsize:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblHullSize, UIControl))

		'txtMinHull initial props
		txtMinHull = New UITextBox(oUILib)
		With txtMinHull
			.ControlName = "txtMinHull"
			.Left = 100
			.Top = 72
			.Width = 65
			.Height = 22
			.Enabled = bEnabled
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 8
			.BorderColor = muSettings.InterfaceBorderColor
			.bNumericOnly = True
		End With
		Me.AddChild(CType(txtMinHull, UIControl))

		'txtMaxHull initial props
		txtMaxHull = New UITextBox(oUILib)
		With txtMaxHull
			.ControlName = "txtMaxHull"
			.Left = 170
			.Top = 72
			.Width = 65
			.Height = 22
			.Enabled = bEnabled
			.Visible = True
			.Caption = ""
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 8
			.bNumericOnly = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(txtMaxHull, UIControl))

		'lblMinRacers initial props
		lblMinRacers = New UILabel(oUILib)
		With lblMinRacers
			.ControlName = "lblMinRacers"
			.Left = 10
			.Top = 120
			.Width = 85
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Min. Racers:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblMinRacers, UIControl))

		'txtMinRacers initial props
		txtMinRacers = New UITextBox(oUILib)
		With txtMinRacers
			.ControlName = "txtMinRacers"
			.Left = 100
			.Top = 122
			.Width = 40
			.Height = 22
			.Enabled = bEnabled
			.Visible = True
			.Caption = "3"
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
		Me.AddChild(CType(txtMinRacers, UIControl))

		'lblFirstPlace initial props
		lblFirstPlace = New UILabel(oUILib)
		With lblFirstPlace
			.ControlName = "lblFirstPlace"
			.Left = 10
			.Top = 170
			.Width = 24
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "1st:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblFirstPlace, UIControl))

		'txtFirstPlace initial props
		txtFirstPlace = New UITextBox(oUILib)
		With txtFirstPlace
			.ControlName = "txtFirstPlace"
			.Left = 40
			.Top = 172
			.Width = 35
			.Height = 22
			.Enabled = bEnabled
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
		Me.AddChild(CType(txtFirstPlace, UIControl))

		'lblSecondplace initial props
		lblSecondplace = New UILabel(oUILib)
		With lblSecondplace
			.ControlName = "lblSecondplace"
			.Left = 85
			.Top = 170
			.Width = 30
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "2nd:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblSecondplace, UIControl))

		'lblThirdPlace initial props
		lblThirdPlace = New UILabel(oUILib)
		With lblThirdPlace
			.ControlName = "lblThirdPlace"
			.Left = 165
			.Top = 170
			.Width = 27
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "3rd:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblThirdPlace, UIControl))

		'txtSecondplace initial props
		txtSecondplace = New UITextBox(oUILib)
		With txtSecondplace
			.ControlName = "txtSecondplace"
			.Left = 120
			.Top = 172
			.Width = 35
			.Height = 22
			.Enabled = bEnabled
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
		Me.AddChild(CType(txtSecondplace, UIControl))

		'txtThirdPlace initial props
		txtThirdPlace = New UITextBox(oUILib)
		With txtThirdPlace
			.ControlName = "txtThirdPlace"
			.Left = 200
			.Top = 172
			.Width = 35
			.Height = 22
			.Enabled = bEnabled
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
		Me.AddChild(CType(txtThirdPlace, UIControl))

		'lblGuildTake initial props
		lblGuildTake = New UILabel(oUILib)
		With lblGuildTake
			.ControlName = "lblGuildTake"
			.Left = 50
			.Top = 200
			.Width = 80
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Guild Take:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblGuildTake, UIControl))

		'txtGuildTake initial props
		txtGuildTake = New UITextBox(oUILib)
		With txtGuildTake
			.ControlName = "txtGuildTake"
			.Left = 135
			.Top = 200
			.Width = 35
			.Height = 22
			.Enabled = bEnabled
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
		Me.AddChild(CType(txtGuildTake, UIControl))

		'lblCourse initial props
		lblCourse = New UILabel(oUILib)
		With lblCourse
			.ControlName = "lblCourse"
			.Left = 255
			.Top = 35
			.Width = 106
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Course Details:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblCourse, UIControl))

		'lstCourse initial props
		lstCourse = New UIListBox(oUILib)
		With lstCourse
			.ControlName = "lstCourse"
			.Left = 255
			.Top = 55
			.Width = 230
			.Height = 135
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
		End With
		Me.AddChild(CType(lstCourse, UIControl))

		'btnAddWP initial props
		btnAddWP = New UIButton(oUILib)
		With btnAddWP
			.ControlName = "btnAddWP"
			.Left = 255
			.Top = 195
			.Width = 110
			.Height = 24
			.Enabled = bEnabled
			.Visible = True
			.Caption = "Add Waypoint"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnAddWP, UIControl))

		'btnRemoveWP initial props
		btnRemoveWP = New UIButton(oUILib)
		With btnRemoveWP
			.ControlName = "btnRemoveWP"
			.Left = 380
			.Top = 195
			.Width = 110
			.Height = 24
			.Enabled = bEnabled
			.Visible = True
			.Caption = "Remove"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnRemoveWP, UIControl))

		'lblLaps initial props
		lblLaps = New UILabel(oUILib)
		With lblLaps
			.ControlName = "lblLaps"
			.Left = 155
			.Top = 120
			.Width = 36
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Laps:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblLaps, UIControl))

		'txtLaps initial props
		txtLaps = New UITextBox(oUILib)
		With txtLaps
			.ControlName = "txtLaps"
			.Left = 200
			.Top = 122
			.Width = 35
			.Height = 22
			.Enabled = bEnabled
			.Visible = True
			.Caption = "1"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 5
			.BorderColor = muSettings.InterfaceBorderColor
			.bNumericOnly = True
		End With
		Me.AddChild(CType(txtLaps, UIControl))

		'chkGroundOnly initial props
		chkGroundOnly = New UICheckBox(oUILib)
		With chkGroundOnly
			.ControlName = "chkGroundOnly"
			.Left = 55
			.Top = 95
			.Width = 154
			.Height = 24
			.Enabled = bEnabled
			.Visible = True
			.Caption = "Ground-Based Only"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
		End With
		Me.AddChild(CType(chkGroundOnly, UIControl))

		'lblPotSplit initial props
		lblPotSplit = New UILabel(oUILib)
		With lblPotSplit
			.ControlName = "lblPotSplit"
			.Left = 10
			.Top = 150
			.Width = 152
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Pot Split Percentages"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblPotSplit, UIControl))

		'btnSave initial props
		btnSave = New UIButton(oUILib)
		With btnSave
			.ControlName = "btnSave"
			.Left = 315
			.Top = 225
			.Width = 120
			.Height = 24
			.Enabled = bEnabled
			.Visible = True
			.Caption = "Save Config"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnSave, UIControl))

		'btnClose initial props
		btnClose = New UIButton(oUILib)
		With btnClose
			.ControlName = "btnClose"
			.Left = 485
			.Top = 4
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

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)

		Dim yMsg(6) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eAdvancedEventConfig).CopyTo(yMsg, 0)
		yMsg(2) = 0	'request
		System.BitConverter.GetBytes(mlEventID).CopyTo(yMsg, 3)
		MyBase.moUILib.SendMsgToPrimary(yMsg)

	End Sub

	Private Sub btnAddWP_Click(ByVal sName As String) Handles btnAddWP.Click
		MyBase.moUILib.yRenderUI = 1
		MyBase.moUILib.lUISelectState = UILib.eSelectState.eSelectRaceWaypoint
		MyBase.moUILib.AddNotification("Left-click a location for this way point.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		MyBase.moUILib.AddNotification("Right-click to cancel.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnRemoveWP_Click(ByVal sName As String) Handles btnRemoveWP.Click
		If lstCourse.ListIndex > -1 Then
			Dim lIdx As Int32 = lstCourse.ItemData(lstCourse.ListIndex)
			For X As Int32 = lIdx To mlWPUB - 1
				muWPs(X) = muWPs(X + 1)
			Next X
			mlWPUB -= 1
			RefreshWPList()
		End If
	End Sub

	Private Sub btnSave_Click(ByVal sName As String) Handles btnSave.Click
		'update the server with the race config
		'let's validate our data
		If txtEntryFee.Caption = "" Then txtEntryFee.Caption = "0"
		If txtMinHull.Caption = "" Then txtMinHull.Caption = "0"
		If txtMaxHull.Caption = "" Then txtMaxHull.Caption = "0"
		If txtMinRacers.Caption = "" Then txtMinRacers.Caption = "0"
		If txtFirstPlace.Caption = "" Then txtFirstPlace.Caption = "0"
		If txtSecondplace.Caption = "" Then txtSecondplace.Caption = "0"
		If txtThirdPlace.Caption = "" Then txtThirdPlace.Caption = "0"
		If txtGuildTake.Caption = "" Then txtGuildTake.Caption = "0"
		If txtLaps.Caption = "" Then txtLaps.Caption = "0"

		Dim blEntryFee As Int64 = CLng(txtEntryFee.Caption)
		Dim lMinHull As Int32 = CInt(txtMinHull.Caption)
		Dim lMaxHull As Int32 = CInt(txtMaxHull.Caption)
		Dim lMinRacers As Int32 = CInt(txtMinRacers.Caption)
		Dim lFirstPlace As Int32 = CInt(txtFirstPlace.Caption)
		Dim lSecondPlace As Int32 = CInt(txtSecondplace.Caption)
		Dim lThirdPlace As Int32 = CInt(txtThirdPlace.Caption)
		Dim lGuildTake As Int32 = CInt(txtGuildTake.Caption)
		Dim lLaps As Int32 = CInt(txtLaps.Caption)

		Dim sTest As String = txtEntryFee.Caption & txtMinHull.Caption & txtMaxHull.Caption & txtMinRacers.Caption & txtFirstPlace.Caption & txtSecondplace.Caption & _
		 txtThirdPlace.Caption & txtGuildTake.Caption & txtLaps.Caption
		If sTest.Contains(".") = True Then
			MyBase.moUILib.AddNotification("All values must be whole numbers, no decimal places.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		If blEntryFee < 0 Then
			MyBase.moUILib.AddNotification("Entry Fee cannot be a negative number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		If lMinHull < 0 Then
			MyBase.moUILib.AddNotification("Minimum Hull cannot be a negative number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		If lMaxHull < 0 Then
			MyBase.moUILib.AddNotification("Maxmimum Hull cannot be a negative number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		If lMinHull > lMaxHull Then
			MyBase.moUILib.AddNotification("Minimum Hull cannot be greater than Maximum Hull.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		If lMinRacers < 0 Then
			MyBase.moUILib.AddNotification("Minimum Racers cannot be a negative number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		If lFirstPlace < 0 OrElse lSecondPlace < 0 OrElse lThirdPlace < 0 OrElse lGuildTake < 0 Then
			MyBase.moUILib.AddNotification("Prize distribution values cannot be a negative number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		Dim lTotal As Int32 = lFirstPlace + lSecondPlace + lThirdPlace + lGuildTake
		If lTotal <> 100 AndAlso blEntryFee <> 0 Then
			MyBase.moUILib.AddNotification("Sum of Prize and Guild Take exceeds 100%. Numbers must sum up to 100% if entry fee is used.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		If lLaps < 1 Then
			MyBase.moUILib.AddNotification("Laps must be greater than 0.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		If chkGroundOnly.Value = True Then
			For X As Int32 = 0 To mlWPUB
				If muWPs(X).iEnvirTypeID = ObjectType.eSolarSystem Then
					MyBase.moUILib.AddNotification("Ground Only cannot be checked if waypoints exist in space.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					Return
				End If
			Next X
		End If

		If ConfirmValueRange(lFirstPlace, 0, 100, "First Place Prize") = False Then Return
		If ConfirmValueRange(lSecondPlace, 0, 100, "Second Place Prize") = False Then Return
		If ConfirmValueRange(lThirdPlace, 0, 100, "Third Place Prize") = False Then Return
		If ConfirmValueRange(lGuildTake, 0, 100, "Guild Take") = False Then Return
		If ConfirmValueRange(lLaps, 0, 255, "Laps") = False Then Return
		If ConfirmValueRange(lMinRacers, 0, 255, "Minimum Racers") = False Then Return

		Dim yFirstPlace As Byte = CByte(lFirstPlace)
		Dim ySecondPlace As Byte = CByte(lSecondPlace)
		Dim yThirdPlace As Byte = CByte(lThirdPlace)
		Dim yGuildTake As Byte = CByte(lGuildTake)
		Dim yLaps As Byte = CByte(lLaps)
		Dim yMinRacers As Byte = CByte(lMinRacers)

		Dim lWPCnt As Int32 = mlWPUB + 1

		Dim yMsg(33 + (lWPCnt * 14)) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eAdvancedEventConfig).CopyTo(yMsg, lPos) : lPos += 2
		yMsg(lPos) = 1 : lPos += 1		'indicating we are adding a race config
		System.BitConverter.GetBytes(mlEventID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(blEntryFee).CopyTo(yMsg, lPos) : lPos += 8
		System.BitConverter.GetBytes(lMinHull).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(lMaxHull).CopyTo(yMsg, lPos) : lPos += 4
		yMsg(lPos) = yMinRacers : lPos += 1
		yMsg(lPos) = yFirstPlace : lPos += 1
		yMsg(lPos) = ySecondPlace : lPos += 1
		yMsg(lPos) = yThirdPlace : lPos += 1
		yMsg(lPos) = yGuildTake : lPos += 1
		yMsg(lPos) = yLaps : lPos += 1
		If chkGroundOnly.Value = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 0
		lPos += 1

		System.BitConverter.GetBytes(lWPCnt).CopyTo(yMsg, lPos) : lPos += 4
		For X As Int32 = 0 To mlWPUB
			With muWPs(X)
				System.BitConverter.GetBytes(.lEnvirID).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(.iEnvirTypeID).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(.lX).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(.lZ).CopyTo(yMsg, lPos) : lPos += 4
			End With
		Next X

		MyBase.moUILib.SendMsgToPrimary(yMsg)

		If lWPCnt < 2 Then
			MyBase.moUILib.AddNotification("Race configuration submitted. Missing Course Info.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		Else
			MyBase.moUILib.AddNotification("Race configuration submitted.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		End If

	End Sub

	Private Function ConfirmValueRange(ByVal lValue As Int32, ByVal lMin As Int32, ByVal lMax As Int32, ByVal sValueName As String) As Boolean
		If lValue >= lMin AndAlso lValue <= lMax Then
			Return True
		Else
			MyBase.moUILib.AddNotification(sValueName & " is invalid and must be between " & lMin & " and " & lMax & ".", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return False
		End If
	End Function

	Private Sub chkGroundOnly_Click() Handles chkGroundOnly.Click
		If chkGroundOnly.Value = True Then
			For X As Int32 = 0 To mlWPUB
				If muWPs(X).iEnvirTypeID = ObjectType.eSolarSystem Then
					MyBase.moUILib.AddNotification("The course has waypoints in space. Ground Only cannot be checked.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
					chkGroundOnly.Value = False
					Return
				End If
			Next X
		End If
	End Sub

	Private Sub lstCourse_ItemClick(ByVal lIndex As Integer) Handles lstCourse.ItemClick

		If lIndex > -1 Then
			Dim lIdx As Int32 = lstCourse.ItemData(lstCourse.ListIndex)

			Dim uWP As PlayerComm.WPAttachment
			With uWP
				.AttachNumber = 0
				.EnvirID = muWPs(lIdx).lEnvirID
				.EnvirTypeID = muWPs(lIdx).iEnvirTypeID
				.LocX = muWPs(lIdx).lX
				.LocZ = muWPs(lIdx).lZ
				.sWPName = "WP"
				.JumpToAttachment()
			End With
		End If
	End Sub

	Public Sub SetLocResultVector(ByVal vecLoc As Vector3)

		mlWPUB += 1
		ReDim Preserve muWPs(mlWPUB)
		With muWPs(mlWPUB)
			If goGalaxy Is Nothing = False Then
				If goGalaxy.CurrentSystemIdx > -1 Then
					.lEnvirID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID
					.iEnvirTypeID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjTypeID
					If glCurrentEnvirView = CurrentView.ePlanetView OrElse glCurrentEnvirView = CurrentView.ePlanetMapView Then
						If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx > -1 Then
							Dim oPlanet As Planet = Nothing
							With goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
								oPlanet = .moPlanets(.CurrentPlanetIdx)
							End With
							If oPlanet Is Nothing = False Then
								.lEnvirID = oPlanet.ObjectID
								.iEnvirTypeID = oPlanet.ObjTypeID
							End If
						End If
					End If
				End If
			Else
				If goCurrentEnvir Is Nothing = False Then
					.lEnvirID = goCurrentEnvir.ObjectID
					.iEnvirTypeID = goCurrentEnvir.ObjTypeID
				Else : Return
				End If
			End If

			If .iEnvirTypeID = ObjectType.eSolarSystem Then
				If chkGroundOnly.Value = True Then
					chkGroundOnly.Value = False
				End If
			End If

			.lX = CInt(vecLoc.X)
			.lZ = CInt(vecLoc.Z)
		End With

		RefreshWPList()

		Me.Visible = True
		MyBase.moUILib.yRenderUI = 255
		MyBase.moUILib.lUISelectState = UILib.eSelectState.eNoSelectState
	End Sub

	Private Sub RefreshWPList()
		lstCourse.Clear()
		For X As Int32 = 0 To mlWPUB
			lstCourse.AddItem("Waypoint " & X + 1, False)
			lstCourse.ItemData(lstCourse.NewIndex) = X
		Next X
	End Sub

	Public Sub HandleAdvancedEventConfig(ByRef yData() As Byte)
		Dim lPos As Int32 = 3 'for msgcode and type

		Dim lEventID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		If lEventID <> mlEventID Then Return

		txtEntryFee.Caption = System.BitConverter.ToInt64(yData, lPos).ToString : lPos += 8
		txtMinHull.Caption = System.BitConverter.ToInt32(yData, lPos).ToString : lPos += 4
		txtMaxHull.Caption = System.BitConverter.ToInt32(yData, lPos).ToString : lPos += 4

		txtMinRacers.Caption = yData(lPos).ToString : lPos += 1
		txtFirstPlace.Caption = yData(lPos).ToString : lPos += 1
		txtSecondplace.Caption = yData(lPos).ToString : lPos += 1
		txtThirdPlace.Caption = yData(lPos).ToString : lPos += 1
		txtGuildTake.Caption = yData(lPos).ToString : lPos += 1
		txtLaps.Caption = yData(lPos).ToString : lPos += 1
		chkGroundOnly.Value = yData(lPos) <> 0 : lPos += 1

		mlWPUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
		ReDim muWPs(mlWPUB)

		For X As Int32 = 0 To mlWPUB
			With muWPs(X)
				.lEnvirID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.iEnvirTypeID = System.BitConverter.ToInt16(yData, lPos) : lPos += 2
				.lX = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
				.lZ = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
			End With
		Next X

		RefreshWPList()

	End Sub
End Class