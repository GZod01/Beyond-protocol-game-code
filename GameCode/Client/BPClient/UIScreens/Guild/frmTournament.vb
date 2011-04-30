Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmTournament
	Inherits UIWindow

	Private lblTitle As UILabel
	Private lnDiv1 As UILine
	Private lblMaxPlayers As UILabel
	Private txtMaxPlayers As UITextBox
	Private lblMaxUnits As UILabel
	Private lblEntryFee As UILabel
	Private lblGround As UILabel
	Private lblAir As UILabel
	Private txtMaxUnits As UITextBox
	Private txtMaxGround As UITextBox
	Private txtMaxAir As UITextBox
	Private fraMap As UIWindow
	Private txtEntryFee As UITextBox
	Private lblGuildTake As UILabel
	Private txtGuildTake As UITextBox

	Private WithEvents btnClose As UIButton
	Private WithEvents lscMap As UILabelScroller
	Private WithEvents btnSave As UIButton

	Private mlEventID As Int32 = -1

	Private moMapTex As Texture = Nothing
    Private Shared moSprite As Sprite = Nothing
    Public Shared Sub ReleaseSprite()
        Try
            If moSprite Is Nothing = False Then moSprite.Dispose()
            moSprite = Nothing
        Catch
        End Try
    End Sub

	Public Sub New(ByRef oUILib As UILib, ByVal lEventID As Int32)
		MyBase.New(oUILib)

		mlEventID = lEventID

		'frmTournament initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eTournament
            .ControlName = "frmTournament"
            .Left = 282
            .Top = 263
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
			.Width = 210
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Tournament Configuration"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTitle, UIControl))

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

		'lblMaxPlayers initial props
		lblMaxPlayers = New UILabel(oUILib)
		With lblMaxPlayers
			.ControlName = "lblMaxPlayers"
			.Left = 220
			.Top = 40
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Max Players:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblMaxPlayers, UIControl))

		'txtMaxPlayers initial props
		txtMaxPlayers = New UITextBox(oUILib)
		With txtMaxPlayers
			.ControlName = "txtMaxPlayers"
			.Left = 315
			.Top = 42
			.Width = 100
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "0"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 2
			.BorderColor = muSettings.InterfaceBorderColor
			.bNumericOnly = True
		End With
		Me.AddChild(CType(txtMaxPlayers, UIControl))

		'lblMaxUnits initial props
		lblMaxUnits = New UILabel(oUILib)
		With lblMaxUnits
			.ControlName = "lblMaxUnits"
			.Left = 220
			.Top = 73
			.Width = 150
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Max Units Per Player:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblMaxUnits, UIControl))

		'lblEntryFee initial props
		lblEntryFee = New UILabel(oUILib)
		With lblEntryFee
			.ControlName = "lblEntryFee"
			.Left = 220
			.Top = 130
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

		'lblGround initial props
		lblGround = New UILabel(oUILib)
		With lblGround
			.ControlName = "lblGround"
			.Left = 230
			.Top = 100
			.Width = 57
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Ground:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblGround, UIControl))

		'lblAir initial props
		lblAir = New UILabel(oUILib)
		With lblAir
			.ControlName = "lblAir"
			.Left = 360
			.Top = 100
			.Width = 27
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Air:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblAir, UIControl))

		'txtMaxUnits initial props
		txtMaxUnits = New UITextBox(oUILib)
		With txtMaxUnits
			.ControlName = "txtMaxUnits"
			.Left = 375
			.Top = 72
			.Width = 100
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "0"
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
		Me.AddChild(CType(txtMaxUnits, UIControl))

		'txtMaxGround initial props
		txtMaxGround = New UITextBox(oUILib)
		With txtMaxGround
			.ControlName = "txtMaxGround"
			.Left = 290
			.Top = 102
			.Width = 50
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "0"
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
		Me.AddChild(CType(txtMaxGround, UIControl))

		'txtMaxAir initial props
		txtMaxAir = New UITextBox(oUILib)
		With txtMaxAir
			.ControlName = "txtMaxAir"
			.Left = 390
			.Top = 102
			.Width = 50
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "0"
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
		Me.AddChild(CType(txtMaxAir, UIControl))

		'lscMap initial props
		lscMap = New UILabelScroller(oUILib)
		With lscMap
			.ControlName = "lscMap"
			.Left = 10
			.Top = 235
			.Width = 190
			.Height = 18
			.Enabled = True
			.Visible = True
		End With
		Me.AddChild(CType(lscMap, UIControl))

		'fraMap initial props
		fraMap = New UIWindow(oUILib)
		With fraMap
			.ControlName = "fraMap"
			.Left = 5
			.Top = 35
			.Width = 200
			.Height = 200
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			.BorderLineWidth = 2
		End With
		Me.AddChild(CType(fraMap, UIControl))

		'txtEntryFee initial props
		txtEntryFee = New UITextBox(oUILib)
		With txtEntryFee
			.ControlName = "txtEntryFee"
			.Left = 315
			.Top = 132
			.Width = 100
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "0"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 18
			.bNumericOnly = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(txtEntryFee, UIControl))

		'lblGuildTake initial props
		lblGuildTake = New UILabel(oUILib)
		With lblGuildTake
			.ControlName = "lblGuildTake"
			.Left = 220
			.Top = 160
			.Width = 81
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
			.Left = 315
			.Top = 162
			.Width = 100
			.Height = 22
			.Enabled = True
			.Visible = True
			.Caption = "0"
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
		Me.AddChild(CType(txtGuildTake, UIControl))

		'btnSave initial props
		btnSave = New UIButton(oUILib)
		With btnSave
			.ControlName = "btnSave"
			.Left = 310
			.Top = 215
			.Width = 100
			.Height = 32
			.Enabled = True
			.Visible = True
			.Caption = "Save"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnSave, UIControl))

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)

		FillMapList()

		Dim yMsg(6) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eAdvancedEventConfig).CopyTo(yMsg, 0)
		yMsg(2) = 0	'request
		System.BitConverter.GetBytes(mlEventID).CopyTo(yMsg, 3)
		MyBase.moUILib.SendMsgToPrimary(yMsg)
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub FillMapList()
		lscMap.AddItem(1, "Tutorial Planet")
	End Sub

	Private Sub lscMap_ItemChanged(ByVal lID As Integer, ByVal sDisplay As String) Handles lscMap.ItemChanged
		If moMapTex Is Nothing = False Then moMapTex.Dispose()
		moMapTex = Nothing
		moMapTex = goResMgr.LoadScratchTexture("Map_" & lID & ".dds", "gi.pak")
		Me.IsDirty = True
	End Sub

	Private Sub btnSave_Click(ByVal sName As String) Handles btnSave.Click

		If txtMaxPlayers.Caption = "" Then txtMaxPlayers.Caption = "0"
		If txtMaxUnits.Caption = "" Then txtMaxUnits.Caption = "0"
		If txtMaxGround.Caption = "" Then txtMaxGround.Caption = "0"
		If txtMaxAir.Caption = "" Then txtMaxAir.Caption = "0"
		If txtEntryFee.Caption = "" Then txtEntryFee.Caption = "0"
		If txtGuildTake.Caption = "" Then txtGuildTake.Caption = "0"

		Dim sTest As String = txtMaxPlayers.Caption & txtMaxUnits.Caption & txtMaxGround.Caption & txtMaxAir.Caption & txtEntryFee.Caption & txtGuildTake.Caption
		If sTest.Contains(".") = True Then
			MyBase.moUILib.AddNotification("All values must be non-decimal numbers.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		Dim blEntryFee As Int64 = CLng(txtEntryFee.Caption)
		If blEntryFee < 0 Then
			MyBase.moUILib.AddNotification("Entry Fee cannot be a negative number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		Dim lMaxPlayers As Int32 = CInt(txtMaxPlayers.Caption)
		If lMaxPlayers < 0 Then
			MyBase.moUILib.AddNotification("Max players cannot be a negative number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		Dim lMaxUnits As Int32 = CInt(txtMaxUnits.Caption)
		If lMaxUnits < 0 Then
			MyBase.moUILib.AddNotification("Maxmimum units cannot be a negative number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		Dim lMaxGround As Int32 = CInt(txtMaxGround.Caption)
		If lMaxGround < 0 Then
			MyBase.moUILib.AddNotification("Maximum Ground Units cannot be a negative number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		Dim lMaxAir As Int32 = CInt(txtMaxAir.Caption)
		If lMaxAir < 0 Then
			MyBase.moUILib.AddNotification("Maximum Air Units cannot be a negative number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		Dim lGuildTake As Int32 = CInt(txtGuildTake.Caption)
		
		If ConfirmValueRange(lMaxPlayers, 0, 30, "Max Players") = False Then Return
		If ConfirmValueRange(lMaxUnits, 0, 200, "Max Units") = False Then Return
		If ConfirmValueRange(lMaxGround, 0, 200, "Max Ground Units") = False Then Return
		If ConfirmValueRange(lGuildTake, 0, 100, "Guild Take") = False Then Return
		If ConfirmValueRange(lMaxAir, 0, 200, "Max Air Units") = False Then Return

		Dim yMsg(23) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eAdvancedEventConfig).CopyTo(yMsg, lPos) : lPos += 2
		yMsg(lPos) = 2 : lPos += 1
		System.BitConverter.GetBytes(mlEventID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(blEntryFee).CopyTo(yMsg, lPos) : lPos += 8
		System.BitConverter.GetBytes(lscMap.SelectedItemID).CopyTo(yMsg, lPos) : lPos += 4
		yMsg(lPos) = CByte(lMaxPlayers) : lPos += 1
		yMsg(lPos) = CByte(lMaxUnits) : lPos += 1
		yMsg(lPos) = CByte(lMaxGround) : lPos += 1
		yMsg(lPos) = CByte(lGuildTake) : lPos += 1
		yMsg(lPos) = CByte(lMaxAir) : lPos += 1

		MyBase.moUILib.SendMsgToPrimary(yMsg)
		MyBase.moUILib.AddNotification("Tournament configuration submitted.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

	End Sub

	Private Function ConfirmValueRange(ByVal lValue As Int32, ByVal lMin As Int32, ByVal lMax As Int32, ByVal sValueName As String) As Boolean
		If lValue >= lMin AndAlso lValue <= lMax Then
			Return True
		Else
			MyBase.moUILib.AddNotification(sValueName & " is invalid and must be between " & lMin & " and " & lMax & ".", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return False
		End If
	End Function

	Public Sub HandleAdvancedEventConfig(ByRef yData() As Byte)
		Dim lPos As Int32 = 3 'for msgcode and type

		Dim lEventID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
		If lEventID <> mlEventID Then Return

		txtEntryFee.Caption = System.BitConverter.ToInt64(yData, lPos).ToString : lPos += 8
		lscMap.SelectByID(System.BitConverter.ToInt32(yData, lPos)) : lPos += 4

		txtMaxPlayers.Caption = yData(lPos).ToString : lPos += 1
		txtMaxUnits.Caption = yData(lPos).ToString : lPos += 1
		txtMaxGround.Caption = yData(lPos).ToString : lPos += 1
		txtMaxAir.Caption = yData(lPos).ToString : lPos += 1
		txtGuildTake.Caption = yData(lPos).ToString : lPos += 1

		Me.IsDirty = True
	End Sub 

	Private Sub frmTournament_OnRenderEnd() Handles Me.OnRenderEnd
		If moMapTex Is Nothing Then Return

		If moSprite Is Nothing OrElse moSprite.Disposed = True Then
			Device.IsUsingEventHandlers = False
			If moSprite Is Nothing = False Then moSprite.Dispose()
			moSprite = Nothing
			moSprite = New Sprite(MyBase.moUILib.oDevice)
			Device.IsUsingEventHandlers = True
		End If

		moSprite.Begin(SpriteFlags.AlphaBlend)
		Try
			Dim rcDest As Rectangle = New Rectangle(0, 0, fraMap.Width - 2, fraMap.Height - 2)
			Dim ptDest As Point = fraMap.GetAbsolutePosition()
			ptDest.X += 4
			ptDest.Y += 14
			moSprite.Draw2D(moMapTex, System.Drawing.Rectangle.Empty, rcDest, ptDest, System.Drawing.Color.FromArgb(255, 255, 255, 255))
		Catch
			'do nothing?
		End Try
		moSprite.End()
	End Sub
End Class