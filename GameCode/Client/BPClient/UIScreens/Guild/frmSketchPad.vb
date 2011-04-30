Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmSketchPad
	Inherits UIWindow

	Private txtName As UITextBox
	Private btnCircle As UIButton
	Private btnLine As UIButton
	Private btnX As UIButton
	Private btnBox As UIButton
	Private btnArrow As UIButton
	Private btnText As UIButton
	Private WithEvents btnClear As UIButton
	Private WithEvents btnFinished As UIButton

	Private mlClr As Int32 = 0

	Private mySelectedShape As SketchPad.eySketchShapes = SketchPad.eySketchShapes.Box
	Private mptA As Point = Point.Empty
	Private mptB As Point = Point.Empty

	Private moSketchPad As SketchPad
	Private mlEventID As Int32 = -1

    'Private Shared moFont As Font
    'Public Shared Sub ReleaseFont()
    '    If moFont Is Nothing = False Then moFont.Dispose()
    '    moFont = Nothing
    'End Sub

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'frmSketch initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eSketchPad
            .ControlName = "frmSketchPad"
            .Left = 0
            .Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 128
            .Width = 128
            .Height = 256
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
            .BorderLineWidth = 2
        End With

		'txtName initial props
		txtName = New UITextBox(oUILib)
		With txtName
			.ControlName = "txtName"
			.Left = 2
			.Top = 2
			.Width = 124
			.Height = 20
			.Enabled = True
			.Visible = True
			.Caption = "Unnamed"
			.ForeColor = muSettings.InterfaceTextBoxForeColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
			.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
			.MaxLength = 20
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(txtName, UIControl))

		''btnCircle initial props
		'btnCircle = New UIButton(oUILib)
		'With btnCircle
		'	.ControlName = "btnCircle"
		'	.Left = 2
		'	.Top = 23
		'	.Width = 62
		'	.Height = 24
		'	.Enabled = True
		'	.Visible = False
		'	.Caption = "Circle"
		'	.ForeColor = muSettings.InterfaceBorderColor
		'	.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
		'	.DrawBackImage = True
		'	.FontFormat = CType(5, DrawTextFormat)
		'	.ControlImageRect = New Rectangle(0, 0, 120, 32)
		'End With
		'Me.AddChild(CType(btnCircle, UIControl))

		'btnLine initial props
		btnLine = New UIButton(oUILib)
		With btnLine
			.ControlName = "btnLine"
			.Left = 64
			.Top = 23
			.Width = 62
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Line"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnLine, UIControl))

		'btnX initial props
		btnX = New UIButton(oUILib)
		With btnX
			.ControlName = "btnX"
			.Left = 2
			.Top = 50
			.Width = 62
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
		Me.AddChild(CType(btnX, UIControl))

		'btnBox initial props
		btnBox = New UIButton(oUILib)
		With btnBox
			.ControlName = "btnBox"
			.Left = 64
			.Top = 50
			.Width = 62
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Box"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnBox, UIControl))

		'btnOval initial props
		btnArrow = New UIButton(oUILib)
		With btnArrow
			.ControlName = "btnArrow"
			.Left = 2
			.Top = 23
			.Width = 62
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Arrow"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnArrow, UIControl))

		'btnText initial props
		btnText = New UIButton(oUILib)
		With btnText
			.ControlName = "btnText"
			.Left = Me.Width \ 2 - 31
			.Top = 77
			.Width = 62
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Text"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnText, UIControl))

		'btnClear initial props
		btnClear = New UIButton(oUILib)
		With btnClear
			.ControlName = "btnClear"
			.Left = 2
			.Top = 230
			.Width = 62
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Clear"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnClear, UIControl))

		'btnFinished initial props
		btnFinished = New UIButton(oUILib)
		With btnFinished
			.ControlName = "btnFinished"
			.Left = 64
			.Top = 230
			.Width = 62
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Done"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnFinished, UIControl))

		AddHandler btnArrow.Click, AddressOf ShapeBoxClick
		AddHandler btnBox.Click, AddressOf ShapeBoxClick
		'AddHandler btnCircle.Click, AddressOf ShapeBoxClick
		AddHandler btnLine.Click, AddressOf ShapeBoxClick
		AddHandler btnText.Click, AddressOf ShapeBoxClick
		AddHandler btnX.Click, AddressOf ShapeBoxClick

		moSketchPad = New SketchPad()

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)
	End Sub

	Public Sub SetExternalData(ByVal lEventID As Int32, ByRef oSketchPad As SketchPad)
		moSketchPad = oSketchPad
		mlEventID = lEventID
	End Sub

	Private Sub frmSketchPad_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
		Dim pt As Point = Me.GetAbsolutePosition()
		lMouseX -= pt.X
		lMouseY -= pt.Y

		If lMouseY >= 110 AndAlso lMouseY <= 220 Then
			lMouseY -= 110
			lMouseY \= 28
			If lMouseX > 10 Then
				lMouseX -= 10
				lMouseX \= 28
				If lMouseX > 3 Then Return

				mlClr = lMouseY * 4 + lMouseX
				Me.IsDirty = True
			End If
		End If
	End Sub

	Private Sub frmSketchPad_OnNewFrameEnd() Handles Me.OnNewFrameEnd
		Dim lTotalWidth As Int32 = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth
		Dim lTotalHeight As Int32 = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight

        'If moFont Is Nothing OrElse moFont.Disposed = True Then
        '	Device.IsUsingEventHandlers = False
        '	moFont = Nothing
        '	moFont = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '	Device.IsUsingEventHandlers = True
        'End If
        Dim oSysFont As System.Drawing.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0)

		If moSketchPad Is Nothing = False Then
			If moSketchPad.uItems Is Nothing = False Then
				For X As Int32 = 0 To moSketchPad.uItems.GetUpperBound(0)
					Dim uItem As SketchPad.SketchPadItem = moSketchPad.uItems(X)

					'Ok, determine our type
					Select Case uItem.yType
						Case SketchPad.eySketchShapes.Arrow
							'Ok, draw a line from point a to point b....
							' then draw, two lines from pont b extending at a 45 degree angle away from the line angle of the line

							Dim v2Line(4) As Vector2
							Dim fPtAX As Single = uItem.fPtA_X * lTotalWidth
							Dim fPtAY As Single = uItem.fPtA_Y * lTotalHeight
							Dim fPtBX As Single = uItem.fPtB_X * lTotalWidth
							Dim fPtBY As Single = uItem.fPtB_Y * lTotalHeight

							v2Line(0) = New Vector2(uItem.fPtA_X * lTotalWidth, uItem.fPtA_Y * lTotalHeight)
							v2Line(1) = New Vector2(fPtBX, fPtBY)

							Dim fAngle As Single = LineAngleDegrees(CInt(fPtAX), CInt(fPtAY), CInt(fPtBX), CInt(fPtBY)) - 180
							If fAngle < 0 Then fAngle += 360
							Dim fTmpX As Single = 30
							Dim fTmpZ As Single = 0
							RotatePoint(0, 0, fTmpX, fTmpZ, fAngle - 45)
							v2Line(2) = New Vector2(fPtBX + fTmpX, fPtBY + fTmpZ)
							fTmpX = 30
							fTmpZ = 0
							RotatePoint(0, 0, fTmpX, fTmpZ, fAngle + 45)
							v2Line(3) = New Vector2(fPtBX + fTmpX, fPtBY + fTmpZ)
							v2Line(4) = New Vector2(fPtBX, fPtBY)

                            'ValidateBorderLine()
                            'With moBorderLine
                            '    .Antialias = True
                            '    .Width = 3
                            '    .Begin()
                            '    .Draw(v2Line, SketchPad.GetColorFromValue(uItem.yClrVal))
                            '    .End()
                            'End With
                            BPLine.DrawLine(3, True, v2Line, SketchPad.GetColorFromValue(uItem.yClrVal))
						Case SketchPad.eySketchShapes.Box
							Dim lPtAX As Int32 = CInt(uItem.fPtA_X * lTotalWidth)
							Dim lPtAY As Int32 = CInt(uItem.fPtA_Y * lTotalHeight)
							Dim lPtBX As Int32 = CInt(uItem.fPtB_X * lTotalWidth)
							Dim lPtBY As Int32 = CInt(uItem.fPtB_Y * lTotalHeight)
							If lPtAX > lPtBX OrElse lPtAY > lPtBY Then
								Dim lTmp As Int32 = lPtAX
								lPtAX = lPtBX : lPtBX = lTmp
								lTmp = lPtAY
								lPtAY = lPtBY : lPtBY = lTmp
							End If
							Dim rcDim As Rectangle = Rectangle.FromLTRB(lPtAX, lPtAY, lPtBX, lPtBY)
                            RenderBox(rcDim, 3, SketchPad.GetColorFromValue(uItem.yClrVal))
						Case SketchPad.eySketchShapes.Cross
							'Just an X, so render from top left, bottom right, top right, bottom left
							Dim v2Line(1) As Vector2
							Dim v2Line2(1) As Vector2
							Dim fPtAX As Single = uItem.fPtA_X * lTotalWidth
							Dim fPtAY As Single = uItem.fPtA_Y * lTotalHeight
							Dim fPtBX As Single = uItem.fPtB_X * lTotalWidth
							Dim fPtBY As Single = uItem.fPtB_Y * lTotalHeight

							v2Line(0) = New Vector2(Math.Min(fPtAX, fPtBX), Math.Min(fPtAY, fPtBY))
							v2Line(1) = New Vector2(Math.Max(fPtAX, fPtBX), Math.Max(fPtAY, fPtBY))

							v2Line2(0) = New Vector2(Math.Min(fPtAX, fPtBX), Math.Max(fPtAY, fPtBY))
							v2Line2(1) = New Vector2(Math.Max(fPtAX, fPtBX), Math.Min(fPtAY, fPtBY))

                            'ValidateBorderLine()
                            'With moBorderLine
                            '    .Antialias = True
                            '    .Width = 3
                            '    .Begin()
                            '    .Draw(v2Line, SketchPad.GetColorFromValue(uItem.yClrVal))
                            '    .Draw(v2Line2, SketchPad.GetColorFromValue(uItem.yClrVal))
                            '    .End()
                            'End With
                            BPLine.DrawLine(3, True, v2Line, SketchPad.GetColorFromValue(uItem.yClrVal))
                            BPLine.DrawLine(3, True, v2Line2, SketchPad.GetColorFromValue(uItem.yClrVal))
						Case SketchPad.eySketchShapes.Line
							' draw a line from pt a to pt b
							Dim v2Line(1) As Vector2
							Dim fPtAX As Single = uItem.fPtA_X * lTotalWidth
							Dim fPtAY As Single = uItem.fPtA_Y * lTotalHeight
							Dim fPtBX As Single = uItem.fPtB_X * lTotalWidth
							Dim fPtBY As Single = uItem.fPtB_Y * lTotalHeight

							v2Line(0) = New Vector2(fPtAX, fPtAY)
							v2Line(1) = New Vector2(fPtBX, fPtBY)

                            'ValidateBorderLine()
                            'With moBorderLine
                            '    .Antialias = True
                            '    .Width = 3
                            '    .Begin()
                            '    .Draw(v2Line, SketchPad.GetColorFromValue(uItem.yClrVal))
                            '    .End()
                            'End With
                            BPLine.DrawLine(3, True, v2Line, SketchPad.GetColorFromValue(uItem.yClrVal))
						Case SketchPad.eySketchShapes.Text
							'draw our text
							Dim lPtX As Int32 = CInt(uItem.fPtA_X * lTotalWidth)
							Dim lPtY As Int32 = CInt(uItem.fPtA_Y * lTotalHeight)
                            'moFont.DrawText(Nothing, uItem.sText, lPtX, lPtY, SketchPad.GetColorFromValue(uItem.yClrVal))
                            BPFont.DrawText(oSysFont, uItem.sText, New Rectangle(lPtX, lPtY, lTotalWidth, lTotalHeight), DrawTextFormat.Left Or DrawTextFormat.Top, SketchPad.GetColorFromValue(uItem.yClrVal))
					End Select

				Next X
			End If
		End If
	End Sub

	Private Sub frmSketchPad_OnRender() Handles Me.OnRender
		Dim lTop As Int32 = 110		'104
		Dim lLeft As Int32 = 10		'2

		For X As Int32 = 0 To 15
			Dim rcDim As Rectangle = New Rectangle(lLeft, lTop, 28, 28)
			MyBase.moUILib.DoAlphaBlendColorFill(rcDim, SketchPad.GetColorFromValue(X), rcDim.Location)

			If mlClr = X Then
                RenderBox(rcDim, 2, SketchPad.GetColorFromValue(X))
			End If

			If (X + 1) Mod 4 = 0 Then
				lLeft = 10 '2
				lTop += 28
			Else
				lLeft += 28
			End If
		Next X
	End Sub

	Private Sub ShapeBoxClick(ByVal sName As String)
		Select Case sName.ToUpper
			Case "BTNCIRCLE"
				mySelectedShape = SketchPad.eySketchShapes.Circle
			Case "BTNLINE"
				mySelectedShape = SketchPad.eySketchShapes.Line
			Case "BTNX"
				mySelectedShape = SketchPad.eySketchShapes.Cross
			Case "BTNBOX"
				mySelectedShape = SketchPad.eySketchShapes.Box
			Case "BTNARROW"
				mySelectedShape = SketchPad.eySketchShapes.Arrow
			Case "BTNTEXT"
				mySelectedShape = SketchPad.eySketchShapes.Text
		End Select

		MyBase.moUILib.lUISelectState = UILib.eSelectState.eSketchPad_SelectPoint
		mptA = Point.Empty
		mptB = Point.Empty
	End Sub

	Private Sub btnClear_Click(ByVal sName As String) Handles btnClear.Click
		ReDim moSketchPad.uItems(-1)
	End Sub

	Private Sub btnFinished_Click(ByVal sName As String) Handles btnFinished.Click
		'save the sketchpad and add as an attachment to the event
		If moSketchPad Is Nothing Then Return
		MyBase.moUILib.yRenderUI = 255
		With moSketchPad
			.sName = txtName.Caption
			.lEnvirID = goCurrentEnvir.ObjectID
			.iEnvirTypeID = goCurrentEnvir.ObjTypeID
			.CameraAtX = goCamera.mlCameraAtX
			.CameraAtY = goCamera.mlCameraAtY
			.CameraAtZ = goCamera.mlCameraAtZ
			.CameraX = goCamera.mlCameraX
			.CameraY = goCamera.mlCameraY
			.CameraZ = goCamera.mlCameraZ
			.ViewID = glCurrentEnvirView
		End With

		Dim yMsg() As Byte = moSketchPad.GetAsAddObjectMsg(mlEventID)
		If yMsg Is Nothing = False Then
			MyBase.moUILib.SendMsgToPrimary(yMsg)
			MyBase.moUILib.RemoveWindow(Me.ControlName)
		End If
	End Sub

	Public Sub PointSelected(ByVal lX As Int32, ByVal lY As Int32)
		If mptA = Point.Empty Then
			mptA.X = lX : mptA.Y = lY

			Dim lTotalWidth As Int32 = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth
			Dim lTotalHeight As Int32 = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight
			Dim fX As Single = CSng(lX / lTotalWidth)
			Dim fY As Single = CSng(lY / lTotalHeight)

			If mySelectedShape <> SketchPad.eySketchShapes.Text Then
				MyBase.moUILib.AddNotification("Left-Click again to finish shape placement.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				MyBase.moUILib.lUISelectState = UILib.eSelectState.eSketchPad_SelectPoint
			Else
				MyBase.moUILib.AddNotification("Enter the text you wish to display.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				MyBase.moUILib.lUISelectState = UILib.eSelectState.eSketchPad_TextEntry
			End If
			moSketchPad.AddItem(mySelectedShape, fX, fY, fX + 2, fY + 2, CByte(mlClr), "")
			Me.IsDirty = True
		Else
			mptB.X = lX : mptB.Y = lY
		End If
	End Sub

	Public Sub PeakPointSelected(ByVal lX As Int32, ByVal lY As Int32)
		If mptA <> Point.Empty AndAlso mptB = Point.Empty Then
			If moSketchPad Is Nothing = False Then
				If moSketchPad.uItems Is Nothing = False Then
					With moSketchPad.uItems(moSketchPad.uItems.GetUpperBound(0))
						Dim lTotalWidth As Int32 = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth
						Dim lTotalHeight As Int32 = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight
						Dim fX As Single = CSng(lX / lTotalWidth)
						Dim fY As Single = CSng(lY / lTotalHeight)

						.fPtB_X = fX
						.fPtB_Y = fY
						Me.IsDirty = True
					End With
				End If
			End If
		End If
	End Sub

	Public Sub AddCharToText(ByVal sChr As String)
		If moSketchPad Is Nothing = False Then
			If moSketchPad.uItems Is Nothing = False Then
				With moSketchPad.uItems(moSketchPad.uItems.GetUpperBound(0))
					.sText &= sChr
					Me.IsDirty = True
				End With
			End If
		End If
	End Sub

	Public Sub BackSpaceHit()
		If moSketchPad Is Nothing = False Then
			If moSketchPad.uItems Is Nothing = False Then
				With moSketchPad.uItems(moSketchPad.uItems.GetUpperBound(0))
					If .sText Is Nothing = False Then
						.sText = .sText.Substring(0, .sText.Length - 1)
					End If
				End With
			End If
		End If
	End Sub

    'Protected Overrides Sub Finalize()
    '	If moFont Is Nothing = False Then moFont.Dispose()
    '	moFont = Nothing
    '	MyBase.Finalize()
    'End Sub
End Class