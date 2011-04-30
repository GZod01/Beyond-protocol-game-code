Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class UIScrollBar
    Inherits UIControl

	Protected mbVertical As Boolean
    Private mlValue As Int32 = 0
    Public Property Value() As Int32 '= 0
        Get
            Return mlValue
        End Get
        Set(ByVal Value As Int32)
            mlValue = Value
            If mlValue > mlMaxValue Then
                mlValue = mlMaxValue
            ElseIf mlValue < mlMinValue Then
                mlValue = mlMinValue
            End If
            IsDirty = True
        End Set
    End Property
    Private mlMaxValue As Int32 = 1
    Public Property MaxValue() As Int32 '= 1
        Get
            Return mlMaxValue
        End Get
        Set(ByVal Value As Int32)
            mlMaxValue = Value
            If mlValue > mlMaxValue Then
                mlValue = mlMaxValue
            End If
            IsDirty = True
        End Set
    End Property
    Private mlMinValue As Int32 = 0
    Public Property MinValue() As Int32
        Get
            Return mlMinValue
        End Get
        Set(ByVal Value As Int32)
            mlMinValue = Value
            IsDirty = True
        End Set
    End Property
    Public SmallChange As Int32 = 1
    Public LargeChange As Int32 = 4

    Private mbReverseDirection As Boolean = False
    Public Property ReverseDirection() As Boolean
        Get
            Return mbReverseDirection
        End Get
        Set(ByVal Value As Boolean)
            mbReverseDirection = Value
            IsDirty = True
        End Set
    End Property

    Private mbScrolling As Boolean = False
    Private mlScrollingMouseX As Int32
    Private mlScrollingMouseY As Int32

    Protected WithEvents btnDecreaser As UIButton
    Protected WithEvents btnIncreaser As UIButton
    Protected WithEvents btnScroller As UIButton

    Public Event ValueChange()

    Public Property IsVertical() As Boolean
        Get
            Return mbVertical
        End Get
        Set(ByVal Value As Boolean)
            mbVertical = Value
            SetButtonImageRects()
        End Set
    End Property

    Public Sub New(ByRef oUILib As UILib, ByVal bAsVertical As Boolean)
        MyBase.New(oUILib)
        mbVertical = bAsVertical

        Me.AcceptMouseWheelEvents = True

        btnDecreaser = New UIButton(oUILib)
        btnIncreaser = New UIButton(oUILib)
        btnScroller = New UIButton(oUILib)

        btnDecreaser.mbAcceptReprocessEvents = True
        btnIncreaser.mbAcceptReprocessEvents = True
        btnScroller.mbAcceptReprocessEvents = True
        Me.mbAcceptReprocessEvents = True

        btnDecreaser.ControlName = "btnDecreaser"
        btnIncreaser.ControlName = "btnIncreaser"
        btnScroller.ControlName = "btnScroller"

        SetButtonImageRects()

        AddChild(CType(btnDecreaser, UIControl))
        AddChild(CType(btnIncreaser, UIControl))
        AddChild(CType(btnScroller, UIControl))
    End Sub

    Private Sub SetButtonImageRects()
        'here, we need to load the images for the buttons...
        If mbVertical = True Then
            With btnDecreaser
                '.ControlImage_Disabled = oUILib.oResMgr.GetTexture("VScr_Dec_Btn_Disabled.bmp")
                .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eDownArrow_Button_Disabled)
                '.ControlImage_Normal = oUILib.oResMgr.GetTexture("VScr_Dec_Btn_Normal.bmp")
                .ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eDownArrow_Button_Normal)
                '.ControlImage_Pressed = oUILib.oResMgr.GetTexture("VScr_Dec_Btn_Pressed.bmp")
                .ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eDownArrow_Button_Down)
            End With
            With btnIncreaser
                '.ControlImage_Disabled = oUILib.oResMgr.GetTexture("VScr_Inc_Btn_Disabled.bmp")
                .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eUpArrow_Button_Disabled)
                '.ControlImage_Normal = oUILib.oResMgr.GetTexture("VScr_Inc_Btn_Normal.bmp")
                .ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eUpArrow_Button_Normal)
                '.ControlImage_Pressed = oUILib.oResMgr.GetTexture("VScr_Inc_Btn_Pressed.bmp")
                .ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eUpArrow_Button_Down)
            End With
            With btnScroller
                '.ControlImage_Disabled = oUILib.oResMgr.GetTexture("VScr_Scr_Btn_Disabled.bmp")
                .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eSmall_Button_Disabled)
                '.ControlImage_Normal = oUILib.oResMgr.GetTexture("VScr_Scr_Btn_Normal.bmp")
                .ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eSmall_Button_Normal)
                '.ControlImage_Pressed = oUILib.oResMgr.GetTexture("VScr_Scr_Btn_Pressed.bmp")
                .ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eSmall_Button_Down)
            End With
        Else
            With btnDecreaser
                '.ControlImage_Disabled = oUILib.oResMgr.GetTexture("HScr_Dec_Btn_Disabled.bmp")
                .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eLeftArrow_Button_Disabled)
                '.ControlImage_Normal = oUILib.oResMgr.GetTexture("HScr_Dec_Btn_Normal.bmp")
                .ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eLeftArrow_Button_Normal)
                '.ControlImage_Pressed = oUILib.oResMgr.GetTexture("HScr_Dec_Btn_Pressed.bmp")
                .ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eLeftArrow_Button_Down)
            End With
            With btnIncreaser
                '.ControlImage_Disabled = oUILib.oResMgr.GetTexture("HScr_Inc_Btn_Disabled.bmp")
                .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eRightArrow_Button_Disabled)
                '.ControlImage_Normal = oUILib.oResMgr.GetTexture("HScr_Inc_Btn_Normal.bmp")
                .ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eRightArrow_Button_Normal)
                '.ControlImage_Pressed = oUILib.oResMgr.GetTexture("HScr_Inc_Btn_Pressed.bmp")
                .ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eRightArrow_Button_Down)
            End With
            With btnScroller
                '.ControlImage_Disabled = oUILib.oResMgr.GetTexture("HScr_Scr_Btn_Disabled.bmp")
                .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eSmall_Button_Disabled)
                '.ControlImage_Normal = oUILib.oResMgr.GetTexture("HScr_Scr_Btn_Normal.bmp")
                .ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eSmall_Button_Normal)
                '.ControlImage_Pressed = oUILib.oResMgr.GetTexture("HScr_Scr_Btn_Pressed.bmp")
                .ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eSmall_Button_Down)
            End With
        End If
        IsDirty = True
    End Sub

	Private Sub btnDecreaser_Click(ByVal sName As String) Handles btnDecreaser.Click
        Dim lNewVal As Int32 = Value
        If ReverseDirection = False Then
            lNewVal -= SmallChange
            If lNewVal < MinValue Then lNewVal = MinValue
        Else
            lNewVal += SmallChange
            If lNewVal > MaxValue Then lNewVal = MaxValue
        End If
        If Me.Value <> lNewVal Then
            Me.Value = lNewVal
            RaiseEvent ValueChange()
        End If
	End Sub

	Private Sub btnIncreaser_Click(ByVal sName As String) Handles btnIncreaser.Click
        Dim lNewVal As Int32 = Value
        If ReverseDirection = False Then
            lNewVal += SmallChange
            If lNewVal > MaxValue Then lNewVal = MaxValue
        Else
            lNewVal -= SmallChange
            If lNewVal < MinValue Then lNewVal = MinValue
        End If
        If Me.Value <> lNewVal Then
            Me.Value = lNewVal
            RaiseEvent ValueChange()
        End If
	End Sub

    Public Sub btnScroller_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles btnScroller.OnMouseDown
        'get our mouse X and mouse Y on this control
        Dim oLoc As Point = GetAbsolutePosition()
        mlScrollingMouseX = lMouseX - oLoc.X
        mlScrollingMouseY = lMouseY - oLoc.Y
        mbScrolling = True
        MyBase.moUILib.lUISelectState = UILib.eSelectState.eScrollBarDrag
        MyBase.moUILib.oScrollingBar = Me
        oLoc = Nothing
    End Sub

    Public Sub btnScroller_OnMouseUp(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles btnScroller.OnMouseUp
        mbScrolling = False
        MyBase.moUILib.lUISelectState = UILib.eSelectState.eNoSelectState
        MyBase.moUILib.oScrollingBar = Nothing
        btnScroller.ControlImageRect = btnScroller.ControlImageRect_Normal
        btnScroller.UIButton_OnMouseUp(lMouseX, lMouseY, lButton)
    End Sub

    Public Sub btnScroller_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles btnScroller.OnMouseMove
        If mbScrolling Then
            'ok, user is dragging scrollbar
            Dim lX As Int32
            Dim lY As Int32
            Dim lTemp As Int32

            Dim lMinExt As Int32
            Dim lMaxExt As Int32

            Dim fMult As Single = 0.0F

            If lButton = MouseButtons.Left Then

                If Me.IsVertical Then
                    'Ok, we only care about the MouseY, so get our extents
                    lX = btnDecreaser.GetAbsolutePosition.Y
                    lY = btnIncreaser.GetAbsolutePosition.Y

                    'Regardless of direction, 'btnIncreaser' is always the top button
                    lY += btnIncreaser.Height

                    lMinExt = Math.Min(lX, lY)
                    lMaxExt = Math.Max(lX, lY)

                    lY = lMouseY - lMinExt
                    lTemp = lMaxExt - lMinExt
                    If lY < 0 Then
                        'then Value is at top... 
                        fMult = 1
                        If Me.ReverseDirection = True Then
                            fMult = 0
                        End If
                    ElseIf lY > lTemp Then
                        'then value is at bottom...
                        fMult = 0
                        If Me.ReverseDirection = True Then
                            fMult = 1
                        End If
                    Else
                        'value is in between
                        'lTemp = lMaxExt - lMinExt
                        fMult = CSng(lY / lTemp)
                        If Me.ReverseDirection = False Then fMult = 1.0F - fMult
                    End If
                Else
                    'Ok, only care about mouse X
                    lX = btnDecreaser.GetAbsolutePosition.X
                    lY = btnIncreaser.GetAbsolutePosition.X

                    'Regardless of direction, 'btnIncreaser' is on the right
                    lX += btnDecreaser.Height

                    lMinExt = Math.Min(lX, lY)
                    lMaxExt = Math.Max(lX, lY)

                    lX = lMouseX - lMinExt
                    lTemp = lMaxExt - lMinExt
                    If lX < 0 Then
                        'then Value is at top... 
                        fMult = 0
                        If Me.ReverseDirection = True Then
                            fMult = 1
                        End If
                    ElseIf lX > lTemp Then
                        'then value is at bottom...
                        fMult = 1
                        If Me.ReverseDirection = True Then
                            fMult = 0
                        End If
                    Else
                        'value is in between
                        'lTemp = lMaxExt - lMinExt
                        fMult = CSng(lX / lTemp)
                        If Me.ReverseDirection = True Then fMult = 1.0F - fMult
                    End If
                End If

                lTemp = Me.MaxValue - Me.MinValue
                Dim lNewVal As Int32 = Me.MinValue + CInt(fMult * lTemp)
                If Me.Value <> lNewVal Then
                    Me.Value = lNewVal
                    RaiseEvent ValueChange()
                End If
            End If
        End If

    End Sub

	Private Sub UIScrollBar_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown

		If NewTutorialManager.TutorialOn = True AndAlso MyBase.moUILib.CommandAllowed(True, GetFullName()) = False Then
			Return
		End If

		'Adjust for our scrollbar
		lMouseX -= Me.GetAbsolutePosition.X
		lMouseY -= Me.GetAbsolutePosition.Y

		Dim lNewValue As Int32 = Me.Value

		If IsVertical = True Then
			'Ok, VScroll
			If Me.ReverseDirection = True Then
				'Increase Bottom, Decrease Top
				If lMouseY < btnScroller.Top Then
					'decrease
					lNewValue -= LargeChange
				Else
					'increase
					lNewValue += LargeChange
				End If
			Else
				'Increase Top, Decrease Bottom
				If lMouseY < btnScroller.Top Then
					'Increase
					lNewValue += LargeChange
				Else
					'Decrease
					lNewValue -= LargeChange
				End If
			End If
		Else
			'Ok, HScroll
			If Me.ReverseDirection = True Then
				'Increase Left, Decrease Right
				If lMouseX < btnScroller.Left Then
					'Increase
					lNewValue += LargeChange
				Else
					'Decrease
					lNewValue -= LargeChange
				End If
			Else
				'Increase Right, Decrease Left
				If lMouseX < btnScroller.Left Then
					'Decrease
					lNewValue -= LargeChange
				Else
					'Increase
					lNewValue += LargeChange
				End If
			End If
		End If

		If lNewValue < Me.MinValue Then lNewValue = Me.MinValue
		If lNewValue > Me.MaxValue Then lNewValue = Me.MaxValue
		Me.Value = lNewValue
		Me.IsDirty = True
		RaiseEvent ValueChange()
    End Sub

    Private Sub UIScrollBar_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        If mbScrolling = True Then btnScroller_OnMouseMove(lMouseX, lMouseY, lButton)
    End Sub

    Private Sub UIScrollBar_OnMouseWheel(ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.OnMouseWheel
        Dim lNewVal As Int32 = Me.Value '+ 1 ' e.Delta
        If e.Delta < 0 Then
            If Me.mbReverseDirection = True Then
                lNewVal += Me.SmallChange
            Else
                lNewVal -= Me.SmallChange
            End If
        Else
            If Me.mbReverseDirection = True Then
                lNewVal -= Me.SmallChange
            Else
                lNewVal += Me.SmallChange
            End If
            'lNewVal += Me.LargeChange Else lNewVal -= Me.LargeChange
        End If
        If lNewVal < Me.MinValue Then lNewVal = Me.MinValue
        If lNewVal > Me.MaxValue Then lNewVal = Me.MaxValue
        If Me.Value <> lNewVal Then
            Me.Value = lNewVal
            RaiseEvent ValueChange()
        End If

        'Me.IsDirty = True
    End Sub

    'Private Sub UIScrollBar_OnRender(ByRef oImgSprite As Sprite, ByRef oTextSprite As Sprite) Handles MyBase.OnRender
    Private Sub UIScrollBar_OnRender() Handles MyBase.OnRender
        Dim oLoc As Point = GetAbsolutePosition()
		Dim lTemp As Int32

		If mbVertical Then
			'increaser is at the top, decreaser is at the bottom
			With btnIncreaser
				.Top = 0
				.Left = 0
				.Width = Me.Width
				.Height = GetOtherButtonSize()
			End With
			With btnDecreaser
				.Top = Me.Height - GetOtherButtonSize()
				.Left = 0
				.Width = Me.Width
				.Height = GetOtherButtonSize()
			End With

			'the scroller moves up and down
			With btnScroller
				.Width = Me.Width
                .Height = GetOtherButtonSize()

                If btnIncreaser.Height + btnDecreaser.Height + .Height + 5 > Me.Height Then
                    .Height = Math.Max(2, Me.Height - (btnIncreaser.Height + btnDecreaser.Height) - 5)
                End If

				.Left = 0

				lTemp = btnIncreaser.Top + btnIncreaser.Height
				lTemp = btnDecreaser.Top - lTemp

				'ltemp now how distance between buttons... remove the size of our scroll button
				lTemp = lTemp - .Height

				'now, the remaining distance tells us our ratio
				If MaxValue = 0 Then
					lTemp = 0
				Else
					lTemp = CInt((Value / MaxValue) * lTemp)
				End If
				If ReverseDirection = False Then
					.Top = btnDecreaser.Top - lTemp - .Height
				Else
                    .Top = (btnIncreaser.Top + btnIncreaser.Height) + lTemp
				End If

			End With
		Else
			'increaser is on the right, decreaser is on the left
			With btnIncreaser
				.Top = 0
				.Left = Me.Width - GetOtherButtonSize()
				.Width = GetOtherButtonSize()
				.Height = Me.Height
			End With
			With btnDecreaser
				.Top = 0
				.Left = 0
				.Width = GetOtherButtonSize()
				.Height = Me.Height
			End With

			'the scroller moves left and right
			With btnScroller
				.Width = GetOtherButtonSize()
                .Height = Me.Height

                If btnIncreaser.Width + btnDecreaser.Width + .Width + 5 > Me.Width Then
                    .Width = Math.Max(2, Me.Width - (btnIncreaser.Width + btnDecreaser.Width) - 5)
                End If

				.Top = 0

				lTemp = btnDecreaser.Left + btnDecreaser.Width
				lTemp = btnIncreaser.Left - lTemp

				'ltemp now how distance between buttons... remove the size of our scroll button
				lTemp = lTemp - .Width

				'now, the remaining distance tells us our ratio
				If MaxValue = 0 Then
					lTemp = 0
				Else
					lTemp = CInt((Value / MaxValue) * lTemp)
				End If
				If ReverseDirection = False Then
					.Left = btnDecreaser.Left + btnDecreaser.Width + lTemp
				Else
					.Left = btnIncreaser.Left - lTemp - .Width
				End If
			End With
		End If

        'Now, render my background
        'Do a color fill
        Dim oRect As Rectangle = New Rectangle(oLoc.X, oLoc.Y, Width, Height)
        If Enabled Then
            'moUILib.DoAlphaBlendColorFill_EX(oRect, System.Drawing.Color.Black, oLoc, oImgSprite)
            moUILib.DoAlphaBlendColorFill(oRect, System.Drawing.Color.Black, oLoc)
        Else
            'moUILib.DoAlphaBlendColorFill_EX(oRect, System.Drawing.Color.DarkGray, oLoc, oImgSprite)
            moUILib.DoAlphaBlendColorFill(oRect, System.Drawing.Color.DarkGray, oLoc)
        End If
        oRect = Nothing
		'end of color fill

		mlOtherButtonSize = Int32.MinValue
    End Sub

    Protected Overrides Sub Finalize()
        RemoveAllChildren()
        btnDecreaser = Nothing
        btnIncreaser = Nothing
        btnScroller = Nothing
        MyBase.Finalize()
    End Sub

    Private Sub btnIncreaser_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles btnIncreaser.OnMouseMove
        If MyBase.moUILib.lUISelectState = UILib.eSelectState.eScrollBarDrag Then
            btnScroller_OnMouseMove(lMouseX, lMouseY, lButton)
        End If
    End Sub

    Private Sub btnDecreaser_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles btnDecreaser.OnMouseMove
        If MyBase.moUILib.lUISelectState = UILib.eSelectState.eScrollBarDrag Then
            btnScroller_OnMouseMove(lMouseX, lMouseY, lButton)
        End If
    End Sub

	Private mlOtherButtonSize As Int32 = Int32.MinValue
	Private Const ml_OTHER_BTN_SIZE As Int32 = 24
	Private Function GetOtherButtonSize() As Int32
		If mlOtherButtonSize <> Int32.MinValue Then Return mlOtherButtonSize
		mlOtherButtonSize = ml_OTHER_BTN_SIZE
		If btnScroller Is Nothing = False AndAlso btnDecreaser Is Nothing = False AndAlso btnIncreaser Is Nothing = False Then
			If mbVertical = True Then
				Dim lTotal As Int32 = ml_OTHER_BTN_SIZE + ml_OTHER_BTN_SIZE + ml_OTHER_BTN_SIZE
				If lTotal + 4 > Me.Height Then
					'ok, need to resize our btns
					lTotal = Me.Height - 4
					mlOtherButtonSize = lTotal \ 3
				End If
			Else
				Dim lTotal As Int32 = ml_OTHER_BTN_SIZE + ml_OTHER_BTN_SIZE + ml_OTHER_BTN_SIZE
				If lTotal + 4 > Me.Width Then
					'ok, need to resize our btns
					lTotal = Me.Width - 4
					mlOtherButtonSize = lTotal \ 3
				End If
			End If
		End If
		Return mlOtherButtonSize
	End Function
 
End Class
