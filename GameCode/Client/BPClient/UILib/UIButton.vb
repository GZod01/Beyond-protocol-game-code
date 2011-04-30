Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class UIButton
    Inherits UILabel

    'Public ControlImage_Pressed As Texture
    Public ControlImageRect_Pressed As Rectangle
    Protected mbPressed As Boolean = False

	Private moAbsPos As Point
	Private mlLastClick As Int32

	Public Event Click(ByVal sName As String)

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'now, always load the default images
        ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eButton_Down)
        'ControlImage_Pressed = oUILib.oResMgr.GetTexture("Button_Pressed.bmp")
        ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eButton_Normal)
        'ControlImage_Normal = oUILib.oResMgr.GetTexture("Button_Normal.bmp")
        ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eButton_Disabled)
        'ControlImage_Disabled = oUILib.oResMgr.GetTexture("Button_Disabled.bmp")

        Me.DrawBackImage = True
        MyBase.bAcceptEvents = True
        Me.BackImageColor = muSettings.InterfaceButtonColor

        Me.ForeColor = muSettings.InterfaceBorderColor
    End Sub

    Private moBackImageColor_Orig As System.Drawing.Color
    Private moBackImageColor_Disabled As System.Drawing.Color
    Public Shadows Property BackImageColor() As System.Drawing.Color
        Get
            Return MyBase.BackImageColor
        End Get
        Set(ByVal value As System.Drawing.Color)
            moBackImageColor_Orig = value
            MyBase.BackImageColor = value

            Dim lMinValNot0 As Int32 = 255
            If value.R < lMinValNot0 AndAlso value.R <> 0 Then lMinValNot0 = value.R
            If value.G < lMinValNot0 AndAlso value.G <> 0 Then lMinValNot0 = value.G
            If value.B < lMinValNot0 AndAlso value.B <> 0 Then lMinValNot0 = value.B
            If lMinValNot0 > 128 Then lMinValNot0 = 128
            If lMinValNot0 < 1 Then lMinValNot0 = 1
            moBackImageColor_Disabled = System.Drawing.Color.FromArgb(255, lMinValNot0, lMinValNot0, lMinValNot0)
        End Set
    End Property
 
    Public Overrides Sub Control_OnRender() 'Handles MyBase.OnRender
        'render the button
        If Me.Enabled = True Then
            MyBase.moBackImageColor = moBackImageColor_Orig
            If mbPressed Then
                'ControlImage = ControlImage_Pressed
                ControlImageRect = ControlImageRect_Pressed
            Else
                'ControlImage = ControlImage_Normal
                ControlImageRect = ControlImageRect_Normal
            End If
        Else
            'ControlImage = ControlImage_Disabled
            MyBase.moBackImageColor = moBackImageColor_Disabled
            ControlImageRect = ControlImageRect_Disabled
        End If

        'and then render the label
        'MyBase.Control_OnRender(oImgSprite, oTextSprite)
        MyBase.Control_OnRender()
    End Sub

    Public Sub ResetPressed()
        mbPressed = False
        Me.IsDirty = True
    End Sub

	Private Sub UIButton_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseDown
		mbPressed = True
        IsDirty = True
        MyBase.moUILib.oButtonDown = Me
    End Sub

    Public Sub UIButton_OnMouseUp(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseUp
        Dim bRaise As Boolean = mbPressed
        mbPressed = False
        If bRaise Then
			IsDirty = True

			If glCurrentCycle - mlLastClick > 4 Then

				If NewTutorialManager.TutorialOn = True Then
					If MyBase.moUILib.CommandAllowed(True, GetFullName()) = False Then
						Return
					End If
				End If

				mlLastClick = glCurrentCycle
				If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
				RaiseEvent Click(Me.ControlName)
			End If
		End If
    End Sub

    Protected Overrides Sub Finalize()
        Me.Visible = False
        'ControlImage_Pressed = Nothing
        MyBase.Finalize()
    End Sub

    Private Sub UIButton_ReprocessInput() Handles MyBase.ReprocessInput
		If mbPressed Then RaiseEvent Click(Me.ControlName)
    End Sub

    Private Sub UIButton_ResetInterfaceColors(ByVal yType As Byte, ByVal clrPrev As System.Drawing.Color) Handles Me.ResetInterfaceColors
        Select Case yType
            Case 1      'border
                'If BorderColor = clrPrev Then BorderColor = muSettings.InterfaceBorderColor
                If ForeColor = clrPrev Then ForeColor = muSettings.InterfaceBorderColor
            Case 2      'fill
                'If FillColor = clrPrev Then FillColor = muSettings.InterfaceFillColor
            Case 3      'textboxfore
                'unused
            Case 4      'textboxfill
                'unused
            Case 5      'button
                If Me.BackImageColor = clrPrev Then Me.BackImageColor = muSettings.InterfaceButtonColor
        End Select
    End Sub
End Class
