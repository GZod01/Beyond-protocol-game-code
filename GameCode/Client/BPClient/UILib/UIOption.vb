Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class UIOption
    Inherits UILabel

    Public Enum eOptionTypes As Integer
        eLargeOption = 0        ' these are now the same
        eSmallOption
    End Enum

    Private mbValue As Boolean = False
    Public Property Value() As Boolean
        Get
            Return mbValue
        End Get
        Set(ByVal Value As Boolean)
            mbValue = Value
            IsDirty = True
        End Set
    End Property
    Public ControlImageRect_Pressed As Rectangle

    Protected mbPressed As Boolean = False
	Private mlType As eOptionTypes = eOptionTypes.eSmallOption

	Public Locked As Boolean = False

    Public Overridable Property DisplayType() As eOptionTypes
        Get
            Return mlType
        End Get
        Set(ByVal Value As eOptionTypes)
            mlType = Value

            'If mlType = eOptionTypes.eSmallOption Then
            ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eOption_Disabled)
            ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eOption_Normal)
            ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eOption_Marked)
            'Else
            '    ControlImageRect_Disabled = System.Drawing.Rectangle.FromLTRB(72, 96, 95, 119)
            '    ControlImageRect_Normal = System.Drawing.Rectangle.FromLTRB(96, 96, 119, 119)
            '    ControlImageRect_Pressed = System.Drawing.Rectangle.FromLTRB(120, 96, 143, 119)
            'End If

        End Set
    End Property

    Public Event Click()

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        DisplayType = eOptionTypes.eLargeOption
        MyBase.bAcceptEvents = True
        Me.ForeColor = System.Drawing.Color.White
    End Sub

    'Public Overrides Sub Control_OnRender(ByRef oImgSprite As Sprite, ByRef oTextSprite As Sprite) Handles MyBase.OnRender
    Public Overrides Sub Control_OnRender() 'Handles MyBase.OnRender
        'render the button
        Dim oLoc As Point = GetAbsolutePosition()
        Dim oRect As Rectangle

        Dim clrVal As System.Drawing.Color = muSettings.InterfaceBorderColor


        If Value Then
            ControlImageRect = ControlImageRect_Pressed
        Else
            ControlImageRect = ControlImageRect_Normal
        End If
        If Me.Enabled = False Then
            'ControlImageRect = ControlImageRect_Disabled
            clrVal = System.Drawing.Color.FromArgb(clrVal.A \ 2, clrVal.R \ 2, clrVal.G \ 2, clrVal.B \ 2)
        End If

        'oLoc = GetAbsolutePosition()
        With oRect
            .Location() = oLoc
            .Width = Width
            .Height = Height

            Dim lOptOffsetY As Int32 = CInt((Height - ControlImageRect.Height) / 2)
            If moUILib.oInterfaceTexture Is Nothing = True Then Return

            'render the back image if there is one
            Dim bNeedToEnd As Boolean = False
            Dim bNeedToEndText As Boolean = False
            With moUILib.Pen
                'With oImgSprite
                '.Begin(SpriteFlags.AlphaBlend)
                moUILib.BeginPenSprite(SpriteFlags.AlphaBlend)
                .Draw2D(moUILib.oInterfaceTexture, ControlImageRect, System.Drawing.Rectangle.FromLTRB(oLoc.X, oLoc.Y + lOptOffsetY, oLoc.X + ControlImageRect.Width, oLoc.Y + lOptOffsetY + ControlImageRect.Height), System.Drawing.Point.Empty, 0, New Point(oLoc.X, oLoc.Y + lOptOffsetY), clrVal)
                '.End()
                moUILib.EndPenSprite()
            End With
            'Breaks too many checkboxes, will review later.
            'FontFormat = DrawTextFormat.Left
            'oRect.X = oRect.X + 12
            'oRect.Y = oRect.Y + 1
            If Caption <> "" Then
                BPFont.DrawText(moSysFont, Caption, oRect, FontFormat, ForeColor)
            End If
        End With
    End Sub

    Public Sub ResetPressed()
        mbPressed = False
        Me.IsDirty = True
    End Sub

    Private Sub UIButton_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseDown
        If Locked = False Then
            mbPressed = True
            MyBase.moUILib.oOptionDown = Me
        End If
    End Sub

    Public Sub UIButton_OnMouseUp(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseUp
        If mbPressed = True AndAlso Locked = False Then
            If NewTutorialManager.TutorialOn = True AndAlso MyBase.moUILib.CommandAllowed(True, GetFullName()) = False Then
                Return
            End If

            Value = Not Value
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            RaiseEvent Click()
        End If
        mbPressed = False
    End Sub

    Protected Overrides Sub Finalize()
        Me.Visible = False
        'ControlImage_Pressed = Nothing
        MyBase.Finalize()
    End Sub

    Private Sub UIOption_ResetInterfaceColors(ByVal yType As Byte, ByVal clrPrev As System.Drawing.Color) Handles Me.ResetInterfaceColors
        Select Case yType
            Case 1      'border
                If ForeColor = clrPrev Then ForeColor = muSettings.InterfaceBorderColor
            Case 2      'fill
                'unused
            Case 3      'textboxfore
                'unused
            Case 4      'textboxfill
                'unused
            Case 5      'button
                'unused
        End Select
    End Sub
End Class
