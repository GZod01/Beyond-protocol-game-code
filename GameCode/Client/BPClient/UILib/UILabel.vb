Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class UILabel
    Inherits UIControl

    Protected mbRaiseTextChangeEvent As Boolean = False

    Public Event TextChanged()

    Public bFilterBadWords As Boolean = True

    Protected msCaption As String
    Public Property Caption() As String
        Get
            Return msCaption
        End Get
        Set(ByVal Value As String)
            If bFilterBadWords = True Then Value = FilterBadWords(Value)
            msCaption = Value
            IsDirty = True
            If mbRaiseTextChangeEvent = True Then RaiseEvent TextChanged()
        End Set
    End Property
    'Private moTrueFont As Font = Nothing
    'Protected Property moFont() As Font
    '    Get
    '        If moTrueFont Is Nothing = True OrElse moTrueFont.Disposed = True Then
    '            SetFont(moSysFont)
    '        End If
    '        Return moTrueFont
    '    End Get
    '    Set(ByVal value As Font)
    '        moTrueFont = value
    '    End Set
    'End Property
    Protected moSysFont As System.Drawing.Font
    Private moForeColor As System.Drawing.Color = System.Drawing.Color.Black
    Public Property ForeColor() As System.Drawing.Color
        Get
            Return moForeColor
        End Get
        Set(ByVal Value As System.Drawing.Color)
            moForeColor = Value
            IsDirty = True
        End Set
    End Property
    Private moFontFormat As DrawTextFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
    Public Property FontFormat() As DrawTextFormat
        Get
            Return moFontFormat
        End Get
        Set(ByVal Value As DrawTextFormat)
            moFontFormat = Value
            IsDirty = True
        End Set
    End Property
    Private mbDrawBackImage As Boolean = False
    Public Property DrawBackImage() As Boolean
        Get
            Return mbDrawBackImage
        End Get
        Set(ByVal Value As Boolean)
            mbDrawBackImage = Value
            IsDirty = True
        End Set
    End Property
    Protected moBackImageColor As System.Drawing.Color = muSettings.InterfaceBorderColor
    Public Property BackImageColor() As System.Drawing.Color
        Get
            Return moBackImageColor
        End Get
        Set(ByVal value As System.Drawing.Color)
            moBackImageColor = value
            IsDirty = True
        End Set
    End Property

    Public Sub ReleaseFont()
        'If moTrueFont Is Nothing = False Then moTrueFont.Dispose()
        'moTrueFont = Nothing
    End Sub

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'Default Font
        moSysFont = New System.Drawing.Font("MS Sans Serif", 10)

        MyBase.bAcceptEvents = False

    End Sub

    Public Sub SetFont(ByVal oFont As System.Drawing.Font)
        If moUILib.oDevice Is Nothing Then Return
        moSysFont = oFont
        IsDirty = True
    End Sub

    Public Function GetFont() As System.Drawing.Font
        Return moSysFont
    End Function

    'Public Overridable Sub Control_OnRender(ByRef oImgSprite As Sprite, ByRef oTextSprite As Sprite) Handles MyBase.OnRender
    Public Overridable Sub Control_OnRender() Handles MyBase.OnRender
        Dim oRect As Rectangle
        Dim oLoc As System.Drawing.Point

        Dim fMultX As Single
        Dim fMultY As Single

        oLoc = GetAbsolutePosition()
        With oRect
            .Location() = oLoc
            .Width = Width
            .Height = Height

            'render the back image if there is one
            If DrawBackImage = True AndAlso moUILib.oInterfaceTexture Is Nothing = False Then
                With moUILib.Pen
                    'With oImgSprite
                    '.Begin(SpriteFlags.AlphaBlend)
                    moUILib.BeginPenSprite(SpriteFlags.AlphaBlend)
                    fMultX = CSng(ControlImageRect.Width / Width)
                    fMultY = CSng(ControlImageRect.Height / Height)
                    .Draw2D(moUILib.oInterfaceTexture, ControlImageRect, System.Drawing.Rectangle.FromLTRB(oLoc.X, oLoc.Y, oLoc.X + Width, oLoc.Y + Height), System.Drawing.Point.Empty, 0, New Point(CInt(oLoc.X * fMultX), CInt(oLoc.Y * fMultY)), moBackImageColor)
                    '.End()
                    moUILib.EndPenSprite()
                End With
            End If
			If Caption <> "" Then
                BPFont.DrawText(moSysFont, Caption, oRect, FontFormat, ForeColor) 'BPFont.CFFont.DrawText(Caption, oRect, FontFormat, ForeColor, moSysFont.Size)
			End If
            'If Caption <> "" Then moFont.DrawText(oTextSprite, Caption, oRect, FontFormat, ForeColor)
        End With
    End Sub

    Public Function GetTextDimensions() As Rectangle
        'returns point filled with the width and height of the caption
        Return BPFont.MeasureString(moSysFont, Caption, FontFormat) 'BPFont.CFFont.GetTextDimensions(Caption, moSysFont.Size)
    End Function

    Public Function GetTextDimensions(ByVal sText As String) As Rectangle
        Return BPFont.MeasureString(moSysFont, sText, FontFormat) ' moFont.MeasureString(moUILib.Pen, sText, FontFormat, ForeColor)
    End Function

    Protected Overrides Sub Finalize()
        If moSysFont Is Nothing = False Then moSysFont.Dispose()
        moSysFont = Nothing
        MyBase.Finalize()
    End Sub

    Public Sub ManualDrawText(ByRef oSprite As Sprite)
        'assumes oSprite is already Begun
        If Me.Visible = True AndAlso Caption <> "" Then
            Dim oRect As Rectangle
            With oRect
                .Location() = GetAbsolutePosition()
                .Width = Width
                .Height = Height
            End With
            'moFont.DrawText(oSprite, Caption, oRect, FontFormat, ForeColor)
            BPFont.DrawText(moSysFont, Caption, oRect, FontFormat, ForeColor) 'BPFont.CFFont.DrawText(Caption, oRect, FontFormat, ForeColor, moSysFont.Size)
        End If
    End Sub

    Public Sub ManualDrawSprite(ByRef oSprite As Sprite)
        'assumes oSprite is already begun
        If Me.Visible = True Then
            Dim fMultX As Single = CSng(ControlImageRect.Width / Width)
            Dim fMultY As Single = CSng(ControlImageRect.Height / Height)
            Dim oLoc As Point = GetAbsolutePosition()
            oSprite.Draw2D(moUILib.oInterfaceTexture, ControlImageRect, System.Drawing.Rectangle.FromLTRB(oLoc.X, oLoc.Y, oLoc.X + Width, oLoc.Y + Height), System.Drawing.Point.Empty, 0, New Point(CInt(oLoc.X * fMultX), CInt(oLoc.Y * fMultY)), moBackImageColor)
        End If
    End Sub

    Private Sub UILabel_ResetInterfaceColors(ByVal yType As Byte, ByVal clrPrev As System.Drawing.Color) Handles Me.ResetInterfaceColors
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
