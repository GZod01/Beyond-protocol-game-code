Option Strict On

'this class is responsible for all user interfaces...
Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D


Public Class UILine
    Inherits UIControl

    'Private Shared moLinePen As Line
    'Public Shared Sub ReleasePen()
    '    If moLinePen Is Nothing = False Then moLinePen.Dispose()
    '    moLinePen = Nothing
    'End Sub
    Private Shared muVecs(1) As Vector2

    Public BorderColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 255)
    Public LineWidth As Int32 = 2

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)
    End Sub

    'Public Sub Control_OnRender(ByRef oImgSprite As Sprite, ByRef oTextSprite As Sprite) Handles MyBase.OnRender
    Public Sub Control_OnRender() Handles MyBase.OnRender
        Dim oLoc As System.Drawing.Point

        oLoc = GetAbsolutePosition()

        muVecs(0).X = oLoc.X
        muVecs(0).Y = oLoc.Y

        muVecs(1).X = oLoc.X + Width
        muVecs(1).Y = oLoc.Y + Height

        'If moLinePen Is Nothing OrElse moLinePen.Disposed = True Then
        '	Device.IsUsingEventHandlers = False
        '	moLinePen = New Line(moUILib.oDevice)
        '	Device.IsUsingEventHandlers = True
        'End If

        '      moLinePen.Antialias = True
        'moLinePen.Begin()
        '      moLinePen.Draw(muVecs, BorderColor)
        '      moLinePen.End() 
        BPLine.DrawLine(LineWidth, True, muVecs, BorderColor)

    End Sub

    Private Sub UILine_ResetInterfaceColors(ByVal yType As Byte, ByVal clrPrev As System.Drawing.Color) Handles Me.ResetInterfaceColors
        Select Case yType
            Case 1      'border
                If BorderColor = clrPrev Then BorderColor = muSettings.InterfaceBorderColor
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
