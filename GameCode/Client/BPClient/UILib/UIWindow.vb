Option Strict On

'this class is responsible for all user interfaces...
Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class UIWindow
    Inherits UIControl

    Public lWindowMetricID As Int32 = -1     'the id of the metric lookup for this window

    'windows can move
    Private mlMouseDownX As Int32
    Private mlMouseDownY As Int32

    'Protected Shared moBorderLine As Line

    Public FillColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 64, 128, 192)
    Public BorderColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 255)

    Public FullScreen As Boolean = False
    Public Moveable As Boolean = True

    Public DrawBorder As Boolean = True

    Public Resizable As Boolean = False
    Public ResizeLimits As System.Drawing.Rectangle = New System.Drawing.Rectangle(128, 64, 1024, 1024) 'Min-Width, Min-Height, Max-Width, Max-Height

	Protected Enum ResizeType As Int32
		NoResizeType = 0
		Left = 1
		Right = 2
		Up = 4
		Down = 8
	End Enum
	Protected mlResizeType As ResizeType
	Private mlResizeMouseX As Int32
	Private mlResizeMouseY As Int32

    Public BorderLineWidth As Int32 = 4

    Public FillWindow As Boolean = True

    Public Caption As String = ""

    Public Event WindowMoved()

    Private mlBaseWindowAlpha As Int32 = Int32.MinValue
    Private mlAlphaModifier As Int32
    Private mlLastUpdate As Int32 = 0

    Protected Event WindowClosed()

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)
        '      Device.IsUsingEventHandlers = False
        '      moBorderLine = New Line(oUILib.oDevice)
        'Device.IsUsingEventHandlers = True
	End Sub

    Private Sub UIWindow_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseDown
        'get our mouse X and mouse Y on this control
        Dim oLoc As Point = GetAbsolutePosition()
        mlMouseDownX = lMouseX - oLoc.X
		mlMouseDownY = lMouseY - oLoc.Y

		If Me.Resizable = True Then
			mlResizeType = ResizeType.NoResizeType
			If mlMouseDownX < 2 Then
				mlResizeType = mlResizeType Or ResizeType.Left
			ElseIf mlMouseDownX > Me.Width - 2 Then
				mlResizeType = mlResizeType Or ResizeType.Right
			End If
			If mlMouseDownY < 2 Then
				mlResizeType = mlResizeType Or ResizeType.Up
			ElseIf mlMouseDownY > Me.Height - 2 Then
				mlResizeType = mlResizeType Or ResizeType.Down
			End If
			If mlResizeType <> ResizeType.NoResizeType Then
				MyBase.moUILib.oMovingWindow = Me
				MyBase.moUILib.lUISelectState = UILib.eSelectState.eResizeWindow
				mlResizeMouseX = lMouseX
                mlResizeMouseY = lMouseY
                oLoc = Nothing
                Exit Sub
			End If
        End If
        MyBase.moUILib.oMovingWindow = Me

        oLoc = Nothing
    End Sub

    Public Sub UIWindow_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseMove
        If MyBase.moUILib.oButtonDown Is Nothing = False OrElse MyBase.moUILib.oOptionDown Is Nothing = False Then Return

        If lButton = MouseButtons.Left AndAlso Moveable = True AndAlso mlResizeType = ResizeType.NoResizeType _
        AndAlso MyBase.moUILib.oMovingWindow Is Nothing = False AndAlso MyBase.moUILib.oMovingWindow.ControlName = Me.ControlName Then
            MyBase.moUILib.lUISelectState = UILib.eSelectState.eMoveWindow
            MyBase.moUILib.oMovingWindow = Me
            'ok, user is dragging window..
            Dim oLoc As Point = GetAbsolutePosition()
            Dim lX As Int32
            Dim lY As Int32

            lX = oLoc.X - Me.Left
            lY = oLoc.Y - Me.Top

            Me.Left = lMouseX - lX - mlMouseDownX
            If Me.Left < 0 Then
                Me.Left = 0
            ElseIf Me.Left + Me.Width > moUILib.oDevice.PresentationParameters.BackBufferWidth Then
                Me.Left = moUILib.oDevice.PresentationParameters.BackBufferWidth - Me.Width
            End If

            Me.Top = lMouseY - lY - mlMouseDownY
            If Me.Top < 0 Then
                Me.Top = 0
            ElseIf Me.Top + Me.Height > moUILib.oDevice.PresentationParameters.BackBufferHeight Then
                Me.Top = moUILib.oDevice.PresentationParameters.BackBufferHeight - Me.Height
            End If

            RaiseEvent WindowMoved()
        ElseIf lButton = MouseButtons.None AndAlso Resizable = True Then
            Dim ptLoc As Point = GetAbsolutePosition()
            Dim lTmpX As Int32 = lMouseX - ptLoc.X
            Dim lTmpY As Int32 = lMouseY - ptLoc.Y

            Dim bEW As Boolean = lTmpX < 2 OrElse lTmpX > Me.Width - 2
            Dim bNS As Boolean = lTmpY < 2 OrElse lTmpY > Me.Height - 2
            If bNS And bEW Then
                System.Windows.Forms.Cursor.Current = Cursors.SizeAll
            ElseIf bNS Then
                System.Windows.Forms.Cursor.Current = Cursors.SizeNS
            ElseIf bEW Then
                System.Windows.Forms.Cursor.Current = Cursors.SizeWE
            ElseIf mlResizeType = ResizeType.NoResizeType Then
                System.Windows.Forms.Cursor.Current = Cursors.Default
            End If
        End If
    End Sub

    'Private Sub UIWindow_OnNewFrame() Handles Me.OnNewFrame
    '    If mlBaseWindowAlpha = Int32.MinValue Then
    '        mlBaseWindowAlpha = FillColor.A
    '        mlAlphaModifier = 1
    '    End If
    '    If glCurrentCycle - mlLastUpdate > 15 Then
    '        mlLastUpdate = glCurrentCycle
    '        Dim lAlpha As Int32 = FillColor.A
    '        lAlpha += mlAlphaModifier
    '        If lAlpha > 255 Then
    '            lAlpha = 255
    '            mlAlphaModifier *= -1
    '        ElseIf lAlpha < mlBaseWindowAlpha Then
    '            lAlpha = mlBaseWindowAlpha
    '            mlAlphaModifier *= -1
    '        End If

    '        If FillColor.A <> lAlpha Then
    '            FillColor = System.Drawing.Color.FromArgb(lAlpha, FillColor.R, FillColor.G, FillColor.B)
    '            Me.IsDirty = True
    '        End If
    '    End If
    'End Sub

    'Private Sub UIWindow_OnRender(ByRef oImgSprite As Sprite, ByRef oTextSprite As Sprite) Handles MyBase.OnRender
    Private Sub UIWindow_OnRender() Handles MyBase.OnRender
        Dim oLoc As Point = GetAbsolutePosition()
        Dim oRect As Rectangle

        With oRect
            .Location() = oLoc
            .Width = Width
            .Height = Height
        End With

        'Do a color fill
        If FillWindow = True Then
            'If DrawBorder = True Then
            '    Dim lTmpOffSet As Int32 = BorderLineWidth \ 2
            '    Dim ptTemp As Point = oLoc
            '    ptTemp.X += (lTmpOffSet * 2)
            '    ptTemp.Y += lTmpOffSet

            '    Dim rcTemp As Rectangle = oRect
            '    rcTemp.Location = ptTemp
            '    rcTemp.Width -= (2 * lTmpOffSet)
            '    rcTemp.Height -= (2 * lTmpOffSet)
            '    'moUILib.DoAlphaBlendColorFill_EX(rcTemp, FillColor, ptTemp, oImgSprite)
            '    moUILib.DoAlphaBlendColorFill(rcTemp, FillColor, ptTemp)
            '    'Else : moUILib.DoAlphaBlendColorFill_EX(oRect, FillColor, oLoc, oImgSprite)
            'Else :
            moUILib.DoAlphaBlendColorFill(oRect, FillColor, oLoc)
            'End If
        End If
        'end of color fill

        If DrawBorder = True Then
            Dim lOffset As Int32 = CInt(Math.Floor(BorderLineWidth / 2))
            Dim lTopLineX As Int32 = 0

            'Draw a box border around our window...
            If Me.Caption.Length = 0 Then
                'With moBorderLineVerts(0)
                '    .X = oLoc.X + lOffset
                '    .Y = oLoc.Y + lOffset
                'End With
                lTopLineX = 0
            Else
                Using oFont As Font = New Font(MyBase.moUILib.oDevice, MyBase.moUILib.oFrameFont)
                    Dim rcTmp As Rectangle = oFont.MeasureString(Nothing, Caption, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                    'With moBorderLineVerts(0)
                    '    .X = oLoc.X + rcTmp.Width + 10 + lOffset
                    '    .Y = oLoc.Y + lOffset
                    'End With
                    lTopLineX = rcTmp.Width

                    If Caption <> "" Then
                        'If moUILib.bTextPenBegun = False Then moUILib.TextPen.Begin(SpriteFlags.AlphaBlend)
                        oFont.DrawText(Nothing, Caption, New Rectangle(oLoc.X + 5, oLoc.Y - 7, rcTmp.Width, rcTmp.Height), DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)
                        'If moUILib.bTextPenBegun = False Then moUILib.TextPen.End()
                    End If
                End Using
            End If

            RenderBorder(lTopLineX, 0)

            ''Draw a box border around our window...
            'With moBorderLineVerts(1)
            '    '.X = oLoc.X + Width
            '    .X = oLoc.X + Width - lOffset
            '    '.Y = oLoc.Y
            '    .Y = oLoc.Y + lOffset
            'End With
            'With moBorderLineVerts(2)
            '    '.X = oLoc.X + Width
            '    .X = oLoc.X + Width - lOffset
            '    '.Y = oLoc.Y + Height
            '    .Y = oLoc.Y + Height - lOffset
            'End With
            'With moBorderLineVerts(3)
            '    '.X = oLoc.X
            '    .X = oLoc.X + lOffset
            '    '.Y = oLoc.Y + Height
            '    .Y = oLoc.Y + Height - lOffset
            'End With
            'With moBorderLineVerts(4)
            '    '.X = oLoc.X
            '    .X = oLoc.X + lOffset
            '    '.Y = oLoc.Y
            '    .Y = oLoc.Y + lOffset
            'End With

            ''Draw a box border around our window...
            'With moBorderLine
            '    .Antialias = True
            '    .Width = BorderLineWidth
            '    .Begin()
            '    .Draw(moBorderLineVerts, BorderColor)
            '    .End()
            'End With
            ''End of the border drawing
        End If

        oLoc = Nothing
    End Sub

    'Private Sub UIWindow_OnResize() Handles MyBase.OnResize
    '    Dim oLoc As Point = GetAbsolutePosition()
    '    With moBorderLineVerts(0)
    '        .X = oLoc.X
    '        .Y = oLoc.Y
    '    End With
    '    With moBorderLineVerts(1)
    '        .X = oLoc.X + Width
    '        .Y = oLoc.Y
    '    End With
    '    With moBorderLineVerts(2)
    '        .X = oLoc.X + Width
    '        .Y = oLoc.Y + Height
    '    End With
    '    With moBorderLineVerts(3)
    '        .X = oLoc.X
    '        .Y = oLoc.Y + Height
    '    End With
    '    With moBorderLineVerts(4)
    '        .X = oLoc.X
    '        .Y = oLoc.Y
    '    End With
    'End Sub

    'Protected Overrides Sub Finalize()
    '    If moBorderLine Is Nothing = False Then moBorderLine.Dispose()
    '    moBorderLine = Nothing

    '    MyBase.Finalize()
    'End Sub
    Public bRoundedBorder As Boolean = True
    Protected Sub RenderBorder(ByVal lTopLineX As Int32, ByVal lAddToHeight As Int32)
        Dim v2Lines() As Vector2
        Dim oLoc As Point = GetAbsolutePosition()
        Dim lOffset As Int32 = CInt(Math.Floor(BorderLineWidth / 2))
        'Dim lOffset As Single = 0.5F

        If bRoundedBorder = True Then 'And True = False Then
            ReDim v2Lines(8)

            With v2Lines(0)
                .X = oLoc.X + lTopLineX + lOffset + 3
                If lTopLineX <> 0 Then lTopLineX += 10
                .Y = oLoc.Y + lOffset
            End With
            With v2Lines(1)
                .X = oLoc.X + Width - 3 - lOffset
                .Y = oLoc.Y + lOffset
            End With
            With v2Lines(2)
                .X = oLoc.X + Width - lOffset
                .Y = oLoc.Y + 3 + lOffset
            End With
            With v2Lines(3)
                .X = oLoc.X + Width - lOffset
                .Y = oLoc.Y + Height - 3 - lOffset + lAddToHeight
            End With
            With v2Lines(4)
                .X = oLoc.X + Width - 3 - lOffset
                .Y = oLoc.Y + Height - lOffset + lAddToHeight
            End With
            With v2Lines(5)
                .X = oLoc.X + 3 + lOffset
                .Y = oLoc.Y + Height - lOffset + lAddToHeight
            End With
            With v2Lines(6)
                .X = oLoc.X + lOffset
                .Y = oLoc.Y + Height - 3 - lOffset + lAddToHeight
            End With
            With v2Lines(7)
                .X = oLoc.X + lOffset
                .Y = oLoc.Y + 3 + lOffset
            End With
            With v2Lines(8)
                .X = oLoc.X + 3 + lOffset
                .Y = oLoc.Y + lOffset
            End With
        Else
            ReDim v2Lines(4)

            With v2Lines(0)
                .X = oLoc.X + lTopLineX + lOffset
                If lTopLineX <> 0 Then .X += 10
                .Y = oLoc.Y + lOffset
            End With
            With v2Lines(1)
                .X = oLoc.X + Width - lOffset
                .Y = oLoc.Y + lOffset
            End With
            With v2Lines(2)
                .X = oLoc.X + Width - lOffset
                .Y = oLoc.Y + Height - lOffset + lAddToHeight
            End With
            With v2Lines(3)
                .X = oLoc.X + lOffset
                .Y = oLoc.Y + Height - lOffset + lAddToHeight
            End With
            With v2Lines(4)
                .X = oLoc.X + lOffset
                .Y = oLoc.Y + lOffset
            End With
        End If

        'BEGIN OF THE SOFT LINE / SOFT BORDER ====================================
        'Dim oTex As Texture = goResMgr.GetTexture("Particle.dds", GFXResourceManager.eGetTextureType.NoSpecifics, "")
        'Dim oSpr As BPSprite = New BPSprite()
        'oSpr.BeginRender(4, oTex, MyBase.moUILib.oDevice) 
        'Dim rcSrc As Rectangle = New Rectangle(0, 0, 32, 32)
        'Dim lLineWd As Int32 = BorderLineWidth * 3
        'Dim lHlfWd As Int32 = lLineWd \ 2
        'Dim lQrtWd As Int32 = lHlfWd \ 2

        'With MyBase.moUILib.oDevice
        '    .RenderState.SourceBlend = Blend.SourceAlpha
        '    .RenderState.DestinationBlend = Blend.One
        '    .RenderState.AlphaBlendEnable = True
        'End With

        ''Top horizontal bar
        'rcSrc = New Rectangle(0, 16, 32, 16)
        'oSpr.Draw2D(rcSrc, New Rectangle(oLoc.X, oLoc.Y, Me.Width, lLineWd), BorderColor)

        ''Right vertical bar
        'rcSrc = New Rectangle(0, 0, 16, 32)
        'oSpr.Draw2D(rcSrc, New Rectangle((oLoc.X + Me.Width) - lLineWd, oLoc.Y, lLineWd, Me.Height), BorderColor)

        ''Left vertical bar
        'rcSrc = New Rectangle(16, 0, 16, 32)
        'oSpr.Draw2D(rcSrc, New Rectangle(oLoc.X, oLoc.Y, lLineWd, Me.Height), BorderColor)

        ''Bottom horizontal bar
        'rcSrc = New Rectangle(0, 0, 32, 16)
        'oSpr.Draw2D(rcSrc, New Rectangle(oLoc.X, (oLoc.Y + Me.Height) - lLineWd, Me.Width, lLineWd), BorderColor)

        'oSpr.EndRender()
        'oSpr.DisposeMe()

        'With MyBase.moUILib.oDevice
        '    .RenderState.SourceBlend = Blend.SourceAlpha 'Blend.One
        '    .RenderState.DestinationBlend = Blend.InvSourceAlpha 'Blend.Zero
        '    .RenderState.AlphaBlendEnable = True
        'End With
        'END OF SOFT LINE / SOFT BORDER =====================

         'Draw a box border around our window...
        'ValidateBorderLine()
        'Try
        '    With moBorderLine
        '        .Antialias = True
        '        .Width = BorderLineWidth
        '        .Begin()
        '        .Draw(v2Lines, BorderColor)
        '        .End()
        '    End With
        'Catch
        'End Try
        BPLine.DrawLine(1, True, v2Lines, BorderColor)
    End Sub

    'Protected Sub ValidateBorderLine()
    '    If moBorderLine Is Nothing = True OrElse moBorderLine.Disposed = True Then
    '        Device.IsUsingEventHandlers = False
    '        moBorderLine = New Line(MyBase.moUILib.oDevice)
    '        Device.IsUsingEventHandlers = True
    '    End If
    'End Sub
    'Public Shared Sub ReleaseBorderLine()
    '    If moBorderLine Is Nothing = False Then moBorderLine.Dispose()
    '    moBorderLine = Nothing
    'End Sub

    Protected Sub RenderRoundedBorder(ByVal rcDimensions As Rectangle, ByVal lBorderLineWidth As Int32, ByVal clrColor As System.Drawing.Color)
        Dim v2Lines(8) As Vector2
        Dim oLoc As Point = rcDimensions.Location
        Dim lOffset As Int32 = lBorderLineWidth \ 2

        With v2Lines(0)
            .X = oLoc.X + lOffset + 3
            .Y = oLoc.Y + lOffset
        End With
        With v2Lines(1)
            .X = oLoc.X + rcDimensions.Width - 3 - lOffset
            .Y = oLoc.Y + lOffset
        End With
        With v2Lines(2)
            .X = oLoc.X + rcDimensions.Width - lOffset
            .Y = oLoc.Y + 3 + lOffset
        End With
        With v2Lines(3)
            .X = oLoc.X + rcDimensions.Width - lOffset
            .Y = oLoc.Y + rcDimensions.Height - 3 - lOffset
        End With
        With v2Lines(4)
            .X = oLoc.X + rcDimensions.Width - 3 - lOffset
            .Y = oLoc.Y + rcDimensions.Height - lOffset
        End With
        With v2Lines(5)
            .X = oLoc.X + 3 + lOffset
            .Y = oLoc.Y + rcDimensions.Height - lOffset
        End With
        With v2Lines(6)
            .X = oLoc.X + lOffset
            .Y = oLoc.Y + rcDimensions.Height - 3 - lOffset
        End With
        With v2Lines(7)
            .X = oLoc.X + lOffset
            .Y = oLoc.Y + 3 + lOffset
        End With
        With v2Lines(8)
            .X = oLoc.X + 3 + lOffset
            .Y = oLoc.Y + lOffset
        End With

        'Draw a box border around our window...
        'ValidateBorderLine()
        'With moBorderLine
        '    .Antialias = True
        '    .Width = lBorderLineWidth
        '    .Begin()
        '    .Draw(v2Lines, clrColor)
        '    .End()
        'End With
        BPLine.DrawLine(lBorderLineWidth, True, v2Lines, clrColor)
    End Sub

    Protected Shared Sub RenderBox(ByVal rcDimensions As Rectangle, ByVal lBorderLineWidth As Int32, ByVal clrColor As System.Drawing.Color)
        Dim v2Lines(4) As Vector2
        Dim oLoc As Point = rcDimensions.Location
        Dim lOffset As Int32 = lBorderLineWidth \ 2

        With v2Lines(0)
            .X = oLoc.X + lOffset
            .Y = oLoc.Y + lOffset
        End With
        With v2Lines(1)
            .X = oLoc.X + rcDimensions.Width - lOffset
            .Y = oLoc.Y + lOffset
        End With
        With v2Lines(2)
            .X = oLoc.X + rcDimensions.Width - lOffset
            .Y = oLoc.Y + rcDimensions.Height - lOffset
        End With
        With v2Lines(3)
            .X = oLoc.X + lOffset
            .Y = oLoc.Y + rcDimensions.Height - lOffset
        End With
        With v2Lines(4)
            .X = oLoc.X + lOffset
            .Y = oLoc.Y + lOffset
        End With

        'Draw a box border around our window...
        'ValidateBorderLine()
        'With moBorderLine
        '    .Antialias = True
        '    .Width = lBorderLineWidth
        '    .Begin()
        '    .Draw(v2Lines, clrColor)
        '    .End()
        'End With
        BPLine.DrawLine(lBorderLineWidth, True, v2Lines, clrColor)
    End Sub

	Public Sub HandleResize(ByVal lNewX As Int32, ByVal lNewY As Int32)
        Dim lNewWidth As Int32 = Int32.MinValue
        If (mlResizeType And ResizeType.Left) <> 0 Then
            lNewWidth = Me.Width - (lNewX - mlResizeMouseX)
        ElseIf (mlResizeType And ResizeType.Right) <> 0 Then
            lNewWidth = Me.Width + (lNewX - mlResizeMouseX)
        End If
        Dim lNewHeight As Int32 = Int32.MinValue
        If (mlResizeType And ResizeType.Up) <> 0 Then
            lNewHeight = Me.Height - (lNewY - mlResizeMouseY)
        ElseIf (mlResizeType And ResizeType.Down) <> 0 Then
            lNewHeight = Me.Height + (lNewY - mlResizeMouseY)
        End If

        'Special booleans used for mlResizeType of both width and hight at the same time.  If ResizeLimits hit in one direction, dont block the other.
        Dim bNoWidth As Boolean = False
        Dim bNoHeight As Boolean = False
        If ResizeLimits.X > -1 AndAlso lNewWidth <> Int32.MinValue AndAlso lNewWidth < ResizeLimits.X Then bNoWidth = True
        If ResizeLimits.Y > -1 AndAlso lNewHeight <> Int32.MinValue AndAlso lNewHeight < ResizeLimits.Y Then bNoHeight = True
        If ResizeLimits.Width > -1 AndAlso lNewWidth <> Int32.MinValue AndAlso lNewWidth > ResizeLimits.Width Then bNoWidth = True
        If ResizeLimits.Height > -1 AndAlso lNewHeight <> Int32.MinValue AndAlso lNewHeight > ResizeLimits.Height Then bNoHeight = True

        If bNoWidth = False Then
            If (mlResizeType And ResizeType.Left) <> 0 Then
                Dim lOldX As Int32 = Me.Width
                Me.Width -= lNewX - mlResizeMouseX
                If lOldX <> Me.Width Then
                    Me.Left += lNewX - mlResizeMouseX
                End If
            ElseIf (mlResizeType And ResizeType.Right) <> 0 Then
                Me.Width += lNewX - mlResizeMouseX
            End If
            mlResizeMouseX = lNewX
        End If
        If bNoHeight = False Then
            If (mlResizeType And ResizeType.Up) <> 0 Then
                Dim lOldY As Int32 = Me.Height
                Me.Height -= lNewY - mlResizeMouseY
                If lOldY <> Me.Height Then
                    Me.Top += lNewY - mlResizeMouseY
                End If
            ElseIf (mlResizeType And ResizeType.Down) <> 0 Then
                Me.Height += lNewY - mlResizeMouseY
            End If
            mlResizeMouseY = lNewY
        End If
    End Sub

    Public Sub RemovedFromUILibList()
        RaiseEvent WindowClosed()
    End Sub

    Private Sub UIWindow_ResetInterfaceColors(ByVal yType As Byte, ByVal clrPrev As System.Drawing.Color) Handles Me.ResetInterfaceColors
        Select Case yType
            Case 1      'border
                If BorderColor = clrPrev Then BorderColor = muSettings.InterfaceBorderColor
            Case 2      'fill
                If FillColor = clrPrev Then FillColor = muSettings.InterfaceFillColor
            Case 3      'textboxfore
                'unused
            Case 4      'textboxfill
                'unused
            Case 5      'button
                'unused
        End Select
    End Sub

    Public Sub DoWindowInitialPosition(ByVal iLeft As Int32, ByVal iTop As Int32, ByVal iWidth As Int32, ByVal iHeight As Int32, ByVal iSettingX As Int32, ByVal iSettingY As Int32, ByVal iSettingWidth As Int32, ByVal iSettingHeight As Int32, ByVal bNoTutorial As Boolean)
        With Me
            Dim lLeft As Int32 = iLeft
            Dim lTop As Int32 = iTop
            Dim lWidth As Int32 = iWidth
            Dim lHeight As Int32 = iHeight

            If bNoTutorial = False OrElse NewTutorialManager.TutorialOn = False Then
                If iSettingX <> -1 Then lLeft = iSettingX
                If iSettingY <> -1 Then lTop = iSettingY
                If iSettingWidth <> -1 Then iWidth = iSettingWidth
                If iSettingHeight <> -1 Then lHeight = iSettingHeight
            End If

            'Left - Top checking
            If lLeft < 0 Then lLeft = Me.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - iWidth
            If lTop < 0 Then lTop = Me.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - iHeight
            If lLeft + .Width > Me.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = Me.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If lTop + .Height > Me.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = Me.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Left = lLeft
            .Top = lTop

            'Check the sanity of the height - width settings.  Drop them down if necessary
            If iSettingHeight > 0 AndAlso .Top + lHeight > Me.moUILib.oDevice.PresentationParameters.BackBufferHeight Then iHeight = Me.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Top
            If iSettingWidth > 0 AndAlso .Left + lWidth > Me.moUILib.oDevice.PresentationParameters.BackBufferWidth Then iWidth = Me.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width

            .Height = iHeight
            .Width = iWidth

        End With
    End Sub

End Class
