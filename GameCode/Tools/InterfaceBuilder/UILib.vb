'this class is responsible for all user interfaces...
Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Enum UILibMsgCode As Short
    eMouseDownMsgCode = 0
    eMouseMoveMsgCode
    eMouseUpMsgCode
    eKeyDownCode
    eKeyUpCode
    eKeyPressCode
End Enum

Public Class UILib
    Public oDevice As Device    'the D3D device object
    'Public oResMgr As EpicaResourceManager      'the resource manager

    Public oRenderTarget As Surface

    Public WindowList As Collection = New Collection()

    Public Pen As Sprite

    Public FocusedControl As UIControl

    Public oInterfaceTexture As Texture

    Public oFrameFont As System.Drawing.Font = New System.Drawing.Font("MS Sans Serif", 10.0F, FontStyle.Bold, GraphicsUnit.Point)

    Public Sub New(ByRef oDev As Device) ', ByRef oRes As EpicaResourceManager)
        oDevice = oDev
        CreateGlobalRectangleList()
        'oResMgr = oRes
        Pen = New Sprite(oDevice)

        oRenderTarget = oDevice.GetRenderTarget(0)

        'TODO: Replace this with intelligent texture loading
        'oInterfaceTexture = oRes.GetTexture("Interface_Main_Texture.bmp")
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If Right(sFile, 1) <> "\" Then sFile &= "\"
        sFile &= "Textures\Interface_Main_Texture.dds"
        oInterfaceTexture = TextureLoader.FromFile(oDevice, sFile, 0, 0, 0, Usage.None, Format.Unknown, Pool.Default, Filter.None, Filter.None, System.Drawing.Color.FromArgb(255, 255, 0, 255).ToArgb)
    End Sub

    Protected Overrides Sub Finalize()
        oDevice = Nothing
        'oResMgr = Nothing

        oRenderTarget.Dispose()
        oRenderTarget = Nothing
        WindowList = Nothing

        If Pen Is Nothing = False Then
            Pen.Dispose()
        End If
        Pen = Nothing
        FocusedControl = Nothing

        If oInterfaceTexture Is Nothing = False Then
            oInterfaceTexture.Dispose()
        End If
        oInterfaceTexture = Nothing
        MyBase.Finalize()

        oFrameFont.Dispose()
        oFrameFont = Nothing
    End Sub

    Public Sub DoAlphaBlendColorFill(ByVal rcDest As Rectangle, ByVal clrFill As System.Drawing.Color, ByVal ptLoc As Point)
        Dim rcSrc As Rectangle

        Dim fX As Single
        Dim fY As Single

        'rcSrc = System.Drawing.Rectangle.FromLTRB(225, 0, 255, 30)
        rcSrc.Location = New Point(192, 0)
        rcSrc.Width = 64
        rcSrc.Height = 64

        'Now, draw it...
        With Pen
            .Begin(SpriteFlags.AlphaBlend)

            fX = ptLoc.X * (rcSrc.Width / CSng(rcDest.Width))
            fY = ptLoc.Y * (rcSrc.Height / CSng(rcDest.Height))

            .Draw2D(oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), clrFill)
            .End()
        End With
    End Sub
End Class

'everything is a control... even a window
Public Class UIControl
    'NOTE: Location coordinates are parent-specific... meaning if this control is inside another control...
    '  then these cordinates represent the location INSIDE that control... NOT absolute
    Protected moLoc As Point = New Point(0, 0)
    Protected mlWidth As Int32
    Protected mlHeight As Int32
    Protected moControlRect As Rectangle
    Public Visible As Boolean = True
    Protected mbEnabled As Boolean = True
    'Private moChildren() As UIControl
    Public moChildren() As UIControl
    Public ChildrenUB As Int32 = -1
    Private mbHasFocus As Boolean

    'Public ControlImage_Disabled As Texture
    'Public ControlImage_Normal As Texture
    'Public ControlImage As Texture      'the current control image
    Public ControlImageRect As Rectangle
    Public ControlImageRect_Normal As Rectangle
	Public ControlImageRect_Disabled As Rectangle

	Public ToolTipText As String

    Protected moUILib As UILib

    Public ControlName As String

    Public mbAcceptReprocessEvents As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        moUILib = oUILib
    End Sub

    Private moParentControl As UIControl

    Public Event OnMouseDown(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons)
    Public Event OnMouseMove(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons)
    Public Event OnMouseUp(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons)
    Public Event OnKeyDown(ByVal e As KeyEventArgs)
    Public Event OnKeyUp(ByVal e As KeyEventArgs)
    Public Event OnKeyPress(ByVal e As KeyPressEventArgs)
    Public Event OnRender()
    Public Event OnGotFocus()
    Public Event OnLostFocus()
    Public Event OnResize()
    Public Event ReprocessInput()       'for sustained inputs
 
    Private Sub UpdateControlRect()
        moControlRect.Location = GetAbsolutePosition()
        moControlRect.Width = Width
        moControlRect.Height = Height
    End Sub

    Public Property ParentControl() As UIControl
        Get
            Return moParentControl
        End Get
        Set(ByVal Value As UIControl)
            moParentControl = Value
            UpdateControlRect()
            RaiseEvent OnResize()
        End Set
    End Property

    Public Property Left() As Int32
        Get
            Return moLoc.X
        End Get
        Set(ByVal Value As Int32)
            moLoc.X = Value
            UpdateControlRect()
            RaiseEvent OnResize()
        End Set
    End Property

    Public Property Top() As Int32
        Get
            Return moLoc.Y
        End Get
        Set(ByVal Value As Int32)
            moLoc.Y = Value
            UpdateControlRect()
            RaiseEvent OnResize()
        End Set
    End Property

    Public Property Width() As Int32
        Get
            Return mlWidth
        End Get
        Set(ByVal Value As Int32)
            mlWidth = Value
            UpdateControlRect()
            RaiseEvent OnResize()
        End Set
    End Property

    Public Property Height() As Int32
        Get
            Return mlHeight
        End Get
        Set(ByVal Value As Int32)
            mlHeight = Value
            UpdateControlRect()
            RaiseEvent OnResize()
        End Set
    End Property

    Public Property HasFocus() As Boolean
        Get
            Return mbHasFocus
        End Get
        Set(ByVal Value As Boolean)
            If mbHasFocus = True AndAlso Value = False Then
                mbHasFocus = Value
                RaiseEvent OnLostFocus()
            Else
                mbHasFocus = Value
            End If
        End Set
    End Property

    Public Property Enabled() As Boolean
        Get
            Return mbEnabled
        End Get
        Set(ByVal Value As Boolean)
            Dim X As Int32
            mbEnabled = Value
            For X = 0 To ChildrenUB
                moChildren(X).Enabled = Value
            Next X
        End Set
    End Property

    Public Sub DrawControl()
        Dim X As Int32
        If Visible Then
            RaiseEvent OnRender()
            For X = 0 To ChildrenUB
                moChildren(X).DrawControl()
            Next X
        End If
    End Sub

    Public Function PostMessage(ByVal lMsgCode As UILibMsgCode, ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons) As Boolean
        Dim oTemp As UIControl

        If TestRegion(lMouseX, lMouseY) Then
            oTemp = PostToChildren(lMsgCode, lMouseX, lMouseY, lButton)
            If oTemp Is Nothing Then
                Select Case lMsgCode
                    Case UILibMsgCode.eMouseDownMsgCode
                        RaiseEvent OnMouseDown(lMouseX, lMouseY, lButton)
                    Case UILibMsgCode.eMouseMoveMsgCode
                        RaiseEvent OnMouseMove(lMouseX, lMouseY, lButton)
                    Case UILibMsgCode.eMouseUpMsgCode
                        RaiseEvent OnMouseUp(lMouseX, lMouseY, lButton)
                End Select
            End If
            Return True
        Else
            Return False
        End If
    End Function

    Public Function PostMessage(ByVal lMsgCode As UILibMsgCode, ByVal eMsg As System.Windows.Forms.KeyEventArgs) As Boolean
        Dim X As Int32

        If HasFocus Then
            Select Case lMsgCode
                Case UILibMsgCode.eKeyDownCode
                    RaiseEvent OnKeyDown(eMsg)
                Case UILibMsgCode.eKeyUpCode
                    RaiseEvent OnKeyUp(eMsg)
            End Select
            Return True
        Else
            'check children
            For X = 0 To ChildrenUB
                If Not moChildren(X) Is Nothing Then
                    If moChildren(X).HasFocus Then
                        moChildren(X).PostMessage(lMsgCode, eMsg)
                        Return True
                    End If
                End If
            Next X
        End If
    End Function

    Public Function PostMessage(ByVal lMsgCode As UILibMsgCode, ByVal eMsg As System.Windows.Forms.KeyPressEventArgs) As Boolean
        Dim X As Int32

        If HasFocus Then
            RaiseEvent OnKeyPress(eMsg)
            Return True
        Else
            'check children
            For X = 0 To ChildrenUB
                If Not moChildren(X) Is Nothing Then
                    If moChildren(X).HasFocus Then
                        moChildren(X).PostMessage(lMsgCode, eMsg)
                        Return True
                    End If
                End If
            Next X
        End If
    End Function

    Public Function TestRegion(ByVal lMouseX As Int32, ByVal lMouseY As Int32) As Boolean
        Dim oLoc As System.Drawing.Point
        Dim lX As Int32
        Dim lY As Int32

        oLoc = GetAbsolutePosition()
        lX = oLoc.X
        lY = oLoc.Y
        oLoc = Nothing

        If Visible And Enabled Then
            If lMouseX >= lX AndAlso lMouseX <= (lX + mlWidth) Then
                If lMouseY >= lY AndAlso lMouseY <= (lY + mlHeight) Then
                    Return True
                End If
            End If
        End If
        Return False
    End Function

    Private Function PostToChildren(ByVal lMsgCode As UILibMsgCode, ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons) As UIControl
        Dim X As Int32

        For X = 0 To ChildrenUB
            If Not moChildren(X) Is Nothing Then
                If moChildren(X).PostMessage(lMsgCode, lMouseX, lMouseY, lButton) Then
                    Return moChildren(X)
                End If
            End If
        Next X
        Return Nothing
    End Function

    Public Sub AddChild(ByRef oChild As UIControl)
        ChildrenUB += 1
        ReDim Preserve moChildren(ChildrenUB)
        oChild.ParentControl = Me
        moChildren(ChildrenUB) = oChild
    End Sub

    Public Sub RemoveChild(ByVal lIndex As Int32)
        Dim X As Int32

        If moChildren.Length >= lIndex Then
            moChildren(lIndex) = Nothing

            'Now, shift all remaining items down
            For X = lIndex To ChildrenUB - 1
                moChildren(X) = moChildren(X + 1)
            Next X
            ChildrenUB -= 1
        End If
    End Sub

    Public Sub RemoveAllChildren()
        Dim X As Int32
        For X = 0 To ChildrenUB
            moChildren(X) = Nothing
        Next X
        ChildrenUB = -1
        Erase moChildren
    End Sub

    Public Function GetAbsolutePosition() As Point
        Dim lX As Int32
        Dim lY As Int32

        lX = moLoc.X
        lY = moLoc.Y

        If ParentControl Is Nothing = False Then
            Dim oTmp As Point
            oTmp = ParentControl.GetAbsolutePosition()
            lX += oTmp.X
            lY += oTmp.Y
            oTmp = Nothing
        End If

        Return New Point(lX, lY)
    End Function

    Public Sub ReprocessInputs()
        Dim X As Int32
        If Visible AndAlso mbAcceptReprocessEvents Then
            RaiseEvent ReprocessInput()
            For X = 0 To ChildrenUB
                moChildren(X).ReprocessInputs()
            Next X
        End If
    End Sub

    Protected Overrides Sub Finalize()
        RemoveAllChildren()
        'ControlImage_Normal = Nothing
        'ControlImage_Disabled = Nothing
        'ControlImage = Nothing
        moUILib = Nothing
        MyBase.Finalize()
    End Sub
End Class

Public Class UIWindow
    Inherits UIControl

    'windows can move
    Private mlMouseDownX As Int32
    Private mlMouseDownY As Int32

    Private moBorderLine As Line
    Private moBorderLineVerts(8) As Vector2

    Public FillColor As System.Drawing.Color = System.Drawing.Color.FromArgb(128, 32, 64, 128)
    Public BorderColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 255)

    Public FullScreen As Boolean = False

    Public Caption As String = ""

    Private moFont As Font = Nothing

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)
        moBorderLine = New Line(oUILib.oDevice)
    End Sub

    Private Sub UIWindow_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseDown
        'get our mouse X and mouse Y on this control
        Dim oLoc As Point = GetAbsolutePosition()
        mlMouseDownX = lMouseX - oLoc.X
        mlMouseDownY = lMouseY - oLoc.Y
        oLoc = Nothing
    End Sub

    Private Sub UIWindow_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseMove
        If lButton = MouseButtons.Left Then
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

        End If
    End Sub

    Private Sub UIWindow_OnRender() Handles MyBase.OnRender
        Dim oLoc As Point = GetAbsolutePosition()

        Dim oRect As Rectangle

        With oRect
            .Location() = oLoc
            .Width = Width
            .Height = Height
        End With

        moUILib.DoAlphaBlendColorFill(oRect, FillColor, oLoc)

        'Draw a box border around our window...
        If Me.Caption.Length = 0 Then
            With moBorderLineVerts(0)
                .X = oLoc.X
                .Y = oLoc.Y
            End With
        Else
            If moFont Is Nothing Then moFont = New Font(MyBase.moUILib.oDevice, MyBase.moUILib.oFrameFont)
            Dim rcTmp As Rectangle = moFont.MeasureString(Nothing, Caption, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, Color.White)
            With moBorderLineVerts(0)
                .X = oLoc.X + rcTmp.Width + 10
                .Y = oLoc.Y
            End With
            If Caption <> "" Then moFont.DrawText(Nothing, Caption, New Rectangle(oLoc.X + 5, oLoc.Y - 7, rcTmp.Width, rcTmp.Height), DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, Color.White)
        End If
        moBorderLineVerts(0).X += 3
        With moBorderLineVerts(1)
            .X = oLoc.X + Width - 3
            .Y = oLoc.Y
        End With
        With moBorderLineVerts(2)
            .X = oLoc.X + Width
            .Y = oLoc.Y + 3
        End With
        With moBorderLineVerts(3)
            .X = oLoc.X + Width
            .Y = oLoc.Y + Height - 3
        End With
        With moBorderLineVerts(4)
            .X = oLoc.X + Width - 3
            .Y = oLoc.Y + Height
        End With
        With moBorderLineVerts(5)
            .X = oLoc.X + 3
            .Y = oLoc.Y + Height
        End With
        With moBorderLineVerts(6)
            .X = oLoc.X
            .Y = oLoc.Y + Height - 3
        End With
        With moBorderLineVerts(7)
            .X = oLoc.X
            .Y = oLoc.Y + 3
        End With
        With moBorderLineVerts(8)
            .X = oLoc.X + 3
            .Y = oloc.Y
        End With

        'Draw a box border around our window...
        With Me.moUILib.oDevice
            moBorderLine.Antialias = True
            moBorderLine.Width = 3
            moBorderLine.Begin()
            moBorderLine.Draw(moBorderLineVerts, BorderColor)
            moBorderLine.End()
        End With
        'End of the border drawing

        oLoc = Nothing
    End Sub

    Private Sub UIWindow_OnResize() Handles MyBase.OnResize
        Dim oLoc As Point = GetAbsolutePosition()
        With moBorderLineVerts(0)
            .X = oLoc.X
            .Y = oLoc.Y
        End With
        With moBorderLineVerts(1)
            .X = oLoc.X + Width
            .Y = oLoc.Y
        End With
        With moBorderLineVerts(2)
            .X = oLoc.X + Width
            .Y = oLoc.Y + Height
        End With
        With moBorderLineVerts(3)
            .X = oLoc.X
            .Y = oLoc.Y + Height
        End With
        With moBorderLineVerts(4)
            .X = oLoc.X
            .Y = oLoc.Y
        End With
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub
End Class

Public Class UILine
    Inherits UIControl

    Private Shared moLinePen As Line
    Private Shared muVecs(1) As Vector2

    Public BorderColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 255)

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

    End Sub

    Public Sub Control_OnRender() Handles MyBase.OnRender
        Dim oLoc As System.Drawing.Point

        oLoc = GetAbsolutePosition()

        muVecs(0).X = oLoc.X
        muVecs(0).Y = oLoc.Y

        muVecs(1).X = oLoc.X + Width
        muVecs(1).Y = oLoc.Y + Height

        If moLinePen Is Nothing Then moLinePen = New Line(mouilib.oDevice)

        moLinePen.Antialias = True
        moLinePen.Begin()
        moLinePen.Draw(muVecs, BorderColor)
        moLinePen.End()
    End Sub
End Class

Public Class UILabel
    Inherits UIControl

    Public Caption As String
    Protected moFont As Font
    Private moSysFont As System.Drawing.Font
    Public ForeColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 255)
    Public FontFormat As DrawTextFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter

    Public DrawBackImage As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'Default Font
        moSysFont = New System.Drawing.Font("MS Sans Serif", 10)
        moFont = New Font(oUILib.oDevice, moSysFont)

    End Sub

    Public Sub SetFont(ByVal oFont As System.Drawing.Font)
        moFont = Nothing
        moSysFont = oFont
        moFont = New Font(moUILib.oDevice, oFont)
    End Sub

    Public Function GetFont() As System.Drawing.Font
        Return moSysFont
    End Function

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
            If DrawBackImage Then
                With moUILib.Pen
                    .Begin(SpriteFlags.AlphaBlend)
                    fMultX = CSng(ControlImageRect.Width / Width)
                    fMultY = CSng(ControlImageRect.Height / Height)
                    .Draw2D(moUILib.oInterfaceTexture, ControlImageRect, System.Drawing.Rectangle.FromLTRB(oLoc.X, oLoc.Y, oLoc.X + Width, oLoc.Y + Height), System.Drawing.Point.Empty, 0, New Point(CInt(oLoc.X * fMultX), CInt(oLoc.Y * fMultY)), System.Drawing.Color.FromArgb(&HFFFFFFFF))
                    .End()
                End With
            End If
            If Caption <> "" Then moFont.DrawText(Nothing, Caption, oRect, FontFormat, ForeColor)
        End With
    End Sub

End Class

Public Class UIButton
    Inherits UILabel

    'Public ControlImage_Pressed As Texture
    Public ControlImageRect_Pressed As Rectangle
    Protected mbPressed As Boolean = False

    Private moAbsPos As Point

    Public Event Click()

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'now, always load the default images
        ControlImageRect_Pressed = System.Drawing.Rectangle.FromLTRB(0, 64, 120, 96)
        'ControlImage_Pressed = oUILib.oResMgr.GetTexture("Button_Pressed.bmp")
        ControlImageRect_Normal = System.Drawing.Rectangle.FromLTRB(0, 0, 120, 32)
        'ControlImage_Normal = oUILib.oResMgr.GetTexture("Button_Normal.bmp")
        ControlImageRect_Disabled = System.Drawing.Rectangle.FromLTRB(0, 32, 128, 64)
        'ControlImage_Disabled = oUILib.oResMgr.GetTexture("Button_Disabled.bmp")

        Me.DrawBackImage = True

        Me.ForeColor = System.Drawing.Color.White
    End Sub

    Public Overrides Sub Control_OnRender() Handles MyBase.OnRender
        'render the button
        If Me.Enabled = True Then
            If mbPressed Then
                'ControlImage = ControlImage_Pressed
                ControlImageRect = ControlImageRect_Pressed
            Else
                'ControlImage = ControlImage_Normal
                ControlImageRect = ControlImageRect_Normal
            End If
        Else
            'ControlImage = ControlImage_Disabled
            ControlImageRect = ControlImageRect_Disabled
        End If

        'and then render the label
        MyBase.Control_OnRender()
    End Sub

    Private Sub UIButton_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseDown
        mbPressed = True
    End Sub

    Private Sub UIButton_OnMouseUp(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseUp
        If mbPressed Then RaiseEvent Click()
        mbPressed = False
    End Sub

    Protected Overrides Sub Finalize()
        Me.Visible = False
        'ControlImage_Pressed = Nothing
        MyBase.Finalize()
    End Sub

    Private Sub UIButton_ReprocessInput() Handles MyBase.ReprocessInput
        If mbPressed Then RaiseEvent Click()
    End Sub
End Class

Public Class UITextBox
    Inherits UILabel

    Public SelStart As Int32 = 0      'current position in the string that the cursor is located
    Public SelLength As Int32       'length of the selection

    Public BackColorEnabled As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 32, 64, 92)
    Public BackColorDisabled As System.Drawing.Color = System.Drawing.Color.DimGray

    Public MaxLength As Int32

    Private moBorderLine As Line
    Private moBorderLineVerts(4) As Vector2
    Private moCursorLine As Line
    Private moCursorVerts(1) As Vector2

    Public BorderColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 255)

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'Me.ControlImage_Normal = oUILib.oResMgr.GetTexture("Textbox_Normal.bmp")
        'Me.ControlImage_Disabled = oUILib.oResMgr.GetTexture("Textbox_Disabled.bmp")
        Me.ForeColor = System.Drawing.Color.White
        Me.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        moBorderLine = New Line(oUILib.oDevice)
        moCursorLine = New Line(oUILib.oDevice)
    End Sub

    Public Function GetSelectedString() As String
        Return Mid$(Caption, SelStart, SelLength)
    End Function

    Private Sub UITextBox_OnRender() Handles MyBase.OnRender
        Dim oRect As Rectangle
        Dim oLoc As System.Drawing.Point
        Dim lFillColor As System.Drawing.Color

        If Enabled Then
            lFillColor = BackColorEnabled
        Else
            lFillColor = BackColorDisabled
        End If

        oLoc = GetAbsolutePosition()
        With oRect
            .Location() = oLoc
            .Width = Width
            .Height = Height

            'Do a color fill
            On Error Resume Next
            moUILib.oDevice.ColorFill(moUILib.oRenderTarget, oRect, lFillColor)
            'end of color fill

            'Draw a box border around our window...
            With moBorderLineVerts(0)
                .X = oLoc.X
                .Y = oLoc.Y
            End With
            With moBorderLineVerts(1)
                .X = oLoc.X + Width
                .Y = oLoc.Y
            End With
            With moBorderLineVerts(2)
                .X = oLoc.X + Width
                .Y = oLoc.Y + Height
            End With
            With moBorderLineVerts(3)
                .X = oLoc.X
                .Y = oLoc.Y + Height
            End With
            With moBorderLineVerts(4)
                .X = oLoc.X
                .Y = oLoc.Y
            End With

            moBorderLine.Begin()
            moBorderLine.Draw(moBorderLineVerts, BorderColor)
            moBorderLine.End()
            'End of the border drawing

            'Draw the caption, but we need to scoot the rect over a bit
            oRect.X += 5
            oRect.Width -= 5
            moFont.DrawText(Nothing, Caption, oRect, FontFormat, ForeColor)

            If HasFocus Then
                oRect = moFont.MeasureString(Nothing, Mid$(Caption, 1, SelStart), FontFormat, ForeColor)

                'ok, now, draw a line there
                moCursorVerts(0).X = oRect.Left + oRect.Width + oLoc.X + 5
                moCursorVerts(0).Y = oRect.Top + oLoc.Y + 5
                moCursorVerts(1).X = moCursorVerts(0).X
                moCursorVerts(1).Y = oRect.Height + oLoc.Y + 5

                moCursorLine.Begin()
                moCursorLine.Draw(moCursorVerts, System.Drawing.Color.Black)
                moCursorLine.End()
            End If

        End With

    End Sub

    Private Sub UITextBox_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseDown
        If moUILib.FocusedControl Is Nothing = False Then
            moUILib.FocusedControl.HasFocus = False
        End If
        HasFocus = True
        moUILib.FocusedControl = Me
    End Sub

    Private Sub UITextBox_OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.OnKeyPress
        Dim sTemp As String
        If Char.IsControl(e.KeyChar) = False Then
            If MaxLength <> 0 Then
                If Len(Caption) + 1 <= MaxLength Then
                    sTemp = Mid$(Caption, 1, SelStart) & e.KeyChar & Mid$(Caption, SelStart + 1)
                    SelStart += 1
                    Caption = sTemp
                End If
            Else
                sTemp = Mid$(Caption, 1, SelStart) & e.KeyChar & Mid$(Caption, SelStart + 1)
                SelStart += 1
                Caption = sTemp
            End If
        End If

    End Sub

    Private Sub UITextBox_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.OnKeyDown
        If e.KeyCode = Keys.Back Then
            If SelStart > 0 Then
                Caption = Mid$(Caption, 1, SelStart - 1) & Mid$(Caption, SelStart + 1)
                SelStart -= 1
            End If
        ElseIf e.KeyCode = Keys.Delete Then
            Caption = Mid$(Caption, 1, SelStart) & Mid$(Caption, SelStart + 2)
        ElseIf e.KeyCode = Keys.Right Then
            If SelStart < Len(Caption) Then SelStart += 1
        ElseIf e.KeyCode = Keys.Left Then
            If SelStart > 0 Then SelStart -= 1
        ElseIf e.KeyCode = Keys.Home Then
            SelStart = 0
        ElseIf e.KeyCode = Keys.End Then
            SelStart = Len(Caption)
        End If
    End Sub

End Class

Public Class UIOption
    'Inherits UIButton
    Inherits UILabel

    Public Value As Boolean
    'Public ControlImage_Pressed As Texture
    Public ControlImageRect_Pressed As Rectangle

    Protected mbPressed As Boolean = False

    Public Event Click()

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'load default images
        ControlImageRect_Disabled = System.Drawing.Rectangle.FromLTRB(72, 96, 95, 119)
        'Me.ControlImage_Disabled = oUILib.oResMgr.GetTexture("Option_Disabled.bmp")
        ControlImageRect_Normal = System.Drawing.Rectangle.FromLTRB(96, 96, 119, 119)
        'Me.ControlImage_Normal = oUILib.oResMgr.GetTexture("Option_Normal.bmp")
        ControlImageRect_Pressed = System.Drawing.Rectangle.FromLTRB(120, 96, 143, 119)
        'Me.ControlImage_Pressed = oUILib.oResMgr.GetTexture("Option_Pressed.bmp")

        'Me.DrawBackImage = True

        ControlImageRect_Disabled = System.Drawing.Rectangle.FromLTRB(10, 225, 20, 235)
        ControlImageRect_Normal = System.Drawing.Rectangle.FromLTRB(0, 225, 10, 235)
        ControlImageRect_Pressed = System.Drawing.Rectangle.FromLTRB(20, 225, 30, 235)

        Me.ForeColor = System.Drawing.Color.White
    End Sub

    Private Sub UIOption_OnRender() Handles MyBase.OnRender
        'render the button
        Dim oLoc As Point = GetAbsolutePosition()
        Dim oRect As Rectangle

        If Me.Enabled = True Then
            If Value Then
                ControlImageRect = ControlImageRect_Pressed
            Else
                ControlImageRect = ControlImageRect_Normal
            End If
        Else
            ControlImageRect = ControlImageRect_Disabled
        End If

        'oLoc = GetAbsolutePosition()
        With oRect
            .Location() = oLoc
            .Width = Width
            .Height = Height

            Dim lOptOffsetY As Int32 = (Height - ControlImageRect.Height) \ 2I

            'render the back image if there is one
            With moUILib.Pen
                .Begin(SpriteFlags.AlphaBlend)
                .Draw2D(moUILib.oInterfaceTexture, ControlImageRect, System.Drawing.Rectangle.FromLTRB(oLoc.X, oLoc.Y + lOptOffsetY, oLoc.X + ControlImageRect.Width, oLoc.Y + lOptOffsetY + ControlImageRect.Height), System.Drawing.Point.Empty, 0, New Point(oLoc.X, oLoc.Y + lOptOffsetY), System.Drawing.Color.FromArgb(&HFFFFFFFF))
                .End()
            End With
            If Caption <> "" Then moFont.DrawText(Nothing, Caption, oRect, FontFormat, ForeColor)
        End With
    End Sub

    Private Sub UIButton_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseDown
        mbPressed = True
    End Sub

    Private Sub UIButton_OnMouseUp(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseUp
        If mbPressed Then
            Value = Not Value
            RaiseEvent Click()
        End If
        mbPressed = False
    End Sub

    Protected Overrides Sub Finalize()
        Me.Visible = False
        'ControlImage_Pressed = Nothing
        MyBase.Finalize()
    End Sub
End Class

Public Class UICheckBox
    Inherits UIOption

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'load default images
        ControlImageRect_Disabled = System.Drawing.Rectangle.FromLTRB(120, 0, 143, 25)
        'Me.ControlImage_Disabled = oUILib.oResMgr.GetTexture("Checkbox_Disabled.bmp")
        ControlImageRect_Normal = System.Drawing.Rectangle.FromLTRB(144, 0, 167, 25)
        'Me.ControlImage_Normal = oUILib.oResMgr.GetTexture("Checkbox_Normal.bmp")
        ControlImageRect_Pressed = System.Drawing.Rectangle.FromLTRB(168, 0, 191, 25)
        'Me.ControlImage_Pressed = oUILib.oResMgr.GetTexture("Checkbox_Pressed.bmp")

        ControlImageRect_Disabled = System.Drawing.Rectangle.FromLTRB(10, 215, 20, 225)
        ControlImageRect_Normal = System.Drawing.Rectangle.FromLTRB(0, 215, 10, 225)
        ControlImageRect_Pressed = System.Drawing.Rectangle.FromLTRB(30, 215, 40, 225)

    End Sub

End Class

Public Class UIScrollBar
    Inherits UIControl

    Private Const ml_BTN_OTHER_SIZE As Int32 = 24

    Protected mbVertical As Boolean
    Public Value As Int32 = 0
    Public MaxValue As Int32 = 1
    Public MinValue As Int32 = 0
    Public SmallChange As Int32 = 1
    Public LargeChange As Int32 = 1

    Public ReverseDirection As Boolean = False

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

        btnDecreaser = New UIButton(oUILib)
        btnIncreaser = New UIButton(oUILib)
        btnScroller = New UIButton(oUILib)

        btnDecreaser.mbAcceptReprocessEvents = True
        btnIncreaser.mbAcceptReprocessEvents = True
        btnScroller.mbAcceptReprocessEvents = True

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

    End Sub


    Private Sub btnDecreaser_Click() Handles btnDecreaser.Click
        If ReverseDirection = False Then
            Value -= SmallChange
            If Value < MinValue Then Value = MinValue
        Else
            Value += SmallChange
            If Value > MaxValue Then Value = MaxValue
        End If
        RaiseEvent ValueChange()
    End Sub

    Private Sub btnIncreaser_Click() Handles btnIncreaser.Click
        If ReverseDirection = False Then
            Value += SmallChange
            If Value > MaxValue Then Value = MaxValue
        Else
            Value -= SmallChange
            If Value < MinValue Then Value = MinValue
        End If
        RaiseEvent ValueChange()
    End Sub

    Private Sub btnScroller_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles btnScroller.OnMouseDown
        'get our mouse X and mouse Y on this control
        Dim oLoc As Point = GetAbsolutePosition()
        mlScrollingMouseX = lMouseX - oLoc.X
        mlScrollingMouseY = lMouseY - oLoc.Y
        mbScrolling = True
        oLoc = Nothing
    End Sub

    Private Sub btnScroller_OnMouseUp(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles btnScroller.OnMouseUp
        mbScrolling = False
    End Sub

    Private Sub btnScroller_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles btnScroller.OnMouseMove
        If mbScrolling Then


            '    'ok, user is dragging window..
            '    Dim oLoc As Point = GetAbsolutePosition()
            '    Dim lX As Int32
            '    Dim lY As Int32
            '    Dim lTemp As Int32

            '    'basically, this is reverse placement... wherever the mouse is located, we want to find the value
            '    If mbVertical Then
            '        'only care about Y
            '        'we want to know the LocY within the Scrollbar....
            '        lY = lMouseY - oLoc.Y '- mlScrollingMouseY
            '        'lY is where the top of the scroller would be
            '        lTemp = btnDecreaser.Top - (btnIncreaser.Top + btnIncreaser.Height) - btnScroller.Height
            '        'lTemp is number of units
            '        lTemp = (lY / lTemp) * MaxValue
            '        Value = lTemp
            '    Else
            '        'only care about x
            '        lX = lMouseX - oLoc.X '- mlScrollingMouseX
            '        'lX is where the front of the scroller would be
            '        lTemp = btnIncreaser.Left - (btnDecreaser.Left + btnDecreaser.Width) - btnScroller.Width
            '        'lTemp is number of units we *could* actually move
            '        lTemp = (lX / lTemp) * MaxValue
            '        Value = lTemp
            '    End If


            '    'btnScroller.Left = lMouseX - lX - mlScrollingMouseX
            '    'btnScroller.Top = lMouseY - lY - mlScrollingMouseY
        End If
    End Sub

    Private Sub UIScrollBar_OnRender() Handles MyBase.OnRender
        Dim oLoc As Point = GetAbsolutePosition()
        Dim lTemp As Int32

        If mbVertical Then
            'increaser is at the top, decreaser is at the bottom
            With btnIncreaser
                .Top = 0
                .Left = 0
                .Width = Me.Width
                .Height = ml_BTN_OTHER_SIZE
            End With
            With btnDecreaser
                .Top = Me.Height - ml_BTN_OTHER_SIZE
                .Left = 0
                .Width = Me.Width
                .Height = ml_BTN_OTHER_SIZE
            End With

            'the scroller moves up and down
            With btnScroller
                .Width = Me.Width
                .Height = ml_BTN_OTHER_SIZE
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
                    .Top = (btnIncreaser.Top + btnIncreaser.Width) + lTemp
                End If

            End With
        Else
            'increaser is on the right, decreaser is on the left
            With btnIncreaser
                .Top = 0
                .Left = Me.Width - ml_BTN_OTHER_SIZE
                .Width = ml_BTN_OTHER_SIZE
                .Height = Me.Height
            End With
            With btnDecreaser
                .Top = 0
                .Left = 0
                .Width = ml_BTN_OTHER_SIZE
                .Height = Me.Height
            End With

            'the scroller moves left and right
            With btnScroller
                .Width = ml_BTN_OTHER_SIZE
                .Height = Me.Height
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
        On Error Resume Next
        If Enabled Then
            moUILib.oDevice.ColorFill(moUILib.oRenderTarget, oRect, System.Drawing.Color.FromArgb(255, 0, 0, 0))
        Else
            moUILib.oDevice.ColorFill(moUILib.oRenderTarget, oRect, System.Drawing.Color.DarkGray)
        End If
        oRect = Nothing
        'end of color fill
    End Sub

    Protected Overrides Sub Finalize()
        RemoveAllChildren()
        btnDecreaser = Nothing
        btnIncreaser = Nothing
        btnScroller = Nothing
        MyBase.Finalize()
    End Sub
End Class

Public Class UIListBox
    Inherits UIControl

    Public Event ItemClick(ByVal lIndex As Int32)
    Public Event ItemDblClick(ByVal lIndex As Int32)

    Protected WithEvents moScroll As UIScrollBar

    Private moBorderLine As Line
    Private moBorderLineVerts(4) As Vector2

    Private moListItems() As UILabel
    Private mlListItemUB As Int32 = -1      'not the same as our real item list... this is for display purposes

    Public NewIndex As Int32 = -1

    Public BorderColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 255)
    Public FillColor As System.Drawing.Color = System.Drawing.Color.FromArgb(128, 32, 64, 128)
    Public HighlightColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 32, 64, 128)
    Private moForeColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 255)

    Private moSysFont As System.Drawing.Font

    Protected msItems() As String
    Protected mlItemUB As Int32 = -1
    Protected mlItemData() As Int32
    Private mlListIndex As Int32

    Public Sub AddItem(ByVal sItem As String)
        Dim X As Int32

        mlItemUB += 1
        ReDim Preserve msItems(mlItemUB)
        ReDim Preserve mlItemData(mlItemUB)

        msItems(mlItemUB) = sItem
        mlItemData(mlItemUB) = 0

        NewIndex = mlItemUB

        X = mlItemUB - mlListItemUB
        If X > -1 Then
            moScroll.MaxValue = X
            'moScroll.Visible = True
            moScroll.Enabled = True
        Else
            moScroll.MaxValue = 1
            moScroll.Value = 0
            'moScroll.Visible = False
            moScroll.Enabled = False
        End If
    End Sub

    Public Sub Clear()
        Dim X As Int32
        NewIndex = -1
        mlItemUB = -1
        ReDim msItems(-1)

        X = mlItemUB - mlListItemUB
        If X > -1 Then
            moScroll.MaxValue = X
            'moScroll.Visible = True
            moScroll.Enabled = True
        Else
            moScroll.MaxValue = 1
            moScroll.Value = 0
            'moScroll.Visible = False
            moScroll.Enabled = False
        End If
    End Sub

    Public Sub RemoveItem(ByVal lIndex As Int32)
        Dim X As Int32

        If lIndex <= mlItemUB Then
            For X = lIndex To mlItemUB - 1
                'shift from one up down
                msItems(X) = msItems(X + 1)
                mlItemData(X) = mlItemData(X + 1)
            Next X

            mlItemUB -= 1
            ReDim Preserve msItems(mlItemUB)
            ReDim Preserve mlItemData(mlItemUB)
        End If
    End Sub

    Public Property ItemData(ByVal lIndex As Int32) As Int32
        Get
            If lIndex <= mlItemUB Then Return mlItemData(lIndex)
        End Get
        Set(ByVal Value As Int32)
            If lIndex <= mlItemUB Then
                mlItemData(lIndex) = Value
            End If
        End Set
    End Property
    Public Property List(ByVal lIndex As Int32) As String
        Get
            If lIndex <= mlItemUB Then Return msItems(lIndex) Else Return ""
        End Get
        Set(ByVal Value As String)
            If lIndex <= mlItemUB Then
                msItems(lIndex) = Value
            End If
        End Set
    End Property
    Public ReadOnly Property ListCount() As Int32
        Get
            Return mlItemUB + 1
        End Get
    End Property
    Public Property ListIndex() As Int32
        Get
            Return mlListIndex
        End Get
        Set(ByVal Value As Int32)
            If Value <= mlItemUB Then
                mlListIndex = Value
                RaiseEvent ItemClick(mlListIndex)
            End If
        End Set
    End Property

    Public Property ForeColor() As System.Drawing.Color
        Get
            Return moForeColor
        End Get
        Set(ByVal Value As System.Drawing.Color)
            moForeColor = Value
            Dim X As Int32

            'For X = 0 To mlListItemUB
            '    moListItems(X).ForeColor = moForeColor
            'Next X
        End Set
    End Property

    Public Sub SetFont(ByVal oFont As System.Drawing.Font)
        Dim X As Int32

        moSysFont = oFont

        'For X = 0 To mlListItemUB
        '    moListItems(X).SetFont(moSysFont)
        'Next X

    End Sub

    Public Function GetFont() As System.Drawing.Font
        Return moSysFont
    End Function

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        moBorderLine = New Line(oUILib.oDevice)

        'Default font
        SetFont(New System.Drawing.Font("MS Sans Serif", 10))


        moScroll = New UIScrollBar(oUILib, True)
        moScroll.ReverseDirection = True
        Me.AddChild(CType(moScroll, UIControl))
    End Sub

    Private Sub Label_MouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons)
        Dim lTemp As Int32
        Dim oLoc As Point = GetAbsolutePosition()

        Dim lY As Int32

        lY = lMouseY - oLoc.Y

        'Now, lY is relative to within the listbox, find the Label they clicked on
        lTemp = CInt(Int(lY / moListItems(0).Height)) + moScroll.Value

        ListIndex = lTemp
    End Sub

    Private Sub Label_MouseUp(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons)
        '
    End Sub

    Private Sub UIListBox_OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.OnKeyUp
        '
    End Sub

    Private Sub UIListBox_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.OnKeyDown
        '
    End Sub

    Private Sub UIListBox_OnMouseUp(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseUp
        '
    End Sub

    Private Sub UIListBox_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseDown
        '
    End Sub

    Private Sub UIListBox_OnRender() Handles MyBase.OnRender
        Dim oLoc As System.Drawing.Point = GetAbsolutePosition()
        Dim X As Int32
        Dim lMax As Int32

        'Do a color fill of White
        On Error Resume Next
        Dim oRect As Rectangle = New Rectangle(oLoc.X, oLoc.Y, Width, Height)
        moUILib.oDevice.ColorFill(moUILib.oRenderTarget, oRect, FillColor)
        oRect = Nothing
        'end of color fill

        'Draw a box border around our window...
        With moBorderLineVerts(0)
            .X = oLoc.X
            .Y = oLoc.Y
        End With
        With moBorderLineVerts(1)
            .X = oLoc.X + Width
            .Y = oLoc.Y
        End With
        With moBorderLineVerts(2)
            .X = oLoc.X + Width
            .Y = oLoc.Y + Height
        End With
        With moBorderLineVerts(3)
            .X = oLoc.X
            .Y = oLoc.Y + Height
        End With
        With moBorderLineVerts(4)
            .X = oLoc.X
            .Y = oLoc.Y
        End With
        With Me.moUILib.oDevice
            moBorderLine.Width = 2
            moBorderLine.Antialias = True
            moBorderLine.Begin()
            moBorderLine.Draw(moBorderLineVerts, BorderColor)
            moBorderLine.End()
        End With
        'End of the border drawing

        ''render the selected label
        'If ListIndex <> -1 Then
        '    If ListIndex >= moScroll.Value AndAlso ListIndex <= moScroll.Value + mlListItemUB Then
        '        'ok, it exists here...
        '        oRect = New System.Drawing.Rectangle(oLoc.X, oLoc.Y + moListItems(ListIndex - moScroll.Value).Top, Width, moListItems(ListIndex - moScroll.Value).Height)
        '        On Error Resume Next
        '        moUILib.oDevice.ColorFill(moUILib.oRenderTarget, oRect, HighlightColor)
        '        oRect = Nothing
        '    End If
        'End If

        ''Now, render our labels
        'For X = 0 To mlListItemUB
        '    If moScroll.Value + X > mlItemUB Then
        '        moListItems(X).Caption = ""
        '    Else
        '        moListItems(X).Caption = msItems(moScroll.Value + X)
        '    End If
        'Next X

    End Sub

    Private Sub UIListBox_OnResize() Handles MyBase.OnResize
        Dim X As Int32

        If Width = 0 Or Height = 0 Then Exit Sub
        'If True = True Then Return

        'set up our scrollbar
        With moScroll
            .Width = 24
            .Left = Me.Width - .Width
            .Height = Me.Height
            .Top = 0
        End With

        'remove our children from previous
        Me.RemoveAllChildren()
        Me.AddChild(CType(moScroll, UIControl))

        'now, set up our item labels
        mlListItemUB = CInt(Math.Floor(Me.Height / 24I)) - 1

        'ReDim moListItems(mlListItemUB)
        'For X = 0 To mlListItemUB
        '    moListItems(X) = New UILabel(moUILib)
        '    With moListItems(X)
        '        .Left = 3
        '        .Top = X * 24
        '        .Width = Me.Width - moScroll.Width - 3
        '        .Height = 24
        '        .Visible = True
        '        .Caption = ""
        '        .ForeColor = ForeColor
        '        .FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
        '    End With
        '    Me.AddChild(CType(moListItems(X), UIControl))

        '    AddHandler moListItems(X).OnMouseDown, AddressOf Label_MouseDown
        '    AddHandler moListItems(X).OnMouseUp, AddressOf Label_MouseUp
        'Next X

        'finally, go back and set up the scrollbar's max
        X = mlItemUB - mlListItemUB
        If X > -1 Then moScroll.MaxValue = X Else moScroll.MaxValue = 1

    End Sub

    Public Sub EnsureItemVisible(ByVal lIndex As Int32)
        'Ok, ensure that the item is visible
    End Sub
End Class

Public Class UIComboBox
    Inherits UIControl

    Protected WithEvents moDropDownList As UIListBox
    Protected WithEvents moDisplayValue As UILabel
    Protected WithEvents moDropDownButton As UIButton

    Private mbListBoxDisplayed As Boolean = False
    Private mlTempHeight As Int32       'used for expand/collapse

    Private moBorderLine As Line
    Private moBorderLineVerts(4) As Vector2

    Public BorderColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 255)
    Private moFillColor As System.Drawing.Color = System.Drawing.Color.White
    Private moSysFont As System.Drawing.Font
    Private moForeColor As System.Drawing.Color = System.Drawing.Color.Black

    Public Event ItemSelected(ByVal lItemIndex As Int32)

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        moBorderLine = New Line(oUILib.oDevice)

        moDropDownButton = New UIButton(oUILib)
        With moDropDownButton
            .Caption = ""
            .Width = 24
            .Left = Me.Width - .Width

            .ControlImageRect_Disabled = System.Drawing.Rectangle.FromLTRB(144, 96, 167, 119)
            .ControlImageRect_Normal = System.Drawing.Rectangle.FromLTRB(168, 96, 191, 119)
            .ControlImageRect_Pressed = System.Drawing.Rectangle.FromLTRB(192, 96, 215, 119)
        End With

        Me.AddChild(CType(moDropDownButton, UIControl))

        moDropDownList = New UIListBox(oUILib)
        moDropDownList.Visible = False
        moDropDownList.Width = Me.Width
        moDropDownList.BorderColor = System.Drawing.Color.Black
        Me.AddChild(CType(moDropDownList, UIControl))

        moDisplayValue = New UILabel(oUILib)
        With moDisplayValue
            .Caption = ""
            .Left = 3
            .Width = Me.Width - moDropDownButton.Width - 3
            .Height = Me.Height
            .FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(moDisplayValue, UIControl))

        'default font
        SetFont(New System.Drawing.Font("MS Sans Serif", 10))

    End Sub

#Region " Pass-thru to the ListBox Control "
    Public Sub AddItem(ByVal sItem As String)
        moDropDownList.AddItem(sItem)
    End Sub
    Public Sub Clear()
        moDropDownList.Clear()
    End Sub
    Public Sub RemoveItem(ByVal lIndex As Int32)
        moDropDownList.RemoveItem(lIndex)
    End Sub
    Public Property ItemData(ByVal lIndex As Int32) As Int32
        Get
            Return moDropDownList.ItemData(lIndex)
        End Get
        Set(ByVal Value As Int32)
            moDropDownList.ItemData(lIndex) = Value
        End Set
    End Property
    Public Property List(ByVal lIndex As Int32) As String
        Get
            Return moDropDownList.List(lIndex)
        End Get
        Set(ByVal Value As String)
            moDropDownList.List(lIndex) = Value
        End Set
    End Property
    Public ReadOnly Property ListCount() As Int32
        Get
            Return moDropDownList.ListCount
        End Get
    End Property
    Public Property ListIndex() As Int32
        Get
            Return moDropDownList.ListIndex
        End Get
        Set(ByVal Value As Int32)
            moDropDownList.ListIndex = Value
        End Set
    End Property

    Public Property FillColor() As System.Drawing.Color
        Get
            Return moFillColor
        End Get
        Set(ByVal Value As System.Drawing.Color)
            moFillColor = Value
            moDropDownList.FillColor = moFillColor
        End Set
    End Property
    Public Property DropDownListBorderColor() As System.Drawing.Color
        Get
            Return moDropDownList.BorderColor
        End Get
        Set(ByVal Value As System.Drawing.Color)
            moDropDownList.BorderColor = Value
        End Set
    End Property
    Public Property HighlightColor() As System.Drawing.Color
        Get
            Return moDropDownList.HighlightColor
        End Get
        Set(ByVal Value As System.Drawing.Color)
            moDropDownList.HighlightColor = Value
        End Set
    End Property
#End Region

    Public Property ForeColor() As System.Drawing.Color
        Get
            Return moForeColor
        End Get
        Set(ByVal Value As System.Drawing.Color)
            moForeColor = Value
            moDisplayValue.ForeColor = moForeColor
            moDropDownList.ForeColor = moForeColor
        End Set
    End Property

    Public Sub SetFont(ByVal oFont As System.Drawing.Font)
        moSysFont = oFont

        moDisplayValue.SetFont(moSysFont)
        moDropDownList.SetFont(moSysFont)

    End Sub

    Public Function GetFont() As System.Drawing.Font
        Return moSysFont
    End Function

    Private Sub moDropDownButton_Click() Handles moDropDownButton.Click
        If mbListBoxDisplayed = False Then
            'Display the listbox
            ExpandListBox()
        Else
            'undisplay the listbox
            CollapseListBox()
        End If
    End Sub

    Private Sub moDropDownList_ItemClick(ByVal lIndex As Integer) Handles moDropDownList.ItemClick
        'undisplay the listbox
        CollapseListBox()

        'set the label
        moDisplayValue.Caption = moDropDownList.List(lIndex)

        'raise the event
        RaiseEvent ItemSelected(lIndex)
    End Sub

    Private Sub UIComboBox_OnResize() Handles MyBase.OnResize
        If Width = 0 Or Height = 0 Then Exit Sub

        moDropDownList.Width = Me.Width

        If mbListBoxDisplayed = False Then
            moDisplayValue.Height = Me.Height
            moDisplayValue.Width = Me.Width - moDropDownButton.Width - 3
            moDropDownButton.Height = Me.Height
            moDropDownButton.Left = Me.Width - moDropDownButton.Width
        End If

    End Sub

    Private Sub UIComboBox_OnRender() Handles MyBase.OnRender
        Dim oLoc As System.Drawing.Point = GetAbsolutePosition()
        Dim X As Int32
        Dim lMax As Int32

        'Do a color fill of White
        On Error Resume Next
        Dim oRect As Rectangle = New Rectangle(oLoc.X, oLoc.Y, Width, Height)
        moUILib.oDevice.ColorFill(moUILib.oRenderTarget, oRect, moFillColor)
        oRect = Nothing
        'end of color fill

        'Draw a box border around our window...
        With moBorderLineVerts(0)
            .X = oLoc.X
            .Y = oLoc.Y
        End With
        With moBorderLineVerts(1)
            .X = oLoc.X + Width
            .Y = oLoc.Y
        End With
        With moBorderLineVerts(2)
            .X = oLoc.X + Width
            .Y = oLoc.Y + Height
        End With
        With moBorderLineVerts(3)
            .X = oLoc.X
            .Y = oLoc.Y + Height
        End With
        With moBorderLineVerts(4)
            .X = oLoc.X
            .Y = oLoc.Y
        End With
        With Me.moUILib.oDevice
            moBorderLine.Width = 2
            moBorderLine.Antialias = True
            moBorderLine.Begin()
            moBorderLine.Draw(moBorderLineVerts, BorderColor)
            moBorderLine.End()
        End With
        'End of the border drawing

    End Sub

    Private Sub CollapseListBox()
        mbListBoxDisplayed = False
        Me.Height = mlTempHeight
        moDropDownList.Visible = False
    End Sub

    Private Sub ExpandListBox()
        If mbListBoxDisplayed = True Then Exit Sub
        mbListBoxDisplayed = True
        mlTempHeight = Me.Height
        moDropDownList.Top = Me.Height + 1

        moDropDownList.Height = (Math.Max(1, Math.Min(4, moDropDownList.ListCount)) * 24)

        Me.Height = Me.Height + moDropDownList.Height
        moDropDownList.Visible = True

        If ListIndex <> -1 Then
            moDropDownList.EnsureItemVisible(ListIndex)
        End If
        mbListBoxDisplayed = True
    End Sub

End Class

Public Class UIHullSlots
    Inherits UIControl

    Private Const ml_GRID_SIZE_WH As Int32 = 30     'number of squares in each direction
    Private Const ml_SQUARE_SIZE As Int32 = 16      'Width and Height of each square

    Public Enum SlotType As Byte
        eUnused = 0
        eLeft = 1
        eFront = 2
        eRight = 3
        eRear = 4
    End Enum
    Public Enum SlotConfig As Integer
        eArmorConfig = 0
        eCrewQuarters = 1
        eCargoBay = 2
        eRadar = 3
        eEngines = 4
        eHangar = 5
        eShields = 6
        eComms = 7
        eWeapons = 8
    End Enum
    Private Structure HullSlot
        Public lX As Int32
        Public lY As Int32
        Public yType As SlotType
        Public lConfig As SlotConfig
        Public lGroupNum As Int32       'for separate groups
    End Structure

    Private moSlots(,) As HullSlot
    Public AutoRefresh As Boolean = True

    Private mcolVal() As System.Drawing.Color

    Public Event HullSlotClick(ByVal lIndexX As Int32, ByVal lIndexY As Int32)

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        ReDim moSlots(ml_GRID_SIZE_WH - 1, ml_GRID_SIZE_WH - 1)

        For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
            For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                With moSlots(X, Y)
                    .lX = X
                    .lY = Y
                    .lConfig = SlotConfig.eArmorConfig
                    .yType = SlotType.eFront 'SlotType.eUnused
                    .lGroupNum = 0
                End With
            Next X
        Next Y

        ReDim mcolVal(4)
        mcolVal(0) = System.Drawing.Color.FromArgb(255, 32, 32, 32)
        mcolVal(1) = System.Drawing.Color.FromArgb(255, 128, 0, 0)
        mcolVal(2) = System.Drawing.Color.FromArgb(255, 0, 128, 0)
        mcolVal(3) = System.Drawing.Color.FromArgb(255, 0, 0, 128)
        mcolVal(4) = System.Drawing.Color.FromArgb(255, 128, 0, 128)

        Me.Width = (ml_SQUARE_SIZE * ml_GRID_SIZE_WH) + 2
        Me.Height = Me.Width
    End Sub

    Private Sub UIHullSlots_OnMouseUp(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseUp
        Dim oLoc As System.Drawing.Point = Me.GetAbsolutePosition()
        Dim lTmpX As Int32 = lMouseX - oLoc.X - 1
        Dim lTmpY As Int32 = lMouseY - oLoc.Y - 1

        lTmpX \= ml_SQUARE_SIZE
        lTmpY \= ml_SQUARE_SIZE

        If lTmpX > 0 AndAlso lTmpY > 0 AndAlso lTmpX < ml_GRID_SIZE_WH AndAlso lTmpY < ml_GRID_SIZE_WH Then
            RaiseEvent HullSlotClick(lTmpX, lTmpY)
        End If
    End Sub

    Private Sub UIHullSlots_OnRender() Handles Me.OnRender
        Dim oLoc As System.Drawing.Point = GetAbsolutePosition()

        Dim oBorderLineVerts(4) As Vector2
        'Draw a box border around our window...
        With oBorderLineVerts(0)
            .X = oLoc.X
            .Y = oLoc.Y
        End With
        With oBorderLineVerts(1)
            .X = oLoc.X + Width
            .Y = oLoc.Y
        End With
        With oBorderLineVerts(2)
            .X = oLoc.X + Width
            .Y = oLoc.Y + Height
        End With
        With oBorderLineVerts(3)
            .X = oLoc.X
            .Y = oLoc.Y + Height
        End With
        With oBorderLineVerts(4)
            .X = oLoc.X
            .Y = oLoc.Y
        End With

        'Draw a box border around our window...
        Using moBorderLine As Line = New Line(MyBase.moUILib.oDevice)
            moBorderLine.Antialias = True
            moBorderLine.Width = 2
            moBorderLine.Begin()
            moBorderLine.Draw(oBorderLineVerts, Color.White)    'musettings.InterfaceBorderColor
            moBorderLine.End()
        End Using
        'End of the border drawing

        'Now, render our stuff
        Dim lTmpX As Int32
        Dim lTmpY As Int32

        Using oSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)

            Dim rcSrc As Rectangle = Rectangle.FromLTRB(0, 215, 10, 225)
            Dim rcDest As Rectangle
            Dim fX As Single
            Dim fY As Single

            oSpr.Begin(SpriteFlags.AlphaBlend)

            For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    With moSlots(X, Y)
                        If .yType <> SlotType.eUnused Then
                            lTmpX = oLoc.X + (X * ml_SQUARE_SIZE) + 1
                            lTmpY = oLoc.Y + (Y * ml_SQUARE_SIZE) + 1

                            fX = lTmpX * (rcSrc.Width / CSng(16))
                            fY = lTmpY * (rcSrc.Height / CSng(16))

                            rcDest = Rectangle.FromLTRB(oLoc.X, oLoc.Y, oLoc.X + 16, oLoc.Y + 16)

                            oSpr.Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcDest, System.Drawing.Point.Empty, 0, New Point(CInt(fX), CInt(fY)), mcolVal(.yType))
                        End If
                    End With
                Next X
            Next Y
            oSpr.End()
            oSpr.Dispose()
        End Using

        Using oFont As Font = New Font(MyBase.moUILib.oDevice, New System.Drawing.Font("Arial", 12, FontStyle.Bold, GraphicsUnit.Point, 0))
            Dim sCap As String
            Dim rcDest As Rectangle
            Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                oTextSpr.Begin(SpriteFlags.AlphaBlend)
                For Y As Int32 = 0 To ml_GRID_SIZE_WH - 1
                    For X As Int32 = 0 To ml_GRID_SIZE_WH - 1
                        With moSlots(X, Y)
                            lTmpX = oLoc.X + 1 + (X * ml_SQUARE_SIZE)
                            lTmpY = oLoc.Y + 1 + (Y * ml_SQUARE_SIZE)

                            rcDest = Rectangle.FromLTRB(lTmpX, lTmpY, lTmpX + 16, lTmpY + 16)
                            If .yType <> SlotType.eUnused Then
                                Select Case .yType
                                    Case SlotType.eFront
                                        sCap = "F"
                                    Case SlotType.eLeft
                                        sCap = "L"
                                    Case SlotType.eRear
                                        sCap = "B"
                                    Case SlotType.eRight
                                        sCap = "R"
                                    Case Else
                                        sCap = ""
                                End Select
                                If sCap <> "" Then
                                    oFont.DrawText(oTextSpr, sCap, rcDest, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, Color.White)
                                End If
                            End If
                        End With
                    Next X
                Next Y
                oTextSpr.End()
                oTextSpr.Dispose()
            End Using

            oFont.Dispose()
        End Using

    End Sub
End Class

Public Class UILabelScroller
    Inherits UIWindow

    Private WithEvents btnDecreaser As UIButton
    Private WithEvents lblCurrent As UILabel
    Private WithEvents btnIncreaser As UIButton

    Private mlIDs() As Int32
    Private msVals() As String
    Private mlUB As Int32 = -1

    Public Sub AddItem(ByVal lID As Int32, ByVal sDisplay As String)
        For X As Int32 = 0 To mlUB
            If mlIDs(X) = lID Then
                msVals(X) = sDisplay
                Return
            End If
        Next X

        mlUB += 1
        ReDim Preserve msVals(mlUB)
        ReDim Preserve mlIDs(mlUB)
    End Sub

    Public Sub RemoveItem(ByVal lID As Int32)
        For X As Int32 = 0 To mlUB
            If mlIDs(X) = lID Then
                For Y As Int32 = X + 1 To mlUB
                    mlIDs(Y - 1) = mlIDs(Y)
                    msVals(Y - 1) = msVals(Y)
                Next Y
                mlUB -= 1
                ReDim Preserve mlIDs(mlUB)
                ReDim Preserve msVals(mlUB)
                Exit For
            End If
        Next X
    End Sub

    Public Sub Clear()
        mlUB = -1
        Erase mlIDs
        Erase msVals
    End Sub

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        btnDecreaser = New UIButton(oUILib)
        With btnDecreaser
            .ControlName = "btnDecreaser"
            .Left = 1
            .Top = 1
            .Width = 24
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(145, 48, 24, 24)
        End With
        Me.AddChild(CType(btnDecreaser, UIControl))
 
        lblCurrent = New UILabel(oUILib)
        With lblCurrent
            .ControlName = "lblCurrent"
            .Left = 25
            .Top = 0
            .Width = 249
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = " "
            .ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblCurrent, UIControl))

        btnIncreaser = New UIButton(oUILib)
        With btnIncreaser
            .ControlName = "btnIncreaser"
            .Left = 150 - 19
            .Top = 1
            .Width = 24
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(145, 48, 24, 24)
        End With
        Me.AddChild(CType(btnIncreaser, UIControl))

        With Me
            .ControlName = "UILabelScroller"
            .BorderColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .FillColor = System.Drawing.Color.FromArgb(128, 32, 64, 128)
            .FullScreen = False
            .Height = 18
            .Width = 150
        End With
        SetButtonImageRects()
    End Sub

    Private Sub SetButtonImageRects()
        'here, we need to load the images for the buttons...
        With btnDecreaser
            '.ControlImage_Disabled = oUILib.oResMgr.GetTexture("HScr_Dec_Btn_Disabled.bmp")
            .ControlImageRect_Disabled = System.Drawing.Rectangle.FromLTRB(120, 48, 143, 71)
            '.ControlImage_Normal = oUILib.oResMgr.GetTexture("HScr_Dec_Btn_Normal.bmp")
            .ControlImageRect_Normal = System.Drawing.Rectangle.FromLTRB(144, 48, 167, 71)
            '.ControlImage_Pressed = oUILib.oResMgr.GetTexture("HScr_Dec_Btn_Pressed.bmp")
            .ControlImageRect_Pressed = System.Drawing.Rectangle.FromLTRB(168, 48, 191, 71)
        End With
        With btnIncreaser
            '.ControlImage_Disabled = oUILib.oResMgr.GetTexture("HScr_Inc_Btn_Disabled.bmp")
            .ControlImageRect_Disabled = System.Drawing.Rectangle.FromLTRB(120, 72, 143, 95)
            '.ControlImage_Normal = oUILib.oResMgr.GetTexture("HScr_Inc_Btn_Normal.bmp")
            .ControlImageRect_Normal = System.Drawing.Rectangle.FromLTRB(144, 72, 167, 95)
            '.ControlImage_Pressed = oUILib.oResMgr.GetTexture("HScr_Inc_Btn_Pressed.bmp")
            .ControlImageRect_Pressed = System.Drawing.Rectangle.FromLTRB(168, 72, 191, 95)
        End With
    End Sub

    Private Sub UILabelScroller_OnResize() Handles Me.OnResize
        Me.btnIncreaser.Left = Me.Width - (1 + Me.btnIncreaser.Width)
        Me.lblCurrent.Width = Me.Width - (2 + Me.btnIncreaser.Width + Me.btnDecreaser.Width)
    End Sub
End Class