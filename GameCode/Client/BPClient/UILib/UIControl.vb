Option Strict On

'this class is responsible for all user interfaces...
Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D
'everything is a control... even a window
Public Class UIControl
    Private Shared msTextBoxTypeName As String = GetType(UITextBox).ToString

    Public Shared mbToolTipSet As Boolean = False

    'NOTE: Location coordinates are parent-specific... meaning if this control is inside another control...
    '  then these cordinates represent the location INSIDE that control... NOT absolute
    Protected moLoc As Point = New Point(0, 0)
    Protected mlWidth As Int32
    Protected mlHeight As Int32
    Protected moControlRect As Rectangle
    Protected mbVisible As Boolean = True
    Public Property Visible() As Boolean
        Get
            Return mbVisible
        End Get
        Set(ByVal Value As Boolean)
            mbVisible = Value
            IsDirty = True
        End Set
    End Property
    Protected mbEnabled As Boolean = True
    Public moChildren() As UIControl
    Public ChildrenUB As Int32 = -1
    Private mbHasFocus As Boolean

    Public ControlImageRect As Rectangle
    Public ControlImageRect_Normal As Rectangle
    Public ControlImageRect_Disabled As Rectangle

    Protected moUILib As UILib

    Protected mbInDragMode As Boolean = False

    Public ControlName As String

    Public ToolTipText As String = ""

    Public mbAcceptReprocessEvents As Boolean = False

    Public AcceptMouseWheelEvents As Boolean = False
    Public Event OnMouseWheel(ByVal e As MouseEventArgs)

    'Always true unless otherwise stated...
    Public bAcceptEvents As Boolean = True

    Public Sub New(ByRef oUILib As UILib)
        moUILib = oUILib
    End Sub

    Private moParentControl As UIControl
    Private mbIsDirty As Boolean = True
    Private oCanvas As Texture
    Private mlCanvasWidth As Int32
    Private mlCanvasHeight As Int32

    Public ClickThru As Boolean = False

    'For Metrics Recording
    Protected mlIndexInsideParent As Int32 = -1
    Protected mlTopWindowMetricID As Int32 = -1
    'End Metrics Recording

    Public TabStop As Boolean = False
    Public TabIndex As Int32 = -1

    Public Event OnMouseDown(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons)
    Public Event OnMouseMove(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons)
    Public Event OnMouseUp(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons)
    Public Event OnKeyDown(ByVal e As KeyEventArgs)
    Public Event OnKeyUp(ByVal e As KeyEventArgs)
    Public Event OnKeyPress(ByVal e As KeyPressEventArgs)
    'Public Event OnRender(ByRef oImgSprite As Sprite, ByRef oTextSprite As Sprite)
    Public Event OnRender()
    Public Event OnRenderEnd()
    Public Event OnGotFocus()
    Public Event OnLostFocus()
    Public Event OnResize()
    Public Event ReprocessInput()       'for sustained inputs
    Public Event OnNewFrame()           'when a new frame occurs
    Public Event OnNewFrameEnd()          'when a new frame is finished
    Protected Event ResetInterfaceColors(ByVal yType As Byte, ByVal clrPrev As System.Drawing.Color)

    Public Property IsDirty() As Boolean
        Get
            Return mbIsDirty
        End Get
        Set(ByVal Value As Boolean)

            'If we're already dirty, then our parents are already dirty so if we're setting Dirty, then who cares?
            If Value = True AndAlso mbIsDirty = True Then Exit Property

            mbIsDirty = Value
            If mbIsDirty = True Then
                If moParentControl Is Nothing = False Then
                    moParentControl.IsDirty = True  'dirty up to the top-most parent
                End If
            End If
        End Set
    End Property

    Private Sub UpdateControlRect()
        moControlRect.Location = GetAbsolutePosition()
        moControlRect.Width = Width
        moControlRect.Height = Height
        IsDirty = True
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
			'RaiseEvent OnResize()
        End Set
    End Property

    Public Property Top() As Int32
        Get
            Return moLoc.Y
        End Get
        Set(ByVal Value As Int32)
            moLoc.Y = Value
            UpdateControlRect()
			'RaiseEvent OnResize()
        End Set
    End Property

    Public Property Width() As Int32
        Get
            Return mlWidth
        End Get
        Set(ByVal Value As Int32)
            mlWidth = Value
            UpdateControlRect()

            'because we are width or height, we need to destroy our canvas
            If oCanvas Is Nothing = False Then
                oCanvas.Dispose()
                oCanvas = Nothing
            End If

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

            'because we are width or height, we need to destroy our canvas
            If oCanvas Is Nothing = False Then
                oCanvas.Dispose()
                oCanvas = Nothing
            End If

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

                For X As Int32 = 0 To ChildrenUB
                    If moChildren(X) Is Nothing = False Then moChildren(X).HasFocus = False
                Next X
                RaiseEvent OnLostFocus()
            ElseIf mbHasFocus = False AndAlso Value = True Then
                mbHasFocus = Value
                RaiseEvent OnGotFocus()
            Else
                mbHasFocus = Value
            End If
            IsDirty = True

        End Set
    End Property

    Public Property Enabled() As Boolean
        Get
            Return mbEnabled
        End Get
        Set(ByVal Value As Boolean)
            Dim X As Int32
            mbEnabled = Value
            IsDirty = True      'this could get rough...
            For X = 0 To ChildrenUB
                moChildren(X).Enabled = Value
            Next X
        End Set
    End Property

    'Public Sub DrawControl()
    '    Dim X As Int32
    '    If Visible = True Then
    '        RaiseEvent OnRender()
    '        For X = 0 To ChildrenUB
    '            moChildren(X).DrawControl()
    '        Next X
    '    End If
    'End Sub

    'Public Sub DrawControl(ByRef oImgSprite As Sprite, ByRef oTextSprite As Sprite)
    Public Sub DrawControl()
        Dim X As Int32
        Dim bSetIsDirty As Boolean = False

        If GFXEngine.gbDeviceLost = True OrElse GFXEngine.gbPaused = True Then Return

        If Visible = True Then

            glUI_Rendered += 1

            RaiseEvent OnNewFrame()

            If moParentControl Is Nothing Then
                'MSC - added this so that moving windows appears to perform excessively better
                '  What this does is it will only redraw the window if the window is not moving or if the canvas was lost
                '  this little state machine ensures that the window does redraw once the window is done moving (in its final location)
                If moUILib.lUISelectState = UILib.eSelectState.eMoveWindow AndAlso Object.Equals(moUILib.oMovingWindow, Me) = True Then
                    'Ok, I am moving...
                    bSetIsDirty = True
                    mbIsDirty = False
                End If
                If mbIsDirty = True OrElse oCanvas Is Nothing Then
                    Dim oOrigSurf As Surface
                    Dim lTempX As Int32 = Left
                    Dim lTempY As Int32 = Top
                    Left = 0
                    Top = 0

                    If oCanvas Is Nothing Then
                        If Me.Width < 65 Then
                            mlCanvasWidth = 64
                        ElseIf Me.Width < 129 Then
                            mlCanvasWidth = 128
                        ElseIf Me.Width < 257 Then
                            mlCanvasWidth = 256
                        ElseIf Me.Width < 513 Then
                            mlCanvasWidth = 512
                        Else : mlCanvasWidth = 1024
                        End If

                        If Me.Height < 65 Then
                            mlCanvasHeight = 64
                        ElseIf Me.Height < 129 Then
                            mlCanvasHeight = 128
                        ElseIf Me.Height < 257 Then
                            mlCanvasHeight = 256
                        ElseIf Me.Height < 513 Then
                            mlCanvasHeight = 512
                        Else : mlCanvasHeight = 1024
                        End If

                        If mlCanvasWidth > moUILib.oDevice.DepthStencilSurface.Description.Width Then
                            mlCanvasWidth = moUILib.oDevice.DepthStencilSurface.Description.Width
                        End If
                        If mlCanvasHeight > moUILib.oDevice.DepthStencilSurface.Description.Height Then
                            mlCanvasHeight = moUILib.oDevice.DepthStencilSurface.Description.Height
                        End If

                        Device.IsUsingEventHandlers = False
                        oCanvas = New Texture(moUILib.oDevice, mlCanvasWidth, mlCanvasHeight, 1, Usage.RenderTarget, Format.A8R8G8B8, Pool.Default)
                        Device.IsUsingEventHandlers = True
                    End If
                    Dim oTmpCanvas As Texture = oCanvas
                    If oTmpCanvas Is Nothing = False Then
                        'moUILib.oDevice.RenderState.ZBufferWriteEnable = False
                        moUILib.oDevice.RenderState.ZBufferEnable = False

                        oOrigSurf = moUILib.oDevice.GetRenderTarget(0)
                        If oTmpCanvas Is Nothing = False Then
                            Try
                                moUILib.oDevice.SetRenderTarget(0, oTmpCanvas.GetSurfaceLevel(0))
                            Catch
                                Return
                            End Try
                        Else : Return
                        End If
                        moUILib.oDevice.Clear(ClearFlags.Target, System.Drawing.Color.FromArgb(0, 0, 0, 0), 1, 0)

                        RaiseEvent OnRender()
                        Dim lCurUB As Int32 = -1
                        If moChildren Is Nothing = False Then lCurUB = Math.Min(ChildrenUB, moChildren.GetUpperBound(0))
                        For X = 0 To lCurUB
                            If moChildren(X) Is Nothing = False Then moChildren(X).DrawControl()
                        Next X
                        RaiseEvent OnRenderEnd()

                        moUILib.oDevice.SetRenderTarget(0, oOrigSurf)
                        'If moUILib.oDevice.RenderState.ZBufferWriteEnable = True OrElse moUILib.oDevice.RenderState.ZBufferEnable = True Then
                        '	goUILib.AddNotification("Shazbot", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        'End If
                        'moUILib.oDevice.RenderState.ZBufferWriteEnable = True
                        moUILib.oDevice.RenderState.ZBufferEnable = True

                        oOrigSurf.Dispose() : oOrigSurf = Nothing
                    End If
                    oTmpCanvas = Nothing


                    Left = lTempX
                    Top = lTempY
                End If

                'Now, do our render to the final location
                Dim oRect As Rectangle
                Dim oLoc As System.Drawing.Point

                oLoc = GetAbsolutePosition()
                With oRect
                    .Location() = oLoc
                    .Width = Width
                    .Height = Height

                    'Dim uVerts(3) As CustomVertex.TransformedTextured

                    'uVerts(0) = New CustomVertex.TransformedTextured(oLoc.X - 0.5F, oLoc.Y - 0.5F, 1.0F, 1.0F, 0, 0)
                    'uVerts(1) = New CustomVertex.TransformedTextured(oLoc.X + .Width - 0.5F, oLoc.Y - 0.5F, 1.0F, 1.0F, 1, 0)
                    'uVerts(2) = New CustomVertex.TransformedTextured(oLoc.X - 0.5F, oLoc.Y + .Height - 0.5F, 1.0F, 1.0F, 0, 1)
                    'uVerts(3) = New CustomVertex.TransformedTextured(oLoc.X + .Width - 0.5F, oLoc.Y + .Height - 0.5F, 1.0F, 1.0F, 1, 1)

                    'With moUILib.oDevice
                    '    .VertexFormat = CustomVertex.TransformedTextured.Format
                    '    .SetTexture(0, oCanvas)
                    '    Dim oMat As Material
                    '    oMat.Ambient = Color.White
                    '    oMat.Diffuse = Color.White
                    '    .Material = oMat
                    '    .DrawUserPrimitives(PrimitiveType.TriangleList, 2, uVerts)
                    'End With
                    'moUILib.oDevice.RenderState.ZBufferEnable = False
                    If oCanvas Is Nothing = False Then
                        With moUILib.Pen
                            Try
                                'If moUILib.bPenBegun = False Then .Begin(SpriteFlags.AlphaBlend)
                                '.Begin(SpriteFlags.AlphaBlend)
                                moUILib.BeginPenSprite(SpriteFlags.AlphaBlend)
                                Try
                                    .Draw2D(oCanvas, Rectangle.FromLTRB(0, 0, Me.Width + 1, Me.Height + 1), System.Drawing.Rectangle.FromLTRB(oLoc.X, oLoc.Y, oLoc.X + Width + 1, oLoc.Y + Height + 1), System.Drawing.Point.Empty, 0, New Point(oLoc.X, oLoc.Y), System.Drawing.Color.FromArgb(&HFFFFFFFF))
                                Catch
                                End Try
                                'If moUILib.bPenBegun = False Then .End()
                                '.End()
                            Catch
                            Finally
                                Try
                                    moUILib.EndPenSprite()
                                Catch
                                End Try
                            End Try
                        End With
                    End If
                    'moUILib.oDevice.RenderState.ZBufferEnable = True
                End With

                RaiseEvent OnNewFrameEnd()
            Else
                'RaiseEvent OnRender(oImgSprite, oTextSprite)
                RaiseEvent OnRender()
                Dim lCurUB As Int32 = -1
                If moChildren Is Nothing = False Then lCurUB = Math.Min(ChildrenUB, moChildren.GetUpperBound(0))
                For X = 0 To lCurUB
                    'moChildren(X).DrawControl(oImgSprite, oTextSprite)
                    'moUILib.oDevice.RenderState.ZBufferEnable = False
                    moChildren(X).DrawControl()
                    'moUILib.oDevice.RenderState.ZBufferEnable = True
                Next X
                RaiseEvent OnRenderEnd()
            End If

            If bSetIsDirty = True Then mbIsDirty = True Else mbIsDirty = False
		End If
    End Sub

    Public Function PostMouseWheel(ByVal e As MouseEventArgs) As Boolean

        If Me.Visible = True AndAlso Me.Enabled = True AndAlso (Me.ParentControl Is Nothing = False OrElse TestRegion(e.X, e.Y) = True) Then
            'post to children?
            For X As Int32 = 0 To ChildrenUB
                If moChildren(X) Is Nothing = False Then
                    If moChildren(X).PostMouseWheel(e) = True Then Return True
                End If
            Next X

            'if not, check my type... do i accept mouse wheel events?
            If Me.AcceptMouseWheelEvents = True Then
                'Ok, raise my mouse wheel event
                RaiseEvent OnMouseWheel(e)
                Return True
            End If
        End If

        Return False
    End Function


    Public Function PostMessage(ByVal lMsgCode As UILibMsgCode, ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons) As Boolean
        If ClickThru = True Then Return False

        Dim oTemp As UIControl = Nothing

        If Me.bAcceptEvents = False OrElse Enabled = False Then
            If lMsgCode = UILibMsgCode.eMouseMoveMsgCode Then
                If TestRegion(lMouseX, lMouseY) = True AndAlso Me.Visible = True Then
                    If ToolTipText <> "" Then
                        moUILib.SetToolTip(ToolTipText, lMouseX, GetAbsolutePosition.Y + Height + 5)
                        mbToolTipSet = True
                    Else
                        moUILib.SetToolTip(False)
                    End If
                End If
            End If
            Return False
        End If

        If Visible = True AndAlso Enabled = True AndAlso TestRegion(lMouseX, lMouseY) = True Then
            If lButton <> MouseButtons.None AndAlso Me.ParentControl Is Nothing Then ' = False Then
                moUILib.BringWindowToForeground(Me.ControlName)
            End If

            If mbInDragMode = False Then
                oTemp = PostToChildren(lMsgCode, lMouseX, lMouseY, lButton)
            End If

            If oTemp Is Nothing Then
                Select Case lMsgCode
                    Case UILibMsgCode.eMouseDownMsgCode
                        BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eUIElementClick, Me.mlTopWindowMetricID, Me.mlIndexInsideParent)
                        RaiseEvent OnMouseDown(lMouseX, lMouseY, lButton)
                    Case UILibMsgCode.eMouseMoveMsgCode
                        'If NewTutorialManager.TutorialOn = True Then
                        '	If moUILib.CommandAllowed(True, Me.GetFullName) = False Then Return True
                        'End If
                        If mbToolTipSet = False Then
                            If ToolTipText <> "" Then
                                moUILib.SetToolTip(ToolTipText, lMouseX, GetAbsolutePosition.Y + Height + 5)
                            ElseIf Me.ParentControl Is Nothing = False Then
                                moUILib.SetToolTip(False)
                            End If
                        End If
                        'UIElementHover goes here
                        RaiseEvent OnMouseMove(lMouseX, lMouseY, lButton)
                    Case UILibMsgCode.eMouseUpMsgCode
                        RaiseEvent OnMouseUp(lMouseX, lMouseY, lButton)
                End Select
            End If

            Return True
        Else
            mbInDragMode = False
            Return False
        End If
    End Function

    Public Function PostMessage(ByVal lMsgCode As UILibMsgCode, ByVal eMsg As System.Windows.Forms.KeyEventArgs) As Boolean
        Dim X As Int32

        If HasFocus Then
            If Me.ParentControl Is Nothing = False Then
                moUILib.BringWindowToForeground(Me.ControlName)
            End If

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
                        If Me.ParentControl Is Nothing = False Then
                            moUILib.BringWindowToForeground(Me.ControlName)
                        End If
                        moChildren(X).PostMessage(lMsgCode, eMsg)
                        Return True
                    ElseIf moChildren(X).ChildrenUB <> -1 Then
                        'check children's children...
                        If Me.ParentControl Is Nothing = False Then
                            moUILib.BringWindowToForeground(Me.ControlName)
                        End If
                        If moChildren(X).PostToChildren(lMsgCode, eMsg) = True Then Return True
                    End If
                End If
            Next X
        End If
    End Function

    Public Function PostMessage(ByVal lMsgCode As UILibMsgCode, ByVal eMsg As System.Windows.Forms.KeyPressEventArgs) As Boolean
        Dim X As Int32

		If HasFocus Then
            RaiseEvent OnKeyPress(eMsg)
            If Me.ParentControl Is Nothing = False Then
                moUILib.BringWindowToForeground(Me.ControlName)
            End If
			Return True
		Else
			'check children
			For X = 0 To ChildrenUB
				If Not moChildren(X) Is Nothing Then
					If moChildren(X).HasFocus Then
                        If Me.ParentControl Is Nothing = False Then
                            moUILib.BringWindowToForeground(Me.ControlName)
                        End If
                        moChildren(X).PostMessage(lMsgCode, eMsg)
						Return True
					ElseIf moChildren(X).ChildrenUB <> -1 Then
                        If moChildren(X).PostToChildren(lMsgCode, eMsg) = True Then
                            If Me.ParentControl Is Nothing = False Then
                                moUILib.BringWindowToForeground(Me.ControlName)
                            End If
                            Return True
                        End If
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

		'If Visible And Enabled Then
		If lMouseX >= lX AndAlso lMouseX <= (lX + mlWidth) Then
			If lMouseY >= lY AndAlso lMouseY <= (lY + mlHeight) Then
				Return True
			End If
		End If
		'End If
		Return False
    End Function

    Private Function PostToChildren(ByVal lMsgCode As UILibMsgCode, ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As MouseButtons) As UIControl
        Dim X As Int32

        For X = 0 To ChildrenUB
			If Not moChildren(X) Is Nothing AndAlso moChildren(X).Enabled = True AndAlso moChildren(X).Visible = True Then
				If moChildren(X).PostMessage(lMsgCode, lMouseX, lMouseY, lButton) = True Then
					Return moChildren(X)
				ElseIf moChildren(X).ChildrenUB <> -1 Then
					Dim oTmp As UIControl = moChildren(X).PostToChildren(lMsgCode, lMouseX, lMouseY, lButton)
					If oTmp Is Nothing = False Then Return oTmp
				End If
			End If
        Next X
        Return Nothing
    End Function

    Private Function PostToChildren(ByVal lMsgCode As UILibMsgCode, ByVal eMsg As System.Windows.Forms.KeyEventArgs) As Boolean
        Dim X As Int32

        For X = 0 To ChildrenUB
            If Not moChildren(X) Is Nothing Then
                If moChildren(X).PostMessage(lMsgCode, eMsg) Then
                    Return True
                End If
            End If
        Next X
        Return False
    End Function

    Private Function PostToChildren(ByVal lMsgCode As UILibMsgCode, ByVal eMsg As System.Windows.Forms.KeyPressEventArgs) As Boolean
        Dim X As Int32

        For X = 0 To ChildrenUB
            If Not moChildren(X) Is Nothing Then
                If moChildren(X).PostMessage(lMsgCode, eMsg) Then
                    Return True
                End If
            End If
        Next X
        Return False
    End Function

    Public Sub AddChild(ByRef oChild As UIControl)
        Dim sNewChildTypeStr As String = ""
        If oChild Is Nothing = False AndAlso oChild.ControlName Is Nothing = False Then sNewChildTypeStr = oChild.ControlName.ToUpper

        If sNewChildTypeStr.StartsWith("TXT") = True Then
            Dim oTxt As UITextBox = CType(oChild, UITextBox)
            If oTxt.Locked = False Then
                oTxt.TabStop = True
            End If
        ElseIf sNewChildTypeStr.StartsWith("LST") = True Then
            oChild.TabStop = True
        ElseIf sNewChildTypeStr.StartsWith("TP") = True Then
            oChild.TabStop = True
        End If

        'Now, place it
        If oChild.TabStop = True AndAlso oChild.TabIndex < 0 Then
            Dim lNextIdx As Int32 = 0
            Dim lNextHighestTop As Int32 = Int32.MaxValue
            For X As Int32 = 0 To ChildrenUB
                If moChildren(X) Is Nothing = False Then
                    With moChildren(X)
                        If .TabStop = True Then
                            If .Top > oChild.Top Then
                                If lNextHighestTop > .Top Then
                                    lNextHighestTop = .Top
                                    lNextIdx = .TabIndex
                                End If
                            Else
                                If .TabIndex >= lNextIdx Then lNextIdx = .TabIndex + 1
                            End If
                        End If
                    End With
                End If
            Next X
            oChild.TabIndex = lNextIdx
            For X As Int32 = 0 To ChildrenUB
                If moChildren(X) Is Nothing = False Then
                    With moChildren(X)
                        If .TabStop = True Then
                            If .TabIndex >= lNextIdx Then .TabIndex += 1
                        End If
                    End With
                End If
            Next X
        End If

        ChildrenUB += 1
        ReDim Preserve moChildren(ChildrenUB)
        oChild.ParentControl = Me
        moChildren(ChildrenUB) = oChild
        oChild.mlIndexInsideParent = ChildrenUB
        oChild.mlTopWindowMetricID = oChild.GetParentMetricID()

        IsDirty = True

        If oChild.mbAcceptReprocessEvents = True Then Me.mbAcceptReprocessEvents = True
    End Sub

    Protected Function GetParentMetricID() As Int32
        Dim oParent As UIControl = Me.ParentControl
        If oParent Is Nothing Then Return -1
        While oParent.ParentControl Is Nothing = False
            oParent = oParent.ParentControl
        End While
        If TypeOf oParent Is UIWindow Then
            Return CType(oParent, UIWindow).lWindowMetricID
        End If
        Return -1
    End Function

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
        IsDirty = True
    End Sub

    Public Sub RemoveAllChildren()
        Dim X As Int32
        For X = 0 To ChildrenUB
            moChildren(X) = Nothing
        Next X
        ChildrenUB = -1
        Erase moChildren
        IsDirty = True
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

    Public Function GetRelativeTop() As Int32 'Relative to this UiWindow (Including +'s for Sub Frames / UiWindows), not the entire display.
        Dim lTop As Int32
        Dim lLastParentTop As Int32

        lTop = Me.Top
        Dim oParent As UIControl = Me.ParentControl
        While oParent Is Nothing = False
            lLastParentTop = oParent.Top
            lTop += lLastParentTop
            oParent = oParent.ParentControl
        End While
        lTop -= lLastParentTop 'Negate the last top as it's the main Form's window.  
        Return lTop
    End Function

    Public Sub ReprocessInputs()
        Dim X As Int32
        If Visible AndAlso mbAcceptReprocessEvents Then
            RaiseEvent ReprocessInput()
            For X = 0 To ChildrenUB
                If moChildren(X) Is Nothing = False Then moChildren(X).ReprocessInputs()
            Next X
        End If
    End Sub

    Protected Overrides Sub Finalize()
		'Debug.Write("Control '" & Me.ControlName & "' Finalized" & vbCrLf)
		If oCanvas Is Nothing = False Then oCanvas.Dispose()
		oCanvas = Nothing
        RemoveAllChildren()
        moUILib = Nothing
        MyBase.Finalize()
    End Sub

    Public Sub ProcessTabButton(ByVal sFromControl As String)

        Dim lBaseValue As Int32 = -1
        If sFromControl = "" Then
            'Ok, first time calling this... do I have children?
            If Me.ChildrenUB > -1 Then
                lBaseValue = -1
            End If
        Else
            If moUILib.FocusedControl Is Nothing = False AndAlso moUILib.FocusedControl.ControlName = sFromControl Then
                lBaseValue = moUILib.FocusedControl.TabIndex
            Else
                For X As Int32 = 0 To Me.ChildrenUB
                    If Me.moChildren(X).ControlName = sFromControl Then
                        lBaseValue = Me.moChildren(X).TabIndex
                        Exit For
                    End If
                Next X
            End If
        End If
 
        Dim lIdx As Int32 = -1
        Dim lLowest As Int32 = Int32.MaxValue
        For X As Int32 = 0 To Me.ChildrenUB
            With moChildren(X)
                'Debug.Print(.ControlName & " " & .TabIndex.ToString & " " & .Enabled.ToString & " " & .TabStop.ToString)
                If .Enabled = True AndAlso .Visible = True AndAlso .TabStop = True Then
                    If .TabIndex > lBaseValue AndAlso .TabIndex < lLowest Then
                        lIdx = X
                        lLowest = .TabIndex
                    End If
                End If
            End With
        Next X

        'Ok, did we find one?
        If lIdx <> -1 Then
            'Yes... set it to focused
            If moUILib.FocusedControl Is Nothing = False Then moUILib.FocusedControl.HasFocus = False
            moUILib.FocusedControl = Nothing
            moUILib.FocusedControl = Me.moChildren(lIdx)
            Me.moChildren(lIdx).HasFocus = True

            Return
        ElseIf lBaseValue <> -1 AndAlso Me.ParentControl Is Nothing = True Then
            For X As Int32 = 0 To Me.ChildrenUB
                With moChildren(X)
                    'Debug.Print(.ControlName & " " & .TabIndex.ToString & " " & .Enabled.ToString & " " & .TabStop.ToString)
                    If .Enabled = True AndAlso .Visible = True AndAlso .TabStop = True Then
                        If .TabIndex > -1 AndAlso .TabIndex < lLowest Then
                            lIdx = X
                            lLowest = .TabIndex
                        End If
                    End If
                End With
            Next X

            If lIdx <> -1 Then
                'Yes... set it to focused
                If moUILib.FocusedControl Is Nothing = False Then moUILib.FocusedControl.HasFocus = False
                moUILib.FocusedControl = Nothing
                moUILib.FocusedControl = Me.moChildren(lIdx)
                Me.moChildren(lIdx).HasFocus = True

                Return
            End If
        End If

        'Ok, either we don't have children, or I did not find one suitable to go to next... so we will raise this to my parent
        If Me.ParentControl Is Nothing = False Then
            'TODO: When passing back to a parent, tell it my name, so the parent can say 'Ok.. you are at the last tab item, on this Frame, goto Next frame in .Top order; don't assume PArent -> control 0
            Me.ParentControl.ProcessTabButton(Me.ControlName)
        End If

    End Sub

    Public Sub ProcessShiftTabButton(ByVal sFromControl As String)
        'Same as process tab button except, we go in reverse...
        Dim lBaseValue As Int32 = -1
        If sFromControl = "" Then
            'Ok, first time calling this... do I have children?
            If Me.ChildrenUB > -1 Then
                lBaseValue = -1
            End If
        Else
            If moUILib.FocusedControl Is Nothing = False AndAlso moUILib.FocusedControl.ControlName = sFromControl Then
                lBaseValue = moUILib.FocusedControl.TabIndex
            Else
                For X As Int32 = 0 To Me.ChildrenUB
                    If Me.moChildren(X).ControlName = sFromControl Then
                        lBaseValue = Me.moChildren(X).TabIndex
                        Exit For
                    End If
                Next X
            End If
        End If

        Dim lIdx As Int32 = -1
        Dim lLowest As Int32 = Int32.MinValue
        For X As Int32 = 0 To Me.ChildrenUB
            With moChildren(X)
                If .Enabled = True AndAlso .Visible = True AndAlso .TabStop = True Then
                    If .TabIndex > -1 AndAlso .TabIndex < lBaseValue AndAlso .TabIndex > lLowest Then
                        lIdx = X
                        lLowest = .TabIndex
                    End If
                End If
            End With
        Next X

        'Ok, did we find one?
        If lIdx <> -1 Then
            'Yes... set it to focused
            If moUILib.FocusedControl Is Nothing = False Then moUILib.FocusedControl.HasFocus = False
            moUILib.FocusedControl = Nothing
            moUILib.FocusedControl = Me.moChildren(lIdx)
            Me.moChildren(lIdx).HasFocus = True

            Return
        ElseIf lBaseValue <> -1 AndAlso Me.ParentControl Is Nothing = True Then
            For X As Int32 = 0 To Me.ChildrenUB
                With moChildren(X)
                    If .Enabled = True AndAlso .Visible = True AndAlso .TabStop = True Then
                        If .TabIndex > -1 AndAlso .TabIndex > lLowest Then
                            lIdx = X
                            lLowest = .TabIndex
                        End If
                    End If
                End With
            Next X

            If lIdx <> -1 Then
                'Yes... set it to focused
                If moUILib.FocusedControl Is Nothing = False Then moUILib.FocusedControl.HasFocus = False
                moUILib.FocusedControl = Nothing
                moUILib.FocusedControl = Me.moChildren(lIdx)
                Me.moChildren(lIdx).HasFocus = True

                Return
            End If
        End If

        'Ok, either we don't have children, or I did not find one suitable to go to next... so we will raise this to my parent
        If Me.ParentControl Is Nothing = False Then
            Me.ParentControl.ProcessShiftTabButton(Me.ControlName)
        End If
	End Sub

	Public Sub KillCanvas()
		Try
			If oCanvas Is Nothing = False Then oCanvas.Dispose()
			oCanvas = Nothing
		Catch
		End Try
    End Sub

    Public Sub KillChildrenDefaultPoolResources()
        Dim sTypeStr As String = Me.GetType.ToString.ToUpper

        If sTypeStr.Contains("UILABEL") = True OrElse sTypeStr.Contains("UIOPTION") = True OrElse sTypeStr.Contains("UICHECKBOX") = True OrElse sTypeStr.Contains("UIBUTTON") = True OrElse sTypeStr.Contains("UITEXTBOX") = True Then
            CType(Me, UILabel).ReleaseFont()
        End If

        For X As Int32 = 0 To Me.ChildrenUB
            If moChildren(X) Is Nothing = False Then moChildren(X).KillChildrenDefaultPoolResources()
        Next X
    End Sub

	Public Function GetFullName() As String
		Dim sFullName As String = Me.ControlName
		Dim oCtrl As UIControl = Me

		While oCtrl.ParentControl Is Nothing = False
			sFullName = oCtrl.ParentControl.ControlName & "." & sFullName
			oCtrl = oCtrl.ParentControl
		End While
		Return sFullName
    End Function

    Public Sub DoResetInterfaceColors(ByVal yType As Byte, ByVal clrPrev As System.Drawing.Color)
        RaiseEvent ResetInterfaceColors(yType, clrPrev)
        For X As Int32 = 0 To Me.ChildrenUB
            If Me.moChildren(X) Is Nothing = False Then Me.moChildren(X).DoResetInterfaceColors(yType, clrPrev)
        Next X
    End Sub

End Class
