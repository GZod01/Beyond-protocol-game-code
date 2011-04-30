Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class UITreeView
	Inherits UIControl

	Public Class UITreeViewItem
		Public sItem As String = ""
		Public lItemData As Int32
		Public lItemData2 As Int32
		Public lItemData3 As Int32
		Public bItemBold As Boolean
		Public bItemLocked As Boolean
        Public clrItemColor As System.Drawing.Color
        Public bUseItemColor As Boolean = False
        Public bClickIcon As Boolean = False

        Public clrFillColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 128, 128, 128)
		Public bUseFillColor As Boolean = False

		Public bExpanded As Boolean = False		'indicates this node item is expanded

        Public oRelatedObject As Object = Nothing

        Public lTop As Int32    'inside of the control

        Public bRenderIcon As Boolean = False
        Public IconRect As Rectangle
        Public IconColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 255)

		'Linked List Definitions...
		Public oParentNode As UITreeViewItem = Nothing
		Public oFirstChild As UITreeViewItem = Nothing
		Public oNextSibling As UITreeViewItem = Nothing
		Public oPrevSibling As UITreeViewItem = Nothing

		Public lTreeViewIndex As Int32 = -1

		Public Function GetLeftOffset() As Int32
			Dim lOffset As Int32 = 10
			Dim oParent As UITreeViewItem = oParentNode
			While oParent Is Nothing = False
				lOffset += 10
				oParent = oParent.oParentNode
			End While
			Return lOffset
		End Function

		Public Sub ExpandToRoot()
			bExpanded = True
			Dim oTmp As UITreeViewItem = oParentNode
			While oTmp Is Nothing = False
				oTmp.bExpanded = True
				oTmp = oTmp.oParentNode
			End While
		End Sub
	End Class

	Protected WithEvents mvscrVert As UIScrollBar		'vertical scroller
	Protected WithEvents mhscrHoriz As UIScrollBar		'horizontal scroller

	Public oRootNode As UITreeViewItem = Nothing	   'Root nodes

	Private moAllNodes() As UITreeViewItem
	Private mlAllNodeUB As Int32 = -1

    'Private moBorderLine As Line
	Private moBorderLineVerts(4) As Vector2

	Private mlListItemHeight As Int32 = 18
	Private moSysFont As System.Drawing.Font
	Private moSysFontBold As System.Drawing.Font

    Public Event NodeDoubleClicked()
    Public Event NodeExpanded(ByVal oNode As UITreeViewItem)
    Public Event NodeUnexpanded(ByRef oNode As UITreeViewItem)
    Public Event NodeIconClicked(ByRef oNode As UITreeViewItem)

    Public oIconTexture As Texture
    Public rcIconDimensions As Rectangle

	Private moSelectedNode As UITreeViewItem = Nothing
	Public Property oSelectedNode() As UITreeViewItem
		Get
			Return moSelectedNode
		End Get
		Set(ByVal value As UITreeViewItem)
            'Dim bDifferent As Boolean = Object.Equals(value, moSelectedNode) = False
            moSelectedNode = value
            If moSelectedNode Is Nothing = False Then
                If moSelectedNode.oFirstChild Is Nothing = False Then
                    If moSelectedNode.bExpanded = False Then
                        moSelectedNode.bExpanded = True
                        RaiseEvent NodeExpanded(value)
                    End If
                End If
            End If
            'If bdifferent = True Then
            RaiseEvent NodeSelected(value)
            'End If
        End Set
    End Property
    Public Property oSelectedNodeNoExpand() As UITreeViewItem
        Get
            Return moSelectedNode
        End Get
        Set(ByVal value As UITreeViewItem)
            'Dim bDifferent As Boolean = Object.Equals(value, moSelectedNode) = False
            moSelectedNode = value
            'If bdifferent = True Then
            RaiseEvent NodeSelected(value)
            'End If
        End Set
    End Property

    Public Sub SelectNodeByObject(ByRef oObj As Object)
        For X As Int32 = 0 To mlAllNodeUB
            If moAllNodes(X) Is Nothing = False Then
                If moAllNodes(X).oRelatedObject Is Nothing = False Then
                    If Object.Equals(moAllNodes(X).oRelatedObject, oObj) = True Then
                        oSelectedNode = moAllNodes(X)
                        Exit For
                    End If
                End If
            End If
        Next X
    End Sub

    Public Event NodeSelected(ByRef oNode As UITreeViewItem)

    Public ReadOnly Property TotalNodeCount() As Int32
        Get
            Return mlAllNodeUB + 1
        End Get
    End Property

#Region "  Control Specific Properties  "
	Private moBorderColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 255)
	Public Property BorderColor() As System.Drawing.Color
		Get
			Return moBorderColor
		End Get
		Set(ByVal Value As System.Drawing.Color)
			moBorderColor = Value
			IsDirty = True
		End Set
	End Property
	Private moFillColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 255, 255)
	Public Property FillColor() As System.Drawing.Color
		Get
			Return moFillColor
		End Get
		Set(ByVal Value As System.Drawing.Color)
			moFillColor = Value
			IsDirty = True
		End Set
	End Property
	Private moHighlightColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 32, 64, 128)
	Private moHighlightTextColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 255, 0, 0)
	Public Property HighlightColor() As System.Drawing.Color
		Get
			Return moHighlightColor
		End Get
		Set(ByVal Value As System.Drawing.Color)
			moHighlightColor = Value
			moHighlightTextColor = Color.FromArgb(255, 255 - moHighlightColor.R, 255 - moHighlightColor.G, 255 - moHighlightColor.B)
			IsDirty = True
		End Set
	End Property
	Private moForeColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 0, 0)
	Public Property ForeColor() As System.Drawing.Color
		Get
			Return moForeColor
		End Get
		Set(ByVal Value As System.Drawing.Color)
			moForeColor = Value
			IsDirty = True
		End Set
	End Property
#End Region

	Public Sub SetFont(ByVal oFont As System.Drawing.Font)
		'Dim X As Int32

		moSysFont = oFont
		moSysFontBold = New System.Drawing.Font(oFont, FontStyle.Bold)

		'UITreeView_OnResize()
		IsDirty = True

	End Sub

	Public Function GetFont() As System.Drawing.Font
		Return moSysFont
	End Function

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		Me.mbAcceptReprocessEvents = True

        'Device.IsUsingEventHandlers = False
        'moBorderLine = New Line(oUILib.oDevice)
        'Device.IsUsingEventHandlers = True

		'Default font
		SetFont(New System.Drawing.Font("MS Sans Serif", 10))

		mvscrVert = New UIScrollBar(oUILib, True)
		mvscrVert.ReverseDirection = True
		mvscrVert.mbAcceptReprocessEvents = True
		With mvscrVert
			.Width = 24
			.Left = Me.Width - .Width
			.Height = Me.Height - 2
			.Top = 1
			.Visible = True
			.mbAcceptReprocessEvents = True
		End With
		Me.AddChild(CType(mvscrVert, UIControl))
	End Sub 

	Private Sub UITreeView_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.OnKeyDown
		'Select Case e.KeyCode
		'	Case Keys.Home
		'		ListIndex = 0
		'		EnsureItemVisible(0)
		'	Case Keys.End
		'		ListIndex = ListCount - 1
		'		EnsureItemVisible(ListIndex)
		'	Case Keys.Up
		'		mlLastSetListIndexCall = 0
		'		If ListIndex <> 0 Then ListIndex -= 1
		'		EnsureItemVisible(ListIndex)
		'	Case Keys.Down
		'		mlLastSetListIndexCall = 0
		'		If ListIndex <> ListCount - 1 Then ListIndex += 1
		'		EnsureItemVisible(ListIndex)
		'End Select
	End Sub

    Private mlLastMouseDown As Int32 = 0
	Private Sub UITreeView_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseDown
		MyBase.HasFocus = True
		If goUILib.FocusedControl Is Nothing = False Then goUILib.FocusedControl.HasFocus = False
		goUILib.FocusedControl = Me

		Dim oLoc As System.Drawing.Point = GetAbsolutePosition()
		Dim lY As Int32 = lMouseY - oLoc.Y - 5

		'divide by our height to determine the item
		lY \= mlListItemHeight

		'Now, get our first visible node
		Dim oNode As UITreeViewItem = GetFirstVisibleNode()
		If oNode Is Nothing = False Then
			While lY <> 0
				lY -= 1
				oNode = TraverseNextNode(oNode)
				If oNode Is Nothing Then Return
			End While
			'ok, we are here...
			If oNode Is Nothing = False Then
				'ok, now, determine our left
				Dim lX As Int32 = lMouseX - GetAbsolutePosition.X - 5
				Dim lOffset As Int32 = oNode.GetLeftOffset
				If lX < lOffset Then
					If oNode.oFirstChild Is Nothing = False Then
                        oNode.bExpanded = Not oNode.bExpanded
                        If oNode.bExpanded = True Then RaiseEvent NodeExpanded(oNode) Else RaiseEvent NodeUnexpanded(oNode)
						Me.IsDirty = True
						Return
                    End If
                ElseIf oNode.bClickIcon = True AndAlso oNode.bRenderIcon = True AndAlso lX < lOffset + oNode.IconRect.Width Then
                    RaiseEvent NodeIconClicked(oNode)
				Else
                    If glCurrentCycle - mlLastMouseDown < 15 AndAlso oSelectedNode Is Nothing = False AndAlso oSelectedNode.lTreeViewIndex = oNode.lTreeViewIndex Then
                        oSelectedNode = oNode
                        RaiseEvent NodeDoubleClicked()
                    Else
                        oSelectedNode = oNode

                    End If
                    mlLastMouseDown = glCurrentCycle
				End If
			End If
		End If

	End Sub

	Private Sub UITreeView_OnRender() Handles MyBase.OnRender
		Dim oLoc As System.Drawing.Point = GetAbsolutePosition()

		Dim lRenderNodeCnt As Int32 = GetTotalRenderableNodeCount()
		Dim lItemCnt As Int32 = (Me.Height - 10) \ mlListItemHeight
		If lRenderNodeCnt > lItemCnt Then
			If mvscrVert.Enabled = False Then mvscrVert.Enabled = True
			If mvscrVert.MaxValue <> (lRenderNodeCnt - lItemCnt) Then mvscrVert.MaxValue = (lRenderNodeCnt - lItemCnt)
			If mvscrVert.Value > mvscrVert.MaxValue Then mvscrVert.Value = mvscrVert.MaxValue
		Else
			If mvscrVert.Enabled = True Then mvscrVert.Enabled = False
			If mvscrVert.Value <> 0 Then mvscrVert.Value = 0
		End If


        For X As Int32 = 0 To mlAllNodeUB
            If moAllNodes(X) Is Nothing = False Then moAllNodes(X).lTop = Int32.MaxValue
        Next X

		'Do a color fill of White
		Dim oRect As Rectangle = New Rectangle(oLoc.X, oLoc.Y, Width, Height)
		moUILib.DoAlphaBlendColorFill(oRect, FillColor, oLoc)
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
        Using moBorderLine As New Line(MyBase.moUILib.oDevice)
            With Me.moUILib.oDevice
                moBorderLine.Width = 3 '2
                moBorderLine.Antialias = True
                moBorderLine.Begin()
                moBorderLine.Draw(moBorderLineVerts, BorderColor)
                moBorderLine.End()
            End With
        End Using
		'End of Border Drawing

        Dim bHasArrows As Boolean = False
        Dim bRenderIcons As Boolean = oIconTexture Is Nothing = False
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return
		Using oFont As Font = New Font(MyBase.moUILib.oDevice, moSysFont)
			Using oBoldFont As Font = New Font(MyBase.moUILib.oDevice, moSysFontBold)
                Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                    oTextSpr.Begin(SpriteFlags.AlphaBlend)

                    Try
                        'Ok, need to determine the first visible node... start with the first parent node
                        Dim lWidth As Int32 = Me.Width - mvscrVert.Width - 10
                        Dim oVisibleNode As UITreeViewItem = GetFirstVisibleNode()
                        If oVisibleNode Is Nothing = False Then
                            'Ok, now, go through our list rendering what we need to
                            For X As Int32 = 0 To lItemCnt - 1
                                'Ok, render our visible node

                                'On a line, we have the arrow (if this item has children)
                                'then, we have the text
                                Dim lOffsetLeft As Int32 = oVisibleNode.GetLeftOffset()
                                If bRenderIcons = True Then lOffsetLeft += 16
                                Dim rcDest As Rectangle = New Rectangle(lOffsetLeft + 5 + oLoc.X, oLoc.Y + 5 + (X * mlListItemHeight), lWidth - lOffsetLeft + 10, mlListItemHeight)

                                oVisibleNode.lTop = rcDest.Top

                                Dim clrItem As System.Drawing.Color = moForeColor
                                If oVisibleNode.bUseItemColor = True Then clrItem = oVisibleNode.clrItemColor

                                'Is this the selected node?
                                If Object.Equals(oVisibleNode, oSelectedNode) = True Then
                                    clrItem = moHighlightTextColor
                                    Try
                                        Dim oT As Surface = MyBase.moUILib.oDevice.GetRenderTarget(0)
                                        moUILib.oDevice.ColorFill(oT, rcDest, HighlightColor)
                                        oT.Dispose()
                                        oT = Nothing
                                    Catch
                                        'Do nothing
                                    End Try
                                ElseIf oVisibleNode.bUseFillColor = True Then
                                    Try
                                        Dim oT As Surface = MyBase.moUILib.oDevice.GetRenderTarget(0)
                                        Dim tmpClr As System.Drawing.Color = oVisibleNode.clrFillColor
                                        moUILib.oDevice.ColorFill(oT, rcDest, tmpClr)
                                        oT.Dispose()
                                        oT = Nothing
                                    Catch
                                        'Do nothing
                                    End Try
                                End If

                                'Now, draw our text...
                                If oVisibleNode.sItem <> "" Then
                                    If bRenderIcons = True Then rcDest.X += 2
                                    If oVisibleNode.bItemBold = True Then
                                        oBoldFont.DrawText(oTextSpr, oVisibleNode.sItem, rcDest, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, clrItem)
                                    Else
                                        oFont.DrawText(oTextSpr, oVisibleNode.sItem, rcDest, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, clrItem)
                                    End If
                                End If

                                If oVisibleNode.oFirstChild Is Nothing = False Then bHasArrows = True

                                ' do xp
                                'Try
                                '    If Me.ParentControl.ControlName = "frmCommand" AndAlso oVisibleNode Is Nothing = False AndAlso goCurrentEnvir Is Nothing = False AndAlso oVisibleNode.lItemData2 = ObjectType.eUnit Then
                                '        For Y As Int32 = 0 To goCurrentEnvir.lEntityUB
                                '            If goCurrentEnvir.lEntityIdx(Y) <> -1 AndAlso goCurrentEnvir.oEntity(Y).ObjTypeID = oVisibleNode.lItemData2 AndAlso goCurrentEnvir.oEntity(Y).ObjectID = oVisibleNode.lItemData Then
                                '                Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(Y)
                                '                'If oEntity.ObjectID = 2972566 Then Stop
                                '                'If oEntity.ObjectID = 2972566 Then Stop
                                '                'If oEntity.ObjectID = 3088802 Then Stop
                                '                Dim lXPRank As Int32 = Math.Abs((CInt(oEntity.Exp_Level) - 1) \ 25) '  CInt(Math.Floor(Math.Abs(.Exp_Level - 1) / 25))
                                '                If oEntity.Exp_Level > -1 Then
                                '                    Using oImgSprite As New Sprite(MyBase.moUILib.oDevice)
                                '                        oImgSprite.Begin(SpriteFlags.AlphaBlend)
                                '                        Dim llOffsetLeft As Int32 = oVisibleNode.GetLeftOffset() - 10
                                '                        Dim rccDest As Rectangle = New Rectangle(oLoc.X + 5, oLoc.Y + 9 + (X * mlListItemHeight), 8, 10)
                                '                        Dim rccSrc As Rectangle = New Rectangle(0, 0, 16, 16)
                                '                        Dim iOffsetX As Int32 = llOffsetLeft + 13
                                '                        Dim iOffsetY As Int32 = -1
                                '                        Select Case lXPRank
                                '                            Case 0  'Green
                                '                                'oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_EmptyDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top + iOffsetY), Color.White)
                                '                            Case 1  'Trained
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top + iOffsetY), Color.White)
                                '                            Case 2  'Experienced
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top - 4 + iOffsetY), Color.White)
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top + 4 + iOffsetY), Color.White)
                                '                            Case 3  'Adept
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top - 4 + iOffsetY), Color.White)
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 21 + iOffsetX, rccDest.Top + 4 + iOffsetY), Color.White)
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 13 + iOffsetX, rccDest.Top + 4 + iOffsetY), Color.White)
                                '                            Case 4  'Veteran
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top - 4 + iOffsetY), Color.White)
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Arrow), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top + 4 + iOffsetY), Color.White)
                                '                            Case 5  'Ace
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Arrow), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top - 4 + iOffsetY), Color.White)
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Arrow), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top + 4 + iOffsetY), Color.White)
                                '                            Case 6  'Top Ace
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Arrow), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top + 6 + iOffsetY), Color.White)
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Arrow), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top + iOffsetY), Color.White)
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Arrow), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top - 6 + iOffsetY), Color.White)
                                '                            Case 7  'Distinguished
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_SolidDot), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top - 4 + iOffsetY), Color.White)
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Bar), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top + 4 + iOffsetY), Color.White)
                                '                            Case 8  'Revered
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Bar), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top - 4 + iOffsetY), Color.White)
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Bar), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top + 4 + iOffsetY), Color.White)
                                '                            Case Else   'Elite
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Bar), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top + 6 + iOffsetY), Color.White)
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Bar), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top + iOffsetY), Color.White)
                                '                                oImgSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, grc_UI(elInterfaceRectangle.eXPRank_Bar), New Rectangle(0, 0, 16, 16), Point.Empty, 0, New Point(rccDest.Width - 17 + iOffsetX, rccDest.Top - 6 + iOffsetY), Color.White)
                                '                        End Select
                                '                        oImgSprite.End()
                                '                        oImgSprite.Dispose()
                                '                    End Using
                                '                End If
                                '                Exit For
                                '            End If
                                '        Next Y
                                '    End If
                                'Catch ex As Exception
                                'End Try
                                oVisibleNode = TraverseNextNode(oVisibleNode)
                                If oVisibleNode Is Nothing Then Exit For
                            Next X

                        End If
                    Catch
                    End Try

                    oTextSpr.End()
                    oTextSpr.Dispose()
                End Using
				oBoldFont.Dispose()
			End Using
			oFont.Dispose()
		End Using

        'Now, render our icons
        Dim lIconCnt As Int32 = 0
        If GFXEngine.gbPaused = False AndAlso bHasArrows = True Then
            Device.IsUsingEventHandlers = False
            Using oSprite As New Sprite(MyBase.moUILib.oDevice)
                oSprite.Begin(SpriteFlags.AlphaBlend)

                Dim lWidth As Int32 = Me.Width - mvscrVert.Width - 10

                Dim oVisibleNode As UITreeViewItem = GetFirstVisibleNode()
                If oVisibleNode Is Nothing = False Then
                    'Ok, now, go through our list rendering what we need to
                    For X As Int32 = 0 To lItemCnt - 1

                        If oVisibleNode.oFirstChild Is Nothing = False Then
                            Dim lOffsetLeft As Int32 = oVisibleNode.GetLeftOffset() - 10
                            Dim rcDest As Rectangle
                            Dim rcSrc As Rectangle

                            If oVisibleNode.bExpanded = True Then
                                rcSrc = grc_UI(elInterfaceRectangle.eDownExpander)
                                rcDest = New Rectangle(oLoc.X + 5 + lOffsetLeft, oLoc.Y + 9 + (X * mlListItemHeight), 10, 10)
                            Else
                                rcSrc = grc_UI(elInterfaceRectangle.eRightExpander)
                                rcDest = New Rectangle(oLoc.X + 5 + lOffsetLeft, oLoc.Y + 9 + (X * mlListItemHeight), 8, 10)
                            End If

                            oSprite.Draw2D(MyBase.moUILib.oInterfaceTexture, rcSrc, rcDest, Point.Empty, 0, rcDest.Location, muSettings.InterfaceBorderColor)
                        End If
                        If oVisibleNode.bRenderIcon = True Then lIconCnt += 1
                        oVisibleNode = TraverseNextNode(oVisibleNode)
                        If oVisibleNode Is Nothing Then Exit For
                    Next X
                End If

                oSprite.End()
                oSprite.Dispose()
            End Using
            Device.IsUsingEventHandlers = True
        Else
            Dim oVisibleNode As UITreeViewItem = GetFirstVisibleNode()
            If oVisibleNode Is Nothing = False Then
                'Ok, now, go through our list rendering what we need to
                For X As Int32 = 0 To lItemCnt - 1
                    If oVisibleNode.bRenderIcon = True Then lIconCnt += 1
                    oVisibleNode = TraverseNextNode(oVisibleNode)
                    If oVisibleNode Is Nothing Then Exit For
                Next X
            End If
        End If

        If GFXEngine.gbPaused = False AndAlso bRenderIcons = True AndAlso lIconCnt > 0 Then
            Dim oBPS As New BPSprite()
            oBPS.BeginRender(0, oIconTexture, GFXEngine.moDevice)

            'Dim lWidth As Int32 = Me.Width - mvscrVert.Width - 10

            Dim oVisibleNode As UITreeViewItem = GetFirstVisibleNode()
            If oVisibleNode Is Nothing = False Then
                'Ok, now, go through our list rendering what we need to
                For X As Int32 = 0 To lItemCnt - 1
                    If oVisibleNode.bRenderIcon = True AndAlso bRenderIcons = True Then
                        Dim lOffsetLeft As Int32 = oVisibleNode.GetLeftOffset() '- 10
                        Dim rcDest As Rectangle = New Rectangle(oLoc.X + lOffsetLeft + 6, oLoc.Y + (X * mlListItemHeight) + 6, 15, 15)
                        'oSprite.Draw2D(oIconTexture, oVisibleNode.IconRect, rcDest, Point.Empty, 0, rcDest.Location, oVisibleNode.IconColor)
                        oBPS.Draw2D(oVisibleNode.IconRect, rcDest, oVisibleNode.IconColor)
                    End If

                    oVisibleNode = TraverseNextNode(oVisibleNode)
                    If oVisibleNode Is Nothing Then Exit For
                Next X
            End If
            GFXEngine.moDevice.RenderState.AlphaBlendEnable = False
            oBPS.EndRenderNoStateChange()
            GFXEngine.moDevice.RenderState.AlphaBlendEnable = True
            oBPS.DisposeMe()
        End If

	End Sub

	Private Sub UITreeView_OnResize() Handles Me.OnResize
		If Width = 0 OrElse Height = 0 Then Return
		If mvscrVert Is Nothing = False Then
			With mvscrVert
				.Width = 24
				.Left = Me.Width - .Width
				.Height = Me.Height - 2
				.Top = 1
				.Visible = True
			End With
		End If
	End Sub

	Public Function GetFirstVisibleNode() As UITreeViewItem
		Dim oResult As UITreeViewItem = Nothing
		If mvscrVert Is Nothing = False Then
			If oRootNode Is Nothing = False Then
				Dim lCnt As Int32 = 0
				If mvscrVert.Enabled = True Then lCnt = mvscrVert.Value
				Dim oCurrNode As UITreeViewItem = oRootNode
				While lCnt <> 0
					'Is current node expanded?
					lCnt -= 1

					Dim oNextNode As UITreeViewItem = TraverseNextNode(oCurrNode)
					If oNextNode Is Nothing Then Return oCurrNode
					oCurrNode = oNextNode

				End While
				oResult = oCurrNode
			End If
		End If
		Return oResult
	End Function

	Public Shared Function TraverseNextNode(ByRef oCurrNode As UITreeViewItem) As UITreeViewItem
		If oCurrNode.bExpanded = True AndAlso oCurrNode.oFirstChild Is Nothing = False Then
			'  Yes: Current node = Current Node's First Child
			Return oCurrNode.oFirstChild
		Else
			'  No: Is Current Node's Next Sibling Nothing?
			If oCurrNode.oNextSibling Is Nothing = False Then
				'	 No: Current Node = Current Node's Next Sibling
				Return oCurrNode.oNextSibling
			Else
				'	 Yes: Is Current Node's Parent Node Nothing?
				If oCurrNode.oParentNode Is Nothing = False Then
					'		No: Is Current Node's Parent Node's Next Sibling Nothing?
					If oCurrNode.oParentNode.oNextSibling Is Nothing = False Then
						'			No: Current Node = Parent Node's Next Sibling
						Return oCurrNode.oParentNode.oNextSibling
					Else
						Return FindNextParentSibling(oCurrNode)
						'ok, check if the parent's parent is nothing
						'If oCurrNode.oParentNode.oParentNode Is Nothing = False Then
						'	If oCurrNode.oParentNode.oParentNode.oNextSibling Is Nothing = False Then
						'		Return oCurrNode.oParentNode.oParentNode.oNextSibling
						'	Else : Return Nothing
						'	End If
						'Else : Return Nothing
						'End If
					End If
				Else : Return Nothing
				End If
			End If
		End If
		Return Nothing
    End Function

    Public Shared Function TraverseNextNode_NoExpand(ByRef oCurrNode As UITreeViewItem) As UITreeViewItem
        If oCurrNode.oFirstChild Is Nothing = False Then
            '  Yes: Current node = Current Node's First Child
            Return oCurrNode.oFirstChild
        Else
            '  No: Is Current Node's Next Sibling Nothing?
            If oCurrNode.oNextSibling Is Nothing = False Then
                '	 No: Current Node = Current Node's Next Sibling
                Return oCurrNode.oNextSibling
            Else
                '	 Yes: Is Current Node's Parent Node Nothing?
                If oCurrNode.oParentNode Is Nothing = False Then
                    '		No: Is Current Node's Parent Node's Next Sibling Nothing?
                    If oCurrNode.oParentNode.oNextSibling Is Nothing = False Then
                        '			No: Current Node = Parent Node's Next Sibling
                        Return oCurrNode.oParentNode.oNextSibling
                    Else
                        Return FindNextParentSibling(oCurrNode)
                        'ok, check if the parent's parent is nothing
                        'If oCurrNode.oParentNode.oParentNode Is Nothing = False Then
                        '	If oCurrNode.oParentNode.oParentNode.oNextSibling Is Nothing = False Then
                        '		Return oCurrNode.oParentNode.oParentNode.oNextSibling
                        '	Else : Return Nothing
                        '	End If
                        'Else : Return Nothing
                        'End If
                    End If
                Else : Return Nothing
                End If
            End If
        End If
        Return Nothing
    End Function

	Public Shared Function FindNextParentSibling(ByRef oNode As UITreeViewItem) As UITreeViewItem
		'my parent node's next sibling is nothing, so we need to go to the parent node's parent and check it...
		Dim oNewParent As UITreeViewItem = oNode.oParentNode
		If oNewParent Is Nothing Then Return Nothing
		If oNewParent.oNextSibling Is Nothing Then
			Return FindNextParentSibling(oNewParent)
		Else : Return oNewParent.oNextSibling
		End If
	End Function

	Private Function GetTotalRenderableNodeCount() As Int32
		Dim lCnt As Int32 = 0
		Dim oNode As UITreeViewItem = oRootNode
		While oNode Is Nothing = False
			lCnt += 1
			oNode = TraverseNextNode(oNode)
		End While
		Return lCnt
	End Function

	Protected Overrides Sub Finalize()
		mvscrVert = Nothing
		mhscrHoriz = Nothing
        'If moBorderLine Is Nothing = False Then moBorderLine.Dispose()
        'moBorderLine = Nothing

		If moSysFont Is Nothing = False Then moSysFont.Dispose()
		moSysFont = Nothing
		If moSysFontBold Is Nothing = False Then moSysFontBold.Dispose()
		moSysFontBold = Nothing


		MyBase.Finalize()
	End Sub

	Public Function GetNodeByItemData1(ByVal lItemData As Int32) As UITreeViewItem
		For X As Int32 = 0 To mlAllNodeUB
			If moAllNodes(X) Is Nothing = False Then
				With moAllNodes(X)
					If .lItemData = lItemData Then Return moAllNodes(X)
				End With
			End If
		Next X
		Return Nothing
	End Function
	Public Function GetNodeByItemData2(ByVal lItemData As Int32, ByVal lItemData2 As Int32) As UITreeViewItem
		For X As Int32 = 0 To mlAllNodeUB
			If moAllNodes(X) Is Nothing = False Then
                With moAllNodes(X)
                    If .lItemData = lItemData AndAlso .lItemData2 = lItemData2 Then Return moAllNodes(X)
                End With
			End If
		Next X
		Return Nothing
	End Function
	Public Function GetNodeByItemData3(ByVal lItemData As Int32, ByVal lItemData2 As Int32, ByVal lItemData3 As Int32) As UITreeViewItem
		For X As Int32 = 0 To mlAllNodeUB
			If moAllNodes(X) Is Nothing = False Then
				With moAllNodes(X)
					If .lItemData = lItemData AndAlso .lItemData2 = lItemData2 AndAlso .lItemData3 = lItemData3 Then Return moAllNodes(X)
				End With
			End If
		Next X
		Return Nothing
	End Function
	Public Function GetNodeByItemText(ByVal sText As String, ByVal bCaseSensative As Boolean, ByVal bExactMatch As Boolean, ByVal lStartNodeIndex As Int32) As UITreeViewItem
		If bCaseSensative = False Then sText = sText.ToUpper
		lStartNodeIndex += 1
		If lStartNodeIndex < 0 Then lStartNodeIndex = 0
		If lStartNodeIndex > mlAllNodeUB Then lStartNodeIndex = Math.Min(0, mlAllNodeUB)
		For X As Int32 = lStartNodeIndex To mlAllNodeUB
			If moAllNodes(X) Is Nothing = False Then
				With moAllNodes(X)
					If bCaseSensative = True Then
						If bExactMatch = True Then
							If .sItem = sText Then Return moAllNodes(X)
						Else
							If .sItem.Contains(sText) = True Then Return moAllNodes(X)
						End If
					Else
						If bExactMatch = True Then
							If .sItem.ToUpper = sText Then Return moAllNodes(X)
						Else
							If .sItem.ToUpper.Contains(sText) = True Then Return moAllNodes(X)
						End If

					End If

				End With
			End If
		Next X
		Return Nothing
	End Function

	Public Function AddNode(ByVal sItem As String, ByVal lItemData As Int32, ByVal lItemData2 As Int32, ByVal lItemData3 As Int32, ByRef oParent As UITreeViewItem, ByRef oBefore As UITreeViewItem) As UITreeViewItem
        Me.IsDirty = True
        Dim oNode As UITreeViewItem = New UITreeViewItem
		With oNode
			.oParentNode = oParent
			.sItem = sItem
			.lItemData = lItemData
			.lItemData2 = lItemData2
			.lItemData3 = lItemData3

			'Now, check, verify and associate...
			If .oParentNode Is Nothing = False Then
				'ok, my parent is not nothing...
				If .oParentNode.oFirstChild Is Nothing = False Then
					If oBefore Is Nothing = False AndAlso Object.Equals(oBefore.oParentNode, .oParentNode) = True Then
						If oBefore.oPrevSibling Is Nothing = False Then
							oBefore.oPrevSibling.oNextSibling = oNode
							oNode.oPrevSibling = oBefore.oPrevSibling
						End If
						oBefore.oPrevSibling = oNode
						oNode.oNextSibling = oBefore
					Else
						'place at end
						Dim oTmpNode As UITreeViewItem = .oParentNode.oFirstChild
						If oTmpNode Is Nothing Then
							.oParentNode.oFirstChild = oNode
						Else
							While oTmpNode.oNextSibling Is Nothing = False
								oTmpNode = oTmpNode.oNextSibling
							End While
							If oTmpNode Is Nothing = False Then
								oTmpNode.oNextSibling = oNode
								oNode.oPrevSibling = oTmpNode
							End If
						End If
						
					End If
				Else
					.oParentNode.oFirstChild = oNode
				End If
			Else
				'Ok, my parent is nothing, check the root
				If oRootNode Is Nothing = False Then
					'Check the before and after, are their parent's nothing?

					If oBefore Is Nothing = False AndAlso Object.Equals(oBefore, oRootNode) = True Then
						oNode.oNextSibling = oRootNode
						oRootNode.oPrevSibling = oNode
						oRootNode = oNode
					ElseIf oBefore Is Nothing = False AndAlso oBefore.oParentNode Is Nothing AndAlso oRootNode.oNextSibling Is Nothing = False Then
						'ok, need to put me before this node...
						If oBefore.oPrevSibling Is Nothing = False Then
							oBefore.oPrevSibling.oNextSibling = oNode
							.oPrevSibling = oBefore.oPrevSibling
						End If
						oBefore.oPrevSibling = oNode
						.oNextSibling = oBefore
					Else
						'ok, put me at the end
						Dim oTmpNode As UITreeViewItem = oRootNode.oNextSibling
						If oTmpNode Is Nothing Then
							oRootNode.oNextSibling = oNode
							oNode.oPrevSibling = oRootNode
						Else
							While oTmpNode.oNextSibling Is Nothing = False
								oTmpNode = oTmpNode.oNextSibling
							End While
							If oTmpNode Is Nothing = False Then
								oTmpNode.oNextSibling = oNode
								oNode.oPrevSibling = oTmpNode
							End If
						End If
					End If
				Else : oRootNode = oNode
				End If
			End If
		End With

		mlAllNodeUB += 1
		ReDim Preserve moAllNodes(mlAllNodeUB)
		oNode.lTreeViewIndex = mlAllNodeUB
		moAllNodes(mlAllNodeUB) = oNode

		Return oNode
	End Function

	Public Sub RemoveNode(ByRef oNode As UITreeViewItem)
		If Object.Equals(oNode, oRootNode) = True Then
			If oNode.oNextSibling Is Nothing = False Then
				oRootNode = oNode.oNextSibling
				oRootNode.oPrevSibling = Nothing
			Else : oRootNode = Nothing
			End If
		Else
			If oNode.oParentNode Is Nothing = False Then
				'ok, I have a parent, am I the parent's first node?
				If Object.Equals(oNode.oParentNode.oFirstChild, oNode) = True Then
					If oNode.oNextSibling Is Nothing = False Then
						oNode.oParentNode.oFirstChild = oNode.oNextSibling
						oNode.oParentNode.oFirstChild.oPrevSibling = Nothing
					Else : oNode.oParentNode.oFirstChild = Nothing
					End If
				End If
			End If
			If oNode.oPrevSibling Is Nothing = False Then
				oNode.oPrevSibling.oNextSibling = oNode.oNextSibling
			End If
			If oNode.oNextSibling Is Nothing = False Then
				oNode.oNextSibling.oPrevSibling = oNode.oPrevSibling
			End If
		End If

		'Now, remove me and any of my children from the array
		moAllNodes(oNode.lTreeViewIndex) = Nothing

		Dim oTmpNode As UITreeViewItem = oNode.oFirstChild
		While oTmpNode Is Nothing = False
			moAllNodes(oTmpNode.lTreeViewIndex) = Nothing
			oTmpNode = oTmpNode.oNextSibling
		End While

		'Now, shift our nodes
		For X As Int32 = 0 To mlAllNodeUB
			If moAllNodes(X) Is Nothing Then
				'ok, got a gap...
				For Y As Int32 = X + 1 To mlAllNodeUB
					If moAllNodes(Y) Is Nothing = False Then
						moAllNodes(X) = moAllNodes(Y)
						moAllNodes(Y) = Nothing
						moAllNodes(X).lTreeViewIndex = X
						Exit For
					End If
				Next Y
			End If
		Next X

		For X As Int32 = 0 To mlAllNodeUB
			If moAllNodes(X) Is Nothing Then
				mlAllNodeUB = X
				Exit For
			End If
		Next X
		Me.IsDirty = True

	End Sub

	Public Sub Clear()
		oRootNode = Nothing
		mlAllNodeUB = -1
		ReDim moAllNodes(-1)
	End Sub

	Public Function SingleScreenNodeRenderCnt() As Int32
		Return (Me.Height - 10) \ mlListItemHeight
	End Function

    Private Sub UITreeView_ResetInterfaceColors(ByVal yType As Byte, ByVal clrPrev As System.Drawing.Color) Handles Me.ResetInterfaceColors
        Select Case yType
            Case 1      'border
                If BorderColor = clrPrev Then BorderColor = muSettings.InterfaceBorderColor
                If ForeColor = clrPrev Then ForeColor = muSettings.InterfaceBorderColor
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

    Public Function GetNodeUnderMouse(ByVal lMouseX As Int32, ByVal lMouseY As Int32) As UITreeViewItem
        Dim oLoc As System.Drawing.Point = GetAbsolutePosition()
        Dim lY As Int32 = lMouseY - oLoc.Y - 5

        'divide by our height to determine the item
        lY \= mlListItemHeight

        'Now, get our first visible node
        Dim oNode As UITreeViewItem = GetFirstVisibleNode()
        If oNode Is Nothing = False Then
            While lY <> 0
                lY -= 1
                oNode = TraverseNextNode(oNode)
                If oNode Is Nothing Then Return Nothing
            End While
            'ok, we are here...
            If oNode Is Nothing = False Then
                'ok, now, determine our left
                Dim lX As Int32 = lMouseX - GetAbsolutePosition.X - 5
                Dim lOffset As Int32 = oNode.GetLeftOffset
                If lX < lOffset Then
                    If oNode.oFirstChild Is Nothing = False Then
                        Return oNode
                    End If
                Else
                    Return oNode
                End If
            End If
        End If
        Return Nothing
    End Function

    Public Property VScrollBarValue() As Int32
        Get
            If mvscrVert Is Nothing Then Return 0
            Return mvscrVert.Value
        End Get
        Set(ByVal value As Int32)
            If mvscrVert Is Nothing = False Then mvscrVert.Value = Math.Min(Math.Max(value, mvscrVert.MinValue), mvscrVert.MaxValue)
        End Set
    End Property

    Public Property HScrollBarValue() As Int32
        Get
            If mhscrHoriz Is Nothing Then Return 0
            Return mhscrHoriz.Value
        End Get
        Set(ByVal value As Int32)
            If mhscrHoriz Is Nothing = False Then mhscrHoriz.Value = Math.Min(Math.Max(value, mhscrHoriz.MinValue), mhscrHoriz.MaxValue)
        End Set
    End Property
End Class

''Interface created from Interface Builder
'Public Class TreeviewTest
'	Inherits UIWindow

'	Private WithEvents NewControl1 As UITreeView
'	Public Sub New(ByRef oUILib As UILib)
'		MyBase.New(oUILib)

'		' initial props
'		With Me
'			.ControlName = "asd"
'			.Left = 230
'			.Top = 277
'			.Width = 256
'			.Height = 256
'			.Enabled = True
'			.Visible = True
'			.BorderColor = muSettings.InterfaceBorderColor
'			.FillColor = muSettings.InterfaceFillColor
'			.FullScreen = True
'		End With

'		'NewControl1 initial props
'		NewControl1 = New UITreeView(oUILib)
'		With NewControl1
'			.ControlName = "NewControl1"
'			.Left = 5
'			.Top = 5
'			.Width = 244
'			.Height = 242
'			.Enabled = True
'			.Visible = True
'			.BorderColor = muSettings.InterfaceBorderColor
'			.FillColor = muSettings.InterfaceFillColor
'			.ForeColor = muSettings.InterfaceBorderColor
'			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'		End With
'		Me.AddChild(CType(NewControl1, UIControl))

'		Dim oRootNode As UITreeView.UITreeViewItem
'		oRootNode = NewControl1.AddNode("Root", 1, 0, 0, Nothing, Nothing)
'		NewControl1.AddNode("Root 1 Child 1", 1, 1, 0, oRootNode, Nothing)
'		Dim oRoot2 As UITreeView.UITreeViewItem = NewControl1.AddNode("Root 2", 2, 0, 0, Nothing, Nothing)
'		Dim oNode As UITreeView.UITreeViewItem = NewControl1.AddNode("Root 1 Child 2", 1, 2, 0, oRootNode, Nothing)
'		oNode = NewControl1.AddNode("Root 1 child 2 child 1", 1, 2, 1, oNode, Nothing)
'		NewControl1.AddNode("Root 1 child 2 child 1 child 1", 1, 2, 2, oNode, Nothing)

'		NewControl1.AddNode("Root 2 child 1", 2, 1, 0, oRoot2, Nothing)
'		oNode = NewControl1.AddNode("Root 2 child 2", 2, 2, 0, oRoot2, Nothing)
'		NewControl1.AddNode("Root 2 child 3 (before 2)", 2, 3, 0, oRoot2, oNode)

'		MyBase.moUILib.AddWindow(Me)
'	End Sub
'End Class