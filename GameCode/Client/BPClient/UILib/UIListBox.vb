Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class UIListBox
    Inherits UIControl

    Public Event ItemClick(ByVal lIndex As Int32)
    Public Event ItemDblClick(ByVal lIndex As Int32)
    Public Event HeaderRowClick(ByVal lX As Int32)
    Public Event ItemMouseOver(ByVal lIndex As Int32, ByVal lX As Int32, ByVal lY As Int32)

    Protected WithEvents moScroll As UIScrollBar

    Private msHeaderRow As String = ""
    Public Property sHeaderRow() As String
        Get
            Return msHeaderRow
        End Get
        Set(ByVal value As String)
            msHeaderRow = value
            UIListBox_OnResize()
            Me.IsDirty = True
        End Set
    End Property

    'Private moBorderLine As Line
    Private moBorderLineVerts(4) As Vector2

    Private mlListItemHeight As Int32 = 18
    Private mlListItemUB As Int32 = -1      'not the same as our real item list... this is for display purposes

    Public NewIndex As Int32 = -1

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
    Private moHighlightColor As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 0, 255, 255)
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

    Private moSysFont As System.Drawing.Font
    Private moSysFontBold As System.Drawing.Font

    Protected Class uListBoxItem
        Public sItem As String
        Public lItemData As Int32
		Public lItemData2 As Int32
		Public lItemData3 As Int32
        Public bItemBold As Boolean
        Public bItemLocked As Boolean

        Public bUseItemCustomColor As Boolean
        Public clrItemCustomColor As System.Drawing.Color

        Public bUseItemCustomBackColor As Boolean = False
        Public clrItemCustomBackColor As System.Drawing.Color
        Public clrItemCustomBackPercent As Int32 = -1

        Public rcIconSrcRect As Rectangle = Rectangle.Empty
        Public clrIconForeColor As System.Drawing.Color = Color.White
        Public bDoIconOffset As Boolean = False

        Public bUpdated As Boolean = False
    End Class

    'Protected msItems() As String
    Protected muItems() As uListBoxItem
    Protected mlItemUB As Int32 = -1
    'Protected mlItemData() As Int32
    'Protected mlItemData2() As Int32    'extended item data
    'Protected mbItemBold() As Boolean
    'Protected mbItemLocked() As Boolean     'if true, indicates that the item is not clickable

    'Protected myItemCustomColor() As Byte   'indicates that the item uses a custom color in the mcolItemColor array
    'Protected mcolItemCustomColor() As System.Drawing.Color

    Private mlListIndex As Int32 = -1

    Public oIconTexture As Texture          'pointer only

    Private mlLastSetListIndexCall As Int32
	Private mbRenderIcons As Boolean = False
    Public Property RenderIcons() As Boolean
        Get
            Return mbRenderIcons
        End Get
        Set(ByVal value As Boolean)
            If mbRenderIcons <> value Then
                mbRenderIcons = value
                Me.IsDirty = True
            End If
        End Set
    End Property 

    Public Sub AddItemBeforeIndex(ByVal sItem As String, ByVal lBeforeIndex As Int32, ByVal bBold As Boolean)
        If lBeforeIndex > mlItemUB Then
            AddItem(sItem, False)
        Else
            If lBeforeIndex < 0 Then lBeforeIndex = 0
            IncrementArraySize()
            'Shift our list down from the index
            For X As Int32 = lBeforeIndex + 1 To mlItemUB
                muItems(X) = muItems(X - 1)
                'msItems(X) = msItems(X - 1)
                'mlItemData(X) = mlItemData(X - 1)
                'mlItemData2(X) = mlItemData2(X - 1)
                'mbItemBold(X) = mbItemBold(X - 1)
                'mcolItemCustomColor(X) = mcolItemCustomColor(X - 1)
                'myItemCustomColor(X) = myItemCustomColor(X - 1)
                'mbItemLocked(X) = mbItemLocked(X - 1)
            Next X

            'Now, add our item there
            With muItems(lBeforeIndex)
                .sItem = sItem
                .lItemData = 0
				.lItemData2 = 0
				.lItemData3 = 0
                .bItemBold = bBold
                .bUseItemCustomColor = False
                .bItemLocked = False
                .rcIconSrcRect = Rectangle.Empty
                .clrIconForeColor = Color.White
            End With
            'msItems(lBeforeIndex) = sItem
            'mlItemData(lBeforeIndex) = 0
            'mlItemData2(lBeforeIndex) = 0
            'mbItemBold(lBeforeIndex) = bBold
            'myItemCustomColor(lBeforeIndex) = 0
            'mbItemLocked(lBeforeIndex) = False
            NewIndex = lBeforeIndex

            Dim lValue As Int32 = mlItemUB - mlListItemUB
            If lValue > -1 Then
                moScroll.MaxValue = lValue
                moScroll.Enabled = True
            Else
                moScroll.MaxValue = 1
                moScroll.Value = 0
                moScroll.Enabled = False
            End If
        End If

        Me.IsDirty = True
    End Sub

    Public Sub AddItemAfterIndex(ByVal sItem As String, ByVal lAfterIndex As Int32, ByVal bBold As Boolean)
        'Ok, this is just like before Index except that the index passed in is retained... so...
        AddItemBeforeIndex(sItem, lAfterIndex + 1, bBold)
    End Sub

    Private Sub IncrementArraySize()
        mlItemUB += 1
        If mlItemUB > -1 Then
            ReDim Preserve muItems(mlItemUB)
            muItems(mlItemUB) = New uListBoxItem()
        End If
        'ReDim Preserve msItems(mlItemUB)
        'ReDim Preserve mlItemData(mlItemUB)
        'ReDim Preserve mlItemData2(mlItemUB)
        'ReDim Preserve mbItemBold(mlItemUB)
        'ReDim Preserve myItemCustomColor(mlItemUB)
        'ReDim Preserve mcolItemCustomColor(mlItemUB)
        'ReDim Preserve mbItemLocked(mlItemUB)
    End Sub

    Public Sub AddItem(ByVal sItem As String, Optional ByVal bBold As Boolean = False)
        Dim X As Int32

        IncrementArraySize()

        'msItems(mlItemUB) = sItem
        'mlItemData(mlItemUB) = 0
        'mlItemData2(mlItemUB) = 0
        'mbItemBold(mlItemUB) = bBold
        'myItemCustomColor(mlItemUB) = 0
        'mbItemLocked(mlItemUB) = False
        With muItems(mlItemUB)
            .sItem = sItem
            .lItemData = 0
			.lItemData2 = 0
			.lItemData3 = 0
            .bItemBold = bBold
            .bUseItemCustomColor = False
            .bItemLocked = False
            .rcIconSrcRect = Rectangle.Empty
            .clrIconForeColor = Color.White
        End With
        NewIndex = mlItemUB

        X = mlItemUB - mlListItemUB
        If X > -1 Then
            moScroll.MaxValue = X
            moScroll.Enabled = True
        Else
            moScroll.MaxValue = 0
            moScroll.Value = 0
            moScroll.Enabled = False
        End If
        If (mlItemUB >= miListIndex AndAlso Me.ListIndex <> miListIndex) OrElse (mlItemUB >= miFirstVisible AndAlso miFirstVisible >= moScroll.MaxValue AndAlso moScroll.Value <> miFirstVisible) Then
            'Me.RestorePosition()
        End If
        Me.IsDirty = True
    End Sub

    Public Sub Clear()
        Me.SavePosition()
        Dim X As Int32
        NewIndex = -1
        mlItemUB = -1
        'ReDim msItems(-1)
        'ReDim mbItemBold(-1)
        'ReDim mlItemData(-1)
        'ReDim mlItemData2(-1)
        'ReDim myItemCustomColor(-1)
        'ReDim mcolItemCustomColor(-1)
        'ReDim mbItemLocked(-1)
        ReDim muItems(-1)

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
        Try
            If lIndex <= mlItemUB Then
                For X = lIndex To mlItemUB - 1
                    'shift from one up down
                    'msItems(X) = msItems(X + 1)
                    'mlItemData(X) = mlItemData(X + 1)
                    'mlItemData2(X) = mlItemData2(X + 1)
                    'mbItemBold(X) = mbItemBold(X + 1)
                    'mcolItemCustomColor(X) = mcolItemCustomColor(X + 1)
                    'myItemCustomColor(X) = myItemCustomColor(X + 1)
                    'mbItemLocked(X) = mbItemLocked(X + 1)
                    muItems(X) = muItems(X + 1)
                Next X

                mlItemUB -= 1
                'ReDim Preserve msItems(mlItemUB)
                'ReDim Preserve mlItemData(mlItemUB)
                'ReDim Preserve mlItemData2(mlItemUB)
                'ReDim Preserve mbItemBold(mlItemUB)
                'ReDim Preserve mcolItemCustomColor(mlItemUB)
                'ReDim Preserve myItemCustomColor(mlItemUB)
                'ReDim Preserve mbItemLocked(mlItemUB)
                ReDim Preserve muItems(mlItemUB)
            End If
        Catch
            Me.Clear()
        End Try

        IsDirty = True
    End Sub

    Public Property ItemBold(ByVal lIndex As Int32) As Boolean
        Get
            'If lIndex <= mlItemUB Then Return  mbItemBold(lIndex) Else Return False
            If lIndex <= mlItemUB AndAlso lIndex > -1 AndAlso muItems(lIndex) Is Nothing = False Then Return muItems(lIndex).bItemBold Else Return False
        End Get
        Set(ByVal value As Boolean)
            If lIndex <= mlItemUB Then
                'If value <> mbItemBold(lIndex) Then
                If value <> muItems(lIndex).bItemBold Then
                    'mbItemBold(lIndex) = value
                    muItems(lIndex).bItemBold = value
                    Me.IsDirty = True
                End If
            End If
        End Set
    End Property
    Public Property ItemData(ByVal lIndex As Int32) As Int32
        Get
            'If lIndex <= mlItemUB AndAlso lIndex > -1 Then Return mlItemData(lIndex)
            If muItems Is Nothing = False AndAlso lIndex <= muItems.GetUpperBound(0) AndAlso lIndex > -1 AndAlso muItems(lIndex) Is Nothing = False Then Return muItems(lIndex).lItemData
        End Get
        Set(ByVal Value As Int32)
            If lIndex <= mlItemUB Then
                'mlItemData(lIndex) = Value
                muItems(lIndex).lItemData = Value
            End If
        End Set
    End Property
    Public Property ItemData2(ByVal lIndex As Int32) As Int32
        Get 
			If muItems Is Nothing = False AndAlso lIndex <= muItems.GetUpperBound(0) AndAlso lIndex > -1 Then Return muItems(lIndex).lItemData2
		End Get
        Set(ByVal Value As Int32)
            If lIndex <= mlItemUB Then 
				muItems(lIndex).lItemData2 = Value
			End If
        End Set
	End Property
	Public Property ItemData3(ByVal lIndex As Int32) As Int32
		Get
			If muItems Is Nothing = False AndAlso lIndex <= muItems.GetUpperBound(0) AndAlso lIndex > -1 Then Return muItems(lIndex).lItemData3
		End Get
		Set(ByVal Value As Int32)
			If lIndex <= mlItemUB Then
				muItems(lIndex).lItemData3 = Value
			End If
		End Set
	End Property
    Public Property List(ByVal lIndex As Int32) As String
        Get
            'If lIndex <= mlItemUB AndAlso lIndex > -1 Then Return msItems(lIndex) Else Return ""
            If lIndex <= mlItemUB AndAlso lIndex > -1 Then Return muItems(lIndex).sItem Else Return ""
        End Get
        Set(ByVal Value As String)
            If lIndex <= mlItemUB Then
                'msItems(lIndex) = Value
                muItems(lIndex).sItem = Value
                Me.IsDirty = True
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
            Try
                If NewTutorialManager.TutorialOn = True Then
                    Dim sParmList() As String = {"-1", "-1", "-1"}
                    If Value > -1 Then
                        Try
                            sParmList(0) = ItemData(Value).ToString
                            sParmList(1) = ItemData2(Value).ToString
                            sParmList(2) = ItemData3(Value).ToString
                        Catch
                        End Try
                    End If
                    If MyBase.moUILib.CommandAllowedWithParms(True, GetFullName(), sParmList, False) = False Then
                        Return
                    End If
                End If

                Me.IsDirty = True
                mlListIndex = Value
                If Value <> -1 AndAlso Value <= mlItemUB AndAlso muItems(Value).bItemLocked = False Then
                    If glCurrentCycle - mlLastSetListIndexCall < 10 AndAlso glCurrentCycle <> mlLastSetListIndexCall Then
                        RaiseEvent ItemDblClick(mlListIndex)
                    Else : RaiseEvent ItemClick(mlListIndex)
                    End If
                    mlLastSetListIndexCall = glCurrentCycle
                Else : mlListIndex = -1
                End If
            Catch
            End Try
        End Set
    End Property
    Public Property ItemCustomColor(ByVal lIndex As Int32) As System.Drawing.Color
        Get
            If lIndex <= mlItemUB AndAlso lIndex > -1 Then
                If muItems(lIndex).bUseItemCustomColor = False Then Return moForeColor Else Return muItems(lIndex).clrItemCustomColor
            Else : Return moForeColor
            End If
        End Get
        Set(ByVal value As System.Drawing.Color)
            If lIndex <= mlItemUB AndAlso lIndex > -1 Then
                If value = moForeColor Then
                    muItems(lIndex).bUseItemCustomColor = False
                Else
                    muItems(lIndex).bUseItemCustomColor = True
                    muItems(lIndex).clrItemCustomColor = value
                End If
                Me.IsDirty = True
            End If
        End Set
    End Property
    Public Property ItemCustomBackColor(ByVal lIndex As Int32) As System.Drawing.Color
        Get
            If lIndex <= mlItemUB AndAlso lIndex > -1 Then
                If muItems(lIndex).bUseItemCustomBackColor = False Then Return moFillColor Else Return muItems(lIndex).clrItemCustomBackColor
            Else : Return moFillColor
            End If
        End Get
        Set(ByVal value As System.Drawing.Color)
            If lIndex <= mlItemUB AndAlso lIndex > -1 Then
                If value = moForeColor Then
                    muItems(lIndex).bUseItemCustomBackColor = False
                Else
                    muItems(lIndex).bUseItemCustomBackColor = True
                    muItems(lIndex).clrItemCustomBackColor = value
                End If
                Me.IsDirty = True
            End If
        End Set
    End Property
    Public Property ItemCustomBackPercent(ByVal lIndex As Int32) As Int32
        Get
            If lIndex <= mlItemUB AndAlso lIndex > -1 Then
                Return muItems(lIndex).clrItemCustomBackPercent
            Else : Return -1
            End If
        End Get
        Set(ByVal value As Int32)
            If lIndex <= mlItemUB AndAlso lIndex > -1 Then
                If value = -1 Then
                    muItems(lIndex).bUseItemCustomBackColor = False
                Else
                    muItems(lIndex).bUseItemCustomBackColor = True
                    muItems(lIndex).clrItemCustomBackPercent = value
                End If
                Me.IsDirty = True
            End If
        End Set
    End Property

    Public Property ItemLocked(ByVal lIndex As Int32) As Boolean
        Get
            If lIndex <= mlItemUB Then Return muItems(lIndex).bItemLocked Else Return False
        End Get
        Set(ByVal value As Boolean)
            If lIndex <= mlItemUB Then
                If value <> muItems(lIndex).bItemLocked Then
                    muItems(lIndex).bItemLocked = value
                    Me.IsDirty = True
                End If
            End If
        End Set
    End Property
    Public Property rcIconRectangle(ByVal lIndex As Int32) As Rectangle
        Get
            If lIndex <= mlItemUB Then Return muItems(lIndex).rcIconSrcRect Else Return Rectangle.Empty
        End Get
        Set(ByVal value As Rectangle)
            If lIndex <= mlItemUB Then
                If muItems(lIndex).rcIconSrcRect <> value Then
                    muItems(lIndex).rcIconSrcRect = value
                    Me.IsDirty = True
                End If
            End If
        End Set
    End Property
    Public Property IconForeColor(ByVal lIndex As Int32) As System.Drawing.Color
        Get
            If lIndex <= mlItemUB AndAlso lIndex > -1 Then
                Return muItems(lIndex).clrIconForeColor
            Else : Return moForeColor
            End If
        End Get
        Set(ByVal value As System.Drawing.Color)
            If lIndex <= mlItemUB AndAlso lIndex > -1 Then
                If value <> muItems(lIndex).clrIconForeColor Then
                    muItems(lIndex).clrIconForeColor = value
                End If
                Me.IsDirty = True
            End If
        End Set
    End Property
    Public Property ApplyIconOffset(ByVal lIndex As Int32) As Boolean
        Get
            If lIndex <= mlItemUB Then Return muItems(lIndex).bDoIconOffset Else Return False
        End Get
        Set(ByVal value As Boolean)
            If lIndex <= mlItemUB Then
                If value <> muItems(lIndex).bDoIconOffset Then
                    muItems(lIndex).bDoIconOffset = value
                    Me.IsDirty = True
                End If
            End If
        End Set
    End Property
    Public Property ForeColor() As System.Drawing.Color
        Get
            Return moForeColor
        End Get
        Set(ByVal Value As System.Drawing.Color)
            moForeColor = Value
            IsDirty = True
        End Set
    End Property
    Public Property ItemUpdated(ByVal lIndex As Int32) As Boolean
        Get
            If lIndex <= mlItemUB Then Return muItems(lIndex).bUpdated Else Return False
        End Get
        Set(ByVal value As Boolean)
            If lIndex <= mlItemUB Then muItems(lIndex).bUpdated = value
        End Set
    End Property
    Public Sub SetFont(ByVal oFont As System.Drawing.Font)
        'Dim X As Int32

        moSysFont = oFont
        moSysFontBold = New System.Drawing.Font(oFont, FontStyle.Bold)

        'For X = 0 To mlListItemUB
        '    moListItems(X).SetFont(moSysFont)
        'Next X
        UIListBox_OnResize()
        IsDirty = True

    End Sub

    Public Function GetFont() As System.Drawing.Font
        Return moSysFont
    End Function

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'Device.IsUsingEventHandlers = False
        'moBorderLine = New Line(oUILib.oDevice)
        'Device.IsUsingEventHandlers = True

        'Default font
        SetFont(New System.Drawing.Font("MS Sans Serif", 10))

		moScroll = New UIScrollBar(oUILib, True)
		moScroll.ControlName = "moScroll"
        moScroll.ReverseDirection = True
        moScroll.mbAcceptReprocessEvents = True
        Me.AddChild(CType(moScroll, UIControl))
    End Sub

    Private Sub UIListBox_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.OnKeyDown
        Select Case e.KeyCode
            Case Keys.Home
                ListIndex = 0
                EnsureItemVisible(0)
            Case Keys.End
                ListIndex = ListCount - 1
                EnsureItemVisible(ListIndex)
            Case Keys.Up
                mlLastSetListIndexCall = 0
                If ListIndex <> 0 Then ListIndex -= 1
                EnsureItemVisible(ListIndex)
            Case Keys.Down
                mlLastSetListIndexCall = 0
                If ListIndex <> ListCount - 1 Then ListIndex += 1
                EnsureItemVisible(ListIndex)
        End Select
    End Sub

    Private Sub UIListBox_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles MyBase.OnMouseDown
        MyBase.HasFocus = True
        If goUILib.FocusedControl Is Nothing = False Then goUILib.FocusedControl.HasFocus = False
        goUILib.FocusedControl = Me

        Dim oLoc As System.Drawing.Point = GetAbsolutePosition()
        Dim lY As Int32 = lMouseY - oLoc.Y - 5

        'Now, lY tells us our exact location within the listbox
        Dim lNewIndex As Int32 = (lY \ mlListItemHeight) + moScroll.Value
        If sHeaderRow <> "" Then lNewIndex -= 1
        ListIndex = lNewIndex

        If sHeaderRow <> "" AndAlso lNewIndex = -1 Then
            RaiseEvent HeaderRowClick(lMouseX - oLoc.X)
        End If

    End Sub

    Private Sub UIListBox_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        Dim oLoc As System.Drawing.Point = GetAbsolutePosition()
        Dim lY As Int32 = lMouseY - oLoc.Y - 5

        'Now, lY tells us our exact location within the listbox
        Dim lNewIndex As Int32 = (lY \ mlListItemHeight) + moScroll.Value
        If sHeaderRow <> "" Then lNewIndex -= 1

        If lNewIndex > -1 Then
            RaiseEvent ItemMouseOver(lNewIndex, lMouseX, lMouseY)
        End If
        
    End Sub

    'Private Sub UIListBox_OnRender(ByRef oImgSprite As Sprite, ByRef oTextSprite As Sprite) Handles MyBase.OnRender
    Private Sub UIListBox_OnRender() Handles MyBase.OnRender

        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return

        Dim oLoc As System.Drawing.Point = GetAbsolutePosition()

        'Do a color fill of White
        Dim oRect As Rectangle = New Rectangle(oLoc.X, oLoc.Y, Width, Height)
        'moUILib.DoAlphaBlendColorFill_EX(oRect, FillColor, oLoc, oImgSprite)
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

        'MyBase.moUILib.AddLineListItem(moBorderLineVerts, True, 3, BorderColor)
        'End of the border drawing

        Dim lSelectedListItemIndex As Int32 = -1

        'render the selected label
        If ListIndex <> -1 Then
            lSelectedListItemIndex = ListIndex - moScroll.Value
        End If

        Dim lIconOffsetLeft As Int32 = 0
        If mbRenderIcons = True Then
            lIconOffsetLeft = Math.Max(mlListItemHeight, 16)
        End If


        Using oFont As Font = New Font(MyBase.moUILib.oDevice, moSysFont)
            Using oBoldFont As Font = New Font(MyBase.moUILib.oDevice, moSysFontBold)
                Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                    Dim bBegun As Boolean = False

                    Try
                        oTextSpr.Begin(SpriteFlags.AlphaBlend)
                        bBegun = True

                        Dim lWidth As Int32 = Me.Width - moScroll.Width - 10

                        Dim lFirstX As Int32 = 0
                        If sHeaderRow <> "" Then
                            lFirstX = 1

                            Dim rcDest As Rectangle = New Rectangle(oLoc.X + 5, oLoc.Y + 5, lWidth, mlListItemHeight)
                            oBoldFont.DrawText(oTextSpr, sHeaderRow, rcDest, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, muSettings.InterfaceBorderColor)

                            Using oTmpLine As New Line(MyBase.moUILib.oDevice)
                                Dim uVecs(1) As Vector2
                                uVecs(0) = New Vector2(oLoc.X + 5, oLoc.Y + 5 + mlListItemHeight)
                                uVecs(1) = New Vector2(oLoc.X + 5 + lWidth, uVecs(0).Y)

                                With oTmpLine
                                    .Antialias = True
                                    .Begin()
                                    .Draw(uVecs, BorderColor)
                                    .End()
                                    .Dispose()
                                End With
                            End Using
                        End If

                        For X As Int32 = lFirstX To mlListItemUB + lFirstX
                            Dim lTrueIdx As Int32 = moScroll.Value + X - lFirstX
                            Dim lTrueX As Int32 = X - lFirstX

                            If lTrueIdx <= mlItemUB Then
                                'ok... figure out where we are drawing this...
                                Dim lOffsetLeft As Int32 = 0
                                If muItems(lTrueIdx).bDoIconOffset = True Then lOffsetLeft = lIconOffsetLeft

                                Dim rcDest As Rectangle = New Rectangle(oLoc.X + 5 + lOffsetLeft, oLoc.Y + 5 + (X * mlListItemHeight), lWidth, mlListItemHeight)
                                Dim clrItem As System.Drawing.Color = ItemCustomColor(lTrueIdx) 'moForeColor

                                'Is this item selected?
                                If ItemCustomBackPercent(lTrueIdx) <> -1 Then
                                    Try
                                        'rcDest.Width = CInt(rcDest.Width * (CDbl(ItemCustomBackPercent(lTrueIdx)) * 0.01))
                                        Dim rcFillDest As Rectangle = New Rectangle(oLoc.X + 5 + lOffsetLeft, oLoc.Y + 5 + (X * mlListItemHeight), lWidth, mlListItemHeight)


                                        rcFillDest.Width = CInt(rcFillDest.Width * (CDbl(ItemCustomBackPercent(lTrueIdx)) * 0.01)) + 5
                                        Dim lWidthTmp As Int32 = lWidth + 11
                                        Dim lWidthPct As Int32 = lWidthTmp - CInt(lWidthTmp * (CDbl(100 - ItemCustomBackPercent(lTrueIdx)) * 0.01))
                                        rcFillDest.X = lWidthPct

                                        rcFillDest.Width = lWidthTmp - lWidthPct
                                        Dim oT As Surface = MyBase.moUILib.oDevice.GetRenderTarget(0)
                                        moUILib.oDevice.ColorFill(oT, rcFillDest, ItemCustomBackColor(lTrueIdx))
                                        oT.Dispose()
                                        oT = Nothing
                                    Catch
                                        'Do nothing
                                    End Try
                                ElseIf lTrueX = lSelectedListItemIndex Then
                                    clrItem = moHighlightTextColor
                                    Try
                                        Dim oT As Surface = MyBase.moUILib.oDevice.GetRenderTarget(0)
                                        moUILib.oDevice.ColorFill(oT, rcDest, HighlightColor)
                                        oT.Dispose()
                                        oT = Nothing
                                    Catch
                                        'Do nothing
                                    End Try
                                ElseIf muItems(lTrueIdx).bItemLocked = True Then
                                    Try
                                        Dim oT As Surface = MyBase.moUILib.oDevice.GetRenderTarget(0)

                                        Dim tmpClr As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                                        moUILib.oDevice.ColorFill(oT, rcDest, tmpClr)
                                        oT.Dispose()
                                        oT = Nothing
                                    Catch
                                        'Do nothing
                                    End Try
                                End If

                                'Now, draw our text...
                                If muItems(lTrueIdx).sItem <> "" Then
                                    If muItems(lTrueIdx).bItemBold = True Then
                                        oBoldFont.DrawText(oTextSpr, muItems(lTrueIdx).sItem, rcDest, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, clrItem)
                                    Else
                                        oFont.DrawText(oTextSpr, muItems(lTrueIdx).sItem, rcDest, DrawTextFormat.Left Or DrawTextFormat.VerticalCenter, clrItem)
                                    End If
                                End If
                                End If
                        Next X
                    Catch
                    End Try

                    If bBegun = True Then oTextSpr.End()
                    oTextSpr.Dispose()
                End Using
                oBoldFont.Dispose()
            End Using
            oFont.Dispose()
        End Using

        'Now, render our icons
        If mbRenderIcons = True AndAlso oIconTexture Is Nothing = False AndAlso oIconTexture.Disposed = False Then
            Using oSprite As New Sprite(MyBase.moUILib.oDevice)
                Dim bBegun As Boolean = False

                Try
                    oSprite.Begin(SpriteFlags.AlphaBlend)
                    bBegun = True

                    Dim lWidth As Int32 = Me.Width - moScroll.Width - 10

                    Dim lFirstX As Int32 = 0
                    If sHeaderRow <> "" Then lFirstX = 1

                    For X As Int32 = lFirstX To mlListItemUB + lFirstX
                        Dim lTrueIdx As Int32 = moScroll.Value + X - lFirstX
                        Dim lTrueX As Int32 = X - lFirstX
                        If lTrueIdx <= mlItemUB AndAlso muItems(lTrueIdx).rcIconSrcRect <> Rectangle.Empty Then
                            'ok... figure out where we are drawing this...
                            Dim rcDest As Rectangle = New Rectangle(oLoc.X + 5, oLoc.Y + 5 + (X * mlListItemHeight), lIconOffsetLeft, lIconOffsetLeft)
                            If rcDest.Width > muItems(lTrueIdx).rcIconSrcRect.Width Then
                                Dim lDiff As Int32 = rcDest.Width - muItems(lTrueIdx).rcIconSrcRect.Width
                                rcDest.Width = muItems(lTrueIdx).rcIconSrcRect.Width
                                rcDest.X += (lDiff \ 2)
                            End If
                            If rcDest.Height > muItems(lTrueIdx).rcIconSrcRect.Height Then
                                Dim lDiff As Int32 = rcDest.Height - muItems(lTrueIdx).rcIconSrcRect.Height
                                rcDest.Height = muItems(lTrueIdx).rcIconSrcRect.Height
                                rcDest.Y += (lDiff \ 2)
                            End If
                            oSprite.Draw2D(oIconTexture, muItems(lTrueIdx).rcIconSrcRect, rcDest, Point.Empty, 0, rcDest.Location, muItems(lTrueIdx).clrIconForeColor)
                        End If

                    Next X
                Catch
                End Try
                If bBegun = True Then oSprite.End()
                oSprite.Dispose()
            End Using
        End If
    End Sub

    Private Sub UIListBox_OnResize() Handles MyBase.OnResize
        Dim X As Int32

        If Width = 0 Or Height = 0 Then Return

        'set up our scrollbar
        With moScroll
            .Width = 24
            .Left = Me.Width - .Width
            .Height = Me.Height - 6
            .Top = 4
        End With

        'remove our children from previous
        Me.RemoveAllChildren()
        Me.AddChild(CType(moScroll, UIControl))

        'now, set up our item labels

        Device.IsUsingEventHandlers = False
        Using moTmpFont As New Font(moUILib.oDevice, moSysFont)
            Dim rcTemp As Rectangle = moTmpFont.MeasureString(moUILib.Pen, "W", DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, Color.Black)
            mlListItemHeight = Math.Max(rcTemp.Height, 16)
            rcTemp = Nothing
            moTmpFont.Dispose()
        End Using
        Device.IsUsingEventHandlers = True

        Dim lModVal As Int32 = 1
        If sHeaderRow <> "" Then lModVal = 2
        mlListItemUB = (Me.Height - 5) \ mlListItemHeight - lModVal

        'finally, go back and set up the scrollbar's max
        X = mlItemUB - mlListItemUB
        If X > -1 Then moScroll.MaxValue = X Else moScroll.MaxValue = 1

    End Sub

    Public Sub EnsureItemVisible(ByVal lIndex As Int32)
        'Ok, ensure that the item is visible
        'easiest way to do this is to put the scroll value to it...
        If mlItemUB - lIndex > mlListItemUB Then
            moScroll.Value = lIndex - 1
        Else : moScroll.Value = mlItemUB - mlListItemUB
        End If
    End Sub

    Protected Overrides Sub Finalize()
        moScroll = Nothing
        'If moBorderLine Is Nothing = False Then moBorderLine.Dispose()
        'moBorderLine = Nothing
        mlListItemUB = -1

        If moSysFont Is Nothing = False Then moSysFont.Dispose()
        moSysFont = Nothing
        If moSysFontBold Is Nothing = False Then moSysFontBold.Dispose()
        moSysFontBold = Nothing


        MyBase.Finalize()
    End Sub

    Public Sub SetAllUpdates(ByVal bValue As Boolean)
        For X As Int32 = 0 To mlItemUB
            muItems(X).bUpdated = bValue
        Next X
    End Sub

    Private Shared Function GetRomanNumeralSortStr(ByVal sVal As String) As String
        If sVal.StartsWith("I") = False AndAlso sVal.StartsWith("X") = False AndAlso sVal.StartsWith("V") = False Then Return sVal

        Select Case sVal.ToUpper
            Case "I"
                Return "001"
            Case "II"
                Return "002"
            Case "III"
                Return "003"
            Case "IV"
                Return "004"
            Case "V"
                Return "005"
            Case "VI"
                Return "006"
            Case "VII"
                Return "007"
            Case "VIII"
                Return "008"
            Case "IX"
                Return "009"
            Case "X"
                Return "010"
            Case "XI"
                Return "011"
            Case "XII"
                Return "012"
            Case "XIII"
                Return "013"
            Case "XIV"
                Return "014"
            Case "XV"
                Return "015"
            Case "XVI"
                Return "016"
            Case "XVII"
                Return "017"
            Case "XVIII"
                Return "018"
            Case "XIX"
                Return "019"
            Case "XX"
                Return "020"
            Case "XXI"
                Return "021"
            Case "XXII"
                Return "022"
            Case "XXIII"
                Return "023"
            Case "XXIV"
                Return "024"
            Case "XXV"
                Return "025"
            Case "XXVI"
                Return "026"
            Case "XXVII"
                Return "027"
            Case "XXVIII"
                Return "028"
            Case "XXIX"
                Return "029"
            Case "XXX"
                Return "030"
            Case "XXXI"
                Return "031"
            Case Else
                Return sVal
        End Select
    End Function

    Public Sub SortList(ByVal bSortNumerals As Boolean, ByVal bHasNone As Boolean)
        Dim uTemp(mlItemUB) As uListBoxItem

        Dim lSorted(mlItemUB) As Int32
        Dim lSortedUB As Int32 = -1
        Dim sSortedVal(mlItemUB) As String

        Dim lStartIdx As Int32 = 0
        If bHasNone = True Then
            lSortedUB += 1
            ReDim Preserve lSorted(lSortedUB)
            ReDim Preserve sSortedVal(lSortedUB)

            lSorted(lSortedUB) = 0
            sSortedVal(lSortedUB) = "None"

            lStartIdx = 1
        End If

        For X As Int32 = lStartIdx To mlItemUB
            Dim sCurrent As String = muItems(X).sItem.ToUpper
            If bSortNumerals = True Then
                Dim sVals() As String = Split(sCurrent.Replace("(", "").Replace(")", ""), " ")
                If sVals Is Nothing = False AndAlso sVals.GetUpperBound(0) > -1 Then
                    Dim sTemp As String = ""
                    For lTmpVal As Int32 = 0 To sVals.GetUpperBound(0)
                        sCurrent &= GetRomanNumeralSortStr(sVals(lTmpVal))
                    Next lTmpVal
                End If
            End If

            Dim lIdx As Int32 = -1
            For Y As Int32 = lStartIdx To lSortedUB
                If sSortedVal(Y) > sCurrent Then
                    lIdx = Y
                    Exit For
                End If
            Next Y
            lSortedUB += 1
            If lSortedUB > lSorted.GetUpperBound(0) Then
                ReDim Preserve lSorted(lSortedUB)
                ReDim Preserve sSortedVal(lSortedUB)
            End If
            If lIdx = -1 Then
                lSorted(lSortedUB) = X
                sSortedVal(lSortedUB) = sCurrent
            Else
                For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                    lSorted(Y) = lSorted(Y - 1)
                    sSortedVal(Y) = sSortedVal(Y - 1)
                Next Y
                lSorted(lIdx) = X
                sSortedVal(lIdx) = sCurrent
            End If
        Next X

        If lSortedUB <> mlItemUB Then Return

        For X As Int32 = 0 To lSortedUB
            uTemp(X) = muItems(lSorted(X))
        Next X
        muItems = uTemp
    End Sub

    Private Sub UIListBox_ResetInterfaceColors(ByVal yType As Byte, ByVal clrPrev As System.Drawing.Color) Handles Me.ResetInterfaceColors
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

    Public Property ScrollBarValue() As Int32
        Get
            If moScroll Is Nothing Then Return 0
            Return moScroll.Value
        End Get
        Set(ByVal value As Int32)
            If moScroll Is Nothing = False Then moScroll.Value = Math.Min(Math.Max(value, moScroll.MinValue), moScroll.MaxValue)
        End Set
    End Property

    Shared miListIndex As Int32
    Shared miFirstVisible As Int32

    Public Sub SavePosition()
        miListIndex = Me.ListIndex
        miFirstVisible = moScroll.Value
    End Sub

    Public Sub RestorePosition()
        moScroll.Value = miFirstVisible
        Me.ListIndex = miListIndex
    End Sub
End Class
