Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class UIComboBox
    Inherits UIControl

    Protected WithEvents moDropDownList As UIListBox
    Protected WithEvents moDisplayValue As UILabel
    Protected WithEvents moDropDownButton As UIButton

    Private mbListBoxDisplayed As Boolean = False
    Private mlTempHeight As Int32       'used for expand/collapse
    Private mlTempWidth As Int32        'used for expand/collapse
    Public l_ListBoxHeight As Int32 = 96
    Public l_ListBoxWidth As Int32 = 0

    'Private Shared moBorderLine As Line
    'Public Shared Sub ReleaseBorderLine()
    '    If moBorderLine Is Nothing = False Then moBorderLine.Dispose()
    '    moBorderLine = Nothing
    'End Sub
    Private moBorderLineVerts(4) As Vector2

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
    Private moFillColor As System.Drawing.Color = System.Drawing.Color.White
    Private moSysFont As System.Drawing.Font
    Private moForeColor As System.Drawing.Color = System.Drawing.Color.Black

    Public Event ItemSelected(ByVal lItemIndex As Int32)
    Public Event DropDownExpanded(ByVal sComboBoxName As String)

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'Device.IsUsingEventHandlers = False
        'moBorderLine = New Line(oUILib.oDevice)
        'Device.IsUsingEventHandlers = True

        moDropDownButton = New UIButton(oUILib)
        With moDropDownButton
            .Caption = ""
            .Width = 24
            .Left = Me.Width - .Width

            .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eDownArrow_Button_Disabled)
            .ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eDownArrow_Button_Normal)
            .ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eDownArrow_Button_Down)
        End With

        Me.AddChild(CType(moDropDownButton, UIControl))

        moDropDownList = New UIListBox(oUILib)
        moDropDownList.Visible = False
        moDropDownList.Width = Me.Width
        moDropDownList.Width = l_ListBoxWidth
        moDropDownList.BorderColor = System.Drawing.Color.Black
        moDropDownList.mbAcceptReprocessEvents = True       'we set this to true so that if the parent Combobox has their set to true, it will work
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

    End Sub

#Region " Pass-thru to the ListBox Control "
    Public Sub AddItem(ByVal sItem As String)
        moDropDownList.AddItem(sItem)
    End Sub
    Public Sub Clear()
        moDropDownList.Clear()
    End Sub
    Public Sub SortList(ByVal bSortNumerals As Boolean, ByVal bHasNone As Boolean)
        moDropDownList.SortList(bSortNumerals, bHasNone)
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
    Public Property ItemData2(ByVal lIndex As Int32) As Int32
        Get
            Return moDropDownList.ItemData2(lIndex)
        End Get
        Set(ByVal value As Int32)
            moDropDownList.ItemData2(lIndex) = value
        End Set
    End Property
    Public ReadOnly Property NewIndex() As Int32
        Get
            Return moDropDownList.NewIndex
        End Get
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
            If Value = -1 Then moDisplayValue.Caption = ""
        End Set
    End Property

    Public Property FillColor() As System.Drawing.Color
        Get
            Return moFillColor
        End Get
        Set(ByVal Value As System.Drawing.Color)
            'moFillColor = muSettings.InterfaceTextBoxFillColor
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

    Public Property bReadOnly() As Boolean
        Get
            Return Not moDropDownButton.Enabled
        End Get
        Set(ByVal value As Boolean)
            moDropDownButton.Enabled = Not value
        End Set
    End Property

    Public Property ForeColor() As System.Drawing.Color
        Get
            Return moForeColor
        End Get
        Set(ByVal Value As System.Drawing.Color)
            moForeColor = Value
            moDisplayValue.ForeColor = moForeColor
            moDropDownList.ForeColor = moForeColor
            IsDirty = True
        End Set
    End Property

    Public Sub SetFont(ByVal oFont As System.Drawing.Font)
        moSysFont = oFont

        moDisplayValue.SetFont(moSysFont)
        moDropDownList.SetFont(moSysFont)

        IsDirty = True
    End Sub

    Public Function GetFont() As System.Drawing.Font
        Return moSysFont
    End Function

	Private Sub moDropDownButton_Click(ByVal sName As String) Handles moDropDownButton.Click
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

		If NewTutorialManager.TutorialOn = True Then
			Dim sParmList() As String = {"-1", "-1"}
			If lIndex > -1 Then
				sParmList(0) = ItemData(lIndex).ToString
				sParmList(1) = ItemData2(lIndex).ToString
			End If
			If MyBase.moUILib.CommandAllowedWithParms(True, GetFullName(), sParmList, False) = False Then
				Return
			End If
		End If

        'set the label
        moDisplayValue.Caption = moDropDownList.List(lIndex)

        'raise the event
        RaiseEvent ItemSelected(lIndex)
    End Sub

    Private Sub UIComboBox_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.OnKeyDown
        If Me.IsExpanded = True Then
            Dim bFound As Boolean = False
            For X As Int32 = ListIndex + 1 To ListCount - 1
                If List(X).ToUpper.StartsWith(e.KeyCode.ToString) = True Then
                    ListIndex = X
                    ExpandListBox()
                    moDropDownList.EnsureItemVisible(X)
                    bFound = True
                    Exit For
                End If
            Next X
            If bFound = False AndAlso ListIndex + 1 <> 0 Then
                For X As Int32 = 0 To ListIndex
                    If List(X).ToUpper.StartsWith(e.KeyCode.ToString) = True Then
                        ListIndex = X
                        ExpandListBox()
                        moDropDownList.EnsureItemVisible(X)
                        bFound = True
                        Exit For
                    End If
                Next X
            End If
        End If
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

    'Private Sub UIComboBox_OnRender(ByRef oImgSprite As Sprite, ByRef oTextSprite As Sprite) Handles MyBase.OnRender
    Private Sub UIComboBox_OnRender() Handles MyBase.OnRender
        Dim oLoc As System.Drawing.Point = GetAbsolutePosition()

        'If moBorderLine Is Nothing = True OrElse moBorderLine.Disposed = True Then
        '    Device.IsUsingEventHandlers = False
        '    moBorderLine = New Line(MyBase.moUILib.oDevice)
        '    Device.IsUsingEventHandlers = True
        'End If

        'Do a color fill of White
        Dim oRect As Rectangle = New Rectangle(oLoc.X, oLoc.Y, Width, Height)
        'moUILib.DoAlphaBlendColorFill_EX(oRect, moFillColor, oLoc, oImgSprite)
        moUILib.DoAlphaBlendColorFill(oRect, moFillColor, oLoc)
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
        'Try
        '    With Me.moUILib.oDevice
        '        moBorderLine.Width = 3 '2
        '        moBorderLine.Antialias = True
        '        moBorderLine.Begin()
        '        moBorderLine.Draw(moBorderLineVerts, BorderColor)
        '        moBorderLine.End()
        '    End With
        'Catch
        'End Try 
        BPLine.DrawLine(3, True, moBorderLineVerts, BorderColor)

    End Sub

    Public Sub CollapseListBox()
        If mbListBoxDisplayed = True Then
            mbListBoxDisplayed = False
            Me.Height = mlTempHeight
            Me.Width = mlTempWidth
            moDropDownList.Visible = False
            MyBase.moUILib.CurrentComboBoxSelected = Nothing
            IsDirty = True

            If mbReverseExpanded = True Then
                Me.Top = mlOriginalTop
                moDisplayValue.Top = 0
                moDropDownButton.Top = 0
            End If
        End If
    End Sub
    Private mbReverseExpanded As Boolean = False
    Private mlOriginalTop As Int32 

    Private Sub ExpandListBox()
        If mbListBoxDisplayed = True Then Exit Sub
        mbListBoxDisplayed = True
        mlTempHeight = Me.Height
        mlTempWidth = Me.Width
        mlOriginalTop = Me.Top
        moDropDownList.Top = Me.Height + 1

        MyBase.moUILib.CurrentComboBoxSelected = Me

        moDropDownList.Height = l_ListBoxHeight  '?

        Me.Height = Me.Height + moDropDownList.Height

        mbReverseExpanded = False

        'Need to iterate all parent controls in case the drop-down is inside a frame.. inside a frame.. inside a form.. etc
        Dim oParent As UIControl = Me.ParentControl
        Dim lParentHeight As Int32 = 0
        While oParent Is Nothing = False
            lParentHeight += oParent.Height
            oParent = oParent.ParentControl
        End While

        'Only do this entire code section if I'm out of bounds.. else don't do the (this sucks) part
        If (lParentHeight > 0 AndAlso Me.Height + Me.Top > lParentHeight) Then
            If Me.Height + Me.Top > Me.ParentControl.Height Then
                mbReverseExpanded = True
                Me.Top = mlOriginalTop - l_ListBoxHeight
                moDropDownList.Top = 0
                moDisplayValue.Top = moDropDownList.Top + moDropDownList.Height
                moDropDownButton.Top = moDisplayValue.Top
            End If
        End If

        'Always force me to the end to ensure that i render last (placing me at the top)
        If Me.ParentControl.ChildrenUB > -1 Then
            Dim oPreviousEndControl As UIControl = Me.ParentControl.moChildren(Me.ParentControl.ChildrenUB)

            'this sucks, I have to search through the parent controls' child array to determine where I am in it
            For X As Int32 = 0 To Me.ParentControl.ChildrenUB
                If Object.Equals(Me.ParentControl.moChildren(X), Me) = True Then
                    Me.ParentControl.moChildren(X) = oPreviousEndControl
                    Me.ParentControl.moChildren(Me.ParentControl.ChildrenUB) = Me
                    Exit For
                End If
            Next X
        End If

        'If l_ListBoxWidth > 0 Then
        '    moDropDownList.Width = l_ListBoxWidth
        '    If l_ListBoxWidth > Me.Width Then Me.Width = Me.Width + (l_ListBoxWidth - Me.Width)
        'End If
        moDropDownList.Visible = True

        If ListIndex <> -1 Then
            moDropDownList.EnsureItemVisible(ListIndex)
        End If
        mbListBoxDisplayed = True
        IsDirty = True

        If MyBase.moUILib.FocusedControl Is Nothing = False Then MyBase.moUILib.FocusedControl.HasFocus = False
        Me.HasFocus = True
        MyBase.moUILib.FocusedControl = Me

        RaiseEvent DropDownExpanded(Me.ControlName)
    End Sub

    Public ReadOnly Property IsExpanded() As Boolean
        Get
            Return mbListBoxDisplayed
        End Get
    End Property

    Protected Overrides Sub Finalize()
        moDropDownList = Nothing
        moDisplayValue = Nothing
        moDropDownButton = Nothing
        mbListBoxDisplayed = False
        'If moBorderLine Is Nothing = False Then moBorderLine.Dispose()
        'moBorderLine = Nothing

        MyBase.Finalize()
    End Sub

    Public Function FindComboItemData(ByVal lVal As Int32) As Boolean
        For X As Int32 = 0 To ListCount - 1
			If ItemData(X) = lVal Then
				Dim bTutOn As Boolean = NewTutorialManager.TutorialOn
				NewTutorialManager.TutorialOn = False
				moDropDownList.ListIndex = X
				moDropDownList.Visible = False
				moDisplayValue.Caption = moDropDownList.List(X)
				NewTutorialManager.TutorialOn = bTutOn
				Return True
			End If
        Next X
        Return False
    End Function

    Public Shadows Property ToolTipText() As String
        Get
            Return MyBase.ToolTipText
        End Get
        Set(ByVal value As String)
            MyBase.ToolTipText = value
            For X As Int32 = 0 To MyBase.ChildrenUB
                MyBase.moChildren(X).ToolTipText = value
            Next X
        End Set
    End Property

    Private Sub UIComboBox_ResetInterfaceColors(ByVal yType As Byte, ByVal clrPrev As System.Drawing.Color) Handles Me.ResetInterfaceColors
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

    Private Sub UIComboBox_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        If moDropDownButton.Enabled = True Then
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
            moDropDownButton_Click(Nothing)
        End If
    End Sub

    Public Function ContainsItemData(ByVal lVal As Int32) As Boolean
        For X As Int32 = 0 To ListCount - 1
            If ItemData(X) = lVal Then
                Return True
            End If
        Next X
        Return False
    End Function
    Public Function ContainsItemData2(ByVal lVal As Int32, ByVal lVal2 As Int32) As Boolean
        For X As Int32 = 0 To ListCount - 1
            If ItemData(X) = lVal AndAlso ItemData2(X) = lVal2 Then
                Return True
            End If
        Next X
        Return False
    End Function 
End Class
