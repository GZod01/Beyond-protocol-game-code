Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmRouteTemplate
    Inherits UIWindow

    Private Const ml_ROUTE_ITEMS As Int32 = 7

    Private vscScroll As UIScrollBar

    Private WithEvents btnClose As UIButton
    Private WithEvents btnClear As UIButton
    Private lblTemplates As UILabel
    Private WithEvents cboTemplates As UIComboBox
    Private WithEvents btnSaveTemplate As UIButton
    Private WithEvents btnDeleteTemplate As UIButton
    Private WithEvents btnAssignToSelected As UIButton
    Private WithEvents btnAssignAndRunToSelected As UIButton
    Private WithEvents btnCopyRoute As UIButton
    Private txtTemplateName As UITextBox
    Private lblTemplateName As UILabel

    Private lnDiv1 As UILine
    Private lnDiv2 As UILine
    Private lblTitle As UILabel
    Private lblDestTitle As UILabel
    Private lblLocTitle As UILabel
    Private lblMineralTitle As UILabel

    Private msCurrentSetMineralControl As String = ""

    Private msCurrentSetLocControl As String = ""

    Private moPopup As frmRouteConfig.fraSelectMin = Nothing
    Private mbInPopup As Boolean = False

    Public Shared oTemplates() As RouteTemplate
    Public Shared lTemplateUB As Int32 = -1
    Public Shared Sub RequestTemplates()
        Dim yMsg(1) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eGetRouteTemplates).CopyTo(yMsg, lPos) : lPos += 2
        If goUILib Is Nothing = False Then goUILib.SendMsgToPrimary(yMsg)
    End Sub
    Public Shared Sub HandleRequestTemplates(ByVal yData() As Byte)
        Dim lPos As Int32 = 2

        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim oTmp(lCnt - 1) As RouteTemplate

        For X As Int32 = 0 To lCnt - 1
            oTmp(X) = New RouteTemplate
            With oTmp(X)
                .lTemplateID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lPlayerID = glPlayerID
                .sTemplateName = GetStringFromBytes(yData, lPos, 20) : lPos += 20

                .lItemUB = System.BitConverter.ToInt32(yData, lPos) - 1 : lPos += 4
                ReDim .uItems(.lItemUB)
                For Y As Int32 = 0 To .lItemUB
                    With .uItems(Y)
                        .lOrderNum = Y
                        .FillFromMsg(yData, lPos)
                    End With
                Next Y
            End With
        Next X

        lTemplateUB = -1
        oTemplates = oTmp
        lTemplateUB = lCnt - 1
    End Sub

    Private moCurrentTemplate As RouteTemplate
    Private mbNeedsSaved As Boolean = False
    Private mbScrollOne As Boolean = False

#Region "  Route Config Lines  "
    Private lblDestination() As UILabel
    Private lblMineral() As UILabel
    Private btnRemoveItem() As UIButton
    Private btnSetMineral() As UIButton
    Private btnSetLocation() As UIButton
    Private lnDivs() As UILine


    Private Sub ConfigureItemAsUnused(ByVal lIndex As Int32)
        If lblDestination(lIndex).Visible = True Then lblDestination(lIndex).Visible = False
        If lblMineral(lIndex).Visible = True Then lblMineral(lIndex).Visible = False
        If btnRemoveItem(lIndex).Visible = True Then btnRemoveItem(lIndex).Visible = False
        If btnSetMineral(lIndex).Visible = True Then btnSetMineral(lIndex).Visible = False
        If btnSetLocation(lIndex).Visible = True Then btnSetLocation(lIndex).Visible = False
        If lblMineral(lIndex).Caption <> "Unload" Then lblMineral(lIndex).Caption = "Unload"
        If lblDestination(lIndex).Caption <> "Unassigned" Then lblDestination(lIndex).Caption = "Unassigned"
    End Sub
    Private Sub ConfigureItemAsAdd(ByVal lIndex As Int32)
        If lblDestination(lIndex).Visible = False Then lblDestination(lIndex).Visible = True
        If lblMineral(lIndex).Visible = False Then lblMineral(lIndex).Visible = True
        If btnRemoveItem(lIndex).Visible = True Then btnRemoveItem(lIndex).Visible = False
        If btnSetMineral(lIndex).Visible = True Then btnSetMineral(lIndex).Visible = False
        If btnSetLocation(lIndex).Visible = False Then btnSetLocation(lIndex).Visible = True
        If lblMineral(lIndex).Caption <> "" Then lblMineral(lIndex).Caption = ""
        If lblDestination(lIndex).Caption <> "" Then lblDestination(lIndex).Caption = ""
    End Sub
    Private Sub ConfigureItemAsUsed(ByVal lIndex As Int32)
        If lblDestination(lIndex).Visible = False Then lblDestination(lIndex).Visible = True
        If lblMineral(lIndex).Visible = False Then lblMineral(lIndex).Visible = True
        If btnRemoveItem(lIndex).Visible = False Then btnRemoveItem(lIndex).Visible = True
        If btnSetMineral(lIndex).Visible = False Then btnSetMineral(lIndex).Visible = True
        If btnSetLocation(lIndex).Visible = False Then btnSetLocation(lIndex).Visible = True
    End Sub

#End Region

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmRouteConfig initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eRouteConfig
            .ControlName = "frmRouteTemplate"

            .Left = 293
            .Top = 160
            .Width = 511
            .Height = 511
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 1
        End With

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 485
            .Top = 1
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "X"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 1
            .Top = 25
            .Width = 511
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = 1
            .Top = 55
            .Width = 511
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 0
            .Width = 175
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Route Template Management"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lblTemplates initial props
        lblTemplates = New UILabel(oUILib)
        With lblTemplates
            .ControlName = "lblTemplates"
            .Left = 30
            .Top = 28
            .Width = 150
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Templates:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTemplates, UIControl))

        lblTemplateName = New UILabel(oUILib)
        With lblTemplateName
            .ControlName = "lblTemplateName"
            .Left = 30
            .Top = 60
            .Width = 50
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Name:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTemplateName, UIControl))

        txtTemplateName = New UITextBox(oUILib)
        With txtTemplateName
            .ControlName = "txtTemplateName"
            .Left = lblTemplateName.Left + lblTemplateName.Width + 5
            .Top = 63
            .Width = 155
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 20
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        MyBase.AddChild(CType(txtTemplateName, UIControl))

        'lblDest initial props
        lblDestTitle = New UILabel(oUILib)
        With lblDestTitle
            .ControlName = "lblDest"
            .Left = 30
            .Top = 80
            .Width = 150
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Destination Name"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblDestTitle, UIControl))

        'lblMineral initial props
        lblMineralTitle = New UILabel(oUILib)
        With lblMineralTitle
            .ControlName = "lblMineral"
            .Left = 280
            .Top = 80 '25
            .Width = 150
            .Height = 24
            .Enabled = True
            .Visible = True
            '.Caption = "Pickup Mineral"
            .Caption = "Waypoint Action"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblMineralTitle, UIControl))


        'vscScroll initial props
        vscScroll = New UIScrollBar(oUILib, True)
        With vscScroll
            .ControlName = "vscScroll"
            .Left = 485
            .Top = (lblDestTitle.Top + lblDestTitle.Height + 2)
            .Width = 24
            .Height = 175
            .Enabled = True
            .Visible = True
            .Value = 0
            .MaxValue = 100
            .MinValue = 0
            .SmallChange = 1
            .LargeChange = 1
            .ReverseDirection = True
        End With
        Me.AddChild(CType(vscScroll, UIControl))

        'Now, our line items, we have 7 of them
        ReDim lblDestination(ml_ROUTE_ITEMS - 1)
        ReDim lblMineral(ml_ROUTE_ITEMS - 1)
        ReDim btnRemoveItem(ml_ROUTE_ITEMS - 1)
        ReDim btnSetMineral(ml_ROUTE_ITEMS - 1)
        ReDim btnSetLocation(ml_ROUTE_ITEMS - 1)
        ReDim lnDivs(ml_ROUTE_ITEMS)

        For X As Int32 = 0 To ml_ROUTE_ITEMS - 1
            Dim lItemTop As Int32 = (lblDestTitle.Top + lblDestTitle.Height + 2) + (X * 25)
            Dim lItemLeft As Int32 = 2

            lnDivs(X) = New UILine(oUILib)
            With lnDivs(X)
                .ControlName = "lnDiv1(" & X & ")"
                .Left = lItemLeft
                .Top = lItemTop - 1
                .Width = vscScroll.Left - lItemLeft
                .Height = 0
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
            End With
            Me.AddChild(CType(lnDivs(X), UIControl))

            btnSetLocation(X) = New UIButton(oUILib)
            With btnSetLocation(X)
                .ControlName = "btnSetLocation(" & X & ")"
                .Left = lItemLeft + 3
                .Top = lItemTop
                .Width = 24
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "+"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
                .ToolTipText = "Click to set the location of this item adding it to the list of the route."
            End With
            Me.AddChild(CType(btnSetLocation(X), UIControl))

            lblDestination(X) = New UILabel(oUILib)
            With lblDestination(X)
                .ControlName = "lblDestination(" & X & ")"
                .Left = btnSetLocation(X).Left + btnSetLocation(X).Width
                .Top = lItemTop
                .Width = 250
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Unassigned"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblDestination(X), UIControl))


            lblMineral(X) = New UILabel(oUILib)
            With lblMineral(X)
                .ControlName = "lblMineral(" & X & ")"
                .Left = lblDestination(X).Left + lblDestination(X).Width
                .Top = lItemTop
                .Width = 150
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Unload"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMineral(X), UIControl))

            btnRemoveItem(X) = New UIButton(oUILib)
            With btnRemoveItem(X)
                .ControlName = "btnRemoveItem(" & X & ")"
                .Left = vscScroll.Left - 25 ' btnSetMineral(X).Left + btnSetMineral(X).Width 
                .Top = lItemTop
                .Width = 24
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "X"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
                .ToolTipText = "Click to remove this item from the route list."
            End With
            Me.AddChild(CType(btnRemoveItem(X), UIControl))

            btnSetMineral(X) = New UIButton(oUILib)
            With btnSetMineral(X)
                .ControlName = "btnSetMineral(" & X & ")"
                .Left = btnRemoveItem(X).Left - 25
                .Top = lItemTop
                .Width = 24
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "M"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
                .ToolTipText = "Click to change the mineral to be picked up upon arrival at destination." & vbCrLf & "This requires that destination is a facility or unit."
            End With
            Me.AddChild(CType(btnSetMineral(X), UIControl))
        Next X

        lnDivs(ml_ROUTE_ITEMS) = New UILine(oUILib)
        With lnDivs(ml_ROUTE_ITEMS)
            .ControlName = "lnDiv1(" & ml_ROUTE_ITEMS & ")"
            .Left = 2
            .Top = (lblDestTitle.Top + lblDestTitle.Height + 2) + (ml_ROUTE_ITEMS * 25)
            .Width = vscScroll.Left - 2
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDivs(ml_ROUTE_ITEMS), UIControl))

        'btnClear initial props
        btnClear = New UIButton(oUILib)
        With btnClear
            .ControlName = "btnClear"
            .Left = Me.Width \ 2 - 50
            .Top = (lblDestTitle.Top + lblDestTitle.Height + 2) + 5 + (ml_ROUTE_ITEMS * 25)
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Clear"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClear, UIControl))

        'btnClear initial props
        btnSaveTemplate = New UIButton(oUILib)
        With btnSaveTemplate
            .ControlName = "btnSaveTemplate"
            .Left = Me.Width - 105
            .Top = btnClear.Top
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Save"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSaveTemplate, UIControl))

        btnDeleteTemplate = New UIButton(oUILib)
        With btnDeleteTemplate
            .ControlName = "btnDeleteTemplate"
            .Left = 5
            .Top = btnClear.Top
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Delete"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnDeleteTemplate, UIControl))

        btnCopyRoute = New UIButton(oUILib)
        With btnCopyRoute
            .ControlName = "btnCopyRoute"
            .Left = btnDeleteTemplate.Left + btnDeleteTemplate.Width + 0
            .Top = btnClear.Top
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Copy"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnCopyRoute, UIControl))

        btnAssignToSelected = New UIButton(oUILib)
        With btnAssignToSelected
            .ControlName = "btnAssignToSelected"
            .Left = 5
            .Top = btnClear.Top + btnClear.Height + 5
            .Width = 150
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Assign to Selected"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAssignToSelected, UIControl))

        btnAssignAndRunToSelected = New UIButton(oUILib)
        With btnAssignAndRunToSelected
            .ControlName = "btnAssignAndRunToSelected"
            .Left = Me.Width - 155
            .Top = btnClear.Top + btnClear.Height + 5
            .Width = 150
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Assign And Begin"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAssignAndRunToSelected, UIControl))

        Me.Height = btnAssignAndRunToSelected.Top + btnAssignAndRunToSelected.Height + 10

        cboTemplates = New UIComboBox(oUILib)
        With cboTemplates
            .ControlName = "cboTemplates"
            .Left = lblTemplates.Left + lblTemplates.Width + 5
            .Top = lblTemplates.Top + 3
            .Width = 173
            .Height = 18
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .l_ListBoxHeight = Me.Height - .Top - .Height - 4 '128
        End With
        Me.AddChild(CType(cboTemplates, UIControl))

        For X As Int32 = 0 To ml_ROUTE_ITEMS - 1
            AddHandler btnRemoveItem(X).Click, AddressOf HandleRemoveItemButtonClick
            AddHandler btnSetMineral(X).Click, AddressOf HandleSetMineralButtonClick
            AddHandler btnSetLocation(X).Click, AddressOf HandleSetLocationButtonClick
        Next X

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        RequestTemplates()
        moCurrentTemplate = New RouteTemplate
    End Sub

    Private Sub FillTemplates()
        Dim lID As Int32 = -1
        If cboTemplates.ListIndex > -1 Then lID = cboTemplates.ItemData(cboTemplates.ListIndex)

        cboTemplates.Clear()
        Try

            Dim lSorted() As Int32 = Nothing
            Dim lSortedUB As Int32 = -1
            For X As Int32 = 0 To lTemplateUB
                Dim lIdx As Int32 = -1

                Dim sName As String = oTemplates(X).sTemplateName

                For Y As Int32 = 0 To lSortedUB
                    Dim sOtherName As String = oTemplates(lSorted(Y)).sTemplateName
                    If sOtherName > sName Then
                        lIdx = Y
                        Exit For
                    End If
                Next Y
                lSortedUB += 1
                ReDim Preserve lSorted(lSortedUB)
                If lIdx = -1 Then
                    lSorted(lSortedUB) = X
                Else
                    For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                        lSorted(Y) = lSorted(Y - 1)
                    Next Y
                    lSorted(lIdx) = X
                End If
            Next X


            For X As Int32 = 0 To lTemplateUB
                If oTemplates(lSorted(X)) Is Nothing = False Then
                    cboTemplates.AddItem(oTemplates(lSorted(X)).sTemplateName)
                    cboTemplates.ItemData(cboTemplates.NewIndex) = oTemplates(lSorted(X)).lTemplateID
                End If
            Next X
        Catch
        End Try

        If lID > -1 Then cboTemplates.FindComboItemData(lID)

    End Sub

    Private Sub frmRouteTemplate_OnNewFrame() Handles Me.OnNewFrame
        If mbInPopup = True Then Return
        Try
            If cboTemplates Is Nothing = False Then
                If cboTemplates.ListCount <> lTemplateUB + 1 Then
                    FillTemplates()
                End If
            End If

            If moCurrentTemplate Is Nothing = False Then
                If moCurrentTemplate.lItemUB = -1 Then
                    ConfigureItemAsAdd(0)
                    For X As Int32 = 1 To ml_ROUTE_ITEMS - 1
                        ConfigureItemAsUnused(X)
                    Next X
                Else
                    'ok, the person has items...
                    Dim lScrMax As Int32 = Math.Max((moCurrentTemplate.lItemUB + 1) - ml_ROUTE_ITEMS + 1, 0)
                    If vscScroll.MaxValue <> lScrMax Then vscScroll.MaxValue = lScrMax
                    If vscScroll.Value > vscScroll.MaxValue Then vscScroll.Value = vscScroll.MaxValue

                    'Now, go through our items
                    For X As Int32 = vscScroll.Value To vscScroll.Value + ml_ROUTE_ITEMS - 1
                        Dim lListIdx As Int32 = X - vscScroll.Value
                        If X > moCurrentTemplate.lItemUB Then
                            If X = moCurrentTemplate.lItemUB + 1 Then
                                ConfigureItemAsAdd(lListIdx)
                            Else : ConfigureItemAsUnused(lListIdx)
                            End If
                            Continue For
                        End If
                        With moCurrentTemplate.uItems(X)
                            ConfigureItemAsUsed(lListIdx)
                            Dim sTemp As String = ""
                            If .lDestID > -1 AndAlso .iDestTypeID > -1 Then
                                If .iDestTypeID = ObjectType.eWormhole Then
                                    sTemp = "Wormhole"
                                    If goGalaxy Is Nothing = False Then
                                        Dim bFound As Boolean = False
                                        For iSystem As Int32 = 0 To goGalaxy.mlSystemUB
                                            For iWormhole As Int32 = 0 To goGalaxy.moSystems(iSystem).WormholeUB
                                                If goGalaxy.moSystems(iSystem).moWormholes(iWormhole).ObjectID = .lDestID Then
                                                    If .lLocX = goGalaxy.moSystems(iSystem).moWormholes(iWormhole).LocX1 AndAlso .lLocZ = goGalaxy.moSystems(iSystem).moWormholes(iWormhole).LocY1 Then
                                                        sTemp = "Wormhole To " & goGalaxy.moSystems(iSystem).moWormholes(iWormhole).System2.SystemName
                                                    ElseIf .lLocX = goGalaxy.moSystems(iSystem).moWormholes(iWormhole).LocX2 AndAlso .lLocZ = goGalaxy.moSystems(iSystem).moWormholes(iWormhole).LocY2 Then
                                                        sTemp = "Wormhole To " & goGalaxy.moSystems(iSystem).moWormholes(iWormhole).System1.SystemName
                                                    End If
                                                    bFound = True
                                                    Exit For
                                                End If
                                            Next
                                            If bFound = True Then Exit For
                                        Next
                                    End If
                                Else
                                    sTemp = GetCacheObjectValue(.lDestID, .iDestTypeID)
                                End If
                            Else : sTemp = ""
                            End If
                            If .lLocationID > 0 AndAlso .iLocationTypeID > 0 Then
                                sTemp &= " (" & GetCacheObjectValue(.lLocationID, .iLocationTypeID) & ")"
                            End If
                            If lblDestination(lListIdx).Caption <> sTemp Then lblDestination(lListIdx).Caption = sTemp

                            If .lLoadItemID = -1 AndAlso .iLoadItemTypeID = ObjectType.eMineral Then
                                sTemp = "Load: Any/All Minerals"
                            ElseIf .lLoadItemID = -1 AndAlso .iLoadItemTypeID = ObjectType.eComponentCache Then
                                sTemp = "Load: Any Components"
                            ElseIf .iLoadItemTypeID = ObjectType.eMineral Then
                                sTemp = ""
                                For Z As Int32 = 0 To glMineralUB
                                    If glMineralIdx(Z) = .lLoadItemID Then
                                        sTemp = "Load: " & goMinerals(Z).MineralName
                                        Exit For
                                    End If
                                Next Z
                            ElseIf .iLoadItemTypeID = ObjectType.eColonists Then
                                sTemp = "Load Colonists"
                            ElseIf .iLoadItemTypeID = ObjectType.eOfficers Then
                                sTemp = "Load Officers"
                            ElseIf .iLoadItemTypeID = ObjectType.eEnlisted Then
                                sTemp = "Load Enlisted"
                            ElseIf .lLoadItemID = -1 AndAlso .iLoadItemTypeID = -1 Then
                                sTemp = "Any Cargo"
                            ElseIf .lLoadItemID = -1 AndAlso .iLoadItemTypeID = ObjectType.eArmorTech Then
                                sTemp = "All Armor"
                            ElseIf .lLoadItemID = -1 AndAlso .iLoadItemTypeID = ObjectType.eEngineTech Then
                                sTemp = "All Engines"
                            ElseIf .lLoadItemID = -1 AndAlso .iLoadItemTypeID = ObjectType.eRadarTech Then
                                sTemp = "All Radar"
                            ElseIf .lLoadItemID = -1 AndAlso .iLoadItemTypeID = ObjectType.eShieldTech Then
                                sTemp = "All Shields"
                            ElseIf .lLoadItemID = -1 AndAlso .iLoadItemTypeID = ObjectType.eWeaponTech Then
                                sTemp = "All Weapons"
                            ElseIf .iLoadItemTypeID = ObjectType.eUnitDef OrElse .iLoadItemTypeID = ObjectType.eFacilityDef Then
                                sTemp = ""
                                For Z As Int32 = 0 To glEntityDefUB
                                    If glEntityDefIdx(Z) = .lLoadItemID AndAlso goEntityDefs(Z).ObjTypeID = .iLoadItemTypeID Then
                                        sTemp = "Build: " & goEntityDefs(X).DefName
                                        Exit For
                                    End If
                                Next Z
                            ElseIf .iLoadItemTypeID = Int16.MaxValue Then
                                sTemp = "Wait " & .lLoadItemID & " seconds"
                            ElseIf .iLoadItemTypeID = ObjectType.eWormhole Then
                                sTemp = "Jump"
                            ElseIf .lLoadItemID > -1 AndAlso .iLoadItemTypeID > -1 Then
                                sTemp = GetCacheObjectValue(.lLoadItemID, .iLoadItemTypeID)
                            Else : sTemp = ""
                            End If

                            If (.yFlag And 1) <> 0 Then
                                If sTemp = "" Then sTemp = "Unload Only" Else sTemp &= " (Unload)"
                            End If

                            If .iDestTypeID = ObjectType.eWormhole Then
                                sTemp = "Jump"
                            End If

                            If lblMineral(lListIdx).Caption <> sTemp Then lblMineral(lListIdx).Caption = sTemp
                        End With
                    Next X
                End If
            End If
            If mbScrollOne = True AndAlso vscScroll.Value < vscScroll.MaxValue Then
                mbScrollOne = False
                vscScroll.Value += 1
            End If
        Catch
        End Try
    End Sub

    Private Sub HandleRemoveItemButtonClick(ByVal sName As String)
        Try
            Dim lIdx As Int32 = sName.IndexOf("("c)
            If lIdx > -1 Then
                lIdx = CInt(Val(sName.Substring(lIdx + 1)))
                lIdx += vscScroll.Value

                If moCurrentTemplate.lItemUB >= lIdx Then
                    For X As Int32 = lIdx To moCurrentTemplate.lItemUB - 1
                        moCurrentTemplate.uItems(X) = moCurrentTemplate.uItems(X + 1)
                    Next X
                    moCurrentTemplate.lItemUB -= 1
                    ReDim Preserve moCurrentTemplate.uItems(moCurrentTemplate.lItemUB)
                End If
                mbNeedsSaved = True
            End If
        Catch
        End Try
    End Sub
    Private Sub HandleSetMineralButtonClick(ByVal sName As String)
        mbInPopup = True

        SetAllForPopup()

        msCurrentSetMineralControl = sName
        If moPopup Is Nothing Then
            moPopup = New frmRouteConfig.fraSelectMin(MyBase.moUILib)
            Me.AddChild(CType(moPopup, UIControl))
            AddHandler moPopup.CancelClicked, AddressOf Popup_CancelClick
            AddHandler moPopup.ItemAccepted, AddressOf Popup_ItemClick
        End If
        moPopup.Left = Me.Width \ 2 - moPopup.Width \ 2
        moPopup.Top = Me.Height \ 2 - moPopup.Height \ 2
        Try
            Dim lIdx As Int32 = sName.IndexOf("("c)
            If lIdx > -1 Then
                lIdx = CInt(Val(sName.Substring(lIdx + 1)))
                lIdx += vscScroll.Value

                Dim iTypeID As Int16 = 0
                Dim yFlag As Byte = 1

                If moCurrentTemplate.lItemUB >= lIdx Then
                    Dim lRouteIdx As Int32 = lIdx
                    lIdx = moCurrentTemplate.uItems(lRouteIdx).lLoadItemID
                    iTypeID = moCurrentTemplate.uItems(lRouteIdx).iLoadItemTypeID
                    yFlag = moCurrentTemplate.uItems(lRouteIdx).yFlag
                End If
                moPopup.SetFromSelected(lIdx, iTypeID, yFlag)
                mbNeedsSaved = True
            End If
        Catch
        End Try

        moPopup.Visible = True


    End Sub
    Private Sub HandleSetLocationButtonClick(ByVal sName As String)
        msCurrentSetLocControl = sName
        Me.Visible = False
        MyBase.moUILib.lUISelectState = UILib.eSelectState.eSelectRouteTemplateLoc
        MyBase.moUILib.AddNotification("Left-click a location for this way point.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        MyBase.moUILib.AddNotification("Right-click to cancel.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
    End Sub

    Private Sub Popup_ItemClick(ByVal lItemID As Int32, ByVal iTypeID As Int16, ByVal bUnload As Boolean)
        Try
            Dim lIdx As Int32 = msCurrentSetMineralControl.IndexOf("("c)
            If lIdx > -1 Then
                lIdx = CInt(Val(msCurrentSetMineralControl.Substring(lIdx + 1)))
                lIdx += vscScroll.Value

                If moCurrentTemplate.lItemUB >= lIdx Then
                    With moCurrentTemplate.uItems(lIdx)
                        .lLoadItemID = lItemID
                        .iLoadItemTypeID = iTypeID
                        If bUnload = True Then .yFlag = 1 Else .yFlag = 0
                    End With
                End If
 
            End If
        Catch
        End Try
        moPopup.Visible = False
        mbInPopup = False
        SetAllForPopup()
    End Sub
    Private Sub Popup_CancelClick()
        moPopup.Visible = False
        mbInPopup = False
        SetAllForPopup()
    End Sub
    Private Sub SetAllForPopup()
        Dim bValue As Boolean = Not mbInPopup
        For X As Int32 = 0 To ml_ROUTE_ITEMS - 1
            btnRemoveItem(X).Enabled = bValue
            btnSetMineral(X).Enabled = bValue
            btnSetLocation(X).Enabled = bValue
        Next X

        btnAssignAndRunToSelected.Enabled = bValue
        btnAssignToSelected.Enabled = bValue
        btnClear.Enabled = bValue
        btnClose.Enabled = bValue
        btnDeleteTemplate.Enabled = bValue
        btnCopyRoute.Enabled = bValue
        lblTemplates.Enabled = bValue

        cboTemplates.Enabled = bValue
        btnSaveTemplate.Enabled = bValue
        btnDeleteTemplate.Enabled = bValue
    End Sub

    Public Sub SetLocResultVector(ByVal vecLoc As Vector3)
        Try
            Dim sName As String = msCurrentSetLocControl
            Dim lIdx As Int32 = sName.IndexOf("("c)

            If lIdx > -1 Then

                lIdx = CInt(Val(sName.Substring(lIdx + 1)))
                lIdx += vscScroll.Value


                Dim lDestID As Int32 = -1
                Dim iDestTypeID As Int16 = -1
                If goGalaxy.CurrentSystemIdx <> -1 Then
                    If goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx <> -1 AndAlso (glCurrentEnvirView = CurrentView.ePlanetMapView OrElse glCurrentEnvirView = CurrentView.ePlanetView) Then
                        lDestID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).moPlanets(goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).CurrentPlanetIdx).ObjectID
                        iDestTypeID = ObjectType.ePlanet
                    Else
                        lDestID = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx).ObjectID
                        iDestTypeID = ObjectType.eSolarSystem
                    End If
                Else
                    Return
                End If

                If moCurrentTemplate.lItemUB < lIdx Then
                    moCurrentTemplate.lItemUB += 1
                    ReDim Preserve moCurrentTemplate.uItems(moCurrentTemplate.lItemUB)
                End If

                If moCurrentTemplate.lItemUB >= lIdx Then
                    With moCurrentTemplate.uItems(lIdx)
                        .lDestID = lDestID
                        .iDestTypeID = iDestTypeID

                        .lLocX = CInt(vecLoc.X)
                        .lLocZ = CInt(vecLoc.Z)
                    End With
                End If

            End If
        Catch
        End Try

        Me.Visible = True
        MyBase.moUILib.lUISelectState = UILib.eSelectState.eNoSelectState
        mbScrollOne = True
    End Sub

    Public Sub SetLocResultGUID(ByRef oEntity As BaseEntity)
        Try
            Dim sName As String = msCurrentSetLocControl
            Dim lIdx As Int32 = sName.IndexOf("("c)

            If lIdx > -1 Then
                lIdx = CInt(Val(sName.Substring(lIdx + 1)))
                lIdx += vscScroll.Value

                If moCurrentTemplate.lItemUB < lIdx Then
                    moCurrentTemplate.lItemUB += 1
                    ReDim Preserve moCurrentTemplate.uItems(moCurrentTemplate.lItemUB)
                End If

                If moCurrentTemplate.lItemUB >= lIdx Then
                    With moCurrentTemplate.uItems(lIdx)
                        .lDestID = oEntity.ObjectID
                        .iDestTypeID = oEntity.ObjTypeID

                        .lLocX = CInt(oEntity.LocX)
                        .lLocZ = CInt(oEntity.LocZ)
                    End With
                End If

            End If
        Catch
        End Try

        Me.Visible = True
        MyBase.moUILib.lUISelectState = UILib.eSelectState.eNoSelectState
        mbScrollOne = True
    End Sub

    Public Sub SetLocResultWormhole(ByRef oWormhole As Wormhole, ByVal bSystem1 As Boolean)
        Try
            Dim sName As String = msCurrentSetLocControl
            Dim lIdx As Int32 = sName.IndexOf("("c)
            If lIdx > -1 Then
                lIdx = CInt(Val(sName.Substring(lIdx + 1)))
                lIdx += vscScroll.Value

                If moCurrentTemplate.lItemUB < lIdx Then
                    moCurrentTemplate.lItemUB += 1
                    ReDim Preserve moCurrentTemplate.uItems(moCurrentTemplate.lItemUB)
                End If

                If moCurrentTemplate.lItemUB >= lIdx Then
                    With moCurrentTemplate.uItems(lIdx)
                        .lDestID = oWormhole.ObjectID
                        .iDestTypeID = ObjectType.eWormhole

                        If bSystem1 = True Then
                            .lLocX = oWormhole.LocX1
                            .lLocZ = oWormhole.LocY1
                        Else
                            .lLocX = oWormhole.LocX2
                            .lLocZ = oWormhole.LocY2
                        End If
                    End With
                End If

            End If
        Catch
        End Try

        Me.Visible = True
        MyBase.moUILib.lUISelectState = UILib.eSelectState.eNoSelectState
        mbScrollOne = True
    End Sub
 
    Private Sub btnClear_Click(ByVal sName As String) Handles btnClear.Click
        If btnClear.Caption.ToUpper = "CONFIRM" Then

            moCurrentTemplate.lItemUB = -1
            mbNeedsSaved = True

            btnClear.Caption = "Clear"
        Else : btnClear.Caption = "Confirm"
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub
 
    Private Sub cboTemplates_ItemSelected(ByVal lItemIndex As Integer) Handles cboTemplates.ItemSelected
        If cboTemplates.ListIndex > -1 Then
            Dim lID As Int32 = cboTemplates.ItemData(cboTemplates.ListIndex)
            Try
                For X As Int32 = 0 To lTemplateUB
                    If oTemplates(X) Is Nothing = False AndAlso oTemplates(X).lTemplateID = lID Then

                        moCurrentTemplate = New RouteTemplate

                        With oTemplates(X)
                            moCurrentTemplate.lPlayerID = glPlayerID
                            moCurrentTemplate.lTemplateID = .lTemplateID
                            moCurrentTemplate.sTemplateName = .sTemplateName

                            txtTemplateName.Caption = .sTemplateName

                            moCurrentTemplate.lItemUB = .lItemUB
                            ReDim moCurrentTemplate.uItems(.lItemUB)
                            For Y As Int32 = 0 To .lItemUB
                                moCurrentTemplate.uItems(Y) = .uItems(Y)
                            Next Y
                        End With

                        Exit For
                    End If
                Next X
            Catch
            End Try
        End If
    End Sub

    Public Sub SetFromTemplateID(ByVal lTemplateID As Int32)
        cboTemplates.FindComboItemData(lTemplateID)
    End Sub

    Private Sub btnDeleteTemplate_Click(ByVal sName As String) Handles btnDeleteTemplate.Click
        If btnDeleteTemplate.Caption.ToUpper = "CONFIRM" Then
            btnDeleteTemplate.Caption = "Delete"

            If moCurrentTemplate.lTemplateID > 0 Then
                'Send the delete off
                Dim yMsg(6) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eEditRouteTemplate).CopyTo(yMsg, lPos) : lPos += 2
                yMsg(lPos) = 0 : lPos += 1        'edit code of 0 for delete
                System.BitConverter.GetBytes(moCurrentTemplate.lTemplateID).CopyTo(yMsg, lPos) : lPos += 4
                MyBase.moUILib.SendMsgToPrimary(yMsg)
            Else
                btnClear_Click(sName)
            End If
        Else
            btnDeleteTemplate.Caption = "Confirm"
        End If
    End Sub

    Private Sub btnSaveTemplate_Click(ByVal sName As String) Handles btnSaveTemplate.Click
        'send the save off... confirm if the one they have selected shuold be overwritten or not
        Dim yMsg(30 + ((moCurrentTemplate.lItemUB + 1) * 21)) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eEditRouteTemplate).CopyTo(yMsg, lPos) : lPos += 2

        If moCurrentTemplate.lTemplateID > 0 Then
            yMsg(lPos) = 1 : lPos += 1         'edit code of 1-254 for edit
            System.BitConverter.GetBytes(moCurrentTemplate.lTemplateID).CopyTo(yMsg, lPos) : lPos += 4
        Else
            yMsg(lPos) = 255 : lPos += 1        'edit code of 255 for add
            System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
        End If

        Dim sTemplateName As String = txtTemplateName.Caption
        If sTemplateName.Length > 20 Then sTemplateName = sTemplateName.Substring(0, 20)

        System.Text.ASCIIEncoding.ASCII.GetBytes(sTemplateName).CopyTo(yMsg, lPos) : lPos += 20

        System.BitConverter.GetBytes(moCurrentTemplate.lItemUB + 1).CopyTo(yMsg, lPos) : lPos += 4

        For X As Int32 = 0 To moCurrentTemplate.lItemUB
            With moCurrentTemplate.uItems(X)
                System.BitConverter.GetBytes(.lDestID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.iDestTypeID).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(.lLoadItemID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.iLoadItemTypeID).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(.lLocX).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(.lLocZ).CopyTo(yMsg, lPos) : lPos += 4
                yMsg(lPos) = .yFlag : lPos += 1
            End With
        Next X

        MyBase.moUILib.SendMsgToPrimary(yMsg)
        FillTemplates()
        mbNeedsSaved = False
    End Sub

    Private Sub btnAssignAndRunToSelected_Click(ByVal sName As String) Handles btnAssignAndRunToSelected.Click
        SubmitAssignCmd(GlobalMessageCode.eBeginRouteForEntities)
    End Sub

    Private Sub btnAssignToSelected_Click(ByVal sName As String) Handles btnAssignToSelected.Click
        SubmitAssignCmd(GlobalMessageCode.eApplyRouteToEntities)
    End Sub

    Private Sub SubmitAssignCmd(ByVal iCmd As Int16)
        Dim lCnt As Int32 = 0

        If moCurrentTemplate.lTemplateID < 1 OrElse mbNeedsSaved = True Then
            MyBase.moUILib.AddNotification("You must save this template first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        Try
            If goCurrentEnvir Is Nothing = False Then
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) > -1 Then
                        Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(X)
                        If oEntity Is Nothing = False Then
                            If oEntity.OwnerID = glPlayerID AndAlso oEntity.bSelected = True AndAlso oEntity.ObjTypeID = ObjectType.eUnit Then
                                lCnt += 1
                            End If
                        End If
                    End If
                Next X


                Dim yMsg(9 + (lCnt * 4)) As Byte
                Dim lPos As Int32 = 0

                System.BitConverter.GetBytes(iCmd).CopyTo(yMsg, lPos) : lPos += 2
                System.BitConverter.GetBytes(moCurrentTemplate.lTemplateID).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4

                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) > -1 Then
                        Dim oEntity As BaseEntity = goCurrentEnvir.oEntity(X)
                        If oEntity Is Nothing = False Then
                            If oEntity.OwnerID = glPlayerID AndAlso oEntity.bSelected = True AndAlso oEntity.ObjTypeID = ObjectType.eUnit Then
                                System.BitConverter.GetBytes(oEntity.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                            End If
                        End If
                    End If
                Next X

                MyBase.moUILib.SendMsgToPrimary(yMsg)

            End If
        Catch
        End Try
    End Sub

    Private Sub btnCopyRoute_Click(ByVal sName As String) Handles btnCopyRoute.Click
        moCurrentTemplate.lTemplateID = 0
        txtTemplateName.Caption = ""
    End Sub
End Class