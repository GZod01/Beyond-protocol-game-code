Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmBuilderWrksht
    Inherits UIWindow

    Private Enum elItemPropIdx As Int32
        eBaseHull = 0
        eEngine = 1
        eShield = 2
        eRadar = 3
        eWpn = 4        'starting index

        eArmor = 14
        eCrew = 15
        eProduction = 16
    End Enum

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private lblHull As UILabel
    Private lblPower As UILabel
    Private txtHull() As UITextBox
    Private txtPower() As UITextBox
    Private cboComponent() As UIComboBox
    Private txtQty() As UITextBox
    Private lblItem() As UILabel
    Private lblQty As UILabel
    Private lblDescription As UILabel
    Private lnDiv2 As UILine
    Private lnDiv3 As UILine
    Private btnView() As UIButton
    Private lblTotalHull As UILabel
    Private lblTotalPower As UILabel

    Private WithEvents chkFilter As UICheckBox
    Private WithEvents btnClose As UIButton
    Private WithEvents btnMinMax As UIButton

    Private mbCrewLocked As Boolean = False
    Private lblLockCrew As UILabel
    Private mlPowerMod As Int32 = 1

    Private mbFilter As Boolean = True
    Private mbLoading As Boolean = True
    Private mlComboIDs() As Int32
    Private mlQtys() As Int32

    Private mbMinimized As Boolean
    Private mlStartY As Int32

    Private mlLastComponentRefresh As Int32 = 0

    Public Shared Sub DoTest()
        Dim oNew As New frmBuilderWrksht(goUILib)
        oNew.FullScreen = True
        oNew.Visible = True

    End Sub

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        mbMinimized = False

        'frmBuilderWrksht initial props
        With Me
            .ControlName = "frmBuilderWrksht"
            .Left = 348
            .Top = 98
            If NewTutorialManager.TutorialOn = False Then
                If muSettings.BuilderWorksheetX <> -1 Then .Left = muSettings.BuilderWorksheetX
                If muSettings.BuilderWorksheetY <> -1 Then .Top = muSettings.BuilderWorksheetY
            End If
            .Width = 472
            .Height = 511
            .Enabled = True
            .Visible = True

            If .Left < 0 OrElse .Left + .Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If .Top < 0 OrElse .Top + .Height > oUILib.oDevice.PresentationParameters.BackBufferHeight Then .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - .Height

            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 133
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Builder Worksheet"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 24 - Me.BorderLineWidth
            .Top = Me.BorderLineWidth
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

        'btnMinMax initial props
        btnMinMax = New UIButton(oUILib)
        With btnMinMax
            .ControlName = "btnMinMax"
            .Left = btnClose.Left - 25
            .Top = Me.BorderLineWidth
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(144, 96, 23, 23)
            .ControlImageRect_Normal = .ControlImageRect
            .ControlImageRect_Disabled = New Rectangle(121, 96, 23, 23)
            .ControlImageRect_Pressed = New Rectangle(168, 96, 23, 23)
        End With
        Me.AddChild(CType(btnMinMax, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 2
            .Top = 26
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lblHull initial props
        lblHull = New UILabel(oUILib)
        With lblHull
            .ControlName = "lblHull"
            .Left = 345
            .Top = 65
            .Width = 50
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Hull"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblHull, UIControl))

        'lblPower initial props
        lblPower = New UILabel(oUILib)
        With lblPower
            .ControlName = "lblPower"
            .Left = 410
            .Top = 65
            .Width = 50
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Power"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblPower, UIControl))

        'Private txtHull() As UITextBox
        'Private txtPower() As UITextBox
        'Private cboComponent() As UIComboBox
        'Private txtQty() As UITextBox
        'Private lblItem() As UILabel
        Dim lItemUB As Int32 = 16
        ReDim txtHull(lItemUB)
        ReDim txtPower(lItemUB)
        ReDim cboComponent(lItemUB)
        ReDim txtQty(lItemUB)
        ReDim lblItem(lItemUB)
        ReDim btnView(lItemUB)

        'For X As Int32 = 0 To lItemUB
        '    AddControl(oUILib, X)
        'Next X
        'For X As Int32 = 11 To 0 Step -1
        '    AddControl(oUILib, X)
        'Next
        'For X As Int32 = 12 To lItemUB
        '    AddControl(oUILib, X)
        'Next X

        For X As Int32 = lItemUB To 0 Step -1
            AddControl(oUILib, X)
        Next X

        'lblQty initial props
        lblQty = New UILabel(oUILib)
        With lblQty
            .ControlName = "lblQty"
            .Left = 280
            .Top = 65
            .Width = 60
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Qty"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblQty, UIControl))

        'lblLockCrew initial props
        lblLockCrew = New UILabel(oUILib)
        With lblLockCrew
            .ControlName = "lblLockCrew"
            .Left = txtQty(elItemPropIdx.eCrew).Left - 16
            .Top = txtQty(elItemPropIdx.eCrew).Top
            .Width = 16
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = grc_UI(elInterfaceRectangle.eLock)
            .BackImageColor = muSettings.InterfaceBorderColor
            .bAcceptEvents = True
            .ToolTipText = "Click to lock the crew count to the current value."
        End With
        Me.AddChild(CType(lblLockCrew, UIControl))
        AddHandler lblLockCrew.OnMouseDown, AddressOf LockCrewClick

        'lblDescription initial props
        lblDescription = New UILabel(oUILib)
        With lblDescription
            .ControlName = "lblDescription"
            .Left = 0
            .Top = 30
            .Width = Me.Width
            .Height = 36
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.WordBreak
            .Caption = "Use this window to calculate your hull and power needs for a design while you design. To do so, select or enter the hull and power requirements."
        End With
        Me.AddChild(CType(lblDescription, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = 2
            .Top = 66
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        'lnDiv3 initial props
        Dim lTotalBarLineTop As Int32 = Me.Height - 30
        lnDiv3 = New UILine(oUILib)
        With lnDiv3
            .ControlName = "lnDiv3"
            .Left = 2
            If lblItem Is Nothing = False Then
                lTotalBarLineTop = lblItem(lblItem.GetUpperBound(0)).Top + lblItem(lblItem.GetUpperBound(0)).Height + 2
            End If
            .Top = lTotalBarLineTop
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv3, UIControl))

        'lblTotalHull initial props
        lblTotalHull = New UILabel(oUILib)
        With lblTotalHull
            .ControlName = "lblTotalHull"
            .Left = 0
            .Top = lTotalBarLineTop + 5
            .Width = Me.Width - 5
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Hull Remaining: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
            .ToolTipText = "If this value is negative, the design will not work."
        End With
        Me.AddChild(CType(lblTotalHull, UIControl))

        'lblTotalPower initial props
        lblTotalPower = New UILabel(oUILib)
        With lblTotalPower
            .ControlName = "lblTotalPower"
            .Left = 0
            .Top = lTotalBarLineTop + 25
            .Width = Me.Width - 5
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Power Remaining: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
            .ToolTipText = "If this value is negative, the design will not work" & vbCrLf & _
                           "unless the design is for a facility that is in a colony."
        End With
        Me.AddChild(CType(lblTotalPower, UIControl))

        'chkFilter initial props
        chkFilter = New UICheckBox(oUILib)
        With chkFilter
            .ControlName = "chkFilter"
            .Left = 300
            .Top = 5
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Filter Archived"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkFilter, UIControl))

        For X As Int32 = lItemUB To 0 Step -1
            If cboComponent(X) Is Nothing = False Then
                Me.AddChild(CType(cboComponent(X), UIControl))
            End If
        Next X

        FillComponents()
        SetComboIDs()
        SetQtyVals()

        mbLoading = False

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        mlStartY = Me.Top
    End Sub

    Private Sub AddControl(ByRef oUILib As UILib, ByVal X As Int32)

        Dim lTopVal As Int32 = 85 + (X * 22)
        Dim lLeftVal As Int32 = 5

        'lblItem() initial props
        lblItem(X) = New UILabel(oUILib)
        With lblItem(X)
            .ControlName = "lblItem(" & X & ")"
            .Left = lLeftVal
            .Top = lTopVal
            .Width = 55
            .Height = 18
            .Enabled = True
            .Visible = True
            lLeftVal += .Width + 5
            Select Case X
                Case elItemPropIdx.eArmor
                    .Caption = "Armor"
                Case elItemPropIdx.eBaseHull
                    .Caption = "Hull"
                Case elItemPropIdx.eCrew
                    .Caption = "Crew"
                    .Width = 210
                Case elItemPropIdx.eEngine
                    .Caption = "Engine"
                Case elItemPropIdx.eProduction
                    .Caption = "Production"
                    .Width = 75
                Case elItemPropIdx.eRadar
                    .Caption = "Radar"
                Case elItemPropIdx.eShield
                    .Caption = "Shield"
                Case Else       'likely a weapon
                    If X >= elItemPropIdx.eWpn AndAlso X < elItemPropIdx.eWpn + 10 Then
                        .Caption = "Weapon"
                    Else
                        .Caption = "Unknown"
                    End If
            End Select
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblItem(X), UIControl))

        'cboComponent() initial props
        If X < elItemPropIdx.eCrew Then
            'btnView() inital props
            btnView(X) = New UIButton(oUILib)
            With btnView(X)
                .ControlName = "btnView(" & X & ")"
                .Left = lLeftVal
                .Top = lTopVal
                .Width = 18
                .Height = 18
                lLeftVal += .Width + 5
                .Enabled = False
                .Visible = True
                .Caption = "V"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
                .ToolTipText = "Click to view this design."
                AddHandler btnView(X).Click, AddressOf ViewDesign
            End With
            Me.AddChild(CType(btnView(X), UIControl))

            cboComponent(X) = New UIComboBox(oUILib)
            With cboComponent(X)
                .ControlName = "cboComponent(" & X & ")"
                .Left = lLeftVal
                .Top = lTopVal
                .Width = 180
                .Height = 18
                lLeftVal += .Width + 5
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
                .l_ListBoxHeight = 197 '192
                If X = elItemPropIdx.eBaseHull Then
                    AddHandler cboComponent(X).ItemSelected, AddressOf HullCombo
                Else
                    AddHandler cboComponent(X).ItemSelected, AddressOf ComboBoxItemSelect
                End If
            End With
            'Me.AddChild(CType(cboComponent(X), UIControl))
        Else
            lLeftVal += 208
        End If

        'txtQty() initial props
        If X > elItemPropIdx.eRadar Then
            txtQty(X) = New UITextBox(oUILib)
            With txtQty(X)
                .ControlName = "txtQty(" & X & ")"
                .Left = lLeftVal
                .Top = lTopVal
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = X > 3
                .Caption = "0"
                .ForeColor = muSettings.InterfaceTextBoxForeColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
                .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)

                If X >= elItemPropIdx.eArmor Then
                    .MaxLength = 8
                Else
                    .MaxLength = 2
                End If
                .bNumericOnly = True

                .Locked = X < 4     'hull, engine, shield, radar
                If .Locked = True Then .Caption = "1"

                .BorderColor = muSettings.InterfaceBorderColor

                .DoNotRender = UITextBox.DoNotRenderSetting.eDoNotRenderControl

                AddHandler txtQty(X).TextChanged, AddressOf QtyTextValueChange
            End With
            Me.AddChild(CType(txtQty(X), UIControl))
        End If
        lLeftVal += 65

        'txtHull() initial props
        txtHull(X) = New UITextBox(oUILib)
        With txtHull(X)
            .ControlName = "txtHull(" & X & ")"
            .Left = lLeftVal
            .Top = lTopVal
            .Width = 65
            .Height = 18
            lLeftVal += .Width + 5
            .Enabled = True
            .Visible = True
            .Caption = "0"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor

            .DoNotRender = UITextBox.DoNotRenderSetting.eDoNotRenderControl

            AddHandler txtHull(X).TextChanged, AddressOf HullPowerTextValueChange
        End With
        Me.AddChild(CType(txtHull(X), UIControl))

        'txtPower() initial props
        txtPower(X) = New UITextBox(oUILib)
        With txtPower(X)
            .ControlName = "txtPower(" & X & ")"
            .Left = lLeftVal
            .Top = lTopVal
            .Width = 60
            .Height = 18
            lLeftVal += .Width + 5
            .Enabled = True
            .Visible = True
            .Caption = "0"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor

            .DoNotRender = UITextBox.DoNotRenderSetting.eDoNotRenderControl

            AddHandler txtPower(X).TextChanged, AddressOf HullPowerTextValueChange
        End With
        Me.AddChild(CType(txtPower(X), UIControl))
    End Sub

    Private Sub HullCombo(ByVal lIndex As Int32)
        If mbLoading = True Then Return

        'mbLoading = True
        'FillComponents()
        'SetComboIDs()
        'SetQtyVals()
        'mbLoading = False
        'RecalcAll()

        Dim oHullTech As HullTech = Nothing

        If goCurrentPlayer Is Nothing = False AndAlso cboComponent Is Nothing = False AndAlso cboComponent.GetUpperBound(0) >= elItemPropIdx.eBaseHull Then
            Dim cbo As UIComboBox = cboComponent(elItemPropIdx.eBaseHull)
            If cbo Is Nothing = False AndAlso cbo.ListIndex > 0 Then
                If btnView(0).Enabled <> True Then btnView(0).Enabled = True
                oHullTech = CType(goCurrentPlayer.GetTech(cbo.ItemData(cbo.ListIndex), CShort(cbo.ItemData2(cbo.ListIndex))), HullTech)
            Else
                btnView(0).Enabled = False
            End If
        End If

        Dim bProdVis As Boolean = True
        mlPowerMod = 1

        If oHullTech Is Nothing = False Then
            If oHullTech.yTypeID = 2 Then
                Select Case oHullTech.ySubTypeID
                    Case 0, 1, 3, 5, 6, 7, 8, 9 ', 10
                        If (oHullTech.ModelID And 255) = 148 Then
                            bProdVis = False
                        End If
                    Case 4
                        If (oHullTech.ModelID And 255) = 111 Then
                            mlPowerMod = 2
                        Else : mlPowerMod = 3
                        End If
                    Case Else
                        bProdVis = False
                End Select
            ElseIf oHullTech.yTypeID = 7 Then
                If oHullTech.ySubTypeID <> 0 AndAlso oHullTech.ySubTypeID <> 1 AndAlso oHullTech.ySubTypeID <> 3 Then
                    bProdVis = False
                End If
            ElseIf oHullTech.yTypeID = 8 Then
                If oHullTech.ySubTypeID <> 6 Then
                    bProdVis = False
                End If
            End If
        End If

        mbLoading = True
        mlComboIDs = Nothing
        For X As Int32 = 1 To cboComponent.GetUpperBound(0)
            If cboComponent(X) Is Nothing = False Then
                cboComponent(X).ListIndex = -1
                If btnView(X).Enabled <> False Then btnView(X).Enabled = False
            End If
        Next X

        FillComponents()
        SetComboIDs()
        SetQtyVals()
        mbLoading = False
        RecalcAll()

        'lblItem(elItemPropIdx.eProduction).Visible = bProdVis
        If bProdVis = True Then lblItem(elItemPropIdx.eProduction).Caption = "Production" Else lblItem(elItemPropIdx.eProduction).Caption = "Does not produce"
        txtQty(elItemPropIdx.eProduction).Visible = bProdVis
        txtHull(elItemPropIdx.eProduction).Visible = bProdVis
        txtPower(elItemPropIdx.eProduction).Visible = bProdVis
    End Sub
    Private Sub ComboBoxItemSelect(ByVal lIndex As Int32)
        If mbLoading = True Then Return

        'Ok, a qty textbox value changed...
        If cboComponent Is Nothing = False Then
            For X As Int32 = 1 To cboComponent.GetUpperBound(0)
                If cboComponent(X) Is Nothing = False Then
                    Dim bEnabled As Boolean = (cboComponent(X).ListIndex > -1) AndAlso (cboComponent(X).ItemData(cboComponent(X).ListIndex) > -1)
                    If btnView(X).Enabled <> bEnabled Then btnView(X).Enabled = bEnabled
                End If
            Next
            If mlComboIDs Is Nothing OrElse mlComboIDs.GetUpperBound(0) <> cboComponent.GetUpperBound(0) Then
                SetComboIDs()
                RecalcAll()
                Return
            End If

            Dim bRecalc As Boolean = False
            For X As Int32 = 0 To cboComponent.GetUpperBound(0)
                If cboComponent(X) Is Nothing = False Then
                    Dim lNewVal As Int32 = -1

                    If cboComponent(X).ListIndex > -1 Then
                        lNewVal = cboComponent(X).ItemData(cboComponent(X).ListIndex)
                    End If

                    If lNewVal <> mlComboIDs(X) Then
                        bRecalc = True
                        mlComboIDs(X) = lNewVal
                        SetHullAndPower(X)
                    End If
                End If
            Next X
            If bRecalc = True Then RecalcTotals(True)
        End If
    End Sub
    Private Sub QtyTextValueChange()
        If mbLoading = True Then Return

        Dim bRecalcCrew As Boolean = False

        'Ok, a qty textbox value changed...
        If txtQty Is Nothing = False Then
            If mlQtys Is Nothing OrElse mlQtys.GetUpperBound(0) <> txtQty.GetUpperBound(0) Then
                SetQtyVals()
                RecalcAll()
                Return
            End If

            Dim bRecalc As Boolean = False
            For X As Int32 = 0 To txtQty.GetUpperBound(0)
                If txtQty(X) Is Nothing = False Then
                    Dim lNewVal As Int32 = 0
                    Dim sVal As String = txtQty(X).Caption
                    If sVal Is Nothing Then sVal = "0"
                    sVal = sVal.Replace(",", "")
                    sVal = sVal.Replace(".", "")
                    If IsNumeric(sVal) = True Then
                        lNewVal = CInt(Val(sVal))
                    End If

                    If lNewVal < 0 Then
                        lNewVal = 0
                        Dim bTempLoad As Boolean = mbLoading
                        mbLoading = True
                        txtQty(X).Caption = lNewVal.ToString
                        mbLoading = bTempLoad
                    End If

                    If lNewVal <> mlQtys(X) Then
                        If X <> elItemPropIdx.eCrew Then bRecalcCrew = True
                        bRecalc = True
                        mlQtys(X) = lNewVal

                        SetHullAndPower(X)
                    End If
                End If
            Next X
            If bRecalc = True Then RecalcTotals(bRecalcCrew)
        End If
    End Sub
    Private Sub HullPowerTextValueChange()
        If mbLoading = True Then Return
        RecalcTotals(False)
    End Sub

    Private Sub SetHullAndPower(ByVal lIdx As Int32)
        Dim lPower As Int32 = 0
        Dim lHull As Int32 = 0

        If lIdx = elItemPropIdx.eCrew Then
            lHull = -11
            If goCurrentPlayer Is Nothing = False Then
                lHull = -goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBetterHullToResidence)
            End If
        ElseIf lIdx = elItemPropIdx.eProduction Then
            lHull = -1
        Else
            If cboComponent Is Nothing = False AndAlso lIdx > -1 AndAlso lIdx <= cboComponent.GetUpperBound(0) Then
                Dim cbo As UIComboBox = cboComponent(lIdx)

                If cbo Is Nothing = False AndAlso cbo.ListIndex > -1 Then
                    If cbo.ItemData(cbo.ListIndex) <> -1 Then
                        If goCurrentPlayer Is Nothing = False Then
                            Dim oTech As Base_Tech = goCurrentPlayer.GetTech(cbo.ItemData(cbo.ListIndex), CShort(cbo.ItemData2(cbo.ListIndex)))
                            If oTech Is Nothing = False Then
                                Select Case oTech.ObjTypeID
                                    Case ObjectType.eArmorTech
                                        lHull = -CType(oTech, ArmorTech).HullUsagePerPlate
                                    Case ObjectType.eEngineTech
                                        With CType(oTech, EngineTech)
                                            lHull = -.HullRequired
                                            lPower = .PowerProd * mlPowerMod
                                        End With
                                    Case ObjectType.eHullTech
                                        With CType(oTech, HullTech)
                                            lHull = .HullSize
                                            lPower = -.PowerRequired
                                        End With
                                    Case ObjectType.eRadarTech
                                        With CType(oTech, RadarTech)
                                            lHull = -.HullRequired
                                            lPower = -.PowerRequired
                                        End With
                                    Case ObjectType.eShieldTech
                                        With CType(oTech, ShieldTech)
                                            lHull = -.HullRequired
                                            lPower = -.PowerRequired
                                        End With
                                    Case ObjectType.eWeaponTech
                                        With CType(oTech, WeaponTech)
                                            lHull = -.HullRequired
                                            lPower = -.PowerRequired
                                        End With
                                End Select
                            End If
                        End If
                    End If
                End If
            End If
        End If

        Dim lQty As Int32 = 0
        If txtQty(lIdx) Is Nothing = False AndAlso txtQty(lIdx).Visible = True Then
            Dim sVal As String = txtQty(lIdx).Caption
            If sVal Is Nothing Then sVal = "0"
            sVal = sVal.Replace(",", "")
            sVal = sVal.Replace(".", "")
            If IsNumeric(sVal) = True Then lQty = CInt(Val(sVal))
        Else
            lQty = 1
        End If

        Try
            lPower *= lQty
        Catch
            lPower = Int32.MaxValue
        End Try
        Try
            lHull *= lQty
        Catch
            lHull = Int32.MaxValue
        End Try

        Dim bPrevLoad As Boolean = mbLoading
        mbLoading = True
        If lPower = Int32.MaxValue Then txtPower(lIdx).Caption = "Invalid" Else txtPower(lIdx).Caption = lPower.ToString("#,##0")
        If lHull = Int32.MaxValue Then txtHull(lIdx).Caption = "Invalid" Else txtHull(lIdx).Caption = lHull.ToString("#,##0")
        mbLoading = bPrevLoad
    End Sub

    Private Sub RecalcCrew()
        Dim lTotalCrew As Int32 = 0
        Dim lRecCrew As Int32 = 0

        If lblItem Is Nothing = False Then
            For X As Int32 = 0 To lblItem.GetUpperBound(0)
                lTotalCrew += GetLineItemCrew(X)
            Next X
        End If

        lRecCrew = CInt(GetPositiveHullsOnly() / 100)

        'Now, update our thing
        Dim oHullTech As HullTech = Nothing
        If goCurrentPlayer Is Nothing = False AndAlso cboComponent Is Nothing = False AndAlso cboComponent.GetUpperBound(0) >= elItemPropIdx.eBaseHull Then
            Dim cbo As UIComboBox = cboComponent(elItemPropIdx.eBaseHull)
            If cbo Is Nothing = False AndAlso cbo.ListIndex > -1 Then
                oHullTech = CType(goCurrentPlayer.GetTech(cbo.ItemData(cbo.ListIndex), CShort(cbo.ItemData2(cbo.ListIndex))), HullTech)
                If oHullTech Is Nothing = False Then
                    If oHullTech.HullSize <= 750 Then
                        lTotalCrew = 2
                        lRecCrew = 2
                    Else
                        lRecCrew = CInt(oHullTech.HullSize / 100)
                        If oHullTech.yTypeID = 2 Then lRecCrew = 0
                    End If
                End If
            End If
        End If

        'Now, set our captions
        Dim bPrevLoad As Boolean = mbLoading
        mbLoading = True

        If txtQty Is Nothing = False AndAlso txtQty.GetUpperBound(0) >= elItemPropIdx.eCrew Then
            If mbCrewLocked = True Then
                lTotalCrew = 0
                Dim sVal As String = txtQty(elItemPropIdx.eCrew).Caption
                If sVal Is Nothing Then sVal = "0"
                sVal = sVal.Replace(",", "")
                sVal = sVal.Replace(".", "")
                If IsNumeric(sVal) = False Then sVal = "0"
                lTotalCrew = CInt(Val(sVal))
            End If
            txtQty(elItemPropIdx.eCrew).Caption = lTotalCrew.ToString()
            SetHullAndPower(elItemPropIdx.eCrew)
            If lTotalCrew < lRecCrew Then
                lblItem(elItemPropIdx.eCrew).ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            Else
                lblItem(elItemPropIdx.eCrew).ForeColor = muSettings.InterfaceBorderColor
            End If
        End If

        If lblItem Is Nothing = False AndAlso lblItem.GetUpperBound(0) >= elItemPropIdx.eCrew Then
            If lRecCrew = 0 Then
                lblItem(elItemPropIdx.eCrew).Caption = "Crew"
            Else
                lblItem(elItemPropIdx.eCrew).Caption = "Crew (Recommended: " & lRecCrew.ToString() & ")"
            End If
        End If

        mbLoading = bPrevLoad
    End Sub

    Private Function GetLineItemCrew(ByVal lIdx As Int32) As Int32

        If lIdx = elItemPropIdx.eCrew Then Return 0
        If lIdx = elItemPropIdx.eProduction Then Return 0

        Dim lResult As Int32 = 0
        If cboComponent Is Nothing = False AndAlso lIdx > -1 AndAlso lIdx <= cboComponent.GetUpperBound(0) Then
            Dim cbo As UIComboBox = cboComponent(lIdx)
            If cbo Is Nothing = False Then
                If cbo.ListIndex > -1 Then
                    If goCurrentPlayer Is Nothing = False Then
                        Dim oTech As Base_Tech = goCurrentPlayer.GetTech(cbo.ItemData(cbo.ListIndex), CShort(cbo.ItemData2(cbo.ListIndex)))
                        If oTech Is Nothing = False Then
                            If oTech.oProductionCost Is Nothing = False Then
                                With oTech.oProductionCost
                                    lResult = .ColonistCost + .EnlistedCost + .OfficerCost


                                    Dim lQty As Int32 = 0
                                    If txtQty(lIdx) Is Nothing = False AndAlso txtQty(lIdx).Visible = True Then
                                        Dim sVal As String = txtQty(lIdx).Caption
                                        If sVal Is Nothing Then sVal = "0"
                                        sVal = sVal.Replace(",", "")
                                        sVal = sVal.Replace(".", "")
                                        If IsNumeric(sVal) = True Then lQty = CInt(Val(sVal))
                                    Else
                                        lQty = 1
                                    End If

                                    lResult *= lQty
                                End With
                            End If
                        End If
                    End If
                End If
            End If
        End If

        Return lResult
    End Function

    Private Sub RecalcAll()
        If lblItem Is Nothing = False Then
            For X As Int32 = 0 To lblItem.GetUpperBound(0)
                SetHullAndPower(X)
            Next X
        End If
        RecalcTotals(True)
    End Sub
    Private Sub RecalcTotals(ByVal bIncludeCrew As Boolean)
        RecalcCrew()

        'all this does is go thru the txtHull array summing and the txtPower array summing
        Dim lTotalHull As Int32 = 0
        Dim lTotalPower As Int32 = 0

        If txtHull Is Nothing = False Then
            For X As Int32 = 0 To txtHull.GetUpperBound(0)
                If txtHull(X) Is Nothing = False AndAlso txtHull(X).Caption Is Nothing = False Then
                    Dim sVal As String = txtHull(X).Caption
                    If sVal Is Nothing Then sVal = "0"
                    sVal = sVal.Replace(",", "")
                    sVal = sVal.Replace(".", "")
                    If IsNumeric(sVal) = False Then sVal = "0"

                    Dim lVal As Int32 = CInt(Val(sVal))

                    lTotalHull += lVal
                End If
            Next X
        End If
        If txtPower Is Nothing = False Then
            For X As Int32 = 0 To txtPower.GetUpperBound(0)
                If txtPower(X) Is Nothing = False AndAlso txtPower(X).Caption Is Nothing = False Then
                    Dim sVal As String = txtPower(X).Caption
                    If sVal Is Nothing Then sVal = "0"
                    sVal = sVal.Replace(",", "")
                    sVal = sVal.Replace(".", "")
                    If IsNumeric(sVal) = False Then sVal = "0"

                    Dim lVal As Int32 = CInt(Val(sVal))
                    lTotalPower += lVal
                End If
            Next
        End If

        lblTotalHull.Caption = "Hull Remaining: " & lTotalHull.ToString("#,##0")
        lblTotalPower.Caption = "Power Remaining: " & lTotalPower.ToString("#,##0")

        If lTotalHull < 0 Then lblTotalHull.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0) Else lblTotalHull.ForeColor = muSettings.InterfaceBorderColor
        If lTotalPower < 0 Then lblTotalPower.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0) Else lblTotalPower.ForeColor = muSettings.InterfaceBorderColor

    End Sub
    Private Function GetPositiveHullsOnly() As Int32
        Dim lTotalHull As Int32 = 0
        Dim lTotalPower As Int32 = 0

        If txtHull Is Nothing = False Then
            For X As Int32 = 0 To txtHull.GetUpperBound(0)
                If txtHull(X) Is Nothing = False AndAlso txtHull(X).Caption Is Nothing = False AndAlso IsNumeric(txtHull(X).Caption) = True Then
                    Dim lVal As Int32 = CInt(Val(txtHull(X).Caption))
                    If lVal > 0 Then lTotalHull += lVal
                End If
            Next X
        End If

        Return lTotalHull
    End Function
    Private Sub FillComponents()
        If cboComponent Is Nothing = False Then

            Dim bDoHull As Boolean = True
            If cboComponent Is Nothing = False AndAlso cboComponent.GetUpperBound(0) >= elItemPropIdx.eBaseHull Then
                If cboComponent(elItemPropIdx.eBaseHull) Is Nothing = False Then
                    If cboComponent(elItemPropIdx.eBaseHull).ListIndex > -1 Then bDoHull = False
                End If
            End If

            For X As Int32 = 0 To cboComponent.GetUpperBound(0)
                If cboComponent(X) Is Nothing = False Then
                    If bDoHull = False AndAlso X = elItemPropIdx.eBaseHull Then Continue For
                    cboComponent(X).Clear()
                    cboComponent(X).AddItem("None")
                    cboComponent(X).ItemData(cboComponent(X).NewIndex) = -1
                End If
            Next X

            If goCurrentPlayer Is Nothing = False Then

                'Ok, first, determine if the hull box is set
                Dim yLimitToHullType As Byte = 255
                If cboComponent(elItemPropIdx.eBaseHull) Is Nothing = False Then
                    Dim cbo As UIComboBox = cboComponent(elItemPropIdx.eBaseHull)
                    If cbo Is Nothing = False Then
                        If cbo.ListIndex > -1 Then
                            Dim lHullID As Int32 = cbo.ItemData(cbo.ListIndex)
                            Dim oHull As HullTech = CType(goCurrentPlayer.GetTech(lHullID, ObjectType.eHullTech), HullTech)
                            If oHull Is Nothing = False Then
                                yLimitToHullType = HullTech.GetHullTypeID(oHull.yTypeID, oHull.ySubTypeID)
                            End If
                        End If
                    End If
                End If

                For X As Int32 = 0 To goCurrentPlayer.mlTechUB
                    Dim oTech As Base_Tech = goCurrentPlayer.moTechs(X)
                    If oTech Is Nothing = False Then
                        If oTech.bArchived = True AndAlso mbFilter = True Then Continue For

                        Select Case oTech.ObjTypeID
                            Case ObjectType.eArmorTech
                                Dim cbo As UIComboBox = cboComponent(elItemPropIdx.eArmor)
                                If cbo Is Nothing = False Then
                                    cbo.AddItem(oTech.GetComponentName())
                                    cbo.ItemData(cbo.NewIndex) = oTech.ObjectID
                                    cbo.ItemData2(cbo.NewIndex) = oTech.ObjTypeID
                                End If
                            Case ObjectType.eEngineTech
                                If yLimitToHullType = 255 OrElse CType(oTech, EngineTech).HullTypeID = yLimitToHullType OrElse CType(oTech, EngineTech).HullTypeID = 255 Then
                                    Dim cbo As UIComboBox = cboComponent(elItemPropIdx.eEngine)
                                    If cbo Is Nothing = False Then
                                        cbo.AddItem(oTech.GetComponentName())
                                        cbo.ItemData(cbo.NewIndex) = oTech.ObjectID
                                        cbo.ItemData2(cbo.NewIndex) = oTech.ObjTypeID
                                    End If
                                End If
                            Case ObjectType.eHullTech
                                If bDoHull = False Then Continue For
                                Dim cbo As UIComboBox = cboComponent(elItemPropIdx.eBaseHull)
                                If cbo Is Nothing = False Then
                                    cbo.AddItem(oTech.GetComponentName())
                                    cbo.ItemData(cbo.NewIndex) = oTech.ObjectID
                                    cbo.ItemData2(cbo.NewIndex) = oTech.ObjTypeID
                                End If
                            Case ObjectType.eRadarTech
                                If yLimitToHullType = 255 OrElse CType(oTech, RadarTech).HullTypeID = yLimitToHullType OrElse CType(oTech, RadarTech).HullTypeID = 255 Then
                                    Dim cbo As UIComboBox = cboComponent(elItemPropIdx.eRadar)
                                    If cbo Is Nothing = False Then
                                        cbo.AddItem(oTech.GetComponentName())
                                        cbo.ItemData(cbo.NewIndex) = oTech.ObjectID
                                        cbo.ItemData2(cbo.NewIndex) = oTech.ObjTypeID
                                    End If
                                End If
                            Case ObjectType.eShieldTech
                                If yLimitToHullType = 255 OrElse CType(oTech, ShieldTech).HullTypeID = yLimitToHullType OrElse CType(oTech, ShieldTech).HullTypeID = 255 Then
                                    Dim cbo As UIComboBox = cboComponent(elItemPropIdx.eShield)
                                    If cbo Is Nothing = False Then
                                        cbo.AddItem(oTech.GetComponentName())
                                        cbo.ItemData(cbo.NewIndex) = oTech.ObjectID
                                        cbo.ItemData2(cbo.NewIndex) = oTech.ObjTypeID
                                    End If
                                End If
                            Case ObjectType.eWeaponTech
                                If yLimitToHullType = 255 OrElse CType(oTech, WeaponTech).HullTypeID = yLimitToHullType OrElse CType(oTech, WeaponTech).HullTypeID = 255 Then
                                    For Y As Int32 = elItemPropIdx.eWpn To elItemPropIdx.eWpn + 9
                                        Dim cbo As UIComboBox = cboComponent(Y)
                                        If cbo Is Nothing = False Then
                                            cbo.AddItem(oTech.GetComponentName())
                                            cbo.ItemData(cbo.NewIndex) = oTech.ObjectID
                                            cbo.ItemData2(cbo.NewIndex) = oTech.ObjTypeID
                                        End If
                                    Next Y
                                End If
                        End Select

                    End If
                Next X
            End If
            For X As Int32 = 0 To cboComponent.GetUpperBound(0)
                If cboComponent(X) Is Nothing = False Then cboComponent(X).SortList(False, True)
            Next X
        End If


    End Sub
    Private Sub SetComboIDs()
        If cboComponent Is Nothing = False Then
            If mlComboIDs Is Nothing Then ReDim mlComboIDs(-1)
            If mlComboIDs.GetUpperBound(0) <> cboComponent.GetUpperBound(0) Then
                ReDim mlComboIDs(cboComponent.GetUpperBound(0))
                For X As Int32 = 0 To mlComboIDs.GetUpperBound(0)
                    mlComboIDs(X) = -1
                Next X
            End If
            For X As Int32 = 0 To cboComponent.GetUpperBound(0)
                Dim lID As Int32 = -1
                If cboComponent(X) Is Nothing = False Then
                    If cboComponent(X).ListIndex > -1 Then
                        lID = cboComponent(X).ItemData(cboComponent(X).ListIndex)
                    End If
                End If
                mlComboIDs(X) = lID
            Next X
        Else
            ReDim mlComboIDs(-1)
        End If
    End Sub
    Private Sub SetQtyVals()
        If txtQty Is Nothing = False Then
            If mlQtys Is Nothing Then ReDim mlQtys(-1)
            If mlQtys.GetUpperBound(0) <> txtQty.GetUpperBound(0) Then
                ReDim mlQtys(txtQty.GetUpperBound(0))
                For X As Int32 = 0 To mlQtys.GetUpperBound(0)
                    mlQtys(X) = 0
                Next X
            End If
            For X As Int32 = 0 To txtQty.GetUpperBound(0)
                Dim lVal As Int32 = 0
                If txtQty(X) Is Nothing = False Then
                    Dim sVal As String = txtQty(X).Caption
                    If sVal Is Nothing Then sVal = "0"
                    sVal = sVal.Replace(",", "")
                    sVal = sVal.Replace(".", "")
                    If IsNumeric(sVal) = False Then sVal = "0"

                    lVal = CInt(Val(sVal))
                End If
                mlQtys(X) = lVal
            Next X
        Else
            ReDim mlQtys(-1)
        End If
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnMinMax_Click(ByVal sName As String) Handles btnMinMax.Click
        'Dim rcTemp As System.Drawing.Rectangle
        Dim lNewY As Int32
        Dim lNewWidth As Int32
        Dim lNewHeight As Int32
        Dim X As Int32

        mbMinimized = Not mbMinimized

        If mbMinimized = True Then
            'Ok, we are now minimized, let's set up our values
            mlStartY = Me.Top
            lNewY = mlStartY + Me.Height - btnMinMax.Height - 4
            lNewWidth = 200 '384
            lNewHeight = btnMinMax.Height + 6
            btnClose.Left = lNewWidth - 24 - Me.BorderLineWidth
            btnMinMax.Left = btnClose.Left - 25
            Me.bRoundedBorder = False
        Else
            'ok, we are now maximized, let's set up our values
            'lNewY = mlStartY
            lNewY = Me.Top - 511 + Me.Height - 2
            lNewWidth = 472 '312
            lNewHeight = 511
            btnClose.Left = lNewWidth - 24 - Me.BorderLineWidth
            btnMinMax.Left = btnClose.Left - 25

            Me.bRoundedBorder = True
        End If
        'Now, set the btn's loc and window props
        Me.Top = lNewY
        Me.Width = lNewWidth
        Me.Height = lNewHeight

        If Me.Top + Me.Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then
            Dim lTempVal As Int32 = Me.Top + Me.Height - MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight
            Me.Top -= lTempVal
        End If

        For X = 0 To MyBase.ChildrenUB
            MyBase.moChildren(X).Visible = Not mbMinimized
        Next X
        btnMinMax.Visible = True
        lblTitle.Visible = True
        btnClose.Visible = True

        'Now, if we are minimized, set our new rect accordingly
        With btnMinMax
            If mbMinimized = True Then
                .ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eUpArrow_Button_Normal)
                .ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eUpArrow_Button_Down)
                .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eUpArrow_Button_Disabled)
            Else
                .ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eDownArrow_Button_Normal)
                .ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eDownArrow_Button_Down)
                .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eDownArrow_Button_Disabled)
            End If
        End With

    End Sub

    Private Sub frmBuilderWrksht_OnNewFrame() Handles Me.OnNewFrame
        If glCurrentCycle - mlLastComponentRefresh > 90 Then
            mlLastComponentRefresh = glCurrentCycle

            'Ok, go through our player techs 
            If goCurrentPlayer Is Nothing = False AndAlso cboComponent Is Nothing = False Then
                Try

                    'Ok, get our hull type
                    Dim yLimitToHullType As Byte = 255
                    If cboComponent(elItemPropIdx.eBaseHull) Is Nothing = False Then
                        Dim cbo As UIComboBox = cboComponent(elItemPropIdx.eBaseHull)
                        If cbo Is Nothing = False Then
                            If cbo.ListIndex > -1 Then
                                Dim lHullID As Int32 = cbo.ItemData(cbo.ListIndex)
                                Dim oHull As HullTech = CType(goCurrentPlayer.GetTech(lHullID, ObjectType.eHullTech), HullTech)
                                If oHull Is Nothing = False Then
                                    yLimitToHullType = HullTech.GetHullTypeID(oHull.yTypeID, oHull.ySubTypeID)
                                End If
                            End If
                        End If
                    End If

                    Dim bResortArmor As Boolean = False
                    Dim bResortEngine As Boolean = False
                    Dim bResortHull As Boolean = False
                    Dim bResortRadar As Boolean = False
                    Dim bResortShield As Boolean = False
                    Dim bResortWeapon As Boolean = False

                    For X As Int32 = 0 To goCurrentPlayer.mlTechUB
                        Dim oTech As Base_Tech = goCurrentPlayer.moTechs(X)
                        If oTech Is Nothing = False Then
                            If oTech.bArchived = True Then Continue For

                            Select Case oTech.ObjTypeID
                                Case ObjectType.eArmorTech
                                    Dim cbo As UIComboBox = cboComponent(elItemPropIdx.eArmor)
                                    If cbo Is Nothing = False Then
                                        If cbo.ContainsItemData2(oTech.ObjectID, oTech.ObjTypeID) = False Then
                                            cbo.AddItem(oTech.GetComponentName())
                                            cbo.ItemData(cbo.NewIndex) = oTech.ObjectID
                                            cbo.ItemData2(cbo.NewIndex) = oTech.ObjTypeID
                                            bResortArmor = True
                                        End If
                                    End If
                                Case ObjectType.eEngineTech
                                    If oTech.GetTechHullTypeID <> yLimitToHullType AndAlso yLimitToHullType <> 255 Then Continue For

                                    Dim cbo As UIComboBox = cboComponent(elItemPropIdx.eEngine)
                                    If cbo Is Nothing = False Then
                                        If cbo.ContainsItemData2(oTech.ObjectID, oTech.ObjTypeID) = False Then
                                            cbo.AddItem(oTech.GetComponentName())
                                            cbo.ItemData(cbo.NewIndex) = oTech.ObjectID
                                            cbo.ItemData2(cbo.NewIndex) = oTech.ObjTypeID
                                            bResortEngine = True
                                        End If
                                    End If
                                Case ObjectType.eHullTech
                                    Dim cbo As UIComboBox = cboComponent(elItemPropIdx.eBaseHull)
                                    If cbo Is Nothing = False Then
                                        If cbo.ContainsItemData2(oTech.ObjectID, oTech.ObjTypeID) = False Then
                                            cbo.AddItem(oTech.GetComponentName())
                                            cbo.ItemData(cbo.NewIndex) = oTech.ObjectID
                                            cbo.ItemData2(cbo.NewIndex) = oTech.ObjTypeID
                                            bResortHull = True
                                        End If
                                    End If
                                Case ObjectType.eRadarTech
                                    If oTech.GetTechHullTypeID <> yLimitToHullType AndAlso yLimitToHullType <> 255 Then Continue For

                                    Dim cbo As UIComboBox = cboComponent(elItemPropIdx.eRadar)
                                    If cbo Is Nothing = False Then
                                        If cbo.ContainsItemData2(oTech.ObjectID, oTech.ObjTypeID) = False Then
                                            cbo.AddItem(oTech.GetComponentName())
                                            cbo.ItemData(cbo.NewIndex) = oTech.ObjectID
                                            cbo.ItemData2(cbo.NewIndex) = oTech.ObjTypeID
                                            bResortRadar = True
                                        End If
                                    End If
                                Case ObjectType.eShieldTech
                                    If oTech.GetTechHullTypeID <> yLimitToHullType AndAlso yLimitToHullType <> 255 Then Continue For

                                    Dim cbo As UIComboBox = cboComponent(elItemPropIdx.eShield)
                                    If cbo Is Nothing = False Then
                                        If cbo.ContainsItemData2(oTech.ObjectID, oTech.ObjTypeID) = False Then
                                            cbo.AddItem(oTech.GetComponentName())
                                            cbo.ItemData(cbo.NewIndex) = oTech.ObjectID
                                            cbo.ItemData2(cbo.NewIndex) = oTech.ObjTypeID
                                            bResortShield = True
                                        End If
                                    End If
                                Case ObjectType.eWeaponTech
                                    If oTech.GetTechHullTypeID <> yLimitToHullType AndAlso yLimitToHullType <> 255 Then Continue For

                                    For Y As Int32 = elItemPropIdx.eWpn To elItemPropIdx.eWpn + 9
                                        Dim cbo As UIComboBox = cboComponent(Y)
                                        If cbo Is Nothing = False Then
                                            If cbo.ContainsItemData2(oTech.ObjectID, oTech.ObjTypeID) = False Then
                                                cbo.AddItem(oTech.GetComponentName())
                                                cbo.ItemData(cbo.NewIndex) = oTech.ObjectID
                                                cbo.ItemData2(cbo.NewIndex) = oTech.ObjTypeID
                                                bResortWeapon = True
                                            End If
                                        End If
                                    Next Y
                            End Select
                        End If
                    Next X


                    'Now, check our resorts
                    If bResortArmor = True Then
                        With cboComponent(elItemPropIdx.eArmor)
                            Dim bPrevLoad As Boolean = mbLoading
                            mbLoading = True

                            Dim lTmpID As Int32 = -1
                            If .ListIndex > -1 Then lTmpID = .ItemData(.ListIndex)
                            .SortList(False, True)
                            .FindComboItemData(lTmpID)
                            mbLoading = bPrevLoad
                        End With 
                    End If
                    If bResortEngine = True Then
                        With cboComponent(elItemPropIdx.eEngine)
                            Dim bPrevLoad As Boolean = mbLoading
                            mbLoading = True

                            Dim lTmpID As Int32 = -1
                            If .ListIndex > -1 Then lTmpID = .ItemData(.ListIndex)
                            .SortList(False, True)
                            .FindComboItemData(lTmpID)
                            mbLoading = bPrevLoad
                        End With
                    End If
                    If bResortHull = True Then
                        With cboComponent(elItemPropIdx.eBaseHull)
                            Dim bPrevLoad As Boolean = mbLoading
                            mbLoading = True

                            Dim lTmpID As Int32 = -1
                            If .ListIndex > -1 Then lTmpID = .ItemData(.ListIndex)
                            .SortList(False, True)
                            .FindComboItemData(lTmpID)
                            mbLoading = bPrevLoad
                        End With
                    End If
                    If bResortRadar = True Then
                        With cboComponent(elItemPropIdx.eRadar)
                            Dim bPrevLoad As Boolean = mbLoading
                            mbLoading = True

                            Dim lTmpID As Int32 = -1
                            If .ListIndex > -1 Then lTmpID = .ItemData(.ListIndex)
                            .SortList(False, True)
                            .FindComboItemData(lTmpID)
                            mbLoading = bPrevLoad
                        End With
                    End If
                    If bResortShield = True Then
                        With cboComponent(elItemPropIdx.eShield)
                            Dim bPrevLoad As Boolean = mbLoading
                            mbLoading = True

                            Dim lTmpID As Int32 = -1
                            If .ListIndex > -1 Then lTmpID = .ItemData(.ListIndex)
                            .SortList(False, True)
                            .FindComboItemData(lTmpID)
                            mbLoading = bPrevLoad
                        End With
                    End If
                    If bResortWeapon = True Then
                        For X As Int32 = elItemPropIdx.eWpn To elItemPropIdx.eWpn + 9
                            With cboComponent(X)
                                Dim bPrevLoad As Boolean = mbLoading
                                mbLoading = True

                                Dim lTmpID As Int32 = -1
                                If .ListIndex > -1 Then lTmpID = .ItemData(.ListIndex)
                                .SortList(False, True)
                                .FindComboItemData(lTmpID)
                                mbLoading = bPrevLoad
                            End With
                        Next X
                    End If
                    
                Catch
                End Try

            End If
        End If
    End Sub

    Private Sub frmBuilderWrksht_OnRenderEnd() Handles Me.OnRenderEnd

        Try
            If txtHull Is Nothing = False Then
                For X As Int32 = 0 To txtHull.GetUpperBound(0)
                    If txtHull(X) Is Nothing = False Then
                        txtHull(X).DoNotRender = UITextBox.DoNotRenderSetting.eRenderAll
                        txtHull(X).PrepareForBulkRender()
                    End If
                    If txtPower(X) Is Nothing = False Then
                        txtPower(X).DoNotRender = UITextBox.DoNotRenderSetting.eRenderAll
                        txtPower(X).PrepareForBulkRender()
                    End If
                    If txtQty(X) Is Nothing = False Then
                        txtQty(X).DoNotRender = UITextBox.DoNotRenderSetting.eRenderAll
                        txtQty(X).PrepareForBulkRender()
                    End If
                Next X


                Using oSprite As New Sprite(GFXEngine.moDevice)
                    Try
                        oSprite.Begin(SpriteFlags.AlphaBlend)

                        For X As Int32 = 0 To txtHull.GetUpperBound(0)
                            If txtHull(X) Is Nothing = False Then txtHull(X).BulkBackColorFill(oSprite)
                            If txtPower(X) Is Nothing = False Then txtPower(X).BulkBackColorFill(oSprite)
                            If txtQty(X) Is Nothing = False Then txtQty(X).BulkBackColorFill(oSprite)
                        Next X

                        oSprite.End()
                        oSprite.Dispose()
                    Catch
                    End Try
                End Using

                Using oBLine As New Line(MyBase.moUILib.oDevice)
                    Try
                        With oBLine
                            .Antialias = True
                            .Width = 1
                            .Begin()

                            For X As Int32 = 0 To txtHull.GetUpperBound(0)
                                If txtHull(X) Is Nothing = False Then txtHull(X).BulkRenderBorder(oBLine)
                                If txtPower(X) Is Nothing = False Then txtPower(X).BulkRenderBorder(oBLine)
                                If txtQty(X) Is Nothing = False Then txtQty(X).BulkRenderBorder(oBLine)
                            Next X

                            .End()
                        End With
                    Catch
                    End Try
                End Using

                Dim oFont As Font = BPFont.AcquireDXFont(txtHull(0).GetFont)
                Try
                    For X As Int32 = 0 To txtHull.GetUpperBound(0)
                        If txtHull(X) Is Nothing = False Then
                            txtHull(X).BulkRenderText(oFont)
                            txtHull(X).DoNotRender = UITextBox.DoNotRenderSetting.eDoNotRenderControl
                        End If
                        If txtPower(X) Is Nothing = False Then
                            txtPower(X).BulkRenderText(oFont)
                            txtPower(X).DoNotRender = UITextBox.DoNotRenderSetting.eDoNotRenderControl
                        End If
                        If txtQty(X) Is Nothing = False Then
                            txtQty(X).BulkRenderText(oFont)
                            txtQty(X).DoNotRender = UITextBox.DoNotRenderSetting.eDoNotRenderControl
                        End If
                    Next X
                Catch
                End Try
            End If 
        Catch
        End Try  
    End Sub

    Private Sub LockCrewClick(ByVal lMouseX As Int32, ByVal lMouseY As Int32, ByVal lButton As Windows.Forms.MouseButtons)
        mbCrewLocked = Not mbCrewLocked
        If mbCrewLocked = True Then
            lblLockCrew.BackImageColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
        Else
            lblLockCrew.BackImageColor = muSettings.InterfaceBorderColor
        End If
        Me.IsDirty = True
    End Sub

    Private Sub frmBuilderWrksht_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.BuilderWorksheetY = Me.Top
            muSettings.BuilderWorksheetX = Me.Left
        End If
    End Sub

    Private Sub ViewDesign(ByVal sName As String)
        If HasAliasedRights(AliasingRights.eViewTechDesigns) = False Then
            MyBase.moUILib.AddNotification("You lack rights to view technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If

        Dim lIdx As Int32 = CType(Val(sName.Substring(8)), Int32)
        Dim cbo As UIComboBox = cboComponent(lIdx)
        If cbo Is Nothing = True OrElse cbo.ListIndex < 1 Then Return
        Dim oTech As Base_Tech = goCurrentPlayer.GetTech(cbo.ItemData(cbo.ListIndex), CShort(cbo.ItemData2(cbo.ListIndex)))
        If oTech Is Nothing Then Return

        'TODO: Figure out how to get the prod factor...
        Dim lProdFactor As Int32 = 100

        With oTech
            Select Case .ObjTypeID
                Case ObjectType.eAlloyTech
                    Dim ofrmAlloy As frmAlloyBuilder = New frmAlloyBuilder(MyBase.moUILib)
                    ofrmAlloy.ViewResults(CType(goCurrentPlayer.GetTech(.ObjectID, .ObjTypeID), AlloyTech), lProdFactor)
                    ofrmAlloy = Nothing
                Case ObjectType.eArmorTech
                    Dim ofrmArmor As frmArmorBuilder = New frmArmorBuilder(MyBase.moUILib)
                    ofrmArmor.ViewResults(CType(goCurrentPlayer.GetTech(.ObjectID, .ObjTypeID), ArmorTech), lProdFactor)
                    ofrmArmor = Nothing
                Case ObjectType.eEngineTech
                    Dim ofrmEngine As frmEngineBuilder = New frmEngineBuilder(MyBase.moUILib)
                    ofrmEngine.ViewResults(CType(goCurrentPlayer.GetTech(.ObjectID, .ObjTypeID), EngineTech), lProdFactor)
                    ofrmEngine = Nothing
                Case ObjectType.eHullTech
                    Dim ofrmHull As frmHullBuilder = New frmHullBuilder(MyBase.moUILib)
                    ofrmHull.ViewResults(CType(goCurrentPlayer.GetTech(.ObjectID, .ObjTypeID), HullTech), lProdFactor)
                    ofrmHull = Nothing
                Case ObjectType.ePrototype
                    Dim ofrmPrototype As frmPrototypeBuilder = New frmPrototypeBuilder(MyBase.moUILib)
                    ofrmPrototype.ViewResults(CType(goCurrentPlayer.GetTech(.ObjectID, .ObjTypeID), PrototypeTech), lProdFactor)
                    ofrmPrototype = Nothing
                Case ObjectType.eRadarTech
                    Dim ofrmRadar As frmRadarBuilder = New frmRadarBuilder(MyBase.moUILib)
                    ofrmRadar.ViewResults(CType(goCurrentPlayer.GetTech(.ObjectID, .ObjTypeID), RadarTech), lProdFactor)
                    ofrmRadar = Nothing
                Case ObjectType.eShieldTech
                    Dim ofrmShield As frmShieldBuilder = New frmShieldBuilder(MyBase.moUILib)
                    ofrmShield.ViewResults(CType(goCurrentPlayer.GetTech(.ObjectID, .ObjTypeID), ShieldTech), lProdFactor)
                    ofrmShield = Nothing
                Case ObjectType.eWeaponTech
                    Dim ofrmWeapon As frmWeaponBuilder = New frmWeaponBuilder(MyBase.moUILib)
                    ofrmWeapon.ViewResults(CType(goCurrentPlayer.GetTech(.ObjectID, .ObjTypeID), WeaponTech), lProdFactor)
                    ofrmWeapon = Nothing
            End Select
        End With
    End Sub

    Private Sub chkFilter_Click() Handles chkFilter.Click
        mbFilter = chkFilter.Value
        FillComponents()
    End Sub
End Class