Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmPrototypeBuilder
    Inherits UIWindow

    Private Structure WeaponPlacement
        Public WeaponTechID As Int32
        Public WeaponGroupTypeID As WeaponGroupType
        Public SlotX As Byte
        Public SlotY As Byte
        Public lGroupNum As Int32
    End Structure

    'fraBasicComp
    Private WithEvents fraBasicComp As UIWindow
    Private WithEvents lblHull As UILabel
    Private WithEvents cboHull As UIComboBox
    Private WithEvents lblEngine As UILabel
    Private WithEvents cboEngine As UIComboBox
    Private WithEvents lblArmor As UILabel
    Private WithEvents cboArmor As UIComboBox
    Private WithEvents lnDiv1 As UILine
    Private lblCrewQtrs As UILabel
    Private WithEvents lblRadar As UILabel
    Private WithEvents lblShield As UILabel
    Private WithEvents cboRadar As UIComboBox
    Private WithEvents cboShield As UIComboBox
    Private WithEvents lblArmorAlloc As UILabel

    Private tpFrontArmor As ctlTechProp
    Private tpLeftArmor As ctlTechProp
    Private tpRightArmor As ctlTechProp
    Private tpRearArmor As ctlTechProp
    Private tpProduction As ctlTechProp
    Private tpSecurity As ctlTechProp
    'Private WithEvents lblFrontArmor As UILabel
    'Private WithEvents hscrFront As UIScrollBar
    'Private WithEvents lblLeftArmor As UILabel
    'Private WithEvents hscrLeft As UIScrollBar
    'Private WithEvents lblRearArmor As UILabel
    'Private WithEvents hscrRear As UIScrollBar
    'Private WithEvents lblRightArmor As UILabel
    'Private WithEvents hscrRight As UIScrollBar
    'Private WithEvents txtForeArmor As UITextBox
    'Private WithEvents txtLeftArmor As UITextBox
    'Private WithEvents txtRearArmor As UITextBox
    'Private WithEvents txtRightArmor As UITextBox
    'Private WithEvents lblProduction As UILabel
    'Private WithEvents txtProduction As UITextBox
    'Private WithEvents hscrProduction As UIScrollBar

    'fraWeapons
    Private WithEvents fraWeapons As UIWindow
    Private WithEvents lblType As UILabel
    Private WithEvents lblWeapon As UILabel
    Private WithEvents cboWeaponType As UIComboBox
    Private WithEvents cboWeapon As UIComboBox
    Private WithEvents txtWpnStats As UITextBox
    Private WithEvents btnPlaceWpn As UIButton
    Private WithEvents btnRemoveWpn As UIButton
    Private WithEvents cboWeaponGroupType As UIComboBox
    Private WithEvents lblWeaponGroupType As UILabel
    'Private WithEvents btnRename As UIButton

    'fraPrototypeAttrs
    Private WithEvents fraPrototypeAttrs As UIWindow
    Private WithEvents optExpected As UIOption
    Private WithEvents optActual As UIOption
    Private WithEvents txtAttrs As UITextBox

    'Form's Other Controls...
    Private WithEvents lblName As UILabel
    Private WithEvents txtPrototypeName As UITextBox
    Private WithEvents btnDesign As UIButton
    Private WithEvents btnCancel As UIButton
    Private WithEvents btnDelete As UIButton
    Private WithEvents btnHelp As UIButton
    Private WithEvents hulMain As UIHullSlots
    Private WithEvents chkFilter As UICheckBox
    Private lblSuggestedCrew As UILabel

    'Form's properties...
    Private mbIgnoreScrollEvents As Boolean = True
    Private mlEntityIndex As Int32 = -1
    Private moWpns() As WeaponPlacement
    Private mlWpnUB As Int32 = -1

    Private moTech As PrototypeTech = Nothing
    Private mfrmResCost As frmProdCost = Nothing
    Private mfrmProdCost As frmProdCost = Nothing

    Private mbUnitImmobile As Boolean = False

    Private mbInPlaceWpn As Boolean = False
    Private mbInRemoveWpn As Boolean = False
    Private mbMeetsSuggestedCrew As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmPrototypeBuilder initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.ePrototypeBuilder
            .ControlName = "frmPrototypeBuilder"
            .Left = 10 '(MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - 400
            .Top = 10 '(MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 365
            .Width = 782
            .Height = 730
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .mbAcceptReprocessEvents = True
            .BorderLineWidth = 1
        End With

        '=================================== Prototype Name Box ====================================
        'lblName initial props
        lblName = New UILabel(oUILib)
        With lblName
            .ControlName = "lblName"
            .Left = 10
            .Top = 10
            .Width = 118
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Prototype Name:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblName, UIControl))

        'txtPrototypeName initial props
        txtPrototypeName = New UITextBox(oUILib)
        With txtPrototypeName
            .ControlName = "txtPrototypeName"
            .Left = 140
            .Top = 10
            .Width = 200
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
		Me.AddChild(CType(txtPrototypeName, UIControl))

		'chkFilter initial props
		chkFilter = New UICheckBox(oUILib)
		With chkFilter
			.ControlName = "chkFilter"
			.Left = txtPrototypeName.Left + txtPrototypeName.Width + 10
			.Top = txtPrototypeName.Top
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

        '=================================== fraBasicComps =========================================
        'fraBasicComp initial props
        fraBasicComp = New UIWindow(oUILib)
        With fraBasicComp
            .ControlName = "fraBasicComp"
            .Left = lblName.Left
            .Top = txtPrototypeName.Top + txtPrototypeName.Height + 10
            .Width = 480
            .Height = 190 '160
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Caption = "Select the Basic Components"
            .Moveable = False
            .mbAcceptReprocessEvents = True
            .BorderLineWidth = 1
        End With
        'MSC - 10/08/08 - moved to end to fix combobox dropdowns
        'Me.AddChild(CType(fraBasicComp, UIControl))

        'lblHull initial props
        lblHull = New UILabel(oUILib)
        With lblHull
            .ControlName = "lblHull"
            .Left = 10
            .Top = 15
            .Width = 55
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hull:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraBasicComp.AddChild(CType(lblHull, UIControl))

        'lblCrewQtrs initial props
        lblCrewQtrs = New UILabel(oUILib)
        With lblCrewQtrs
            .ControlName = "lblCrewQtrs"
            .Left = 250
            .Top = 15
            .Width = 250
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Crew Quarters: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Indicates the number of crew quarters required for this design. Crew quarters are general areas of" & vbCrLf & _
                           "access and living and working areas for crew that operate this prototype design. This number is" & vbCrLf & _
                           "deducted from available armor and production slots and is represented in HULL."
        End With
        fraBasicComp.AddChild(CType(lblCrewQtrs, UIControl))

        'lblEngine initial props
        lblEngine = New UILabel(oUILib)
        With lblEngine
            .ControlName = "lblEngine"
            .Left = 10
            .Top = 40
            .Width = 55
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Engine:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraBasicComp.AddChild(CType(lblEngine, UIControl))

        'lblArmor initial props
        lblArmor = New UILabel(oUILib)
        With lblArmor
            .ControlName = "lblArmor"
            .Left = 10
            .Top = 65
            .Width = 55
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Armor:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraBasicComp.AddChild(CType(lblArmor, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 0
            .Top = 90
            .Width = 480
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        fraBasicComp.AddChild(CType(lnDiv1, UIControl))

        ''lblHangar initial props
        'lblHangar = New UILabel(oUILib)
        'With lblHangar
        '    .ControlName = "lblHangar"
        '    .Left = 240
        '    .Top = 15
        '    .Width = 55
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Hangar:"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(4, DrawTextFormat)
        'End With
        'fraBasicComp.AddChild(CType(lblHangar, UIControl))

        'lblRadar initial props
        lblRadar = New UILabel(oUILib)
        With lblRadar
            .ControlName = "lblRadar"
            .Left = 250
            .Top = 40
            .Width = 55
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Radar:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraBasicComp.AddChild(CType(lblRadar, UIControl))

        'lblShield initial props
        lblShield = New UILabel(oUILib)
        With lblShield
            .ControlName = "lblShield"
            .Left = 250
            .Top = 65
            .Width = 55
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Shield:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraBasicComp.AddChild(CType(lblShield, UIControl))

        'lblArmorAlloc initial props
        lblArmorAlloc = New UILabel(oUILib)
        With lblArmorAlloc
            .ControlName = "lblArmorAlloc"
            .Left = 10
            .Top = 93
            .Width = 120
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Armor Allocation:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraBasicComp.AddChild(CType(lblArmorAlloc, UIControl))

        '     'lblFrontArmor initial props
        '     lblFrontArmor = New UILabel(oUILib)
        '     With lblFrontArmor
        '         .ControlName = "lblFrontArmor"
        '         .Left = 20
        '         .Top = 110
        '         .Width = 36 '99
        '         .Height = 18
        '         .Enabled = True
        '         .Visible = True
        '         .Caption = "Front"
        '         .ForeColor = muSettings.InterfaceBorderColor
        '         .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '         .DrawBackImage = False
        '         .FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
        '     End With
        '     fraBasicComp.AddChild(CType(lblFrontArmor, UIControl))

        '     'txtForeArmor initial props
        '     txtForeArmor = New UITextBox(oUILib)
        '     With txtForeArmor
        '         .ControlName = "txtForeArmor"
        '         .Left = 60
        '         .Top = 110
        '         .Width = 83
        '         .Height = 18
        '         .Enabled = True
        '         .Visible = True
        '         .Caption = "0"
        '         .ForeColor = muSettings.InterfaceTextBoxForeColor
        '         .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '         .DrawBackImage = False
        '         .FontFormat = CType(4, DrawTextFormat)
        '         .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '         .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '         .MaxLength = 10
        '         .BorderColor = muSettings.InterfaceBorderColor
        '         .bNumericOnly = True
        '     End With
        '     fraBasicComp.AddChild(CType(txtForeArmor, UIControl))

        '     'hscrFront initial props
        '     hscrFront = New UIScrollBar(oUILib, False)
        '     With hscrFront
        '         .ControlName = "hscrFront"
        '         .Left = 145
        '         .Top = 111
        '         .Width = 100
        '         .Height = 18
        '         .Enabled = True
        '         .Visible = True
        '         .Value = 0
        '         .MaxValue = 0
        '         .MinValue = 0
        '         .SmallChange = 1
        '         .LargeChange = 10
        '         .ReverseDirection = False
        '         .mbAcceptReprocessEvents = True
        '     End With
        '     fraBasicComp.AddChild(CType(hscrFront, UIControl))
        tpFrontArmor = New ctlTechProp(oUILib)
        With tpFrontArmor
            .ControlName = "tpFrontArmor"
            .Enabled = True
            .Height = 18
            .Left = 22
            .MaxValue = 1000000
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Front:"
            .PropertyValue = 0
            .Top = 118
            .Visible = True
            .Width = 200
            .SetExtendedProps(False, 36, 70, False)
        End With
        fraBasicComp.AddChild(CType(tpFrontArmor, UIControl))

        '     'lblLeftArmor initial props
        '     lblLeftArmor = New UILabel(oUILib)
        '     With lblLeftArmor
        '         .ControlName = "lblLeftArmor"
        '         .Left = 20
        '         .Top = 135
        '         .Width = 36
        '         .Height = 18
        '         .Enabled = True
        '         .Visible = True
        '         .Caption = "Left"
        '         .ForeColor = muSettings.InterfaceBorderColor
        '         .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '         .DrawBackImage = False
        '         .FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
        '     End With
        '     fraBasicComp.AddChild(CType(lblLeftArmor, UIControl))

        '     'txtLeftArmor initial props
        '     txtLeftArmor = New UITextBox(oUILib)
        '     With txtLeftArmor
        '         .ControlName = "txtLeftArmor"
        '         .Left = 60
        '         .Top = 135
        '         .Width = 83
        '         .Height = 18
        '         .Enabled = True
        '         .Visible = True
        '         .Caption = "0"
        '         .ForeColor = muSettings.InterfaceTextBoxForeColor
        '         .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '         .DrawBackImage = False
        '         .FontFormat = CType(4, DrawTextFormat)
        '         .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '         .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '         .MaxLength = 10
        '         .BorderColor = muSettings.InterfaceBorderColor
        '         .bNumericOnly = True
        '     End With
        '     fraBasicComp.AddChild(CType(txtLeftArmor, UIControl))

        '     'hscrLeft initial props
        '     hscrLeft = New UIScrollBar(oUILib, False)
        '     With hscrLeft
        '         .ControlName = "hscrLeft"
        '         .Left = 145
        '         .Top = 136
        '         .Width = 100
        '         .Height = 18
        '         .Enabled = True
        '         .Visible = True
        '         .Value = 0
        '         .MaxValue = 0
        '         .MinValue = 0
        '         .SmallChange = 1
        '         .LargeChange = 10
        '         .ReverseDirection = False
        '         .mbAcceptReprocessEvents = True
        '     End With
        '     fraBasicComp.AddChild(CType(hscrLeft, UIControl))
        tpLeftArmor = New ctlTechProp(oUILib)
        With tpLeftArmor
            .ControlName = "tpLeftArmor"
            .Enabled = True
            .Height = 18
            .Left = 22
            .MaxValue = 1000000
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Left:"
            .PropertyValue = 0
            .Top = 141
            .Visible = True
            .Width = 200
            .SetExtendedProps(False, 36, 70, False)
        End With
        fraBasicComp.AddChild(CType(tpLeftArmor, UIControl))

        '     'lblRearArmor initial props
        '     lblRearArmor = New UILabel(oUILib)
        '     With lblRearArmor
        '         .ControlName = "lblRearArmor"
        '         .Left = 250
        '         .Top = 110
        '         .Width = 36
        '         .Height = 18
        '         .Enabled = True
        '         .Visible = True
        '         .Caption = "Rear"
        '         .ForeColor = muSettings.InterfaceBorderColor
        '         .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '         .DrawBackImage = False
        '         .FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
        '     End With
        '     fraBasicComp.AddChild(CType(lblRearArmor, UIControl))

        '     'txtRearArmor initial props
        '     txtRearArmor = New UITextBox(oUILib)
        '     With txtRearArmor
        '         .ControlName = "txtRearArmor"
        '         .Left = 290
        '         .Top = 110
        '         .Width = 83
        '         .Height = 18
        '         .Enabled = True
        '         .Visible = True
        '         .Caption = "0"
        '         .ForeColor = muSettings.InterfaceTextBoxForeColor
        '         .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '         .DrawBackImage = False
        '         .FontFormat = CType(4, DrawTextFormat)
        '         .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '         .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '         .MaxLength = 10
        '         .BorderColor = muSettings.InterfaceBorderColor
        '         .bNumericOnly = True
        '     End With
        '     fraBasicComp.AddChild(CType(txtRearArmor, UIControl))

        '     'hscrRear initial props
        '     hscrRear = New UIScrollBar(oUILib, False)
        '     With hscrRear
        '         .ControlName = "hscrRear"
        '         .Left = 375
        '         .Top = 111
        '         .Width = 100
        '         .Height = 18
        '         .Enabled = True
        '         .Visible = True
        '         .Value = 0
        '         .MaxValue = 0
        '         .MinValue = 0
        '         .SmallChange = 1
        '         .LargeChange = 10
        '         .ReverseDirection = False
        '         .mbAcceptReprocessEvents = True
        '     End With
        '     fraBasicComp.AddChild(CType(hscrRear, UIControl))
        tpRearArmor = New ctlTechProp(oUILib)
        With tpRearArmor
            .ControlName = "tpRearArmor"
            .Enabled = True
            .Height = 18
            .Left = 250
            .MaxValue = 1000000
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Rear:"
            .PropertyValue = 0
            .Top = 118
            .Visible = True
            .Width = 200
            .SetExtendedProps(False, 36, 70, False)
        End With
        fraBasicComp.AddChild(CType(tpRearArmor, UIControl))

        '     'lblRightArmor initial props
        '     lblRightArmor = New UILabel(oUILib)
        '     With lblRightArmor
        '         .ControlName = "lblRightArmor"
        '         .Left = 250
        '         .Top = 135
        '         .Width = 36
        '         .Height = 18
        '         .Enabled = True
        '         .Visible = True
        '         .Caption = "Right"
        '         .ForeColor = muSettings.InterfaceBorderColor
        '         .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '         .DrawBackImage = False
        '         .FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
        '     End With
        '     fraBasicComp.AddChild(CType(lblRightArmor, UIControl))

        '     'txtRightArmor initial props
        '     txtRightArmor = New UITextBox(oUILib)
        '     With txtRightArmor
        '         .ControlName = "txtRightArmor"
        '         .Left = 290
        '         .Top = 135
        '         .Width = 83
        '         .Height = 18
        '         .Enabled = True
        '         .Visible = True
        '         .Caption = "0"
        '         .ForeColor = muSettings.InterfaceTextBoxForeColor
        '         .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '         .DrawBackImage = False
        '         .FontFormat = CType(4, DrawTextFormat)
        '         .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '         .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '         .MaxLength = 10
        '         .BorderColor = muSettings.InterfaceBorderColor
        '         .bNumericOnly = True
        '     End With
        '     fraBasicComp.AddChild(CType(txtRightArmor, UIControl))

        '     'hscrRight initial props
        '     hscrRight = New UIScrollBar(oUILib, False)
        '     With hscrRight
        '         .ControlName = "hscrRight"
        '         .Left = 375
        '         .Top = 135
        '         .Width = 100
        '         .Height = 18
        '         .Enabled = True
        '         .Visible = True
        '         .Value = 0
        '         .MaxValue = 0
        '         .MinValue = 0
        '         .SmallChange = 1
        '         .LargeChange = 10
        '         .ReverseDirection = False
        '         .mbAcceptReprocessEvents = True
        '     End With
        '     fraBasicComp.AddChild(CType(hscrRight, UIControl))
        tpRightArmor = New ctlTechProp(oUILib)
        With tpRightArmor
            .ControlName = "tpRightArmor"
            .Enabled = True
            .Height = 18
            .Left = 250
            .MaxValue = 1000000
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Right:"
            .PropertyValue = 0
            .Top = 141
            .Visible = True
            .Width = 200
            .SetExtendedProps(False, 36, 70, False)
        End With
        fraBasicComp.AddChild(CType(tpRightArmor, UIControl))

        '     'lblProduction initial props
        '     lblProduction = New UILabel(oUILib)
        '     With lblProduction
        '         .ControlName = "lblProduction"
        '         .Left = 20
        '         .Top = 160
        '.Width = 120
        '         .Height = 18
        '         .Enabled = True
        '         .Visible = True
        '         .Caption = "Production:"
        '         .ForeColor = muSettings.InterfaceBorderColor
        '         .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '         .DrawBackImage = False
        '         .FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
        '         .ToolTipText = "Indicates to use Armor Slot space for Production Slots. Directly impacts how well this item will produce."
        '     End With
        '     fraBasicComp.AddChild(CType(lblProduction, UIControl))

        '     'txtProduction initial props
        '     txtProduction = New UITextBox(oUILib)
        '     With txtProduction
        '         .ControlName = "txtProduction"
        '.Left = 145
        '         .Top = 160
        '         .Width = 83
        '         .Height = 18
        '         .Enabled = True
        '         .Visible = True
        '         .Caption = "0"
        '         .ForeColor = muSettings.InterfaceTextBoxForeColor
        '         .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '         .DrawBackImage = False
        '         .FontFormat = CType(4, DrawTextFormat)
        '         .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '         .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '         .MaxLength = 10
        '         .BorderColor = muSettings.InterfaceBorderColor
        '         .bNumericOnly = True
        '     End With
        '     fraBasicComp.AddChild(CType(txtProduction, UIControl))

        '     'hscrProduction initial props
        '     hscrProduction = New UIScrollBar(oUILib, False)
        '     With hscrProduction
        '         .ControlName = "hscrProduction"
        '.Left = 230
        '         .Top = 160
        '         .Width = 100
        '         .Height = 18
        '         .Enabled = True
        '         .Visible = True
        '         .Value = 0
        '         .MaxValue = 0
        '         .MinValue = 0
        '         .SmallChange = 1
        '         .LargeChange = 10
        '         .ReverseDirection = False
        '         .mbAcceptReprocessEvents = True
        '     End With
        '     fraBasicComp.AddChild(CType(hscrProduction, UIControl))
        tpProduction = New ctlTechProp(oUILib)
        With tpProduction
            .ControlName = "tpProduction"
            .Enabled = True
            .Height = 18
            .Left = 20
            .MaxValue = 1000000
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Production:"
            .PropertyValue = 0
            .Top = 165
            .Visible = True
            .Width = 430
            .SetExtendedProps(False, 120, 70, False)
            .ToolTipText = "Indicates to use Armor Slot space for Production Slots. Directly impacts how well this item will produce."
        End With
        fraBasicComp.AddChild(CType(tpProduction, UIControl))

        tpSecurity = New ctlTechProp(oUILib)
        With tpSecurity
            .ControlName = "tpSecurity"
            .Enabled = True
            .Height = 18
            .Left = 236 '250
            .MaxValue = 10000
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Security:"
            .PropertyValue = 0
            .Top = 95
            .Visible = True
            .Width = 214
            .SetExtendedProps(False, 50, 70, False)
            .ToolTipText = "Additional crew placed on board to help defend against boarding parties, pirate raids and agent missions."
        End With
        fraBasicComp.AddChild(CType(tpSecurity, UIControl))

        '=========================================== fraWeapons Attrs ======================================
        'fraWeapons initial props
        fraWeapons = New UIWindow(oUILib)
        With fraWeapons
            .ControlName = "fraWeapons"
            .Left = 495
            .Top = 10
            .Width = 267
            .Height = 200 '(fraBasicComp.Top + fraBasicComp.Height) - .Top
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Caption = "Weapons"
            .Moveable = False
            .BorderLineWidth = 1
            .mbAcceptReprocessEvents = True
        End With
        'MSC - 10/08/08 - moved fraWeapons to end so the comboboxes render properly
        'Me.AddChild(CType(fraWeapons, UIControl))

        'lblType initial props
        lblType = New UILabel(oUILib)
        With lblType
            .ControlName = "lblType"
            .Left = 15
            .Top = 10
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Type:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraWeapons.AddChild(CType(lblType, UIControl))

        'lblWeapon initial props
        lblWeapon = New UILabel(oUILib)
        With lblWeapon
            .ControlName = "lblWeapon"
            .Left = 15
            .Top = 35
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Weapon:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraWeapons.AddChild(CType(lblWeapon, UIControl))

        'lblWeaponGroupType initial props
        lblWeaponGroupType = New UILabel(oUILib)
        With lblWeaponGroupType
            .ControlName = "lblWeaponGroupType"
            .Left = 15
            .Top = 60
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Group:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraWeapons.AddChild(CType(lblWeaponGroupType, UIControl))

        'lstWpnStats initial props
        txtWpnStats = New UITextBox(oUILib)
        With txtWpnStats
            .ControlName = "txtWpnStats"
            .Left = 10
            .Top = 85
            .Width = 250
            .Height = 81
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .MultiLine = True
        End With
        fraWeapons.AddChild(CType(txtWpnStats, UIControl))

        'btnPlaceWpn initial props
        btnPlaceWpn = New UIButton(oUILib)
        With btnPlaceWpn
            .ControlName = "btnPlaceWpn"
            .Left = txtWpnStats.Left
            .Top = fraWeapons.Height - 30
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Place"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        fraWeapons.AddChild(CType(btnPlaceWpn, UIControl))

        'btnRemoveWpn initial props
        btnRemoveWpn = New UIButton(oUILib)
        With btnRemoveWpn
            .ControlName = "btnRemoveWpn"
            .Left = (txtWpnStats.Left + txtWpnStats.Width) - 100
            .Top = btnPlaceWpn.Top
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Remove"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        fraWeapons.AddChild(CType(btnRemoveWpn, UIControl))

        '=================================== fraPrototypeAttrs =====================================
        'fraPrototypeAttrs initial props
        fraPrototypeAttrs = New UIWindow(oUILib)
        With fraPrototypeAttrs
            .ControlName = "fraPrototypeAttrs"
            .Left = fraWeapons.Left
            .Top = fraWeapons.Top + fraWeapons.Height + 10 'fraBasicComp.Top + fraBasicComp.Height + 5
            .Width = 266
            '.Height = (hulMain.Top + hulMain.Height) - .Top
            .Height = (fraBasicComp.Top + fraBasicComp.Height + 455) - .Top
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Caption = "Prototype Attributes"
            .Moveable = False
            .BorderLineWidth = 1
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(fraPrototypeAttrs, UIControl))

        'optExpected initial props
        optExpected = New UIOption(oUILib)
        With optExpected
            .ControlName = "optExpected"
            .Left = 10
            .Top = 10
            .Width = 120
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Expected Results"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
            .DisplayType = UIOption.eOptionTypes.eSmallOption
        End With
        fraPrototypeAttrs.AddChild(CType(optExpected, UIControl))

        'optActual initial props
        optActual = New UIOption(oUILib)
        With optActual
            .ControlName = "optActual"
            .Left = 150
            .Top = 10
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Actual Results"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UIOption.eOptionTypes.eSmallOption
        End With
        fraPrototypeAttrs.AddChild(CType(optActual, UIControl))

        'lstAttrs initial props
        txtAttrs = New UITextBox(oUILib)
        With txtAttrs
            .ControlName = "txtAttrs"
            .Left = 5
            .Top = 32
            .Width = 256
            .Height = fraPrototypeAttrs.Height - .Top - 5
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Left Or DrawTextFormat.Top
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .MultiLine = True
            .Locked = True
            .mbAcceptReprocessEvents = True
        End With
        fraPrototypeAttrs.AddChild(CType(txtAttrs, UIControl))

        '=========================================== frmPrototypeBuilder Controls =======================================
        'btnHelp initial props
        btnHelp = New UIButton(oUILib)
        With btnHelp
            .ControlName = "btnHelp"
            .Left = fraWeapons.Left - 26
            .Top = Me.BorderLineWidth
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "?"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 14.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
            .ToolTipText = "Click to begin the tutorial for this window."
        End With
        Me.AddChild(CType(btnHelp, UIControl))

        'btnDesign initial props
        btnDesign = New UIButton(oUILib)
        With btnDesign
            .ControlName = "btnDesign"
            .Left = fraWeapons.Left
            .Top = fraPrototypeAttrs.Top + fraPrototypeAttrs.Height + 10
            .Width = 80
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Submit"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnDesign, UIControl))

        'btnCancel initial props
        btnCancel = New UIButton(oUILib)
        With btnCancel
            .ControlName = "btnCancel"
            .Left = fraWeapons.Left + fraWeapons.Width - 70
            .Top = btnDesign.Top
            .Width = 80
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Cancel"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnCancel, UIControl))

        'btnDelete initial props
        btnDelete = New UIButton(oUILib)
        With btnDelete
            .ControlName = "btnDelete"
            Dim lTemp As Int32 = btnCancel.Left - (btnDesign.Left + btnDesign.Width)
            lTemp \= 2
            lTemp -= 40
            .Left = btnDesign.Left + btnDesign.Width + lTemp
            .Top = btnCancel.Top
            .Width = 80
            .Height = 24
            .Enabled = True
            .Visible = False
            .Caption = "Delete"
            .ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnDelete, UIControl))

        'hulMain initial props
        hulMain = New UIHullSlots(oUILib)
        With hulMain
            .AutoRefresh = False
            .ControlName = "hulMain"
            .Enabled = True
            .Height = 480
            .Left = fraBasicComp.Left
            .Top = fraBasicComp.Top + fraBasicComp.Height + 5
            .Visible = True
            .Width = 480
            .AutoRefresh = True
        End With
        Me.AddChild(CType(hulMain, UIControl))

        lblSuggestedCrew = New UILabel(oUILib)
        With lblSuggestedCrew
            .ControlName = "lblSuggestedCrew"
            .Left = 13
            .Top = hulMain.Top + 3
            .Width = 300
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Suggested Crew: 0"
            .ForeColor = System.Drawing.Color.FromArgb(192, muSettings.InterfaceBorderColor.R, muSettings.InterfaceBorderColor.G, muSettings.InterfaceBorderColor.B)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSuggestedCrew, UIControl))

        Me.AddChild(CType(fraWeapons, UIControl))
        Me.AddChild(CType(fraBasicComp, UIControl))




        '=================================== COMBO BOXES FOR ALL FRM FRA =================================
        'cboShield initial props
        cboShield = New UIComboBox(oUILib)
        With cboShield
            .ControlName = "cboShield"
            .Left = 295
            .Top = 65
            .Width = 180
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
            .l_ListBoxHeight = Me.Height \ 2
            .l_ListBoxWidth = 180
        End With
        fraBasicComp.AddChild(CType(cboShield, UIControl))

        'cboRadar initial props
        cboRadar = New UIComboBox(oUILib)
        With cboRadar
            .ControlName = "cboRadar"
            .Left = 295
            .Top = 40
            .Width = 180
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
            .l_ListBoxHeight = Me.Height \ 2
            .l_ListBoxWidth = 180
        End With
        fraBasicComp.AddChild(CType(cboRadar, UIControl))

        ''cboHangar initial props
        'cboHangar = New UIComboBox(oUILib)
        'With cboHangar
        '    .ControlName = "cboHangar"
        '    .Left = 295
        '    .Top = 15
        '    .Width = 150
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        '    .DropDownListBorderColor = muSettings.InterfaceBorderColor
        'End With
        'fraBasicComp.AddChild(CType(cboHangar, UIControl))

        'cboArmor initial props
        cboArmor = New UIComboBox(oUILib)
        With cboArmor
            .ControlName = "cboArmor"
            .Left = 65
            .Top = 65
            .Width = 180
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
            .l_ListBoxHeight = Me.Height \ 2
            .l_ListBoxWidth = 180
        End With
        fraBasicComp.AddChild(CType(cboArmor, UIControl))

        'cboEngine initial props
        cboEngine = New UIComboBox(oUILib)
        With cboEngine
            .ControlName = "cboEngine"
            .Left = 65
            .Top = 40
            .Width = 180
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
            .l_ListBoxHeight = Me.Height \ 2
            .l_ListBoxWidth = 180
        End With
        fraBasicComp.AddChild(CType(cboEngine, UIControl))

        'cboHull initial props
        cboHull = New UIComboBox(oUILib)
        With cboHull
            .ControlName = "cboHull"
            .Left = 65
            .Top = 15
            .Width = 180
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
            .l_ListBoxHeight = Me.Height \ 2
            .l_ListBoxWidth = 180
        End With
        fraBasicComp.AddChild(CType(cboHull, UIControl))

        'cboWeaponGroupType initial props
        cboWeaponGroupType = New UIComboBox(oUILib)
        With cboWeaponGroupType
            .ControlName = "cboWeaponGroupType"
            .Left = 80
            .Top = 60
            .Width = 180
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
            .l_ListBoxHeight = 115
        End With
        fraWeapons.AddChild(CType(cboWeaponGroupType, UIControl))

        'cboWeapon initial props
        cboWeapon = New UIComboBox(oUILib)
        With cboWeapon
            .ControlName = "cboWeapon"
            .Left = 80
            .Top = 35
            .Width = 180
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
            .l_ListBoxHeight = Me.Height \ 2
        End With
        fraWeapons.AddChild(CType(cboWeapon, UIControl))

        'cboWeaponType initial props
        cboWeaponType = New UIComboBox(oUILib)
        With cboWeaponType
            .ControlName = "cboWeaponType"
            .Left = 80
            .Top = 10
            .Width = 180
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
        End With
        fraWeapons.AddChild(CType(cboWeaponType, UIControl))

        AddHandler tpFrontArmor.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpLeftArmor.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpRightArmor.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpRearArmor.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpProduction.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpSecurity.PropertyValueChanged, AddressOf tp_ValueChanged

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        glCurrentEnvirView = CurrentView.ePrototypeBuilder

        If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
        goFullScreenBackground = Nothing

        'Dim oTmpWin As UIWindow = MyBase.moUILib.GetWindow("frmChat")
        'If oTmpWin Is Nothing = False Then oTmpWin.Visible = False
        'oTmpWin = Nothing

        mbIgnoreScrollEvents = False

        'FillComponentLists()
        FillHullList()
        RefreshAttributeBox()
    End Sub

    Private Sub FillHullList()
		cboHull.Clear()

		Dim bFilter As Boolean = chkFilter.Value

        For X As Int32 = 0 To goCurrentPlayer.mlTechUB
			If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
				If goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eHullTech AndAlso (goCurrentPlayer.moTechs(X).bArchived = False OrElse bFilter = False) Then
					cboHull.AddItem(goCurrentPlayer.moTechs(X).GetComponentName)
					cboHull.ItemData(cboHull.NewIndex) = goCurrentPlayer.moTechs(X).ObjectID
					cboHull.ItemData2(cboHull.NewIndex) = ObjectType.eHullTech
				End If
			End If
        Next X

    End Sub

	Private Sub FillComponentLists()
		If goCurrentPlayer Is Nothing Then Return

		'Clear our lists out... (in case they have something already)
		cboArmor.Clear() : cboEngine.Clear() ': cboHangar.Clear()
		cboRadar.Clear() : cboShield.Clear() : cboWeapon.Clear() : cboWeaponType.Clear()
		cboWeaponType.Clear() : cboWeaponGroupType.Clear()

		If cboHull.ListIndex > -1 Then

			Dim oHullTech As HullTech = CType(goCurrentPlayer.GetTech(cboHull.ItemData(cboHull.ListIndex), ObjectType.eHullTech), HullTech)

			If oHullTech Is Nothing = False Then

				Dim fEngineCap As Single = oHullTech.GetHullAllotment(SlotType.eNoChange, SlotConfig.eEngines)
				Dim fHangarCap As Single = oHullTech.GetHullAllotment(SlotType.eNoChange, SlotConfig.eHangar)
				Dim fRadarCap As Single = oHullTech.GetHullAllotment(SlotType.eNoChange, SlotConfig.eRadar)
				Dim fShieldCap As Single = oHullTech.GetHullAllotment(SlotType.eNoChange, SlotConfig.eShields)

                Dim yHullTypeID As Byte = HullTech.GetHullTypeID(oHullTech.yTypeID, oHullTech.ySubTypeID)

				'Add a none to each item...
				cboArmor.AddItem("None")
				cboArmor.ItemData(cboArmor.NewIndex) = 0
				cboEngine.AddItem("None")
				cboEngine.ItemData(cboEngine.NewIndex) = 0
				'cboHangar.AddItem("None")
				'cboHangar.ItemData(cboHangar.NewIndex) = 0
				cboRadar.AddItem("None")
				cboRadar.ItemData(cboRadar.NewIndex) = 0
				cboShield.AddItem("None")
				cboShield.ItemData(cboShield.NewIndex) = 0

				Dim bFilter As Boolean = chkFilter.Value

				For X As Int32 = 0 To goCurrentPlayer.mlTechUB
					If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
                        If goCurrentPlayer.moTechs(X).bArchived = False OrElse bFilter = False Then
                            Select Case goCurrentPlayer.moTechs(X).ObjTypeID
                                Case ObjectType.eArmorTech
                                    cboArmor.AddItem(goCurrentPlayer.moTechs(X).GetComponentName)
                                    cboArmor.ItemData(cboArmor.NewIndex) = goCurrentPlayer.moTechs(X).ObjectID
                                    cboArmor.ItemData2(cboArmor.NewIndex) = ObjectType.eArmorTech
                                Case ObjectType.eEngineTech
                                    With CType(goCurrentPlayer.moTechs(X), EngineTech)
                                        If .HullRequired <= fEngineCap AndAlso (.HullTypeID = yHullTypeID OrElse .HullTypeID = 255) Then
                                            'facilities cannot accept engines with thrust,maneuver, or maxspeed
                                            If oHullTech.yTypeID <> 2 OrElse (.Maneuver = 0 AndAlso .Speed = 0 AndAlso .Thrust = 0) Then
                                                cboEngine.AddItem(.sEngineName)
                                                cboEngine.ItemData(cboEngine.NewIndex) = .ObjectID
                                                cboEngine.ItemData2(cboEngine.NewIndex) = ObjectType.eEngineTech
                                            End If
                                        End If
                                    End With

                                    'Case ObjectType.eHangarTech
                                    '    If CType(goCurrentPlayer.moTechs(X), HangarTech).HullRequired < lHangarCap Then
                                    '        cboHangar.AddItem(goCurrentPlayer.moTechs(X).GetComponentName)
                                    '        cboHangar.ItemData(cboHangar.NewIndex) = goCurrentPlayer.moTechs(X).ObjectID
                                    '    End If
                                Case ObjectType.eRadarTech
                                    With CType(goCurrentPlayer.moTechs(X), RadarTech)
                                        If .HullRequired <= fRadarCap AndAlso (.HullTypeID = yHullTypeID OrElse .HullTypeID = 255) Then
                                            cboRadar.AddItem(.sRadarName)
                                            cboRadar.ItemData(cboRadar.NewIndex) = goCurrentPlayer.moTechs(X).ObjectID
                                            cboRadar.ItemData2(cboRadar.NewIndex) = ObjectType.eRadarTech
                                        End If
                                    End With
                                Case ObjectType.eShieldTech
                                    With CType(goCurrentPlayer.moTechs(X), ShieldTech)
                                        If .HullRequired <= fShieldCap AndAlso (.HullTypeID = yHullTypeID OrElse .HullTypeID = 255) Then
                                            If .lProjectionHullSize >= oHullTech.HullSize Then
                                                cboShield.AddItem(.sShieldName)
                                                cboShield.ItemData(cboShield.NewIndex) = goCurrentPlayer.moTechs(X).ObjectID
                                                cboShield.ItemData2(cboShield.NewIndex) = ObjectType.eShieldTech
                                            End If
                                        End If
                                    End With
                            End Select
                        End If
					End If
				Next X

				'Now, for player knowledge techs
				For X As Int32 = 0 To glPlayerTechKnowledgeUB
					If glPlayerTechKnowledgeIdx(X) <> -1 AndAlso goPlayerTechKnowledge(X).yKnowledgeType > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
						If goPlayerTechKnowledge(X).bArchived = False OrElse bFilter = False Then
							Dim oTech As Base_Tech = goPlayerTechKnowledge(X).oTech
							Select Case oTech.ObjTypeID
								Case ObjectType.eArmorTech
									cboArmor.AddItem(oTech.GetComponentName)
									cboArmor.ItemData(cboArmor.NewIndex) = oTech.ObjectID
									cboArmor.ItemData2(cboArmor.NewIndex) = ObjectType.eArmorTech
                                Case ObjectType.eEngineTech
                                    With CType(oTech, EngineTech)
                                        If .HullRequired < fEngineCap AndAlso (.HullTypeID = yHullTypeID OrElse .HullTypeID = 255) Then
                                            cboEngine.AddItem(oTech.GetComponentName)
                                            cboEngine.ItemData(cboEngine.NewIndex) = oTech.ObjectID
                                            cboEngine.ItemData2(cboEngine.NewIndex) = ObjectType.eEngineTech
                                        End If
                                    End With
                                Case ObjectType.eRadarTech
                                    With CType(oTech, RadarTech)
                                        If .HullRequired < fRadarCap AndAlso (.HullTypeID = yHullTypeID OrElse .HullTypeID = 255) Then
                                            cboRadar.AddItem(oTech.GetComponentName)
                                            cboRadar.ItemData(cboRadar.NewIndex) = oTech.ObjectID
                                            cboRadar.ItemData2(cboRadar.NewIndex) = ObjectType.eRadarTech
                                        End If
                                    End With
                                Case ObjectType.eShieldTech
                                    With CType(oTech, ShieldTech)
                                        If .HullRequired < fShieldCap AndAlso (.HullTypeID = yHullTypeID OrElse .HullTypeID = 255) Then
                                            If .lProjectionHullSize >= oHullTech.HullSize Then
                                                cboShield.AddItem(oTech.GetComponentName)
                                                cboShield.ItemData(cboShield.NewIndex) = oTech.ObjectID
                                                cboShield.ItemData2(cboShield.NewIndex) = ObjectType.eShieldTech
                                            End If
                                        End If
                                    End With
                            End Select
						End If
					End If
				Next X

				'Fill in our weapontype... (in alphabetical)
				cboWeapon.Clear()
				cboWeaponType.Clear()
				cboWeaponType.AddItem("Bomb")
				cboWeaponType.ItemData(cboWeaponType.NewIndex) = WeaponClassType.eBomb
				cboWeaponType.AddItem("Energy/Beam")
				cboWeaponType.ItemData(cboWeaponType.NewIndex) = WeaponClassType.eEnergyBeam
				cboWeaponType.AddItem("Energy/Pulse")
				cboWeaponType.ItemData(cboWeaponType.NewIndex) = WeaponClassType.eEnergyPulse
                'cboWeaponType.AddItem("Mine")
                'cboWeaponType.ItemData(cboWeaponType.NewIndex) = WeaponClassType.eMine
				cboWeaponType.AddItem("Missile")
				cboWeaponType.ItemData(cboWeaponType.NewIndex) = WeaponClassType.eMissile
				cboWeaponType.AddItem("Projectile")
				cboWeaponType.ItemData(cboWeaponType.NewIndex) = WeaponClassType.eProjectile

				'Fill in our WeaponGroupType...
				cboWeaponGroupType.Clear()
				cboWeaponGroupType.AddItem("Primary")
				cboWeaponGroupType.ItemData(cboWeaponGroupType.NewIndex) = WeaponGroupType.PrimaryWeaponGroup
				cboWeaponGroupType.AddItem("Secondary")
				cboWeaponGroupType.ItemData(cboWeaponGroupType.NewIndex) = WeaponGroupType.SecondaryWeaponGroup
				cboWeaponGroupType.AddItem("Point Defense")
                cboWeaponGroupType.ItemData(cboWeaponGroupType.NewIndex) = WeaponGroupType.PointDefenseWeaponGroup

                If oHullTech.yTypeID = 1 AndAlso oHullTech.ySubTypeID = 3 Then
                    cboWeaponGroupType.AddItem("Bomb")
                    cboWeaponGroupType.ItemData(cboWeaponGroupType.NewIndex) = WeaponGroupType.BombWeaponGroup
                End If

				
				cboWeaponGroupType.FindComboItemData(WeaponGroupType.PrimaryWeaponGroup)
			End If
		End If

	End Sub

	Private Sub cboWeaponType_ItemSelected(ByVal lItemIndex As Integer) Handles cboWeaponType.ItemSelected
		cboWeapon.Clear()
		SetWeaponPlaceState(False, False)
		SetWeaponPlaceState(False, True)
        cboWeapon.ListIndex = -1

        Dim oHullTech As HullTech = Nothing
        If cboHull.ListIndex > -1 Then
            oHullTech = CType(goCurrentPlayer.GetTech(cboHull.ItemData(cboHull.ListIndex), ObjectType.eHullTech), HullTech)
        End If
        Dim yHullTypeID As Byte = 255
        If oHullTech Is Nothing = False Then yHullTypeID = HullTech.GetHullTypeID(oHullTech.yTypeID, oHullTech.ySubTypeID)

		If NewTutorialManager.TutorialOn = True Then
			If lItemIndex > -1 Then NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.ePrototypeWeaponTypeSelected, cboWeaponType.ItemData(lItemIndex), -1, -1, "")
		End If

		If lItemIndex > -1 Then
			Dim bFilter As Boolean = chkFilter.Value
			Dim yClassTypeID As Byte = CByte(cboWeaponType.ItemData(lItemIndex))
            For X As Int32 = 0 To goCurrentPlayer.mlTechUB
                If goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eWeaponTech Then
                    With CType(goCurrentPlayer.moTechs(X), WeaponTech)
                        If .ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched AndAlso .WeaponClassTypeID = yClassTypeID AndAlso (.HullTypeID = yHullTypeID OrElse .HullTypeID = 255) Then
                            If .bArchived = False OrElse bFilter = False Then
                                cboWeapon.AddItem(.GetComponentName)
                                cboWeapon.ItemData(cboWeapon.NewIndex) = goCurrentPlayer.moTechs(X).ObjectID
                            End If
                        End If
                    End With
                End If
            Next X
			For X As Int32 = 0 To glPlayerTechKnowledgeUB
				If glPlayerTechKnowledgeIdx(X) <> -1 AndAlso goPlayerTechKnowledge(X).yKnowledgeType > PlayerTechKnowledge.KnowledgeType.eSettingsLevel1 Then
					If goPlayerTechKnowledge(X).bArchived = False OrElse bFilter = False Then
						Dim oTech As Base_Tech = goPlayerTechKnowledge(X).oTech
                        If oTech.ObjTypeID = ObjectType.eWeaponTech AndAlso CType(oTech, WeaponTech).WeaponClassTypeID = yClassTypeID AndAlso (CType(oTech, WeaponTech).HullTypeID = yHullTypeID OrElse CType(oTech, WeaponTech).HullTypeID = 255) Then
                            cboWeapon.AddItem(oTech.GetComponentName)
                            cboWeapon.ItemData(cboWeapon.NewIndex) = oTech.ObjectID
                        End If
					End If
				End If
			Next X
		End If
	End Sub

	Private Sub cboWeapon_ItemSelected(ByVal lItemIndex As Integer) Handles cboWeapon.ItemSelected

		Dim sFinal As String = ""
		If lItemIndex > -1 Then
            Dim lID As Int32 = cboWeapon.ItemData(lItemIndex)

            Dim bPD As Boolean = False
            If cboWeaponGroupType.ListIndex > -1 Then
                If cboWeaponGroupType.ItemData(cboWeaponGroupType.ListIndex) = WeaponGroupType.PointDefenseWeaponGroup Then
                    bPD = True
                End If
            End If

			If NewTutorialManager.TutorialOn = True Then
				NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.ePrototypeWeaponSelected, lID, ObjectType.eWeaponTech, -1, "")
			End If

			For X As Int32 = 0 To goCurrentPlayer.mlTechUB
				If goCurrentPlayer.moTechs(X).ObjectID = lID AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eWeaponTech Then
                    With CType(goCurrentPlayer.moTechs(X), WeaponTech)

                        If bPD = True AndAlso .WeaponTypeID >= WeaponType.eMissile_Color_1 AndAlso .WeaponTypeID <= WeaponType.eMissile_Color_9 Then
                            MyBase.moUILib.AddNotification("Missiles cannot be Point Defense weapons.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            cboWeaponGroupType.ListIndex = -1
                            Return
                        End If

                        sFinal = .GetPrototypeBuilderString(bPD)
                    End With

					Exit For
				End If
			Next X

			If sFinal = "" Then
				For X As Int32 = 0 To glPlayerTechKnowledgeUB
					If glPlayerTechKnowledgeIdx(X) = lID AndAlso goPlayerTechKnowledge(X).oTech.ObjTypeID = ObjectType.eWeaponTech Then
                        With CType(goPlayerTechKnowledge(X).oTech, WeaponTech)
                            If bPD = True AndAlso .WeaponTypeID >= WeaponType.eMissile_Color_1 AndAlso .WeaponTypeID <= WeaponType.eMissile_Color_9 Then
                                MyBase.moUILib.AddNotification("Missiles cannot be Point Defense weapons.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                                cboWeaponGroupType.ListIndex = -1
                                Return
                            End If

                            sFinal = .GetPrototypeBuilderString(bPD)
                        End With
						Exit For
					End If
				Next X
			End If
		End If

		txtWpnStats.Caption = sFinal
	End Sub

    Public Sub SetWeaponPlaceState(ByVal bVal As Boolean, ByVal bRemoval As Boolean)
        If bVal = True Then
            'we are placing a weapon or removing a weapon
            If bRemoval = True Then
                btnPlaceWpn.Enabled = False
                btnRemoveWpn.Caption = "Cancel"

                hulMain.SetSelectingSlotConfig(SlotConfig.eWeapons, -1, 0)
                mbInRemoveWpn = True
            Else
                'Ok, we are setting to true to place a wepaon, so make sure there is a weapon to place
                If cboHull.ListIndex > -1 Then
                    If cboWeaponType.ListIndex > -1 AndAlso cboWeapon.ListIndex > -1 AndAlso cboWeaponGroupType.ListIndex > -1 Then
                        'Ok, get our weapon
                        Dim oWpn As WeaponTech = CType(goCurrentPlayer.GetTech(cboWeapon.ItemData(cboWeapon.ListIndex), ObjectType.eWeaponTech), WeaponTech)
                        Dim oHull As HullTech = CType(goCurrentPlayer.GetTech(cboHull.ItemData(cboHull.ListIndex), ObjectType.eHullTech), HullTech)

                        If oWpn Is Nothing = False AndAlso oHull Is Nothing = False Then
                            Dim lHullRequired As Int32 = oWpn.HullRequired
                            If cboWeaponGroupType.ItemData(cboWeaponGroupType.ListIndex) = WeaponGroupType.PointDefenseWeaponGroup Then
                                lHullRequired = CInt(lHullRequired * 0.5F)
                            End If

                            hulMain.SetSelectingSlotConfig(SlotConfig.eWeapons, lHullRequired, oHull.HullSize)
                            mbInPlaceWpn = True
                            btnPlaceWpn.Caption = "Cancel"
                            btnRemoveWpn.Enabled = False
                        End If
                    Else
                        MyBase.moUILib.AddNotification("Please select a weapon and a weapon group type.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    End If
                Else
                    MyBase.moUILib.AddNotification("Please select a Hull first.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
            End If
        Else
            'we are ending placing a weapon or removing a weapon
            'If bRemoval = True Then
            '    btnPlaceWpn.Enabled = True
            '    btnRemoveWpn.Caption = "Remove"
            'Else
            '    btnRemoveWpn.Enabled = True
            '    btnPlaceWpn.Caption = "Place"
            'End If
            'Nero 7/7/2010: It no longer matters, cancel both
            If btnPlaceWpn.Enabled <> True Then btnPlaceWpn.Enabled = True
            If btnPlaceWpn.Caption <> "Place" Then btnPlaceWpn.Caption = "Place"
            If btnRemoveWpn.Enabled <> True Then btnRemoveWpn.Enabled = True
            If btnRemoveWpn.Caption <> "Remove" Then btnRemoveWpn.Caption = "Remove"
            hulMain.ClearSelectingSlotConfig()
            mbInPlaceWpn = False
            mbInRemoveWpn = False
        End If
        RefreshAttributeBox()
    End Sub

	Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
		CloseMe()
	End Sub

	Private Sub CloseMe()
        'Dim oTmpWin As UIWindow = MyBase.moUILib.GetWindow("frmChat")
        'If oTmpWin Is Nothing = False Then oTmpWin.Visible = True
        'oTmpWin = Nothing

		If mfrmProdCost Is Nothing = False Then MyBase.moUILib.RemoveWindow(mfrmProdCost.ControlName)
		If mfrmResCost Is Nothing = False Then MyBase.moUILib.RemoveWindow(mfrmResCost.ControlName)

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		ReturnToPreviousView()
		If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
        goFullScreenBackground = Nothing
    End Sub

	Private Sub RefreshAttributeBox()
		Dim oArmorTech As ArmorTech = Nothing
		Dim oEngineTech As EngineTech = Nothing
		'Dim oHangarTech As HangarTech = Nothing
		Dim oHullTech As HullTech = Nothing
		Dim oRadarTech As RadarTech = Nothing
		Dim oShieldTech As ShieldTech = Nothing

		'set our labels too
		'lblFrontArmor.Caption = "Front (" & hscrFront.Value & ")"
        'lblLeftArmor.Caption = "Left (" & tpLeftArmor.PropertyValue & ")"
        'lblRearArmor.Caption = "Rear (" & tpRearArmor.PropertyValue & ")"
        'lblRightArmor.Caption = "Right (" & tpRightArmor.PropertyValue & ")"

        SetHscrBarMaxVals()

        Dim bResetScrollEvents As Boolean = mbIgnoreScrollEvents = False
        mbIgnoreScrollEvents = True

        'txtForeArmor.Caption = hscrFront.Value.ToString
        'txtLeftArmor.Caption = tpLeftArmor.PropertyValue.ToString
        'txtRearArmor.Caption = tpRearArmor.PropertyValue.ToString
        'txtRightArmor.Caption = tpRightArmor.PropertyValue.ToString
        UpdateProductionScroller()
        'txtProduction.Caption = tpProduction.PropertyValue.ToString
        If bResetScrollEvents = True Then mbIgnoreScrollEvents = False

        Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder
        Dim cPad As Char = "."c
        Dim sFormat As String = " #,##0"

        If optExpected.Value = True Then
            'This is purely based on the given values... so get them from our combo boxes
            If cboArmor.ListIndex > -1 Then oArmorTech = CType(goCurrentPlayer.GetTech(cboArmor.ItemData(cboArmor.ListIndex), ObjectType.eArmorTech), ArmorTech)
            If cboEngine.ListIndex > -1 Then oEngineTech = CType(goCurrentPlayer.GetTech(cboEngine.ItemData(cboEngine.ListIndex), ObjectType.eEngineTech), EngineTech)
            If cboHull.ListIndex > -1 Then oHullTech = CType(goCurrentPlayer.GetTech(cboHull.ItemData(cboHull.ListIndex), ObjectType.eHullTech), HullTech)
            If cboRadar.ListIndex > -1 Then oRadarTech = CType(goCurrentPlayer.GetTech(cboRadar.ItemData(cboRadar.ListIndex), ObjectType.eRadarTech), RadarTech)
            If cboShield.ListIndex > -1 Then oShieldTech = CType(goCurrentPlayer.GetTech(cboShield.ItemData(cboShield.ListIndex), ObjectType.eShieldTech), ShieldTech)

            Dim lSpecTraitID As elModelSpecialTrait = elModelSpecialTrait.NoSpecialTrait
            If oHullTech Is Nothing = False Then
                Dim oModelDef As ModelDef = goModelDefs.GetModelDef(oHullTech.ModelID)
                If oModelDef Is Nothing = False Then
                    lSpecTraitID = CType(oModelDef.lSpecialTraitID, elModelSpecialTrait)
                End If
            End If

            'Ok, figure out our hull used...
            Dim lHullUsed As Int32 = 0
            Dim lPowerUsed As Int32 = 0
            Dim lArmorHullRemaining As Int32 = 0

            If oArmorTech Is Nothing = False Then lHullUsed += (CInt(tpFrontArmor.PropertyValue + tpLeftArmor.PropertyValue + tpRearArmor.PropertyValue + tpRightArmor.PropertyValue) * oArmorTech.HullUsagePerPlate)
            If oEngineTech Is Nothing = False Then lHullUsed += oEngineTech.HullRequired
            If oHullTech Is Nothing = False Then lPowerUsed += oHullTech.PowerRequired
            If oRadarTech Is Nothing = False Then
                lHullUsed += oRadarTech.HullRequired
                lPowerUsed += oRadarTech.PowerRequired
            End If
            If oShieldTech Is Nothing = False Then
                lHullUsed += oShieldTech.HullRequired
                lPowerUsed += oShieldTech.PowerRequired
            End If
            lHullUsed += GetTotalCrewQuartersHull()

            'TODO: Get MSC Approval
            'If oHullTech Is Nothing = False Then
            '    lHullUsed += CInt(oHullTech.GetHullAllotment(SlotType.eNoChange, SlotConfig.eCargoBay))
            '    lHullUsed += CInt(oHullTech.GetHullAllotment(SlotType.eNoChange, SlotConfig.eHangar))
            'End If
            For X As Int32 = 0 To mlWpnUB
                Dim bFound As Boolean = False
                For Y As Int32 = 0 To goCurrentPlayer.mlTechUB
                    If goCurrentPlayer.moTechs(Y).ObjTypeID = ObjectType.eWeaponTech AndAlso goCurrentPlayer.moTechs(Y).ObjectID = moWpns(X).WeaponTechID Then
                        bFound = True
                        If moWpns(X).WeaponGroupTypeID = WeaponGroupType.PointDefenseWeaponGroup Then
                            lHullUsed += (CType(goCurrentPlayer.moTechs(Y), WeaponTech).HullRequired \ 2)
                        Else
                            lHullUsed += CType(goCurrentPlayer.moTechs(Y), WeaponTech).HullRequired
                        End If

                        lPowerUsed += CType(goCurrentPlayer.moTechs(Y), WeaponTech).PowerRequired
                        Exit For
                    End If
                Next Y
                If bFound = False Then
                    For Y As Int32 = 0 To glPlayerTechKnowledgeUB
                        If glPlayerTechKnowledgeIdx(Y) = moWpns(X).WeaponTechID Then
                            Dim oPTK As PlayerTechKnowledge = goPlayerTechKnowledge(Y)
                            If oPTK Is Nothing = False Then
                                If oPTK.oTech.ObjectID = moWpns(X).WeaponTechID AndAlso oPTK.oTech.ObjTypeID = ObjectType.eWeaponTech Then
                                    If moWpns(X).WeaponGroupTypeID = WeaponGroupType.PointDefenseWeaponGroup Then
                                        lHullUsed += (CType(oPTK.oTech, WeaponTech).HullRequired \ 2)
                                    Else
                                        lHullUsed += CType(oPTK.oTech, WeaponTech).HullRequired
                                    End If

                                    lPowerUsed += CType(oPTK.oTech, WeaponTech).PowerRequired
                                    Exit For
                                End If
                            End If
                        End If
                    Next Y
                End If
            Next X

            oSB.AppendLine("Basic Attributes")
            If oHullTech Is Nothing = False Then
                oSB.AppendLine("  Hull Size" & oHullTech.HullSize.ToString(sFormat).PadLeft(19, cPad))
                oSB.AppendLine("  Hull Used" & lHullUsed.ToString(sFormat).PadLeft(19, cPad))
                If oHullTech Is Nothing = False Then
                    Dim fTmpVal As Single = oHullTech.GetHullAllotment(SlotType.eNoChange, SlotConfig.eArmorConfig)
                    fTmpVal -= GetTotalCrewQuartersHull()
                    fTmpVal -= tpProduction.PropertyValue 'tpProduction.PropertyValue
                    If oArmorTech Is Nothing = False Then fTmpVal -= (CInt(tpFrontArmor.PropertyValue + tpLeftArmor.PropertyValue + tpRearArmor.PropertyValue + tpRightArmor.PropertyValue) * oArmorTech.HullUsagePerPlate)
                    oSB.AppendLine("  Space Remaining" & CInt(Math.Floor(fTmpVal)).ToString(sFormat).PadLeft(13, cPad))
                End If
                Dim lSugCrew As Int32 = CInt(oHullTech.HullSize / 100)
                If oHullTech.yTypeID = 2 Then lSugCrew = 0
                oSB.AppendLine("  Suggested Crew" & lSugCrew.ToString(sFormat).PadLeft(14, cPad))
                lblSuggestedCrew.Caption = "Suggested Crew: " & lSugCrew
            Else : oSB.AppendLine("  Hull Size Undecided")
            End If
            If oEngineTech Is Nothing = False Then
                Dim lPowerProd As Int32 = oEngineTech.PowerProd
                If lSpecTraitID = elModelSpecialTrait.PowerGen2 Then
                    lPowerProd *= 2
                ElseIf lSpecTraitID = elModelSpecialTrait.PowerGen3 Then
                    lPowerProd *= 3
                End If
                oSB.AppendLine("  Power Generated" & lPowerProd.ToString(sFormat).PadLeft(13, cPad)) 'oEngineTech.PowerProd.ToString)
                oSB.AppendLine("  Power Remaining" & (lPowerProd - lPowerUsed).ToString(sFormat).PadLeft(13, cPad))
            End If

            oSB.AppendLine("Defenses")
            If oShieldTech Is Nothing = False Then
                oSB.AppendLine("  Shield Strength" & oShieldTech.MaxHitPoints.ToString(sFormat).PadLeft(13, cPad))
                oSB.AppendLine("  Recharge Rate" & oShieldTech.RechargeRate.ToString(sFormat).PadLeft(15, cPad))
                oSB.AppendLine("    Per second" & (oShieldTech.RechargeRate / (oShieldTech.RechargeFreq / 30.0F)).ToString(sFormat & ".00").PadLeft(16, cPad))
            End If
            If oArmorTech Is Nothing = False Then
                oSB.AppendLine("  Armor Resists")
                oSB.AppendLine("    Beam/Energy Resist" & oArmorTech.BeamResist.ToString(sFormat).PadLeft(8, cPad))
                oSB.AppendLine("    Burn Resist" & oArmorTech.BurnResist.ToString(sFormat).PadLeft(15, cPad))
                oSB.AppendLine("    Chemical Resist" & oArmorTech.ChemicalResist.ToString(sFormat).PadLeft(11, cPad))
                oSB.AppendLine("    Impact Resist" & oArmorTech.ImpactResist.ToString(sFormat).PadLeft(13, cPad))
                oSB.AppendLine("    Magnetic Resist" & oArmorTech.MagneticResist.ToString(sFormat).PadLeft(11, cPad))
                oSB.AppendLine("    Piercing Resist" & oArmorTech.PiercingResist.ToString(sFormat).PadLeft(11, cPad))
                oSB.AppendLine("  Armor Hitpoints")


                Dim lArmorMult As Int32 = 1
                If lSpecTraitID = elModelSpecialTrait.Armor1000 Then lArmorMult *= 10
                If oHullTech Is Nothing = False AndAlso oHullTech.yTypeID = 2 AndAlso oHullTech.ySubTypeID <> 9 Then lArmorMult *= 10

                'If lSpecTraitID = elModelSpecialTrait.Armor1000 Then
                oSB.AppendLine("    Forward" & (oArmorTech.HPPerPlate * tpFrontArmor.PropertyValue * lArmorMult).ToString(sFormat).PadLeft(19, cPad))
                oSB.AppendLine("    Left" & (oArmorTech.HPPerPlate * tpLeftArmor.PropertyValue * lArmorMult).ToString(sFormat).PadLeft(22, cPad))
                oSB.AppendLine("    Rear" & (oArmorTech.HPPerPlate * tpRearArmor.PropertyValue * lArmorMult).ToString(sFormat).PadLeft(22, cPad))
                oSB.AppendLine("    Right" & (oArmorTech.HPPerPlate * tpRightArmor.PropertyValue * lArmorMult).ToString(sFormat).PadLeft(21, cPad))
                'Else
                '    oSB.AppendLine("    Forward: " & (oArmorTech.HPPerPlate * tpFrontArmor.PropertyValue).ToString("0.#") & " hps")
                '    oSB.AppendLine("    Left: " & (oArmorTech.HPPerPlate * tpLeftArmor.PropertyValue).ToString("0.#") & " hps")
                '    oSB.AppendLine("    Rear: " & (oArmorTech.HPPerPlate * tpRearArmor.PropertyValue).ToString("0.#") & " hps")
                '    oSB.AppendLine("    Right: " & (oArmorTech.HPPerPlate * tpRightArmor.PropertyValue).ToString("0.#") & " hps")
                'End If
            End If

            If oHullTech Is Nothing = False Then oSB.AppendLine("  Structure" & oHullTech.StructuralHitPoints.ToString(sFormat).PadLeft(19, cPad))

            If oRadarTech Is Nothing = False Then
                oSB.AppendLine("Electronics")
                oSB.AppendLine("  Targeting Accuracy" & oRadarTech.WeaponAcc.ToString(sFormat).PadLeft(10, cPad))
                oSB.AppendLine("  Scan Resolution" & oRadarTech.ScanResolution.ToString(sFormat).PadLeft(13, cPad))
                If lSpecTraitID = elModelSpecialTrait.ScanRange10 Then
                    oSB.AppendLine("  Optimum Radar Range" & Math.Min(255, CInt(oRadarTech.OptimumRange) + 10).ToString(sFormat).PadLeft(9, cPad))
                    oSB.AppendLine("  Maximum Radar Range" & Math.Min(255, CInt(oRadarTech.MaximumRange) + 10).ToString(sFormat).PadLeft(9, cPad))
                ElseIf lSpecTraitID = elModelSpecialTrait.ScanRange15 Then
                    oSB.AppendLine("  Optimum Radar Range" & Math.Min(255, CInt(oRadarTech.OptimumRange) + 15).ToString(sFormat).PadLeft(9, cPad))
                    oSB.AppendLine("  Maximum Radar Range" & Math.Min(255, CInt(oRadarTech.MaximumRange) + 15).ToString(sFormat).PadLeft(9, cPad))
                Else
                    oSB.AppendLine("  Optimum Radar Range" & oRadarTech.OptimumRange.ToString(sFormat).PadLeft(9, cPad))
                    oSB.AppendLine("  Maximum Radar Range" & oRadarTech.MaximumRange.ToString(sFormat).PadLeft(9, cPad))
                End If

                oSB.AppendLine("  Disruption Resist" & oRadarTech.DisruptionResistance.ToString(sFormat).PadLeft(11, cPad))

                If lSpecTraitID = elModelSpecialTrait.NoJammer Then
                    oSB.AppendLine("  Jamming Abilities Negated")
                Else
                    oSB.AppendLine("  Jamming Strength" & oRadarTech.JamStrength.ToString(sFormat).PadLeft(12, cPad))
                    If oRadarTech.JamTargets = 128 Then
                        oSB.AppendLine("    Targets" & " Area Effect".PadLeft(19, cPad))
                    Else
                        oSB.AppendLine("    Targets" & oRadarTech.JamTargets.ToString(sFormat).PadLeft(19, cPad))
                    End If
                End If

                If oHullTech Is Nothing = False AndAlso oHullTech.yTypeID = 2 Then oSB.AppendLine("  +50 Point Defense Accuracy")

            Else : oSB.AppendLine("Electronics" & " None".PadLeft(19, cPad))
            End If

            If mlWpnUB = -1 Then
                oSB.AppendLine("Weapons" & " None".PadLeft(23, cPad))
            Else
                oSB.AppendLine("Weapons" & (mlWpnUB + 1).ToString(sFormat).PadLeft(23, cPad))
                For X As Int32 = 0 To mlWpnUB
                    Dim oWpn As WeaponTech = CType(goCurrentPlayer.GetTech(moWpns(X).WeaponTechID, ObjectType.eWeaponTech), WeaponTech)
                    If oWpn Is Nothing = False Then
                        Select Case moWpns(X).WeaponGroupTypeID
                            Case WeaponGroupType.BombWeaponGroup
                                oSB.AppendLine("  Weapon " & moWpns(X).lGroupNum & " (Bomb)")
                            Case WeaponGroupType.PointDefenseWeaponGroup
                                oSB.AppendLine("  Weapon " & moWpns(X).lGroupNum & " (Point Defense)")
                            Case WeaponGroupType.PrimaryWeaponGroup
                                oSB.AppendLine("  Weapon " & moWpns(X).lGroupNum & " (Primary)")
                            Case WeaponGroupType.SecondaryWeaponGroup
                                oSB.AppendLine("  Weapon " & moWpns(X).lGroupNum & " (Secondary)")
                            Case Else
                                oSB.AppendLine("  Weapon " & moWpns(X).lGroupNum)
                        End Select
                        oWpn.GetPrototypeBuilderAttributeString(oSB, moWpns(X).WeaponGroupTypeID = WeaponGroupType.PointDefenseWeaponGroup)
                    End If
                Next X
            End If

            oSB.AppendLine("Mobility")
            If oEngineTech Is Nothing = False Then

                'TODO: Get MSC Approval 
                'Dim fTemp As Single = CSng(oEngineTech.Thrust / lHullUsed)
                Dim fTemp As Single = CSng(oEngineTech.Thrust / oHullTech.HullSize)
                Dim lManeuver As Int32 = CInt(Math.Min(oEngineTech.Maneuver, oEngineTech.Maneuver * fTemp))
                Dim lSpeed As Int32 = CInt(Math.Min(oEngineTech.Speed, oEngineTech.Speed * fTemp))

                Select Case lSpecTraitID
                    Case elModelSpecialTrait.Maneuver1
                        If lManeuver <> 0 Then lManeuver += 1
                    Case elModelSpecialTrait.Maneuver10
                        If lManeuver <> 0 Then lManeuver += 10
                    Case elModelSpecialTrait.Maneuver10Critical1
                        If lManeuver <> 0 Then lManeuver += 10
                    Case elModelSpecialTrait.Maneuver10Critical2
                        If lManeuver <> 0 Then lManeuver += 10
                    Case elModelSpecialTrait.Maneuver2
                        If lManeuver <> 0 Then lManeuver += 2
                    Case elModelSpecialTrait.Maneuver3
                        If lManeuver <> 0 Then lManeuver += 3
                    Case elModelSpecialTrait.Maneuver30
                        If lManeuver <> 0 Then lManeuver += 30
                    Case elModelSpecialTrait.Maneuver4
                        If lManeuver <> 0 Then lManeuver += 4
                    Case elModelSpecialTrait.Maneuver5
                        If lManeuver <> 0 Then lManeuver += 5
                    Case elModelSpecialTrait.Maneuver5Critical2
                        If lManeuver <> 0 Then lManeuver += 5
                    Case elModelSpecialTrait.Speed1
                        If lSpeed <> 0 Then lSpeed += 1
                    Case elModelSpecialTrait.Speed2
                        If lSpeed <> 0 Then lSpeed += 2
                    Case elModelSpecialTrait.Speed5Critical2
                        If lSpeed <> 0 Then lSpeed += 5
                    Case elModelSpecialTrait.SpeedAndManeuver1
                        If lSpeed <> 0 Then lSpeed += 1
                        If lManeuver <> 0 Then lManeuver += 1
                End Select

                mbUnitImmobile = False
                If oEngineTech.Speed = 0 AndAlso oEngineTech.Maneuver = 0 Then
                    oSB.AppendLine("  No Movement ability")
                Else
                    oSB.AppendLine("  Maneuver" & lManeuver.ToString(sFormat).PadLeft(20, cPad))
                    oSB.AppendLine("  Max Speed" & lSpeed.ToString(sFormat).PadLeft(19, cPad))
                    If lSpeed = 0 OrElse lManeuver = 0 Then mbUnitImmobile = True
                End If
            End If

            'If oHullTech Is Nothing = False Then
            '    Dim fFuelCap As Single = oHullTech.GetHullAllotment(SlotType.eNoChange, SlotConfig.eFuelBay)
            '    If fFuelCap <> 0.0F Then
            '        oSB.AppendLine("  Fuel Capacity: " & fFuelCap.ToString("0.#"))
            '    End If
            'End If

            oSB.AppendLine("Utility")
            If oHullTech Is Nothing = False Then
                oSB.Append(oHullTech.GetHangarDoorDetails)

                Dim fCargoCap As Single = oHullTech.GetHullAllotment(SlotType.eNoChange, SlotConfig.eCargoBay)
                Select Case lSpecTraitID
                    Case elModelSpecialTrait.Cargo10
                        fCargoCap = fCargoCap * 1.1F
                    Case elModelSpecialTrait.Cargo200
                        fCargoCap = fCargoCap * 3.0F
                    Case elModelSpecialTrait.Cargo20
                        fCargoCap = fCargoCap * 1.2F
                    Case elModelSpecialTrait.Cargo5
                        fCargoCap = fCargoCap * 1.05F
                    Case elModelSpecialTrait.CargoAndHangar10
                        fCargoCap = fCargoCap * 1.1F
                    Case elModelSpecialTrait.CargoAndHangar3
                        fCargoCap = fCargoCap * 1.03F
                    Case elModelSpecialTrait.CargoAndHangar6
                        fCargoCap = fCargoCap * 1.06F
                End Select

                oSB.AppendLine("  Cargo Capacity" & Math.Floor(fCargoCap).ToString(sFormat).PadLeft(14, cPad))

                'MSC: we only care about the mesh portion of the model
                Dim iMesh As Int16 = (oHullTech.ModelID And 255S)
                If oHullTech.yTypeID = 2 Then
                    Select Case oHullTech.ySubTypeID
                        Case 0  'CC
                            oSB.AppendLine("Production: Command Center Ability")
                            oSB.AppendLine("  Based On: Production Slots")
                        Case 1 'Mine
                            oSB.AppendLine("Production: Mining")
                            oSB.AppendLine("  Based On: Production Slots")
                        Case 3 'personnel
                            If iMesh = 6 Then
                                oSB.AppendLine("Production: Enlisted")
                                oSB.AppendLine("  Based On: Production Slots")
                            ElseIf iMesh = 7 Then
                                oSB.AppendLine("Production: Officers")
                                oSB.AppendLine("  Based On: Production Slots")
                            End If
                        Case 4 'Power
                            oSB.AppendLine("Production: Power")
                            oSB.AppendLine("  Based On: Power Remaining")
                        Case 5 'Production
                            If iMesh = 13 OrElse iMesh = 102 Then
                                oSB.AppendLine("Production: Air/Space Production")
                            ElseIf iMesh = 137 Then
                                oSB.AppendLine("Production: Naval Production")
                            Else
                                oSB.AppendLine("Production: Ground Production")
                            End If
                            oSB.AppendLine("  Based On: Production Slots and Hangar Capacity")
                        Case 6 'Refining
                            oSB.AppendLine("Production: Refining")
                        Case 7 'Research
                            oSB.AppendLine("Production: Research")
                            oSB.AppendLine("  Based On: Production Slots")
                        Case 8 'Residence
                            oSB.AppendLine("Production: Residence")
                            oSB.AppendLine("  Based On: Production Slots")
                        Case 9 'SpaceStation
                            oSB.AppendLine("Production: Space Station Ability")
                        Case 10
                            oSB.AppendLine("Production: Galactic Trading")
                    End Select
                    oSB.AppendLine("  Production Slots: " & tpProduction.PropertyValue)
                ElseIf oHullTech.yTypeID = 7 Then
                    Select Case oHullTech.ySubTypeID
                        Case 0, 1
                            oSB.AppendLine("Production: Facilities")
                        Case 3
                            oSB.AppendLine("Production: Mining")
                    End Select
                End If
            End If

        Else
            If moTech Is Nothing OrElse moTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eComponentDesign Then
                oSB.AppendLine("Unknown")
                oSB.AppendLine("  We will need to research this")
                oSB.AppendLine("  design before knowing how well")
                oSB.AppendLine("  it matches up to our expected")
                oSB.AppendLine("  results.")
            Else
                'Display the actual results here...
                If moTech.oActualResults Is Nothing = False Then
                    With moTech.oActualResults
                        oSB.AppendLine("Basic Attributes")
                        oSB.AppendLine("  Hull Size" & .HullSize.ToString(sFormat).PadLeft(19, cPad))

                        'Power generated?

                        oSB.AppendLine("Defenses")
                        If .Shield_MaxHP <> 0 Then oSB.AppendLine("  Shield Strength" & .Shield_MaxHP.ToString(sFormat).PadLeft(13, cPad))
                        If .ShieldRecharge <> 0 Then oSB.AppendLine("  Recharge Rate" & .ShieldRecharge.ToString(sFormat).PadLeft(15, cPad))
                        If .ShieldRechargeFreq <> 0 Then oSB.AppendLine("    Per second" & (.ShieldRecharge / (.ShieldRechargeFreq / 30)).ToString(sFormat & ".00").PadLeft(16, cPad))

                        'display armor values
                        If .Armor_MaxHP(UnitArcs.eBackArc) <> 0 OrElse .Armor_MaxHP(UnitArcs.eForwardArc) <> 0 OrElse .Armor_MaxHP(UnitArcs.eLeftArc) <> 0 OrElse .Armor_MaxHP(UnitArcs.eRightArc) <> 0 Then
                            ' if resist is stored as 100, meaning no resist, then do not display
                            If .BeamResist = 100 AndAlso .FlameResist = 100 AndAlso .ChemicalResist = 100 AndAlso .ImpactResist = 100 AndAlso .ECMResist = 100 AndAlso .PiercingResist = 100 Then
                                oSB.AppendLine("  Armor Resists" & " None".PadLeft(15, cPad))
                            Else
                                oSB.AppendLine("  Armor Resists")
                                If .BeamResist <> 100 Then oSB.AppendLine("    Beam/Energy Resist" & (100 - .BeamResist).ToString(sFormat).PadLeft(8, cPad))
                                If .FlameResist <> 100 Then oSB.AppendLine("    Burn Resist" & (100 - .FlameResist).ToString(sFormat).PadLeft(15, cPad))
                                If .ChemicalResist <> 100 Then oSB.AppendLine("    Chemical Resist" & (100 - .ChemicalResist).ToString(sFormat).PadLeft(11, cPad))
                                If .ImpactResist <> 100 Then oSB.AppendLine("    Impact Resist" & (100 - .ImpactResist).ToString(sFormat).PadLeft(13, cPad))
                                If .ECMResist <> 100 Then oSB.AppendLine("    Magnetic Resist" & (100 - .ECMResist).ToString(sFormat).PadLeft(11, cPad))
                                If .PiercingResist <> 100 Then oSB.AppendLine("    Piercing Resist" & (100 - .PiercingResist).ToString(sFormat).PadLeft(11, cPad))
                            End If
                            oSB.AppendLine("    Integrity" & (100 - .yArmorIntegrityRoll).ToString(sFormat).PadLeft(16, cPad) & "%")
                            oSB.AppendLine("  Armor Hitpoints")
                            oSB.AppendLine("    Forward" & .Armor_MaxHP(UnitArcs.eForwardArc).ToString(sFormat).PadLeft(19, cPad))
                            oSB.AppendLine("    Left" & .Armor_MaxHP(UnitArcs.eLeftArc).ToString(sFormat).PadLeft(22, cPad))
                            oSB.AppendLine("    Rear" & .Armor_MaxHP(UnitArcs.eBackArc).ToString(sFormat).PadLeft(22, cPad))
                            oSB.AppendLine("    Right" & .Armor_MaxHP(UnitArcs.eRightArc).ToString(sFormat).PadLeft(21, cPad))
                        Else
                            oSB.AppendLine("  Armor Plating" & " None".PadLeft(15, cPad))
                        End If
                        oSB.AppendLine("  Structure" & .Structure_MaxHP.ToString(sFormat).PadLeft(19, cPad))

                        'display radar values
                        If .Weapon_Acc <> 0 OrElse .ScanResolution <> 0 OrElse .OptRadarRange <> 0 OrElse .MaxRadarRange <> 0 OrElse .DisruptionResistance <> 0 OrElse .JamStrength <> 0 OrElse .JamTargets <> 0 Then
                            oSB.AppendLine("Electronics")
                            If .Weapon_Acc <> 0 Then oSB.AppendLine("  Targeting Accuracy" & .Weapon_Acc.ToString(sFormat).PadLeft(10, cPad))
                            If .ScanResolution <> 0 Then oSB.AppendLine("  Scan Resolution" & .ScanResolution.ToString(sFormat).PadLeft(13, cPad))
                            If .OptRadarRange <> 0 Then oSB.AppendLine("  Optimum Radar Range" & .OptRadarRange.ToString(sFormat).PadLeft(9, cPad))
                            If .MaxRadarRange <> 0 Then oSB.AppendLine("  Maximum Radar Range" & .MaxRadarRange.ToString(sFormat).PadLeft(9, cPad))
                            If .DisruptionResistance <> 0 Then oSB.AppendLine("  Disruption Resistance" & .DisruptionResistance.ToString(sFormat).PadLeft(7, cPad))
                            'display jamming values
                            If .JamStrength <> 0 AndAlso .JamTargets <> 0 Then
                                oSB.AppendLine("  Jamming Strength" & .JamStrength.ToString(sFormat).PadLeft(12, cPad))
                                If .JamTargets = 128 Then
                                    oSB.AppendLine("    Targets" & " Area Effect".PadLeft(19, cPad))
                                Else
                                    oSB.AppendLine("    Targets" & .JamTargets.ToString(sFormat).PadLeft(19, cPad))
                                End If
                            End If
                        Else
                            oSB.AppendLine("Electronics" & " None".PadLeft(19, cPad))
                        End If

                        'display weapons
                        If .WeaponDefUB = -1 Then
                            oSB.AppendLine("Weapons" & " None".PadLeft(23, cPad))
                        Else
                            oSB.AppendLine("Weapons" & (.WeaponDefUB + 1).ToString(sFormat).PadLeft(23, cPad))
                            For X As Int32 = 0 To .WeaponDefUB
                                Dim sTmp As String = "  "

                                Select Case .WeaponDefs(X).ArcID
                                    Case UnitArcs.eForwardArc
                                        sTmp &= "Front arc"
                                    Case UnitArcs.eLeftArc
                                        sTmp &= "Left arc"
                                    Case UnitArcs.eRightArc
                                        sTmp &= "Right arc"
                                    Case UnitArcs.eBackArc
                                        sTmp &= "Rear arc"
                                    Case Else
                                        sTmp &= "All arcs"
                                End Select
                                Select Case .WeaponDefs(X).WpnGroup
                                    Case WeaponGroupType.BombWeaponGroup
                                        sTmp &= " (Bomb)"
                                    Case WeaponGroupType.PointDefenseWeaponGroup
                                        sTmp &= " (Point Defense)"
                                    Case WeaponGroupType.PrimaryWeaponGroup
                                        sTmp &= " (Primary)"
                                    Case WeaponGroupType.SecondaryWeaponGroup
                                        sTmp &= " (Secondary)"
                                End Select
                                oSB.AppendLine(sTmp)

                                oSB.AppendLine("    " & .WeaponDefs(X).WeaponName)

                                oSB.AppendLine("    RoF" & (.WeaponDefs(X).ROF_Delay / 30.0F).ToString(sFormat & ".00").PadLeft(23, cPad))

                                Dim lMinDmg As Int32 = .WeaponDefs(X).BeamMinDmg + .WeaponDefs(X).ChemicalMinDmg + .WeaponDefs(X).ECMMinDmg + .WeaponDefs(X).FlameMinDmg + .WeaponDefs(X).ImpactMinDmg + .WeaponDefs(X).PiercingMinDmg
                                Dim lMaxDmg As Int32 = .WeaponDefs(X).BeamMaxDmg + .WeaponDefs(X).ChemicalMaxDmg + .WeaponDefs(X).ECMMaxDmg + .WeaponDefs(X).FlameMaxDmg + .WeaponDefs(X).ImpactMaxDmg + .WeaponDefs(X).PiercingMaxDmg
                                oSB.AppendLine("    Damage" & (" " & lMinDmg & " - " & lMaxDmg).PadLeft(20, cPad))
                            Next X
                        End If
                        If .Maneuver = 0 AndAlso .MaxSpeed = 0 Then
                            oSB.AppendLine("Mobility" & " None".PadLeft(22, cPad))
                        Else
                            oSB.AppendLine("Mobility")
                            If .Maneuver <> 0 Then oSB.AppendLine("  Maneuver" & .Maneuver.ToString(sFormat).PadLeft(20, cPad))
                            If .MaxSpeed <> 0 Then oSB.AppendLine("  Max Speed" & .MaxSpeed.ToString(sFormat).PadLeft(19, cPad))
                        End If

                        'If .Fuel_Cap <> 0 Then oSB.AppendLine("  Fuel Capacity: " & .Fuel_Cap)



                        'Hangar is never altered through noises
                        Dim sHangerDetails As String = ""
                        oHullTech = CType(goCurrentPlayer.GetTech(moTech.HullID, ObjectType.eHullTech), HullTech)
                        If oHullTech Is Nothing = False Then sHangerDetails = oHullTech.GetHangarDoorDetails()
                        If .Cargo_Cap <> 0 OrElse sHangerDetails <> "" Then
                            oSB.AppendLine("Utility")
                            If .Cargo_Cap <> 0 Then oSB.AppendLine("  Cargo Capacity" & .Cargo_Cap.ToString(sFormat).PadLeft(14, cPad))
                            If sHangerDetails <> "" Then oSB.Append(sHangerDetails)
                        Else
                            oSB.AppendLine("Utility" & " None".PadLeft(23, cPad))
                        End If



                        If .ProdFactor <> 0 Then
                            oSB.AppendLine("Production")
                            If moTech.lActualResultsWorkers <> 0 Then oSB.AppendLine("  Jobs" & moTech.lActualResultsWorkers.ToString(sFormat).PadLeft(24, cPad))
                            oSB.AppendLine("  Production Rating" & .ProdFactor.ToString(sFormat).PadLeft(11, cPad))
                            oSB.AppendLine("  Power Rating" & .PowerFactor.ToString(sFormat).PadLeft(16, cPad))

                            Dim sProdType As String = ""
                            Select Case CType(.ProductionTypeID, ProductionType)
                                Case ProductionType.eAerialProduction
                                    sProdType = "Air/Space"
                                Case ProductionType.eColonists
                                    sProdType = "Residence"
                                Case ProductionType.eCommandCenterSpecial
                                    sProdType = "Command Center"
                                Case ProductionType.eCredits
                                    sProdType = "Financials"
                                Case ProductionType.eEnlisted
                                    sProdType = "Enlisted"
                                Case ProductionType.eEnlistedAndOfficers
                                    sProdType = "Enlisted/Officers"
                                Case ProductionType.eFacility
                                    sProdType = "Facilities"
                                Case ProductionType.eFood
                                    sProdType = "Food"
                                Case ProductionType.eLandProduction
                                    sProdType = "Ground"
                                Case ProductionType.eMining
                                    sProdType = "Mining"
                                Case ProductionType.eMorale
                                    sProdType = "Morale"
                                Case ProductionType.eNavalProduction
                                    sProdType = "Naval"
                                Case ProductionType.eOfficers
                                    sProdType = "Officers"
                                Case ProductionType.ePowerCenter
                                    sProdType = "Power"
                                Case ProductionType.eRefining
                                    sProdType = "Refining"
                                Case ProductionType.eResearch
                                    sProdType = "Research"
                                Case ProductionType.eSpaceStationSpecial
                                    sProdType = "Space Station"
                                Case ProductionType.eWareHouse
                                    sProdType = "Warehouse"
                            End Select

                            If sProdType <> "" Then oSB.AppendLine("  Type" & (" " & sProdType).PadLeft(24, cPad))
                        End If
                        Dim iCombatRating As Int32 = GetCombatRating(moTech.oActualResults)
                        If iCombatRating > 0 Then
                            oSB.AppendLine(vbCrLf & "Combat Rating" & iCombatRating.ToString(sFormat).PadLeft(17, cPad))
                            'If oHullTech Is Nothing = False AndAlso oHullTech.yTypeID <> 2 Then
                            '    If .WeaponDefUB > -1 Then
                            '        oSB.AppendLine("Warpoint Cost" & (iCombatRating \ 10000).ToString(" #,##0").PadLeft(17, cPad))
                            '    Else
                            '        oSB.AppendLine("Warpoint Cost............... 1")
                            '    End If
                            'End If
                        End If
                    End With
                End If
            End If
        End If

        txtAttrs.Caption = oSB.ToString
    End Sub

    Private Function GetCombatRating(ByVal oEntity As EntityDef) As Int32

        Dim CombatRating As Int32
        Try
            With oEntity
                Dim lTotalArmor As Int32 = .Armor_MaxHP(0) + .Armor_MaxHP(1) + .Armor_MaxHP(2) + .Armor_MaxHP(3)
                Dim lResists As Int32 = CInt(.PiercingResist) + CInt(.ImpactResist) + CInt(.BeamResist) + CInt(.ECMResist) + CInt(.FlameResist) + CInt(.ChemicalResist)
                lResists = CInt(lResists / 6.0F)
                Dim fShieldRegenPerSec As Single
                If .ShieldRechargeFreq > 0 Then
                    fShieldRegenPerSec = .ShieldRecharge * (30.0F / .ShieldRechargeFreq)
                Else : fShieldRegenPerSec = 0
                End If
                Dim SideWeaponPower(3) As Int32
                Dim lAllArcPower As Int32 = 0
                For X As Int32 = 0 To .WeaponDefUB
                    If .WeaponDefs(X).ArcID <> UnitArcs.eAllArcs Then
                        SideWeaponPower(.WeaponDefs(X).ArcID) += .WeaponDefs(X).lFirePowerRating
                    Else : lAllArcPower += .WeaponDefs(X).lFirePowerRating
                    End If
                Next X
                Dim lWpnStr As Int32 = SideWeaponPower(0) + SideWeaponPower(1) + SideWeaponPower(2) + SideWeaponPower(3) + (lAllArcPower * 2)
                'old CombatRating = CInt(((((lTotalArmor / ((lResists + 1) / 100.0F)) + (.Shield_MaxHP * fShieldRegenPerSec))) + (.MaxSpeed * (2 * .Maneuver))) + (lWpnStr * 5))
                CombatRating = CInt(((lTotalArmor * ((100 - .yArmorIntegrityRoll) / 100)) / ((lResists + 1) / 100.0F) + .Shield_MaxHP))
                If .Shield_MaxHP > 0 Then CombatRating = CInt(Math.Pow(CombatRating, (1 + (fShieldRegenPerSec / .Shield_MaxHP))))
                CombatRating += (.MaxSpeed * 2 * .Maneuver) + lWpnStr

                'For Primary-Server (Military Score)'s "Correct" CR use this
                'CombatRating = (lTotalArmor \ ((lResists \ 100) + 1))
                'CombatRating += CInt(.Shield_MaxHP * fShieldRegenPerSec)
                'CombatRating += CInt(.MaxSpeed * (2I * .Maneuver))
                'CombatRating += (lWpnStr * 5)

            End With
        Catch
        End Try
        If CombatRating < 1 Then CombatRating = 1
        Return CombatRating
    End Function

    Private Function GetTotalCrewQuartersHull() As Int32
        Dim oHullTech As HullTech = Nothing
        Dim oArmorTech As ArmorTech = Nothing
        Dim oEngineTech As EngineTech = Nothing
        Dim oShieldTech As ShieldTech = Nothing
        Dim oRadarTech As RadarTech = Nothing

        'Now, for crew requirements
        Dim lTotalCrew As Int32 = 0

        If cboHull.ListIndex > -1 Then oHullTech = CType(goCurrentPlayer.GetTech(cboHull.ItemData(cboHull.ListIndex), ObjectType.eHullTech), HullTech)
        If cboEngine.ListIndex > -1 Then oEngineTech = CType(goCurrentPlayer.GetTech(cboEngine.ItemData(cboEngine.ListIndex), ObjectType.eEngineTech), EngineTech)
        If cboShield.ListIndex > -1 Then oShieldTech = CType(goCurrentPlayer.GetTech(cboShield.ItemData(cboShield.ListIndex), ObjectType.eShieldTech), ShieldTech)
        If cboRadar.ListIndex > -1 Then oRadarTech = CType(goCurrentPlayer.GetTech(cboRadar.ItemData(cboRadar.ListIndex), ObjectType.eRadarTech), RadarTech)

        If oEngineTech Is Nothing = False AndAlso oEngineTech.oProductionCost Is Nothing = False Then
            With oEngineTech.oProductionCost
                lTotalCrew += .ColonistCost + .EnlistedCost + .OfficerCost
            End With
        End If
        If oShieldTech Is Nothing = False AndAlso oShieldTech.oProductionCost Is Nothing = False Then
            With oShieldTech.oProductionCost
                lTotalCrew += .ColonistCost + .EnlistedCost + .OfficerCost
            End With
        End If
        If oRadarTech Is Nothing = False AndAlso oRadarTech.oProductionCost Is Nothing = False Then
            With oRadarTech.oProductionCost
                lTotalCrew += .ColonistCost + .EnlistedCost + .OfficerCost
            End With
        End If
        For X As Int32 = 0 To mlWpnUB
            If moWpns(X).WeaponTechID <> -1 Then
                Dim oWpn As WeaponTech = CType(goCurrentPlayer.GetTech(moWpns(X).WeaponTechID, ObjectType.eWeaponTech), WeaponTech)
                If oWpn Is Nothing = False AndAlso oWpn.oProductionCost Is Nothing = False Then
                    With oWpn.oProductionCost
                        lTotalCrew += .ColonistCost + .EnlistedCost + .OfficerCost
                    End With
                End If
            End If
        Next X

        'Multiply by hull usage per residence to get hull consumption of all crew
        lTotalCrew += CInt(tpSecurity.PropertyValue)
        Dim lOriginalCrew As Int32 = lTotalCrew

        Dim lBaseCrew As Int32 = 11
        Dim bIgnoreCrew As Boolean = False
        If oHullTech Is Nothing = False AndAlso oHullTech.HullSize <= 750 Then
            lBaseCrew = 2
            lTotalCrew = 2
            lOriginalCrew = 2
            bIgnoreCrew = True
        Else
            lTotalCrew *= Math.Min(lBaseCrew, goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBetterHullToResidence))
        End If

        'lblCrewQtrs.Caption = "Crew Quarters: " & lTotalCrew & " hull for " & lOriginalCrew & " crew"  ' (" & lOriginalCrew & ")"
        lblCrewQtrs.Caption = "Crew: " & lOriginalCrew.ToString("#,##0") & " crew in " & lTotalCrew.ToString("#,##0") & " hull"  ' (" & lOriginalCrew & ")"

        If oHullTech Is Nothing = False Then
            Dim lSugCrew As Int32 = CInt(oHullTech.HullSize / 100)
            If oHullTech.yTypeID = 2 Then lSugCrew = 0
            If bIgnoreCrew = True Then lSugCrew = 0
            If lOriginalCrew < lSugCrew Then
                lblSuggestedCrew.ForeColor = System.Drawing.Color.FromArgb(192, 255, 0, 0)
                mbMeetsSuggestedCrew = False
            Else
                lblSuggestedCrew.ForeColor = System.Drawing.Color.FromArgb(192, muSettings.InterfaceBorderColor.R, muSettings.InterfaceBorderColor.G, muSettings.InterfaceBorderColor.B)
                mbMeetsSuggestedCrew = True
            End If
        End If

        Return lTotalCrew
    End Function


    Private Sub SetHscrBarMaxVals()
        Dim lF As Int32 = CInt(tpFrontArmor.PropertyValue)
        Dim lL As Int32 = CInt(tpLeftArmor.PropertyValue)
        Dim lR As Int32 = CInt(tpRightArmor.PropertyValue)
        Dim lB As Int32 = CInt(tpRearArmor.PropertyValue)
        Dim oHullTech As HullTech = Nothing
        Dim oArmorTech As ArmorTech = Nothing

        mbIgnoreScrollEvents = True

        If cboHull.ListIndex > -1 Then oHullTech = CType(goCurrentPlayer.GetTech(cboHull.ItemData(cboHull.ListIndex), ObjectType.eHullTech), HullTech)
        If cboArmor.ListIndex > -1 Then oArmorTech = CType(goCurrentPlayer.GetTech(cboArmor.ItemData(cboArmor.ListIndex), ObjectType.eArmorTech), ArmorTech)

        Dim lTotalCrew As Int32 = GetTotalCrewQuartersHull()

        tpFrontArmor.PropertyValue = 0 : tpFrontArmor.MaxValue = 0
        tpLeftArmor.PropertyValue = 0 : tpLeftArmor.MaxValue = 0
        tpRightArmor.PropertyValue = 0 : tpRightArmor.MaxValue = 0
        tpRearArmor.PropertyValue = 0 : tpRearArmor.MaxValue = 0
        Dim lSec As Int32 = CInt(tpSecurity.PropertyValue)
        tpSecurity.PropertyValue = 0 : tpSecurity.MaxValue = 0

        If oHullTech Is Nothing = False AndAlso oArmorTech Is Nothing = False Then
            tpFrontArmor.MaxValue = CInt(Math.Floor(oHullTech.GetHullAllotment(SlotType.eFront, SlotConfig.eArmorConfig) / oArmorTech.HullUsagePerPlate))
            'If hscrFront.MaxValue < lTotalCrew Then hscrFront.MaxValue = 0 Else hscrFront.MaxValue -= lTotalCrew
            tpLeftArmor.MaxValue = CInt(Math.Floor(oHullTech.GetHullAllotment(SlotType.eLeft, SlotConfig.eArmorConfig) / oArmorTech.HullUsagePerPlate))
            'If hscrLeft.MaxValue < lTotalCrew Then hscrLeft.MaxValue = 0 Else hscrLeft.MaxValue -= lTotalCrew
            tpRightArmor.MaxValue = CInt(Math.Floor(oHullTech.GetHullAllotment(SlotType.eRight, SlotConfig.eArmorConfig) / oArmorTech.HullUsagePerPlate))
            'If hscrRight.MaxValue < lTotalCrew Then hscrRight.MaxValue = 0 Else hscrRight.MaxValue -= lTotalCrew
            tpRearArmor.MaxValue = CInt(Math.Floor(oHullTech.GetHullAllotment(SlotType.eRear, SlotConfig.eArmorConfig) / oArmorTech.HullUsagePerPlate))
            'If hscrRear.MaxValue < lTotalCrew Then hscrRear.MaxValue = 0 Else hscrRear.MaxValue -= lTotalCrew
        End If

        'Ok, now remove per sides...
        Dim lFMax As Int32 = CInt(tpFrontArmor.MaxValue)
        Dim lLMax As Int32 = CInt(tpLeftArmor.MaxValue)
        Dim lRMax As Int32 = CInt(tpRightArmor.MaxValue)
        Dim lAMax As Int32 = CInt(tpRearArmor.MaxValue)

        Dim lCrewArmorRemoval As Int32 = lTotalCrew
        If oArmorTech Is Nothing = False Then lCrewArmorRemoval = CInt(Math.Ceiling(lTotalCrew / oArmorTech.HullUsagePerPlate))

        While lCrewArmorRemoval > 0 AndAlso lFMax > 0 AndAlso lLMax > 0 AndAlso lRMax > 0 AndAlso lAMax > 0
            Dim lMaxVal As Int32 = Math.Max(Math.Max(Math.Max(lFMax, lLMax), lRMax), lAMax)
            If lFMax = lMaxVal Then
                Dim lOtherMaxVal As Int32 = Math.Max(Math.Max(lLMax, lRMax), lAMax)
                Dim lDiff As Int32 = lMaxVal - lOtherMaxVal
                If lDiff = 0 Then lDiff = 1
                lDiff = Math.Min(lCrewArmorRemoval, lDiff)
                lCrewArmorRemoval -= lDiff
                lFMax -= lDiff
            ElseIf lLMax = lMaxVal Then
                Dim lOtherMaxVal As Int32 = Math.Max(Math.Max(lFMax, lRMax), lAMax)
                Dim lDiff As Int32 = lMaxVal - lOtherMaxVal
                If lDiff = 0 Then lDiff = 1
                lDiff = Math.Min(lCrewArmorRemoval, lDiff)
                lCrewArmorRemoval -= lDiff
                lLMax -= lDiff
            ElseIf lRMax = lMaxVal Then
                Dim lOtherMaxVal As Int32 = Math.Max(Math.Max(lFMax, lLMax), lAMax)
                Dim lDiff As Int32 = lMaxVal - lOtherMaxVal
                If lDiff = 0 Then lDiff = 1
                lDiff = Math.Min(lCrewArmorRemoval, lDiff)
                lCrewArmorRemoval -= lDiff
                lRMax -= lDiff
            ElseIf lAMax = lMaxVal Then
                Dim lOtherMaxVal As Int32 = Math.Max(Math.Max(lFMax, lRMax), lLMax)
                Dim lDiff As Int32 = lMaxVal - lOtherMaxVal
                If lDiff = 0 Then lDiff = 1
                lDiff = Math.Min(lCrewArmorRemoval, lDiff)
                lCrewArmorRemoval -= lDiff
                lAMax -= lDiff
            End If
        End While

        Dim lSecMax As Int32 = lFMax + lAMax + lLMax + lRMax

        'For lValue As Int32 = 0 To lTotalCrew - 1
        '	Dim lMaxVal As Int32 = Math.Max(Math.Max(Math.Max(lFMax, lLMax), lRMax), lAMax)
        '	If lFMax = lMaxVal Then
        '		lFMax -= 1
        '	ElseIf lLMax = lMaxVal Then
        '		lLMax -= 1
        '	ElseIf lRMax = lMaxVal Then
        '		lRMax -= 1
        '	ElseIf lAMax = lMaxVal Then
        '		lAMax -= 1
        '	End If
        'Next lValue
        If lFMax < 0 Then lFMax = 0
        If lLMax < 0 Then lLMax = 0
        If lRMax < 0 Then lRMax = 0
        If lAMax < 0 Then lAMax = 0
        tpFrontArmor.MaxValue = lFMax : tpLeftArmor.MaxValue = lLMax : tpRightArmor.MaxValue = lRMax : tpRearArmor.MaxValue = lAMax

        If tpFrontArmor.MaxValue >= lF Then tpFrontArmor.PropertyValue = lF Else tpFrontArmor.PropertyValue = tpFrontArmor.MaxValue
        If tpLeftArmor.MaxValue >= lL Then tpLeftArmor.PropertyValue = lL Else tpLeftArmor.PropertyValue = tpLeftArmor.MaxValue
        If tpRightArmor.MaxValue >= lR Then tpRightArmor.PropertyValue = lR Else tpRightArmor.PropertyValue = tpRightArmor.MaxValue
        If tpRearArmor.MaxValue >= lB Then tpRearArmor.PropertyValue = lB Else tpRearArmor.PropertyValue = tpRearArmor.MaxValue

        If oHullTech Is Nothing = False Then 'AndAlso oArmorTech Is Nothing = True Then
            lSecMax = CInt(Math.Floor(oHullTech.GetHullAllotment(SlotType.eNoChange, SlotConfig.eArmorConfig)))
            If oArmorTech Is Nothing = False Then
                lSecMax -= (lF + lL + lR + lB) * oArmorTech.HullUsagePerPlate
            End If
        End If

        'If oArmorTech Is Nothing = False Then lSecMax -= (CInt(tpFrontArmor.PropertyValue) + CInt(tpLeftArmor.PropertyValue) + CInt(tpRightArmor.PropertyValue) + CInt(tpRearArmor.PropertyValue)) * oArmorTech.HullUsagePerPlate
        Dim lBHtoR As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBetterHullToResidence)
        lSecMax \= lBHtoR
        If lSecMax < 0 Then lSecMax = 0
        tpSecurity.MaxValue = lSecMax
        If tpSecurity.MaxValue >= lSec Then tpSecurity.PropertyValue = lSec Else tpSecurity.PropertyValue = tpSecurity.MaxValue

        UpdateProductionScroller()

        mbIgnoreScrollEvents = False
    End Sub

    Private Sub UpdateProductionScroller()
        Dim oHullTech As HullTech = Nothing

        If cboHull.ListIndex > -1 Then oHullTech = CType(goCurrentPlayer.GetTech(cboHull.ItemData(cboHull.ListIndex), ObjectType.eHullTech), HullTech)

        If oHullTech Is Nothing = False Then
            If oHullTech.yTypeID = 2 Then
                Select Case oHullTech.ySubTypeID
                    Case 0, 1, 3, 5, 6, 7, 8, 9 ', 10
                        If (oHullTech.ModelID And 255) <> 148 Then
                            tpProduction.Enabled = True
                        Else
                            tpProduction.PropertyValue = 0
                            tpProduction.Enabled = False
                        End If
                    Case Else
                        tpProduction.PropertyValue = 0
                        tpProduction.Enabled = False
                End Select
            ElseIf oHullTech.yTypeID = 7 Then
                Select Case oHullTech.ySubTypeID
                    Case 0, 1, 3
                        tpProduction.Enabled = True
                    Case Else
                        tpProduction.PropertyValue = 0
                        tpProduction.Enabled = False
                End Select
            ElseIf oHullTech.yTypeID = 8 Then
                If oHullTech.ySubTypeID = 6 Then
                    tpProduction.Enabled = True
                Else
                    tpProduction.Enabled = False
                    tpProduction.PropertyValue = 0
                End If
            End If

            If tpProduction.Enabled = True Then
                Dim oArmorTech As ArmorTech = Nothing
                If cboArmor.ListIndex > -1 Then oArmorTech = CType(goCurrentPlayer.GetTech(cboArmor.ItemData(cboArmor.ListIndex), ObjectType.eArmorTech), ArmorTech)
                Dim lP As Int32 = CInt(tpProduction.PropertyValue)
                tpProduction.MaxValue = CInt(Math.Floor(oHullTech.GetHullAllotment(SlotType.eNoChange, SlotConfig.eArmorConfig)))
                If oArmorTech Is Nothing = False Then
                    tpProduction.MaxValue -= Math.Max(CInt((tpFrontArmor.PropertyValue + tpLeftArmor.PropertyValue + tpRightArmor.PropertyValue + tpRearArmor.PropertyValue) * oArmorTech.HullUsagePerPlate), 0)
                End If
                Dim lTotalCrew As Int32 = GetTotalCrewQuartersHull()
                If tpProduction.MaxValue < lTotalCrew Then tpProduction.MaxValue = 0 Else tpProduction.MaxValue -= lTotalCrew
                If tpProduction.MaxValue >= lP Then tpProduction.PropertyValue = lP Else tpProduction.PropertyValue = tpProduction.PropertyValue
            End If
        Else
            tpProduction.PropertyValue = 0
            tpProduction.Enabled = False
        End If

    End Sub

    Private Sub optActual_Click() Handles optActual.Click
        optExpected.Value = False
        optActual.Value = True
        RefreshAttributeBox()
    End Sub

    Private Sub optExpected_Click() Handles optExpected.Click
        optExpected.Value = True
        optActual.Value = False
        RefreshAttributeBox()
    End Sub

    Private Sub tp_ValueChanged(ByVal sPropName As String, ByRef ctl As ctlTechProp)
        If mbIgnoreScrollEvents = True Then Return

        If NewTutorialManager.TutorialOn = True Then
            If tpFrontArmor.MaxValue = tpFrontArmor.PropertyValue AndAlso tpLeftArmor.PropertyValue = tpLeftArmor.MaxValue AndAlso tpRightArmor.PropertyValue = tpRightArmor.MaxValue AndAlso tpRearArmor.MaxValue = tpRearArmor.PropertyValue Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.ePrototypeAllArmorSliderMax, -1, -1, -1, "")
            End If
        End If

        RefreshAttributeBox()
    End Sub

    Private Sub cboArmor_ItemSelected(ByVal lItemIndex As Integer) Handles cboArmor.ItemSelected
        SetHscrBarMaxVals()
        If NewTutorialManager.TutorialOn = True Then
            If lItemIndex > -1 Then NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.ePrototypeComponentSelected, cboArmor.ItemData(lItemIndex), ObjectType.eArmorTech, -1, "")
        End If
        RefreshAttributeBox()
    End Sub

    Private Sub cboEngine_ItemSelected(ByVal lItemIndex As Integer) Handles cboEngine.ItemSelected
        If NewTutorialManager.TutorialOn = True Then
            If lItemIndex > -1 Then NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.ePrototypeComponentSelected, cboEngine.ItemData(lItemIndex), ObjectType.eEngineTech, -1, "")
        End If
        RefreshAttributeBox()
    End Sub

    'Private Sub cboHangar_ItemSelected(ByVal lItemIndex As Integer) Handles cboHangar.ItemSelected
    '    RefreshAttributeBox()
    'End Sub

    Private Sub cboHull_ItemSelected(ByVal lItemIndex As Integer) Handles cboHull.ItemSelected
        SetWeaponPlaceState(False, False)
        'cboWeapon_ItemSelected(cboWeapon.ListIndex)
        'cboWeaponType_ItemSelected(cboWeaponType.ListIndex)
        cboWeapon.ListIndex = -1
        cboWeaponType.ListIndex = -1

        If NewTutorialManager.TutorialOn = True Then
            If lItemIndex > -1 Then NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.ePrototypeComponentSelected, cboHull.ItemData(lItemIndex), ObjectType.eHullTech, -1, "")
        End If

        Dim oHullTech As HullTech = Nothing
        If lItemIndex > -1 Then oHullTech = CType(goCurrentPlayer.GetTech(cboHull.ItemData(cboHull.ListIndex), ObjectType.eHullTech), HullTech)
        hulMain.SetFromHullTech(oHullTech)

        cboArmor.Clear() : cboEngine.Clear() : cboRadar.Clear() : cboShield.Clear()
        mlWpnUB = -1
        ReDim moWpns(-1)

        'Now... reset our indices...
        cboArmor.ListIndex = -1 : cboEngine.ListIndex = -1 : cboRadar.ListIndex = -1 : cboShield.ListIndex = -1
        tpFrontArmor.PropertyValue = 0 : tpLeftArmor.PropertyValue = 0 : tpRearArmor.PropertyValue = 0 : tpRightArmor.PropertyValue = 0

        'Now, enable/disable our items based on the hulltech...
        If oHullTech Is Nothing = False Then
            FillComponentLists()

            If oHullTech.HasSlotConfig(SlotConfig.eArmorConfig) = True Then
                If cboArmor.Enabled = False Then cboArmor.Enabled = True
            ElseIf cboArmor.Enabled = True Then
                cboArmor.Enabled = False : cboArmor.ListIndex = 0
            End If

            If oHullTech.HasSlotConfig(SlotConfig.eEngines) = True Then
                If cboEngine.Enabled = False Then cboEngine.Enabled = True
            ElseIf cboEngine.Enabled = True Then
                cboEngine.Enabled = False : cboEngine.ListIndex = 0
            End If

            'If oHullTech.HasSlotConfig(SlotConfig.eHangar) = True Then
            '    If cboHangar.Enabled = False Then cboHangar.Enabled = True
            'ElseIf cboHangar.Enabled = True Then
            '    cboHangar.Enabled = False : cboHangar.ListIndex = 0
            'End If

            If oHullTech.HasSlotConfig(SlotConfig.eRadar) = True Then
                If cboRadar.Enabled = False Then cboRadar.Enabled = True
            ElseIf cboRadar.Enabled = True Then
                cboRadar.Enabled = False : cboRadar.ListIndex = 0
            End If

            If oHullTech.HasSlotConfig(SlotConfig.eShields) = True Then
                If cboShield.Enabled = False Then cboShield.Enabled = True
            ElseIf cboShield.Enabled = True Then
                cboShield.Enabled = False : cboShield.ListIndex = 0
            End If
        Else
            cboArmor.ListIndex = 0 : cboArmor.Enabled = False
            cboEngine.ListIndex = 0 : cboEngine.Enabled = False
            'cboHangar.ListIndex = 0 : cboHangar.Enabled = False
            cboRadar.ListIndex = 0 : cboRadar.Enabled = False
            cboShield.ListIndex = 0 : cboShield.Enabled = False
        End If

        SetProductionLabel(oHullTech)

        SetHscrBarMaxVals()
        RefreshAttributeBox()
    End Sub

    Private Sub cboRadar_ItemSelected(ByVal lItemIndex As Integer) Handles cboRadar.ItemSelected
        If NewTutorialManager.TutorialOn = True Then
            If lItemIndex > -1 Then NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.ePrototypeComponentSelected, cboRadar.ItemData(lItemIndex), ObjectType.eRadarTech, -1, "")
        End If
        RefreshAttributeBox()
    End Sub

    Private Sub cboShield_ItemSelected(ByVal lItemIndex As Integer) Handles cboShield.ItemSelected
        If NewTutorialManager.TutorialOn = True Then
            If lItemIndex > -1 Then NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.ePrototypeComponentSelected, cboShield.ItemData(lItemIndex), ObjectType.eShieldTech, -1, "")
        End If
        RefreshAttributeBox()
    End Sub

    Private mbConfirmedLowCrew As Boolean = False
    Private Sub ConfirmLowCrewResult(ByVal lResult As MsgBoxResult)
        If lResult = MsgBoxResult.Yes Then
            mbConfirmedLowCrew = True
            btnDesign_Click(btnDesign.Caption)
        End If
    End Sub

    Private Sub btnDesign_Click(ByVal sName As String) Handles btnDesign.Click
        If goCurrentEnvir Is Nothing Then Return
        If mlEntityIndex = -1 Then
            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID AndAlso goCurrentEnvir.oEntity(X).bSelected Then
                    mlEntityIndex = X
                    Exit For
                End If
            Next X
        End If
        If mlEntityIndex = -1 Then Return

        If Me.txtPrototypeName.Caption.Trim.Length = 0 Then
            MyBase.moUILib.AddNotification("You must enter a name for this Prototype.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        If moTech Is Nothing = False AndAlso btnDesign.Caption.ToLower = "research" Then
            'Ok, simply submit this for research now
            Dim yData(13) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yData, 0)
            goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yData, 2)
            moTech.GetGUIDAsString.CopyTo(yData, 8)
            MyBase.moUILib.GetMsgSys.SendToPrimary(yData)

            If NewTutorialManager.TutorialOn = True Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eBuildOrderSent, moTech.ObjectID, moTech.ObjTypeID, 1, "")
            End If

            CloseMe()
            Return
        End If

        Dim oEngineTech As EngineTech = Nothing
        Dim oArmorTech As ArmorTech = Nothing
        'Dim oHangarTech As HangarTech = Nothing
        Dim oHullTech As HullTech = Nothing
        Dim oRadarTech As RadarTech = Nothing
        Dim oShieldTech As ShieldTech = Nothing

        Dim lForeArmor As Int32 = CInt(tpFrontArmor.PropertyValue)
        Dim lRearArmor As Int32 = CInt(tpRearArmor.PropertyValue)
        Dim lLeftArmor As Int32 = CInt(tpLeftArmor.PropertyValue)
        Dim lRightArmor As Int32 = CInt(tpRightArmor.PropertyValue)

        If cboEngine.ListIndex > -1 Then oEngineTech = CType(goCurrentPlayer.GetTech(cboEngine.ItemData(cboEngine.ListIndex), ObjectType.eEngineTech), EngineTech)
        If cboArmor.ListIndex > -1 Then oArmorTech = CType(goCurrentPlayer.GetTech(cboArmor.ItemData(cboArmor.ListIndex), ObjectType.eArmorTech), ArmorTech)
        'If cboHangar.ListIndex > -1 Then oHangarTech = CType(goCurrentPlayer.GetTech(cboHangar.ItemData(cboHangar.ListIndex), ObjectType.eHangarTech), HangarTech)
        If cboHull.ListIndex > -1 Then oHullTech = CType(goCurrentPlayer.GetTech(cboHull.ItemData(cboHull.ListIndex), ObjectType.eHullTech), HullTech)
        If cboRadar.ListIndex > -1 Then oRadarTech = CType(goCurrentPlayer.GetTech(cboRadar.ItemData(cboRadar.ListIndex), ObjectType.eRadarTech), RadarTech)
        If cboShield.ListIndex > -1 Then oShieldTech = CType(goCurrentPlayer.GetTech(cboShield.ItemData(cboShield.ListIndex), ObjectType.eShieldTech), ShieldTech)

        'Validate it first...
        If ValidateDesign() = False Then Return

        If mbMeetsSuggestedCrew = False AndAlso mbConfirmedLowCrew = False Then
            Dim oMsgBox As New frmMsgBox(goUILib, "This prototype does not meet the suggested crew requirement. This provides a higher chance for pirate raids to destroy this prototype." & vbCrLf & vbCrLf & "Are you sure you wish to use this prototype?", MsgBoxStyle.YesNo, "Confirm Prototype")
            oMsgBox.Visible = True
            AddHandler oMsgBox.DialogClosed, AddressOf ConfirmLowCrewResult
            Return
        End If

        lForeArmor = Math.Max(0, lForeArmor)
        lRearArmor = Math.Max(0, lRearArmor)
        lLeftArmor = Math.Max(0, lLeftArmor)
        lRightArmor = Math.Max(0, lRightArmor)

        'Now, get the data...
        Dim yMsg(81 + ((mlWpnUB + 1) * 7)) As Byte
        Dim lPos As Int32 = 0
        Dim sTemp As String = Mid$(txtPrototypeName.Caption, 1, 20)

        Dim oResearchGuid As Base_GUID = goCurrentEnvir.oEntity(mlEntityIndex)
        'ok, if the entity production is station...
        If goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = ProductionType.eSpaceStationSpecial Then
            'Ok, need to find research lab... try our lstFac...
            Dim frmSelFac As frmSelectFac = CType(MyBase.moUILib.GetWindow("frmSelectFac"), frmSelectFac)
            If frmSelFac Is Nothing = False Then
                Dim oChild As StationChild = frmSelFac.GetCurrentChild()
                If oChild Is Nothing = False Then
                    oResearchGuid = New Base_GUID
                    oResearchGuid.ObjectID = oChild.lChildID
                    oResearchGuid.ObjTypeID = oChild.iChildTypeID
                End If
            End If
        End If

        System.BitConverter.GetBytes(GlobalMessageCode.eSubmitComponentPrototype).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(ObjectType.ePrototype).CopyTo(yMsg, lPos) : lPos += 2

        If moTech Is Nothing = False Then
            System.BitConverter.GetBytes(moTech.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
        End If

        oResearchGuid.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6

        'Name
        System.Text.ASCIIEncoding.ASCII.GetBytes(sTemp).CopyTo(yMsg, lPos) : lPos += 20

        'Engine ID
        If oEngineTech Is Nothing = False Then
            System.BitConverter.GetBytes(oEngineTech.ObjectID).CopyTo(yMsg, lPos)
        Else : System.BitConverter.GetBytes(CInt(-1)).CopyTo(yMsg, lPos)
        End If
        lPos += 4

        'ArmorID
        If oArmorTech Is Nothing = False Then
            System.BitConverter.GetBytes(oArmorTech.ObjectID).CopyTo(yMsg, lPos)
        Else : System.BitConverter.GetBytes(CInt(-1)).CopyTo(yMsg, lPos)
        End If
        lPos += 4

        'HangarID
        'If oHangarTech Is Nothing = False Then
        '    System.BitConverter.GetBytes(oHangarTech.ObjectID).CopyTo(yMsg, lPos)
        'Else : System.BitConverter.GetBytes(CInt(-1)).CopyTo(yMsg, lPos)
        'End If
        'lPos += 4

        'HullID
        If oHullTech Is Nothing = False Then
            System.BitConverter.GetBytes(oHullTech.ObjectID).CopyTo(yMsg, lPos)
        Else : System.BitConverter.GetBytes(CInt(-1)).CopyTo(yMsg, lPos)
        End If
        lPos += 4

        'RadarID
        If oRadarTech Is Nothing = False Then
            System.BitConverter.GetBytes(oRadarTech.ObjectID).CopyTo(yMsg, lPos)
        Else : System.BitConverter.GetBytes(CInt(-1)).CopyTo(yMsg, lPos)
        End If
        lPos += 4

        'ShieldID
        If oShieldTech Is Nothing = False Then
            System.BitConverter.GetBytes(oShieldTech.ObjectID).CopyTo(yMsg, lPos)
        Else : System.BitConverter.GetBytes(CInt(-1)).CopyTo(yMsg, lPos)
        End If
        lPos += 4

        'Armor Sides...
        System.BitConverter.GetBytes(lForeArmor).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lRearArmor).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lLeftArmor).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lRightArmor).CopyTo(yMsg, lPos) : lPos += 4

        '4 bytes for the Prod Slot Count
        Dim lProdCount As Int32 = CInt(tpProduction.PropertyValue)
        System.BitConverter.GetBytes(lProdCount).CopyTo(yMsg, lPos) : lPos += 4
        '4 bytes for security
        Dim lSecurity As Int32 = CInt(tpSecurity.PropertyValue)
        System.BitConverter.GetBytes(lSecurity).CopyTo(yMsg, lPos) : lPos += 4


        '4 bytes for the weapon count...
        Dim lCnt As Int32 = mlWpnUB + 1
        System.BitConverter.GetBytes(lCnt).CopyTo(yMsg, lPos) : lPos += 4

        'Weapon Data...
        '  WeaponTechID (4)
        '  SlotX (1)
        '  SlotY (1)
        '  WeaponGroupType (1)
        For X As Int32 = 0 To mlWpnUB
            System.BitConverter.GetBytes(moWpns(X).WeaponTechID).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = moWpns(X).SlotX : lPos += 1
            yMsg(lPos) = moWpns(X).SlotY : lPos += 1
            yMsg(lPos) = moWpns(X).WeaponGroupTypeID : lPos += 1
        Next X

        MyBase.moUILib.SendMsgToPrimary(yMsg)

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eSubmitDesignClick, ObjectType.ePrototype, -1, -1, "")
        End If

        MyBase.moUILib.AddNotification("Prototype Design Submitted", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        CloseMe()
    End Sub

    Private Sub btnPlaceWpn_Click(ByVal sName As String) Handles btnPlaceWpn.Click
        If btnPlaceWpn.Caption = "Cancel" Then
            SetWeaponPlaceState(False, False)
        Else

            If cboWeaponGroupType.ListIndex > -1 Then
                Dim lType As Int32 = cboWeaponGroupType.ItemData(cboWeaponGroupType.ListIndex)
                If lType = WeaponGroupType.BombWeaponGroup Then
                    If cboWeaponType.ListIndex > -1 Then
                        Dim lWpnType As Int32 = cboWeaponType.ItemData(cboWeaponType.ListIndex)
                        If lWpnType <> WeaponClassType.eBomb AndAlso lWpnType <> WeaponClassType.eEnergyBeam Then
                            MyBase.moUILib.AddNotification("Bombs grouping can only be set on Solid Beam or Bomb weapons.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return
                        End If

                        If cboHull.ListIndex > -1 Then
                            Dim oHullTech As HullTech = CType(goCurrentPlayer.GetTech(cboHull.ItemData(cboHull.ListIndex), ObjectType.eHullTech), HullTech)
                            If oHullTech Is Nothing = False Then
                                If oHullTech.yTypeID <> 1 AndAlso oHullTech.ySubTypeID <> 3 Then
                                    MyBase.moUILib.AddNotification("Bombs cannot be placed on non-frigate hulls.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                                    Return
                                End If
                            End If
                        End If
                    End If
                ElseIf cboWeaponType.ListIndex > -1 Then
                    Dim lWpnType As Int32 = cboWeaponType.ItemData(cboWeaponType.ListIndex)
                    If lWpnType = WeaponClassType.eBomb Then
                        MyBase.moUILib.AddNotification("Bombs must be grouped as Bombs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return
                    End If

                End If

            End If

            SetWeaponPlaceState(True, False)
            If NewTutorialManager.TutorialOn = True Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.ePrototypePlaceButtonClick, -1, -1, -1, "")
            End If
        End If
    End Sub

    Public Sub ViewResults(ByRef oTech As PrototypeTech, ByVal lProdFactor As Int32)
        If oTech Is Nothing Then Return

        mbIgnoreScrollEvents = True

        If mbIgnoreCheckFilter = False Then
            chkFilter.Value = False
            chkFilter_Click()
        End If
        

        moTech = oTech

        Me.btnDesign.Enabled = False

        Me.Left = 10
        Me.Top = 10

        With oTech
            'Ok, select our hull first
            cboHull.FindComboItemData(.HullID)
            'Now the other cbo's
            Me.cboArmor.FindComboItemData(.ArmorID)
            Me.cboEngine.FindComboItemData(.EngineID)
            'Me.cboHangar.FindComboItemData(.HangarID)
            Me.cboRadar.FindComboItemData(.RadarID)
            Me.cboShield.FindComboItemData(.ShieldID)

            'Weapon doesn't matter... now... set our hscr's
            Me.tpFrontArmor.PropertyValue = .ForeArmorUnits
            Me.tpLeftArmor.PropertyValue = .LeftArmorUnits
            Me.tpRearArmor.PropertyValue = .RearArmorUnits
            Me.tpRightArmor.PropertyValue = .RightArmorUnits
            UpdateProductionScroller()
            Me.tpProduction.PropertyValue = .ProductionHull
            Me.tpSecurity.PropertyValue = .MaxCrew

            'set our name
            Me.txtPrototypeName.Caption = .PrototypeName

            'Set our moWpns array
            mlWpnUB = .lWeaponUB
            ReDim moWpns(.lWeaponUB)
            For X As Int32 = 0 To .lWeaponUB
                moWpns(X).SlotX = .oWeapons(X).SlotX
                moWpns(X).SlotY = .oWeapons(X).SlotY
                moWpns(X).WeaponGroupTypeID = .oWeapons(X).WeaponGroupTypeID
                moWpns(X).WeaponTechID = .oWeapons(X).WeaponID
            Next X
            If mbWpnGroupNumsSet = False Then FillWeaponGroupNums()
        End With

        If oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eComponentResearching AndAlso HasAliasedRights(AliasingRights.eAddResearch) = True Then
            btnDesign.Caption = "Research"
            btnDesign.Enabled = True
        End If
        If gbAliased = False Then
            btnDelete.Visible = True 'oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched
        End If

        'Now, set actual results option to true
        optActual_Click()

        'Now... what state is the tech in?
        If oTech.ComponentDevelopmentPhase > Base_Tech.eComponentDevelopmentPhase.eComponentDesign Then
            'Ok, show it's research and production cost
            If mfrmResCost Is Nothing Then mfrmResCost = New frmProdCost(MyBase.moUILib, "frmResCost")
            With mfrmResCost
                .Visible = True
                .Left = Me.Left + Me.Width + 10
                .Top = Me.Top '+ Me.Height + 10
                .SetFromProdCost(oTech.oResearchCost, lProdFactor, True, 0, 0)
            End With

            If mfrmProdCost Is Nothing Then mfrmProdCost = New frmProdCost(MyBase.moUILib, "frmProdCost")
            With mfrmProdCost
                .Visible = True
                .Left = mfrmResCost.Left
                .Top = mfrmResCost.Top + mfrmResCost.Height + 10
                .SetFromProdCost(oTech.oProductionCost, 1000, False, 0, 0)
            End With
        ElseIf oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
            If mfrmResCost Is Nothing Then mfrmResCost = New frmProdCost(MyBase.moUILib, "frmResCost")
            With mfrmResCost
                .Visible = True
                .Left = Me.Left + Me.Width + 10
                .Top = Me.Top '+ Me.Height + 10
                .SetFromFailureCode(oTech.ErrorCodeReason)
            End With
        End If

        mbIgnoreScrollEvents = False
    End Sub

    Private mbWpnGroupNumsSet As Boolean = False
    Private Sub FillWeaponGroupNums()
        For X As Int32 = 0 To mlWpnUB
            Dim yType As SlotType
            Dim yConfig As SlotConfig
            Dim lGroup As Int32
            hulMain.GetHullSlotValues(moWpns(X).SlotX, moWpns(X).SlotY, yType, yConfig, lGroup)
            moWpns(X).lGroupNum = lGroup
        Next X
        mbWpnGroupNumsSet = True
    End Sub

    Private Sub hulMain_HullSlotClick(ByVal lIndexX As Integer, ByVal lIndexY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles hulMain.HullSlotClick
        If mbInPlaceWpn = True Then
            Dim yType As SlotType
            Dim lConfig As SlotConfig
            Dim lGroup As Int32
            Dim lIdx As Int32 = -1

            hulMain.GetHullSlotValues(lIndexX, lIndexY, yType, lConfig, lGroup)

            If NewTutorialManager.TutorialOn = True Then
                Dim sParms() As String = {lGroup.ToString}
                If goUILib.CommandAllowedWithParms(True, hulMain.GetFullName, sParms, False) = False Then Return
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.ePrototypeWeaponPlaced, cboWeapon.ItemData(cboWeapon.ListIndex), ObjectType.eWeaponTech, lGroup, "")
            End If

            'now... what... Find out if there is a weapon already there..
            For X As Int32 = 0 To mlWpnUB
                If moWpns(X).lGroupNum = lGroup Then
                    'Yes, so we will just use this moWpn object
                    lIdx = X
                    Exit For
            End If
            Next X
            If lIdx = -1 Then
                mlWpnUB += 1
                ReDim Preserve moWpns(mlWpnUB)
                lIdx = mlWpnUB
            End If

            With moWpns(lIdx)
                .lGroupNum = lGroup
                .SlotX = CByte(lIndexX)
                .SlotY = CByte(lIndexY)
                .WeaponGroupTypeID = CType(cboWeaponGroupType.ItemData(cboWeaponGroupType.ListIndex), WeaponGroupType)
                .WeaponTechID = cboWeapon.ItemData(cboWeapon.ListIndex)
            End With
            SortWeaponGroups()

            'finally, if all goes well, set our status
            If frmMain.mbShiftKeyDown = False Then SetWeaponPlaceState(False, False)
            'SetWeaponPlaceState(False, False)
        ElseIf mbInRemoveWpn = True Then

            If mlWpnUB <> -1 Then
                Dim yType As SlotType
                Dim lConfig As SlotConfig
                Dim lGroup As Int32
                Dim lIdx As Int32 = -1

                hulMain.GetHullSlotValues(lIndexX, lIndexY, yType, lConfig, lGroup)

                'Find the weapon... it will share the same group number
                For X As Int32 = 0 To mlWpnUB
                    If moWpns(X).lGroupNum = lGroup Then
                        'Yes, so we will just use this moWpn object
                        lIdx = X
                    End If
                Next X
                SortWeaponGroups()
                If lIdx <> -1 Then
                    For X As Int32 = lIdx To mlWpnUB - 1
                        moWpns(X) = moWpns(X + 1)
                    Next X
                    mlWpnUB -= 1
                    ReDim Preserve moWpns(mlWpnUB)
                End If

            End If
            'finally, if all goes well, set our status
            'SetWeaponPlaceState(False, True)
            If frmMain.mbShiftKeyDown = False Then SetWeaponPlaceState(False, True)
        End If
    End Sub

    Private Sub btnRemoveWpn_Click(ByVal sName As String) Handles btnRemoveWpn.Click
        If btnRemoveWpn.Caption = "Cancel" Then
            SetWeaponPlaceState(False, True)
        Else : SetWeaponPlaceState(True, True)
        End If
    End Sub

    Private Sub SortWeaponGroups()
        Dim lIdxes() As Int32
        ReDim lIdxes(mlWpnUB)

        Dim lLowGroup As Int32
        Dim lLastLowGroup As Int32
        Dim lIdx As Int32

        lLowGroup = Int32.MaxValue
        lLastLowGroup = Int32.MinValue

        For X As Int32 = 0 To mlWpnUB
            lIdx = -1

            For Y As Int32 = 0 To mlWpnUB
                If moWpns(Y).lGroupNum < lLowGroup AndAlso moWpns(Y).lGroupNum > lLastLowGroup Then
                    lIdx = Y
                    lLowGroup = moWpns(Y).lGroupNum
                End If
            Next Y

            'Ok, we should have one
            lIdxes(X) = lIdx
            lLastLowGroup = lLowGroup
            lLowGroup = Int32.MaxValue
        Next X

        'Now, reset our array
        Dim oTmpArr(mlWpnUB) As WeaponPlacement
        For X As Int32 = 0 To mlWpnUB
            oTmpArr(X) = moWpns(lIdxes(X))
            'moWpns(X) = oTmpArr(lIdxes(X))
        Next
        moWpns = oTmpArr
    End Sub

    Private Function ValidateDesign() As Boolean
        Dim oHullTech As HullTech = Nothing
        Dim oEngineTech As EngineTech = Nothing
        Dim oArmorTech As ArmorTech = Nothing
        Dim oRadarTech As RadarTech = Nothing
        Dim oShieldTech As ShieldTech = Nothing

        Dim lPowerUsed As Int32 = 0

        Dim fHullPerSlot As Single

        'First, check our hull
        If cboHull.ListIndex > -1 Then
            oHullTech = CType(goCurrentPlayer.GetTech(cboHull.ItemData(cboHull.ListIndex), ObjectType.eHullTech), HullTech)
            If oHullTech Is Nothing Then
                MyBase.moUILib.AddNotification("You must select a Hull for this prototype!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
        Else
            MyBase.moUILib.AddNotification("You must select a Hull for this prototype!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If

        lPowerUsed += oHullTech.PowerRequired

        fHullPerSlot = CSng(oHullTech.HullSize / oHullTech.TotalSlots)

        'Now, get our engine...
        If cboEngine.ListIndex > -1 Then
            oEngineTech = CType(goCurrentPlayer.GetTech(cboEngine.ItemData(cboEngine.ListIndex), ObjectType.eEngineTech), EngineTech)

            If oEngineTech Is Nothing = False Then
                If oEngineTech.HullRequired > oHullTech.GetHullAllotment(SlotType.eNoChange, SlotConfig.eEngines) Then
                    MyBase.moUILib.AddNotification("Engines are too big for this hull.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                End If
            End If
        End If

        Dim fFrontArmor As Single = oHullTech.GetHullAllotment(SlotType.eFront, SlotConfig.eArmorConfig)
        Dim fLeftArmor As Single = oHullTech.GetHullAllotment(SlotType.eLeft, SlotConfig.eArmorConfig)
        Dim fRearArmor As Single = oHullTech.GetHullAllotment(SlotType.eRear, SlotConfig.eArmorConfig)
        Dim fRightArmor As Single = oHullTech.GetHullAllotment(SlotType.eRight, SlotConfig.eArmorConfig)

        If cboArmor.ListIndex > -1 Then
            oArmorTech = CType(goCurrentPlayer.GetTech(cboArmor.ItemData(cboArmor.ListIndex), ObjectType.eArmorTech), ArmorTech)
        End If

        Dim lFA As Int32 = CInt(tpFrontArmor.PropertyValue)
        Dim lLA As Int32 = CInt(tpLeftArmor.PropertyValue)
        Dim lAA As Int32 = CInt(tpRearArmor.PropertyValue)
        Dim lRA As Int32 = CInt(tpRightArmor.PropertyValue)

        Dim lProdSlots As Int32 = CInt(tpProduction.PropertyValue)

        If oArmorTech Is Nothing = True AndAlso (lFA > 0 OrElse lLA > 0 OrElse lAA > 0 OrElse lRA > 0) Then
            MyBase.moUILib.AddNotification("In order to equip armor plating, you must select an armor component.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If

        Dim fTotalArmorUsed As Single = 0.0F
        If oArmorTech Is Nothing = False Then
            If lFA * oArmorTech.HullUsagePerPlate > fFrontArmor Then
                MyBase.moUILib.AddNotification("Front Side Armor exceeds hull allocated for armor.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If lLA * oArmorTech.HullUsagePerPlate > fLeftArmor Then
                MyBase.moUILib.AddNotification("Left Side Armor exceeds hull allocated for armor.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If lAA * oArmorTech.HullUsagePerPlate > fRearArmor Then
                MyBase.moUILib.AddNotification("Rear Side Armor exceeds hull allocated for armor.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If lRA * oArmorTech.HullUsagePerPlate > fRightArmor Then
                MyBase.moUILib.AddNotification("Right Side Armor exceeds hull allocated for armor.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            fTotalArmorUsed = (lFA + lLA + lAA + lRA) * oArmorTech.HullUsagePerPlate
        End If

        Dim lSecurity As Int32 = GetTotalCrewQuartersHull() 'CInt(tpSecurity.PropertyValue)
        'lSecurity *= goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBetterHullToResidence)
        ' lSecurity += GetTotalCrewQuartersHull()

        If lProdSlots + fTotalArmorUsed + lSecurity > oHullTech.GetHullAllotment(SlotType.eNoChange, SlotConfig.eArmorConfig) Then
            MyBase.moUILib.AddNotification("Production Slots, Crew, Security and Armor Slots exceed Armor Slot Allotment!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If

        If cboRadar.ListIndex > -1 Then
            oRadarTech = CType(goCurrentPlayer.GetTech(cboRadar.ItemData(cboRadar.ListIndex), ObjectType.eRadarTech), RadarTech)
            If oRadarTech Is Nothing = False Then
                If oRadarTech.HullRequired > oHullTech.GetHullAllotment(SlotType.eNoChange, SlotConfig.eRadar) Then
                    MyBase.moUILib.AddNotification("Radar component is too big for this hull.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                End If

                lPowerUsed += oRadarTech.PowerRequired
            End If
        End If

        If cboShield.ListIndex > -1 Then
            oShieldTech = CType(goCurrentPlayer.GetTech(cboShield.ItemData(cboShield.ListIndex), ObjectType.eShieldTech), ShieldTech)
            If oShieldTech Is Nothing = False Then
                If oShieldTech.HullRequired > oHullTech.GetHullAllotment(SlotType.eNoChange, SlotConfig.eShields) Then
                    MyBase.moUILib.AddNotification("Shield component is too big for this hull.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                End If

                lPowerUsed += oShieldTech.PowerRequired
            End If
        End If

        'Now, for weapons
        Dim oWpn As WeaponTech = Nothing
        For X As Int32 = 0 To mlWpnUB
            oWpn = CType(goCurrentPlayer.GetTech(moWpns(X).WeaponTechID, ObjectType.eWeaponTech), WeaponTech)
            If oWpn Is Nothing = False Then
                Dim lHull As Int32 = oWpn.HullRequired
                If moWpns(X).WeaponGroupTypeID = WeaponGroupType.PointDefenseWeaponGroup Then lHull = CInt(lHull * 0.5F)
                If lHull > oHullTech.GetWeaponHullAllotment(moWpns(X).lGroupNum) Then
                    MyBase.moUILib.AddNotification("Weapon Group " & moWpns(X).lGroupNum & " (" & oWpn.WeaponName & ") is too big for that slot grouping.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                End If

                lPowerUsed += oWpn.PowerRequired
            End If
        Next X

        'Ok, if this is not a facility or the facility is a space station... or the modelid is 138
        If oHullTech.yTypeID <> 2 OrElse oHullTech.ySubTypeID = 9 OrElse (oHullTech.ModelID And 255) = 138 Then
            If oEngineTech Is Nothing = True AndAlso lPowerUsed > 0 Then
                MyBase.moUILib.AddNotification("This design requires power but there is no engine to provide it.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            ElseIf oEngineTech Is Nothing = False AndAlso lPowerUsed > oEngineTech.PowerProd Then
                MyBase.moUILib.AddNotification("The power required for this design exceeds the power generated.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
        End If


        If mbUnitImmobile = True Then
            MyBase.moUILib.AddNotification("This design will be unable to move due to thrust vs. hull size.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If

        Return True
    End Function

    Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
        If moTech Is Nothing Then Return
        If btnDelete.Caption = "Delete" Then
            btnDelete.Caption = "CONFIRM"
        Else
            If moTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
                'Delete the design - permanently
                Dim yMsg(8) As Byte
                Dim lPos As Int32 = 0
                System.BitConverter.GetBytes(GlobalMessageCode.eSetArchiveState).CopyTo(yMsg, lPos) : lPos += 2
                moTech.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
                yMsg(8) = 255
                MyBase.moUILib.SendMsgToPrimary(yMsg)

                For X As Int32 = 0 To glEntityDefUB
                    If glEntityDefIdx(X) > -1 Then
                        If goEntityDefs(X).PrototypeID = moTech.ObjectID Then
                            glEntityDefIdx(X) = -1
                            Exit For
                        End If
                    End If
                Next X
            Else
                'Delete the design
                Dim yMsg(7) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eDeleteDesign).CopyTo(yMsg, 0)
                moTech.GetGUIDAsString.CopyTo(yMsg, 2)
                MyBase.moUILib.SendMsgToPrimary(yMsg)
            End If
            CloseMe()
            ''Delete the design
            'Dim yMsg(7) As Byte
            'System.BitConverter.GetBytes(GlobalMessageCode.eDeleteDesign).CopyTo(yMsg, 0)
            'moTech.GetGUIDAsString.CopyTo(yMsg, 2)
            'MyBase.moUILib.SendMsgToPrimary(yMsg)
            'CloseMe()
        End If
    End Sub

    Private Sub frmPrototypeBuilder_OnRender() Handles Me.OnRender
        'Ok, first, is moTech nothing?
        If moTech Is Nothing = False Then
            'Ok, now, do the values used on the form currently match the tech?
            Dim bChanged As Boolean = False
            With moTech
                'components
                Dim lHullID As Int32 = -1
                Dim lEngineID As Int32 = -1
                Dim lRadarID As Int32 = -1
                Dim lShieldID As Int32 = -1
                Dim lArmorID As Int32 = -1

                If cboHull.ListIndex <> -1 Then lHullID = cboHull.ItemData(cboHull.ListIndex)
                If cboEngine.ListIndex <> -1 Then lEngineID = cboEngine.ItemData(cboEngine.ListIndex)
                If cboRadar.ListIndex <> -1 Then lRadarID = cboRadar.ItemData(cboRadar.ListIndex)
                If cboShield.ListIndex <> -1 Then lShieldID = cboShield.ItemData(cboShield.ListIndex)
                If cboArmor.ListIndex <> -1 Then lArmorID = cboArmor.ItemData(cboArmor.ListIndex)

                If lHullID < 1 Then lHullID = -1
                If lEngineID < 1 Then lEngineID = -1
                If lRadarID < 1 Then lRadarID = -1
                If lShieldID < 1 Then lShieldID = -1
                If lArmorID < 1 Then lArmorID = -1

                bChanged = lHullID <> .HullID OrElse lEngineID <> .EngineID OrElse lRadarID <> .RadarID OrElse _
                  lShieldID <> .ShieldID OrElse lArmorID <> .ArmorID

                If bChanged = False Then
                    If txtPrototypeName.Caption <> .PrototypeName AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eInvalidDesign AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
                        bChanged = True
                    End If
                End If

                If bChanged = False Then
                    'armor sides
                    bChanged = .ForeArmorUnits <> tpFrontArmor.PropertyValue OrElse .LeftArmorUnits <> tpLeftArmor.PropertyValue OrElse _
                      .RightArmorUnits <> tpRightArmor.PropertyValue OrElse .RearArmorUnits <> tpRearArmor.PropertyValue OrElse mlWpnUB <> .lWeaponUB OrElse _
                      .ProductionHull <> tpProduction.PropertyValue OrElse .MaxCrew <> tpSecurity.PropertyValue

                    'weapons
                    If bChanged = False Then
                        For X As Int32 = 0 To mlWpnUB
                            Dim bFound As Boolean = False
                            For Y As Int32 = 0 To .lWeaponUB
                                If .oWeapons(Y).WeaponGroupTypeID = moWpns(X).WeaponGroupTypeID AndAlso .oWeapons(Y).WeaponID = moWpns(X).WeaponTechID Then
                                    If .oWeapons(Y).SlotX = moWpns(X).SlotX AndAlso .oWeapons(Y).SlotY = moWpns(X).SlotY Then
                                        bFound = True
                                        Exit For
                                    End If
                                End If
                            Next Y
                            If bFound = False Then
                                bChanged = True
                                Exit For
                            End If
                        Next X
                    End If
                End If

                If bChanged = True Then
                    If btnDesign.Caption <> "Redesign" Then btnDesign.Caption = "Redesign"
                    If btnDesign.Enabled = False Then btnDesign.Enabled = True
                Else
                    If btnDesign.Caption <> "Research" Then btnDesign.Caption = "Research"
                    If moTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eInvalidDesign AndAlso btnDesign.Enabled = True Then btnDesign.Enabled = False
                End If
            End With
        End If
    End Sub

    'Private Sub txtArmor_TextChanged() Handles txtForeArmor.TextChanged, txtLeftArmor.TextChanged, txtRearArmor.TextChanged, txtRightArmor.TextChanged, txtProduction.TextChanged
    '    If mbIgnoreScrollEvents = True Then Return

    '    mbIgnoreScrollEvents = True

    '    With txtForeArmor
    '        If .Caption <> "" Then
    '            Dim lTemp As Int32 = CInt(Val(.Caption))
    '            If lTemp.ToString <> .Caption Then
    '                .Caption = lTemp.ToString
    '            End If
    '        End If
    '    End With
    '    With txtLeftArmor
    '        If .Caption <> "" Then
    '            Dim lTemp As Int32 = CInt(Val(.Caption))
    '            If lTemp.ToString <> .Caption Then
    '                .Caption = lTemp.ToString
    '            End If
    '        End If
    '    End With
    '    With txtRearArmor
    '        If .Caption <> "" Then
    '            Dim lTemp As Int32 = CInt(Val(.Caption))
    '            If lTemp.ToString <> .Caption Then
    '                .Caption = lTemp.ToString
    '            End If
    '        End If
    '    End With
    '    With txtRightArmor
    '        If .Caption <> "" Then
    '            Dim lTemp As Int32 = CInt(Val(.Caption))
    '            If lTemp.ToString <> .Caption Then
    '                .Caption = lTemp.ToString
    '            End If
    '        End If
    '    End With
    '    With txtProduction
    '        If .Caption <> "" Then
    '            Dim lTemp As Int32 = CInt(Val(.Caption))
    '            If lTemp.ToString <> .Caption Then
    '                .Caption = lTemp.ToString
    '            End If
    '        End If
    '    End With

    '    Dim lValue As Int32 = CInt(Val(txtForeArmor.Caption))
    '    If lValue < hscrFront.MinValue Then lValue = hscrFront.MinValue
    '    If lValue > hscrFront.MaxValue Then lValue = hscrFront.MaxValue
    '    hscrFront.Value = lValue

    '    lValue = CInt(Val(txtLeftArmor.Caption))
    '    If lValue < hscrLeft.MinValue Then lValue = hscrLeft.MinValue
    '    If lValue > hscrLeft.MaxValue Then lValue = hscrLeft.MaxValue
    '    tpLeftArmor.PropertyValue = lValue

    '    lValue = CInt(Val(txtRearArmor.Caption))
    '    If lValue < hscrRear.MinValue Then lValue = hscrRear.MinValue
    '    If lValue > hscrRear.MaxValue Then lValue = hscrRear.MaxValue
    '    tpRearArmor.PropertyValue = lValue

    '    lValue = CInt(Val(txtRightArmor.Caption))
    '    If lValue < hscrRight.MinValue Then lValue = hscrRight.MinValue
    '    If lValue > hscrRight.MaxValue Then lValue = hscrRight.MaxValue
    '    tpRightArmor.PropertyValue = lValue

    '    lValue = CInt(Val(txtProduction.Caption))
    '    If lValue < hscrProduction.MinValue Then lValue = hscrProduction.MinValue
    '    If lValue > hscrProduction.MaxValue Then lValue = hscrProduction.MaxValue
    '    tpProduction.PropertyValue = lValue

    '    RefreshAttributeBox()

    '    mbIgnoreScrollEvents = False

    'End Sub

	Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
		If goTutorial Is Nothing Then goTutorial = New TutorialManager()
		goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.ePrototypeBuilder)
	End Sub

	Private Sub SetProductionLabel(ByRef oHullTech As HullTech)
        If oHullTech Is Nothing = False Then
            If oHullTech.yTypeID = 2 Then
                'MSC: We only care about the mesh portion of the modelid
                Dim iMesh As Int32 = (oHullTech.ModelID And 255S)
                Select Case oHullTech.ySubTypeID
                    Case 0  'CC
                        If tpProduction.PropertyName <> "Command Center:" Then tpProduction.PropertyName = "Command Center:"
                        Return
                    Case 1 'Mine
                        If (oHullTech.ModelID And 255) <> 148 Then
                            If tpProduction.PropertyName <> "Mining Rate:" Then tpProduction.PropertyName = "Mining Rate:"
                            Return
                        End If
                    Case 3 'personnel
                        If iMesh = 6 Then
                            If tpProduction.PropertyName <> "Enlisted Training:" Then tpProduction.PropertyName = "Enlisted Training:"
                        ElseIf iMesh = 7 Then
                            If tpProduction.PropertyName <> "Officer Training:" Then tpProduction.PropertyName = "Officer Training:"
                        End If
                        Return
                    Case 4 'Power 
                        If tpProduction.PropertyName <> "" Then tpProduction.PropertyName = ""
                        Return
                    Case 5 'Production
                        If iMesh = 13 OrElse iMesh = 102 Then
                            If tpProduction.PropertyName <> "Aerial Production:" Then tpProduction.PropertyName = "Aerial Production:"
                        ElseIf iMesh = 137 Then
                            If tpProduction.PropertyName <> "Naval Production:" Then tpProduction.PropertyName = "Naval Production:"
                        Else
                            If tpProduction.PropertyName <> "Ground Production:" Then tpProduction.PropertyName = "Ground Production:"
                        End If
                        Return
                    Case 6 'Refining
                        If tpProduction.PropertyName <> "Refining:" Then tpProduction.PropertyName = "Refining:"
                        Return
                    Case 7 'Research
                        If tpProduction.PropertyName <> "Research:" Then tpProduction.PropertyName = "Research:"
                        Return
                    Case 8 'Residence
                        If tpProduction.PropertyName <> "Residence:" Then tpProduction.PropertyName = "Residence:"
                        Return
                    Case 9 'SpaceStation
                        If iMesh = 24 Then
                            If tpProduction.PropertyName <> "Galactic Trading:" Then tpProduction.PropertyName = "Galactic Trading:"
                        Else
                            If tpProduction.PropertyName <> "Space Production:" Then tpProduction.PropertyName = "Space Production:"
                        End If

                        Return
                    Case 10 'tradepost
                        If tpProduction.PropertyName <> "No Production" Then tpProduction.PropertyName = "No Production"
                        Return
                End Select
            ElseIf oHullTech.yTypeID = 7 Then
                Select Case oHullTech.ySubTypeID
                    Case 0, 1
                        If tpProduction.PropertyName <> "Facility Production:" Then tpProduction.PropertyName = "Facility Production:"
                        Return
                    Case 3
                        If tpProduction.PropertyName <> "Mining:" Then tpProduction.PropertyName = "Mining:"
                        Return
                End Select
            ElseIf oHullTech.yTypeID = 8 Then
                If oHullTech.ySubTypeID = 6 Then
                    If tpProduction.PropertyName <> "Facility Production:" Then tpProduction.PropertyName = "Facility Production:"
                    Return
                End If
            End If
        End If
        If tpProduction.PropertyName <> "No Production" Then tpProduction.PropertyName = "No Production"
	End Sub

    Private mbIgnoreCheckFilter As Boolean = False
    Private Sub chkFilter_Click() Handles chkFilter.Click
        If mbIgnoreCheckFilter = True Then Return
        mbIgnoreCheckFilter = True

        MyBase.moUILib.GetMsgSys.LoadArchived()
        FillHullList()
        FillComponentLists()

        If moTech Is Nothing = False Then ViewResults(moTech, 100)

        mbIgnoreCheckFilter = False
    End Sub

	Private Sub txtPrototypeName_TextChanged() Handles txtPrototypeName.TextChanged
		If NewTutorialManager.TutorialOn = True Then
			NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.ePrototypeNameChange, -1, -1, -1, txtPrototypeName.Caption)
		End If
	End Sub

    Private Sub cboWeaponGroupType_ItemSelected(ByVal lItemIndex As Integer) Handles cboWeaponGroupType.ItemSelected
        cboWeapon_ItemSelected(cboWeapon.ListIndex)
    End Sub
End Class