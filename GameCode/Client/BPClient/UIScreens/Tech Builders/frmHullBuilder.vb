Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class frmHullBuilder
    Inherits UIWindow

    'fraHullDesignDetails Panel
    Private WithEvents fraHullDesignDetails As UIWindow
    'Private lblStructHP As UILabel
    'Private WithEvents txtStructHP As UITextBox
    'Private lblHullSize As UILabel
    'Private WithEvents txtHullSize As UITextBox
    Private tpStructHP As ctlTechProp
    Private tpHullSize As ctlTechProp
    Private lblHullName As UILabel
    Private WithEvents txtHullName As UITextBox
    Private lblStructMat As UILabel
    'Private WithEvents cboStructMat As UIComboBox
    Private lblStructMatItem As UILabel

    'fraSelectFrame
    Private fraSelectFrame As UIWindow
    Private WithEvents lscType As UILabelScroller
    Private WithEvents lscSubType As UILabelScroller
    Private WithEvents lscFrame As UILabelScroller
    'New customization stuff
    Private lscDiffuse As UILabelScroller
    Private lscNormal As UILabelScroller
    Private lscIllum As UILabelScroller

    'fraBaseAttrs
    Private fraBaseAttrs As UIWindow
    Private lblHullSizeRange As UILabel
    Private WithEvents chkGround As UICheckBox
    Private WithEvents chkNaval As UICheckBox
    Private WithEvents chkAtmospheric As UICheckBox
    Private WithEvents chkSpace As UICheckBox
    Private lblCrewReqs As UILabel

    'fraSideDist
    Private fraSideDist As UIWindow
    Private lblForeDist As UILabel
    Private lblLeftDist As UILabel
    Private lblRearDist As UILabel
    Private lblRightDist As UILabel
    Private lblAllArcDist As UILabel

    'fraSlotPalette
    Private fraSlotPalette As UIWindow
    Private WithEvents optArmor As UIOption
    Private WithEvents optCrewQuarters As UIOption
    Private WithEvents optCargo As UIOption
    Private WithEvents optRadarComm As UIOption
    Private WithEvents optEngines As UIOption
    Private WithEvents optHangar As UIOption
    Private WithEvents optShield As UIOption
    Private WithEvents optWeapons As UIOption
    Private WithEvents txtGroupNum As UITextBox 
    Private WithEvents optHangarDoor As UIOption
    Private WithEvents txtHgrDrNum As UITextBox

    'Hull Slots...
    Private WithEvents hulMain As UIHullSlots

    'Other controls...
    Private WithEvents btnSwitchView As UIButton
    Private txtHullAllotment As UITextBox
    Private lblHullAllotments As UILabel
    Private optHullSlotSize As UICheckBox
    Private txtErrors As UITextBox
    Private lblErrors As UILabel
    Private WithEvents btnDesign As UIButton
    Private WithEvents btnCancel As UIButton
    Private WithEvents btnClear As UIButton
    Private WithEvents btnDelete As UIButton
    Private WithEvents btnHelp As UIButton
    Private WithEvents btnShowMinDisplay As UIButton
    Private WithEvents btnRename As UIButton
    Private lstHullLookup As UIListBox
    Private lblHullLookup As UILabel

    Private lblSpecialTrait As UILabel
    Private lblSpecialTraitText As UILabel

    'Form's Variables...
    Private mbLoading As Boolean = True
    Private mlCrewReqPerc As Int32
    Private mlEntityIndex As Int32 = -1
    Private mlMinHull As Int32
    Private mlMaxHull As Int32

    Private moModelViewTex As Texture = Nothing
    Private mbModelView As Boolean = False
    Private mfLastCX As Single = 0
    Private mlLastCY As Int32 = 1000
    Private mfLastCZ As Single = -1000
    Private mlInitialZ As Int32 = Int32.MinValue

    Private moTech As HullTech = Nothing
    Private mfrmResCost As frmProdCost = Nothing
    Private mfrmProdCost As frmProdCost = Nothing
    Private mbIgnoreLabelScrollEvents As Boolean = False

    Private mbIgnoreCheckEvents As Boolean = False
    Private mbSpaceLocked As Boolean = False
    Private mbGroundLocked As Boolean = False
    Private mbNavalLocked As Boolean = False
    Private mbAtmosLocked As Boolean = False

    'Private mbSpaceStation As Boolean = False
    'Private mbShip As Boolean = False
    Private myHullRequirements As UIHullSlots.eyHullRequirements

    Public Shared sParms() As String = Nothing

    Private moShader As ModelShader = Nothing

    Private mlMineralID As Int32 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmHullBuilder initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eHullBuilder
            .ControlName = "frmHullBuilder"
            '.Left = 5
            .Left = 10 ' (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - 475
            '.Top = 5
            .Top = 10 '(oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 357
            '.Width = oUILib.oDevice.PresentationParameters.BackBufferWidth - 10
            .Width = 950
            '.Height = oUILib.oDevice.PresentationParameters.BackBufferHeight - 10
            .Height = 715
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .DrawBorder = False
            .Moveable = False
            .mbAcceptReprocessEvents = True
        End With

        'fraSlotPalette initial props
        fraSlotPalette = New UIWindow(oUILib)
        With fraSlotPalette
            .ControlName = "fraSlotPalette"
            .Left = 490
            .Top = 120 '187
            .Width = 135
            .Height = 205 '177
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Caption = "Slot Palette"
            .BorderLineWidth = 1
            .Moveable = False
        End With
        Me.AddChild(CType(fraSlotPalette, UIControl))

        'optArmor initial props
        optArmor = New UIOption(oUILib)
        With optArmor
            .ControlName = "optArmor"
            .Left = 5
            .Top = 10
            .Width = 125
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Reserve for Armor" & vbCrLf & "and Crew Quarters"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
            .DisplayType = UIOption.eOptionTypes.eSmallOption
        End With
        fraSlotPalette.AddChild(CType(optArmor, UIControl))

        'optCrewQuarters initial props
        optCrewQuarters = New UIOption(oUILib)
        With optCrewQuarters
            .ControlName = "optCrewQuarters"
            .Left = 5
            .Top = 28
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = False
            .Caption = "Crew Quarters"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UIOption.eOptionTypes.eSmallOption
        End With
        fraSlotPalette.AddChild(CType(optCrewQuarters, UIControl))

        'optCargo initial props
        optCargo = New UIOption(oUILib)
        With optCargo
            .ControlName = "optCargo"
            .Left = 5
            .Top = 46
            .Width = 80
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Cargo Bay"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UIOption.eOptionTypes.eSmallOption
        End With
        fraSlotPalette.AddChild(CType(optCargo, UIControl))

        'optRadarComm initial props
        optRadarComm = New UIOption(oUILib)
        With optRadarComm
            .ControlName = "optRadarComm"
            .Left = 5
            .Top = 64
            .Width = 104
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Radar/Comms"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UIOption.eOptionTypes.eSmallOption
        End With
        fraSlotPalette.AddChild(CType(optRadarComm, UIControl))

        'optEngines initial props
        optEngines = New UIOption(oUILib)
        With optEngines
            .ControlName = "optEngines"
            .Left = 5
            .Top = 82
            .Width = 65
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Engines"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UIOption.eOptionTypes.eSmallOption
        End With
        fraSlotPalette.AddChild(CType(optEngines, UIControl))

        'optHangar initial props
        optHangar = New UIOption(oUILib)
        With optHangar
            .ControlName = "optHangar"
            .Left = 5
            .Top = 100
            .Width = 61
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hangar"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UIOption.eOptionTypes.eSmallOption
        End With
        fraSlotPalette.AddChild(CType(optHangar, UIControl))

        'optShield initial props
        optShield = New UIOption(oUILib)
        With optShield
            .ControlName = "optShield"
            .Left = 5
            .Top = 118
            .Width = 61
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Shields"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UIOption.eOptionTypes.eSmallOption
        End With
        fraSlotPalette.AddChild(CType(optShield, UIControl))

        'optWeapons initial props
        optWeapons = New UIOption(oUILib)
        With optWeapons
            .ControlName = "optWeapons"
            .Left = 5
            .Top = 136
            .Width = 76
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Weapons"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UIOption.eOptionTypes.eSmallOption
        End With
        fraSlotPalette.AddChild(CType(optWeapons, UIControl))

        'txtGroupNum initial props
        txtGroupNum = New UITextBox(oUILib)
        With txtGroupNum
            .ControlName = "txtGroupNum"
            .Left = 90
            .Top = 136
            .Width = 32
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "1"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 2
            .BorderColor = muSettings.InterfaceBorderColor
            .bNumericOnly = True
            .ToolTipText = ""
        End With
        fraSlotPalette.AddChild(CType(txtGroupNum, UIControl))

        ''optFuelBay initial props
        'optFuelBay = New UIOption(oUILib)
        'With optFuelBay
        '    .ControlName = "optFuelBay"
        '    .Left = 5
        '    .Top = 154
        '    .Width = 69
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Fuel Bay"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(6, DrawTextFormat)
        '    .Value = False
        '    .DisplayType = UIOption.eOptionTypes.eSmallOption
        'End With
        'fraSlotPalette.AddChild(CType(optFuelBay, UIControl))

        'optHangarDoor initial props
        optHangarDoor = New UIOption(oUILib)
        With optHangarDoor
            .ControlName = "optHangarDoor"
            .Left = 5
            .Top = 172
            .Width = 73
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Bay Door"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UIOption.eOptionTypes.eSmallOption
        End With
        fraSlotPalette.AddChild(CType(optHangarDoor, UIControl))

        'txtHgrDrNum initial props
        txtHgrDrNum = New UITextBox(oUILib)
        With txtHgrDrNum
            .ControlName = "txtHgrDrNum"
            .Left = 90
            .Top = 172
            .Width = 32
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "1"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 2
            .BorderColor = muSettings.InterfaceBorderColor
            .bNumericOnly = True
        End With
        fraSlotPalette.AddChild(CType(txtHgrDrNum, UIControl))

        'fraSideDist initial props
        fraSideDist = New UIWindow(oUILib)
        With fraSideDist
            .ControlName = "fraSideDist"
            .Left = 490
            .Top = 15 '87
            .Width = 135
            .Height = 90
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Caption = "Side Distribution"
            .BorderLineWidth = 1
            .Moveable = False
        End With
        Me.AddChild(CType(fraSideDist, UIControl))

        'lblForeDist initial props
        lblForeDist = New UILabel(oUILib)
        With lblForeDist
            .ControlName = "lblForeDist"
            .Left = 5
            .Top = 5
            .Width = 130
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Forward: 0%"
            .ForeColor = System.Drawing.Color.FromArgb(255, 0, 0, 192)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraSideDist.AddChild(CType(lblForeDist, UIControl))

        'lblLeftDist initial props
        lblLeftDist = New UILabel(oUILib)
        With lblLeftDist
            .ControlName = "lblLeftDist"
            .Left = 5
            .Top = 22
            .Width = 130
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Left: 0%"
            .ForeColor = System.Drawing.Color.FromArgb(255, 192, 0, 192)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraSideDist.AddChild(CType(lblLeftDist, UIControl))

        'lblRearDist initial props
        lblRearDist = New UILabel(oUILib)
        With lblRearDist
            .ControlName = "lblRearDist"
            .Left = 5
            .Top = 39
            .Width = 130
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Rear: 0%"
            .ForeColor = System.Drawing.Color.FromArgb(255, 0, 192, 0)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraSideDist.AddChild(CType(lblRearDist, UIControl))

        'lblRightDist initial props
        lblRightDist = New UILabel(oUILib)
        With lblRightDist
            .ControlName = "lblRightDist"
            .Left = 5
            .Top = 56
            .Width = 130
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Right: 0%"
            .ForeColor = System.Drawing.Color.FromArgb(255, 192, 0, 0)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraSideDist.AddChild(CType(lblRightDist, UIControl))

        'lblAllArcDist initial props
        lblAllArcDist = New UILabel(oUILib)
        With lblAllArcDist
            .ControlName = "lblAllArcDist"
            .Left = 5
            .Top = 72
            .Width = 130
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "All-Arc: 0%"
            .ForeColor = System.Drawing.Color.FromArgb(255, 192, 192, 0)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraSideDist.AddChild(CType(lblAllArcDist, UIControl))

        'fraBaseAttrs initial props
        fraBaseAttrs = New UIWindow(oUILib)
        With fraBaseAttrs
            .ControlName = "fraBaseAttrs"
            .Left = 287
            .Top = 86 '87
            .Width = 200
            .Height = 91 '90
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Caption = "Frame Base Attributes"
            .BorderLineWidth = 1
            .Moveable = False
        End With
        Me.AddChild(CType(fraBaseAttrs, UIControl))

        'lblHullSizeRange initial props
        lblHullSizeRange = New UILabel(oUILib)
        With lblHullSizeRange
            .ControlName = "lblHullSizeRange"
            .Left = 5
            .Top = 10
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraBaseAttrs.AddChild(CType(lblHullSizeRange, UIControl))

        'chkGround initial props
        chkGround = New UICheckBox(oUILib)
        With chkGround
            .ControlName = "chkGround"
            .Left = 5
            .Top = 25
            .Width = 104
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Ground-Based"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            .ToolTipText = "Check if the object moves along land"
        End With
        fraBaseAttrs.AddChild(CType(chkGround, UIControl))

        'chkNaval initial props
        chkNaval = New UICheckBox(oUILib)
        With chkNaval
            .ControlName = "chkNaval"
            .Left = 138
            .Top = 25
            .Width = 55
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Naval"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            .ToolTipText = "Check if the object can move along water/liquids"
        End With
        fraBaseAttrs.AddChild(CType(chkNaval, UIControl))

        'chkAtmospheric initial props
        chkAtmospheric = New UICheckBox(oUILib)
        With chkAtmospheric
            .ControlName = "chkAtmospheric"
            .Left = 5
            .Top = 45
            .Width = 92
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Atmospheric"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            .ToolTipText = "Check if the object will be able to fly in atmospheres"
        End With
        fraBaseAttrs.AddChild(CType(chkAtmospheric, UIControl))

        'chkSpace initial props
        chkSpace = New UICheckBox(oUILib)
        With chkSpace
            .ControlName = "chkSpace"
            .Left = 138
            .Top = 45
            .Width = 58
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Space"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            .ToolTipText = "Check if the object can exist in space"
        End With
        fraBaseAttrs.AddChild(CType(chkSpace, UIControl))

        'fraSelectFrame initial props
        fraSelectFrame = New UIWindow(oUILib)
        With fraSelectFrame
            .ControlName = "fraSelectFrame"
            .Left = 5
            .Top = 86 '87
            .Width = 279
            .Height = 91 '90
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Caption = "Select a Frame"
            .BorderLineWidth = 1
            .Moveable = False
        End With
        Me.AddChild(CType(fraSelectFrame, UIControl))

        'lscType initial props
        lscType = New UILabelScroller(oUILib)
        With lscType
            .ControlName = "lscType"
            .Left = 15
            .Top = 8 '15
            .Width = 250
            .Height = 18
            .Enabled = True
            .Visible = True
        End With
        fraSelectFrame.AddChild(CType(lscType, UIControl))

        'lscSubType initial props
        lscSubType = New UILabelScroller(oUILib)
        With lscSubType
            .ControlName = "lscSubType"
            .Left = 15
            .Top = 29 '40
            .Width = 250
            .Height = 18
            .Enabled = True
            .Visible = True
        End With
        fraSelectFrame.AddChild(CType(lscSubType, UIControl))

        'lscFrame initial props
        lscFrame = New UILabelScroller(oUILib)
        With lscFrame
            .ControlName = "lscFrame"
            .Left = 15
            .Top = 50 '65
            .Width = 250
            .Height = 18
            .Enabled = True
            .Visible = True
        End With
        fraSelectFrame.AddChild(CType(lscFrame, UIControl))

        'lscDiffuse initial props
        lscDiffuse = New UILabelScroller(oUILib)
        With lscDiffuse
            .ControlName = "lscDiffuse"
            .Left = 15
            .Top = 71
            .Width = 80 '250
            .Height = 18
            .Enabled = True
            .Visible = True
            .ToolTipText = "Changes the overall texture of the model. Only available when viewing with the High Quality lighting setting on."
        End With
        fraSelectFrame.AddChild(CType(lscDiffuse, UIControl))

        'lscNormal initial props
        lscNormal = New UILabelScroller(oUILib)
        With lscNormal
            .ControlName = "lscNormal"
            .Left = lscDiffuse.Left + lscDiffuse.Width + 5
            .Top = lscDiffuse.Top
            .Width = 80 '250
            .Height = 18
            .Enabled = True
            .Visible = True
            .ToolTipText = "Changes the bumpiness of the model. Only available when viewing with the High Quality lighting setting on."
        End With
        fraSelectFrame.AddChild(CType(lscNormal, UIControl))

        'lscIllum initial props
        lscIllum = New UILabelScroller(oUILib)
        With lscIllum
            .ControlName = "lscIllum"
            .Left = lscNormal.Left + lscNormal.Width + 5
            .Top = lscDiffuse.Top
            .Width = 80 '250
            .Height = 18
            .Enabled = True
            .Visible = True
            .ToolTipText = "Changes the illumination texture of the model. Only available when viewing with the High Quality lighting setting on."
        End With
        fraSelectFrame.AddChild(CType(lscIllum, UIControl))

        'fraHullDesignDetails initial props
        fraHullDesignDetails = New UIWindow(oUILib)
        With fraHullDesignDetails
            .ControlName = "fraHullDesignDetails"
            .Left = 5
            .Top = 15
            .Width = 485 '490 ' 564
            .Height = 61
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
            .Caption = "Hull Design Details"
            .BorderLineWidth = 1
        End With
        Me.AddChild(CType(fraHullDesignDetails, UIControl))

        ''lblStructHP initial props
        'lblStructHP = New UILabel(oUILib)
        'With lblStructHP
        '    .ControlName = "lblStructHP"
        '    .Left = 220 '265
        '    .Top = 35
        '    .Width = 70 '130 '135
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Hitpoints:"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(4, DrawTextFormat)
        'End With
        'fraHullDesignDetails.AddChild(CType(lblStructHP, UIControl))

        ''txtStructHP initial props
        'txtStructHP = New UITextBox(oUILib)
        'With txtStructHP
        '    .ControlName = "txtStructHP"
        '    .Left = 300 '285 '345 '405
        '    .Top = 35
        '    .Width = 150
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "0"
        '    .ForeColor = muSettings.InterfaceTextBoxForeColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(4, DrawTextFormat)
        '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '    .MaxLength = 9
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .bNumericOnly = True
        'End With
        'fraHullDesignDetails.AddChild(CType(txtStructHP, UIControl))
        tpStructHP = New ctlTechProp(oUILib)
        With tpStructHP
            .ControlName = "tpStructHP"
            .Enabled = True
            .Height = 18
            .Left = 220
            .MaxValue = Int32.MaxValue
            .MinValue = 0
            .bNoMaxValue = True
            .blAbsoluteMaximum = Int32.MaxValue
            .PropertyLocked = False
            .PropertyName = "Hitpoints:"
            .PropertyValue = 0
            .Top = 35
            .Visible = True
            .Width = 250
            .SetExtendedProps(False, 75, 80, False)
            .ToolTipText = "Indicates how many hitpoints of damage the structure of this hull can withstand."
        End With
        fraHullDesignDetails.AddChild(CType(tpStructHP, UIControl))

        'lblHullSize initial props
        'lblHullSize = New UILabel(oUILib)
        'With lblHullSize
        '    .ControlName = "lblHullSize"
        '    .Left = 10
        '    .Top = 35
        '    .Width = 40 '80
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Size:"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(4, DrawTextFormat)
        'End With
        'fraHullDesignDetails.AddChild(CType(lblHullSize, UIControl))

        ''txtHullSize initial props
        'txtHullSize = New UITextBox(oUILib)
        'With txtHullSize
        '    .ControlName = "txtHullSize"
        '    .Left = 60 '100
        '    .Top = 35
        '    .Width = 130 ' 150
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "0"
        '    .ForeColor = muSettings.InterfaceTextBoxForeColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(4, DrawTextFormat)
        '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '    .MaxLength = 0
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .bNumericOnly = True
        'End With
        'fraHullDesignDetails.AddChild(CType(txtHullSize, UIControl))
        tpHullSize = New ctlTechProp(oUILib)
        With tpHullSize
            .ControlName = "tpHullSize"
            .Enabled = True
            .Height = 18
            .Left = 10
            .MaxValue = 200
            .MinValue = 0
            '.bNoMaxValue = True
            '.blAbsoluteMaximum = Int32.MaxValue
            .PropertyLocked = False
            .PropertyName = "Size:"
            .PropertyValue = 0
            .TextMaxLength = 10
            .Top = 35
            .Visible = True
            .Width = 180 '170
            .SetExtendedProps(False, 45, 70, False)
            .ToolTipText = "Indicates the hull size of this hull."
        End With
        fraHullDesignDetails.AddChild(CType(tpHullSize, UIControl))

        'lblHullName initial props
        lblHullName = New UILabel(oUILib)
        With lblHullName
            .ControlName = "lblHullName"
            .Left = 10
            .Top = 13
            .Width = 40 '80
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Name:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraHullDesignDetails.AddChild(CType(lblHullName, UIControl))

        'txtHullName initial props
        txtHullName = New UITextBox(oUILib)
        With txtHullName
            .ControlName = "txtHullName"
            .Left = 60 '100
            .Top = 13
            .Width = 140 '130
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
        fraHullDesignDetails.AddChild(CType(txtHullName, UIControl))

        'lblStructMat initial props
        lblStructMat = New UILabel(oUILib)
        With lblStructMat
            .ControlName = "lblStructMat"
            .Left = 220 '205 '265
            .Top = 13
            .Width = 70 '130 '135
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Material:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        fraHullDesignDetails.AddChild(CType(lblStructMat, UIControl))

        'btnShowMinDisplay initial props
        btnShowMinDisplay = New UIButton(oUILib)
        With btnShowMinDisplay
            .ControlName = "btnShowMinDisplay"
            .Left = 440
            .Top = 13
            .Width = 40
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Set"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
            .ToolTipText = "Click to view the mineral details window"
        End With
        fraHullDesignDetails.AddChild(CType(btnShowMinDisplay, UIControl))

        'hulMain initial props
        hulMain = New UIHullSlots(oUILib)
        With hulMain
            .AutoRefresh = False
            .ControlName = "hulMain"
            .Enabled = True
            .Height = 480
            .Left = 5
            '.ToolTipText = "Click to set slot to selected type in Palette."
            .Top = fraSelectFrame.Top + fraSelectFrame.Height + 6
            .Visible = True
            .Width = 480
            .AutoRefresh = True
        End With
        Me.AddChild(CType(hulMain, UIControl))

        'lblCrewReqs initial props
        lblCrewReqs = New UILabel(oUILib)
        With lblCrewReqs
            .ControlName = "lblCrewReqs"
            .ToolTipText = "Suggested crew to have onboard to help protect against boarding parties, pirate raids and agents."
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
        Me.AddChild(CType(lblCrewReqs, UIControl))

        'btnSwitchView initial props
        btnSwitchView = New UIButton(oUILib)
        With btnSwitchView
            .ControlName = "btnSwitchView"
            .Left = hulMain.Left + (hulMain.Width \ 2) - 55
            .Top = hulMain.Top + hulMain.Height + 10
            .Width = 110
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Switch View"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSwitchView, UIControl))

        'btnClear initial props
        btnClear = New UIButton(oUILib)
        With btnClear
            .ControlName = "btnClear"
            .Left = btnSwitchView.Left + btnSwitchView.Width + 10
            .Top = btnSwitchView.Top
            .Width = 110
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Clear"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnClear, UIControl))

        'lblHullAllotments initial props
        lblHullAllotments = New UILabel(oUILib)
        With lblHullAllotments
            .ControlName = "lblHullAllotments"
            .Left = fraSideDist.Left + fraSideDist.Width + 5
            .Top = 5
            .Width = 111
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hull Allotments:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblHullAllotments, UIControl))

        'txtHullAllotment initial props
        txtHullAllotment = New UITextBox(oUILib)
        With txtHullAllotment
            .ControlName = "txtHullAllotment"
            .Left = lblHullAllotments.Left
            .Top = lblHullAllotments.Top + lblHullAllotments.Height + 5
            .Width = 300
            .Height = ((fraSlotPalette.Top + fraSlotPalette.Height) - .Top - 25) \ 2
            .Enabled = True
            .Visible = True
            .Caption = ""
            .BorderColor = muSettings.InterfaceBorderColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Left Or DrawTextFormat.Top
            .BackColorEnabled = muSettings.InterfaceFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(-9868951)
            .MaxLength = 0
            .Locked = True
            .MultiLine = True
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(txtHullAllotment, UIControl))

        'lblHullLookup initial props
        lblHullLookup = New UILabel(oUILib)
        With lblHullLookup
            .ControlName = "lblHullLookup"
            .Left = lblHullAllotments.Left
            .Top = txtHullAllotment.Top + txtHullAllotment.Height + 5
            .Width = 131
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hull Usage Lookup"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblHullLookup, UIControl))

        'lstHullLookup initial props
        lstHullLookup = New UIListBox(oUILib)
        With lstHullLookup
            .ControlName = "lstHullLookup"
            .Left = lblHullLookup.Left
            .Top = lblHullLookup.Top + lblHullLookup.Height + 3
            .Width = 300
            .Height = txtHullAllotment.Height
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(lstHullLookup, UIControl))

        'optHullSlotSize initial props
        optHullSlotSize = New UICheckBox(oUILib)
        With optHullSlotSize
            .ControlName = "optHullSlotSize"
            .Left = 10
            .Top = btnSwitchView.Top + (btnSwitchView.Height \ 2) - 9
            .Width = 140
            .Height = 18
            .Enabled = False
            .Visible = True
            .Caption = "= 0 Hull Per Slot"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Right Or DrawTextFormat.VerticalCenter
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallBlock
            .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eCheck_Disabled)
        End With
        Me.AddChild(CType(optHullSlotSize, UIControl))

        'lblSpecialTrait initial props
        lblSpecialTrait = New UILabel(oUILib)
        With lblSpecialTrait
            .ControlName = "lblSpecialTrait"
            .Left = fraSlotPalette.Left
            .Top = fraSlotPalette.Top + fraSlotPalette.Height + 7
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Special Trait: "
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSpecialTrait, UIControl))

        'lblSpecialTraitText initial props
        lblSpecialTraitText = New UILabel(oUILib)
        With lblSpecialTraitText
            .ControlName = "lblSpecialTraitText"
            .Left = lblSpecialTrait.Left + lblSpecialTrait.Width
            .Top = lblSpecialTrait.Top
            .Width = txtHullAllotment.Width + (txtHullAllotment.Left - .Left)
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "None"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblSpecialTraitText, UIControl))

        'lblErrors initial props
        lblErrors = New UILabel(oUILib)
        With lblErrors
            .ControlName = "lblErrors"
            .Left = fraSlotPalette.Left
            .Top = lblSpecialTrait.Top + lblSpecialTrait.Height + 5 'fraSlotPalette.Top + fraSlotPalette.Height + 5
            .Width = 111
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Design Flaws:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblErrors, UIControl))

        'txtErrors initial props
        txtErrors = New UITextBox(oUILib)
        With txtErrors
            .ControlName = "txtErrors"
            .Left = lblErrors.Left
            .Top = lblErrors.Top + lblErrors.Height + 5
            .Width = txtHullAllotment.Width + (txtHullAllotment.Left - .Left)
            .Height = (hulMain.Top + hulMain.Height) - .Top
            .Enabled = True
            .Visible = True
            .Caption = ""
            .BorderColor = muSettings.InterfaceBorderColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Left Or DrawTextFormat.Top
            .BackColorEnabled = muSettings.InterfaceFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(-9868951)
            .MaxLength = 0
            .Locked = True
            .MultiLine = True
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(txtErrors, UIControl))

        'btnHelp initial props
        btnHelp = New UIButton(oUILib)
        With btnHelp
            .ControlName = "btnHelp"
            .Left = Me.Width - 24 - Me.BorderLineWidth
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

        'btnCancel initial props
        btnCancel = New UIButton(oUILib)
        With btnCancel
            .ControlName = "btnCancel"
            .Left = (txtErrors.Left + txtErrors.Width) - 100
            .Top = txtErrors.Top + txtErrors.Height + 10
            .Width = 100
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Cancel"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnCancel, UIControl))

        'btnDesign initial props
        btnDesign = New UIButton(oUILib)
        With btnDesign
            .ControlName = "btnDesign"
            .Left = btnCancel.Left - 110
            .Top = btnCancel.Top
            .Width = 100
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Submit Design"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnDesign, UIControl))

        'btnDelete initial props
        btnDelete = New UIButton(oUILib)
        With btnDelete
            .ControlName = "btnDelete"
            .Left = btnDesign.Left - 110 'fraSlotPalette.Left
            .Top = btnCancel.Top
            .Width = 100
            .Height = 32
            .Enabled = True
            .Visible = False
            .Caption = "Delete Design"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnDelete, UIControl))

        'btnRename initial props
        btnRename = New UIButton(oUILib)
        With btnRename
            .ControlName = "btnRename"
            .Left = btnDelete.Left - 110 'txtHullName.Left + txtHullName.Width + 5
            .Top = btnDesign.Top
            .Width = 100
            .Height = btnDelete.Height
            .Enabled = True
            .Visible = False
            .Caption = "Rename"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnRename, UIControl))

        ''cboStructMat initial props
        'cboStructMat = New UIComboBox(oUILib)
        'With cboStructMat
        '    .ControlName = "cboStructMat"
        '    .Left = 300 '285 '345 '405
        '    .Top = 13
        '    .Width = 150
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .FillColor = muSettings.InterfaceTextBoxFillColor
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        '    .DropDownListBorderColor = muSettings.InterfaceBorderColor

        '    '.ToolTipText = "We are not quite sure of the best way to engineer this."
        '    'If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.PlayerKnowsProperty("HARDNESS") = True Then
        '    .ToolTipText = "Hard materials work best for this."
        '    'End If
        '    .mbAcceptReprocessEvents = True
        'End With
        'fraHullDesignDetails.AddChild(CType(cboStructMat, UIControl))
        'lblStructMatItem initial props
        lblStructMatItem = New UILabel(oUILib)
        With lblStructMatItem
            .ControlName = "lblStructMatItem"
            .Left = 300
            .Top = 13
            .Width = 150
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Unselected"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
        End With
        fraHullDesignDetails.AddChild(CType(lblStructMatItem, UIControl))

        AddHandler tpHullSize.PropertyValueChanged, AddressOf tpHullSize_ValueChange
        AddHandler tpStructHP.PropertyValueChanged, AddressOf tpStructHP_ValueChanged

        AddHandler lscDiffuse.ItemChanged, AddressOf DiffuseScroller_ItemChanged
        AddHandler lscNormal.ItemChanged, AddressOf AppearanceScroller_ItemChanged
        AddHandler lscIllum.ItemChanged, AddressOf AppearanceScroller_ItemChanged

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        glCurrentEnvirView = CurrentView.eHullResearch

        If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
        goFullScreenBackground = Nothing

        'Dim oTmpWin As UIWindow = MyBase.moUILib.GetWindow("frmChat")
        'If oTmpWin Is Nothing = False Then oTmpWin.Visible = False
        'oTmpWin = Nothing

        mbLoading = False

        FillTypesList()
        'FillStaticList()

    End Sub

    Private Sub hulMain_HullSlotClick(ByVal lIndexX As Integer, ByVal lIndexY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles hulMain.HullSlotClick
        If lButton <> MouseButtons.Left AndAlso lButton <> MouseButtons.Right Then Return
        Dim lVal As SlotConfig
        Dim lGrp As Int32 = 1

        If NewTutorialManager.TutorialOn = True Then
            If goUILib.CommandAllowed(True, hulMain.GetFullName) = False Then Return
        End If

        If lButton = MouseButtons.Left Then
            If optArmor.Value = True Then
                lVal = SlotConfig.eArmorConfig
            ElseIf optCargo.Value = True Then
                lVal = SlotConfig.eCargoBay
            ElseIf optCrewQuarters.Value = True Then
                lVal = SlotConfig.eCrewQuarters
            ElseIf optEngines.Value = True Then
                lVal = SlotConfig.eEngines
                'ElseIf optFuelBay.Value = True Then
                '    lVal = SlotConfig.eFuelBay
            ElseIf optHangar.Value = True Then
                lVal = SlotConfig.eHangar
            ElseIf optRadarComm.Value = True Then
                lVal = SlotConfig.eRadar
            ElseIf optShield.Value = True Then
                lVal = SlotConfig.eShields
            ElseIf optWeapons.Value = True Then
                lVal = SlotConfig.eWeapons
                lGrp = CInt(Val(txtGroupNum.Caption))
                Dim lMaxWpns As Int32 = HullTech.MaxWpnSlots(lscType.SelectedItemID, lscSubType.SelectedItemID) ' CInt(Val(txtHullSize.Caption)))
                If lGrp < 1 OrElse lGrp > lMaxWpns Then
                    MyBase.moUILib.AddNotification("Group Number for weapons must be a number from 1 to " & lMaxWpns & ".", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If
            ElseIf optHangarDoor.Value = True Then
                lVal = SlotConfig.eHangarDoor
                lGrp = CInt(Val(txtHgrDrNum.Caption))
                If lGrp < 1 Then
                    MyBase.moUILib.AddNotification("Group Number for hangar doors must be greater than 0.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If
            End If
        Else
            lVal = SlotConfig.eArmorConfig
        End If
        If lVal = SlotConfig.eArmorConfig Then lGrp = 0

        'If NewTutorialManager.TutorialOn = True Then
        '	If goUILib.CommandAllowed(True, "frmHullBuilder.hulMain") = False Then Return
        '	If sParms Is Nothing = False Then
        '		For X As Int32 = 0 To sParms.GetUpperBound(0)
        '			Select Case X
        '				Case 0 'slot type
        '					If CInt(sParms(X)) <> lVal Then Return
        '				Case 1 'groupnum
        '					If CInt(sParms(X)) <> -1 AndAlso CInt(sParms(X)) <> lGrp Then Return
        '				Case 2 'totalslots
        '					Dim lReq As Int32 = CInt(sParms(X))
        '					Dim lCur As Int32 = 0
        '					If CInt(sParms(X - 1)) = -1 Then
        '						lCur = hulMain.GetSlotConfigCnt(SlotType.eNoChange, lVal)
        '					Else : lCur = hulMain.GetSlotConfigCntWithGroup(SlotType.eNoChange, lVal, lGrp)
        '					End If
        '					lCur += 1
        '					If lCur > lReq Then Return
        '				Case 3 'side
        '					Dim yType As SlotType = SlotType.eNoChange
        '					Dim lConfig As SlotConfig = SlotConfig.eArmorConfig
        '					Dim lGrpNum As Int32 = 1
        '					hulMain.GetHullSlotValues(lIndexX, lIndexY, yType, lConfig, lGrpNum)
        '					If yType <> CInt(sParms(X)) Then Return
        '				Case 4 'allowerrors
        '					If sParms(X) = "0" Then
        '						If hulMain.WillCauseErrors(lIndexX, lIndexY, lVal, lGrp, CInt(Val(txtHullSize.Caption)), mbShip, mbSpaceStation) = True Then Return
        '					End If
        '			End Select
        '		Next X
        '	End If
        'End If

        'if we are holding shift down, we are deleting a group, so only if we set to armor
        If frmMain.mbShiftKeyDown = True AndAlso lVal = SlotConfig.eArmorConfig Then
            hulMain.ClearConfigGroup(lIndexX, lIndexY)
        Else
            hulMain.SetHullSlot(lIndexX, lIndexY, SlotType.eNoChange, lVal, lGrp)
        End If

        If NewTutorialManager.TutorialOn = True Then
            Dim lTotalSlots As Int32 = 0
            If lVal = SlotConfig.eWeapons Then
                lTotalSlots = hulMain.GetSlotConfigCntWithGroup(SlotType.eNoChange, SlotConfig.eWeapons, lGrp)
            Else : lTotalSlots = hulMain.GetSlotConfigCnt(SlotType.eNoChange, lVal)
            End If
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderSlotConfigChange, lVal, lGrp, lTotalSlots, "")
        End If

        'txtHullAllotment.Caption = hulMain.GetHullSummary(CInt(Val(txtHullSize.Caption)))
        txtHullAllotment.Caption = hulMain.GetHullSummary(CInt(tpHullSize.PropertyValue))
        RefreshErrorList()

        If lVal <> SlotConfig.eArmorConfig Then hulMain.HighLightAcceptableSlots(lVal, lGrp)
    End Sub

    Private Sub tpHullSize_ValueChange(ByVal sPropName As String, ByRef ctl As ctlTechProp) 'Handles txtHullSize.TextChanged
        If mbLoading = True Then Return

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderHullSizeChange, CInt(tpHullSize.PropertyValue), -1, -1, "")
        End If

        'ok... got a new value so.... 
        'txtHullAllotment.Caption = hulMain.GetHullSummary(CInt(Val(txtHullSize.Caption)))
        txtHullAllotment.Caption = hulMain.GetHullSummary(CInt(tpHullSize.PropertyValue))

        tpStructHP.MaxValue = tpHullSize.PropertyValue * 5

        mlCrewReqPerc = CInt(tpHullSize.PropertyValue / 100)
        If lscType.SelectedItemID = 2 Then mlCrewReqPerc = 0
        lblCrewReqs.Caption = "Suggested Crew: " & mlCrewReqPerc


        RefreshErrorList()
    End Sub
 
    Private Sub RefreshErrorList()

        Dim oSB As New System.Text.StringBuilder()

        lblSpecialTraitText.Caption = "None"
        Dim oModelDef As ModelDef = goModelDefs.GetModelDef(lscFrame.SelectedItemID)
        If oModelDef Is Nothing = False Then
            Dim sTemp As String = ModelDef.GetSpecialTraitText(oModelDef.lSpecialTraitID)
            If sTemp Is Nothing = False AndAlso sTemp <> "" Then
                lblSpecialTraitText.Caption = sTemp
            End If
        End If

        If lscFrame.SelectedItemID = 148 Then
            oSB.AppendLine("NOTE - This hull must be placed over a planet ring")
        End If

        If lscType.SelectedItemID = 2 Then
            oSB.AppendLine("NOTE - This hull cannot accept engines with a speed, maneuver or thrust that are greater than 0")
            oSB.AppendLine("NOTE - This hull receives a +50 Point Defense Accuracy Bonus")
        End If

        Dim lMaxGuns As Int32 = HullTech.MaxWpnSlots(lscType.SelectedItemID, lscSubType.SelectedItemID) ' CInt(Val(txtHullSize.Caption)))
        If hulMain.GetMaxWeaponGroupNum() > lMaxGuns Then
            oSB.AppendLine("ERROR - Max Weapon Group Number Exceeds Max Weapon Groups (" & lMaxGuns & ")")
        Else : oSB.AppendLine("NOTE - Max Weapon Groups is " & lMaxGuns)
        End If

        If (myHullRequirements And UIHullSlots.eyHullRequirements.RequiresFiftyThousandStructHP) <> 0 Then
            If tpStructHP.PropertyValue < 50000 Then
                oSB.AppendLine("ERROR - Hull Design must have at least 50000 hitpoints. Structural hitpoints determine how many modules can be installed.")
            End If
        End If

        oSB.Append(hulMain.HasErrorList(CInt(tpHullSize.PropertyValue), mlCrewReqPerc, myHullRequirements))

        If mlMineralID < 1 Then
            oSB.AppendLine("ERROR - Structural Material must be selected")
        End If

        Dim lDensity As Int32 = 0
        Dim bResMin As Boolean = False
        If mlMineralID > 0 Then
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = mlMineralID AndAlso goMinerals(X) Is Nothing = False Then
                    bResMin = goMinerals(X).bDiscovered
                    lDensity = goMinerals(X).MinPropValueScore(goMinerals(X).GetPropertyIndex(1))
                    Exit For
                End If
            Next X
        End If
        If bResMin = False Then
            oSB.AppendLine("ERROR - Structural Material must be Researched!")
            oSB.AppendLine("  The material is unknown to our people.")
            oSB.AppendLine("  Research the material or select another material.")
        End If
        'If Me.cboStructMat.ListIndex = -1 Then
        '    oSB.AppendLine("ERROR - Structural Material must be selected")
        'ElseIf Me.cboStructMat.ListIndex > -1 Then
        '    Dim lTmp As Int32 = cboStructMat.ItemData(cboStructMat.ListIndex)
        '    If lTmp < 1 Then
        '        oSB.AppendLine("ERROR - Structural Material must be selected")
        '    End If
        'End If

        If tpStructHP.PropertyValue < 1 Then
            oSB.AppendLine("ERROR - Structural Hitpoints must be greater than 0")
        End If

        Dim lTotal As Int32 = hulMain.TotalSlots
        Dim lVal As Int32 = CInt(tpHullSize.PropertyValue)

        If lVal > mlMaxHull OrElse lVal < mlMinHull Then
            oSB.AppendLine("ERROR - Hull Size Must be between " & mlMinHull & " and " & mlMaxHull)
        End If

        If hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eWeapons) > 0 Then
            If hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eRadar) = 0 Then
                oSB.AppendLine("ERROR - Weapon slots allocated but no radar slots")
            End If
        End If

        If tpStructHP.PropertyValue > tpHullSize.PropertyValue Then
            oSB.AppendLine("WARNING - Structural hitpoints exceeds the hull size!")
            oSB.AppendLine("  This can result is a Poor Design and excessive research time.")
        End If

        If lscType.SelectedItemID = 2 AndAlso lscSubType.SelectedItemID = 5 Then
            If hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eHangar) = 0 OrElse _
              hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eHangarDoor) = 0 Then
                oSB.AppendLine("WARNING - Production-Based facility without Hangar Bay or Hangar Doors cannot produce units!")
            End If
        End If

        Dim lPower As Int32 = hulMain.GetApproxPowerRequired(oModelDef.FrameTypeID, lVal, lDensity)
        oSB.AppendLine("NOTE - Estimated Power Usage: " & lPower.ToString("#,##0"))

        'If lscType.SelectedItemID = 2 AndAlso lscSubType.SelectedItemID = 1 Then
        '    Dim lHullSlots As Int32 = hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eHangar)
        '    Dim lHullDoorSlots As Int32 = hulMain.GetBiggestHangarDoorsize()
        '    Dim lCargoBaySlots As Int32 = hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eCargoBay)
        '    If lTotal <> 0 Then
        '        Dim fHullPerSlot As Single = CSng(lVal / lTotal)
        '        fHullPerSlot *= 100.0F
        '        fHullPerSlot = CSng(Math.Floor(fHullPerSlot))
        '        fHullPerSlot /= 100.0F

        '        If lHullSlots * fHullPerSlot < 500 OrElse lHullDoorSlots * fHullPerSlot < 500 Then
        '            oSB.AppendLine("ERROR - Mining facilities must contain at least 500 hull size for a hangar door and hangar")
        '        End If
        '        If lCargoBaySlots * fHullPerSlot < 1000 Then
        '            oSB.AppendLine("ERROR - Mining facilities must have at least 1000 hull configured as cargo bay")
        '        End If
        '    Else
        '        oSB.AppendLine("ERROR - Mining facilities must contain at least 500 hull size for a hangar door and hangar")
        '    End If
        'End If


        txtErrors.Caption = oSB.ToString

        If lTotal = 0 Then
            optHullSlotSize.Caption = "= 0 Hull Per Slot"
        Else
            Dim fHullPerSlot As Single = CSng(lVal / lTotal)
            fHullPerSlot *= 100.0F
            fHullPerSlot = CSng(Math.Floor(fHullPerSlot))
            fHullPerSlot /= 100.0F
            optHullSlotSize.Caption = "= " & fHullPerSlot.ToString("#,###.0#") & " Hull Per Slot"
            Dim rcTemp As Rectangle = optHullSlotSize.GetTextDimensions()
            optHullSlotSize.Width = rcTemp.Width + 20
        End If
    End Sub

#Region "Option Button Clicks"
    Private Sub optArmor_Click() Handles optArmor.Click
        hulMain.ClearSelectingSlotConfig()
        optArmor.Value = True
        optCargo.Value = False
        optCrewQuarters.Value = False
        optEngines.Value = False
        'optFuelBay.Value = False
        optHangar.Value = False
        optRadarComm.Value = False
        optShield.Value = False
        optWeapons.Value = False
        optHangarDoor.Value = False
        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderSlotPaletteSelected, SlotConfig.eArmorConfig, -1, -1, "")
        End If

        hulMain.HighLightAcceptableSlots(SlotConfig.eArmorConfig, -1)
    End Sub

    Private Sub optHangarDoor_Click() Handles optHangarDoor.Click
        hulMain.FilterEdgeSlots(True)
        optArmor.Value = False
        optCargo.Value = False
        optCrewQuarters.Value = False
        optEngines.Value = False
        'optFuelBay.Value = False
        optHangar.Value = False
        optRadarComm.Value = False
        optShield.Value = False
        optWeapons.Value = False
        optHangarDoor.Value = True
        If MyBase.moUILib.FocusedControl Is Nothing = False Then
            MyBase.moUILib.FocusedControl.HasFocus = False
        End If
        MyBase.moUILib.FocusedControl = txtHgrDrNum
        txtHgrDrNum.HasFocus = True
        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderSlotPaletteSelected, SlotConfig.eHangarDoor, -1, -1, "")
        End If

        Dim lVal As Int32 = CInt(Val(txtHgrDrNum.Caption))
        hulMain.HighLightAcceptableSlots(SlotConfig.eHangarDoor, lVal)
    End Sub

    Private Sub optCargo_Click() Handles optCargo.Click
        hulMain.ClearSelectingSlotConfig()
        optArmor.Value = False
        optCargo.Value = True
        optCrewQuarters.Value = False
        optEngines.Value = False
        'optFuelBay.Value = False
        optHangar.Value = False
        optRadarComm.Value = False
        optShield.Value = False
        optWeapons.Value = False
        optHangarDoor.Value = False
        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderSlotPaletteSelected, SlotConfig.eCargoBay, -1, -1, "")
        End If

        hulMain.HighLightAcceptableSlots(SlotConfig.eCargoBay, -1)
    End Sub

    Private Sub optCrewQuarters_Click() Handles optCrewQuarters.Click
        hulMain.ClearSelectingSlotConfig()
        optArmor.Value = False
        optCargo.Value = False
        optCrewQuarters.Value = True
        optEngines.Value = False
        'optFuelBay.Value = False
        optHangar.Value = False
        optRadarComm.Value = False
        optShield.Value = False
        optWeapons.Value = False
        optHangarDoor.Value = False
        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderSlotPaletteSelected, SlotConfig.eArmorConfig, -1, -1, "")
        End If

        hulMain.HighLightAcceptableSlots(SlotConfig.eCrewQuarters, -1)
    End Sub

    Private Sub optEngines_Click() Handles optEngines.Click
        hulMain.ClearSelectingSlotConfig()
        optArmor.Value = False
        optCargo.Value = False
        optCrewQuarters.Value = False
        optEngines.Value = True
        'optFuelBay.Value = False
        optHangar.Value = False
        optRadarComm.Value = False
        optShield.Value = False
        optWeapons.Value = False
        optHangarDoor.Value = False
        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderSlotPaletteSelected, SlotConfig.eEngines, -1, -1, "")
        End If

        hulMain.HighLightAcceptableSlots(SlotConfig.eEngines, -1)
    End Sub

    'Private Sub optFuelBay_Click() Handles optFuelBay.Click
    '    hulMain.ClearSelectingSlotConfig()
    '    optArmor.Value = False
    '    optCargo.Value = False
    '    optCrewQuarters.Value = False
    '    optEngines.Value = False
    '    optFuelBay.Value = True
    '    optHangar.Value = False
    '    optRadarComm.Value = False
    '    optShield.Value = False
    '    optWeapons.Value = False
    '    optHangarDoor.Value = False
    'End Sub

    Private Sub optHangar_Click() Handles optHangar.Click
        hulMain.ClearSelectingSlotConfig()
        optArmor.Value = False
        optCargo.Value = False
        optCrewQuarters.Value = False
        optEngines.Value = False
        'optFuelBay.Value = False
        optHangar.Value = True
        optRadarComm.Value = False
        optShield.Value = False
        optWeapons.Value = False
        optHangarDoor.Value = False
        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderSlotPaletteSelected, SlotConfig.eHangar, -1, -1, "")
        End If

        hulMain.HighLightAcceptableSlots(SlotConfig.eHangar, -1)
    End Sub

    Private Sub optRadarComm_Click() Handles optRadarComm.Click
        hulMain.ClearSelectingSlotConfig()
        optArmor.Value = False
        optCargo.Value = False
        optCrewQuarters.Value = False
        optEngines.Value = False
        'optFuelBay.Value = False
        optHangar.Value = False
        optRadarComm.Value = True
        optShield.Value = False
        optWeapons.Value = False
        optHangarDoor.Value = False
        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderSlotPaletteSelected, SlotConfig.eRadar, -1, -1, "")
        End If

        hulMain.HighLightAcceptableSlots(SlotConfig.eRadar, -1)
    End Sub

    Private Sub optShield_Click() Handles optShield.Click
        hulMain.ClearSelectingSlotConfig()
        optArmor.Value = False
        optCargo.Value = False
        optCrewQuarters.Value = False
        optEngines.Value = False
        'optFuelBay.Value = False
        optHangar.Value = False
        optRadarComm.Value = False
        optShield.Value = True
        optWeapons.Value = False
        optHangarDoor.Value = False
        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderSlotPaletteSelected, SlotConfig.eShields, -1, -1, "")
        End If

        hulMain.HighLightAcceptableSlots(SlotConfig.eShields, -1)
    End Sub

    Private Sub optWeapons_Click() Handles optWeapons.Click
        hulMain.ClearSelectingSlotConfig()
        optArmor.Value = False
        optCargo.Value = False
        optCrewQuarters.Value = False
        optEngines.Value = False
        'optFuelBay.Value = False
        optHangar.Value = False
        optRadarComm.Value = False
        optShield.Value = False
        optWeapons.Value = True
        optHangarDoor.Value = False
        If MyBase.moUILib.FocusedControl Is Nothing = False Then
            MyBase.moUILib.FocusedControl.HasFocus = False
        End If
        MyBase.moUILib.FocusedControl = txtGroupNum
        txtGroupNum.HasFocus = True

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderSlotPaletteSelected, SlotConfig.eWeapons, -1, -1, "")
        End If

        Dim lVal As Int32 = CInt(Val(txtGroupNum.Caption))
        hulMain.HighLightAcceptableSlots(SlotConfig.eWeapons, lVal)
    End Sub
#End Region

#Region "Label Scrolls"
    Private Sub lscType_ItemChanged(ByVal lID As Integer, ByVal sDisplay As String) Handles lscType.ItemChanged
        If mbIgnoreLabelScrollEvents = True Then Return

        lscSubType.Clear()

        Dim bHasBC As Boolean = False
        Dim bHasBS As Boolean = False
        Dim bHasEscort As Boolean = False
        Dim bHasCorvette As Boolean = False
        Dim bHasFrig As Boolean = False
        Dim bHasDest As Boolean = False
        Dim bHasCruiser As Boolean = False
        Dim lMaxHull As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eEnginePowerThrustLimit)

        Dim bHasNavalUnit As Boolean = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eNavalAbility) > 0
        Dim bHasNavalFac As Boolean = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eNavalAbility) > 1

        Dim bHasNavalBS As Boolean = False
        Dim bHasNavalCarrier As Boolean = False
        Dim bHasNavalCruiser As Boolean = False
        Dim bHasNavalDestroyer As Boolean = False
        Dim bHasNavalFrigate As Boolean = False
        Dim bHasNavalSub As Boolean = False
        'lMaxHull = 10000000

        For X As Int32 = 0 To goModelDefs.MaxModelDefUB
            Dim oModelDef As ModelDef = goModelDefs.GetModelDef(X)
            If oModelDef Is Nothing = False AndAlso oModelDef.MinHull < lMaxHull Then
                Select Case oModelDef.TypeID
                    Case 0  'capital
                        If oModelDef.SubTypeID = 2 Then bHasBC = True
                        If oModelDef.SubTypeID = 0 Then bHasBS = True
                    Case 1
                        Select Case oModelDef.SubTypeID
                            Case 4 : bHasEscort = True
                            Case 0 : bHasCorvette = True
                            Case 3 : bHasFrig = True
                            Case 2 : bHasDest = True
                            Case 1 : bHasCruiser = True
                        End Select
                    Case 8  'naval
                        Select Case oModelDef.SubTypeID
                            Case 0 : bHasNavalBS = True
                            Case 1 : bHasNavalCarrier = True
                            Case 2 : bHasNavalCruiser = True
                            Case 3 : bHasNavalDestroyer = True
                            Case 4 : bHasNavalFrigate = True
                            Case 5 : bHasNavalSub = True
                        End Select
                End Select
            End If
        Next X

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderScrollTypeSelected, lID, -1, -1, "")
        End If

        Select Case lID
            Case 0  'capital
                If bHasBC = True Then lscSubType.AddItem(2, "Battle Cruiser")
                If bHasBS = True Then lscSubType.AddItem(0, "Battleship")
                'lscSubType.AddItem(1, "Carrier")
            Case 1  'escort
                If bHasEscort = True Then lscSubType.AddItem(4, "Escort")
                If bHasCorvette = True Then lscSubType.AddItem(0, "Corvette")
                If bHasFrig = True Then lscSubType.AddItem(3, "Frigate")
                If bHasDest = True Then lscSubType.AddItem(2, "Destroyer")
                If bHasCruiser = True Then lscSubType.AddItem(1, "Cruiser")
            Case 2  'facility
                lscSubType.AddItem(0, "Command Center")
                lscSubType.AddItem(1, "Mining")
                lscSubType.AddItem(2, "Other")
                lscSubType.AddItem(3, "Personnel")
                lscSubType.AddItem(4, "Power Generator")
                lscSubType.AddItem(5, "Production")
                lscSubType.AddItem(6, "Refining")
                lscSubType.AddItem(7, "Research")
                lscSubType.AddItem(8, "Residence")
                lscSubType.AddItem(9, "Space Station")
                lscSubType.AddItem(10, "Tradepost")
            Case 3  'fighter
                lscSubType.AddItem(0, "Light")
                lscSubType.AddItem(1, "Medium")
                lscSubType.AddItem(2, "Heavy")
            Case 4  'small vehicle
                lscSubType.AddItem(0, "Quad")
            Case 5  'tank
                'lscSubType.AddItem(0, "Half-Track")
                lscSubType.AddItem(1, "Hover")
                'lscSubType.AddItem(2, "Track")
            Case 6  'transport
                lscSubType.AddItem(0, "Cargo Transport")
                'lscSubType.AddItem(1, "Supply Transport")
                'lscSubType.AddItem(2, "Unit Transport")
            Case 7  'utility vehicle
                lscSubType.AddItem(0, "Land Engineer")
                lscSubType.AddItem(1, "Space Engineer")
                'lscSubType.AddItem(2, "Cargo Truck")
                'lscSubType.AddItem(3, "Mining Truck")
            Case 8  'naval
                If bHasNavalBS = True Then lscSubType.AddItem(0, "Battleship")
                If bHasNavalCarrier = True Then lscSubType.AddItem(1, "Carrier")
                If bHasNavalCruiser = True Then lscSubType.AddItem(2, "Cruiser")
                If bHasNavalDestroyer = True Then lscSubType.AddItem(3, "Destroyer")
                If bHasNavalFrigate = True Then lscSubType.AddItem(4, "Frigate")
                If bHasNavalSub = True Then lscSubType.AddItem(5, "Submarine")
                lscSubType.AddItem(6, "Utility")
        End Select

        lscSubType.ResetCurrentIndex()
    End Sub

    Private Sub lscSubType_ItemChanged(ByVal lID As Integer, ByVal sDisplay As String) Handles lscSubType.ItemChanged
        If mbIgnoreLabelScrollEvents = True Then Return
        lscFrame.Clear()
        goModelDefs.AddLabelScrollVals(lscFrame, CByte(lscType.SelectedItemID), CByte(lID))
        lscFrame.ResetCurrentIndex()

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderScrollSubTypeSelected, lID, -1, -1, "")
        End If

        If lscFrame.ListCount = 0 Then
            hulMain.SetByModelID(0)
        End If

        'Ok, determine the hulltypeid from the type/subtype
        ' then, fill our statics list accordingly...
        Dim yHullTypeID As Byte = HullTech.GetHullTypeID(lscType.SelectedItemID, lscSubType.SelectedItemID)
        FillLookupList(yHullTypeID)

    End Sub

    Private Sub lscFrame_ItemChanged(ByVal lID As Integer, ByVal sDisplay As String) Handles lscFrame.ItemChanged
        If mbIgnoreLabelScrollEvents = True Then Return

        mlInitialZ = Int32.MinValue

        'clear our config data
        hulMain.ClearAllConfigs()

        optArmor_Click()

        'Hull slots should be set now...
        hulMain.SetByModelID(lID)

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderScrollFrameSelected, lID, -1, -1, "")
        End If

        'Now, set our Distributions
        Dim lTotal As Int32 = hulMain.TotalSlots
        If lTotal <> 0 Then
            lblForeDist.Caption = "Forward: " & ((hulMain.FrontSlots / lTotal) * 100).ToString("#0.#") & "%"
            lblLeftDist.Caption = "Left: " & ((hulMain.LeftSlots / lTotal) * 100).ToString("#0.#") & "%"
            lblRearDist.Caption = "Rear: " & ((hulMain.RearSlots / lTotal) * 100).ToString("#0.#") & "%"
            lblRightDist.Caption = "Right: " & ((hulMain.RightSlots / lTotal) * 100).ToString("#0.#") & "%"
            lblAllArcDist.Caption = "All-Arc: " & ((hulMain.AllArcSlots / lTotal) * 100).ToString("#0.#") & "%"
        End If

        Dim oModelDef As ModelDef = goModelDefs.GetModelDef(lID)

        If oModelDef Is Nothing = False Then
            With oModelDef

                If .bMTTexCapable = True Then
                    Dim oMesh As BaseMesh = goResMgr.GetMesh(lID)
                    If oMesh Is Nothing = False AndAlso oMesh.oSharedTex Is Nothing = False Then
                        lscDiffuse.Enabled = True
                        lscNormal.Enabled = True
                        lscIllum.Enabled = True

                        lscDiffuse.Clear()
                        lscNormal.Clear()
                        lscIllum.Clear()
                        For X As Int32 = 0 To oMesh.oSharedTex.GetUpperBound(0)
                            lscDiffuse.AddItem(X, (X + 1).ToString)
                        Next X
                        Dim bView As Boolean = mbModelView
                        lscDiffuse.SelectByID(oMesh.lDefaultSharedTexIdx)
                        If mbModelView <> bView Then btnSwitchView_Click(btnSwitchView.ControlName)
                        mbIgnoreLabelScrollEvents = True
                        lscNormal.SelectByID(0)
                        lscIllum.SelectByID(0)
                        mbIgnoreLabelScrollEvents = False
                    End If
                Else
                    lscDiffuse.Enabled = False
                    lscNormal.Enabled = False
                    lscIllum.Enabled = False
                End If

                lblHullSizeRange.Caption = "Hull Size: " & .MinHull.ToString("#,##0") & "-" & .MaxHull.ToString("#,##0")
                'lblCrewReqs.Caption = "Crew Requirement: " & .CrewReqPerc.ToString & "%"

                Me.txtGroupNum.ToolTipText = "Specific weapon group number to be configured." & vbCrLf & "Each weapon group indicates an individual" & vbCrLf & "weapon on the hull. Maximum value is " & HullTech.MaxWpnSlots(.TypeID, .SubTypeID) & "."

                'mlCrewReqPerc = CInt(Val(.CrewReqPerc))

                '.TypeID 
                Select Case .TypeID
                    Case 0, 1, 3           'capital, escort ship, fighter
                        'mbSpaceStation = False
                        'mbShip = True
                        myHullRequirements = UIHullSlots.eyHullRequirements.EngineIsToBeInRear Or UIHullSlots.eyHullRequirements.RequiresAnEngine
                    Case 2              'facility
                        'mbSpaceStation = (.FrameTypeID = 4) 
                        'mbShip = False
                        If .FrameTypeID = 4 Then
                            If .ModelID = 138 Then
                                'space defense
                                myHullRequirements = UIHullSlots.eyHullRequirements.RequiresAnEngine
                            Else
                                myHullRequirements = UIHullSlots.eyHullRequirements.RequiresAnEngine Or UIHullSlots.eyHullRequirements.RequiresBayDoor Or UIHullSlots.eyHullRequirements.RequiresFiftyThousandStructHP
                            End If
                        Else
                            myHullRequirements = 0
                        End If
                    Case Else
                        'small vehicle, tanks, transports, utility
                        'mbShip = (.FrameTypeID <> 2)
                        'mbSpaceStation = False
                        If .FrameTypeID <> 2 Then
                            myHullRequirements = UIHullSlots.eyHullRequirements.RequiresAnEngine Or UIHullSlots.eyHullRequirements.EngineIsToBeInRear
                        Else
                            myHullRequirements = UIHullSlots.eyHullRequirements.RequiresAnEngine
                        End If
                End Select

                mbIgnoreCheckEvents = True
                chkSpace.Value = False : chkGround.Value = False : chkAtmospheric.Value = False : chkNaval.Value = False
                chkSpace.Enabled = True : chkGround.Enabled = True : chkAtmospheric.Enabled = True : chkNaval.Enabled = True
                Select Case .FrameTypeID
                    Case 0
                        'space and atmosphere
                        chkSpace.Value = True
                        chkAtmospheric.Value = True
                        chkGround.Enabled = False
                        chkNaval.Enabled = False
                    Case 1
                        'atmosphere only
                        chkAtmospheric.Value = True
                        chkGround.Enabled = False
                        chkNaval.Enabled = False
                    Case 2
                        'ground based
                        chkGround.Value = True
                        chkAtmospheric.Enabled = False
                        chkSpace.Enabled = False
                    Case 3
                        'naval based
                        chkNaval.Value = True
                        chkGround.Enabled = False
                        chkAtmospheric.Enabled = False
                        chkSpace.Enabled = False
                    Case 4
                        'space only
                        chkSpace.Value = True
                        chkGround.Enabled = False
                        chkNaval.Enabled = False
                        'If .TypeID = 2 Then
                        'space-based facilities cannot enter atmospheres
                        chkAtmospheric.Enabled = False
                        'End If
                End Select

                chkNaval.Enabled = False
                If goCurrentPlayer Is Nothing = False AndAlso (goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eNavalAbility) And 2) <> 0 Then
                    If .TypeID = 2 AndAlso .FrameTypeID <> 4 Then
                        chkNaval.Enabled = True
                    End If
                End If


                mbGroundLocked = chkGround.Value : mbSpaceLocked = chkSpace.Value : mbAtmosLocked = chkAtmospheric.Value : mbNavalLocked = chkNaval.Value
                mbIgnoreCheckEvents = False

                tpHullSize.MinValue = .MinHull
                tpHullSize.MaxValue = .MaxHull
                tpHullSize.PropertyValue = tpHullSize.MinValue
                mlMinHull = .MinHull
                mlMaxHull = .MaxHull
            End With
        Else
            lblHullSizeRange.Caption = ""
            lblCrewReqs.Caption = ""
            mlCrewReqPerc = 0
            mlMinHull = 0
            mlMaxHull = 0
        End If

        RefreshErrorList()
    End Sub

    Private Sub FillTypesList()

        Dim lThrustLimit As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eEnginePowerThrustLimit)
        Dim bHasCapitalShip As Boolean = lThrustLimit > 110000

        With lscType
            If bHasCapitalShip = True Then .AddItem(0, "Capital Ship")
            .AddItem(1, "Escort Ship")
            .AddItem(2, "Facility")
            .AddItem(3, "Fighter")
            Dim bHasNavalUnit As Boolean = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eNavalAbility) > 0
            Dim bHasNavalFac As Boolean = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eNavalAbility) > 1
            If bHasNavalUnit = True OrElse bHasNavalFac = True Then
                .AddItem(8, "Naval")
            End If
            .AddItem(4, "Small Vehicle")
            .AddItem(5, "Tank")
            If lThrustLimit > 30000 Then .AddItem(6, "Transport")
            .AddItem(7, "Utility Vehicle")
        End With
        lscType.ResetCurrentIndex()
    End Sub
#End Region

    Private Sub txtGroupNum_TextChanged() Handles txtGroupNum.TextChanged
        If mbLoading = True Then Return
        mbLoading = True
        If txtGroupNum.Caption <> "" Then
            Dim lValue As Int32 = CInt(Val(txtGroupNum.Caption).ToString)
            Dim lMaxGuns As Int32 = HullTech.MaxWpnSlots(lscType.SelectedItemID, lscSubType.SelectedItemID) ' CInt(Val(txtHullSize.Caption)))
            If lValue > lMaxGuns Then lValue = lMaxGuns

            If NewTutorialManager.TutorialOn = True Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderWeaponGroupChange, lValue, -1, -1, "")
            End If

            txtGroupNum.Caption = lValue.ToString
        End If
        mbLoading = False

        If optWeapons.Value = True Then
            If CInt(Val(txtGroupNum.Caption)) > 0 Then hulMain.HighLightAcceptableSlots(SlotConfig.eWeapons, CInt(Val(txtGroupNum.Caption)))
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
        CloseMe()
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
        If mlEntityIndex = -1 Then
            MyBase.moUILib.AddNotification("A Research Facility must be selected!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        'If we're in the tutorial do a validate versus the data here and what is expected
        If NewTutorialManager.TutorialOn = True Then
            If ValidateTutorialSettings() = False Then Return
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

        If Me.txtHullName.Caption.Trim.Length = 0 Then
            MyBase.moUILib.AddNotification("You must specify a name for this Hull.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        'txtErrors.Caption = hulMain.HasErrorList(CInt(Val(txtHullSize.Caption)), mlCrewReqPerc, mbShip, mbSpaceStation)
        RefreshErrorList()
        If txtErrors.Caption.ToUpper.Contains("ERROR") = True Then
            MyBase.moUILib.AddNotification("Please fix the errors listed in Design Flaws.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        Dim lHullSize As Int32 = CInt(tpHullSize.PropertyValue)
        If lHullSize > mlMaxHull OrElse lHullSize < mlMinHull Then
            MyBase.moUILib.AddNotification("Hull Size must be between " & mlMinHull & " and " & mlMaxHull, Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        If lHullSize = 0 Then
            MyBase.moUILib.AddNotification("Please select a hull frame first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        Dim lStructHP As Int32 = CInt(tpStructHP.PropertyValue)

        Dim ySlotData() As Byte = hulMain.GetMsgRdySlotList()
        Dim yMsg(52 + ySlotData.Length) As Byte
        Dim lPos As Int32 = 0

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
        System.BitConverter.GetBytes(ObjectType.eHullTech).CopyTo(yMsg, lPos) : lPos += 2

        If moTech Is Nothing = False Then
            System.BitConverter.GetBytes(moTech.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
        Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
        End If

        oResearchGuid.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        System.Text.ASCIIEncoding.ASCII.GetBytes(txtHullName.Caption).CopyTo(yMsg, lPos) : lPos += 20
        System.BitConverter.GetBytes(lHullSize).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(mlMineralID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(lStructHP).CopyTo(yMsg, lPos) : lPos += 4

        yMsg(lPos) = CByte(lscType.SelectedItemID) : lPos += 1
        yMsg(lPos) = CByte(lscSubType.SelectedItemID) : lPos += 1
        Dim lFinalModelID As Int32 = ConstructModelID()
        System.BitConverter.GetBytes(lFinalModelID).CopyTo(yMsg, lPos) : lPos += 4

        'Now, compile our desired chassis type
        Dim yChassisType As Byte = 0
        If chkGround.Value = True Then yChassisType += ChassisType.eGroundBased
        If chkAtmospheric.Value = True Then yChassisType += ChassisType.eAtmospheric
        If chkSpace.Value = True Then yChassisType += ChassisType.eSpaceBased
        If chkNaval.Value = True Then yChassisType = yChassisType Or ChassisType.eNavalBased
        yMsg(lPos) = yChassisType : lPos += 1

        ySlotData.CopyTo(yMsg, lPos)

        MyBase.moUILib.SendMsgToPrimary(yMsg)

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eSubmitDesignClick, ObjectType.eHullTech, -1, -1, "")
        End If

        MyBase.moUILib.AddNotification("Hull Design Submitted", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

        CloseMe()
    End Sub

    Private Sub FillLookupList(ByVal yHullTypeID As Byte)
        Dim lSorted() As Int32

        'Now, for the Hull Usage Lookup
        Dim iTypes() As Int16 = {ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech}
        lstHullLookup.Clear()

        Dim lCrewMult As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBetterHullToResidence)

        'lstHullLookup.sHeaderRow = "Component".PadRight(21, " "c) & "Hull Usage"
        For X As Int32 = 0 To iTypes.GetUpperBound(0)
            lSorted = Nothing
            Dim lSortedUB As Int32 = -1

            For Y As Int32 = 0 To goCurrentPlayer.mlTechUB
                If goCurrentPlayer.moTechs(Y) Is Nothing = False AndAlso goCurrentPlayer.moTechs(Y).ObjTypeID = iTypes(X) Then
                    Dim lIdx As Int32 = -1
                    For Z As Int32 = 0 To lSortedUB
                        If goCurrentPlayer.moTechs(lSorted(Z)).GetComponentName > goCurrentPlayer.moTechs(Y).GetComponentName Then
                            lIdx = Z
                            Exit For
                        End If
                    Next Z
                    lSortedUB += 1
                    ReDim Preserve lSorted(lSortedUB)
                    If lIdx = -1 Then
                        lSorted(lSortedUB) = Y
                    Else
                        For Z As Int32 = lSortedUB To lIdx + 1 Step -1
                            lSorted(Z) = lSorted(Z - 1)
                        Next Z
                        lSorted(lIdx) = Y
                    End If
                End If
            Next Y

            lstHullLookup.AddItem(Base_Tech.GetComponentTypeName(iTypes(X)), True)
            lstHullLookup.ItemLocked(lstHullLookup.NewIndex) = True
            For Y As Int32 = 0 To lSortedUB
                Dim sTemp As String = goCurrentPlayer.moTechs(lSorted(Y)).GetComponentName.PadRight(22, " "c)
                Select Case iTypes(X)
                    Case ObjectType.eArmorTech
                        sTemp &= CType(goCurrentPlayer.moTechs(lSorted(Y)), ArmorTech).HullUsagePerPlate.ToString.PadLeft(11, " "c)
                    Case ObjectType.eEngineTech
                        With CType(goCurrentPlayer.moTechs(lSorted(Y)), EngineTech)
                            If .HullTypeID <> yHullTypeID AndAlso .HullTypeID <> 255 Then Continue For
                            If .oProductionCost Is Nothing = False Then
                                Dim lCrew As Int32 = .oProductionCost.ColonistCost + .oProductionCost.EnlistedCost + .oProductionCost.OfficerCost
                                lCrew *= lCrewMult
                                If lCrew <> 0 Then sTemp &= (.HullRequired.ToString & "(+" & lCrew.ToString & ")").PadLeft(11, " "c) Else sTemp &= .HullRequired.ToString.PadLeft(11, " "c)
                            Else
                                sTemp &= .HullRequired.ToString.PadLeft(11, " "c)
                            End If
                        End With
                    Case ObjectType.eRadarTech
                        With CType(goCurrentPlayer.moTechs(lSorted(Y)), RadarTech)
                            If .HullTypeID <> yHullTypeID AndAlso .HullTypeID <> 255 Then Continue For
                            If .oProductionCost Is Nothing = False Then
                                Dim lCrew As Int32 = .oProductionCost.ColonistCost + .oProductionCost.EnlistedCost + .oProductionCost.OfficerCost
                                lCrew *= lCrewMult
                                If lCrew <> 0 Then sTemp &= (.HullRequired.ToString & "(+" & lCrew.ToString & ")").PadLeft(11, " "c) Else sTemp &= .HullRequired.ToString.PadLeft(11, " "c)
                            Else
                                sTemp &= .HullRequired.ToString.PadLeft(11, " "c)
                            End If
                        End With
                    Case ObjectType.eShieldTech
                        With CType(goCurrentPlayer.moTechs(lSorted(Y)), ShieldTech)
                            If .HullTypeID <> yHullTypeID AndAlso .HullTypeID <> 255 Then Continue For
                            If .oProductionCost Is Nothing = False Then
                                Dim lCrew As Int32 = .oProductionCost.ColonistCost + .oProductionCost.EnlistedCost + .oProductionCost.OfficerCost
                                lCrew *= lCrewMult
                                If lCrew <> 0 Then sTemp &= (.HullRequired.ToString & "(+" & lCrew.ToString & ")").PadLeft(11, " "c) Else sTemp &= .HullRequired.ToString.PadLeft(11, " "c)
                            Else
                                sTemp &= .HullRequired.ToString.PadLeft(11, " "c)
                            End If
                        End With
                    Case ObjectType.eWeaponTech
                        With CType(goCurrentPlayer.moTechs(lSorted(Y)), WeaponTech)
                            If .HullTypeID <> yHullTypeID AndAlso .HullTypeID <> 255 Then Continue For
                            If .oProductionCost Is Nothing = False Then
                                Dim lCrew As Int32 = .oProductionCost.ColonistCost + .oProductionCost.EnlistedCost + .oProductionCost.OfficerCost
                                lCrew *= lCrewMult
                                If lCrew <> 0 Then sTemp &= (.HullRequired.ToString & "(+" & lCrew.ToString & ")").PadLeft(11, " "c) Else sTemp &= .HullRequired.ToString.PadLeft(11, " "c)
                            Else
                                sTemp &= .HullRequired.ToString.PadLeft(11, " "c)
                            End If
                        End With
                End Select

                lstHullLookup.AddItem(sTemp, False)
                If goCurrentPlayer.moTechs(lSorted(Y)).ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
                    lstHullLookup.ItemCustomColor(lstHullLookup.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                End If
            Next Y

            ReDim lSorted(-1)
            lSortedUB = -1
            For Y As Int32 = 0 To glPlayerTechKnowledgeUB
                If glPlayerTechKnowledgeIdx(Y) <> -1 AndAlso goPlayerTechKnowledge(Y).oTech Is Nothing = False AndAlso goPlayerTechKnowledge(Y).oTech.ObjTypeID = iTypes(X) Then
                    If goPlayerTechKnowledge(Y).yKnowledgeType = PlayerTechKnowledge.KnowledgeType.eSettingsLevel2 Then
                        Dim lIdx As Int32 = -1
                        For Z As Int32 = 0 To lSortedUB
                            If goPlayerTechKnowledge(lSorted(Z)).oTech.GetComponentName.ToUpper > goPlayerTechKnowledge(Y).oTech.GetComponentName.ToUpper Then
                                lIdx = Z
                                Exit For
                            End If
                        Next Z
                        lSortedUB += 1
                        ReDim Preserve lSorted(lSortedUB)
                        If lIdx = -1 Then
                            lSorted(lSortedUB) = Y
                        Else
                            For Z As Int32 = lSortedUB To lIdx + 1 Step -1
                                lSorted(Z) = lSorted(Z - 1)
                            Next Z
                            lSorted(lIdx) = Y
                        End If
                    End If
                End If
            Next Y

            Dim clrNonOwner As System.Drawing.Color = System.Drawing.Color.FromArgb(255, CInt(muSettings.InterfaceBorderColor.R * 0.75F), CInt(muSettings.InterfaceBorderColor.G * 0.75F), CInt(muSettings.InterfaceBorderColor.B * 0.75F))
            For Y As Int32 = 0 To lSortedUB
                Dim sTemp As String = goPlayerTechKnowledge(lSorted(Y)).oTech.GetComponentName.PadRight(22, " "c)
                Dim oTmpTech As Base_Tech = goPlayerTechKnowledge(lSorted(Y)).oTech
                Select Case iTypes(X)
                    Case ObjectType.eArmorTech
                        sTemp &= CType(oTmpTech, ArmorTech).HullUsagePerPlate.ToString.PadLeft(11, " "c)
                    Case ObjectType.eEngineTech
                        sTemp &= CType(oTmpTech, EngineTech).HullRequired.ToString.PadLeft(11, " "c)
                    Case ObjectType.eRadarTech
                        sTemp &= CType(oTmpTech, RadarTech).HullRequired.ToString.PadLeft(11, " "c)
                    Case ObjectType.eShieldTech
                        sTemp &= CType(oTmpTech, ShieldTech).HullRequired.ToString.PadLeft(11, " "c)
                    Case ObjectType.eWeaponTech
                        sTemp &= CType(oTmpTech, WeaponTech).HullRequired.ToString.PadLeft(11, " "c)
                End Select

                lstHullLookup.AddItem(sTemp, False)
                lstHullLookup.ItemCustomColor(lstHullLookup.NewIndex) = clrNonOwner
            Next Y

        Next X
    End Sub
 
    Private Sub CloseMe()
        'Dim oTmpWin As UIWindow = MyBase.moUILib.GetWindow("frmChat")
        'If oTmpWin Is Nothing = False Then oTmpWin.Visible = True
        'oTmpWin = Nothing

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.RemoveWindow("frmResCost")
        MyBase.moUILib.RemoveWindow("frmProdCost")
        MyBase.moUILib.RemoveWindow("frmMinDetail")
        ReturnToPreviousView()
        If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
        goFullScreenBackground = Nothing
    End Sub
 
    Private Sub btnSwitchView_Click(ByVal sName As String) Handles btnSwitchView.Click
        mbModelView = Not mbModelView
        hulMain.Visible = Not mbModelView
    End Sub

    Private Sub frmHullBuilder_OnNewFrame() Handles Me.OnNewFrame
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return
        If mbModelView = True Then

            'Ok, gotta render the model... get our model id
            Dim iModelID As Int16 = ConstructModelID() '(CShort(lscFrame.SelectedItemID), CShort(lscDiffuse.SelectedItemID), CShort(lscNormal.SelectedItemID), CShort(lscIllum.SelectedItemID))
            Dim oOriginal As Surface
            Dim oScene As Surface = Nothing
            Dim matView As Matrix
            Dim matProj As Matrix

            Dim lCX As Int32
            Dim lCY As Int32
            Dim lCZ As Int32
            Dim lCAtX As Int32
            Dim lCAtY As Int32
            Dim lCAtZ As Int32

            'Get our current camera location
            With goCamera
                lCX = .mlCameraX
                lCY = .mlCameraY
                lCZ = .mlCameraZ
                lCAtX = .mlCameraAtX
                lCAtY = .mlCameraAtY
                lCAtZ = .mlCameraAtZ

                .mlCameraAtX = 0
                .mlCameraAtY = 0
                .mlCameraAtZ = 0
                .mlCameraX = CInt(mfLastCX)
                .mlCameraY = mlLastCY
                .mlCameraZ = CInt(mfLastCZ)

                RotatePoint(0, 0, mfLastCX, mfLastCZ, 1.0F)
            End With

            'Set no events
            Device.IsUsingEventHandlers = False

            'Create our texture...
            If moModelViewTex Is Nothing Then moModelViewTex = New Texture(MyBase.moUILib.oDevice, 512, 512, 1, Usage.RenderTarget, Format.R5G6B5, Pool.Default)
            With MyBase.moUILib.oDevice

                'Store our matrices beforehand...
                matView = MyBase.moUILib.oDevice.Transform.View
                matProj = MyBase.moUILib.oDevice.Transform.Projection

                'Ok, store our original surface
                oOriginal = .GetRenderTarget(0)

                'Get our surface to render to
                oScene = moModelViewTex.GetSurfaceLevel(0)

                'Now, set our render target to the texture's surface
                .SetRenderTarget(0, oScene)

                'Clear out our surface
                .Clear(ClearFlags.Target Or ClearFlags.ZBuffer, System.Drawing.Color.FromArgb(255, 0, 0, 0), 1.0F, 0)

                'Set up our new matrices
                goCamera.SetupMatrices(MyBase.moUILib.oDevice, glCurrentEnvirView)

                .RenderState.ZBufferWriteEnable = True
                .RenderState.AlphaBlendEnable = False

                'render our model here... NOTE: No breaking out of the code here... the render targets need to be reset!
                If iModelID <> 0 Then
                    Dim oObj As BaseMesh = goResMgr.GetMesh(iModelID)

                    If oObj.oMesh Is Nothing = False Then
                        .Transform.World = Matrix.Identity
                        .RenderState.CullMode = Cull.None

                        If CInt(oObj.ShieldXZRadius) > mlInitialZ Then
                            mlInitialZ = CInt(oObj.ShieldXZRadius)
                            mfLastCX = 0
                            goCamera.mlCameraX = 0
                            mlLastCY = mlInitialZ
                            goCamera.mlCameraY = mlLastCY
                            mfLastCZ = mlInitialZ * -2
                            goCamera.mlCameraZ = CInt(mfLastCZ)
                        End If

                        If muSettings.LightQuality > EngineSettings.LightQualitySetting.VSPS1 Then
                            If moShader Is Nothing = True Then moShader = New ModelShader()
                            moShader.RenderHullBuilderMesh(oObj, .Transform.World, iModelID)
                        Else
                            For X As Int32 = 0 To oObj.NumOfMaterials - 1
                                .Material = oObj.Materials(X)
                                .SetTexture(0, oObj.Textures(0))
                                oObj.oMesh.DrawSubset(X)
                            Next X

                            'Now, check for a turret...
                            If oObj.bTurretMesh = True Then
                                .Material = oObj.Materials(0)
                                .SetTexture(0, oObj.Textures(0))
                                oObj.oTurretMesh.DrawSubset(0)
                            End If
                        End If

                    End If
                    'End of model rendering...
                    .RenderState.CullMode = Cull.CounterClockwise
                    oObj = Nothing      'do not dispose, simply release the pointer
                End If

                'Now, restore our original surface to the device
                .SetRenderTarget(0, oOriginal)

                'now, re-enable our Z Buffer
                .RenderState.ZBufferWriteEnable = True
                .RenderState.AlphaBlendEnable = True

                'restore our matrices
                .Transform.View = matView
                .Transform.Projection = matProj
                .Transform.World = Matrix.Identity

                'Release all our objects
                If oScene Is Nothing = False Then oScene.Dispose()
                If oOriginal Is Nothing = False Then oOriginal.Dispose()
                oScene = Nothing
                oOriginal = Nothing

            End With

            'turn events back on
            Device.IsUsingEventHandlers = True

            'Reset our camera
            With goCamera
                'mfLastCX = .mlCameraX
                mlLastCY = .mlCameraY
                'mfLastCZ = .mlCameraZ

                .mlCameraAtX = lCAtX
                .mlCameraAtY = lCAtY
                .mlCameraAtZ = lCAtZ
                .mlCameraX = lCX
                .mlCameraY = lCY
                .mlCameraZ = lCZ
            End With
        End If
    End Sub

    Private Sub frmHullBuilder_OnNewFrameEnd() Handles Me.OnNewFrameEnd
        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return
        'Now... create a sprite
        If mbModelView = True Then
            Using oTmpSprite As Sprite = New Sprite(MyBase.moUILib.oDevice)
                oTmpSprite.Begin(SpriteFlags.AlphaBlend)

                Dim oLoc As System.Drawing.Point = hulMain.GetAbsolutePosition

                Dim fX As Single = oLoc.X
                Dim fY As Single = oLoc.Y

                Dim rcDest As Rectangle = Rectangle.FromLTRB(oLoc.X, oLoc.Y, oLoc.X + hulMain.Width, oLoc.Y + hulMain.Height)
                Dim rcSrc As Rectangle = Rectangle.FromLTRB(64, 0, 448, 512) '384 width

                fX = oLoc.X * CSng(rcSrc.Width / rcDest.Width)
                fY = oLoc.Y * CSng(rcSrc.Height / rcDest.Height)

                oTmpSprite.Draw2D(moModelViewTex, rcSrc, rcDest, Point.Empty, 0, New Point(CInt(fX), CInt(fY)), System.Drawing.Color.White)

                oTmpSprite.End()
                oTmpSprite.Dispose()
            End Using
        End If
    End Sub

    Public Sub AdjustZoom(ByVal lAmt As Int32)
        Dim oVec3 As Vector3 = New Vector3(mfLastCX, mlLastCY, mfLastCZ)
        oVec3.Normalize()
        mfLastCX += lAmt * oVec3.X
        mlLastCY += CInt(lAmt * oVec3.Y)
        mfLastCZ += lAmt * oVec3.Z
        oVec3 = Nothing
    End Sub

    Private Sub btnClear_Click(ByVal sName As String) Handles btnClear.Click
        hulMain.ClearAllConfigs()
        RefreshErrorList()
    End Sub

    Private Sub FrameTypeCheck_Click() Handles chkAtmospheric.Click, chkGround.Click, chkNaval.Click, chkSpace.Click
        If mbIgnoreCheckEvents = True Then Return

        mbIgnoreCheckEvents = True
        chkAtmospheric.Value = chkAtmospheric.Value Or mbAtmosLocked
        chkSpace.Value = chkSpace.Value Or mbSpaceLocked
        chkGround.Value = chkGround.Value Or mbGroundLocked
        chkNaval.Value = chkNaval.Value Or mbNavalLocked
        mbIgnoreCheckEvents = False
    End Sub

    Private Sub tpStructHP_ValueChanged(ByVal sPropName As String, ByRef ctl As ctlTechProp) 'Handles txtStructHP.TextChanged
        If mbLoading = True Then Return
        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderHitpointsChange, CInt(tpStructHP.PropertyValue), -1, -1, "")
        End If
        RefreshErrorList()
    End Sub

    Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
        If moTech Is Nothing Then Return
        If btnDelete.Caption = "Delete Design" Then
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

    Private Sub frmHullBuilder_OnRender() Handles Me.OnRender
        'Ok, first, is moTech nothing?
        If moTech Is Nothing = False Then
            'Ok, now, do the values used on the form currently match the tech?
            Dim bChanged As Boolean = False
            With moTech

                'Dim lID As Int32 = -1
                'If cboStructMat.ListIndex <> -1 Then lID = cboStructMat.ItemData(cboStructMat.ListIndex)
                bChanged = ((.yChassisType And ChassisType.eGroundBased) <> 0) <> chkGround.Value OrElse _
                  ((.yChassisType And ChassisType.eAtmospheric) <> 0) <> chkAtmospheric.Value OrElse _
                  ((.yChassisType And ChassisType.eSpaceBased) <> 0) <> chkSpace.Value OrElse _
                  ((.yChassisType And ChassisType.eNavalBased) <> 0) <> chkNaval.Value OrElse _
                  tpStructHP.PropertyValue <> .StructuralHitPoints OrElse tpHullSize.PropertyValue <> .HullSize OrElse _
                  mlMineralID <> .StructuralMineralID OrElse lscType.SelectedItemID <> .yTypeID OrElse lscSubType.SelectedItemID <> .ySubTypeID OrElse _
                  .ModelID <> ConstructModelID() ' lscFrame.SelectedItemID

                If bChanged = False Then
                    'Ok, check the slots specifically
                    bChanged = hulMain.HasChanged(moTech)
                End If

                If bChanged = True Then
                    If btnDesign.Caption <> "Redesign" Then btnDesign.Caption = "Redesign"
                    If btnDesign.Enabled = False Then btnDesign.Enabled = True

                    If lblErrors.Visible = False Then lblErrors.Visible = True
                    If txtErrors.Visible = False Then txtErrors.Visible = True
                    If mfrmResCost Is Nothing = False AndAlso mfrmResCost.Visible = True Then mfrmResCost.Visible = False
                    If mfrmProdCost Is Nothing = False AndAlso mfrmProdCost.Visible = True Then mfrmProdCost.Visible = False
                ElseIf txtHullName.Caption <> .HullName AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eInvalidDesign AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
                    If btnDesign.Caption <> "Redesign" Then btnDesign.Caption = "Redesign"
                    If btnDesign.Enabled = False Then btnDesign.Enabled = True
                Else
                    If btnDesign.Caption <> "Research" Then btnDesign.Caption = "Research"
                    If moTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eInvalidDesign AndAlso btnDesign.Enabled = True Then btnDesign.Enabled = False
                End If
            End With
        End If
    End Sub

    Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
        If goTutorial Is Nothing Then goTutorial = New TutorialManager()
        goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eHullBuilder)
    End Sub

    Private Sub btnShowMinDisplay_Click(ByVal sName As String) Handles btnShowMinDisplay.Click
        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
        If ofrm Is Nothing Then
            ofrm = New frmMinDetail(goUILib)
            ofrm.ShowMineralDetail(Me.Left, Me.Top, 512, -1)
            ofrm.SetAsHullBuilder()
            ofrm.ClearHighlights()
            For X As Int32 = 0 To glMineralPropertyUB
                If glMineralPropertyIdx(X) <> -1 AndAlso goMineralProperty(X).MineralPropertyName.ToUpper = "HARDNESS" Then
                    ofrm.HighlightProperty(goMineralProperty(X).ObjectID, 2, 10)
                    Exit For
                End If
            Next X
        Else : goUILib.RemoveWindow(ofrm.ControlName)
        End If
        ofrm.bRaiseSelectEvent = True
        Try
            RemoveHandler ofrm.MineralSelected, AddressOf Mineral_Selected
        Catch
        End Try
        AddHandler ofrm.MineralSelected, AddressOf Mineral_Selected
    End Sub

    Private Function ValidateTutorialSettings() As Boolean
        Dim lStepID As Int32 = NewTutorialManager.GetTutorialStepID
        If lStepID < 135 Then
            'building the tank hull... ensure the following:
            ' Storm II hull is selected
            If lscFrame.SelectedItemID <> 33 Then
                MyBase.moUILib.AddNotification("Invalid frame selected, select the Storm II Hover Tank frame.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                NewTutorialManager.AlertControl("frmHullBuilder.fraSelectFrame.lscFrame")
                Return False
            End If
            ' hull size is 540
            If tpHullSize.PropertyValue <> 540 Then
                MyBase.moUILib.AddNotification("The Hull Size must equal 540.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                NewTutorialManager.AlertControl("frmHullBuilder.fraHullDesignDetails.txtHullSize")
                Return False
            End If
            ' enochine
            If mlMineralID < 1 Then
                MyBase.moUILib.AddNotification("Select Enochine for the hull material.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                NewTutorialManager.AlertControl("frmHullBuilder.fraHullDesignDetails.cboStructMat")
                Return False
            End If
            ' hitpoints is 540
            If tpStructHP.PropertyValue <> 540 Then
                MyBase.moUILib.AddNotification("The Structural Hitpoints must equal 540.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                NewTutorialManager.AlertControl("frmHullBuilder.fraHullDesignDetails.txtStructHP")
                Return False
            End If
            ' 20 slots for engines
            If hulMain.GetSlotConfigCnt(SlotType.eRear, SlotConfig.eEngines) <> 20 Then
                MyBase.moUILib.AddNotification("Allocate 20 slots for engines in the rear slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            ' 10 radar slots in yellow
            If hulMain.GetSlotConfigCnt(SlotType.eAllArc, SlotConfig.eRadar) <> 10 Then
                MyBase.moUILib.AddNotification("Allocate 10 slots for radar in the yellow all-arc slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            ' 6 weapon 1 slots in yellow
            If hulMain.GetSlotConfigCntWithGroup(SlotType.eAllArc, SlotConfig.eWeapons, 1) <> 6 Then
                MyBase.moUILib.AddNotification("Allocate 6 slots for weapon group 1 in the yellow all-arc slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            ' 6 weapon 2 slots in yellow
            If hulMain.GetSlotConfigCntWithGroup(SlotType.eAllArc, SlotConfig.eWeapons, 2) <> 6 Then
                MyBase.moUILib.AddNotification("Allocate 6 slots for weapon group 2 in the yellow all-arc slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            ' 6 weapon 3 slots in yellow
            If hulMain.GetSlotConfigCntWithGroup(SlotType.eAllArc, SlotConfig.eWeapons, 3) <> 6 Then
                MyBase.moUILib.AddNotification("Allocate 6 slots for weapon group 3 in the yellow all-arc slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            ' 6 weaopn 4 slots in yellow
            If hulMain.GetSlotConfigCntWithGroup(SlotType.eAllArc, SlotConfig.eWeapons, 4) <> 6 Then
                MyBase.moUILib.AddNotification("Allocate 6 slots for weapon group 4 in the yellow all-arc slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            ' 12 cargo slots in yellow
            If hulMain.GetSlotConfigCnt(SlotType.eAllArc, SlotConfig.eCargoBay) <> 12 Then
                MyBase.moUILib.AddNotification("Allocate 12 slots for cargo bay in the yellow all-arc slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            'now, check our overall counts
            If hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eCargoBay) <> 12 Then
                MyBase.moUILib.AddNotification("Allocate only 12 slots of cargo bay in the yellow all-arc slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eEngines) <> 20 Then
                MyBase.moUILib.AddNotification("Allocate only 20 slots of engines in the rear arc slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eHangar) <> 0 Then
                MyBase.moUILib.AddNotification("Allocate exactly 0 slots for hangar.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eRadar) <> 10 Then
                MyBase.moUILib.AddNotification("Allocate exactly 10 radar slots in the yellow all-arc slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eShields) <> 0 Then
                MyBase.moUILib.AddNotification("Allocate exactly 0 slots for shields.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eWeapons) <> 24 Then
                MyBase.moUILib.AddNotification("Allocate only 24 slots for weapons, 6 for 4 weapon groups.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
        ElseIf lStepID > 150 AndAlso lStepID < 160 Then
            'building the turret hull... ensure the following:
            ' Defense I hull is selected
            If lscFrame.SelectedItemID <> 16 Then
                MyBase.moUILib.AddNotification("Invalid frame selected, select the Defense I frame.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                NewTutorialManager.AlertControl("frmHullBuilder.fraSelectFrame.lscFrame")
                Return False
            End If
            ' hull is 270
            If tpHullSize.PropertyValue <> 2700 Then
                MyBase.moUILib.AddNotification("The Hull Size must equal 2700.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                NewTutorialManager.AlertControl("frmHullBuilder.fraHullDesignDetails.txtHullSize")
                Return False
            End If
            ' enochine
            If mlMineralID < 1 Then
                MyBase.moUILib.AddNotification("Select Enochine for the hull material.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                NewTutorialManager.AlertControl("frmHullBuilder.fraHullDesignDetails.cboStructMat")
                Return False
            End If
            ' hitpoints is 270
            If tpStructHP.PropertyValue <> 2700 Then
                MyBase.moUILib.AddNotification("The Structural Hitpoints must equal 2700.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                NewTutorialManager.AlertControl("frmHullBuilder.fraHullDesignDetails.txtStructHP")
                Return False
            End If
            ' 13 radar in red right arc
            If hulMain.GetSlotConfigCnt(SlotType.eRight, SlotConfig.eRadar) <> 13 Then
                MyBase.moUILib.AddNotification("Allocate 13 slots for radar in the red right arc slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            ' 2 weapon slots in blue
            If hulMain.GetSlotConfigCntWithGroup(SlotType.eFront, SlotConfig.eWeapons, 1) <> 2 Then
                MyBase.moUILib.AddNotification("Allocate 2 slots for weapon group 1 in the front blue slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            ' 2 weapon slots in red
            If hulMain.GetSlotConfigCntWithGroup(SlotType.eRight, SlotConfig.eWeapons, 2) <> 2 Then
                MyBase.moUILib.AddNotification("Allocate 2 slots for weapon group 2 in the red right slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            ' 2 weapon slots in green
            If hulMain.GetSlotConfigCntWithGroup(SlotType.eRear, SlotConfig.eWeapons, 3) <> 2 Then
                MyBase.moUILib.AddNotification("Allocate 2 slots for weapon group 3 in the green rear slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            ' 2 weapon slots in purple 
            If hulMain.GetSlotConfigCntWithGroup(SlotType.eLeft, SlotConfig.eWeapons, 4) <> 2 Then
                MyBase.moUILib.AddNotification("Allocate 2 slots for weapon group 4 in the purplse left slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            'now, check our overall counts
            If hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eCargoBay) <> 0 Then
                MyBase.moUILib.AddNotification("Allocate exactly 0 slots of cargo bay.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eEngines) <> 0 Then
                MyBase.moUILib.AddNotification("Allocate exactly 0 slots of engines.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eHangar) <> 0 Then
                MyBase.moUILib.AddNotification("Allocate exactly 0 slots for hangar.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eRadar) <> 13 Then
                MyBase.moUILib.AddNotification("Allocate exactly 13 radar slots in the right red slots.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eShields) <> 0 Then
                MyBase.moUILib.AddNotification("Allocate exactly 0 slots for shields.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If hulMain.GetSlotConfigCnt(SlotType.eNoChange, SlotConfig.eWeapons) <> 8 Then
                MyBase.moUILib.AddNotification("Allocate only 8 slots for weapons, 2 for 4 weapon groups in each arc.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
        End If

        Return True
    End Function
 
    Private Sub Mineral_Selected(ByVal lMineralID As Int32)
        mlMineralID = lMineralID

        Dim sMinName As String = "Unknown Mineral"
        For X As Int32 = 0 To glMineralUB
            If glMineralIdx(X) = lMineralID Then
                sMinName = goMinerals(X).MineralName
                Exit For
            End If
        Next X
        lblStructMatItem.Caption = sMinName

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eHullBuilderMaterialSelected, lMineralID, ObjectType.eMineral, -1, "")
        End If
        RefreshErrorList()

        'If goUILib Is Nothing = False Then goUILib.RemoveWindow("frmMinDetail")
        Me.IsDirty = True
    End Sub

    Public Sub ViewResults(ByRef oTech As Base_Tech, ByVal lProdFactor As Integer)
        If oTech Is Nothing Then Return

        moTech = CType(oTech, HullTech)

        Me.btnDesign.Enabled = False

        mbIgnoreLabelScrollEvents = True
        mbLoading = True

        lblErrors.Visible = False
        txtErrors.Visible = False

        With moTech
            mbIgnoreLabelScrollEvents = False
            lscType.SelectByID(.yTypeID)
            lscSubType.SelectByID(.ySubTypeID)
            Dim iMesh As Int16 = 0
            Dim iTexNum As Int16 = 0
            Dim iBumpNum As Int16 = 0
            Dim iIllumNum As Int16 = 0
            DeconstructModelID(.ModelID, iMesh, iTexNum, iBumpNum, iIllumNum)
            lscFrame.SelectByID(iMesh)
            If lscDiffuse.Enabled = True Then
                Dim bLoad As Boolean = mbLoading
                Dim bIgnore As Boolean = mbIgnoreLabelScrollEvents
                Dim bView As Boolean = mbModelView
                mbLoading = False : mbIgnoreLabelScrollEvents = False
                lscDiffuse.SelectByID(iTexNum)
                mbLoading = bLoad : mbIgnoreLabelScrollEvents = bIgnore
                mbModelView = bView
                hulMain.Visible = Not mbModelView
            End If
            If lscNormal.Enabled = True Then
                Dim bLoad As Boolean = mbLoading
                Dim bIgnore As Boolean = mbIgnoreLabelScrollEvents
                Dim bView As Boolean = mbModelView
                mbLoading = False : mbIgnoreLabelScrollEvents = False
                lscNormal.SelectByID(iBumpNum)
                mbLoading = bLoad : mbIgnoreLabelScrollEvents = bIgnore
                mbModelView = bView
                hulMain.Visible = Not mbModelView
            End If
            If lscIllum.Enabled = True Then
                Dim bLoad As Boolean = mbLoading
                Dim bIgnore As Boolean = mbIgnoreLabelScrollEvents
                Dim bView As Boolean = mbModelView
                mbLoading = False : mbIgnoreLabelScrollEvents = False
                lscIllum.SelectByID(iIllumNum)
                mbLoading = bLoad : mbIgnoreLabelScrollEvents = bIgnore
                mbModelView = bView
                hulMain.Visible = Not mbModelView
            End If

            mbIgnoreLabelScrollEvents = True

            hulMain.SetFromHullTech(moTech)

            'cboStructMat.FindComboItemData(.StructuralMineralID)
            Dim bFound As Boolean = False
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = .StructuralMineralID AndAlso goMinerals(X).bDiscovered = True Then
                    bFound = True
                    Exit For
                End If
            Next X
            If bFound = True Then Mineral_Selected(.StructuralMineralID)

            tpStructHP.PropertyValue = .StructuralHitPoints
            tpStructHP.PropertyLocked = True
            tpHullSize.PropertyValue = .HullSize
            tpHullSize.PropertyLocked = True
            Me.txtHullName.Caption = .HullName

            mbIgnoreCheckEvents = True
            chkGround.Value = (.yChassisType And ChassisType.eGroundBased) <> 0
            chkAtmospheric.Value = (.yChassisType And ChassisType.eAtmospheric) <> 0
            chkSpace.Value = (.yChassisType And ChassisType.eSpaceBased) <> 0
            chkNaval.Value = (.yChassisType And ChassisType.eNavalBased) <> 0
            mbIgnoreCheckEvents = False

            txtHullAllotment.Caption = hulMain.GetHullSummary(.HullSize)
            mlCrewReqPerc = CInt(tpHullSize.PropertyValue / 100)
            If lscType.SelectedItemID = 2 Then mlCrewReqPerc = 0
            lblCrewReqs.Caption = "Suggested Crew: " & mlCrewReqPerc
        End With

        If oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eComponentResearching AndAlso HasAliasedRights(AliasingRights.eAddResearch) = True Then
            btnDesign.Caption = "Research"
            btnDesign.Enabled = True
        End If
        If gbAliased = False Then
            btnRename.Visible = oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched
            btnDelete.Visible = True 'oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched
        End If
        optArmor_Click()

        'Now... what state is the tech in?
        If oTech.ComponentDevelopmentPhase > Base_Tech.eComponentDevelopmentPhase.eComponentDesign Then
            'Ok, show it's research and production cost
            If mfrmResCost Is Nothing Then mfrmResCost = New frmProdCost(MyBase.moUILib, "frmResCost", True)
            With mfrmResCost
                .Visible = True
                .Left = txtErrors.Left
                .Top = txtErrors.Top
                .Width = 220        
                .Height = 290
                .SetFromProdCost(oTech.oResearchCost, lProdFactor, True, 0, 0)
            End With
            Me.AddChild(CType(mfrmResCost, UIControl))

            If mfrmProdCost Is Nothing Then mfrmProdCost = New frmProdCost(MyBase.moUILib, "frmProdCost", True)
            With mfrmProdCost
                .Visible = True
                .Left = mfrmResCost.Left + mfrmResCost.Width + 10
                .Top = mfrmResCost.Top
                .Height = 290
                .Width = 220
                .SetFromProdCost(oTech.oProductionCost, 1000, False, 0, moTech.PowerRequired)
            End With
            Me.AddChild(CType(mfrmProdCost, UIControl))
        ElseIf oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
            If mfrmResCost Is Nothing Then mfrmResCost = New frmProdCost(MyBase.moUILib, "frmResCost", True)
            With mfrmResCost
                .Visible = True
                .Left = txtErrors.Left
                .Top = txtErrors.Top
                .SetFromFailureCode(oTech.ErrorCodeReason)
            End With
            Me.AddChild(CType(mfrmResCost, UIControl))
        End If

        mbIgnoreLabelScrollEvents = False
        mbLoading = False
    End Sub

    Private Sub txtHgrDrNum_TextChanged() Handles txtHgrDrNum.TextChanged
        If mbLoading = True Then Return
        If optHangarDoor.Value = True Then
            If CInt(Val(txtHgrDrNum.Caption)) > 0 Then hulMain.HighLightAcceptableSlots(SlotConfig.eHangarDoor, CInt(Val(txtHgrDrNum.Caption)))
        End If
    End Sub

    Private Sub AppearanceScroller_ItemChanged(ByVal lID As Int32, ByVal sName As String)
        If mbLoading = True OrElse mbIgnoreLabelScrollEvents = True Then Return
        If mbModelView = False Then btnSwitchView_Click(btnSwitchView.ControlName)
    End Sub

    Private Sub DiffuseScroller_ItemChanged(ByVal lID As Int32, ByVal sName As String)
        If mbLoading = True OrElse mbIgnoreLabelScrollEvents = True Then Return
        mbIgnoreLabelScrollEvents = True
        Try
            Dim lModelID As Int32 = lscFrame.SelectedItemID
            Dim oMesh As BaseMesh = goResMgr.GetMesh(lModelID)
            If oMesh Is Nothing = False AndAlso oMesh.bMTTexCapable = True AndAlso oMesh.oSharedTex Is Nothing = False Then
                Dim lTexNum As Int32 = lscDiffuse.SelectedItemID

                'ok, clear our normal and illum list
                lscNormal.Clear()
                lscIllum.Clear()

                With oMesh.oSharedTex(lTexNum)
                    If lModelID = 141 Then
                        'Ok, normal map, we have 2
                        If .oNormalTexture Is Nothing = False Then
                            Dim lPow As Int32 = 1
                            Dim lVal As Int32 = 1
                            For X As Int32 = 0 To Math.Min(1, .oNormalTexture.GetUpperBound(0))
                                If (oMesh.lSharedTexOpt(lTexNum) And lPow) <> 0 OrElse X = 0 Then
                                    lscNormal.AddItem(X, lVal.ToString)
                                    lVal += 1
                                End If
                                lPow *= 2
                            Next X
                        End If

                        'and the illum map has 4
                        'add the default
                        lscIllum.AddItem(0, "1")
                        lscIllum.AddItem(1, "Guild")
                    Else
                        If .oNormalTexture Is Nothing = False Then
                            Dim lPow As Int32 = 1
                            Dim lVal As Int32 = 1
                            For X As Int32 = 0 To .oNormalTexture.GetUpperBound(0)
                                If (oMesh.lSharedTexOpt(lTexNum) And lPow) <> 0 OrElse X = 0 Then
                                    lscNormal.AddItem(X, lVal.ToString)
                                    lVal += 1
                                End If
                                lPow *= 2
                            Next X
                        End If
                        If .oIllumTexture Is Nothing = False Then
                            Dim lPow As Int32 = 16
                            Dim lVal As Int32 = 0
                            For X As Int32 = 0 To .oIllumTexture.GetUpperBound(0)
                                If (oMesh.lSharedTexOpt(lTexNum) And lPow) <> 0 OrElse X = 0 Then
                                    lscIllum.AddItem(X, lVal.ToString)
                                    lVal += 1
                                End If
                                lPow *= 2
                            Next X
                        End If
                    End If
                End With
                lscNormal.SelectByID(0)
                If lModelID = 141 Then lscIllum.SelectByID(1) Else lscIllum.SelectByID(0)

                lscNormal.ResetCurrentIndex()
                lscIllum.ResetCurrentIndex()

                If lscNormal.ListCount > 1 Then lscNormal.Enabled = True Else lscNormal.Enabled = False
                If lscIllum.ListCount > 1 Then lscIllum.Enabled = True Else lscIllum.Enabled = False
            End If
        Catch
        End Try

        mbIgnoreLabelScrollEvents = False
        AppearanceScroller_ItemChanged(lID, sName)
    End Sub

    Private Function ConstructModelID() As Int16 'ByVal iMesh As Int16, ByVal iTexNum As Int16, ByVal iBumpNum As Int16, ByVal iIllumNum As Int16) As Int16
        Dim iMesh As Int16 = CShort(lscFrame.SelectedItemID)
        Dim iTexNum As Int16 = CShort(lscDiffuse.SelectedItemID)
        Dim iNormalNum As Int16 = CShort(lscNormal.SelectedItemID)
        Dim iIllumNum As Int16 = CShort(lscIllum.SelectedItemID)

        If lscDiffuse.Enabled = False Then iTexNum = 0
        If lscNormal.Enabled = False Then iNormalNum = 0
        If lscIllum.Enabled = False Then iIllumNum = 0

        'ok, here we go.... first here is the map of the final result
        '8 - modelid = 255
        '5 - texture number - 31
        '2 - bump map number - 4
        '1 - illumination - 2

        Dim iResult As Int16 = iMesh
        If iMesh = 141 Then
            iResult += (iTexNum * 256S)

            If iIllumNum <> 0 Then
                Dim iActual As Int16 = 0
                If goCurrentPlayer Is Nothing = False Then
                    If goCurrentPlayer.oGuild Is Nothing = False Then
                        Select Case goCurrentPlayer.oGuild.ObjectID
                            Case 13
                                iActual = 1
                            Case 14
                                iActual = 2
                            Case 17
                                iActual = 3
                        End Select
                    End If
                End If
                iIllumNum = iActual
            End If
            iResult += (iIllumNum * 8192S)

            If iNormalNum <> 0 Then iResult = CShort(iResult Or -32768)
        Else
            iResult += (iTexNum * 256S)
            iResult += (iNormalNum * 8192S)

            If iIllumNum <> 0 Then iResult = CShort(iResult Or -32768)
        End If
        

        Return iResult
    End Function

    Private Sub DeconstructModelID(ByVal iModelID As Int16, ByRef iMesh As Int16, ByRef iTexNum As Int16, ByRef iBumpNum As Int16, ByRef iIllumNum As Int16)

        iMesh = (iModelID And 255S)

        If iMesh = 141 Then
            iTexNum = (iModelID And 7936S) \ 256S  'fine
            iIllumNum = (iModelID And 24576S) \ 8192S
            If iIllumNum <> 0 Then iIllumNum = 1
            If (iModelID And -32768) = 0 Then iBumpNum = 0 Else iBumpNum = 1
        Else
            iTexNum = (iModelID And 7936S) \ 256S
            iBumpNum = (iModelID And 24576S) \ 8192S
            If (iModelID And -32768) = 0 Then iIllumNum = 0 Else iIllumNum = 1
        End If

    End Sub

    Private msRenameVal As String = ""
    Private Sub btnRename_Click(ByVal sName As String) Handles btnRename.Click
        If moTech Is Nothing = False Then
            If moTech.OwnerID = glPlayerID Then
                If goCurrentPlayer.blCredits < 10000000 Then
                    MyBase.moUILib.AddNotification("You require 10,000,000 credits to rename a tech.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If
                Dim sVal As String = txtHullName.Caption.Trim
                If sVal.Length > 20 Then
                    MyBase.moUILib.AddNotification("The new name is too long!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                ElseIf sVal.Length = 0 Then
                    MyBase.moUILib.AddNotification("You must enter a name for this tech to be renamed to.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If
                If sVal.ToUpper <> moTech.GetComponentName().ToUpper Then
                    msRenameVal = sVal
                    Dim oFrm As New frmMsgBox(MyBase.moUILib, "Renaming a tech design costs 10,000,000 credits in order to have the GTC update all of the registries. Do you wish to proceed?", MsgBoxStyle.YesNo, "Confirm Rename")
                    oFrm.Visible = True
                    AddHandler oFrm.DialogClosed, AddressOf RenameMsgBoxResult
                Else
                    MyBase.moUILib.AddNotification("The new name is not different enough from the old name.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
            End If
        End If
    End Sub
    Private Sub RenameMsgBoxResult(ByVal lResult As MsgBoxResult)
        If lResult = MsgBoxResult.Yes Then
            Dim yMsg(27) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityName).CopyTo(yMsg, lPos) : lPos += 2
            moTech.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
            System.Text.ASCIIEncoding.ASCII.GetBytes(msRenameVal).CopyTo(yMsg, lPos) : lPos += 20
            MyBase.moUILib.SendMsgToPrimary(yMsg)
            moTech.SetComponentName(msRenameVal)
        End If
    End Sub

    Public Sub SetFromClaimableModel(ByVal lModelID As Int32)
        btnDesign.Enabled = False
        btnDesign.Visible = False

        Dim oModelDef As ModelDef = goModelDefs.GetModelDef(lModelID)
        If oModelDef Is Nothing = False Then
            lscType.Clear()
            Dim sDisplay As String = ""
            Select Case oModelDef.TypeID
                Case 0 : sDisplay = "Capital Ship"
                Case 1 : sDisplay = "Escort Ship"
                Case 2 : sDisplay = "Facility"
                Case 3 : sDisplay = "Fighter"
                Case 4 : sDisplay = "Small Vehicle"
                Case 5 : sDisplay = "Tank"
                Case 6 : sDisplay = "Transport"
                Case 7 : sDisplay = "Utility Vehicle"
                Case 8 : sDisplay = "Naval"
                Case Else
                    sDisplay = "Unknown"
            End Select
            lscType.AddItem(oModelDef.TypeID, sDisplay)
            lscType.SelectByID(oModelDef.TypeID)

            lscSubType.SelectByID(oModelDef.SubTypeID)
            If lscSubType.SelectedItemID <> oModelDef.SubTypeID Then
                lscSubType.AddItem(oModelDef.SubTypeID, "Undiscovered")
                lscSubType.SelectByID(oModelDef.SubTypeID)
            End If

            lscFrame.SelectByID(oModelDef.ModelID)
            If lscFrame.SelectedItemID <> oModelDef.ModelID Then
                lscFrame.AddItem(oModelDef.ModelID, oModelDef.FrameName)
                lscFrame.SelectByID(oModelDef.ModelID)
            End If

            lscFrame.Enabled = False
            lscSubType.Enabled = False
            lscType.Enabled = False
        End If
    End Sub
End Class
