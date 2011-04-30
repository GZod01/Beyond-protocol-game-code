Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmRadarBuilder
    Inherits frmTechBuilder

    'Inherits UIWindow

    Private lblRadarType As UILabel
    'Private WithEvents optActive As UIOption
    'Private WithEvents optPassive As UIOption

    Private cboHullType As UIComboBox

    Private lblEmitter As UILabel
    Private lblDetection As UILabel
    Private lblCollection As UILabel
    Private lblCasing As UILabel

    Private lblJamImmune As UILabel
    Private lblJamEffect As UILabel
    Private cboJamImmune As UIComboBox
    Private cboJamEffect As UIComboBox

    'Private lblWepAcc As UILabel
    'Private lblScanRes As UILabel
    'Private lblOptRng As UILabel
    'Private lblMaxRng As UILabel
    'Private lblDisRes As UILabel
    'Private lblJamStrength As UILabel
    'Private lblJamTargets As UILabel
    'Private txtWepAcc As UITextBox
    'Private txtScanRes As UITextBox
    'Private txtOptRng As UITextBox
    'Private txtMaxRng As UITextBox
    'Private txtDisRes As UITextBox
    'Private txtJamStrength As UITextBox
    'Private txtJamTargets As UITextBox
    Private tpWepAcc As ctlTechProp
    Private tpScanRes As ctlTechProp
    Private tpOptRng As ctlTechProp
    Private tpMaxRng As ctlTechProp
    Private tpDisRes As ctlTechProp
    Private tpJamStrength As ctlTechProp
    Private tpJamTargets As ctlTechProp
    Private chkAreaEffectJamming As UICheckBox

    'Private cboCasing As UIComboBox
    'Private cboCollection As UIComboBox
    'Private cboDetection As UIComboBox
    'Private cboEmitter As UIComboBox
    Private stmCasing As ctlSetTechMineral
    Private stmCollection As ctlSetTechMineral
    Private stmDetection As ctlSetTechMineral
    Private stmEmitter As ctlSetTechMineral

    Private tpCasing As ctlTechProp
    Private tpCollection As ctlTechProp
    Private tpDetection As ctlTechProp
    Private tpEmitter As ctlTechProp

    Private lblRadarBuilder As UILabel

    Private WithEvents btnDesign As UIButton
    Private WithEvents btnCancel As UIButton
    Private WithEvents btnDelete As UIButton
    Private WithEvents btnHelp As UIButton
    Private WithEvents btnRename As UIButton
    Private WithEvents btnAutoFill As UIButton

    Private WithEvents lblTechName As UILabel
    Private WithEvents txtTechName As UITextBox

    Private lblDesignFlaw As UILabel

    Private mlEntityIndex As Int32 = -1

    Private moTech As RadarTech = Nothing
    Private mfrmResCost As frmProdCost = Nothing
    Private mfrmProdCost As frmProdCost = Nothing

    Private mbIgnoreOptionEvents As Boolean = False

    Private mlMaxOptRange As Int32 = 50
    Private mlMaxMaxRange As Int32 = 50
    Private mlMaxScanRes As Int32 = 50
    Private mlMaxWpnAcc As Int32 = 50
    Private mlMaxDisRes As Int32 = 50
    Private mlMaxJamStr As Int32 = 50
    Private mlMaxJamTargets As Int32 = 1

    Private mbIgnoreValueChange As Boolean = False
    Private mbImpossibleDesign As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmRadarBuilder initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eRadarBuilder
            .ControlName = "frmRadarBuilder"
            .Left = 5
            .Top = 10
            .Width = 490 '360
            .Height = 512
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = False
            .mbAcceptReprocessEvents = True
        End With

        'txtTechName initial props
        txtTechName = New UITextBox(oUILib)
        With txtTechName
            .ControlName = "txtTechName"
            .Left = 170
            .Top = 50
            .Width = 175
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 20
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtTechName, UIControl))

        'lblTechName initial props
        lblTechName = New UILabel(oUILib)
        With lblTechName
            .ControlName = "lblTechName"
            .Left = 10
            .Top = 50
            .Width = 164
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Radar Name:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTechName, UIControl))

        'lblRadarType initial props
        lblRadarType = New UILabel(oUILib)
        With lblRadarType
            .ControlName = "lblRadarType"
            .Left = 10
            .Top = 75
            .Width = 161
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hull Type:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblRadarType, UIControl))

        ''optActive initial props
        'optActive = New UIOption(oUILib)
        'With optActive
        '    .ControlName = "optActive"
        '    .Left = 175
        '    .Top = 75
        '    .Width = 55
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Active"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(6, DrawTextFormat)
        '    .DisplayType = UIOption.eOptionTypes.eSmallOption
        '    .Value = False
        '    .ToolTipText = "Click to make this an active radar system." & vbCrLf & "Active radar systems are more robust but also require more power."
        'End With
        'Me.AddChild(CType(optActive, UIControl))

        ''optPassive initial props
        'optPassive = New UIOption(oUILib)
        'With optPassive
        '    .ControlName = "optPassive"
        '    .Left = 275
        '    .Top = 75
        '    .Width = 65
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Passive"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(6, DrawTextFormat)
        '    .DisplayType = UIOption.eOptionTypes.eSmallOption
        '    .Value = True
        '    .ToolTipText = "Click to make this a passive radar system." & vbCrLf & "Passive radar systems are not as robust as active radar but passive requires less power."
        'End With
        'Me.AddChild(CType(optPassive, UIControl))
        '=========================================================================
        ''lblWepAcc initial props
        'lblWepAcc = New UILabel(oUILib)
        'With lblWepAcc
        '    .ControlName = "lblWepAcc"
        '    .Left = 10
        '    .Top = 100
        '    .Width = 161
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Desired Weapon Accuracy:"
        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblWepAcc, UIControl))

        ''txtWepAcc initial props
        'txtWepAcc = New UITextBox(oUILib)
        'With txtWepAcc
        '    .ControlName = "txtWepAcc"
        '    .Left = 170
        '    .Top = 100
        '    .Width = 41
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "0"
        '    '.ForeColor = muSettings.InterfaceBorderColor
        '    .ForeColor = muSettings.InterfaceTextBoxForeColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        '    '.BackColorEnabled = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '    .MaxLength = 3
        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .ToolTipText = "Determines the overall ability of the radar system" & vbCrLf & "to execute firing sequences and assist gunners." & vbCrLf & "Valid Range is any value between 0 and 255."
        '    .bNumericOnly = True
        'End With
        'Me.AddChild(CType(txtWepAcc, UIControl))

        ''lblScanRes initial props
        'lblScanRes = New UILabel(oUILib)
        'With lblScanRes
        '    .ControlName = "lblScanRes"
        '    .Left = 10
        '    .Top = 125
        '    .Width = 161
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Desired Scan Resolution:"
        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblScanRes, UIControl))

        ''lblOptRng initial props
        'lblOptRng = New UILabel(oUILib)
        'With lblOptRng
        '    .ControlName = "lblOptRng"
        '    .Left = 10
        '    .Top = 150
        '    .Width = 161
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Desired Optimum Range:"
        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)'
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblOptRng, UIControl))

        ''lblMaxRng initial props
        'lblMaxRng = New UILabel(oUILib)
        'With lblMaxRng
        '    .ControlName = "lblMaxRng"
        '    .Left = 230
        '    .Top = 150
        '    .Width = 55
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Maximum:"
        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)'
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblMaxRng, UIControl))

        ''lblDisRes initial props
        'lblDisRes = New UILabel(oUILib)
        'With lblDisRes
        '    .ControlName = "lblDisRes"
        '    .Left = 10
        '    .Top = 175
        '    .Width = 161
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Desired Disruption Resist:"
        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblDisRes, UIControl))

        'lblJamImmune initial props
        lblJamImmune = New UILabel(oUILib)
        With lblJamImmune
            .ControlName = "lblJamImmune"
            .Left = 10
            .Top = 200
            .Width = 161
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Jamming Immunity:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblJamImmune, UIControl))

        'lblJamEffect initial props
        lblJamEffect = New UILabel(oUILib)
        With lblJamEffect
            .ControlName = "lblJamEffect"
            .Left = 10
            .Top = 225
            .Width = 161
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Jamming Effect:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblJamEffect, UIControl))

        ''lblJamStrength initial props
        'lblJamStrength = New UILabel(oUILib)
        'With lblJamStrength
        '    .ControlName = "lblJamStrength"
        '    .Left = 10
        '    .Top = 250
        '    .Width = 161
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Jamming Strength:"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblJamStrength, UIControl))

        ''lblJamTargets initial props
        'lblJamTargets = New UILabel(oUILib)
        'With lblJamTargets
        '    .ControlName = "lblJamTargets"
        '    .Left = 10
        '    .Top = 275
        '    .Width = 161
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Jamming Targets:"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblJamTargets, UIControl))


        'lblEmitter initial props
        lblEmitter = New UILabel(oUILib)
        With lblEmitter
            .ControlName = "lblEmitter"
            .Left = 10
            .Top = 300
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Emitter:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblEmitter, UIControl))

        'lblDetection initial props
        lblDetection = New UILabel(oUILib)
        With lblDetection
            .ControlName = "lblDetection"
            .Left = 10
            .Top = 325
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Detection:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblDetection, UIControl))

        'lblCollection initial props
        lblCollection = New UILabel(oUILib)
        With lblCollection
            .ControlName = "lblCollection"
            .Left = 10
            .Top = 350
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Collector:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblCollection, UIControl))

        'lblCasin initial props
        lblCasing = New UILabel(oUILib)
        With lblCasing
            .ControlName = "lblCasing"
            .Left = 10
            .Top = 375
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Casing:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblCasing, UIControl))

        tpWepAcc = New ctlTechProp(oUILib)
        With tpWepAcc
            .ControlName = "tpWepAcc"
            .Enabled = True
            .Height = 18
            .Left = 10
            .MaxValue = 255
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Weapon Accuracy:"
            .PropertyValue = 0
            .Top = 95
            .Visible = True
            '.Width = 512
            .ToolTipText = "Determines the overall ability of the radar system" & vbCrLf & "to execute firing sequences and assist gunners." & vbCrLf & "Valid Range is any value between 0 and 255."
        End With
        Me.AddChild(CType(tpWepAcc, UIControl))

        ''txtScanRes initial props
        'txtScanRes = New UITextBox(oUILib)
        'With txtScanRes
        '    .ControlName = "txtScanRes"
        '    .Left = 170
        '    .Top = 125
        '    .Width = 41
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "0"
        '    '.ForeColor = muSettings.InterfaceBorderColor
        '    .ForeColor = muSettings.InterfaceTextBoxForeColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        '    '.BackColorEnabled = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '    .MaxLength = 3
        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .ToolTipText = "Determines scanning strength of the radar system." & vbCrLf & "Higher values will reward more detailed results in target scans" & vbCrLf & "and in long range (outside optimum, inside maximum) scans." & vbCrLf & "Valid range is any value between 0 and 255."
        '    .bNumericOnly = True
        'End With
        'Me.AddChild(CType(txtScanRes, UIControl))
        tpScanRes = New ctlTechProp(oUILib)
        With tpScanRes
            .ControlName = "tpScanRes"
            .Enabled = True
            .Height = 18
            .Left = 10
            .MaxValue = 255
            .MinValue = 0
            .PropertyLocked = False
            '.PropertyName = "Scan Resolution:"
            .PropertyName = "Point Defense Accuracy:"
            .PropertyValue = 0
            .Top = 115
            .Visible = True
            '.Width = 512
            '.ToolTipText = "Determines scanning strength of the radar system." & vbCrLf & "Higher values will reward more detailed results in target scans" & vbCrLf & "and in long range (outside optimum, inside maximum) scans." & vbCrLf & "Valid range is any value between 0 and 255."
            .ToolTipText = "Accuracy of Point Defense Weapons"
        End With
        Me.AddChild(CType(tpScanRes, UIControl))

        ''txtOptRng initial props
        'txtOptRng = New UITextBox(oUILib)
        'With txtOptRng
        '    .ControlName = "txtOptRng"
        '    .Left = 170
        '    .Top = 150
        '    .Width = 41
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "0"
        '    '.ForeColor = muSettings.InterfaceBorderColor
        '    .ForeColor = muSettings.InterfaceTextBoxForeColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        '    '.BackColorEnabled = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '    .MaxLength = 3
        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .ToolTipText = "Indicates short-range visibility. Units will attack enemies within" & vbCrLf & "this range. Entities within this range will be fully visible." & vbCrLf & "Valid range is any value between 0 and 255."
        '    .bNumericOnly = True
        'End With
        'Me.AddChild(CType(txtOptRng, UIControl))
        tpOptRng = New ctlTechProp(oUILib)
        With tpOptRng
            .ControlName = "tpOptRng"
            .Enabled = True
            .Height = 18
            .Left = 10
            .MaxValue = 255
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Visible Range:"
            .PropertyValue = 0
            .Top = 135
            .Visible = True
            '.Width = 512
            .ToolTipText = "Indicates short-range visibility. Units will attack enemies within" & vbCrLf & "this range. Entities within this range will be fully visible." & vbCrLf & "Valid range is any value between 0 and 255."
        End With
        Me.AddChild(CType(tpOptRng, UIControl))

        ''txtMaxRng initial props
        'txtMaxRng = New UITextBox(oUILib)
        'With txtMaxRng
        '    .ControlName = "txtMaxRng"
        '    .Left = 290 '170
        '    .Top = 150
        '    .Width = 41
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "0"
        '    '.ForeColor = muSettings.InterfaceBorderColor
        '    .ForeColor = muSettings.InterfaceTextBoxForeColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        '    '.BackColorEnabled = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '    .MaxLength = 3
        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .ToolTipText = "Indicates long-range scanning capability. Maximum radar range units are" & vbCrLf & "four times the size of optimum radar range units." & vbCrLf & "Valid range of values is between 0 and 255."
        '    .bNumericOnly = True
        'End With
        'Me.AddChild(CType(txtMaxRng, UIControl))
        tpMaxRng = New ctlTechProp(oUILib)
        With tpMaxRng
            .ControlName = "tpMaxRng"
            .Enabled = True
            .Height = 18
            .Left = 10
            .MaxValue = 255
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Detection Range:"
            .PropertyValue = 0
            .Top = 155
            .Visible = True
            '.Width = 512
            .ToolTipText = "Indicates long-range scanning capability. Detection radar range units are" & vbCrLf & "four times the size of Visible radar range units." & vbCrLf & "Valid range of values is between 0 and 255."
        End With
        Me.AddChild(CType(tpMaxRng, UIControl))

        ''txtDisRes initial props
        'txtDisRes = New UITextBox(oUILib)
        'With txtDisRes
        '    .ControlName = "txtDisRes"
        '    .Left = 170
        '    .Top = 175
        '    .Width = 41
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "0"
        '    '.ForeColor = muSettings.InterfaceBorderColor
        '    .ForeColor = muSettings.InterfaceTextBoxForeColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        '    '.BackColorEnabled = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '    .MaxLength = 3
        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .ToolTipText = "How resilient this radar system is to be against enemy jamming." & vbCrLf & "Valid value ranges are between 0 and 255."
        '    .bNumericOnly = True
        'End With
        'Me.AddChild(CType(txtDisRes, UIControl))
        tpDisRes = New ctlTechProp(oUILib)
        With tpDisRes
            .ControlName = "tpDisRes"
            .Enabled = True
            .Height = 18
            .Left = 10
            .MaxValue = 255
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Disruption Resist:"
            .PropertyValue = 0
            .Top = 175
            .Visible = True
            '.Width = 512
            .ToolTipText = "How resilient this radar system is to be against enemy jamming." & vbCrLf & "Valid value ranges are between 0 and 255."
        End With
        Me.AddChild(CType(tpDisRes, UIControl))

        ''txtJamStrength initial props
        'txtJamStrength = New UITextBox(oUILib)
        'With txtJamStrength
        '    .ControlName = "txtJamStrength"
        '    .Left = 170
        '    .Top = 250
        '    .Width = 41
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "0"
        '    .ForeColor = muSettings.InterfaceTextBoxForeColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '    .MaxLength = 3
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .ToolTipText = "Jamming strength of the radar system against targets." & vbCrLf & "Valid value range is between 0 and 255."
        '    .bNumericOnly = True
        'End With
        'Me.AddChild(CType(txtJamStrength, UIControl))
        tpJamStrength = New ctlTechProp(oUILib)
        With tpJamStrength
            .ControlName = "tpJamStrength"
            .Enabled = True
            .Height = 18
            .Left = 10
            .MaxValue = 255
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Jamming Strength:"
            .PropertyValue = 0
            .Top = 250
            .Visible = True
            '.Width = 512
            .ToolTipText = "Jamming strength of the radar system against targets." & vbCrLf & "Valid value range is between 0 and 255."
        End With
        Me.AddChild(CType(tpJamStrength, UIControl))

        ''txtJamTargets initial props
        'txtJamTargets = New UITextBox(oUILib)
        'With txtJamTargets
        '    .ControlName = "txtJamTargets"
        '    .Left = 170
        '    .Top = 275
        '    .Width = 41
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "0"
        '    .ForeColor = muSettings.InterfaceTextBoxForeColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '    .MaxLength = 3
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .ToolTipText = "Number of targets that can be jammed at the same time." & vbCrLf & "A value of 0 indicates an area of effect jam effecting" & vbCrLf & "units within the Optimum Range of the radar system." & vbCrLf & "Valid value range is between 0 and 255."
        '    .bNumericOnly = True
        'End With
        'Me.AddChild(CType(txtJamTargets, UIControl))
        tpJamTargets = New ctlTechProp(oUILib)
        With tpJamTargets
            .ControlName = "tpJamTargets"
            .Enabled = True
            .Height = 18
            .Left = 10
            .MaxValue = 255
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Jamming Targets:"
            .PropertyValue = 0
            .Top = 275
            .Visible = True
            '.Width = 512
            .ToolTipText = "Number of targets that can be jammed at the same time." & vbCrLf & "A value of 0 indicates an area of effect jam effecting" & vbCrLf & "units within the Optimum Range of the radar system." & vbCrLf & "Valid value range is between 0 and 255."
        End With
        Me.AddChild(CType(tpJamTargets, UIControl))

        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eAreaEffectJamming) <> 0 Then
            tpJamTargets.Width -= 110

            chkAreaEffectJamming = New UICheckBox(oUILib)
            With chkAreaEffectJamming
                .ControlName = "chkAreaEffectJamming"
                .Left = tpJamTargets.Left + tpJamTargets.Width + 15
                .Top = tpJamTargets.Top
                .Width = 90
                .Height = 18
                .Enabled = True
                .Visible = True

                .Caption = "Area Effect"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(6, DrawTextFormat)
                .Value = False
                .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
                .ToolTipText = "If checked, disables the Jamming Targets slider and indicates that the jamming device uses Area Effect Jamming."
            End With
            Me.AddChild(CType(chkAreaEffectJamming, UIControl))
            AddHandler chkAreaEffectJamming.Click, AddressOf chkAreaEffectJamming_Click
        End If

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

        'btnDesign initial props
        btnDesign = New UIButton(oUILib)
        With btnDesign
            .ControlName = "btnDesign"
            .Left = (Me.Width \ 2) - 50
            .Top = 460
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

        'btnCancel initial props
        btnCancel = New UIButton(oUILib)
        With btnCancel
            .ControlName = "btnCancel"
            .Left = Me.Width - 110 '250
            .Top = 460
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

        'btnRename initial props
        btnRename = New UIButton(oUILib)
        With btnRename
            .ControlName = "btnRename"
            .Left = txtTechName.Left + txtTechName.Width + 5
            .Top = txtTechName.Top - 1
            .Width = 100
            .Height = 24 'txtTechName.Height
            .Enabled = True
            .Visible = False
            .Caption = "Rename"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnRename, UIControl))

        'btnDelete initial props
        btnDelete = New UIButton(oUILib)
        With btnDelete
            .ControlName = "btnDelete"
            .Left = 10 'btnDesign.Left - 105
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

        'btnDelete initial props
        btnAutoFill = New UIButton(oUILib)
        With btnAutoFill
            .ControlName = "btnAutoFill"
            .Left = 10
            .Top = btnCancel.Top
            .Width = 100
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Auto Balance"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnAutoFill, UIControl))

        'lblDesignFlaw initial props
        lblDesignFlaw = New UILabel(oUILib)
        With lblDesignFlaw
            .ControlName = "lblDesignFlaw"
            .Left = 15
            .Top = 400
            .Width = Me.Width - (.Left * 2)
            .Height = 56 '18
            .Enabled = True
            .Visible = False
            .Caption = ""
            .ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(lblDesignFlaw, UIControl))

        tpCasing = New ctlTechProp(oUILib)
        With tpCasing
            .ControlName = "tpCasing"
            .Enabled = True
            .Height = 18
            .Left = lblCasing.Left + lblCasing.Width + 180  '10 for left/right and 170 for cbo
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Casing:"
            .PropertyValue = 0
            .Top = lblCasing.Top
            .Visible = True
            .Width = Me.Width - .Left - 15
            .SetNoPropDisplay(True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpCasing, UIControl))

        tpCollection = New ctlTechProp(oUILib)
        With tpCollection
            .ControlName = "tpCollection"
            .Enabled = True
            .Height = 18
            .Left = tpCasing.Left
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Collection:"
            .PropertyValue = 0
            .Top = lblCollection.Top
            .Visible = True
            .Width = tpCasing.Width
            .SetNoPropDisplay(True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpCollection, UIControl))

        tpEmitter = New ctlTechProp(oUILib)
        With tpEmitter
            .ControlName = "tpEmitter"
            .Enabled = True
            .Height = 18
            .Left = tpCasing.Left
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Emitter:"
            .PropertyValue = 0
            .Top = lblEmitter.Top
            .Visible = True
            .Width = tpCasing.Width
            .SetNoPropDisplay(True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpEmitter, UIControl))

        tpDetection = New ctlTechProp(oUILib)
        With tpDetection
            .ControlName = "tpDetection"
            .Enabled = True
            .Height = 18
            .Left = tpCasing.Left
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Collection:"
            .PropertyValue = 0
            .Top = lblDetection.Top
            .Visible = True
            .Width = tpCasing.Width
            .SetNoPropDisplay(True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpDetection, UIControl))

        ''cboCasing initial props
        'cboCasing = New UIComboBox(oUILib)
        'With cboCasing
        '    .ControlName = "cboCasing"
        '    .Left = lblCasing.Left + lblCasing.Width + 5
        '    .Top = 375
        '    .Width = 170
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    '.FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    '.ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    '.HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    '.DropDownListBorderColor = muSettings.InterfaceBorderColor
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .FillColor = muSettings.InterfaceTextBoxFillColor
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
        '    .mbAcceptReprocessEvents = True
        'End With
        'Me.AddChild(CType(cboCasing, UIControl))
        stmCasing = New ctlSetTechMineral(oUILib)
        With stmCasing
            .ControlName = "stmCasing"
            .Left = lblCasing.Left + lblCasing.Width + 5
            .Top = 375
            .Width = 170
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 3
        End With
        Me.AddChild(CType(stmCasing, UIControl))

        ''cboCollection initial props
        'cboCollection = New UIComboBox(oUILib)
        'With cboCollection
        '    .ControlName = "cboCollection"
        '    .Left = lblCollection.Left + lblCollection.Width + 5
        '    .Top = 350
        '    .Width = 170
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    '.FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    '.ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    '.HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    '.DropDownListBorderColor = muSettings.InterfaceBorderColor
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .FillColor = muSettings.InterfaceTextBoxFillColor
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
        '    .mbAcceptReprocessEvents = True
        'End With
        'Me.AddChild(CType(cboCollection, UIControl))
        stmCollection = New ctlSetTechMineral(oUILib)
        With stmCollection
            .ControlName = "stmCollection"
            .Left = stmCasing.Left
            .Top = 350
            .Width = 170
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 2
        End With
        Me.AddChild(CType(stmCollection, UIControl))

        ''cboDetection initial props
        'cboDetection = New UIComboBox(oUILib)
        'With cboDetection
        '    .ControlName = "cboDetection"
        '    .Left = lblDetection.Left + lblDetection.Width + 5
        '    .Top = 325
        '    .Width = 170
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    '.FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    '.ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    '.HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    '.DropDownListBorderColor = muSettings.InterfaceBorderColor
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .FillColor = muSettings.InterfaceTextBoxFillColor
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
        '    .mbAcceptReprocessEvents = True
        'End With
        'Me.AddChild(CType(cboDetection, UIControl))
        stmDetection = New ctlSetTechMineral(oUILib)
        With stmDetection
            .ControlName = "stmDetection"
            .Left = stmCasing.Left
            .Top = 325
            .Width = 170
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 1
        End With
        Me.AddChild(CType(stmDetection, UIControl))

        ''cboEmitter initial props
        'cboEmitter = New UIComboBox(oUILib)
        'With cboEmitter
        '    .ControlName = "cboEmitter"
        '    .Left = lblEmitter.Left + lblEmitter.Width + 5
        '    .Top = 300
        '    .Width = 170
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    '.FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    '.ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    '.HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    '.DropDownListBorderColor = muSettings.InterfaceBorderColor
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .FillColor = muSettings.InterfaceTextBoxFillColor
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
        '    .mbAcceptReprocessEvents = True
        'End With
        'Me.AddChild(CType(cboEmitter, UIControl))
        stmEmitter = New ctlSetTechMineral(oUILib)
        With stmEmitter
            .ControlName = "stmEmitter"
            .Left = stmCasing.Left
            .Top = 300
            .Width = 170
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 0
        End With
        Me.AddChild(CType(stmEmitter, UIControl))

        'cboJamEffect initial props
        cboJamEffect = New UIComboBox(oUILib)
        With cboJamEffect
            .ControlName = "cboJamEffect"
            .Left = 170
            .Top = 225
            .Width = 170
            .Height = 18
            .Enabled = True
            .Visible = True
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
            .ToolTipText = "The type of effect that this radar has when it jams a target." & vbCrLf & "By selecting anything other than No Effect indicates jamming capabilities."
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(cboJamEffect, UIControl))

        'cboJamImmune initial props
        cboJamImmune = New UIComboBox(oUILib)
        With cboJamImmune
            .ControlName = "cboJamImmune"
            .Left = 170
            .Top = 200
            .Width = 170
            .Height = 18
            .Enabled = True
            .Visible = True
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
            .ToolTipText = "Type of Jamming effect that will never affect this radar system."
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(cboJamImmune, UIControl))

        'cboHullType initial props
        cboHullType = New UIComboBox(oUILib)
        With cboHullType
            .ControlName = "cboHullType"
            .Left = cboJamEffect.Left
            .Top = lblRadarType.Top
            .Width = 170
            .Height = 18
            .Enabled = True
            .Visible = True 
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0)) 
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(cboHullType, UIControl))

        'lblRadarBuilder initial props
        lblRadarBuilder = New UILabel(oUILib)
        With lblRadarBuilder
            .ControlName = "lblRadarBuilder"
            .Left = 10
            .Top = 10
            .Width = 172
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Radar Builder"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 18, FontStyle.Bold Or FontStyle.Italic, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblRadarBuilder, UIControl))
        FillValues()
        EnableDisableControls()
        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
        ofrm.ShowMineralDetail(Me.Left + Me.Width + 5, Me.Top, Me.Height, -1)

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        glCurrentEnvirView = CurrentView.eRadarResearch

        'AddHandler cboCasing.DropDownExpanded, AddressOf DropDownExpanded_Event
        'AddHandler cboCollection.DropDownExpanded, AddressOf DropDownExpanded_Event
        'AddHandler cboEmitter.DropDownExpanded, AddressOf DropDownExpanded_Event
        'AddHandler cboDetection.DropDownExpanded, AddressOf DropDownExpanded_Event

        'AddHandler cboCasing.ItemSelected, AddressOf ComboSelectedValue
        'AddHandler cboCollection.ItemSelected, AddressOf ComboSelectedValue
        'AddHandler cboEmitter.ItemSelected, AddressOf ComboSelectedValue
        'AddHandler cboDetection.ItemSelected, AddressOf ComboSelectedValue

        AddHandler stmCasing.SetButtonClicked, AddressOf SetButtonClicked
        AddHandler stmCollection.SetButtonClicked, AddressOf SetButtonClicked
        AddHandler stmDetection.SetButtonClicked, AddressOf SetButtonClicked
        AddHandler stmEmitter.SetButtonClicked, AddressOf SetButtonClicked

        AddHandler cboHullType.ItemSelected, AddressOf HullTypeSelected

        AddHandler tpCasing.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpCollection.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpDetection.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpDisRes.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpEmitter.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpJamStrength.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpJamTargets.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpMaxRng.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpOptRng.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpScanRes.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpWepAcc.PropertyValueChanged, AddressOf tp_ValueChanged

        AddHandler cboJamImmune.ItemSelected, AddressOf ComboSelectedValue
        AddHandler cboJamEffect.ItemSelected, AddressOf ComboSelectedValue


        If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
        goFullScreenBackground = Nothing
        goFullScreenBackground = goResMgr.GetTexture("radarbuilder.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "TechBacks.pak")

        SetBuildPhase()

        mbIgnoreOptionEvents = False
        'optPassive.Value = False
        'optActive.Value = True

    End Sub

    Public Sub FillValues()
        'Dim lSorted() As Int32 = GetSortedMineralIdxArray(False)

        ''Fill our minerals/alloys
        'cboCasing.Clear() : cboCollection.Clear() : cboDetection.Clear() : cboEmitter.Clear()
        'If lSorted Is Nothing = False Then
        '    For X As Int32 = 0 To lSorted.GetUpperBound(0)
        '        If lSorted(X) <> -1 Then
        '            'If goMinerals(lSorted(X)).IsFullyStudied = False Then Continue For
        '            cboCasing.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboCasing.ItemData(cboCasing.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '            cboCollection.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboCollection.ItemData(cboCollection.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '            cboDetection.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboDetection.ItemData(cboDetection.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '            cboEmitter.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboEmitter.ItemData(cboEmitter.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '        End If
        '    Next X
        'End If

        Dim lJamImmunes As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eJammingImmunityAvailable)
        With cboJamImmune
            .Clear()
            .AddItem("No Effect") : .ItemData(.NewIndex) = 0
            If lJamImmunes > 0 Then
                .AddItem("System Degradation") : .ItemData(.NewIndex) = 1
            End If
            If lJamImmunes > 1 Then
                .AddItem("System Interference") : .ItemData(.NewIndex) = 2
            End If
            If lJamImmunes > 2 Then
                .AddItem("System Clutter") : .ItemData(.NewIndex) = 3
            End If
            If lJamImmunes > 3 Then
                .AddItem("Anti-Jamming") : .ItemData(.NewIndex) = 4
            End If
            If lJamImmunes > 4 Then
                .AddItem("Decoy Clutter") : .ItemData(.NewIndex) = 5
            End If
            .ListIndex = 0
        End With
        Dim lJamEffects As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eJammingEffectAvailable)
        With cboJamEffect
            .Clear()
            .AddItem("No Effect") : .ItemData(.NewIndex) = 0
            If lJamEffects > 0 Then
                .AddItem("System Degradation") : .ItemData(.NewIndex) = 1
            End If
            If lJamEffects > 1 Then
                .AddItem("System Interference") : .ItemData(.NewIndex) = 2
            End If
            If lJamEffects > 2 Then
                .AddItem("System Clutter") : .ItemData(.NewIndex) = 3
            End If
            If lJamEffects > 3 Then
                .AddItem("Anti-Jamming") : .ItemData(.NewIndex) = 4
            End If
            If lJamEffects > 4 Then
                .AddItem("Decoy Clutter") : .ItemData(.NewIndex) = 5
            End If
            .ListIndex = 0
        End With


        SetCboHelpers()

        Dim lMaxPowerThrust As Int32 = 10000
        If goCurrentPlayer Is Nothing = False Then lMaxPowerThrust = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eEnginePowerThrustLimit)
        cboHullType.Clear()
        If lMaxPowerThrust > 110000 Then
            cboHullType.AddItem("Battlecruiser") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.BattleCruiser
        End If
        If lMaxPowerThrust > 400000 Then
            cboHullType.AddItem("Battleship") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Battleship
        End If
        cboHullType.AddItem("Corvette") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Corvette
        If lMaxPowerThrust > 57000 Then
            cboHullType.AddItem("Cruiser") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Cruiser
        End If
        If lMaxPowerThrust > 32000 Then
            cboHullType.AddItem("Destroyer") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Destroyer
        End If
        cboHullType.AddItem("Escort") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Escort
        cboHullType.AddItem("Facility") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Facility
        cboHullType.AddItem("Fighter (Light)") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.LightFighter
        cboHullType.AddItem("Fighter (Medium)") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.MediumFighter
        cboHullType.AddItem("Fighter (Heavy)") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.HeavyFighter
        cboHullType.AddItem("Frigate") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Frigate
        Dim bHasNavalUnit As Boolean = False
        If goCurrentPlayer Is Nothing = False Then bHasNavalUnit = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eNavalAbility) > 0
        If bHasNavalUnit = True Then
            If lMaxPowerThrust > 181000 Then
                cboHullType.AddItem("Naval Battleship") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalBattleship
                cboHullType.AddItem("Naval Carrier") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalCarrier
            End If
            If lMaxPowerThrust > 83000 Then
                cboHullType.AddItem("Naval Cruiser") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalCruiser
            End If
            If lMaxPowerThrust > 31000 Then
                cboHullType.AddItem("Naval Destroyer") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalDestroyer
            End If
            If lMaxPowerThrust > 15000 Then
                cboHullType.AddItem("Naval Frigate") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalFrigate
                If lMaxPowerThrust > 42000 Then
                    cboHullType.AddItem("Naval Submarine") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalSub
                End If
            End If
            cboHullType.AddItem("Naval Utility") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Utility
        End If
        cboHullType.AddItem("Small Vehicle") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.SmallVehicle
        cboHullType.AddItem("Space Station") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.SpaceStation
        cboHullType.AddItem("Tank") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Tank
        cboHullType.AddItem("Utility") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Utility
    End Sub

    Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        ReturnToPreviousViewAndReleaseBackground()
    End Sub

    Private Sub ReturnToPreviousViewAndReleaseBackground()
        MyBase.moUILib.RemoveWindow("frmMinDetail")
        If mfrmProdCost Is Nothing = False Then MyBase.moUILib.RemoveWindow(mfrmProdCost.ControlName)
        If mfrmResCost Is Nothing = False Then MyBase.moUILib.RemoveWindow(mfrmResCost.ControlName)
        ReturnToPreviousView()
        If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
        goFullScreenBackground = Nothing
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

        'we drop focus for our controls just incase some invalid value causes recalculating
        If goUILib.FocusedControl Is Nothing = False Then
            goUILib.FocusedControl.HasFocus = False
            goUILib.FocusedControl = Nothing
        End If

        'Check the Techname
        If txtTechName.Caption.Trim.Length = 0 Then
            MyBase.moUILib.AddNotification("You must specify a name for this Radar.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        If mbImpossibleDesign = True Then
            MyBase.moUILib.AddNotification("You must fix the flaws of this design.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
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
            MyBase.moUILib.RemoveWindow(Me.ControlName)
            ReturnToPreviousViewAndReleaseBackground()
            Return
        End If

        'Dim lMinID(3) As Int32

        'If cboEmitter.ListIndex > -1 Then lMinID(0) = cboEmitter.ItemData(cboEmitter.ListIndex)
        'If cboDetection.ListIndex > -1 Then lMinID(1) = cboDetection.ItemData(cboDetection.ListIndex)
        'If cboCollection.ListIndex > -1 Then lMinID(2) = cboCollection.ItemData(cboCollection.ListIndex)
        'If cboCasing.ListIndex > -1 Then lMinID(3) = cboCasing.ItemData(cboCasing.ListIndex)

        For X As Int32 = 0 To mlMineralIDs.Length - 1
            If mlMineralIDs(X) < 1 Then
                MyBase.moUILib.AddNotification("You must select a material for all entries.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            Else
                For Y As Int32 = 0 To glMineralUB
                    If glMineralIdx(Y) = mlMineralIDs(X) Then
                        If goMinerals(Y).bDiscovered = False Then
                            MyBase.moUILib.AddNotification("You must select minerals that you have researched.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return
                        End If
                        Exit For
                    End If
                Next Y
            End If
        Next X

        'If IsNumeric(txtWepAcc.Caption) = False OrElse IsNumeric(txtScanRes.Caption) = False OrElse _
        '  IsNumeric(txtOptRng.Caption) = False OrElse IsNumeric(txtDisRes.Caption) = False OrElse IsNumeric(txtMaxRng.Caption) = False OrElse _
        '  IsNumeric(txtJamStrength.Caption) = False OrElse IsNumeric(txtJamTargets.Caption) = False Then
        '    MyBase.moUILib.AddNotification("Entries must be numeric!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        '    Return
        'End If

        'Dim sTemp As String = txtWepAcc.Caption & txtScanRes.Caption & txtOptRng.Caption & txtDisRes.Caption & txtMaxRng.Caption & txtJamStrength.Caption & txtJamTargets.Caption

        'If sTemp.IndexOf("."c) > -1 Then
        '    MyBase.moUILib.AddNotification("Numeric entries must be whole numbers!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        '    Return
        'ElseIf sTemp.IndexOf("-"c) > -1 Then
        '    MyBase.moUILib.AddNotification("Numeric values must be at least 0!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        '    Return
        'End If

        Dim lWepAcc As Int32 = CInt(tpWepAcc.PropertyValue)
        Dim lScanRes As Int32 = CInt(tpScanRes.PropertyValue)
        Dim lOptRng As Int32 = CInt(tpOptRng.PropertyValue)
        Dim lDisRes As Int32 = CInt(tpDisRes.PropertyValue)
        Dim lMaxRng As Int32 = CInt(tpMaxRng.PropertyValue)

        'If optActive.Value = False AndAlso optPassive.Value = False Then
        '    MyBase.moUILib.AddNotification("You must select whether this is an active or passive system!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        '    Return
        'End If
        'Dim yType As Byte = 0
        'If optActive.Value = True Then yType = 1
        If cboHullType.ListIndex < 0 Then
            MyBase.moUILib.AddNotification("You must select the hull type for this radar system.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If
        Dim yType As Byte = CByte(cboHullType.ItemData(cboHullType.ListIndex))

        Dim lJamEffect As Int32 = 0
        If cboJamEffect.ListIndex <> -1 Then lJamEffect = cboJamEffect.ItemData(cboJamEffect.ListIndex)
        Dim lJamImmunity As Int32 = 0
        If cboJamImmune.ListIndex <> -1 Then lJamImmunity = cboJamImmune.ItemData(cboJamImmune.ListIndex)
        Dim lJamStr As Int32 = CInt(tpJamStrength.PropertyValue)
        If lJamStr < 0 OrElse lJamStr > mlMaxJamStr Then
            MyBase.moUILib.AddNotification("Jamming Strength must be between 0 and " & mlMaxJamStr & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If
        Dim lJamTargets As Int32 = CInt(tpJamTargets.PropertyValue)

        If chkAreaEffectJamming Is Nothing = False AndAlso chkAreaEffectJamming.Value = True Then
            lJamTargets = 128
        Else
            If lJamEffect <> 0 AndAlso lJamStr <> 0 Then
                Dim lMinJamTargets As Int32 = 1
                'If goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eAreaEffectJamming) <> 0 Then lMinJamTargets = -1
                If lJamTargets < lMinJamTargets OrElse lJamTargets > mlMaxJamTargets Then
                    MyBase.moUILib.AddNotification("Jamming Targets must be between " & lMinJamTargets & " and " & mlMaxJamTargets & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If
            End If
            If lJamTargets = -1 Then lJamTargets = 128
        End If
        

        If lWepAcc > mlMaxWpnAcc Then
            MyBase.moUILib.AddNotification("Weapon Accuracy cannot exceed " & mlMaxWpnAcc & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lScanRes > mlMaxScanRes Then
            MyBase.moUILib.AddNotification("Scan Resolution cannot exceed " & mlMaxScanRes & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lOptRng > mlMaxOptRange Then
            MyBase.moUILib.AddNotification("Optimum Range cannot exceed " & mlMaxOptRange & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lDisRes > mlMaxDisRes Then
            MyBase.moUILib.AddNotification("Disruption Resist cannot exceed " & mlMaxDisRes & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lMaxRng > mlMaxMaxRange Then
            MyBase.moUILib.AddNotification("Maximum Range cannot exceed " & mlMaxMaxRange & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

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

        Dim lTechID As Int32 = -1
        If moTech Is Nothing = False Then lTechID = moTech.ObjectID

        Dim lHullReq As Int32 = -1
        Dim lPowerReq As Int32 = -1
        Dim blResCost As Int64 = -1
        Dim blResTime As Int64 = -1
        Dim blProdCost As Int64 = -1
        Dim blProdTime As Int64 = -1
        Dim lColonists As Int32 = -1
        Dim lEnlisted As Int32 = -1
        Dim lOfficers As Int32 = -1
        mfrmBuilderCost.GetLockedValues(lHullReq, lPowerReq, blResCost, blResTime, blProdCost, blProdTime, lColonists, lEnlisted, lOfficers)

        Dim lSpecMin1 As Int32 = -1
        If tpEmitter.PropertyLocked = True Then lSpecMin1 = CInt(tpEmitter.PropertyValue) Else lSpecMin1 = -1
        Dim lSpecMin2 As Int32 = -1
        If tpDetection.PropertyLocked = True Then lSpecMin2 = CInt(tpDetection.PropertyValue) Else lSpecMin2 = -1
        Dim lSpecMin3 As Int32 = -1
        If tpCollection.PropertyLocked = True Then lSpecMin3 = CInt(tpCollection.PropertyValue) Else lSpecMin3 = -1
        Dim lSpecMin4 As Int32 = -1
        If tpCasing.PropertyLocked = True Then lSpecMin4 = CInt(tpCasing.PropertyValue) Else lSpecMin4 = -1

        MyBase.moUILib.GetMsgSys.SubmitRadarDesign(oResearchGuid, txtTechName.Caption, CByte(lWepAcc), _
          CByte(lScanRes), CByte(lOptRng), CByte(lMaxRng), CByte(lDisRes), mlMineralIDs(0), mlMineralIDs(1), mlMineralIDs(2), mlMineralIDs(3), _
          yType, CByte(lJamImmunity), CByte(lJamStr), CByte(lJamTargets), CByte(lJamEffect), lTechID, _
          lHullReq, lPowerReq, blResCost, blResTime, blProdCost, blProdTime, lColonists, lEnlisted, lOfficers, _
          lSpecMin1, lSpecMin2, lSpecMin3, lSpecMin4, -1, -1)

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eSubmitDesignClick, ObjectType.eRadarTech, -1, -1, "")
        End If

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddNotification("Radar Prototype Design Submitted", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

        ReturnToPreviousViewAndReleaseBackground()
    End Sub

    'Private Sub cboCasing_ItemSelected(ByVal lItemIndex As Integer) Handles cboCasing.ItemSelected
    '    If cboCasing.ListIndex > -1 Then
    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboCasing.ItemData(cboCasing.ListIndex))
    '    End If
    'End Sub

    'Private Sub cboCollection_ItemSelected(ByVal lItemIndex As Integer) Handles cboCollection.ItemSelected
    '    If cboCollection.ListIndex > -1 Then
    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboCollection.ItemData(cboCollection.ListIndex))
    '    End If
    'End Sub

    'Private Sub cboDetection_ItemSelected(ByVal lItemIndex As Integer) Handles cboDetection.ItemSelected
    '    If cboDetection.ListIndex > -1 Then
    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboDetection.ItemData(cboDetection.ListIndex))
    '    End If
    'End Sub

    'Private Sub cboEmitter_ItemSelected(ByVal lItemIndex As Integer) Handles cboEmitter.ItemSelected
    '    If cboEmitter.ListIndex > -1 Then
    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboEmitter.ItemData(cboEmitter.ListIndex))
    '    End If
    'End Sub

    Private Sub SetCboHelpers()
        'Dim bReflect As Boolean = goCurrentPlayer.PlayerKnowsProperty("REFLECTION")
        'Dim bElectRes As Boolean = goCurrentPlayer.PlayerKnowsProperty("ELECTRICAL RESISTANCE")
        'Dim bDensity As Boolean = goCurrentPlayer.PlayerKnowsProperty("DENSITY")
        'Dim bSuperC As Boolean = goCurrentPlayer.PlayerKnowsProperty("SUPERCONDUCTIVE POINT")
        'Dim bMalleable As Boolean = goCurrentPlayer.PlayerKnowsProperty("MALLEABLE")
        'Dim bCompress As Boolean = goCurrentPlayer.PlayerKnowsProperty("COMPRESSIBILITY")
        'Dim bThermCond As Boolean = goCurrentPlayer.PlayerKnowsProperty("THERMAL CONDUCTION")
        'Dim bHardness As Boolean = goCurrentPlayer.PlayerKnowsProperty("HARDNESS")

        Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder

        ''Radar.Emitter		Reflection high, Electrical Resist low, density high
        oSB.Length = 0
        'If bReflect = True AndAlso bElectRes = True AndAlso bDensity = True Then
        'oSB.AppendLine("Dense materials with low electrical resistance")
        'oSB.AppendLine("and high reflection properties are important.")
        'ElseIf bReflect = True AndAlso bDensity = True Then
        '    oSB.AppendLine("Dense, highly reflective materials are preferred.")
        'ElseIf bElectRes = True AndAlso bDensity = True Then
        '    oSB.AppendLine("Dense materials with low electrical resistance are important.")
        'ElseIf bReflect = True AndAlso bElectRes = True Then
        '    oSB.AppendLine("Reflective materials with low electrical resistance are preferred.")
        'Else
        '    If bReflect = True Then oSB.AppendLine("Highly Reflective materials are preferred.")
        '    If bElectRes = True Then oSB.AppendLine("Materials with a low Electrical resistance are preferred.")
        '    If bDensity = True Then oSB.AppendLine("High Density materials are important.")
        'End If
        oSB.AppendLine("Low Electrical Resistance properties with high Magnetic Production and Magnetic Reaction properties are a must.")
        If oSB.Length <> 0 Then stmEmitter.ToolTipText = oSB.ToString Else stmEmitter.ToolTipText = "We are not quite sure of the best way to engineer this."

        ''Radar.Detection		electrical resist low, superconductive point low
        oSB.Length = 0
        'If bElectRes = True AndAlso bSuperC = True Then
        'oSB.AppendLine("Materials with a low electrical resistance and low superconductive point work best.")
        'ElseIf bElectRes = True Then
        '    oSB.AppendLine("Low Electrical Resistance properties are best.")
        'ElseIf bSuperC = True Then
        '    oSB.AppendLine("Low Superconductive Point is important for this material.")
        'End If
        oSB.AppendLine("Materials with a high Superconductive Point with a high Magnetic Reactance and a low Magnetic Production.")
        If oSB.Length <> 0 Then stmDetection.ToolTipText = oSB.ToString Else stmDetection.ToolTipText = "We are not quite sure of the best way to engineer this."

        ''Radar.Collection	malleable high, reflection high, compression high
        oSB.Length = 0
        'If bMalleable = True AndAlso bReflect = True AndAlso bCompress = True Then
        'oSB.AppendLine("Highly Compressible materials that are malleable and")
        'oSB.AppendLine("have high Reflective properties are a must.")
        'Else
        '    If bMalleable = True Then oSB.AppendLine("Higher Malleable properties are important.")
        '    If bReflect = True Then oSB.AppendLine("Greater Reflective properties are important.")
        '    If bCompress = True Then oSB.AppendLine("High Compressibility is preferred.")
        'End If
        oSB.AppendLine("Highly Malleable materials with low Electrical Resistance and low Magnetic Production")
        If oSB.Length <> 0 Then stmCollection.ToolTipText = oSB.ToString Else stmCollection.ToolTipText = "We are not quite sure of the best way to engineer this."

        ''Radar.Casing		thermal conduct high, hardness high, density high
        oSB.Length = 0
        'If bThermCond = True AndAlso bHardness = True AndAlso bDensity = True Then
        'oSB.AppendLine("Dense materials that are very hard")
        'oSB.AppendLine("and conduct heat are important.")
        'Else
        '    If bThermCond = True Then oSB.AppendLine("High Thermal Conductance is important.")
        '    If bHardness = True Then oSB.AppendLine("Higher Hardness attributes are preferred.")
        '    If bDensity = True Then oSB.AppendLine("Dense materials work best.")
        'End If
        oSB.AppendLine("Very Dense materials with high Thermal Conductance and low Thermal Expansion, Magnetic Production and Magnetic Reactance properties work best.")
        If oSB.Length <> 0 Then stmCasing.ToolTipText = oSB.ToString Else stmCasing.ToolTipText = "We are not quite sure of the best way to engineer this."

        'Also related to special techs...
        mlMaxOptRange = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eRadarOptRange)
        mlMaxMaxRange = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eRadarMaxRange)
        mlMaxScanRes = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eRadarScanRes)
        mlMaxWpnAcc = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eRadarWpnAcc)
        mlMaxDisRes = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eRadarDisRes)
        mlMaxJamTargets = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eJammingTargets)
        If (mlMaxJamTargets And 128) <> 0 Then mlMaxJamTargets = mlMaxJamTargets Xor 128
        mlMaxJamStr = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eJammingStrength)
        tpWepAcc.ToolTipText = "Determines the overall ability of the radar system" & vbCrLf & "to execute firing sequences and assist gunners." & vbCrLf & "Valid Range is any value between 0 and " & mlMaxWpnAcc & "."
        tpScanRes.ToolTipText = "Determines scanning strength of the radar system." & vbCrLf & "Higher values will reward more detailed results in target scans" & vbCrLf & "and in long range (outside optimum, inside maximum) scans." & vbCrLf & "Valid range is any value between 0 and " & mlMaxScanRes & "."
        tpOptRng.ToolTipText = "Indicates short-range visibility. Units will attack enemies within" & vbCrLf & "this range. Entities within this range will be fully visible." & vbCrLf & "Valid range is any value between 0 and " & mlMaxOptRange & "."
        tpMaxRng.ToolTipText = "Indicates long-range scanning capability. One unit of Maximum Radar Range" & vbCrLf & "is equivalent in size to four units of Optimum Radar Range." & vbCrLf & "Valid range of values is between 0 and " & mlMaxMaxRange & "."
        tpDisRes.ToolTipText = "How resilient this radar system is to be against enemy jamming." & vbCrLf & "Valid value ranges are between 0 and " & mlMaxDisRes & "."

        Dim sFinalTT As String = "Number of targets that can be jammed at the same time." & vbCrLf
        'If goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eAreaEffectJamming) <> 0 Then
        '    sFinalTT &= "A value of -1 indicates an area of effect jam effecting" & vbCrLf & "units within the Optimum Range of the radar system." & vbCrLf & "Valid value range is between -1 and "
        '    tpJamTargets.MinValue = -1
        'Else
        sFinalTT &= "A value greater than 0 must be entered for jamming effects." & vbCrLf & "Valid value range is between 1 and "
        tpJamTargets.MinValue = 1
        'End If
        tpJamTargets.ToolTipText = sFinalTT & mlMaxJamTargets & "."
        tpJamStrength.ToolTipText = "Jamming strength of the radar system against targets." & vbCrLf & "Valid value range is between 0 and " & mlMaxJamStr & "."

        tpOptRng.MaxValue = mlMaxOptRange
        tpMaxRng.MaxValue = mlMaxMaxRange
        tpScanRes.MaxValue = mlMaxScanRes
        tpWepAcc.MaxValue = mlMaxWpnAcc
        tpDisRes.MaxValue = mlMaxDisRes
        tpJamTargets.MaxValue = mlMaxJamTargets
        tpJamStrength.MaxValue = mlMaxJamStr

        tpDisRes.SetToInitialDefault()
        tpMaxRng.SetToInitialDefault()
        tpOptRng.SetToInitialDefault()
        tpScanRes.SetToInitialDefault()
        tpWepAcc.SetToInitialDefault()

    End Sub

    'Private Sub optActive_Click() Handles optActive.Click
    '    If mbIgnoreOptionEvents = True Then Return
    '    mbIgnoreOptionEvents = True
    '    optActive.Value = True
    '    optPassive.Value = False
    '    tpScanRes.Enabled = True
    '    tpMaxRng.Enabled = True
    '    cboJamEffect.Enabled = True
    '    cboJamImmune.Enabled = True
    '    tpJamStrength.Enabled = True
    '    tpJamTargets.Enabled = True
    '    mbIgnoreOptionEvents = False
    'End Sub

    'Private Sub optPassive_Click() Handles optPassive.Click
    '    If mbIgnoreOptionEvents = True Then Return
    '    mbIgnoreOptionEvents = True
    '    optActive.Value = False
    '    optPassive.Value = True
    '    tpScanRes.PropertyValue = 0 : tpScanRes.Enabled = False
    '    tpMaxRng.PropertyValue = 0 : tpMaxRng.Enabled = False
    '    cboJamEffect.ListIndex = -1 : cboJamEffect.Enabled = False
    '    cboJamImmune.ListIndex = -1 : cboJamImmune.Enabled = False
    '    tpJamStrength.PropertyValue = 0 : tpJamStrength.Enabled = False
    '    tpJamTargets.PropertyValue = 0 : tpJamTargets.Enabled = False
    '    mbIgnoreOptionEvents = False
    'End Sub

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

            MyBase.moUILib.RemoveWindow(Me.ControlName)
            ReturnToPreviousViewAndReleaseBackground()
            ''Delete the design
            'Dim yMsg(7) As Byte
            'System.BitConverter.GetBytes(GlobalMessageCode.eDeleteDesign).CopyTo(yMsg, 0)
            'moTech.GetGUIDAsString.CopyTo(yMsg, 2)
            'MyBase.moUILib.SendMsgToPrimary(yMsg)
            'MyBase.moUILib.RemoveWindow(Me.ControlName)
            'ReturnToPreviousViewAndReleaseBackground()
        End If
    End Sub

    Private Sub frmRadarBuilder_OnRender() Handles Me.OnRender
        'Ok, first, is moTech nothing?
        If moTech Is Nothing = False Then
            'Ok, now, do the values used on the form currently match the tech?
            Dim bChanged As Boolean = False
            With moTech
                'bChanged = (.yRadarType = 0 AndAlso optPassive.Value = False) OrElse _
                '   tpWepAcc.PropertyValue <> .WeaponAcc OrElse tpScanRes.PropertyValue <> .ScanResolution OrElse _
                '   tpOptRng.PropertyValue <> .OptimumRange OrElse tpMaxRng.PropertyValue <> .MaximumRange OrElse _
                '   tpDisRes.PropertyValue <> .DisruptionResistance OrElse tpJamStrength.PropertyValue <> .JamStrength OrElse _
                '   tpJamTargets.PropertyValue <> .JamTargets
                bChanged = tpWepAcc.PropertyValue <> .WeaponAcc OrElse tpScanRes.PropertyValue <> .ScanResolution OrElse _
                  tpOptRng.PropertyValue <> .OptimumRange OrElse tpMaxRng.PropertyValue <> .MaximumRange OrElse _
                  tpDisRes.PropertyValue <> .DisruptionResistance OrElse tpJamStrength.PropertyValue <> .JamStrength OrElse _
                  tpJamTargets.PropertyValue <> .JamTargets

                If bChanged = False Then
                    Dim lJamImmune As Int32 = -1
                    Dim lJamEffect As Int32 = -1
                    Dim lCasingID As Int32 = mlMineralIDs(3)
                    Dim lCollID As Int32 = mlMineralIDs(2)
                    Dim lDetID As Int32 = mlMineralIDs(1)
                    Dim lEmitID As Int32 = mlMineralIDs(0)

                    If cboJamImmune.ListIndex <> -1 Then lJamImmune = cboJamImmune.ItemData(cboJamImmune.ListIndex)
                    If cboJamEffect.ListIndex <> -1 Then lJamEffect = cboJamEffect.ItemData(cboJamEffect.ListIndex)
                    'If cboCasing.ListIndex <> -1 Then lCasingID = cboCasing.ItemData(cboCasing.ListIndex)
                    'If cboCollection.ListIndex <> -1 Then lCollID = cboCollection.ItemData(cboCollection.ListIndex)
                    'If cboDetection.ListIndex <> -1 Then lDetID = cboDetection.ItemData(cboDetection.ListIndex)
                    'If cboEmitter.ListIndex <> -1 Then lEmitID = cboEmitter.ItemData(cboEmitter.ListIndex)

                    bChanged = .JamImmunity <> lJamImmune OrElse .JamEffect <> lJamEffect OrElse .lCasingMineralID <> lCasingID OrElse _
                      .lCollectionMineralID <> lCollID OrElse .lDetectionMineralID <> lDetID OrElse .lEmitterMineralID <> lEmitID OrElse _
                      (tpEmitter.PropertyLocked <> (.lSpecifiedMin1 <> -1)) OrElse (tpCollection.PropertyLocked <> (.lSpecifiedMin2 <> -1)) OrElse _
                      (tpCollection.PropertyLocked <> (.lSpecifiedMin3 <> -1)) OrElse (tpCasing.PropertyLocked <> (.lSpecifiedMin4 <> -1))

                    If bChanged = False Then
                        If txtTechName.Caption <> .sRadarName AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eInvalidDesign AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
                            bChanged = True
                        End If
                    End If

                    If bChanged = False Then

                        bChanged = (.lSpecifiedMin1 <> -1 AndAlso .lSpecifiedMin1 <> tpEmitter.PropertyValue) OrElse _
                                    (.lSpecifiedMin2 <> -1 AndAlso .lSpecifiedMin2 <> tpCollection.PropertyValue) OrElse _
                                    (.lSpecifiedMin3 <> -1 AndAlso .lSpecifiedMin3 <> tpCollection.PropertyValue) OrElse _
                                    (.lSpecifiedMin4 <> -1 AndAlso .lSpecifiedMin4 <> tpCasing.PropertyValue) 
                        If bChanged = False Then
                            Dim lHullReq As Int32 = -1
                            Dim lPowerReq As Int32 = -1
                            Dim blResCost As Int64 = -1
                            Dim blResTime As Int64 = -1
                            Dim blProdCost As Int64 = -1
                            Dim blProdTime As Int64 = -1
                            Dim lColonists As Int32 = -1
                            Dim lEnlisted As Int32 = -1
                            Dim lOfficer As Int32 = -1
                            Me.mfrmBuilderCost.GetLockedValues(lHullReq, lPowerReq, blResCost, blResTime, blProdCost, blProdTime, lColonists, lEnlisted, lOfficer)

                            bChanged = lHullReq <> .lSpecifiedHull OrElse lPowerReq <> .lSpecifiedPower OrElse blResCost <> .blSpecifiedResCost OrElse _
                                blResTime <> .blSpecifiedResTime OrElse blProdCost <> .blSpecifiedProdCost OrElse blProdTime <> .blSpecifiedProdTime OrElse _
                                lColonists <> .lSpecifiedColonists OrElse lEnlisted <> .lSpecifiedEnlisted OrElse lOfficer <> .lSpecifiedOfficers
                        End If
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

    Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
        If goTutorial Is Nothing Then goTutorial = New TutorialManager()
        goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eRadarBuilder)
    End Sub

    'Private Sub DropDownExpanded_Event(ByVal sComboBoxName As String)
    '    If sComboBoxName = "" Then Return

    '    Dim oTech As New RadarTechComputer
    '    If cboHullType.ListIndex > -1 Then
    '        oTech.lHullTypeID = cboHullType.ItemData(cboHullType.ListIndex)
    '    End If

    '    Dim lMinIdx As Int32 = -1
    '    Select Case sComboBoxName
    '        Case cboEmitter.ControlName
    '            lMinIdx = 0
    '        Case cboDetection.ControlName
    '            lMinIdx = 1
    '        Case cboCollection.ControlName
    '            lMinIdx = 2
    '        Case cboCasing.ControlName
    '            lMinIdx = 3
    '    End Select
    '    If lMinIdx = -1 Then Return
    '    oTech.MineralCBOExpanded(lMinIdx, ObjectType.eRadarTech)
    'End Sub
    'Try
    '    Dim ofrmMin As frmMinDetail = CType(MyBase.moUILib.GetWindow("frmMinDetail"), frmMinDetail)
    '    If ofrmMin Is Nothing Then Return

    '    Select Case sComboBoxName
    '        Case cboCasing.ControlName
    '            ofrmMin.ClearHighlights()
    '            For X As Int32 = 0 To glMineralPropertyUB
    '                If glMineralPropertyIdx(X) <> -1 Then
    '                    Select Case goMineralProperty(X).MineralPropertyName.ToUpper
    '                        Case "DENSITY"
    '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 4)
    '                        Case "THERMAL CONDUCTION"
    '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 4)
    '                        Case "THERMAL EXPANSION"
    '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
    '                        Case "MAGNETIC PRODUCTION"
    '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
    '                        Case "MAGNETIC REACTANCE"
    '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
    '                    End Select
    '                End If
    '            Next X
    '        Case cboCollection.ControlName
    '            ofrmMin.ClearHighlights()
    '            For X As Int32 = 0 To glMineralPropertyUB
    '                If glMineralPropertyIdx(X) <> -1 Then
    '                    Select Case goMineralProperty(X).MineralPropertyName.ToUpper
    '                        Case "MALLEABLE"
    '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 4)
    '                        Case "ELECTRICAL RESISTANCE"
    '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
    '                        Case "MAGNETIC PRODUCTION"
    '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
    '                    End Select
    '                End If
    '            Next X
    '        Case cboDetection.ControlName
    '            ofrmMin.ClearHighlights()
    '            For X As Int32 = 0 To glMineralPropertyUB
    '                If glMineralPropertyIdx(X) <> -1 Then
    '                    Select Case goMineralProperty(X).MineralPropertyName.ToUpper
    '                        Case "SUPERCONDUCTIVE POINT"
    '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 5)
    '                        Case "MAGNETIC REACTANCE"
    '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
    '                        Case "MAGNETIC PRODUCTION"
    '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 6)
    '                    End Select
    '                End If
    '            Next X
    '        Case cboEmitter.ControlName
    '            ofrmMin.ClearHighlights()
    '            For X As Int32 = 0 To glMineralPropertyUB
    '                If glMineralPropertyIdx(X) <> -1 Then
    '                    Select Case goMineralProperty(X).MineralPropertyName.ToUpper
    '                        Case "ELECTRICAL RESISTANCE"
    '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
    '                        Case "MAGNETIC PRODUCTION"
    '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 5)
    '                        Case "MAGNETIC REACTANCE"
    '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 5)
    '                    End Select
    '                End If
    '            Next X

    '    End Select
    'Catch
    'End Try

    Private Sub tp_ValueChanged(ByVal sPropName As String, ByRef ctl As ctlTechProp)
        BuilderCostValueChange(False)
    End Sub
    Private Sub ComboSelectedValue(ByVal lItemIndex As Int32)
        BuilderCostValueChange(False)
    End Sub


    Private Sub EnableDisableControls()
        If mbIgnoreValueChange = True Then Return
        If cboHullType Is Nothing Then Return

        Dim bValue As Boolean = cboHullType.ListIndex > -1

        SetBuildPhase()
        
        If tpWepAcc.Enabled <> bValue Then tpWepAcc.Enabled = bValue
        If tpScanRes.Enabled <> bValue Then tpScanRes.Enabled = bValue
        If tpOptRng.Enabled <> bValue Then tpOptRng.Enabled = bValue
        If tpMaxRng.Enabled <> bValue Then tpMaxRng.Enabled = bValue
        If tpDisRes.Enabled <> bValue Then tpDisRes.Enabled = bValue
        If tpJamStrength.Enabled <> bValue Then tpJamStrength.Enabled = bValue
        If tpJamTargets.Enabled <> bValue Then tpJamTargets.Enabled = bValue
        If tpCasing.Enabled <> bValue Then tpCasing.Enabled = bValue
        If tpCollection.Enabled <> bValue Then tpCollection.Enabled = bValue
        If tpDetection.Enabled <> bValue Then tpDetection.Enabled = bValue
        If tpEmitter.Enabled <> bValue Then tpEmitter.Enabled = bValue

        If chkAreaEffectJamming Is Nothing = False Then
            If chkAreaEffectJamming.Visible <> bValue Then chkAreaEffectJamming.Visible = bValue
            If chkAreaEffectJamming.Value = True Then tpJamTargets.Enabled = False
        End If
        
    End Sub

    Protected Overrides Sub BuilderCostValueChange(ByVal bAutoFill As Boolean)

        If mbIgnoreValueChange = True Then Return
        EnableDisableControls()
        mbIgnoreValueChange = True
        mbImpossibleDesign = True

        Dim oRadarTechComputer As New RadarTechComputer
        With oRadarTechComputer
            .blDisRes = tpDisRes.PropertyValue
            .blJamStrength = tpJamStrength.PropertyValue

            If chkAreaEffectJamming Is Nothing = False AndAlso chkAreaEffectJamming.Value = True Then .blJamTargets = 20 Else .blJamTargets = tpJamTargets.PropertyValue
            .blMaxRng = tpMaxRng.PropertyValue
            .blOptRng = tpOptRng.PropertyValue
            .blScanRes = tpScanRes.PropertyValue
            .blWepAcc = tpWepAcc.PropertyValue
            If cboHullType.ListIndex > -1 Then
                .lHullTypeID = cboHullType.ItemData(cboHullType.ListIndex)
            End If
            If cboJamEffect.ListIndex > -1 Then
                .lJamType = cboJamEffect.ItemData(cboJamEffect.ListIndex)
            Else : .lJamType = -1
            End If

            .SetMinDAValues(ml_Min1DA, ml_Min2DA, ml_Min3DA, ml_Min4DA, ml_Min5DA, ml_Min6DA)

            'If cboEmitter.ListIndex > -1 Then
            '    .lMineral1ID = cboEmitter.ItemData(cboEmitter.ListIndex)
            'Else : .lMineral1ID = -1
            'End If
            'If cboDetection.ListIndex > -1 Then
            '    .lMineral2ID = cboDetection.ItemData(cboDetection.ListIndex)
            'Else : .lMineral2ID = -1
            'End If
            'If cboCollection.ListIndex > -1 Then
            '    .lMineral3ID = cboCollection.ItemData(cboCollection.ListIndex)
            'Else : .lMineral3ID = -1
            'End If
            'If cboCasing.ListIndex > -1 Then
            '    .lMineral4ID = cboCasing.ItemData(cboCasing.ListIndex)
            'Else : .lMineral4ID = -1
            'End If
            .lMineral1ID = mlMineralIDs(0)
            .lMineral2ID = mlMineralIDs(1)
            .lMineral3ID = mlMineralIDs(2)
            .lMineral4ID = mlMineralIDs(3)
            .lMineral5ID = -1
            .lMineral6ID = -1

            If .lMineral1ID < 1 OrElse .lMineral2ID < 1 OrElse .lMineral3ID < 1 OrElse .lMineral4ID < 1 Then
                lblDesignFlaw.Caption = "All properties and materials need to be defined."
                mbIgnoreValueChange = False
                mbImpossibleDesign = True
                Return
            End If


            .msMin1Name = "Emitter"
            .msMin2Name = "Detection"
            .msMin3Name = "Collection"
            .msMin4Name = "Casing"
            .msMin5Name = ""
            .msMin6Name = ""

            If bAutoFill = True Then
                .DoBuilderCostPreconfigure(lblDesignFlaw, MyBase.mfrmBuilderCost, tpEmitter, tpDetection, tpCollection, tpCasing, Nothing, Nothing, ObjectType.eRadarTech, 0D, False)
            Else
                .BuilderCostValueChange(lblDesignFlaw, MyBase.mfrmBuilderCost, tpEmitter, tpDetection, tpCollection, tpCasing, Nothing, Nothing, ObjectType.eRadarTech, 0D)
            End If
            mbImpossibleDesign = .bImpossibleDesign
        End With



        mbIgnoreValueChange = False


    End Sub

 
    Public Overloads Overrides Sub ViewResults(ByRef oTech As Base_Tech, ByVal lProdFactor As Integer)
        If oTech Is Nothing Then Return

        moTech = CType(oTech, RadarTech)

        Me.btnDesign.Enabled = False

        With moTech

            mbIgnoreValueChange = True
            cboHullType.FindComboItemData(.HullTypeID)

            tpDisRes.PropertyValue = .DisruptionResistance
            If .DisruptionResistance > 0 Then tpDisRes.PropertyLocked = True
            tpOptRng.PropertyValue = .OptimumRange
            If .OptimumRange > 0 Then tpOptRng.PropertyLocked = True
            tpScanRes.PropertyValue = .ScanResolution
            If .ScanResolution > 0 Then tpScanRes.PropertyLocked = True
            Me.txtTechName.Caption = .sRadarName
            tpWepAcc.PropertyValue = .WeaponAcc
            If .WeaponAcc > 0 Then tpWepAcc.PropertyLocked = True
            tpMaxRng.PropertyValue = .MaximumRange
            If .MaximumRange > 0 Then tpMaxRng.PropertyLocked = True

            mlSelectedMineralIdx = 0 : Mineral_Selected(.lEmitterMineralID)
            mlSelectedMineralIdx = 1 : Mineral_Selected(.lDetectionMineralID)
            mlSelectedMineralIdx = 2 : Mineral_Selected(.lCollectionMineralID)
            mlSelectedMineralIdx = 3 : Mineral_Selected(.lCasingMineralID)
            'If .lCasingMineralID > 0 Then Me.cboCasing.FindComboItemData(.lCasingMineralID)
            'If .lCollectionMineralID > 0 Then Me.cboCollection.FindComboItemData(.lCollectionMineralID)
            'If .lDetectionMineralID > 0 Then Me.cboDetection.FindComboItemData(.lDetectionMineralID)
            'If .lEmitterMineralID > 0 Then Me.cboEmitter.FindComboItemData(.lEmitterMineralID)

            tpEmitter.PropertyValue = .lSpecifiedMin1
            If .lSpecifiedMin1 <> -1 Then tpEmitter.PropertyLocked = True
            tpDetection.PropertyValue = .lSpecifiedMin2
            If .lSpecifiedMin2 <> -1 Then tpDetection.PropertyLocked = True
            tpCollection.PropertyValue = .lSpecifiedMin3
            If .lSpecifiedMin3 <> -1 Then tpCollection.PropertyLocked = True
            tpCasing.PropertyValue = .lSpecifiedMin4
            If .lSpecifiedMin4 <> -1 Then tpCasing.PropertyLocked = True

            Me.cboJamEffect.FindComboItemData(.JamEffect)
            Me.cboJamImmune.FindComboItemData(.JamImmunity)
            tpJamStrength.PropertyValue = .JamStrength
            If .JamStrength > 0 Then tpJamStrength.PropertyLocked = True
            If .JamTargets = 128 Then
                tpJamTargets.PropertyValue = 0
                If .JamTargets > 0 Then tpJamTargets.PropertyLocked = True
                If chkAreaEffectJamming Is Nothing = False Then chkAreaEffectJamming.Value = True
            Else
                tpJamTargets.PropertyValue = .JamTargets
                If .JamTargets > 0 Then tpJamTargets.PropertyLocked = True
            End If
            

            'If .yRadarType = 0 Then
            '    optActive.Value = False
            '    optPassive.Value = True
            '    optPassive_Click()
            'Else
            '    optPassive.Value = False
            '    optActive.Value = True
            '    optActive_Click()
            'End If
            'cboHullType.FindComboItemData(.HullTypeID)
            MyBase.mfrmBuilderCost.SetAndLockValues(oTech)

            mbIgnoreValueChange = False
            BuilderCostValueChange(False)

            lblDesignFlaw.Caption = oTech.GetDesignFlawText()
            lblDesignFlaw.Visible = True

        End With

        If oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eComponentResearching AndAlso HasAliasedRights(AliasingRights.eAddResearch) = True Then
            btnDesign.Caption = "Research"
            btnDesign.Enabled = True
        End If
        If gbAliased = False Then
            btnRename.Visible = oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched
            btnDelete.Visible = True ' oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched
            'If btnDelete.Visible = True Then
            '    'Ok, set up the buttons better
            '    btnDesign.Left = (Me.Width \ 2) - (btnDesign.Width \ 2)
            '    btnCancel.Left = Me.Width - btnCancel.Width - 5
            '    btnDelete.Left = 5
            'End If
            btnDelete.Left = 10
            btnCancel.Left = Me.Width - btnCancel.Width - 10
            Dim lNewW As Int32 = (btnCancel.Left - (btnDelete.Left + btnDelete.Width))
            Dim lGapW As Int32 = lNewW - (btnDesign.Width + btnAutoFill.Width)
            lGapW \= 3  'for 3 gaps
            btnDesign.Left = btnCancel.Left - btnDesign.Width - lGapW
            btnAutoFill.Left = btnDelete.Left + btnDelete.Width + lGapW
        End If

        'Now... what state is the tech in?
        If oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
            cboHullType.Enabled = False
        End If
        If oTech.ComponentDevelopmentPhase > Base_Tech.eComponentDevelopmentPhase.eComponentDesign Then
            'Ok, show it's research and production cost
            If mfrmResCost Is Nothing Then mfrmResCost = New frmProdCost(MyBase.moUILib, "frmResCost")
            With mfrmResCost
                .Visible = True
                .Left = Me.Left + Me.Width + 5
                .Top = Me.Top + Me.Height + 8
                .SetFromProdCost(oTech.oResearchCost, lProdFactor, True, 0, 0)
            End With

            If mfrmProdCost Is Nothing Then mfrmProdCost = New frmProdCost(MyBase.moUILib, "frmProdCost")
            With mfrmProdCost
                .Visible = True
                .Left = mfrmResCost.Left + mfrmResCost.Width + 10
                .Top = mfrmResCost.Top
                .SetFromProdCost(oTech.oProductionCost, 1000, False, moTech.HullRequired, moTech.PowerRequired)
            End With
        ElseIf oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
            If mfrmResCost Is Nothing Then mfrmResCost = New frmProdCost(MyBase.moUILib, "frmResCost")
            With mfrmResCost
                .Visible = True
                .Left = Me.Left
                .Top = Me.Top + Me.Height + 8
                .SetFromFailureCode(oTech.ErrorCodeReason)
            End With
        End If
    End Sub

    Private Sub SetBuildPhase()
        'Ok, we do the following steps...
        Dim lPhase As Int32 = 0

        'Hull Type Selected
        If cboHullType.ListIndex > -1 Then lPhase = 1

        'Property Values and mineral selections
        If lPhase = 1 Then
            'because radar does not require any attribute to be non-zero, we can make it all visible.... so determine if the minerals are set
            lPhase = 2
            For X As Int32 = 0 To mlMineralIDs.GetUpperBound(0)
                If mlMineralIDs(X) < 1 Then
                    lPhase = 1
                    Exit For
                End If
            Next X
        End If

        If moTech Is Nothing = False Then lPhase = 2

        Dim bPhase1 As Boolean = lPhase > 0
        Dim bPhase2 As Boolean = lPhase > 1

        If tpWepAcc.Visible <> bPhase1 Then tpWepAcc.Visible = bPhase1
        If tpScanRes.Visible <> bPhase1 Then tpScanRes.Visible = bPhase1
        If tpOptRng.Visible <> bPhase1 Then tpOptRng.Visible = bPhase1
        If tpMaxRng.Visible <> bPhase1 Then tpMaxRng.Visible = bPhase1
        If tpDisRes.Visible <> bPhase1 Then tpDisRes.Visible = bPhase1
        If cboJamEffect.Visible <> bPhase1 Then cboJamEffect.Visible = bPhase1
        If lblJamEffect.Visible <> bPhase1 Then lblJamEffect.Visible = bPhase1
        If cboJamImmune.Visible <> bPhase1 Then cboJamImmune.Visible = bPhase1
        If lblJamImmune.Visible <> bPhase1 Then lblJamImmune.Visible = bPhase1
        If tpJamStrength.Visible <> bPhase1 Then tpJamStrength.Visible = bPhase1
        If tpJamTargets.Visible <> bPhase1 Then tpJamTargets.Visible = bPhase1
        If stmEmitter.Visible <> bPhase1 Then stmEmitter.Visible = bPhase1
        If tpEmitter.Visible <> bPhase1 Then tpEmitter.Visible = bPhase1
        If stmDetection.Visible <> bPhase1 Then stmDetection.Visible = bPhase1
        If tpDetection.Visible <> bPhase1 Then tpDetection.Visible = bPhase1
        If stmCollection.Visible <> bPhase1 Then stmCollection.Visible = bPhase1
        If tpCollection.Visible <> bPhase1 Then tpCollection.Visible = bPhase1
        If stmCasing.Visible <> bPhase1 Then stmCasing.Visible = bPhase1
        If tpCasing.Visible <> bPhase1 Then tpCasing.Visible = bPhase1
        If lblEmitter.Visible <> bPhase1 Then lblEmitter.Visible = bPhase1
        If lblDetection.Visible <> bPhase1 Then lblDetection.Visible = bPhase1
        If lblCollection.Visible <> bPhase1 Then lblCollection.Visible = bPhase1
        If lblCasing.Visible <> bPhase1 Then lblCasing.Visible = bPhase1

        If mfrmBuilderCost.Visible <> bPhase2 Then mfrmBuilderCost.Visible = bPhase2

    End Sub

    Private mlMineralIDs(3) As Int32
    Private Sub SetButtonClicked(ByVal lMinIdx As Int32)

        Dim lHullTechID As Int32 = -1
        If cboHullType.ListIndex > -1 Then
            lHullTechID = cboHullType.ItemData(cboHullType.ListIndex)
        End If

        Dim oTech As New RadarTechComputer
        oTech.lHullTypeID = lHullTechID

        If lMinIdx = -1 Then Return

        mlSelectedMineralIdx = -1
        oTech.MineralCBOExpanded(lMinIdx, ObjectType.eRadarTech)
        mlSelectedMineralIdx = lMinIdx

        Dim ofrmMin As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
        If ofrmMin Is Nothing Then Return
        ofrmMin.bRaiseSelectEvent = True
        Try
            RemoveHandler ofrmMin.MineralSelected, AddressOf Mineral_Selected
        Catch
        End Try
        AddHandler ofrmMin.MineralSelected, AddressOf Mineral_Selected
    End Sub
    Private mlSelectedMineralIdx As Int32 = -1
    Private Sub Mineral_Selected(ByVal lMineralID As Int32)
        If mlSelectedMineralIdx > -1 Then
            mlMineralIDs(mlSelectedMineralIdx) = lMineralID
            Dim sMinName As String = "Unknown Mineral"
            For X As Int32 = 0 To glMineralUB
                If glMineralIdx(X) = lMineralID Then
                    sMinName = goMinerals(X).MineralName
                    Exit For
                End If
            Next X

            Select Case mlSelectedMineralIdx
                Case 0
                    stmEmitter.SetMineralName(sMinName)
                Case 1
                    stmDetection.SetMineralName(sMinName)
                Case 2
                    stmCollection.SetMineralName(sMinName)
                Case 3
                    stmCasing.SetMineralName(sMinName)
            End Select
        End If

        'BuilderCostValueChange(False)
        CheckForDARequest()
    End Sub

    Private Sub HullTypeSelected(ByVal lItemIndex As Int32)
        EnableDisableControls()
        CheckForDARequest()
    End Sub

    Private Sub CheckForDARequest()
        Dim bRequestDA As Boolean = True
        For X As Int32 = 0 To mlMineralIDs.GetUpperBound(0)
            If mlMineralIDs(X) < 1 Then
                bRequestDA = False
                Exit For
            End If
        Next X
        Dim lHullTypeID As Int32 = -1
        If cboHullType Is Nothing = False AndAlso cboHullType.ListIndex > -1 Then
            lHullTypeID = cboHullType.ItemData(cboHullType.ListIndex)
        End If
        If bRequestDA = True AndAlso lHullTypeID <> -1 Then RequestDAValues(ObjectType.eRadarTech, 0, CByte(lHullTypeID), mlMineralIDs(0), mlMineralIDs(1), mlMineralIDs(2), mlMineralIDs(3), -1, -1, 0, 0)
    End Sub

    Private msRenameVal As String = ""
    Private Sub btnRename_Click(ByVal sName As String) Handles btnRename.Click
        If moTech Is Nothing = False Then
            If moTech.OwnerID = glPlayerID Then
                If goCurrentPlayer.blCredits < 10000000 Then
                    MyBase.moUILib.AddNotification("You require 10,000,000 credits to rename a tech.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If
                Dim sVal As String = txtTechName.Caption.Trim
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

    Private Sub btnAutoFill_Click(ByVal sName As String) Handles btnAutoFill.Click
        'we drop focus for our controls just incase some invalid value causes recalculating
        If goUILib.FocusedControl Is Nothing = False Then
            goUILib.FocusedControl.HasFocus = False
            goUILib.FocusedControl = Nothing
        End If
        BuilderCostValueChange(True)
    End Sub

    Private Sub chkAreaEffectJamming_Click()
        If chkAreaEffectJamming.Value = True Then
            tpJamTargets.Enabled = False
            tpJamTargets.MinValue = 0
            tpJamTargets.PropertyValue = 0
        Else
            tpJamTargets.Enabled = True
        End If
    End Sub
End Class