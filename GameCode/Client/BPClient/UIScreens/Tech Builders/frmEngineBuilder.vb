Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Public Class frmEngineBuilder
    Inherits frmTechBuilder

    Private lblEngineType As UILabel
    Private cboEngineType As UIComboBox

    Private tpPowGen As ctlTechProp
    Private tpThrustGen As ctlTechProp
    Private tpMaxSpd As ctlTechProp
    Private tpMan As ctlTechProp

    Private lblStructuralMaterials As UILabel
    Private lblStructBody As UILabel
    Private lblStructFrame As UILabel
    Private lblStructMeld As UILabel
    Private lblDriveMaterials As UILabel
    Private lblDriveBody As UILabel
    Private lblDriveFrame As UILabel
    Private lblDriveMeld As UILabel
    Private lblColor As UILabel
    Private cboColor As UIComboBox

 
    'Private lblStructBodyItem As UILabel
    'Private lblStructFrameItem As UILabel
    'Private lblStructMeldItem As UILabel
    'Private lblDriveBodyItem As UILabel
    'Private lblDriveFrameItem As UILabel
    'Private lblDriveMeldItem As UILabel
    'Private btnSetStructBodyItem As UIButton
    'Private btnSetStructFrameItem As UIButton
    'Private btnSetStructMeldItem As UIButton
    'Private btnSetDriveBodyItem As UIButton
    'Private btnSetDriveFrameItem As UIButton
    'Private btnSetDriveMeldItem As UIButton
    Private stmStructBody As ctlSetTechMineral
    Private stmStructFrame As ctlSetTechMineral
    Private stmStructMeld As ctlSetTechMineral
    Private stmDriveBody As ctlSetTechMineral
    Private stmDriveFrame As ctlSetTechMineral
    Private stmDriveMeld As ctlSetTechMineral

    Private lblTitle As UILabel

    Private tpDriveMeld As ctlTechProp
    Private tpDriveFrame As ctlTechProp
    Private tpDriveBody As ctlTechProp
    Private tpStructMeld As ctlTechProp
    Private tpStructFrame As ctlTechProp
    Private tpStructBody As ctlTechProp

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

    Private moTech As EngineTech = Nothing
    Private mfrmResCost As frmProdCost = Nothing
    Private mfrmProdCost As frmProdCost = Nothing

    Private mlMaxSpeed As Int32 = 30
    Private mlMaxManeuver As Int32 = 20
    Private mlMaxPowerThrust As Int32 = 10000

    Private mbIgnoreValueChange As Boolean = False
    Private mbImpossibleDesign As Boolean = False
 
    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmEngineDesigner initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eEngineBuilder
            .ControlName = "frmEngineBuilder"
            .Left = 5
            .Top = 10
            .Width = 490 '420
            .Height = 512 '520
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
            '.FillColor = System.Drawing.Color.FromArgb(128, 64, 128, 192)
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = False
            .mbAcceptReprocessEvents = True
        End With

        'txtTechName initial props
        txtTechName = New UITextBox(oUILib)
        With txtTechName
            .ControlName = "txtTechName"
            .Left = 170 '220
            .Top = 45
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
            .Left = 15
            .Top = 45
            .Width = 164
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Engine Name:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTechName, UIControl))

        tpPowGen = New ctlTechProp(oUILib)
        With tpPowGen
            .ControlName = "tpPowGen"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 255
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Power Generation:"
            .PropertyValue = 0
            .Top = 89
            .Visible = True
            '.Width = 512
        End With
        Me.AddChild(CType(tpPowGen, UIControl))

        tpThrustGen = New ctlTechProp(oUILib)
        With tpThrustGen
            .ControlName = "tpThrustGen"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 1000000
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Thrust Generation: "
            .PropertyValue = 0
            .Top = 111
            .Visible = True
            '.Width = 512
        End With
        Me.AddChild(CType(tpThrustGen, UIControl))

        tpMaxSpd = New ctlTechProp(oUILib)
        With tpMaxSpd
            .ControlName = "tpMaxSpd"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 255
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Maximum Speed: "
            .PropertyValue = 0
            .Top = 133
            .Visible = True
            '.Width = 512
        End With
        Me.AddChild(CType(tpMaxSpd, UIControl))

        tpMan = New ctlTechProp(oUILib)
        With tpMan
            .ControlName = "tpMan"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 255
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Maneuverability: "
            .PropertyValue = 0
            .Top = 155
            .Visible = True
            '.Width = 512
        End With
        Me.AddChild(CType(tpMan, UIControl))

        'lblStructBody initial props
        lblStructuralMaterials = New UILabel(oUILib)
        With lblStructuralMaterials
            .ControlName = "lblStructuralMaterials"
            .Left = 15
            .Top = 175
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Select Structural Materials:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblStructuralMaterials, UIControl))

        'lblStructBody initial props
        lblStructBody = New UILabel(oUILib)
        With lblStructBody
            .ControlName = "lblStructBody"
            .Left = 15
            .Top = 195
            .Width = 45
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Body:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblStructBody, UIControl))

        tpStructBody = New ctlTechProp(oUILib)
        With tpStructBody
            .ControlName = "tpStructBody"
            .Enabled = True
            .Height = 18
            .Left = 235
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = ""
            .PropertyValue = 0
            .Top = lblStructBody.Top
            .Visible = True
            .Width = 250
            .SetNoPropDisplay(True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpStructBody, UIControl))

        'lblStructFrame initial props
        lblStructFrame = New UILabel(oUILib)
        With lblStructFrame
            .ControlName = "lblStructFrame"
            .Left = 15
            .Top = 220
            .Width = 45
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Frame:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblStructFrame, UIControl))

        tpStructFrame = New ctlTechProp(oUILib)
        With tpStructFrame
            .ControlName = "tpStructFrame"
            .Enabled = True
            .Height = 18
            .Left = 235
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = ""
            .PropertyValue = 0
            .Top = lblStructFrame.Top
            .Visible = True
            .Width = 250
            .SetNoPropDisplay(True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpStructFrame, UIControl))

        'lblStructMeld initial props
        lblStructMeld = New UILabel(oUILib)
        With lblStructMeld
            .ControlName = "lblStructMeld"
            .Left = 15
            .Top = 245
            .Width = 45
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Meld:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblStructMeld, UIControl))

        tpStructMeld = New ctlTechProp(oUILib)
        With tpStructMeld
            .ControlName = "tpStructMeld"
            .Enabled = True
            .Height = 18
            .Left = 235
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = ""
            .PropertyValue = 0
            .Top = lblStructMeld.Top
            .Visible = True
            .Width = 250
            .SetNoPropDisplay(True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpStructMeld, UIControl))

        'lblDriveMaterials initial props
        lblDriveMaterials = New UILabel(oUILib)
        With lblDriveMaterials
            .ControlName = "lblDriveMaterials"
            .Left = 15
            .Top = 270
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Select Drive Materials:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblDriveMaterials, UIControl))

        'lblDriveBody initial props
        lblDriveBody = New UILabel(oUILib)
        With lblDriveBody
            .ControlName = "lblDriveBody"
            .Left = 15
            .Top = 290
            .Width = 45
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Body:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblDriveBody, UIControl))

        tpDriveBody = New ctlTechProp(oUILib)
        With tpDriveBody
            .ControlName = "tpDriveBody"
            .Enabled = True
            .Height = 18
            .Left = 235
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = ""
            .PropertyValue = 0
            .Top = lblDriveBody.Top
            .Visible = True
            .Width = 250
            .SetNoPropDisplay(True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpDriveBody, UIControl))

        'lblDriveFrame initial props
        lblDriveFrame = New UILabel(oUILib)
        With lblDriveFrame
            .ControlName = "lblDriveFrame"
            .Left = 15
            .Top = 315
            .Width = 45
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Frame:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblDriveFrame, UIControl))

        tpDriveFrame = New ctlTechProp(oUILib)
        With tpDriveFrame
            .ControlName = "tpDriveFrame"
            .Enabled = True
            .Height = 18
            .Left = 235
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = ""
            .PropertyValue = 0
            .Top = lblDriveFrame.Top
            .Visible = True
            .Width = 250
            .SetNoPropDisplay(True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpDriveFrame, UIControl))

        'lblDriveMeld initial props
        lblDriveMeld = New UILabel(oUILib)
        With lblDriveMeld
            .ControlName = "lblDriveMeld"
            .Left = 15
            .Top = 340
            .Width = 45
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Meld:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblDriveMeld, UIControl))

        tpDriveMeld = New ctlTechProp(oUILib)
        With tpDriveMeld
            .ControlName = "tpDriveMeld"
            .Enabled = True
            .Height = 18
            .Left = 235
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = ""
            .PropertyValue = 0
            .Top = lblDriveMeld.Top
            .Visible = True
            .Width = 250
            .SetNoPropDisplay(True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpDriveMeld, UIControl))

        'lblColor initial props
        lblColor = New UILabel(oUILib)
        With lblColor
            .ControlName = "lblColor"
            .Left = 15
            .Top = 385
            .Width = 200
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Select Engine Color:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblColor, UIControl))

        'lblDesignFlaw initial props
        lblDesignFlaw = New UILabel(oUILib)
        With lblDesignFlaw
            .ControlName = "lblDesignFlaw"
            .Left = 15
            .Top = 410
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
            .Left = (Me.Width \ 2) - 50 '130
            .Top = 470
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
            .Left = Me.Width - 110 '260
            .Top = 470
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

        'btnDelete initial props
        btnDelete = New UIButton(oUILib)
        With btnDelete
            .ControlName = "btnDelete"
            .Left = btnDesign.Left - 105
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

        'btnAutoFill initial props
        btnAutoFill = New UIButton(oUILib)
        With btnAutoFill
            .ControlName = "btnAutoFill"
            .Left = 10 'btnDesign.Left - 105
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

        'cboColor initial props
        cboColor = New UIComboBox(oUILib)
        With cboColor
            .ControlName = "cboColor"
            .Left = 220
            .Top = 385
            .Width = 175
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
        Me.AddChild(CType(cboColor, UIControl))

        ''lblDriveMeldItem initial props
        'lblDriveMeldItem = New UILabel(oUILib)
        'With lblDriveMeldItem
        '    .ControlName = "lblDriveMeldItem"
        '    .Left = 55
        '    .Top = 340
        '    .Width = 125
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Unselected"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        'End With
        'Me.AddChild(CType(lblDriveMeldItem, UIControl))
        ''btnSetDriveMeldItem initial props
        'btnSetDriveMeldItem = New UIButton(oUILib)
        'With btnSetDriveMeldItem
        '    .ControlName = "btnSetDriveMeldItem"
        '    .Left = lblDriveMeldItem.Left + lblDriveMeldItem.Width
        '    .Top = lblDriveMeldItem.Top
        '    .Width = 50
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Set"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        'End With
        'Me.AddChild(CType(btnSetDriveMeldItem, UIControl))
        stmDriveMeld = New ctlSetTechMineral(oUILib)
        With stmDriveMeld
            .ControlName = "stmDriveMeld"
            .Left = 55
            .Top = 340
            .Width = 175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 5
        End With
        Me.AddChild(CType(stmDriveMeld, UIControl))

        ''cboDriveMeld initial props
        'cboDriveMeld = New UIComboBox(oUILib)
        'With cboDriveMeld
        '    .ControlName = "cboDriveMeld"
        '    .Left = 55 '220
        '    .Top = 340
        '    .Width = 175
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
        'Me.AddChild(CType(cboDriveMeld, UIControl))

        ''lblDriveFrameItem initial props
        'lblDriveFrameItem = New UILabel(oUILib)
        'With lblDriveFrameItem
        '    .ControlName = "lblDriveFrameItem"
        '    .Left = 55
        '    .Top = 315
        '    .Width = 125
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Unselected"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        'End With
        'Me.AddChild(CType(lblDriveFrameItem, UIControl))
        ''btnSetDriveFrameItem initial props
        'btnSetDriveFrameItem = New UIButton(oUILib)
        'With btnSetDriveFrameItem
        '    .ControlName = "btnSetDriveFrameItem"
        '    .Left = btnSetDriveMeldItem.Left
        '    .Top = lblDriveFrameItem.Top
        '    .Width = 50
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Set"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        'End With
        'Me.AddChild(CType(btnSetDriveFrameItem, UIControl))
        stmDriveFrame = New ctlSetTechMineral(oUILib)
        With stmDriveFrame
            .ControlName = "stmDriveFrame"
            .Left = 55
            .Top = 315
            .Width = 175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 4
        End With
        Me.AddChild(CType(stmDriveFrame, UIControl))
        ''cboDriveFrame initial props
        'cboDriveFrame = New UIComboBox(oUILib)
        'With cboDriveFrame
        '    .ControlName = "cboDriveFrame"
        '    .Left = 55 ' 220
        '    .Top = 315
        '    .Width = 175
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
        'Me.AddChild(CType(cboDriveFrame, UIControl))

        'lblDrivebodyItem initial props
        'lblDriveBodyItem = New UILabel(oUILib)
        'With lblDriveBodyItem
        '    .ControlName = "lblDrivebodyItem"
        '    .Left = 55
        '    .Top = 290
        '    .Width = 125
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Unselected"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        'End With
        'Me.AddChild(CType(lblDriveBodyItem, UIControl))
        ''btnSetDriveBodyItem initial props
        'btnSetDriveBodyItem = New UIButton(oUILib)
        'With btnSetDriveBodyItem
        '    .ControlName = "btnSetDriveBodyItem"
        '    .Left = btnSetDriveMeldItem.Left
        '    .Top = lblDriveBodyItem.Top
        '    .Width = 50
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Set"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        'End With
        'Me.AddChild(CType(btnSetDriveBodyItem, UIControl))
        stmDriveBody = New ctlSetTechMineral(oUILib)
        With stmDriveBody
            .ControlName = "stmDriveBody"
            .Left = 55
            .Top = 290
            .Width = 175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 3
        End With
        Me.AddChild(CType(stmDriveBody, UIControl))
        ''cboDriveBody initial props
        'cboDriveBody = New UIComboBox(oUILib)
        'With cboDriveBody
        '    .ControlName = "cboDriveBody"
        '    .Left = 55 '220
        '    .Top = 290
        '    .Width = 175
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
        'Me.AddChild(CType(cboDriveBody, UIControl))

        'lblStructMeldItem initial props
        'lblStructMeldItem = New UILabel(oUILib)
        'With lblStructMeldItem
        '    .ControlName = "lblStructMeldItem"
        '    .Left = 55
        '    .Top = 245
        '    .Width = 125
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Unselected"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        'End With
        'Me.AddChild(CType(lblStructMeldItem, UIControl))
        ''btnSetStructMeldItem initial props
        'btnSetStructMeldItem = New UIButton(oUILib)
        'With btnSetStructMeldItem
        '    .ControlName = "btnSetStructMeldItem"
        '    .Left = btnSetDriveMeldItem.Left
        '    .Top = lblStructMeldItem.Top
        '    .Width = 50
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Set"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        'End With
        'Me.AddChild(CType(btnSetStructMeldItem, UIControl))
        stmStructMeld = New ctlSetTechMineral(oUILib)
        With stmStructMeld
            .ControlName = "stmStructMeld"
            .Left = 55
            .Top = 245
            .Width = 175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 2
        End With
        Me.AddChild(CType(stmStructMeld, UIControl))
        ''cboStructMeld initial props
        'cboStructMeld = New UIComboBox(oUILib)
        'With cboStructMeld
        '    .ControlName = "cboStructMeld"
        '    .Left = 55 '220
        '    .Top = 245
        '    .Width = 175
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
        'Me.AddChild(CType(cboStructMeld, UIControl))

        'lblStructFrameItem initial props
        'lblStructFrameItem = New UILabel(oUILib)
        'With lblStructFrameItem
        '    .ControlName = "lblStructFrameItem"
        '    .Left = 55
        '    .Top = 220
        '    .Width = 125
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Unselected"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        'End With
        'Me.AddChild(CType(lblStructFrameItem, UIControl))
        ''btnSetStructFrameItem initial props
        'btnSetStructFrameItem = New UIButton(oUILib)
        'With btnSetStructFrameItem
        '    .ControlName = "btnSetStructFrameItem"
        '    .Left = btnSetDriveMeldItem.Left
        '    .Top = lblStructFrameItem.Top
        '    .Width = 50
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Set"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        'End With
        'Me.AddChild(CType(btnSetStructFrameItem, UIControl))
        stmStructFrame = New ctlSetTechMineral(oUILib)
        With stmStructFrame
            .ControlName = "stmStructFrame"
            .Left = 55
            .Top = 220
            .Width = 175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 1
        End With
        Me.AddChild(CType(stmStructFrame, UIControl))
        ''cboStructFrame initial props
        'cboStructFrame = New UIComboBox(oUILib)
        'With cboStructFrame
        '    .ControlName = "cboStructFrame"
        '    .Left = 55 '220
        '    .Top = 220
        '    .Width = 175
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
        'Me.AddChild(CType(cboStructFrame, UIControl))

        ''lblStructbodyItem initial props
        'lblStructBodyItem = New UILabel(oUILib)
        'With lblStructBodyItem
        '    .ControlName = "lblStructbodyItem"
        '    .Left = 55
        '    .Top = 195
        '    .Width = 125
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Unselected"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        'End With
        'Me.AddChild(CType(lblStructBodyItem, UIControl))
        ''btnSetStructBodyItem initial props
        'btnSetStructBodyItem = New UIButton(oUILib)
        'With btnSetStructBodyItem
        '    .ControlName = "btnSetStructBodyItem"
        '    .Left = btnSetDriveMeldItem.Left
        '    .Top = lblStructBodyItem.Top
        '    .Width = 50
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Set"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        'End With
        'Me.AddChild(CType(btnSetStructBodyItem, UIControl))
        stmStructBody = New ctlSetTechMineral(oUILib)
        With stmStructBody
            .ControlName = "stmStructbody"
            .Left = 55
            .Top = 195
            .Width = 175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 0
        End With
        Me.AddChild(CType(stmStructBody, UIControl))

        ''cboStructBody initial props
        'cboStructBody = New UIComboBox(oUILib)
        'With cboStructBody
        '    .ControlName = "cboStructBody"
        '    .Left = 55 '220
        '    .Top = 195
        '    .Width = 175
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
        'Me.AddChild(CType(cboStructBody, UIControl))

        'btnRename initial props
        btnRename = New UIButton(oUILib)
        With btnRename
            .ControlName = "btnRename"
            .Left = txtTechName.Left + txtTechName.Width + 5
            .Top = txtTechName.Top - 1
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = False
            .Caption = "Rename"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnRename, UIControl))

        'lblEngineType initial props
        lblEngineType = New UILabel(oUILib)
        With lblEngineType
            .ControlName = "lblEngineType"
            .Left = 15
            .Top = 67
            .Width = 100
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Engine Type:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblEngineType, UIControl))

        'cboEngineType initial props
        cboEngineType = New UIComboBox(oUILib)
        With cboEngineType
            .ControlName = "cboEngineType"
            .Left = txtTechName.Left '220
            .Top = 67
            .Width = 175
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
        Me.AddChild(CType(cboEngineType, UIControl))

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 15
            .Top = 10
            .Width = 214
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Engine Designer"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Arial", 18, FontStyle.Bold Or FontStyle.Italic, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        FillValues()
        EnableDisableControls()

        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
        ofrm.ShowMineralDetail(Me.Left + Me.Width + 5, Me.Top, Me.Height, -1)

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        glCurrentEnvirView = CurrentView.eEngineResearch
  
        AddHandler stmDriveBody.SetButtonClicked, AddressOf SetButtonClicked
        AddHandler stmDriveFrame.SetButtonClicked, AddressOf SetButtonClicked
        AddHandler stmDriveMeld.SetButtonClicked, AddressOf SetButtonClicked
        AddHandler stmStructBody.SetButtonClicked, AddressOf SetButtonClicked
        AddHandler stmStructFrame.SetButtonClicked, AddressOf SetButtonClicked
        AddHandler stmStructMeld.SetButtonClicked, AddressOf SetButtonClicked

        AddHandler tpDriveBody.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpDriveFrame.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpDriveMeld.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpMan.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpMaxSpd.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpPowGen.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpStructBody.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpStructFrame.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpStructMeld.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpThrustGen.PropertyValueChanged, AddressOf tp_PropertyValueChanged

        AddHandler cboEngineType.ItemSelected, AddressOf HullTypeSelected

        If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
        goFullScreenBackground = Nothing
        goFullScreenBackground = goResMgr.GetTexture("enginebuilder.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "TechBacks.pak")

    End Sub
    Public Overrides Sub ViewResults(ByRef oTech As Base_Tech, ByVal lProdFactor As Integer)
        If oTech Is Nothing Then Return

        moTech = CType(oTech, EngineTech)

        Me.btnDesign.Enabled = False

        With moTech
            mbIgnoreValueChange = True

            cboEngineType.FindComboItemData(.HullTypeID)

            tpMan.PropertyValue = .Maneuver
            If .Maneuver > 0 Then tpMan.PropertyLocked = True
            tpMaxSpd.PropertyValue = .Speed
            If .Speed > 0 Then tpMaxSpd.PropertyLocked = True
            tpPowGen.PropertyValue = .PowerProd
            If .PowerProd > 0 Then tpPowGen.PropertyLocked = True
            Me.txtTechName.Caption = .sEngineName
            tpThrustGen.PropertyValue = .Thrust
            If .Thrust > 0 Then tpThrustGen.PropertyLocked = True

            mlSelectedMineralIdx = 0 : Mineral_Selected(.lStructuralBodyMineralID)
            mlSelectedMineralIdx = 1 : Mineral_Selected(.lStructuralFrameMineralID)
            mlSelectedMineralIdx = 2 : Mineral_Selected(.lStructuralMeldMineralID)
            mlSelectedMineralIdx = 3 : Mineral_Selected(.lDriveBodyMineralID)
            mlSelectedMineralIdx = 4 : Mineral_Selected(.lDriveFrameMineralID)
            mlSelectedMineralIdx = 5 : Mineral_Selected(.lDriveMeldMineralID)
            
            tpStructBody.PropertyValue = .lSpecifiedMin1
            If .lSpecifiedMin1 <> -1 Then tpStructBody.PropertyLocked = True
            tpStructFrame.PropertyValue = .lSpecifiedMin2
            If .lSpecifiedMin2 <> -1 Then tpStructFrame.PropertyLocked = True
            tpStructMeld.PropertyValue = .lSpecifiedMin3
            If .lSpecifiedMin3 <> -1 Then tpStructMeld.PropertyLocked = True
            tpDriveBody.PropertyValue = .lSpecifiedMin4
            If .lSpecifiedMin4 <> -1 Then tpDriveBody.PropertyLocked = True
            tpDriveFrame.PropertyValue = .lSpecifiedMin5
            If .lSpecifiedMin5 <> -1 Then tpDriveFrame.PropertyLocked = True
            tpDriveMeld.PropertyValue = .lSpecifiedMin6
            If .lSpecifiedMin6 <> -1 Then tpDriveMeld.PropertyLocked = True

            Me.cboColor.FindComboItemData(.ColorValue)

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
            btnDelete.Visible = True 'oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched
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
            cboEngineType.Enabled = False
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
                .SetFromProdCost(oTech.oProductionCost, 1000, False, moTech.HullRequired, 0)
            End With
        ElseIf oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
            If mfrmResCost Is Nothing Then mfrmResCost = New frmProdCost(MyBase.moUILib, "frmResCost")
            With mfrmResCost
                .Visible = True
                .Left = Me.Left + Me.Width + 5
                .Top = Me.Top + Me.Height + 8
                .SetFromFailureCode(oTech.ErrorCodeReason)
            End With
        End If
    End Sub
 
    Protected Overrides Sub BuilderCostValueChange(ByVal bAutoFill As Boolean)
        If mbIgnoreValueChange = True Then Return
        EnableDisableControls()
        mbIgnoreValueChange = True
        mbImpossibleDesign = True
        Dim oEngineTechComputer As New EngineTechComputer()
        With oEngineTechComputer
            .blMan = tpMan.PropertyValue
            .blMaxSpd = tpMaxSpd.PropertyValue
            .blPowGen = tpPowGen.PropertyValue
            .blThrustGen = tpThrustGen.PropertyValue

            .lHullTypeID = -1
            If cboEngineType.ListIndex > -1 Then
                .lHullTypeID = cboEngineType.ItemData(cboEngineType.ListIndex)
                If .lHullTypeID = Base_Tech.eyHullType.Facility OrElse .lHullTypeID = Base_Tech.eyHullType.SpaceStation Then
                    tpMaxSpd.PropertyValue = 0
                    tpMan.PropertyValue = 0
                    tpThrustGen.PropertyValue = 0
                    tpMaxSpd.MaxValue = 0 : tpMaxSpd.MinValue = 0
                    tpMan.MaxValue = 0 : tpMan.MinValue = 0
                    tpThrustGen.MaxValue = 0 : tpThrustGen.MinValue = 0
                Else
                    If tpMan.MaxValue <> mlMaxManeuver Then tpMan.MaxValue = mlMaxManeuver
                    If tpMaxSpd.MaxValue <> mlMaxSpeed Then tpMaxSpd.MaxValue = mlMaxSpeed
                    'If tpThrustGen.MaxValue <> mlMaxPowerThrust Then tpThrustGen.MaxValue = mlMaxPowerThrust
                End If
            End If

            .SetMinDAValues(ml_Min1DA, ml_Min2DA, ml_Min3DA, ml_Min4DA, ml_Min5DA, ml_Min6DA)

            Dim decNormalizer As Decimal = 1D
            Dim lMaxGuns As Int32 = 1
            Dim lMaxDPS As Int32 = 1
            Dim lMaxHullSize As Int32 = 1
            Dim lHullAvail As Int32 = 1
            Dim lMaxPower As Int32 = 1
            If .lHullTypeID = -1 Then
                lblDesignFlaw.Caption = "Invalid Design: a Hull Type Must be selected!"
                mbIgnoreValueChange = False
                Return
            End If
            TechBuilderComputer.GetTypeValues(.lHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)
            'tpPowGen.MaxValue = Math.Max(tpPowGen.PropertyValue, Math.Min(lMaxPower, mlMaxPowerThrust))
            'tpPowGen.blAbsoluteMaximum = mlMaxPowerThrust
            'If tpPowGen.PropertyValue > tpPowGen.blAbsoluteMaximum Then tpPowGen.PropertyValue = tpPowGen.blAbsoluteMaximum
            'tpThrustGen.MaxValue = Math.Min(lMaxHullSize, mlMaxPowerThrust)
            'If tpThrustGen.PropertyValue > tpThrustGen.MaxValue Then tpThrustGen.PropertyValue = tpThrustGen.MaxValue

            .lMineral1ID = mlMineralIDs(0)
            .lMineral2ID = mlMineralIDs(1)
            .lMineral3ID = mlMineralIDs(2)
            .lMineral4ID = mlMineralIDs(3)
            .lMineral5ID = mlMineralIDs(4)
            .lMineral6ID = mlMineralIDs(5)

            If .lMineral1ID < 1 OrElse .lMineral2ID < 1 OrElse .lMineral3ID < 1 OrElse .lMineral4ID < 1 OrElse .lMineral5ID < 1 OrElse .lMineral6ID < 1 Then
                lblDesignFlaw.Caption = "All properties and materials need to be defined."
                mbIgnoreValueChange = False
                mbImpossibleDesign = True
                Return
            End If


            .msMin1Name = "Structural Body"
            .msMin2Name = "Structural Frame"
            .msMin3Name = "Structural Meld"
            .msMin4Name = "Drive Body"
            .msMin5Name = "Drive Frame"
            .msMin6Name = "Drive Meld"

            If bAutoFill = True Then
                Dim bDoNotLockCrew As Boolean = .lHullTypeID = Base_Tech.eyHullType.Facility OrElse .lHullTypeID = Base_Tech.eyHullType.SpaceStation
                .DoBuilderCostPreconfigure(lblDesignFlaw, MyBase.mfrmBuilderCost, tpStructBody, tpStructFrame, tpStructMeld, tpDriveBody, tpDriveFrame, tpDriveMeld, ObjectType.eEngineTech, 0D, bDoNotLockCrew)
            Else
                .BuilderCostValueChange(lblDesignFlaw, MyBase.mfrmBuilderCost, tpStructBody, tpStructFrame, tpStructMeld, tpDriveBody, tpDriveFrame, tpDriveMeld, ObjectType.eEngineTech, 0D)
            End If
            mbImpossibleDesign = .bImpossibleDesign
        End With

        mbIgnoreValueChange = False
    End Sub

    Private Sub EnableDisableControls()
        If mbIgnoreValueChange = True Then Return
        If cboEngineType Is Nothing Then Return

        Dim bValue As Boolean = cboEngineType.ListIndex > -1

        If tpPowGen.Enabled <> bValue Then tpPowGen.Enabled = bValue
        If tpThrustGen.Enabled <> bValue Then tpThrustGen.Enabled = bValue
        If tpMaxSpd.Enabled <> bValue Then tpMaxSpd.Enabled = bValue
        If tpMan.Enabled <> bValue Then tpMan.Enabled = bValue
        If tpDriveMeld.Enabled <> bValue Then tpDriveMeld.Enabled = bValue
        If tpDriveFrame.Enabled <> bValue Then tpDriveFrame.Enabled = bValue
        If tpDriveBody.Enabled <> bValue Then tpDriveBody.Enabled = bValue
        If tpStructMeld.Enabled <> bValue Then tpStructMeld.Enabled = bValue
        If tpStructFrame.Enabled <> bValue Then tpStructFrame.Enabled = bValue
        If tpStructBody.Enabled <> bValue Then tpStructBody.Enabled = bValue

        Dim lPhase As Int32 = 0
        If bValue = True Then lPhase = 1
        If lPhase = 1 Then
            If tpPowGen.PropertyValue > 0 OrElse tpThrustGen.PropertyValue > 0 Then
                lPhase = 2

                lPhase = 3
                For X As Int32 = 0 To 5
                    If mlMineralIDs(X) < 1 Then
                        lPhase = 2
                        Exit For
                    End If
                Next X
                
            End If
        End If
        If moTech Is Nothing = False Then lPhase = 3

        Dim bPhase1 As Boolean = lPhase > 0
        Dim bPhase2 As Boolean = lPhase > 1
        Dim bPhase3 As Boolean = lPhase > 2

        If tpPowGen.Visible <> bPhase1 Then tpPowGen.Visible = bPhase1
        If tpThrustGen.Visible <> bPhase1 Then tpThrustGen.Visible = bPhase1
        If tpMaxSpd.Visible <> bPhase1 Then tpMaxSpd.Visible = bPhase1
        If tpMan.Visible <> bPhase1 Then tpMan.Visible = bPhase1
        If lblDriveMaterials.Visible <> bPhase2 Then lblDriveMaterials.Visible = bPhase2
        If lblStructuralMaterials.Visible <> bPhase2 Then lblStructuralMaterials.Visible = bPhase2
        If lblStructBody.Visible <> bPhase2 Then lblStructBody.Visible = bPhase2
        If lblStructFrame.Visible <> bPhase2 Then lblStructFrame.Visible = bPhase2
        If lblStructMeld.Visible <> bPhase2 Then lblStructMeld.Visible = bPhase2
        If lblDriveBody.Visible <> bPhase2 Then lblDriveBody.Visible = bPhase2
        If lblDriveFrame.Visible <> bPhase2 Then lblDriveFrame.Visible = bPhase2
        If lblDriveMeld.Visible <> bPhase2 Then lblDriveMeld.Visible = bPhase2
        If stmStructBody.Visible <> bPhase2 Then stmStructBody.Visible = bPhase2
        If stmStructFrame.Visible <> bPhase2 Then stmStructFrame.Visible = bPhase2
        If stmStructMeld.Visible <> bPhase2 Then stmStructMeld.Visible = bPhase2
        If stmDriveBody.Visible <> bPhase2 Then stmDriveBody.Visible = bPhase2
        If stmDriveFrame.Visible <> bPhase2 Then stmDriveFrame.Visible = bPhase2
        If stmDriveMeld.Visible <> bPhase2 Then stmDriveMeld.Visible = bPhase2

        If tpStructBody.Visible <> bPhase2 Then tpStructBody.Visible = bPhase2
        If tpStructFrame.Visible <> bPhase2 Then tpStructFrame.Visible = bPhase2
        If tpStructMeld.Visible <> bPhase2 Then tpStructMeld.Visible = bPhase2
        If tpDriveBody.Visible <> bPhase2 Then tpDriveBody.Visible = bPhase2
        If tpDriveMeld.Visible <> bPhase2 Then tpDriveMeld.Visible = bPhase2
        If tpDriveFrame.Visible <> bPhase2 Then tpDriveFrame.Visible = bPhase2
        If lblColor.Visible <> bPhase2 Then lblColor.Visible = bPhase2
        If cboColor.Visible <> bPhase2 Then cboColor.Visible = bPhase2

        If mfrmBuilderCost.Visible <> bPhase3 Then mfrmBuilderCost.Visible = bPhase3
    End Sub

    Private Sub HullTypeSelected(ByVal lItemIndex As Int32)
        EnableDisableControls()
        If lItemIndex > -1 Then
            Dim decNormalizer As Decimal = 1D
            Dim lMaxGuns As Int32 = 1
            Dim lMaxDPS As Int32 = 1
            Dim lMaxHullSize As Int32 = 1
            Dim lHullAvail As Int32 = 1
            Dim lMaxPower As Int32 = 1

            Dim lHullTypeID As Int32 = cboEngineType.ItemData(cboEngineType.ListIndex)
            TechBuilderComputer.GetTypeValues(lHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)

            tpPowGen.MaxValue = Math.Min(lMaxPower * 4, mlMaxPowerThrust)
            tpPowGen.blAbsoluteMaximum = mlMaxPowerThrust
            If tpPowGen.PropertyValue > lMaxPower Then tpPowGen.PropertyValue = lMaxPower
            If tpPowGen.PropertyValue > tpPowGen.blAbsoluteMaximum Then tpPowGen.PropertyValue = tpPowGen.blAbsoluteMaximum
            tpThrustGen.MaxValue = Math.Min(lMaxHullSize, mlMaxPowerThrust)
            If tpThrustGen.PropertyValue > tpThrustGen.MaxValue Then tpThrustGen.PropertyValue = tpThrustGen.MaxValue
            tpPowGen.UpdateTextDisplay()
            tpThrustGen.UpdateTextDisplay()
            tpThrustGen.ToolTipText = "Amount of hull that this engine will be able to move." & vbCrLf & "Maximum Value: " & tpThrustGen.MaxValue
            tpPowGen.ToolTipText = "Amount of power this engine will produce." & vbCrLf & "Maximum Value: " & tpPowGen.MaxValue

            If lHullTypeID <> Base_Tech.eyHullType.SpaceStation AndAlso lHullTypeID <> Base_Tech.eyHullType.Facility Then
                If lHullTypeID = Base_Tech.eyHullType.Frigate Then
                    tpThrustGen.MinValue = 5000
                ElseIf lHullTypeID = Base_Tech.eyHullType.Utility Then
                    tpThrustGen.MinValue = 0
                Else
                    tpThrustGen.MinValue = CInt(lMaxHullSize * 0.5)
                End If

                tpPowGen.MinValue = CInt(lMaxPower * 0.5)
            End If
        End If
        'BuilderCostValueChange(False)
        CheckForDARequest()
    End Sub

    Private Sub ComboBoxItemSelected(ByVal lItemIndex As Int32)
        BuilderCostValueChange(False)
    End Sub

    Public Sub FillValues()
        'Dim lSorted() As Int32 = GetSortedMineralIdxArray(False)

        ''Fill our minerals/alloys
        ''cboFuelCat.Clear() : cboFuelComp.Clear()
        'cboDriveMeld.Clear() : cboDriveFrame.Clear() : cboDriveBody.Clear() : cboStructMeld.Clear() : cboStructFrame.Clear() : cboStructBody.Clear()
        'If lSorted Is Nothing = False Then
        '    For X As Int32 = 0 To lSorted.GetUpperBound(0)
        '        If lSorted(X) <> -1 Then
        '            'If goMinerals(lSorted(X)).IsFullyStudied = False Then Continue For
        '            'cboFuelCat.AddItem(goMinerals(X).MineralName)
        '            'cboFuelCat.ItemData(cboFuelCat.NewIndex) = goMinerals(X).ObjectID
        '            'cboFuelComp.AddItem(goMinerals(X).MineralName)
        '            'cboFuelComp.ItemData(cboFuelComp.NewIndex) = goMinerals(X).ObjectID
        '            cboDriveMeld.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboDriveMeld.ItemData(cboDriveMeld.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '            cboDriveFrame.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboDriveFrame.ItemData(cboDriveFrame.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '            cboDriveBody.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboDriveBody.ItemData(cboDriveBody.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '            cboStructMeld.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboStructMeld.ItemData(cboStructMeld.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '            cboStructFrame.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboStructFrame.ItemData(cboStructFrame.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '            cboStructBody.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboStructBody.ItemData(cboStructBody.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '        End If
        '    Next X
        'End If

        cboColor.Clear()
        cboColor.AddItem("Light Blue") : cboColor.ItemData(cboColor.NewIndex) = 0
        cboColor.AddItem("Bright Green") : cboColor.ItemData(cboColor.NewIndex) = 1
        cboColor.AddItem("Orange") : cboColor.ItemData(cboColor.NewIndex) = 2
        cboColor.AddItem("Blue") : cboColor.ItemData(cboColor.NewIndex) = 3
        cboColor.AddItem("Purple") : cboColor.ItemData(cboColor.NewIndex) = 4
        cboColor.AddItem("Dark Blue") : cboColor.ItemData(cboColor.NewIndex) = 5
        cboColor.AddItem("Red") : cboColor.ItemData(cboColor.NewIndex) = 6
        cboColor.AddItem("Yellow") : cboColor.ItemData(cboColor.NewIndex) = 7
        cboColor.AddItem("Dark Green") : cboColor.ItemData(cboColor.NewIndex) = 8
        cboColor.ListIndex = 0

        SetCboHelpers()

        cboEngineType.Clear()
        If mlMaxPowerThrust > 110000 Then
            cboEngineType.AddItem("Battlecruiser") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.BattleCruiser
        End If
        If mlMaxPowerThrust > 400000 Then
            cboEngineType.AddItem("Battleship") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.Battleship
        End If
        cboEngineType.AddItem("Corvette") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.Corvette
        If mlMaxPowerThrust > 57000 Then
            cboEngineType.AddItem("Cruiser") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.Cruiser
        End If
        If mlMaxPowerThrust > 32000 Then
            cboEngineType.AddItem("Destroyer") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.Destroyer
        End If
        cboEngineType.AddItem("Escort") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.Escort
        cboEngineType.AddItem("Facility") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.Facility
        cboEngineType.AddItem("Fighter (Light)") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.LightFighter
        cboEngineType.AddItem("Fighter (Medium)") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.MediumFighter
        cboEngineType.AddItem("Fighter (Heavy)") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.HeavyFighter
        cboEngineType.AddItem("Frigate") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.Frigate
        Dim bHasNavalUnit As Boolean = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eNavalAbility) > 0
        If bHasNavalUnit = True Then
            If mlMaxPowerThrust > 181000 Then
                cboEngineType.AddItem("Naval Battleship") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.NavalBattleship
                cboEngineType.AddItem("Naval Carrier") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.NavalCarrier
            End If
            If mlMaxPowerThrust > 83000 Then
                cboEngineType.AddItem("Naval Cruiser") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.NavalCruiser
            End If
            If mlMaxPowerThrust > 31000 Then
                cboEngineType.AddItem("Naval Destroyer") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.NavalDestroyer
            End If
            If mlMaxPowerThrust > 15000 Then
                cboEngineType.AddItem("Naval Frigate") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.NavalFrigate
                If mlMaxPowerThrust > 42000 Then
                    cboEngineType.AddItem("Naval Submarine") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.NavalSub
                End If
            End If
            cboEngineType.AddItem("Naval Utility") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.Utility
        End If
        cboEngineType.AddItem("Small Vehicle") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.SmallVehicle
        cboEngineType.AddItem("Space Station") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.SpaceStation
        cboEngineType.AddItem("Tank") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.Tank
        cboEngineType.AddItem("Utility") : cboEngineType.ItemData(cboEngineType.NewIndex) = EngineTech.eyHullType.Utility

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

        'we drop focus for our controls just incase some invalid value causes recalculating
        If goUILib.FocusedControl Is Nothing = False Then
            goUILib.FocusedControl.HasFocus = False
            goUILib.FocusedControl = Nothing
        End If

        'Check the Techname
        If txtTechName.Caption.Trim.Length = 0 Then
            MyBase.moUILib.AddNotification("You must specify a name for this Engine.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If
        If mbImpossibleDesign = True Then
            MyBase.moUILib.AddNotification("You must fix the flaws of this design.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return
        End If

        'Dim lMinID(5) As Int32

        'If cboStructBody.ListIndex > -1 Then lMinID(0) = cboStructBody.ItemData(cboStructBody.ListIndex)
        'If cboStructFrame.ListIndex > -1 Then lMinID(1) = cboStructFrame.ItemData(cboStructFrame.ListIndex)
        'If cboStructMeld.ListIndex > -1 Then lMinID(2) = cboStructMeld.ItemData(cboStructMeld.ListIndex)

        'If cboDriveBody.ListIndex > -1 Then lMinID(3) = cboDriveBody.ItemData(cboDriveBody.ListIndex)
        'If cboDriveFrame.ListIndex > -1 Then lMinID(4) = cboDriveFrame.ItemData(cboDriveFrame.ListIndex)
        'If cboDriveMeld.ListIndex > -1 Then lMinID(5) = cboDriveMeld.ItemData(cboDriveMeld.ListIndex)

        'If cboFuelComp.ListIndex > -1 Then lMinID(6) = cboFuelComp.ItemData(cboFuelComp.ListIndex)
        'If cboFuelCat.ListIndex > -1 Then lMinID(7) = cboFuelCat.ItemData(cboFuelCat.ListIndex)

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

        Dim sTemp As String = tpPowGen.PropertyValue.ToString & tpMan.PropertyValue.ToString & tpMaxSpd.PropertyValue.ToString & tpThrustGen.PropertyValue.ToString

        If sTemp.IndexOf("."c) > -1 Then
            MyBase.moUILib.AddNotification("Entries must be whole numbers.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf sTemp.IndexOf("-"c) > -1 Then
            MyBase.moUILib.AddNotification("Entries must be greater than 0.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        Dim lPower As Int32 = CInt(tpPowGen.PropertyValue)
        Dim lMan As Int32 = CInt(tpMan.PropertyValue)
        Dim lMaxSpd As Int32 = CInt(tpMaxSpd.PropertyValue)
        Dim lThrust As Int32 = CInt(tpThrustGen.PropertyValue)

        If lPower = 0 AndAlso lMan = 0 AndAlso lMaxSpd = 0 AndAlso lThrust = 0 Then
            MyBase.moUILib.AddNotification("This engine does not do anything. Enter an attribute for one of the properties.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        If lMan > mlMaxManeuver Then
            MyBase.moUILib.AddNotification("Maneuver cannot exceed " & mlMaxManeuver & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lMaxSpd > mlMaxSpeed Then
            MyBase.moUILib.AddNotification("Max Speed cannot exceed " & mlMaxSpeed & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lPower > mlMaxPowerThrust Then
            MyBase.moUILib.AddNotification("Power Production cannot exceed " & mlMaxPowerThrust & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lThrust > mlMaxPowerThrust Then
            MyBase.moUILib.AddNotification("Thrust cannot exceed " & mlMaxPowerThrust & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        If lMan > 0 AndAlso (lThrust = 0 OrElse lMaxSpd = 0) Then
            MyBase.moUILib.AddNotification("Engines cannot maneuver without thrust and speed.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If
        If lThrust > 0 AndAlso (lMan = 0 OrElse lMaxSpd = 0) Then
            MyBase.moUILib.AddNotification("Engines cannot move without maneuver and speed.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If
        If lMaxSpd > 0 AndAlso (lMan = 0 OrElse lThrust = 0) Then
            MyBase.moUILib.AddNotification("Engines cannot move without thrust and maneuver.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
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

        Dim yColorValue As Byte = 0
        If cboColor.ListIndex <> -1 Then yColorValue = CByte(cboColor.ItemData(cboColor.ListIndex))

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
        If tpStructBody.PropertyLocked = True Then lSpecMin1 = CInt(tpStructBody.PropertyValue) Else lSpecMin1 = -1
        Dim lSpecMin2 As Int32 = -1
        If tpStructFrame.PropertyLocked = True Then lSpecMin2 = CInt(tpStructFrame.PropertyValue) Else lSpecMin2 = -1
        Dim lSpecMin3 As Int32 = -1
        If tpStructMeld.PropertyLocked = True Then lSpecMin3 = CInt(tpStructMeld.PropertyValue) Else lSpecMin3 = -1
        Dim lSpecMin4 As Int32 = -1
        If tpDriveBody.PropertyLocked = True Then lSpecMin4 = CInt(tpDriveBody.PropertyValue) Else lSpecMin4 = -1
        Dim lSpecMin5 As Int32 = -1
        If tpDriveFrame.PropertyLocked = True Then lSpecMin5 = CInt(tpDriveFrame.PropertyValue) Else lSpecMin5 = -1
        Dim lSpecMin6 As Int32 = -1
        If tpDriveMeld.PropertyLocked = True Then lSpecMin6 = CInt(tpDriveMeld.PropertyValue) Else lSpecMin6 = -1

        MyBase.moUILib.GetMsgSys.SubmitEngineDesign(oResearchGuid, txtTechName.Caption, lPower, lThrust, _
          CByte(lMaxSpd), CByte(lMan), mlMineralIDs(0), mlMineralIDs(1), mlMineralIDs(2), mlMineralIDs(3), mlMineralIDs(4), mlMineralIDs(5), yColorValue, _
          lTechID, lHullReq, lPowerReq, blResCost, blResTime, blProdCost, blProdTime, lColonists, lEnlisted, lOfficers, _
          lSpecMin1, lSpecMin2, lSpecMin3, lSpecMin4, lSpecMin5, lSpecMin6, CByte(cboEngineType.ItemData(cboEngineType.ListIndex)))

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eSubmitDesignClick, ObjectType.eEngineTech, -1, -1, "")
        End If

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddNotification("Engine Prototype Design Submitted", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

        ReturnToPreviousViewAndReleaseBackground()
    End Sub

    Private Sub SetCboHelpers()
        Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder
        'Return
        'Dim bDensity As Boolean = goCurrentPlayer.PlayerKnowsProperty("DENSITY")
        'Dim bMagProd As Boolean = goCurrentPlayer.PlayerKnowsProperty("MAGNETIC PRODUCTION")
        'Dim bMeltPt As Boolean = goCurrentPlayer.PlayerKnowsProperty("MELTING POINT")
        'Dim bMalleable As Boolean = goCurrentPlayer.PlayerKnowsProperty("MALLEABLE")
        'Dim bHardness As Boolean = goCurrentPlayer.PlayerKnowsProperty("HARDNESS")
        'Dim bCompress As Boolean = goCurrentPlayer.PlayerKnowsProperty("COMPRESSIBILITY")
        'Dim bMagReact As Boolean = goCurrentPlayer.PlayerKnowsProperty("MAGNETIC REACTANCE")
        'Dim bSuperC As Boolean = goCurrentPlayer.PlayerKnowsProperty("SUPERCONDUCTIVE POINT")
        'Dim bChemReact As Boolean = goCurrentPlayer.PlayerKnowsProperty("CHEMICAL REACTANCE")
        'Dim bCombust As Boolean = goCurrentPlayer.PlayerKnowsProperty("COMBUSTIVENESS")

        ''Engines.StrBod		- MagProd Low, Density High, MeltingPt High
        oSB.Length = 0
        'If bDensity = True Then
        oSB.Append("High Density materials ")
        'Else : oSB.Append("Materials ")
        'End If
        'If bMagProd = True Then 
        oSB.Append("with low Magnetic Production ")
        'If bMeltPt = True Then 
        oSB.Append("that have a high Melting Point ")
        oSB.Append("are desired.")
        'If bDensity = True OrElse bMagProd = True OrElse bMeltPt = True Then
        stmStructBody.ToolTipText = oSB.ToString
        'Else : cboStructBody.ToolTipText = "We are not quite sure of the best way to engineer this."
        'End If

        ''Engines.StrFrame   - Malleable High, Hardness High, Compress Low
        oSB.Length = 0
        'If bHardness = True Then
        oSB.Append("Very Hard materials ")
        'Else : oSB.Append("Materials ")
        'End If
        'If bMalleable = True Then 
        oSB.Append("that are malleable ")
        'If bCompress = True Then 
        oSB.Append("with low compressibility ")
        oSB.Append("would work best.")
        'If bHardness = True OrElse bMalleable = True OrElse bCompress = True Then
        stmStructFrame.ToolTipText = oSB.ToString
        'Else : cboStructFrame.ToolTipText = "We are not quite sure of the best way to engineer this."
        'End If

        ''Engines.StrMeld		malleable low, hardness high, meltpt high
        oSB.Length = 0
        'If bMalleable = True Then
        oSB.Append("Non-Malleable materials ")
        'Else : oSB.Append("Materials ")
        'End If
        'If bHardness = True Then 
        oSB.Append("that are very hard ")
        'If bMeltPt = True Then 
        oSB.Append("with a high melting point ")
        oSB.Append("work best.")
        'If bMalleable = True OrElse bMeltPt = True OrElse bHardness = True Then
        stmStructMeld.ToolTipText = oSB.ToString
        'Else : cboStructMeld.ToolTipText = "We are not quite sure of the best way to engineer this."
        'End If

        ''Engines.DrvBody		MagReact High, Compress Low, Density Low
        oSB.Length = 0
        'If bDensity = True Then
        oSB.Append("Low Density materials ")
        'Else : oSB.Append("Materials ")
        'End If
        'If bCompress = True Then 
        oSB.Append("with low Compressibility ")
        'If bMagReact = True Then 
        oSB.Append("that have a high Magnetic Reaction ")
        oSB.Append("would be optimal in this application.")
        'If bDensity = True OrElse bCompress = True OrElse bMagReact = True Then
        stmDriveBody.ToolTipText = oSB.ToString
        'Else : cboDriveBody.ToolTipText = "We are not quite sure of the best way to engineer this."
        'End If

        ''Engines.DrvFrame	SuperC High, Hardness High, Compress Low
        oSB.Length = 0
        'If bSuperC = True Then
        oSB.Append("Materials with a high superconductive point ")
        'Else : oSB.Append("Materials ")
        'End If
        'If bHardness = True Then 
        oSB.Append("that have a high Hardness ")
        'If bCompress = True Then 
        oSB.Append("with low Compressibility ")
        oSB.Append("are preferred.")
        'If bSuperC = True OrElse bHardness = True OrElse bCompress = True Then
        stmDriveFrame.ToolTipText = oSB.ToString
        'Else : cboDriveFrame.ToolTipText = "We are not quite sure of the best way to engineer this."
        'End If

        ''Engines.DrvMeld		ChemReact High, Malleable Low, Combust Low
        oSB.Length = 0
        'If bChemReact = True Then
        oSB.Append("Highly Chemically Reactive materials ")
        'Else : oSB.Append("Materials ")
        'End If
        'If bCombust = True Then 
        oSB.Append("with low Combustive properties ")
        'If bMalleable = True Then 
        oSB.Append("that are non-Malleable ")
        oSB.Append("are ideal.")
        'If bChemReact = True OrElse bCombust = True OrElse bMalleable = True Then
        stmDriveMeld.ToolTipText = oSB.ToString
        'Else : cboDriveMeld.ToolTipText = "We are not quite sure of the best way to engineer this."
        'End If


        'Also special tech related
        If goCurrentPlayer Is Nothing = False Then
            mlMaxManeuver = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eManeuverLimit)
            mlMaxSpeed = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eEngineMaxSpeed)
            mlMaxPowerThrust = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eEnginePowerThrustLimit)
        End If

        tpThrustGen.ToolTipText = "Amount of hull that this engine will be able to move." & vbCrLf & "Maximum Value: " & mlMaxPowerThrust
        tpPowGen.ToolTipText = "Amount of power this engine will produce." & vbCrLf & "Maximum Value: " & mlMaxPowerThrust
        tpMaxSpd.ToolTipText = "Maximum speed this engine will be rated to move." & vbCrLf & "Maximum Value: " & mlMaxSpeed
        tpMan.ToolTipText = "Maneuverability rating of the engine." & vbCrLf & "Maximum Value: " & mlMaxManeuver
        tpThrustGen.MaxValue = mlMaxPowerThrust
        tpPowGen.MaxValue = mlMaxPowerThrust
        tpMaxSpd.MaxValue = mlMaxSpeed
        tpMan.MaxValue = mlMaxManeuver

        tpMan.PropertyValue = 10
        tpMaxSpd.PropertyValue = 10
        tpPowGen.SetToInitialDefault()
        tpThrustGen.SetToInitialDefault()
 
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

            MyBase.moUILib.RemoveWindow(Me.ControlName)
            ReturnToPreviousViewAndReleaseBackground() 
        End If
    End Sub

    Private Sub frmEngineBuilder_OnRender() Handles Me.OnRender
        'Ok, first, is moTech nothing?
        If moTech Is Nothing = False Then
            'Ok, now, do the values used on the form currently match the tech?
            Dim bChanged As Boolean = False
            With moTech
                bChanged = tpPowGen.PropertyValue <> .PowerProd OrElse tpThrustGen.PropertyValue <> .Thrust OrElse _
                  tpMaxSpd.PropertyValue <> .Speed OrElse tpMan.PropertyValue <> .Maneuver

                If bChanged = False Then
                    If txtTechName.Caption <> .sEngineName AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eInvalidDesign AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
                        bChanged = True
                    End If
                End If

                If bChanged = False Then
                    'Dim lID1 As Int32 = -1
                    'Dim lID2 As Int32 = -1
                    'Dim lID3 As Int32 = -1
                    'Dim lID4 As Int32 = -1
                    'Dim lID5 As Int32 = -1
                    'Dim lID6 As Int32 = -1
                    Dim lColor As Int32 = -1

                    'If cboDriveMeld.ListIndex <> -1 Then lID1 = cboDriveMeld.ItemData(cboDriveMeld.ListIndex)
                    'If cboDriveFrame.ListIndex <> -1 Then lID2 = cboDriveFrame.ItemData(cboDriveFrame.ListIndex)
                    'If cboDriveBody.ListIndex <> -1 Then lID3 = cboDriveBody.ItemData(cboDriveBody.ListIndex)
                    'If cboStructMeld.ListIndex <> -1 Then lID4 = cboStructMeld.ItemData(cboStructMeld.ListIndex)
                    'If cboStructFrame.ListIndex <> -1 Then lID5 = cboStructFrame.ItemData(cboStructFrame.ListIndex)
                    'If cboStructBody.ListIndex <> -1 Then lID6 = cboStructBody.ItemData(cboStructBody.ListIndex)
                    If cboColor.ListIndex <> -1 Then lColor = cboColor.ItemData(cboColor.ListIndex)

                    'bChanged = .lDriveMeldMineralID <> lID1 OrElse .lDriveFrameMineralID <> lID2 OrElse _
                    '  .lDriveBodyMineralID <> lID3 OrElse .lStructuralBodyMineralID <> lID6 OrElse _
                    '  .lStructuralFrameMineralID <> lID5 OrElse .lStructuralMeldMineralID <> lID4 OrElse lColor <> .ColorValue
                    bChanged = lColor <> .ColorValue OrElse .lStructuralBodyMineralID <> mlMineralIDs(0) OrElse _
                      .lStructuralFrameMineralID <> mlMineralIDs(1) OrElse .lStructuralMeldMineralID <> mlMineralIDs(2) OrElse _
                      .lDriveBodyMineralID <> mlMineralIDs(3) OrElse .lDriveFrameMineralID <> mlMineralIDs(4) OrElse _
                      .lDriveMeldMineralID <> mlMineralIDs(5) OrElse (tpStructBody.PropertyLocked <> (.lSpecifiedMin1 <> -1)) OrElse _
                          (tpStructFrame.PropertyLocked <> (.lSpecifiedMin2 <> -1)) OrElse (tpStructMeld.PropertyLocked <> (.lSpecifiedMin3 <> -1)) OrElse _
                          (tpDriveBody.PropertyLocked <> (.lSpecifiedMin4 <> -1)) OrElse (tpDriveFrame.PropertyLocked <> (.lSpecifiedMin5 <> -1)) OrElse _
                          (tpDriveMeld.PropertyLocked <> (.lSpecifiedMin6 <> -1))

                    If bChanged = False Then

                        bChanged = (.lSpecifiedMin1 <> -1 AndAlso .lSpecifiedMin1 <> tpStructBody.PropertyValue) OrElse _
                                    (.lSpecifiedMin2 <> -1 AndAlso .lSpecifiedMin2 <> tpStructFrame.PropertyValue) OrElse _
                                    (.lSpecifiedMin3 <> -1 AndAlso .lSpecifiedMin3 <> tpStructMeld.PropertyValue) OrElse _
                                    (.lSpecifiedMin4 <> -1 AndAlso .lSpecifiedMin4 <> tpDriveBody.PropertyValue) OrElse _
                                    (.lSpecifiedMin5 <> -1 AndAlso .lSpecifiedMin5 <> tpDriveFrame.PropertyValue) OrElse _
                                    (.lSpecifiedMin6 <> -1 AndAlso .lSpecifiedMin6 <> tpDriveMeld.PropertyValue)
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
        goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eEngineBuilder)
    End Sub

    Private mlMineralIDs(5) As Int32
    Private Sub SetButtonClicked(ByVal lMinIdx As Int32)
        Dim lHullTechID As Int32 = -1
        If cboEngineType.ListIndex > -1 Then
            lHullTechID = cboEngineType.ItemData(cboEngineType.ListIndex)
        End If

        Dim oTech As New EngineTechComputer
        oTech.lHullTypeID = lHullTechID

        If lMinIdx = -1 Then Return

        mlSelectedMineralIdx = -1
        oTech.MineralCBOExpanded(lMinIdx, ObjectType.eEngineTech)
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
                    stmStructBody.SetMineralName(sMinName)
                Case 1
                    stmStructFrame.SetMineralName(sMinName)
                Case 2
                    stmStructMeld.SetMineralName(sMinName)
                Case 3
                    stmDriveBody.SetMineralName(sMinName)
                Case 4
                    stmDriveFrame.SetMineralName(sMinName)
                Case 5
                    stmDriveMeld.SetMineralName(sMinName)
            End Select 
        End If

        'BuilderCostValueChange(False)
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
        If cboEngineType Is Nothing = False AndAlso cboEngineType.ListIndex > -1 Then
            lHullTypeID = cboEngineType.ItemData(cboEngineType.ListIndex)
        End If
        If bRequestDA = True AndAlso lHullTypeID <> -1 Then RequestDAValues(ObjectType.eEngineTech, 0, CByte(lHullTypeID), mlMineralIDs(0), mlMineralIDs(1), mlMineralIDs(2), mlMineralIDs(3), mlMineralIDs(4), mlMineralIDs(5), 0, 0)
    End Sub

    Private Sub tp_PropertyValueChanged(ByVal sPropName As String, ByRef ctl As ctlTechProp)
        BuilderCostValueChange(False)
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
End Class


''Interface created from Interface Builder
'Public Class frmEngineBuilder
'    Inherits UIWindow

'    Private lblPowGen As UILabel
'    Private lblThrustGen As UILabel
'    Private lblMaxSpd As UILabel
'    Private lblMan As UILabel
'    Private txtPowGen As UITextBox
'    Private txtThrust As UITextBox
'    Private txtMaxSpd As UITextBox
'    Private txtMan As UITextBox
'    Private lblStructBody As UILabel
'    Private lblStructFrame As UILabel
'    Private lblStructMeld As UILabel
'    Private lblDriveBody As UILabel
'    Private lblDriveFrame As UILabel
'    Private lblDriveMeld As UILabel
'    Private lblColor As UILabel
'    Private cboColor As UIComboBox

'    Private cboDriveMeld As UIComboBox
'    Private cboDriveFrame As UIComboBox
'    Private cboDriveBody As UIComboBox
'    Private cboStructMeld As UIComboBox
'    Private cboStructFrame As UIComboBox
'    Private cboStructBody As UIComboBox
'    Private lblTitle As UILabel

'    Private WithEvents btnDesign As UIButton
'    Private WithEvents btnCancel As UIButton
'    Private WithEvents btnDelete As UIButton
'    Private WithEvents btnHelp As UIButton

'    Private WithEvents lblTechName As UILabel
'    Private WithEvents txtTechName As UITextBox

'    Private lblDesignFlaw As UILabel

'    Private mlEntityIndex As Int32 = -1

'    Private moTech As EngineTech = Nothing
'    Private mfrmResCost As frmProdCost = Nothing
'    Private mfrmProdCost As frmProdCost = Nothing

'    Private mlMaxSpeed As Int32 = 30
'    Private mlMaxManeuver As Int32 = 20
'    Private mlMaxPowerThrust As Int32 = 10000

'    Public Sub New(ByRef oUILib As UILib)
'        MyBase.New(oUILib)

'        'frmEngineDesigner initial props
'        With Me
'            .ControlName = "frmEngineDesigner"
'            .Left = 10
'            .Top = 10
'            .Width = 420 '708
'            .Height = 520
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            .BorderColor = muSettings.InterfaceBorderColor
'            '.FillColor = System.Drawing.Color.FromArgb(128, 64, 128, 192)
'            .FillColor = muSettings.InterfaceFillColor
'            .FullScreen = True
'            .Moveable = False
'            .mbAcceptReprocessEvents = True
'        End With

'        'txtTechName initial props
'        txtTechName = New UITextBox(oUILib)
'        With txtTechName
'            .ControlName = "txtTechName"
'            .Left = 220
'            .Top = 45
'            .Width = 175
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = ""
'            .ForeColor = muSettings.InterfaceTextBoxForeColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'            .MaxLength = 20
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(txtTechName, UIControl))

'        'lblTechName initial props
'        lblTechName = New UILabel(oUILib)
'        With lblTechName
'            .ControlName = "lblTechName"
'            .Left = 15
'            .Top = 45
'            .Width = 164
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Engine Name:"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblTechName, UIControl))

'        'lblPowGen initial props
'        lblPowGen = New UILabel(oUILib)
'        With lblPowGen
'            .ControlName = "lblPowGen"
'            .Left = 15
'            .Top = 67
'            .Width = 164
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Desired Power Generation:"
'            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblPowGen, UIControl))

'        'lblThrustGen initial props
'        lblThrustGen = New UILabel(oUILib)
'        With lblThrustGen
'            .ControlName = "lblThrustGen"
'            .Left = 15
'            .Top = 89
'            .Width = 164
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Desired Thrust Generation:"
'            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblThrustGen, UIControl))

'        'lblMaxSpd initial props
'        lblMaxSpd = New UILabel(oUILib)
'        With lblMaxSpd
'            .ControlName = "lblMaxSpd"
'            .Left = 15
'            .Top = 111
'            .Width = 164
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Desired Maximum Speed:"
'            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblMaxSpd, UIControl))

'        'lblMan initial props
'        lblMan = New UILabel(oUILib)
'        With lblMan
'            .ControlName = "lblMan"
'            .Left = 15
'            .Top = 134
'            .Width = 164
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Desired Maneuverability:"
'            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblMan, UIControl))

'        'txtPowGen initial props
'        txtPowGen = New UITextBox(oUILib)
'        With txtPowGen
'            .ControlName = "txtPowGen"
'            .Left = 220
'            .Top = 67
'            .Width = 89
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "0"
'            '.ForeColor = muSettings.InterfaceBorderColor
'            .ForeColor = muSettings.InterfaceTextBoxForeColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'            '.BackColorEnabled = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'            .MaxLength = 0
'            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            .BorderColor = muSettings.InterfaceBorderColor
'            .ToolTipText = "Amount of power this engine will produce."
'            .bNumericOnly = True
'        End With
'        Me.AddChild(CType(txtPowGen, UIControl))

'        'txtThrust initial props
'        txtThrust = New UITextBox(oUILib)
'        With txtThrust
'            .ControlName = "txtThrust"
'            .Left = 220
'            .Top = 89
'            .Width = 89
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "0"
'            '.ForeColor = muSettings.InterfaceBorderColor
'            .ForeColor = muSettings.InterfaceTextBoxForeColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'            '.BackColorEnabled = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'            .MaxLength = 0
'            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            .BorderColor = muSettings.InterfaceBorderColor
'            .ToolTipText = "Amount of hull that this engine will be able to move."
'            .bNumericOnly = True
'        End With
'        Me.AddChild(CType(txtThrust, UIControl))

'        'txtMaxSpd initial props
'        txtMaxSpd = New UITextBox(oUILib)
'        With txtMaxSpd
'            .ControlName = "txtMaxSpd"
'            .Left = 220
'            .Top = 111
'            .Width = 89
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "0"
'            '.ForeColor = muSettings.InterfaceBorderColor
'            .ForeColor = muSettings.InterfaceTextBoxForeColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'            '.BackColorEnabled = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'            .MaxLength = 0
'            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            .BorderColor = muSettings.InterfaceBorderColor
'            .ToolTipText = "Maximum speed this engine will be rated to move."
'            .bNumericOnly = True
'        End With
'        Me.AddChild(CType(txtMaxSpd, UIControl))

'        'txtMan initial props
'        txtMan = New UITextBox(oUILib)
'        With txtMan
'            .ControlName = "txtMan"
'            .Left = 220
'            .Top = 133
'            .Width = 89
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "0"
'            '.ForeColor = muSettings.InterfaceBorderColor
'            .ForeColor = muSettings.InterfaceTextBoxForeColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'            '.BackColorEnabled = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'            .MaxLength = 0
'            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            .BorderColor = muSettings.InterfaceBorderColor
'            .ToolTipText = "Maneuverability rating of the engine."
'            .bNumericOnly = True
'        End With
'        Me.AddChild(CType(txtMan, UIControl))

'        'lblStructBody initial props
'        lblStructBody = New UILabel(oUILib)
'        With lblStructBody
'            .ControlName = "lblStructBody"
'            .Left = 15
'            .Top = 175
'            .Width = 200
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Select Structural Body Material:"
'            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblStructBody, UIControl))

'        'lblStructFrame initial props
'        lblStructFrame = New UILabel(oUILib)
'        With lblStructFrame
'            .ControlName = "lblStructFrame"
'            .Left = 15
'            .Top = 200
'            .Width = 200
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Select Structural Frame Material:"
'            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblStructFrame, UIControl))

'        'lblStructMeld initial props
'        lblStructMeld = New UILabel(oUILib)
'        With lblStructMeld
'            .ControlName = "lblStructMeld"
'            .Left = 15
'            .Top = 225
'            .Width = 200
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Select Structural Meld Material:"
'            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblStructMeld, UIControl))

'        'lblDriveBody initial props
'        lblDriveBody = New UILabel(oUILib)
'        With lblDriveBody
'            .ControlName = "lblDriveBody"
'            .Left = 15
'            .Top = 270
'            .Width = 200
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Select Drive Body Material:"
'            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblDriveBody, UIControl))

'        'lblDriveFrame initial props
'        lblDriveFrame = New UILabel(oUILib)
'        With lblDriveFrame
'            .ControlName = "lblDriveFrame"
'            .Left = 15
'            .Top = 295
'            .Width = 200
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Select Drive Frame Material:"
'            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblDriveFrame, UIControl))

'        'lblDriveMeld initial props
'        lblDriveMeld = New UILabel(oUILib)
'        With lblDriveMeld
'            .ControlName = "lblDriveMeld"
'            .Left = 15
'            .Top = 320
'            .Width = 200
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Select Drive Meld Material:"
'            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblDriveMeld, UIControl))

'        'lblColor initial props
'        lblColor = New UILabel(oUILib)
'        With lblColor
'            .ControlName = "lblColor"
'            .Left = 15
'            .Top = 365
'            .Width = 200
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Select Engine Color:"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblColor, UIControl))

'        'lblDesignFlaw initial props
'        lblDesignFlaw = New UILabel(oUILib)
'        With lblDesignFlaw
'            .ControlName = "lblDesignFlaw"
'            .Left = 15
'            .Top = 390
'            .Width = 400
'            .Height = 56 '18
'            .Enabled = True
'            .Visible = False
'            .Caption = ""
'            .ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
'        End With
'        Me.AddChild(CType(lblDesignFlaw, UIControl))

'        ''lblFuelComp initial props
'        'lblFuelComp = New UILabel(oUILib)
'        'With lblFuelComp
'        '    .ControlName = "lblFuelComp"
'        '    .Left = 15
'        '    .Top = 365
'        '    .Width = 200
'        '    .Height = 18
'        '    .Enabled = True
'        '    .Visible = True
'        '    .Caption = "Select Fuel Composition Material:"
'        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'        '    .ForeColor = muSettings.InterfaceBorderColor
'        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
'        '    .DrawBackImage = False
'        '    .FontFormat = DrawTextFormat.VerticalCenter
'        'End With
'        'Me.AddChild(CType(lblFuelComp, UIControl))

'        ''New Control initial props
'        'lblFuelCat = New UILabel(oUILib)
'        'With lblFuelCat
'        '    .ControlName = "lblFuelCat"
'        '    .Left = 15
'        '    .Top = 390
'        '    .Width = 200
'        '    .Height = 18
'        '    .Enabled = True
'        '    .Visible = True
'        '    .Caption = "Select Fuel Catalyst Material:"
'        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'        '    .ForeColor = muSettings.InterfaceBorderColor
'        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
'        '    .DrawBackImage = False
'        '    .FontFormat = DrawTextFormat.VerticalCenter
'        'End With
'        'Me.AddChild(CType(lblFuelCat, UIControl))

'        'btnHelp initial props
'        btnHelp = New UIButton(oUILib)
'        With btnHelp
'            .ControlName = "btnHelp"
'            .Left = Me.Width - 24 - Me.BorderLineWidth
'            .Top = Me.BorderLineWidth
'            .Width = 24
'            .Height = 24
'            .Enabled = True
'            .Visible = True
'            .Caption = "?"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 14.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
'            .ToolTipText = "Click to begin the tutorial for this window."
'        End With
'        Me.AddChild(CType(btnHelp, UIControl))

'        'btnDesign initial props
'        btnDesign = New UIButton(oUILib)
'        With btnDesign
'            .ControlName = "btnDesign"
'            .Left = 100
'            .Top = 480
'            .Width = 100
'            .Height = 32
'            .Enabled = True
'            .Visible = True
'            .Caption = "Submit Design"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
'        End With
'        Me.AddChild(CType(btnDesign, UIControl))

'        'btnCancel initial props
'        btnCancel = New UIButton(oUILib)
'        With btnCancel
'            .ControlName = "btnCancel"
'            .Left = 230
'            .Top = 480
'            .Width = 100
'            .Height = 32
'            .Enabled = True
'            .Visible = True
'            .Caption = "Cancel"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
'        End With
'        Me.AddChild(CType(btnCancel, UIControl))

'        'btnDelete initial props
'        btnDelete = New UIButton(oUILib)
'        With btnDelete
'            .ControlName = "btnDelete"
'            .Left = btnDesign.Left - 105
'            .Top = btnCancel.Top
'            .Width = 100
'            .Height = 32
'            .Enabled = True
'            .Visible = False
'            .Caption = "Delete Design"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
'        End With
'        Me.AddChild(CType(btnDelete, UIControl))

'        ''cboFuelCat initial props
'        'cboFuelCat = New UIComboBox(oUILib)
'        'With cboFuelCat
'        '    .ControlName = "cboFuelCat"
'        '    .Left = 220
'        '    .Top = 390
'        '    .Width = 175
'        '    .Height = 18
'        '    .Enabled = True
'        '    .Visible = True
'        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'        '    '.FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'        '    '.ForeColor = muSettings.InterfaceBorderColor
'        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'        '    '.HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'        '    '.DropDownListBorderColor = muSettings.InterfaceBorderColor
'        '    .BorderColor = muSettings.InterfaceBorderColor
'        '    .FillColor = muSettings.InterfaceBorderColor
'        '    .ForeColor = muSettings.InterfaceBorderColor
'        '    .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
'        'End With
'        'Me.AddChild(CType(cboFuelCat, UIControl))

'        ''cboFuelComp initial props
'        'cboFuelComp = New UIComboBox(oUILib)
'        'With cboFuelComp
'        '    .ControlName = "cboFuelComp"
'        '    .Left = 220
'        '    .Top = 365
'        '    .Width = 175
'        '    .Height = 18
'        '    .Enabled = True
'        '    .Visible = True
'        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'        '    '.FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'        '    '.ForeColor = muSettings.InterfaceBorderColor
'        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'        '    '.HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'        '    '.DropDownListBorderColor = muSettings.InterfaceBorderColor
'        '    .BorderColor = muSettings.InterfaceBorderColor
'        '    .FillColor = muSettings.InterfaceBorderColor
'        '    .ForeColor = muSettings.InterfaceBorderColor
'        '    .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
'        'End With
'        'Me.AddChild(CType(cboFuelComp, UIControl))

'        'cboColor initial props
'        cboColor = New UIComboBox(oUILib)
'        With cboColor
'            .ControlName = "cboColor"
'            .Left = 220
'            .Top = 365
'            .Width = 175
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceTextBoxFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(cboColor, UIControl))

'        'cboDriveMeld initial props
'        cboDriveMeld = New UIComboBox(oUILib)
'        With cboDriveMeld
'            .ControlName = "cboDriveMeld"
'            .Left = 220
'            .Top = 320
'            .Width = 175
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            '.FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            '.ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            '.HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            '.DropDownListBorderColor = muSettings.InterfaceBorderColor
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceTextBoxFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(cboDriveMeld, UIControl))

'        'cboDriveFrame initial props
'        cboDriveFrame = New UIComboBox(oUILib)
'        With cboDriveFrame
'            .ControlName = "cboDriveFrame"
'            .Left = 220
'            .Top = 295
'            .Width = 175
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            '.FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            '.ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            '.HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            '.DropDownListBorderColor = muSettings.InterfaceBorderColor
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceTextBoxFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(cboDriveFrame, UIControl))

'        'cboDriveBody initial props
'        cboDriveBody = New UIComboBox(oUILib)
'        With cboDriveBody
'            .ControlName = "cboDriveBody"
'            .Left = 220
'            .Top = 270
'            .Width = 175
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            '.FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            '.ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            '.HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            '.DropDownListBorderColor = muSettings.InterfaceBorderColor
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceTextBoxFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(cboDriveBody, UIControl))

'        'cboStructMeld initial props
'        cboStructMeld = New UIComboBox(oUILib)
'        With cboStructMeld
'            .ControlName = "cboStructMeld"
'            .Left = 220
'            .Top = 225
'            .Width = 175
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            '.FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            '.ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            '.HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            '.DropDownListBorderColor = muSettings.InterfaceBorderColor
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceTextBoxFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(cboStructMeld, UIControl))

'        'cboStructFrame initial props
'        cboStructFrame = New UIComboBox(oUILib)
'        With cboStructFrame
'            .ControlName = "cboStructFrame"
'            .Left = 220
'            .Top = 200
'            .Width = 175
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            '.FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            '.ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            '.HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            '.DropDownListBorderColor = muSettings.InterfaceBorderColor
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceTextBoxFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(cboStructFrame, UIControl))

'        'cboStructBody initial props
'        cboStructBody = New UIComboBox(oUILib)
'        With cboStructBody
'            .ControlName = "cboStructBody"
'            .Left = 220
'            .Top = 175
'            .Width = 175
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            '.FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            '.ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
'            '.HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
'            '.DropDownListBorderColor = muSettings.InterfaceBorderColor
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceTextBoxFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(cboStructBody, UIControl))

'        'lblTitle initial props
'        lblTitle = New UILabel(oUILib)
'        With lblTitle
'            .ControlName = "lblTitle"
'            .Left = 15
'            .Top = 10
'            .Width = 214
'            .Height = 32
'            .Enabled = True
'            .Visible = True
'            .Caption = "Engine Designer"
'            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Arial", 18, FontStyle.Bold Or FontStyle.Italic, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblTitle, UIControl))
'        FillValues()

'        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
'        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
'        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, -1)

'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'        MyBase.moUILib.AddWindow(Me)

'        glCurrentEnvirView = CurrentView.eEngineResearch

'        AddHandler cboDriveBody.DropDownExpanded, AddressOf DropDownExpanded_Event
'        AddHandler cboDriveFrame.DropDownExpanded, AddressOf DropDownExpanded_Event
'        AddHandler cboDriveMeld.DropDownExpanded, AddressOf DropDownExpanded_Event
'        AddHandler cboStructBody.DropDownExpanded, AddressOf DropDownExpanded_Event
'        AddHandler cboStructFrame.DropDownExpanded, AddressOf DropDownExpanded_Event
'        AddHandler cboStructMeld.DropDownExpanded, AddressOf DropDownExpanded_Event

'        If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
'        goFullScreenBackground = Nothing
'        goFullScreenBackground = goResMgr.GetTexture("enginebuilder.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "TechBacks.pak")
'    End Sub

'    Public Sub FillValues()
'        Dim lSorted() As Int32 = GetSortedMineralIdxArray(False)

'        'Fill our minerals/alloys
'        'cboFuelCat.Clear() : cboFuelComp.Clear()
'        cboDriveMeld.Clear() : cboDriveFrame.Clear() : cboDriveBody.Clear() : cboStructMeld.Clear() : cboStructFrame.Clear() : cboStructBody.Clear()
'        If lSorted Is Nothing = False Then
'            For X As Int32 = 0 To lSorted.GetUpperBound(0)
'                If lSorted(X) <> -1 Then
'                    'If goMinerals(lSorted(X)).IsFullyStudied = False Then Continue For
'                    'cboFuelCat.AddItem(goMinerals(X).MineralName)
'                    'cboFuelCat.ItemData(cboFuelCat.NewIndex) = goMinerals(X).ObjectID
'                    'cboFuelComp.AddItem(goMinerals(X).MineralName)
'                    'cboFuelComp.ItemData(cboFuelComp.NewIndex) = goMinerals(X).ObjectID
'                    cboDriveMeld.AddItem(goMinerals(lSorted(X)).MineralName)
'                    cboDriveMeld.ItemData(cboDriveMeld.NewIndex) = goMinerals(lSorted(X)).ObjectID
'                    cboDriveFrame.AddItem(goMinerals(lSorted(X)).MineralName)
'                    cboDriveFrame.ItemData(cboDriveFrame.NewIndex) = goMinerals(lSorted(X)).ObjectID
'                    cboDriveBody.AddItem(goMinerals(lSorted(X)).MineralName)
'                    cboDriveBody.ItemData(cboDriveBody.NewIndex) = goMinerals(lSorted(X)).ObjectID
'                    cboStructMeld.AddItem(goMinerals(lSorted(X)).MineralName)
'                    cboStructMeld.ItemData(cboStructMeld.NewIndex) = goMinerals(lSorted(X)).ObjectID
'                    cboStructFrame.AddItem(goMinerals(lSorted(X)).MineralName)
'                    cboStructFrame.ItemData(cboStructFrame.NewIndex) = goMinerals(lSorted(X)).ObjectID
'                    cboStructBody.AddItem(goMinerals(lSorted(X)).MineralName)
'                    cboStructBody.ItemData(cboStructBody.NewIndex) = goMinerals(lSorted(X)).ObjectID
'                End If
'            Next X
'        End If

'        cboColor.Clear()
'        cboColor.AddItem("Light Blue") : cboColor.ItemData(cboColor.NewIndex) = 0
'        cboColor.AddItem("Bright Green") : cboColor.ItemData(cboColor.NewIndex) = 1
'        cboColor.AddItem("Orange") : cboColor.ItemData(cboColor.NewIndex) = 2
'        cboColor.AddItem("Blue") : cboColor.ItemData(cboColor.NewIndex) = 3
'        cboColor.AddItem("Purple") : cboColor.ItemData(cboColor.NewIndex) = 4
'        cboColor.AddItem("Dark Blue") : cboColor.ItemData(cboColor.NewIndex) = 5
'        cboColor.AddItem("Red") : cboColor.ItemData(cboColor.NewIndex) = 6
'        cboColor.AddItem("Yellow") : cboColor.ItemData(cboColor.NewIndex) = 7
'        cboColor.AddItem("Dark Green") : cboColor.ItemData(cboColor.NewIndex) = 8
'        cboColor.ListIndex = 0

'        SetCboHelpers()
'    End Sub

'    Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'        ReturnToPreviousViewAndReleaseBackground()
'    End Sub

'    Private Sub ReturnToPreviousViewAndReleaseBackground()
'        MyBase.moUILib.RemoveWindow("frmMinDetail")
'        If mfrmProdCost Is Nothing = False Then MyBase.moUILib.RemoveWindow(mfrmProdCost.ControlName)
'        If mfrmResCost Is Nothing = False Then MyBase.moUILib.RemoveWindow(mfrmResCost.ControlName)
'        ReturnToPreviousView()
'        If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
'        goFullScreenBackground = Nothing
'    End Sub

'    Private Sub btnDesign_Click(ByVal sName As String) Handles btnDesign.Click
'        If goCurrentEnvir Is Nothing Then Return
'        If mlEntityIndex = -1 Then
'            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
'                If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID AndAlso goCurrentEnvir.oEntity(X).bSelected Then
'                    mlEntityIndex = X
'                    Exit For
'                End If
'            Next X
'        End If
'        If mlEntityIndex = -1 Then Return

'        If moTech Is Nothing = False AndAlso btnDesign.Caption.ToLower = "research" Then
'            'Ok, simply submit this for research now
'            Dim yData(13) As Byte
'            System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yData, 0)
'            goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yData, 2)
'            moTech.GetGUIDAsString.CopyTo(yData, 8)
'            MyBase.moUILib.GetMsgSys.SendToPrimary(yData)
'            If NewTutorialManager.TutorialOn = True Then
'                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eBuildOrderSent, moTech.ObjectID, moTech.ObjTypeID, 1, "")
'            End If
'            MyBase.moUILib.RemoveWindow(Me.ControlName)
'            ReturnToPreviousViewAndReleaseBackground()
'            Return
'        End If

'        'Check the Techname
'        If txtTechName.Caption.Trim.Length = 0 Then
'            MyBase.moUILib.AddNotification("You must specify a name for this Engine.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return
'        End If

'        Dim lMinID(5) As Int32

'        If cboStructBody.ListIndex > -1 Then lMinID(0) = cboStructBody.ItemData(cboStructBody.ListIndex)
'        If cboStructFrame.ListIndex > -1 Then lMinID(1) = cboStructFrame.ItemData(cboStructFrame.ListIndex)
'        If cboStructMeld.ListIndex > -1 Then lMinID(2) = cboStructMeld.ItemData(cboStructMeld.ListIndex)

'        If cboDriveBody.ListIndex > -1 Then lMinID(3) = cboDriveBody.ItemData(cboDriveBody.ListIndex)
'        If cboDriveFrame.ListIndex > -1 Then lMinID(4) = cboDriveFrame.ItemData(cboDriveFrame.ListIndex)
'        If cboDriveMeld.ListIndex > -1 Then lMinID(5) = cboDriveMeld.ItemData(cboDriveMeld.ListIndex)

'        'If cboFuelComp.ListIndex > -1 Then lMinID(6) = cboFuelComp.ItemData(cboFuelComp.ListIndex)
'        'If cboFuelCat.ListIndex > -1 Then lMinID(7) = cboFuelCat.ItemData(cboFuelCat.ListIndex)

'        For X As Int32 = 0 To lMinID.Length - 1
'            If lMinID(X) < 1 Then
'                MyBase.moUILib.AddNotification("You must select a material for all entries.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                Return
'            End If
'        Next X

'        If IsNumeric(txtPowGen.Caption) = False OrElse IsNumeric(txtMan.Caption) = False OrElse IsNumeric(txtMaxSpd.Caption) = False _
'          OrElse IsNumeric(txtThrust.Caption) = False Then
'            MyBase.moUILib.AddNotification("Entries must be numeric.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return
'        End If

'        Dim sTemp As String = txtPowGen.Caption & txtMan.Caption & txtMaxSpd.Caption & txtThrust.Caption

'        If sTemp.IndexOf("."c) > -1 Then
'            MyBase.moUILib.AddNotification("Entries must be whole numbers.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return
'        ElseIf sTemp.IndexOf("-"c) > -1 Then
'            MyBase.moUILib.AddNotification("Entries must be greater than 0.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return
'        End If

'        Dim lPower As Int32 = CInt(Val(txtPowGen.Caption))
'        Dim lMan As Int32 = CInt(Val(txtMan.Caption))
'        Dim lMaxSpd As Int32 = CInt(Val(txtMaxSpd.Caption))
'        Dim lThrust As Int32 = CInt(Val(txtThrust.Caption))

'        If lPower = 0 AndAlso lMan = 0 AndAlso lMaxSpd = 0 AndAlso lThrust = 0 Then
'            MyBase.moUILib.AddNotification("This engine does not do anything. Enter an attribute for one of the properties.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return
'        End If

'        If lMan > mlMaxManeuver Then
'            MyBase.moUILib.AddNotification("Maneuver cannot exceed " & mlMaxManeuver & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return
'        ElseIf lMaxSpd > mlMaxSpeed Then
'            MyBase.moUILib.AddNotification("Max Speed cannot exceed " & mlMaxSpeed & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return
'        ElseIf lPower > mlMaxPowerThrust Then
'            MyBase.moUILib.AddNotification("Power Production cannot exceed " & mlMaxPowerThrust & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return
'        ElseIf lThrust > mlMaxPowerThrust Then
'            MyBase.moUILib.AddNotification("Thrust cannot exceed " & mlMaxPowerThrust & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return
'        End If

'        If lMan > 0 AndAlso (lThrust = 0 OrElse lMaxSpd = 0) Then
'            MyBase.moUILib.AddNotification("Engines cannot maneuver without thrust and speed.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return
'        End If
'        If lThrust > 0 AndAlso (lMan = 0 OrElse lMaxSpd = 0) Then
'            MyBase.moUILib.AddNotification("Engines cannot move without maneuver and speed.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return
'        End If
'        If lMaxSpd > 0 AndAlso (lMan = 0 OrElse lThrust = 0) Then
'            MyBase.moUILib.AddNotification("Engines cannot move without thrust and maneuver.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return
'        End If

'        Dim oResearchGuid As Base_GUID = goCurrentEnvir.oEntity(mlEntityIndex)
'        'ok, if the entity production is station...
'        If goCurrentEnvir.oEntity(mlEntityIndex).yProductionType = ProductionType.eSpaceStationSpecial Then
'            'Ok, need to find research lab... try our lstFac...
'            Dim frmSelFac As frmSelectFac = CType(MyBase.moUILib.GetWindow("frmSelectFac"), frmSelectFac)
'            If frmSelFac Is Nothing = False Then
'                Dim oChild As StationChild = frmSelFac.GetCurrentChild()
'                If oChild Is Nothing = False Then
'                    oResearchGuid = New Base_GUID
'                    oResearchGuid.ObjectID = oChild.lChildID
'                    oResearchGuid.ObjTypeID = oChild.iChildTypeID
'                End If
'            End If
'        End If

'        Dim yColorValue As Byte = 0
'        If cboColor.ListIndex <> -1 Then yColorValue = CByte(cboColor.ItemData(cboColor.ListIndex))

'        Dim lTechID As Int32 = -1
'        If moTech Is Nothing = False Then lTechID = moTech.ObjectID
'        MyBase.moUILib.GetMsgSys.SubmitEngineDesign(oResearchGuid, txtTechName.Caption, lPower, lThrust, _
'          CByte(lMaxSpd), CByte(lMan), lMinID(0), lMinID(1), lMinID(2), lMinID(3), lMinID(4), lMinID(5), yColorValue, lTechID)

'        If NewTutorialManager.TutorialOn = True Then
'            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eSubmitDesignClick, ObjectType.eEngineTech, -1, -1, "")
'        End If

'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'        MyBase.moUILib.AddNotification("Engine Prototype Design Submitted", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

'        ReturnToPreviousViewAndReleaseBackground()
'    End Sub

'    'Private Sub cboDriveBody_ItemSelected(ByVal lItemIndex As Integer) Handles cboDriveBody.ItemSelected
'    '    If cboDriveBody.ListIndex > -1 Then
'    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
'    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
'    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboDriveBody.ItemData(cboDriveBody.ListIndex))
'    '    End If
'    'End Sub

'    'Private Sub cboDriveFrame_ItemSelected(ByVal lItemIndex As Integer) Handles cboDriveFrame.ItemSelected
'    '    If cboDriveFrame.ListIndex > -1 Then
'    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
'    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
'    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboDriveFrame.ItemData(cboDriveFrame.ListIndex))
'    '    End If
'    'End Sub

'    'Private Sub cboDriveMeld_ItemSelected(ByVal lItemIndex As Integer) Handles cboDriveMeld.ItemSelected
'    '    If cboDriveMeld.ListIndex > -1 Then
'    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
'    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
'    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboDriveMeld.ItemData(cboDriveMeld.ListIndex))
'    '    End If
'    'End Sub

'    'Private Sub cboStructBody_ItemSelected(ByVal lItemIndex As Integer) Handles cboStructBody.ItemSelected
'    '    If cboStructBody.ListIndex > -1 Then
'    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
'    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
'    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboStructBody.ItemData(cboStructBody.ListIndex))
'    '    End If
'    'End Sub

'    'Private Sub cboStructFrame_ItemSelected(ByVal lItemIndex As Integer) Handles cboStructFrame.ItemSelected
'    '    If cboStructFrame.ListIndex > -1 Then
'    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
'    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
'    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboStructFrame.ItemData(cboStructFrame.ListIndex))
'    '    End If
'    'End Sub

'    'Private Sub cboStructMeld_ItemSelected(ByVal lItemIndex As Integer) Handles cboStructMeld.ItemSelected
'    '    If cboStructMeld.ListIndex > -1 Then
'    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
'    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
'    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboStructMeld.ItemData(cboStructMeld.ListIndex))
'    '    End If
'    'End Sub

'    Public Sub ViewResults(ByRef oTech As EngineTech, ByVal lProdFactor As Int32)
'        If oTech Is Nothing Then Return

'        moTech = oTech

'        Me.btnDesign.Enabled = False

'        With oTech
'            Me.txtMan.Caption = .Maneuver.ToString
'            Me.txtMaxSpd.Caption = .Speed.ToString
'            Me.txtPowGen.Caption = .PowerProd.ToString
'            Me.txtTechName.Caption = .sEngineName
'            Me.txtThrust.Caption = .Thrust.ToString

'            If .lDriveBodyMineralID > 0 Then Me.cboDriveBody.FindComboItemData(.lDriveBodyMineralID)
'            If .lDriveFrameMineralID > 0 Then Me.cboDriveFrame.FindComboItemData(.lDriveFrameMineralID)
'            If .lDriveMeldMineralID > 0 Then Me.cboDriveMeld.FindComboItemData(.lDriveMeldMineralID)
'            'If .lFuelCatalystMineralID > 0 Then Me.cboFuelCat.FindComboItemData(.lFuelCatalystMineralID)
'            'If .lFuelCompositionMineralID > 0 Then Me.cboFuelComp.FindComboItemData(.lFuelCompositionMineralID)
'            If .lStructuralBodyMineralID > 0 Then Me.cboStructBody.FindComboItemData(.lStructuralBodyMineralID)
'            If .lStructuralFrameMineralID > 0 Then Me.cboStructFrame.FindComboItemData(.lStructuralFrameMineralID)
'            If .lStructuralMeldMineralID > 0 Then Me.cboStructMeld.FindComboItemData(.lStructuralMeldMineralID)

'            Me.cboColor.FindComboItemData(.ColorValue)

'            lblDesignFlaw.Caption = oTech.GetDesignFlawText()
'            lblDesignFlaw.Visible = True
'        End With

'        If oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eComponentResearching AndAlso HasAliasedRights(AliasingRights.eAddResearch) = True Then
'            btnDesign.Caption = "Research"
'            btnDesign.Enabled = True
'        End If
'        If gbAliased = False Then
'            btnDelete.Visible = oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched
'            If btnDelete.Visible = True Then
'                'Ok, set up the buttons better
'                btnDesign.Left = (Me.Width \ 2) - (btnDesign.Width \ 2)
'                btnCancel.Left = Me.Width - btnCancel.Width - 5
'                btnDelete.Left = 5
'            End If
'        End If

'        'Now... what state is the tech in?
'        If oTech.ComponentDevelopmentPhase > Base_Tech.eComponentDevelopmentPhase.eComponentDesign Then
'            'Ok, show it's research and production cost
'            If mfrmResCost Is Nothing Then mfrmResCost = New frmProdCost(MyBase.moUILib, "frmResCost")
'            With mfrmResCost
'                .Visible = True
'                .Left = Me.Left
'                .Top = Me.Top + Me.Height + 10
'                .SetFromProdCost(oTech.oResearchCost, lProdFactor, True, 0, 0)
'            End With

'            If mfrmProdCost Is Nothing Then mfrmProdCost = New frmProdCost(MyBase.moUILib, "frmProdCost")
'            With mfrmProdCost
'                .Visible = True
'                .Left = mfrmResCost.Left + mfrmResCost.Width + 10
'                .Top = mfrmResCost.Top
'                .SetFromProdCost(oTech.oProductionCost, 1000, False, oTech.HullRequired, 0)
'            End With
'        ElseIf oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
'            If mfrmResCost Is Nothing Then mfrmResCost = New frmProdCost(MyBase.moUILib, "frmResCost")
'            With mfrmResCost
'                .Visible = True
'                .Left = Me.Left
'                .Top = Me.Top + Me.Height + 10
'                .SetFromFailureCode(oTech.ErrorCodeReason)
'            End With
'        End If
'    End Sub

'    Private Sub SetCboHelpers()
'        Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder

'        Dim bDensity As Boolean = goCurrentPlayer.PlayerKnowsProperty("DENSITY")
'        Dim bMagProd As Boolean = goCurrentPlayer.PlayerKnowsProperty("MAGNETIC PRODUCTION")
'        Dim bMeltPt As Boolean = goCurrentPlayer.PlayerKnowsProperty("MELTING POINT")
'        Dim bMalleable As Boolean = goCurrentPlayer.PlayerKnowsProperty("MALLEABLE")
'        Dim bHardness As Boolean = goCurrentPlayer.PlayerKnowsProperty("HARDNESS")
'        Dim bCompress As Boolean = goCurrentPlayer.PlayerKnowsProperty("COMPRESSIBILITY")
'        Dim bMagReact As Boolean = goCurrentPlayer.PlayerKnowsProperty("MAGNETIC REACTANCE")
'        Dim bSuperC As Boolean = goCurrentPlayer.PlayerKnowsProperty("SUPERCONDUCTIVE POINT")
'        Dim bChemReact As Boolean = goCurrentPlayer.PlayerKnowsProperty("CHEMICAL REACTANCE")
'        Dim bCombust As Boolean = goCurrentPlayer.PlayerKnowsProperty("COMBUSTIVENESS")

'        ''Engines.StrBod		- MagProd Low, Density High, MeltingPt High
'        oSB.Length = 0
'        'If bDensity = True Then
'        oSB.Append("High Density materials ")
'        'Else : oSB.Append("Materials ")
'        'End If
'        'If bMagProd = True Then 
'        oSB.Append("with low Magnetic Production ")
'        'If bMeltPt = True Then 
'        oSB.Append("that have a high Melting Point ")
'        oSB.Append("are desired.")
'        'If bDensity = True OrElse bMagProd = True OrElse bMeltPt = True Then
'        cboStructBody.ToolTipText = oSB.ToString
'        'Else : cboStructBody.ToolTipText = "We are not quite sure of the best way to engineer this."
'        'End If

'        ''Engines.StrFrame   - Malleable High, Hardness High, Compress Low
'        oSB.Length = 0
'        'If bHardness = True Then
'        oSB.Append("Very Hard materials ")
'        'Else : oSB.Append("Materials ")
'        'End If
'        'If bMalleable = True Then 
'        oSB.Append("that are malleable ")
'        'If bCompress = True Then 
'        oSB.Append("with low compressibility ")
'        oSB.Append("would work best.")
'        'If bHardness = True OrElse bMalleable = True OrElse bCompress = True Then
'        cboStructFrame.ToolTipText = oSB.ToString
'        'Else : cboStructFrame.ToolTipText = "We are not quite sure of the best way to engineer this."
'        'End If

'        ''Engines.StrMeld		malleable low, hardness high, meltpt high
'        oSB.Length = 0
'        'If bMalleable = True Then
'        oSB.Append("Non-Malleable materials ")
'        'Else : oSB.Append("Materials ")
'        'End If
'        'If bHardness = True Then 
'        oSB.Append("that are very hard ")
'        'If bMeltPt = True Then 
'        oSB.Append("with a high melting point ")
'        oSB.Append("work best.")
'        'If bMalleable = True OrElse bMeltPt = True OrElse bHardness = True Then
'        cboStructMeld.ToolTipText = oSB.ToString
'        'Else : cboStructMeld.ToolTipText = "We are not quite sure of the best way to engineer this."
'        'End If

'        ''Engines.DrvBody		MagReact High, Compress Low, Density Low
'        oSB.Length = 0
'        'If bDensity = True Then
'        oSB.Append("Low Density materials ")
'        'Else : oSB.Append("Materials ")
'        'End If
'        'If bCompress = True Then 
'        oSB.Append("with low Compressibility ")
'        'If bMagReact = True Then 
'        oSB.Append("that have a high Magnetic Reaction ")
'        oSB.Append("would be optimal in this application.")
'        'If bDensity = True OrElse bCompress = True OrElse bMagReact = True Then
'        cboDriveBody.ToolTipText = oSB.ToString
'        'Else : cboDriveBody.ToolTipText = "We are not quite sure of the best way to engineer this."
'        'End If

'        ''Engines.DrvFrame	SuperC High, Hardness High, Compress Low
'        oSB.Length = 0
'        'If bSuperC = True Then
'        oSB.Append("Materials with a high superconductive point ")
'        'Else : oSB.Append("Materials ")
'        'End If
'        'If bHardness = True Then 
'        oSB.Append("that have a high Hardness ")
'        'If bCompress = True Then 
'        oSB.Append("with low Compressibility ")
'        oSB.Append("are preferred.")
'        'If bSuperC = True OrElse bHardness = True OrElse bCompress = True Then
'        cboDriveFrame.ToolTipText = oSB.ToString
'        'Else : cboDriveFrame.ToolTipText = "We are not quite sure of the best way to engineer this."
'        'End If

'        ''Engines.DrvMeld		ChemReact High, Malleable Low, Combust Low
'        oSB.Length = 0
'        'If bChemReact = True Then
'        oSB.Append("Highly Chemically Reactive materials ")
'        'Else : oSB.Append("Materials ")
'        'End If
'        'If bCombust = True Then 
'        oSB.Append("with low Combustive properties ")
'        'If bMalleable = True Then 
'        oSB.Append("that are non-Malleable ")
'        oSB.Append("are ideal.")
'        'If bChemReact = True OrElse bCombust = True OrElse bMalleable = True Then
'        cboDriveMeld.ToolTipText = oSB.ToString
'        'Else : cboDriveMeld.ToolTipText = "We are not quite sure of the best way to engineer this."
'        'End If


'        'Also special tech related
'        mlMaxManeuver = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eManeuverLimit)
'        mlMaxSpeed = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eEngineMaxSpeed)
'        mlMaxPowerThrust = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eEnginePowerThrustLimit)
'        txtThrust.ToolTipText = "Amount of hull that this engine will be able to move." & vbCrLf & "Maximum Value: " & mlMaxPowerThrust
'        txtPowGen.ToolTipText = "Amount of power this engine will produce." & vbCrLf & "Maximum Value: " & mlMaxPowerThrust
'        txtMaxSpd.ToolTipText = "Maximum speed this engine will be rated to move." & vbCrLf & "Maximum Value: " & mlMaxSpeed
'        txtMan.ToolTipText = "Maneuverability rating of the engine." & vbCrLf & "Maximum Value: " & mlMaxManeuver
'    End Sub

'    Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
'        If moTech Is Nothing Then Return
'        If btnDelete.Caption = "Delete Design" Then
'            btnDelete.Caption = "CONFIRM"
'        Else
'            'Delete the design
'            Dim yMsg(7) As Byte
'            System.BitConverter.GetBytes(GlobalMessageCode.eDeleteDesign).CopyTo(yMsg, 0)
'            moTech.GetGUIDAsString.CopyTo(yMsg, 2)
'            MyBase.moUILib.SendMsgToPrimary(yMsg)
'            MyBase.moUILib.RemoveWindow(Me.ControlName)
'            ReturnToPreviousViewAndReleaseBackground()
'        End If
'    End Sub

'    Private Sub frmEngineBuilder_OnRender() Handles Me.OnRender
'        'Ok, first, is moTech nothing?
'        If moTech Is Nothing = False Then
'            'Ok, now, do the values used on the form currently match the tech?
'            Dim bChanged As Boolean = False
'            With moTech
'                bChanged = Val(txtPowGen.Caption) <> .PowerProd OrElse Val(txtThrust.Caption) <> .Thrust OrElse _
'                  Val(txtMaxSpd.Caption) <> .Speed OrElse Val(txtMan.Caption) <> .Maneuver

'                If bChanged = False Then
'                    Dim lID1 As Int32 = -1
'                    Dim lID2 As Int32 = -1
'                    Dim lID3 As Int32 = -1
'                    Dim lID4 As Int32 = -1
'                    Dim lID5 As Int32 = -1
'                    Dim lID6 As Int32 = -1
'                    Dim lColor As Int32 = -1

'                    If cboDriveMeld.ListIndex <> -1 Then lID1 = cboDriveMeld.ItemData(cboDriveMeld.ListIndex)
'                    If cboDriveFrame.ListIndex <> -1 Then lID2 = cboDriveFrame.ItemData(cboDriveFrame.ListIndex)
'                    If cboDriveBody.ListIndex <> -1 Then lID3 = cboDriveBody.ItemData(cboDriveBody.ListIndex)
'                    If cboStructMeld.ListIndex <> -1 Then lID4 = cboStructMeld.ItemData(cboStructMeld.ListIndex)
'                    If cboStructFrame.ListIndex <> -1 Then lID5 = cboStructFrame.ItemData(cboStructFrame.ListIndex)
'                    If cboStructBody.ListIndex <> -1 Then lID6 = cboStructBody.ItemData(cboStructBody.ListIndex)
'                    If cboColor.ListIndex <> -1 Then lColor = cboColor.ItemData(cboColor.ListIndex)

'                    bChanged = .lDriveMeldMineralID <> lID1 OrElse .lDriveFrameMineralID <> lID2 OrElse _
'                      .lDriveBodyMineralID <> lID3 OrElse .lStructuralBodyMineralID <> lID6 OrElse _
'                      .lStructuralFrameMineralID <> lID5 OrElse .lStructuralMeldMineralID <> lID4 OrElse lColor <> .ColorValue
'                End If

'                If bChanged = True Then
'                    btnDesign.Caption = "Redesign"
'                    btnDesign.Enabled = True
'                Else
'                    btnDesign.Caption = "Research"
'                    If moTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eInvalidDesign Then btnDesign.Enabled = False
'                End If
'            End With
'        End If
'    End Sub

'    Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
'        If goTutorial Is Nothing Then goTutorial = New TutorialManager()
'        goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eEngineBuilder)
'    End Sub

'    Private Sub DropDownExpanded_Event(ByVal sComboBoxName As String)
'        Try
'            Dim ofrmMin As frmMinDetail = CType(MyBase.moUILib.GetWindow("frmMinDetail"), frmMinDetail)
'            If ofrmMin Is Nothing Then Return

'            Select Case sComboBoxName
'                Case cboDriveBody.ControlName
'                    ofrmMin.ClearHighlights()
'                    For X As Int32 = 0 To glMineralPropertyUB
'                        If glMineralPropertyIdx(X) <> -1 Then
'                            Select Case goMineralProperty(X).MineralPropertyName.ToUpper
'                                Case "DENSITY"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
'                                Case "COMPRESSIBILITY"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
'                                Case "MAGNETIC REACTANCE"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 2, 7)
'                            End Select
'                        End If
'                    Next X
'                Case cboDriveFrame.ControlName
'                    ofrmMin.ClearHighlights()
'                    For X As Int32 = 0 To glMineralPropertyUB
'                        If glMineralPropertyIdx(X) <> -1 Then
'                            Select Case goMineralProperty(X).MineralPropertyName.ToUpper
'                                Case "SUPERCONDUCTIVE POINT"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 2, 7)
'                                Case "COMPRESSIBILITY"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
'                                Case "HARDNESS"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 6)
'                            End Select
'                        End If
'                    Next X
'                Case cboDriveMeld.ControlName
'                    ofrmMin.ClearHighlights()
'                    For X As Int32 = 0 To glMineralPropertyUB
'                        If glMineralPropertyIdx(X) <> -1 Then
'                            Select Case goMineralProperty(X).MineralPropertyName.ToUpper
'                                Case "CHEMICAL REACTANCE"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 6)
'                                Case "COMBUSTIVENESS"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
'                                Case "MALLEABLE"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
'                            End Select
'                        End If
'                    Next X
'                Case cboStructBody.ControlName
'                    ofrmMin.ClearHighlights()
'                    For X As Int32 = 0 To glMineralPropertyUB
'                        If glMineralPropertyIdx(X) <> -1 Then
'                            Select Case goMineralProperty(X).MineralPropertyName.ToUpper
'                                Case "DENSITY"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 6)
'                                Case "MAGNETIC PRODUCTION"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
'                                Case "MELTING POINT"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 6)
'                            End Select
'                        End If
'                    Next X
'                Case cboStructFrame.ControlName
'                    ofrmMin.ClearHighlights()
'                    For X As Int32 = 0 To glMineralPropertyUB
'                        If glMineralPropertyIdx(X) <> -1 Then
'                            Select Case goMineralProperty(X).MineralPropertyName.ToUpper
'                                Case "HARDNESS"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 6)
'                                Case "MALLEABLE"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 5)
'                                Case "COMPRESSIBILITY"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
'                            End Select
'                        End If
'                    Next X
'                Case cboStructMeld.ControlName
'                    ofrmMin.ClearHighlights()
'                    For X As Int32 = 0 To glMineralPropertyUB
'                        If glMineralPropertyIdx(X) <> -1 Then
'                            Select Case goMineralProperty(X).MineralPropertyName.ToUpper
'                                Case "HARDNESS"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 6)
'                                Case "MALLEABLE"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
'                                Case "MELTING POINT"
'                                    ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 6)
'                            End Select
'                        End If
'                    Next X
'            End Select
'        Catch
'        End Try
'    End Sub
'End Class