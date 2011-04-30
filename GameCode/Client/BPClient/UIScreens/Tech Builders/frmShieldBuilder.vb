Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmShieldBuilder
    Inherits frmTechBuilder
    'Inherits UIWindow

    Private lblCoilMat As UILabel
    Private lblAccMat As UILabel
    Private lblCaseMat As UILabel
    Private lblColor As UILabel

    'Private cboCasingMat As UIComboBox
    'Private cboAccMat As UIComboBox
    'Private cboCoilMat As UIComboBox
    Private stmCoilMat As ctlSetTechMineral
    Private stmAccMat As ctlSetTechMineral
    Private stmCasingMat As ctlSetTechMineral

    Private cboColor As UIComboBox

    Private lblHullType As UILabel
    Private cboHullType As UIComboBox

    'Private lblMaxHP As UILabel
    'Private lblRecRate As UILabel
    'Private lblRecInt As UILabel
    'Private lblProjHull As UILabel
    'Private txtProjectionHullSize As UITextBox
    'Private txtRechargeInterval As UITextBox
    'Private txtRechargeRate As UITextBox
    'Private txtMaxHP As UITextBox
    Private tpMaxHP As ctlTechProp
    Private tpRecRate As ctlTechProp
    Private tpRecInt As ctlTechProp
    Private tpProjHull As ctlTechProp

    Private tpCoilMat As ctlTechProp
    Private tpAccMat As ctlTechProp
    Private tpCasingMat As ctlTechProp

    Private lblTitle As UILabel

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

    Private moTech As ShieldTech = Nothing
    Private mfrmResCost As frmProdCost = Nothing
    Private mfrmProdCost As frmProdCost = Nothing

    Private mlMaxHP As Int32 = 50
    Private mlMaxProjSize As Int32 = 500
    Private mlMinRechInt As Int32 = 150
    Private mlMaxRechRate As Int32 = 50
    Private mbImpossibleDesign As Boolean = False

    Private mbIgnoreValueChange As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmShieldBuilder initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eShieldBuilder
            .ControlName = "frmShieldBuilder"
            .Left = 14
            .Top = 13
            .Width = 490 '380
            .Height = 410
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
            .Left = 190
            .Top = 55
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
            .Top = 55
            .Width = 164
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Shield Name:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTechName, UIControl))

        'lblHullType initial props
        lblHullType = New UILabel(oUILib)
        With lblHullType
            .ControlName = "lblHullType"
            .Left = 15
            .Top = 80
            .Width = 164
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hull Type:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblHullType, UIControl))

        'lblCoilMat initial props
        lblCoilMat = New UILabel(oUILib)
        With lblCoilMat
            .ControlName = "lblCoilMat"
            .Left = 15
            .Top = 205 ' 200
            .Width = 70 '167
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Coil:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblCoilMat, UIControl))

        'lblAccMat initial props
        lblAccMat = New UILabel(oUILib)
        With lblAccMat
            .ControlName = "lblAccMat"
            .Left = 15
            .Top = 230 '225
            .Width = 70 '167
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Accelerator:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblAccMat, UIControl))

        'lblCaseMat initial props
        lblCaseMat = New UILabel(oUILib)
        With lblCaseMat
            .ControlName = "lblCaseMat"
            .Left = 15
            .Top = 255 ' 250
            .Width = 70 '167
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
        Me.AddChild(CType(lblCaseMat, UIControl))

         
        tpProjHull = New ctlTechProp(oUILib)
        With tpProjHull
            .ControlName = "tpProjHull"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 255
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Projection Size:"
            .PropertyValue = 0
            .Top = 180
            .Visible = True
            '.Width = 512
            .ToolTipText = "Indicates the maximum size that the hull projector" & vbCrLf & "extends the shield system's protective field" & vbCrLf & "around the entity that the system is placed on." & vbCrLf & "This system cannot be placed on units that have" & vbCrLf & "higher hull size than this value."
        End With
        Me.AddChild(CType(tpProjHull, UIControl))

        
        tpRecInt = New ctlTechProp(oUILib)
        With tpRecInt
            .ControlName = "tpRecInt"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 1500
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Recharge Interval:"
            .PropertyValue = 0
            .Top = 155
            .Visible = True
            '.Width = 512
            .bNoMaxValue = True
            .SetDivisor(30)
            .ToolTipText = "The number of seconds between pulses projected" & vbCrLf & "by the shield system to replenish lost energy."
        End With
        Me.AddChild(CType(tpRecInt, UIControl))

       
        tpRecRate = New ctlTechProp(oUILib)
        With tpRecRate
            .ControlName = "tpRecRate"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 255
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Recharge Rate:"
            .PropertyValue = 0
            .Top = 130
            .Visible = True
            '.Width = 512
            .ToolTipText = "The amount of energy that the shield system" & vbCrLf & "can replenish in a given replenish pulse."
        End With
        Me.AddChild(CType(tpRecRate, UIControl))

       
        tpMaxHP = New ctlTechProp(oUILib)
        With tpMaxHP
            .ControlName = "tpMaxHP"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 255
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Maximum Hitpoints:"
            .PropertyValue = 0
            .Top = 105
            .Visible = True
            '.Width = 512
            .ToolTipText = "The maximum damage the shield system can withstand."
        End With
        Me.AddChild(CType(tpMaxHP, UIControl))

        'NewControl4 initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 15
            .Top = 10
            .Width = 187
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Shield Designer"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Arial", 18, FontStyle.Bold Or FontStyle.Italic, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblTitle, UIControl))

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
            .Left = (Me.Width \ 2) - 50 '110
            .Top = 365 '355
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
            .Left = Me.Width - 110 '270
            .Top = btnDesign.Top
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

        'lblColor initial props
        lblColor = New UILabel(oUILib)
        With lblColor
            .ControlName = "lblColor"
            .Left = 15
            .Top = 285
            .Width = 167
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Select Shield Color:"
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
            .Top = lblColor.Top + lblColor.Height + 5
            .Width = Me.Width - (.Left * 2)
            .Height = 56 '18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(lblDesignFlaw, UIControl))

        tpCasingMat = New ctlTechProp(oUILib)
        With tpCasingMat
            .ControlName = "tpCasingMat"
            .Enabled = True
            .Height = 18
            .Left = lblCaseMat.Left + lblCaseMat.Width + 185    '10 for left/right, 175 for cbo
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Casing:"
            .PropertyValue = 0
            .Top = lblCaseMat.Top
            .Visible = True
            .Width = Me.Width - .Left - 15
            .SetNoPropDisplay(True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpCasingMat, UIControl))

        tpAccMat = New ctlTechProp(oUILib)
        With tpAccMat
            .ControlName = "tpAccMat"
            .Enabled = True
            .Height = 18
            .Left = lblCaseMat.Left + lblCaseMat.Width + 185    '10 for left/right, 175 for cbo
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Accelerator:"
            .PropertyValue = 0
            .Top = lblAccMat.Top
            .Visible = True
            .Width = tpCasingMat.Width
            .SetNoPropDisplay(True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpAccMat, UIControl))

        tpCoilMat = New ctlTechProp(oUILib)
        With tpCoilMat
            .ControlName = "tpCoilMat"
            .Enabled = True
            .Height = 18
            .Left = lblCaseMat.Left + lblCaseMat.Width + 185    '10 for left/right, 175 for cbo
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Coil:"
            .PropertyValue = 0
            .Top = lblCoilMat.Top
            .Visible = True
            .Width = tpCasingMat.Width
            .SetNoPropDisplay(True)
            .SetPercOfPaymentVisibility(True)
        End With
        Me.AddChild(CType(tpCoilMat, UIControl))

        'cboColor initial props
        cboColor = New UIComboBox(oUILib)
        With cboColor
            .ControlName = "cboColor"
            .Left = 190
            .Top = 285
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

        ''lblCoilMatitem initial props
        'lblCoilMatItem = New UILabel(oUILib)
        'With lblCoilMatItem
        '    .ControlName = "lblCoilMatitem"
        '    .Left = lblCaseMat.Left + lblCaseMat.Width + 5
        '    .Top = lblCaseMat.Top
        '    .Width = 125
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Unselected"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        'End With
        'Me.AddChild(CType(lblCoilMatItem, UIControl))
        ''btnSetCoilMatItem initial props
        'btnSetCoilMatItem = New UIButton(oUILib)
        'With btnSetCoilMatItem
        '    .ControlName = "btnSetCoilMatItem"
        '    .Left = lblCoilMatItem.Left + lblCoilMatItem.Width + 1
        '    .Top = lblCoilMatItem.Top
        '    .Width = 50
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = False
        '    .Caption = "Set"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        'End With
        'Me.AddChild(CType(btnSetCoilMatItem, UIControl))
        stmCoilMat = New ctlSetTechMineral(oUILib)
        With stmCoilMat
            .ControlName = "stmCoilMat"
            .Left = lblCoilMat.Left + lblCoilMat.Width + 5
            .Top = lblCoilMat.Top
            .Width = 175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 0
        End With
        Me.AddChild(CType(stmCoilMat, UIControl))

        ''lblAccMatItem initial props
        'lblAccMatItem = New UILabel(oUILib)
        'With lblAccMatItem
        '    .ControlName = "lblAccMatItem"
        '    .Left = lblCoilMatItem.Left
        '    .Top = lblAccMat.Top
        '    .Width = 125
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Unselected"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        'End With
        'Me.AddChild(CType(lblAccMatItem, UIControl))
        ''btnSetAccMatItem initial props
        'btnSetAccMatItem = New UIButton(oUILib)
        'With btnSetAccMatItem
        '    .ControlName = "btnSetAccMatItem"
        '    .Left = lblAccMatItem.Left + lblAccMatItem.Width + 1
        '    .Top = lblAccMatItem.Top
        '    .Width = 50
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = False
        '    .Caption = "Set"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        'End With
        'Me.AddChild(CType(btnSetAccMatItem, UIControl))
        stmAccMat = New ctlSetTechMineral(oUILib)
        With stmAccMat
            .ControlName = "stmAccMat"
            .Left = stmCoilMat.Left
            .Top = lblAccMat.Top
            .Width = 175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 1
        End With
        Me.AddChild(CType(stmAccMat, UIControl))

        ''lblCaseMatItem initial props
        'lblCaseMatItem = New UILabel(oUILib)
        'With lblCaseMatItem
        '    .ControlName = "lblCaseMatItem"
        '    .Left = lblCoilMatItem.Left
        '    .Top = lblCaseMat.Top
        '    .Width = 125
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Unselected"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        'End With
        'Me.AddChild(CType(lblCaseMatItem, UIControl))
        ''btnSetCaseMatItem initial props
        'btnSetCaseMatItem = New UIButton(oUILib)
        'With btnSetCaseMatItem
        '    .ControlName = "btnSetCaseMatItem"
        '    .Left = lblCaseMatItem.Left + lblCaseMatItem.Width + 1
        '    .Top = lblCaseMatItem.Top
        '    .Width = 50
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = False
        '    .Caption = "Set"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        'End With
        'Me.AddChild(CType(btnSetCaseMatItem, UIControl))
        stmCasingMat = New ctlSetTechMineral(oUILib)
        With stmCasingMat
            .ControlName = "stmCasingMat"
            .Left = stmCoilMat.Left
            .Top = lblCaseMat.Top
            .Width = 175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 2
        End With
        Me.AddChild(CType(stmCasingMat, UIControl))

        ''cboCasingMat initial props
        'cboCasingMat = New UIComboBox(oUILib)
        'With cboCasingMat
        '    .ControlName = "cboCasingMat"
        '    .Left = lblCaseMat.Left + lblCaseMat.Width + 5
        '    .Top = lblCaseMat.Top
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
        'Me.AddChild(CType(cboCasingMat, UIControl))

        ''cboAccMat initial props
        'cboAccMat = New UIComboBox(oUILib)
        'With cboAccMat
        '    .ControlName = "cboAccMat"
        '    .Left = cboCasingMat.Left
        '    .Top = lblAccMat.Top
        '    .Width = 175
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    '.FillColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    '.ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    ' .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    '.DropDownListBorderColor = muSettings.InterfaceBorderColor
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .FillColor = muSettings.InterfaceTextBoxFillColor
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
        '    .mbAcceptReprocessEvents = True
        'End With
        'Me.AddChild(CType(cboAccMat, UIControl))

        ''cboCoilMat initial props
        'cboCoilMat = New UIComboBox(oUILib)
        'With cboCoilMat
        '    .ControlName = "cboCoilMat"
        '    .Left = cboCasingMat.Left
        '    .Top = lblCoilMat.Top
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
        'Me.AddChild(CType(cboCoilMat, UIControl))

        'cboHullType initial props
        cboHullType = New UIComboBox(oUILib)
        With cboHullType
            .ControlName = "cboHullType"
            .Left = txtTechName.Left
            .Top = lblHullType.Top
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
        Me.AddChild(CType(cboHullType, UIControl))

        FillValues()
        EnableDisableControls()
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
        ofrm.ShowMineralDetail(Me.Left + Me.Width + 5, Me.Top, 512, -1)

        glCurrentEnvirView = CurrentView.eShieldResearch

        AddHandler stmAccMat.SetButtonClicked, AddressOf SetButtonClicked
        AddHandler stmCasingMat.SetButtonClicked, AddressOf SetButtonClicked
        AddHandler stmCoilMat.SetButtonClicked, AddressOf SetButtonClicked

        AddHandler cboHullType.ItemSelected, AddressOf ComboValueChanged

        AddHandler tpAccMat.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpCasingMat.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpCoilMat.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpMaxHP.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpProjHull.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpRecInt.PropertyValueChanged, AddressOf tp_ValueChanged
        AddHandler tpRecRate.PropertyValueChanged, AddressOf tp_ValueChanged

        AddHandler tpMaxHP.OnLostFocus, AddressOf tpMaxHP_OnLostFocus
        AddHandler tpRecRate.OnLostFocus, AddressOf tpMaxHP_OnLostFocus

        If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
        goFullScreenBackground = Nothing
        goFullScreenBackground = goResMgr.GetTexture("shieldbuilder.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "TechBacks.pak")
    End Sub

    Public Sub FillValues()
        'Dim lSorted() As Int32 = GetSortedMineralIdxArray(False)

        ''Fill our minerals/alloys
        'cboCasingMat.Clear() : cboAccMat.Clear() : cboCoilMat.Clear()
        'If lSorted Is Nothing = False Then
        '    For X As Int32 = 0 To lSorted.GetUpperBound(0)
        '        If lSorted(X) <> -1 Then
        '            'If goMinerals(lSorted(X)).IsFullyStudied = False Then Continue For
        '            cboCasingMat.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboCasingMat.ItemData(cboCasingMat.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '            cboAccMat.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboAccMat.ItemData(cboAccMat.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '            cboCoilMat.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboCoilMat.ItemData(cboCoilMat.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '        End If
        '    Next X
        'End If

        cboColor.Clear()
        cboColor.AddItem("Teal") : cboColor.ItemData(cboColor.NewIndex) = 0
        cboColor.AddItem("White") : cboColor.ItemData(cboColor.NewIndex) = 1
        cboColor.AddItem("Yellow") : cboColor.ItemData(cboColor.NewIndex) = 2
        cboColor.AddItem("Orange") : cboColor.ItemData(cboColor.NewIndex) = 3
        cboColor.AddItem("Blue") : cboColor.ItemData(cboColor.NewIndex) = 4
        cboColor.AddItem("Red") : cboColor.ItemData(cboColor.NewIndex) = 5
        cboColor.AddItem("Purple") : cboColor.ItemData(cboColor.NewIndex) = 6

        cboColor.ListIndex = 0

        SetCboHelpers()

        cboHullType.Clear()
        Dim lMaxPowerThrustProjSize As Int32 = 0
        If goCurrentPlayer Is Nothing = False Then
            lMaxPowerThrustProjSize = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eEnginePowerThrustLimit)
            lMaxPowerThrustProjSize = Math.Min(lMaxPowerThrustProjSize, goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eShieldProjSize))
        End If
        Dim lMinHull As Int32 = 0
        Dim lMaxHull As Int32 = 0
        HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.BattleCruiser, lMinHull, lMaxHull)
        If lMinHull < lMaxPowerThrustProjSize Then
            cboHullType.AddItem("Battlecruiser") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.BattleCruiser
        End If
        HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.Battleship, lMinHull, lMaxHull)
        If lMinHull < lMaxPowerThrustProjSize Then
            cboHullType.AddItem("Battleship") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Battleship
        End If
        HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.Corvette, lMinHull, lMaxHull)
        If lMinHull < lMaxPowerThrustProjSize Then
            cboHullType.AddItem("Corvette") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Corvette
        End If
        HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.Cruiser, lMinHull, lMaxHull)
        If lMinHull < lMaxPowerThrustProjSize Then
            cboHullType.AddItem("Cruiser") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Cruiser
        End If
        HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.Destroyer, lMinHull, lMaxHull)
        If lMinHull < lMaxPowerThrustProjSize Then
            cboHullType.AddItem("Destroyer") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Destroyer
        End If

        HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.Escort, lMinHull, lMaxHull)
        If lMinHull < lMaxPowerThrustProjSize Then
            cboHullType.AddItem("Escort") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Escort
        End If
        HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.Facility, lMinHull, lMaxHull)
        If lMinHull < lMaxPowerThrustProjSize Then
            cboHullType.AddItem("Facility") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Facility
        End If
        HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.LightFighter, lMinHull, lMaxHull)
        If lMinHull < lMaxPowerThrustProjSize Then
            cboHullType.AddItem("Fighter (Light)") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.LightFighter
        End If
        HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.MediumFighter, lMinHull, lMaxHull)
        If lMinHull < lMaxPowerThrustProjSize Then
            cboHullType.AddItem("Fighter (Medium)") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.MediumFighter
        End If
        HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.HeavyFighter, lMinHull, lMaxHull)
        If lMinHull < lMaxPowerThrustProjSize Then
            cboHullType.AddItem("Fighter (Heavy)") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.HeavyFighter
        End If
        HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.Frigate, lMinHull, lMaxHull)
        If lMinHull < lMaxPowerThrustProjSize Then
            cboHullType.AddItem("Frigate") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Frigate
        End If

        Dim bHasNavalUnit As Boolean = False
        If goCurrentPlayer Is Nothing = False Then bHasNavalUnit = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eNavalAbility) > 0
        If bHasNavalUnit = True Then
            HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.NavalBattleship, lMinHull, lMaxHull)
            If lMinHull < lMaxPowerThrustProjSize Then
                cboHullType.AddItem("Naval Battleship") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalBattleship
            End If
            HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.NavalCarrier, lMinHull, lMaxHull)
            If lMinHull < lMaxPowerThrustProjSize Then
                cboHullType.AddItem("Naval Carrier") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalCarrier
            End If
            HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.NavalCruiser, lMinHull, lMaxHull)
            If lMinHull < lMaxPowerThrustProjSize Then
                cboHullType.AddItem("Naval Cruiser") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalCruiser
            End If
            HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.NavalDestroyer, lMinHull, lMaxHull)
            If lMinHull < lMaxPowerThrustProjSize Then
                cboHullType.AddItem("Naval Destroyer") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalDestroyer
            End If
            HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.NavalFrigate, lMinHull, lMaxHull)
            If lMinHull < lMaxPowerThrustProjSize Then
                cboHullType.AddItem("Naval Frigate") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalFrigate
            End If
            HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.NavalSub, lMinHull, lMaxHull)
            If lMinHull < lMaxPowerThrustProjSize Then
                cboHullType.AddItem("Naval Submarine") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.NavalSub
            End If
            HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.Utility, lMinHull, lMaxHull)
            If lMinHull < lMaxPowerThrustProjSize Then
                cboHullType.AddItem("Naval Utility") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Utility
            End If
        End If
        HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.SmallVehicle, lMinHull, lMaxHull)
        If lMinHull < lMaxPowerThrustProjSize Then
            cboHullType.AddItem("Small Vehicle") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.SmallVehicle
        End If
        HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.SpaceStation, lMinHull, lMaxHull)
        If lMinHull < lMaxPowerThrustProjSize Then
            cboHullType.AddItem("Space Station") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.SpaceStation
        End If
        HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.Tank, lMinHull, lMaxHull)
        If lMinHull < lMaxPowerThrustProjSize Then
            cboHullType.AddItem("Tank") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Tank
        End If
        HullTech.GetMinMaxHullOfHullType(Base_Tech.eyHullType.Utility, lMinHull, lMaxHull)
        If lMinHull < lMaxPowerThrustProjSize Then
            cboHullType.AddItem("Utility") : cboHullType.ItemData(cboHullType.NewIndex) = EngineTech.eyHullType.Utility
        End If

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

        'Check the Techname
        If txtTechName.Caption.Trim.Length = 0 Then
            MyBase.moUILib.AddNotification("You must specify a name for this Shield.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        'Dim lMinID(2) As Int32
        'If cboCoilMat.ListIndex > -1 Then lMinID(0) = cboCoilMat.ItemData(cboCoilMat.ListIndex)
        'If cboAccMat.ListIndex > -1 Then lMinID(1) = cboAccMat.ItemData(cboAccMat.ListIndex)
        'If cboCasingMat.ListIndex > -1 Then lMinID(2) = cboCasingMat.ItemData(cboCasingMat.ListIndex)

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

        If mbImpossibleDesign = True Then
            MyBase.moUILib.AddNotification("You must fix the flaws of this design.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
            Return
        End If

        'If IsNumeric(txtProjectionHullSize.Caption) = False OrElse IsNumeric(txtRechargeInterval.Caption) = False OrElse _
        '  IsNumeric(txtRechargeRate.Caption) = False OrElse IsNumeric(txtMaxHP.Caption) = False Then
        '    MyBase.moUILib.AddNotification("Entries must be numeric!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        '    Return
        'End If

        'Dim sTemp As String = txtProjectionHullSize.Caption & txtRechargeRate.Caption & txtMaxHP.Caption
        'If sTemp.IndexOf("."c) > -1 Then
        '    MyBase.moUILib.AddNotification("Numeric entries must be whole numbers!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        '    Return
        'ElseIf sTemp.IndexOf("-"c) > -1 Then
        '    MyBase.moUILib.AddNotification("Numeric entries must be greater than 0!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        '    Return
        'End If

        If tpProjHull.PropertyValue < 1 OrElse tpRecInt.PropertyValue <= 0 OrElse tpRecRate.PropertyValue < 1 OrElse tpMaxHP.PropertyValue < 1 Then
            MyBase.moUILib.AddNotification("Numeric entries must be greater than 0!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        Dim lProjHullSize As Int32 = CInt(tpProjHull.PropertyValue)

        'Dim fTemp As Single = CSng(tpRecInt.PropertyValue) * 30.0F
        If tpRecInt.PropertyValue = 0 OrElse tpRecInt.PropertyValue > 2000000000 Then
            MyBase.moUILib.AddNotification("Please enter a valid entry for Recharge Interval.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        Dim lRechInterval As Int32 = CInt(tpRecInt.PropertyValue)
        Dim lRechRate As Int32 = CInt(tpRecRate.PropertyValue)
        Dim lMaxHP As Int32 = CInt(tpMaxHP.PropertyValue)

        If lMaxHP > mlMaxHP OrElse lMaxHP < 1 Then
            MyBase.moUILib.AddNotification("Please enter a valid entry for Maximum Hitpoints (1 - " & mlMaxHP & ").", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lProjHullSize > mlMaxProjSize OrElse lProjHullSize < 1 Then
            MyBase.moUILib.AddNotification("Please enter a valid entry for Projection Hull Size (1 - " & mlMaxProjSize & ").", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lRechInterval < mlMinRechInt Then
            MyBase.moUILib.AddNotification("Please enter a recharge interval greater than " & (mlMinRechInt / 30.0F).ToString("#,##0.0###") & ".", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lRechRate > mlMaxRechRate Then
            MyBase.moUILib.AddNotification("Please enter a valid entry for Max Recharge Rate (1 - " & Math.Min(mlMaxRechRate, mlMaxHP) & ").", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lRechRate > lMaxHP Then
            MyBase.moUILib.AddNotification("The Recharge Rate cannot exceed the maximum hitpoints.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
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
        If tpCoilMat.PropertyLocked = True Then lSpecMin1 = CInt(tpCoilMat.PropertyValue) Else lSpecMin1 = -1
        Dim lSpecMin2 As Int32 = -1
        If tpAccMat.PropertyLocked = True Then lSpecMin2 = CInt(tpAccMat.PropertyValue) Else lSpecMin2 = -1
        Dim lSpecMin3 As Int32 = -1
        If tpCasingMat.PropertyLocked = True Then lSpecMin3 = CInt(tpCasingMat.PropertyValue) Else lSpecMin3 = -1

        MyBase.moUILib.GetMsgSys.SubmitShieldDesign(oResearchGuid, txtTechName.Caption, _
          lMaxHP, lRechRate, lRechInterval, lProjHullSize, mlMineralIDs(0), mlMineralIDs(1), mlMineralIDs(2), yColorValue, lTechID, _
          lHullReq, lPowerReq, blResCost, blResTime, blProdCost, blProdTime, lColonists, lEnlisted, lOfficers, _
          lSpecMin1, lSpecMin2, lSpecMin3, -1, -1, -1, CByte(cboHullType.ItemData(cboHullType.ListIndex)))

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eSubmitDesignClick, ObjectType.eShieldTech, -1, -1, "")
        End If

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddNotification("Shield Prototype Design Submitted", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        ReturnToPreviousViewAndReleaseBackground()

    End Sub

    'Private Sub cboAccMat_ItemSelected(ByVal lItemIndex As Integer) Handles cboAccMat.ItemSelected
    '    If cboAccMat.ListIndex > -1 Then
    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboAccMat.ItemData(cboAccMat.ListIndex))
    '    End If
    'End Sub

    'Private Sub cboCasingMat_ItemSelected(ByVal lItemIndex As Integer) Handles cboCasingMat.ItemSelected
    '    If cboCasingMat.ListIndex > -1 Then
    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboCasingMat.ItemData(cboCasingMat.ListIndex))
    '    End If
    'End Sub

    'Private Sub cboCoilMat_ItemSelected(ByVal lItemIndex As Integer) Handles cboCoilMat.ItemSelected
    '    If cboCoilMat.ListIndex > -1 Then
    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboCoilMat.ItemData(cboCoilMat.ListIndex))
    '    End If
    'End Sub

    Private Sub SetCboHelpers()
        'Dim bSuperC As Boolean = goCurrentPlayer.PlayerKnowsProperty("SUPERCONDUCTIVE POINT")
        'Dim bDensity As Boolean = goCurrentPlayer.PlayerKnowsProperty("DENSITY")
        'Dim bMagProd As Boolean = goCurrentPlayer.PlayerKnowsProperty("MAGNETIC PRODUCTION")
        'Dim bQuantum As Boolean = goCurrentPlayer.PlayerKnowsProperty("QUANTUM")
        'Dim bMagReact As Boolean = goCurrentPlayer.PlayerKnowsProperty("MAGNETIC REACTANCE")
        'Dim bThermCond As Boolean = goCurrentPlayer.PlayerKnowsProperty("THERMAL CONDUCTION")
        'Dim bTempSens As Boolean = goCurrentPlayer.PlayerKnowsProperty("TEMPERATURE SENSITIVITY")
        Dim oSB As System.Text.StringBuilder = New System.Text.StringBuilder

        ''Shields.Coil		supercond low, density high, magprod high
        oSB.Length = 0
        'If bDensity = True Then
        oSB.Append("High Density Materials ")
        'Else : oSB.Append("Materials ")
        'End If
        'If bSuperC = True Then 
        oSB.Append("with a low Superconductive Point ")
        'If bMagProd = True Then 
        oSB.Append("that have high Magnetic Production ")
        oSB.Append("work best.")
        'If bDensity = True OrElse bSuperC = True OrElse bMagProd = True Then cboCoilMat.ToolTipText = oSB.ToString Else cboCoilMat.ToolTipText = "We are not quite sure of the best way to engineer this."
        stmCoilMat.ToolTipText = oSB.ToString

        ''Shields.Accelerator	quantum high, supercond low, magreact high
        oSB.Length = 0
        'If bQuantum = True Then
        oSB.Append("Quantum-sensitive materials ")
        'Else : oSB.Append("Materials ")
        'End If
        'If bSuperC = True Then 
        oSB.Append("with a low Superconductive Point ")
        'If bMagReact = True Then 
        oSB.Append("that have high Magnetic Reactance ")
        oSB.Append("are preferred.")
        'If bQuantum = True OrElse bSuperC = True OrElse bMagReact = True Then cboAccMat.ToolTipText = oSB.ToString Else cboAccMat.ToolTipText = "We are not quite sure of the best way to engineer this."
        stmAccMat.ToolTipText = oSB.ToString

        ''Shields.Casing		density high, thermal cond high, temp sens low
        oSB.Length = 0
        'If bDensity = True Then
        oSB.Append("High Density materials ")
        'Else : oSB.Append("Materials ")
        'End If
        'If bThermCond = True Then 
        oSB.Append("with a high Thermal Conductance property ")
        'If bTempSens = True Then 
        oSB.Append("that have a low Temperature Sensitivity ")
        oSB.Append("are a must.")
        'If bDensity = True OrElse bThermCond = True OrElse bTempSens = True Then cboCasingMat.ToolTipText = oSB.ToString Else cboCasingMat.ToolTipText = "We are not quite sure of the best way to engineer this."
        stmCasingMat.ToolTipText = oSB.ToString

        If goCurrentPlayer Is Nothing = False Then
            mlMaxHP = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eShieldMaxHP)
            mlMaxProjSize = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eShieldProjSize)
            mlMinRechInt = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eShieldRechargeFreqLow)
            mlMaxRechRate = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eShieldRechargeRate)
        End If

        mlMaxRechRate = Math.Min(mlMaxRechRate, mlMaxHP)
        tpProjHull.ToolTipText = "Indicates the maximum size that the hull projector" & vbCrLf & "extends the shield system's protective field" & vbCrLf & "around the entity that the system is placed on." & vbCrLf & "This system cannot be placed on units that have" & vbCrLf & "higher hull size than this value." & vbCrLf & "Maximum value for this attribute is " & mlMaxProjSize & "."
        tpRecInt.ToolTipText = "The number of seconds between pulses projected" & vbCrLf & "by the shield system to replenish lost energy." & vbCrLf & "Must be at least " & (mlMinRechInt / 30.0F).ToString("#,##0.0###") & "."
        tpRecRate.ToolTipText = "The amount of energy that the shield system" & vbCrLf & "can replenish in a given replenish pulse." & vbCrLf & "This value cannot exceed the Maximum Hitpoints." & vbCrLf & "Maximum value for this attribute is " & mlMaxRechRate & "."
        tpMaxHP.ToolTipText = "The maximum damage the shield system can withstand." & vbCrLf & "Maximum value for this attribute is " & mlMaxHP.ToString & "."

        tpProjHull.MaxValue = mlMaxProjSize
        tpRecInt.MinValue = mlMinRechInt
        'tpRecInt.MaxValue = 30000
        tpRecInt.blAbsoluteMaximum = 30000

        tpRecRate.MaxValue = mlMaxRechRate
        tpMaxHP.MaxValue = mlMaxHP

        tpMaxHP.SetToInitialDefault()
        tpProjHull.SetToInitialDefault()
        tpRecInt.SetToInitialDefault()
        tpRecRate.SetToInitialDefault()

        'lblCoilMatItem.ToolTipText = btnSetCoilMatItem.ToolTipText
        'lblAccMatItem.ToolTipText = btnSetAccMatItem.ToolTipText
        'lblCaseMatItem.ToolTipText = btnSetCaseMatItem.ToolTipText
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
            ''Delete the design
            'Dim yMsg(7) As Byte
            'System.BitConverter.GetBytes(GlobalMessageCode.eDeleteDesign).CopyTo(yMsg, 0)
            'moTech.GetGUIDAsString.CopyTo(yMsg, 2)
            'MyBase.moUILib.SendMsgToPrimary(yMsg)
            'MyBase.moUILib.RemoveWindow(Me.ControlName)
            'ReturnToPreviousViewAndReleaseBackground()
        End If
    End Sub

    Private Sub frmShieldBuilder_OnRender() Handles Me.OnRender
        'Ok, first, is moTech nothing?
        If moTech Is Nothing = False Then
            'Ok, now, do the values used on the form currently match the tech?
            Dim bChanged As Boolean = False
            With moTech
                bChanged = tpProjHull.PropertyValue <> .lProjectionHullSize OrElse tpRecInt.PropertyValue <> .RechargeFreq OrElse _
                  tpRecRate.PropertyValue <> .RechargeRate OrElse tpMaxHP.PropertyValue <> .MaxHitPoints

                If bChanged = False Then
                    If txtTechName.Caption <> .sShieldName AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eInvalidDesign AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
                        bChanged = True
                    End If
                End If

                If bChanged = False Then
                    'Dim lID1 As Int32 = -1
                    'Dim lID2 As Int32 = -1
                    'Dim lID3 As Int32 = -1
                    Dim lColorIdx As Int32 = -1

                    'If cboCasingMat.ListIndex <> -1 Then lID1 = cboCasingMat.ItemData(cboCasingMat.ListIndex)
                    'If cboAccMat.ListIndex <> -1 Then lID2 = cboAccMat.ItemData(cboAccMat.ListIndex)
                    'If cboCoilMat.ListIndex <> -1 Then lID3 = cboCoilMat.ItemData(cboCoilMat.ListIndex)

                    If cboColor.ListIndex <> -1 Then lColorIdx = cboColor.ItemData(cboColor.ListIndex)

                    bChanged = .lCasingMineralID <> mlMineralIDs(2) OrElse .lAcceleratorMineralID <> mlMineralIDs(1) OrElse .lCoilMineralID <> mlMineralIDs(0) OrElse _
                      .ColorValue <> lColorIdx OrElse (tpCoilMat.PropertyLocked <> (.lSpecifiedMin1 <> -1)) OrElse _
                          (tpAccMat.PropertyLocked <> (.lSpecifiedMin2 <> -1)) OrElse (tpCasingMat.PropertyLocked <> (.lSpecifiedMin3 <> -1))


                    If bChanged = False Then

                        bChanged = (.lSpecifiedMin1 <> -1 AndAlso .lSpecifiedMin1 <> tpCoilMat.PropertyValue) OrElse _
                                    (.lSpecifiedMin2 <> -1 AndAlso .lSpecifiedMin2 <> tpAccMat.PropertyValue) OrElse _
                                    (.lSpecifiedMin3 <> -1 AndAlso .lSpecifiedMin3 <> tpCasingMat.PropertyValue)


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
        goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eShieldBuilder)
    End Sub

    Private Sub tp_ValueChanged(ByVal sPropName As String, ByRef ctl As ctlTechProp)
        tpRecInt.MaxValue = Math.Min(Int32.MaxValue, Math.Max(tpRecInt.MaxValue, tpRecInt.PropertyValue + 500))
        BuilderCostValueChange(False)
    End Sub
    Private Sub ComboValueChanged(ByVal lItemIndex As Int32)
        If cboHullType.ListIndex > -1 Then
            '.lHullTypeID = cboHullType.ItemData(cboHullType.ListIndex)
            Dim lHullTypeID As Int32 = cboHullType.ItemData(cboHullType.ListIndex)
            Dim lMinHull As Int32 = Base_Tech.GetMinHullForType(lHullTypeID)
            Dim lMaxHull As Int32 = Base_Tech.GetMaxHullForType(lHullTypeID)

            If mlMaxProjSize < lMinHull Then
                MyBase.moUILib.AddNotification("You cannot build a shield that covers that type of hull.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                lMinHull = 0
                lMaxHull = 0
            Else
                lMaxHull = Math.Min(mlMaxProjSize, lMaxHull)
                tpProjHull.ToolTipText = "Indicates the maximum size that the hull projector" & vbCrLf & "extends the shield system's protective field" & vbCrLf & "around the entity that the system is placed on." & vbCrLf & "This system cannot be placed on units that have" & vbCrLf & "higher hull size than this value." & vbCrLf & "Maximum value for this attribute is " & lMaxHull & "."
            End If

            mbIgnoreValueChange = True
            tpProjHull.MinValue = lMinHull
            tpProjHull.MaxValue = lMaxHull
            tpMaxHP.MaxValue = Math.Min(lMaxHull, mlMaxHP)

            If tpProjHull.PropertyValue < tpProjHull.MinValue Then tpProjHull.PropertyValue = tpProjHull.MinValue
            If tpProjHull.PropertyValue > tpProjHull.MaxValue Then tpProjHull.PropertyValue = tpProjHull.MaxValue
            If tpMaxHP.PropertyValue > tpMaxHP.MaxValue Then tpMaxHP.PropertyValue = tpMaxHP.MaxValue
            mbIgnoreValueChange = False

            tpProjHull.UpdateTextDisplay()
            tpMaxHP.UpdateTextDisplay()
            tpProjHull.IsDirty = True
            tpMaxHP.IsDirty = True
        End If
        'BuilderCostValueChange(False)
        CheckForDARequest()
    End Sub

    Private Sub EnableDisableControls()
        If mbIgnoreValueChange = True Then Return
        If cboHullType Is Nothing Then Return

        Dim bValue As Boolean = cboHullType.ListIndex > -1
        If tpMaxHP.Enabled <> bValue Then tpMaxHP.Enabled = bValue
        If tpRecRate.Enabled <> bValue Then tpRecRate.Enabled = bValue
        If tpRecInt.Enabled <> bValue Then tpRecInt.Enabled = bValue
        If tpProjHull.Enabled <> bValue Then tpProjHull.Enabled = bValue
        If tpCoilMat.Enabled <> bValue Then tpCoilMat.Enabled = bValue
        If tpAccMat.Enabled <> bValue Then tpAccMat.Enabled = bValue
        If tpCasingMat.Enabled <> bValue Then tpCasingMat.Enabled = bValue

        Dim lPhase As Int32 = 0
        If bValue = True Then
            lPhase = 1
            If tpMaxHP.PropertyValue > 0 AndAlso tpProjHull.PropertyValue > 0 AndAlso tpRecInt.PropertyValue > 0 AndAlso tpRecRate.PropertyValue > 0 Then
                lPhase = 3
                For X As Int32 = 0 To 2
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

        If tpMaxHP.Visible <> bPhase1 Then tpMaxHP.Visible = bPhase1
        If tpRecInt.Visible <> bPhase1 Then tpRecInt.Visible = bPhase1
        If tpRecRate.Visible <> bPhase1 Then tpRecRate.Visible = bPhase1
        If tpProjHull.Visible <> bPhase1 Then tpProjHull.Visible = bPhase1

        If lblColor.Visible <> bPhase2 Then lblColor.Visible = bPhase2
        If cboColor.Visible <> bPhase2 Then cboColor.Visible = bPhase2
        If lblCoilMat.Visible <> bPhase2 Then lblCoilMat.Visible = bPhase2

        If stmCoilMat.Visible <> bPhase2 Then stmCoilMat.Visible = bPhase2

        If tpCoilMat.Visible <> bPhase2 Then tpCoilMat.Visible = bPhase2
        If lblCaseMat.Visible <> bPhase2 Then lblCaseMat.Visible = bPhase2

        If stmCasingMat.Visible <> bPhase2 Then stmCasingMat.Visible = bPhase2

        If tpCasingMat.Visible <> bPhase2 Then tpCasingMat.Visible = bPhase2
        If lblAccMat.Visible <> bPhase2 Then lblAccMat.Visible = bPhase2

        If stmAccMat.Visible <> bPhase2 Then stmAccMat.Visible = bPhase2

        If tpAccMat.Visible <> bPhase2 Then tpAccMat.Visible = bPhase2

        If mfrmBuilderCost.Visible <> bPhase3 Then mfrmBuilderCost.Visible = bPhase3
    End Sub

    Protected Overrides Sub BuilderCostValueChange(ByVal bAutoFill As Boolean)
        If mbIgnoreValueChange = True Then Return
        EnableDisableControls()
        mbIgnoreValueChange = True
        lblDesignFlaw.Caption = ""
        mbImpossibleDesign = True

        Dim oShieldTechComputer As New ShieldTechComputer
        With oShieldTechComputer

            .lHullTypeID = -1
            If cboHullType.ListIndex > -1 Then
                .lHullTypeID = cboHullType.ItemData(cboHullType.ListIndex)
            End If

            If .lHullTypeID = -1 Then
                lblDesignFlaw.Caption = "All properties and materials need to be defined."
                mbIgnoreValueChange = False
                mbImpossibleDesign = True
                Return
            End If

            .SetMinDAValues(ml_Min1DA, ml_Min2DA, ml_Min3DA, ml_Min4DA, ml_Min5DA, ml_Min6DA)

            'Dim decNormalizer As Decimal = 1D
            'Dim lMaxGuns As Int32 = 1
            'Dim lMaxDPS As Int32 = 1
            'Dim lMaxHullSize As Int32 = 1
            'Dim lHullAvail As Int32 = 1
            'Dim lMaxPower As Int32 = 1
            'TechBuilderComputer.GetTypeValues(.lHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)
            Dim lMaxHullSize As Int32 = Base_Tech.GetMaxHullForType(.lHullTypeID)
            tpMaxHP.MaxValue = Math.Min(mlMaxHP, lMaxHullSize)
            If tpMaxHP.PropertyValue > tpMaxHP.MaxValue Then tpMaxHP.PropertyValue = tpMaxHP.MaxValue
            tpProjHull.MaxValue = Math.Min(lMaxHullSize, mlMaxProjSize)
            'If tpMaxHP.PropertyValue < tpRecRate.PropertyValue Then
            '    tpRecRate.PropertyValue = tpMaxHP.PropertyValue
            '    tpRecRate.UpdateTextDisplay()
            'End If
            If tpProjHull.PropertyValue > tpProjHull.MaxValue Then
                tpProjHull.PropertyValue = tpProjHull.MaxValue
                tpProjHull.UpdateTextDisplay()
            End If


            Dim lMinHull As Int32 = 0
            Dim lMaxHull As Int32 = 0
            HullTech.GetMinMaxHullOfHullType(.lHullTypeID, lMinHull, lMaxHull)
            tpProjHull.MaxValue = Math.Min(lMaxHull, mlMaxProjSize)
            tpProjHull.MinValue = Math.Min(lMinHull, mlMaxProjSize)
            If tpProjHull.PropertyValue > tpProjHull.MaxValue OrElse tpProjHull.PropertyValue < tpProjHull.MinValue Then
                lblDesignFlaw.Caption = "The design is impossible as the projection hull size" & vbCrLf & "does not match the hull size range of the selected hull type."
                mbIgnoreValueChange = False
                Return
            End If

            Dim decRInt As Decimal = Math.Max(1, tpRecInt.PropertyValue) / 30D
            .decHPS = tpRecRate.PropertyValue / decRInt
            .blMaxHP = tpMaxHP.PropertyValue
            .blProjHull = tpProjHull.PropertyValue

            'If cboCoilMat.ListIndex > -1 Then
            '    .lMineral1ID = cboCoilMat.ItemData(cboCoilMat.ListIndex)
            'Else : .lMineral1ID = -1
            'End If
            'If cboAccMat.ListIndex > -1 Then
            '    .lMineral2ID = cboAccMat.ItemData(cboAccMat.ListIndex)
            'Else : .lMineral2ID = -1
            'End If
            'If cboCasingMat.ListIndex > -1 Then
            '    .lMineral3ID = cboCasingMat.ItemData(cboCasingMat.ListIndex)
            'Else : .lMineral3ID = -1
            'End If
            .lMineral1ID = mlMineralIDs(0)
            .lMineral2ID = mlMineralIDs(1)
            .lMineral3ID = mlMineralIDs(2)

            If .lMineral1ID < 1 OrElse .lMineral2ID < 1 OrElse .lMineral3ID < 1 Then
                lblDesignFlaw.Caption = "All properties and materials need to be defined."
                mbIgnoreValueChange = False
                mbImpossibleDesign = True
                Return
            End If


            .lMineral4ID = -1
            .lMineral5ID = -1
            .lMineral6ID = -1

            .msMin1Name = "Coil"
            .msMin2Name = "Accelerator"
            .msMin3Name = "Casing"
            .msMin4Name = ""
            .msMin5Name = ""
            .msMin6Name = ""

            If bAutoFill = True Then
                .DoBuilderCostPreconfigure(lblDesignFlaw, MyBase.mfrmBuilderCost, tpCoilMat, tpAccMat, tpCasingMat, Nothing, Nothing, Nothing, ObjectType.eShieldTech, 0D, False)
            Else
                .BuilderCostValueChange(lblDesignFlaw, MyBase.mfrmBuilderCost, tpCoilMat, tpAccMat, tpCasingMat, Nothing, Nothing, Nothing, ObjectType.eShieldTech, 0D)
            End If

            mbImpossibleDesign = .bImpossibleDesign
        End With

        mbIgnoreValueChange = False

    End Sub
  
    Public Overloads Overrides Sub ViewResults(ByRef oTech As Base_Tech, ByVal lProdFactor As Integer)
        If oTech Is Nothing Then Return

        moTech = CType(oTech, ShieldTech)

        Me.btnDesign.Enabled = False

        With moTech
            mbIgnoreValueChange = True

            cboHullType.FindComboItemData(.HullTypeID)
            mbIgnoreValueChange = True

            tpMaxHP.PropertyValue = .MaxHitPoints
            If .MaxHitPoints > 0 Then tpMaxHP.PropertyLocked = True
            tpProjHull.PropertyValue = .lProjectionHullSize
            If .lProjectionHullSize > 0 Then tpProjHull.PropertyLocked = True
            If tpRecInt.MaxValue < .RechargeFreq Then tpRecInt.MaxValue = .RechargeFreq
            tpRecInt.PropertyValue = .RechargeFreq
            If .RechargeFreq > 0 Then tpRecInt.PropertyLocked = True
            tpRecRate.PropertyValue = .RechargeRate
            If .RechargeRate > 0 Then tpRecRate.PropertyLocked = True
            Me.txtTechName.Caption = .sShieldName

            mlSelectedMineralIdx = 0 : Mineral_Selected(.lCoilMineralID)
            mlSelectedMineralIdx = 1 : Mineral_Selected(.lAcceleratorMineralID)
            mlSelectedMineralIdx = 2 : Mineral_Selected(.lCasingMineralID) 

            tpCoilMat.PropertyValue = .lSpecifiedMin1
            tpAccMat.PropertyValue = .lSpecifiedMin2
            tpCasingMat.PropertyValue = .lSpecifiedMin3
            If .lSpecifiedMin1 <> -1 Then tpCoilMat.PropertyLocked = True
            If .lSpecifiedMin2 <> -1 Then tpAccMat.PropertyLocked = True
            If .lSpecifiedMin3 <> -1 Then tpCasingMat.PropertyLocked = True

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
            btnDelete.Visible = True
            If btnDelete.Visible = True Then
                'Ok, set up the buttons better
                btnDelete.Left = 10
                btnCancel.Left = Me.Width - btnCancel.Width - 10
                Dim lNewW As Int32 = (btnCancel.Left - (btnDelete.Left + btnDelete.Width))
                Dim lGapW As Int32 = lNewW - (btnDesign.Width + btnAutoFill.Width)
                lGapW \= 3  'for 3 gaps
                btnDesign.Left = btnCancel.Left - btnDesign.Width - lGapW
                btnAutoFill.Left = btnDelete.Left + btnDelete.Width + lGapW

                'cancel on right
                'design left of cancel
                'autobalance left of design
                'delete left

                'btnDesign.Left = (Me.Width \ 2) - (btnDesign.Width \ 2)

                'btnDelete.Left = 5
            End If
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
                Dim oMinFrm As UIWindow = MyBase.moUILib.GetWindow("frmMinDetail")
                If oMinFrm Is Nothing Then .Top = Me.mfrmBuilderCost.Top Else .Top = oMinFrm.Top + oMinFrm.Height
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
                .Left = Me.Left + Me.Width + 5
                Dim oMinFrm As UIWindow = MyBase.moUILib.GetWindow("frmMinDetail")
                If oMinFrm Is Nothing Then .Top = Me.mfrmBuilderCost.Top Else .Top = oMinFrm.Top + oMinFrm.Height
                .SetFromFailureCode(oTech.ErrorCodeReason)
            End With
        End If
    End Sub

    Private Sub tpMaxHP_OnLostFocus()
        If tpMaxHP.PropertyValue < tpRecRate.PropertyValue Then
            tpRecRate.PropertyValue = tpMaxHP.PropertyValue
            tpRecRate.UpdateTextDisplay()
            tpRecRate.IsDirty = True
        End If
    End Sub

    Private mlMineralIDs(2) As Int32
    Private Sub SetButtonClicked(ByVal lMinIdx As Int32)
        Dim lHullTechID As Int32 = -1
        If cboHullType.ListIndex > -1 Then
            lHullTechID = cboHullType.ItemData(cboHullType.ListIndex)
        End If

        Dim oTech As New ShieldTechComputer
        oTech.lHullTypeID = lHullTechID

        If lMinIdx = -1 Then Return

        mlSelectedMineralIdx = -1
        oTech.MineralCBOExpanded(lMinIdx, ObjectType.eShieldTech)
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
            'Dim sControlName As String = ""
            Select Case mlSelectedMineralIdx
                Case 0
                    stmCoilMat.SetMineralName(sMinName)
                Case 1
                    stmAccMat.SetMineralName(sMinName)
                Case 2
                    stmCasingMat.SetMineralName(sMinName)
            End Select
            'If sControlName <> "" Then
            '    sControlName = sControlName.ToUpper
            '    For X As Int32 = 0 To Me.ChildrenUB
            '        If Me.moChildren(X) Is Nothing = False AndAlso Me.moChildren(X).ControlName Is Nothing = False Then
            '            If Me.moChildren(X).ControlName.ToUpper = sControlName Then
            '                CType(Me.moChildren(X), UILabel).Caption = sMinName
            '                Exit For
            '            End If
            '        End If
            '    Next X
            'End If
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
        If cboHullType Is Nothing = False AndAlso cboHullType.ListIndex > -1 Then
            lHullTypeID = cboHullType.ItemData(cboHullType.ListIndex)
        End If
        If bRequestDA = True AndAlso lHullTypeID <> -1 Then RequestDAValues(ObjectType.eShieldTech, 0, CByte(lHullTypeID), mlMineralIDs(0), mlMineralIDs(1), mlMineralIDs(2), -1, -1, -1, 0, 0)
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