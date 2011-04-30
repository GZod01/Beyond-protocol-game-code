Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmArmorBuilder
    Inherits frmTechBuilder
    ' UIWindow

    'Private lblBeam As UILabel
    'Private lblImpact As UILabel
    'Private lblPiercing As UILabel
    'Private lblMagnetic As UILabel
    'Private lblChemical As UILabel
    'Private lblBurn As UILabel
    'Private lblRadar As UILabel
    'Private lblHull As UILabel
    'Private lblHP As UILabel
    Private lblOuter As UILabel
    Private lblMiddle As UILabel
    Private lblInner As UILabel
    Private lblArmorDesigner As UILabel
    'Private txtBeam As UITextBox
    'Private txtImpact As UITextBox
    'Private txtPiercing As UITextBox
    'Private txtMagnetic As UITextBox
    'Private txtChemical As UITextBox
    'Private txtBurn As UITextBox
    'Private txtRadar As UITextBox
    'Private txtHullUsage As UITextBox
    'Private txtHPPerPlate As UITextBox

    Private tpBeam As ctlTechProp
    Private tpImpact As ctlTechProp
    Private tpPiercing As ctlTechProp
    Private tpMagnetic As ctlTechProp
    Private tpChemical As ctlTechProp
    Private tpBurn As ctlTechProp 
    Private tpHullUsage As ctlTechProp
    Private tpHPPerPlate As ctlTechProp

    'Private WithEvents cboOuter As UIComboBox
    'Private WithEvents cboMiddle As UIComboBox
    'Private WithEvents cboInner As UIComboBox
    Private stmOuter As ctlSetTechMineral
    Private stmMiddle As ctlSetTechMineral
    Private stmInner As ctlSetTechMineral

    Private tpOuter As ctlTechProp
    Private tpMiddle As ctlTechProp
    Private tpInner As ctlTechProp

    'Private lblHiLite As UILabel
    'Private chkBeamHilite As UICheckBox
    'Private chkImpactHiLite As UICheckBox
    'Private chkPiercingHiLite As UICheckBox
    'Private chkMagneticHiLite As UICheckBox
    'Private chkChemicalHiLite As UICheckBox
    'Private chkBurnHiLite As UICheckBox 

    Private WithEvents btnDesign As UIButton
    Private WithEvents btnCancel As UIButton
    Private WithEvents btnDelete As UIButton

    Private WithEvents btnRename As UIButton

    Private WithEvents btnHelp As UIButton

    Private lblTechName As UILabel
    Private WithEvents txtTechName As UITextBox

    Private lblDesignFlaw As UILabel
    Private lblArmorIntegrity As UILabel

    Private mlEntityIndex As Int32 = -1

    Private moTech As ArmorTech = Nothing
    Private mfrmResCost As frmProdCost = Nothing
    Private mfrmProdCost As frmProdCost = Nothing

    Private mbImpossibleDesign As Boolean = False
    Private mbIgnoreValueChange As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmArmorBuilder initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eArmorBuilder
            .ControlName = "frmArmorBuilder"
            .Left = 4
            .Top = 5
            .Width = 512 '450
            .Height = 512
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

        'lblTechName initial props
        lblTechName = New UILabel(oUILib)
        With lblTechName
            .ControlName = "lblTechName"
            .Left = 15
            .Top = 45
            .Width = 100
            .Height = 17
            .Enabled = True
            .Visible = True
            .Caption = "Armor Name:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTechName, UIControl))

        'txtTechName initial props
        txtTechName = New UITextBox(oUILib)
        With txtTechName
            .ControlName = "txtTechName"
            .Left = 180 '210
            .Top = 45
            .Width = 184
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
        Me.AddChild(CType(txtTechName, UIControl))

        ''lblHiLite initial props
        'lblHiLite = New UILabel(oUILib)
        'With lblHiLite
        '    .ControlName = "lblHiLite"
        '    .Left = 420
        '    .Top = 45
        '    .Width = 80
        '    .Height = 17
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Highlight"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
        'End With
        'Me.AddChild(CType(lblHiLite, UIControl))

        ''lblBeam initial props
        'lblBeam = New UILabel(oUILib)
        'With lblBeam
        '    .ControlName = "lblBeam"
        '    .Left = 30
        '    .Top = 68
        '    .Width = 184
        '    .Height = 24
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Desired Beam Resist:"
        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblBeam, UIControl))
        tpBeam = New ctlTechProp(oUILib)
        With tpBeam
            .ControlName = "tpBeam"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 100
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Beam Resist:"
            .PropertyValue = 0
            .Top = 70
            .Visible = True
            '.Width = 512
            .ToolTipText = "This percentage value of damage will be completely mitigated before being applied to armor." & vbCrLf & _
                           "This value requires that the armor pass the Armor Integrity test before any reduction."
        End With
        Me.AddChild(CType(tpBeam, UIControl))

        ''lblImpact initial props
        'lblImpact = New UILabel(oUILib)
        'With lblImpact
        '    .ControlName = "lblImpact"
        '    .Left = 30
        '    .Top = 93
        '    .Width = 184
        '    .Height = 24
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Desired Impact Resist:"
        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblImpact, UIControl))
        tpImpact = New ctlTechProp(oUILib)
        With tpImpact
            .ControlName = "tpImpact"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 100
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Impact Resist:"
            .PropertyValue = 0
            .Top = 95
            .Visible = True
            '.Width = 512
            .ToolTipText = "This percentage value of damage will be completely mitigated before being applied to armor." & vbCrLf & _
                           "This value requires that the armor pass the Armor Integrity test before any reduction."
        End With
        Me.AddChild(CType(tpImpact, UIControl))

        ''lblPiercing initial props
        'lblPiercing = New UILabel(oUILib)
        'With lblPiercing
        '    .ControlName = "lblPiercing"
        '    .Left = 30
        '    .Top = 118
        '    .Width = 184
        '    .Height = 24
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Desired Piercing Resist:"
        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblPiercing, UIControl))
        tpPiercing = New ctlTechProp(oUILib)
        With tpPiercing
            .ControlName = "tpPiercing"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 100
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Piercing Resist:"
            .PropertyValue = 0
            .Top = 120
            .Visible = True
            '.Width = 512
            .ToolTipText = "This percentage value of damage will be completely mitigated before being applied to armor." & vbCrLf & _
                           "This value requires that the armor pass the Armor Integrity test before any reduction."
        End With
        Me.AddChild(CType(tpPiercing, UIControl))

        ''lblMagnetic initial props
        'lblMagnetic = New UILabel(oUILib)
        'With lblMagnetic
        '    .ControlName = "lblMagnetic"
        '    .Left = 30
        '    .Top = 143
        '    .Width = 184
        '    .Height = 24
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Desired Magnetic Resist:"
        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblMagnetic, UIControl))
        tpMagnetic = New ctlTechProp(oUILib)
        With tpMagnetic
            .ControlName = "tpMagnetic"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 100
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Magnetic Resist:"
            .PropertyValue = 0
            .Top = 145
            .Visible = True
            '.Width = 512
            .ToolTipText = "This percentage value of damage will be completely mitigated before being applied to armor." & vbCrLf & _
                           "This value requires that the armor pass the Armor Integrity test before any reduction."
        End With
        Me.AddChild(CType(tpMagnetic, UIControl))

        ''lblChemical initial props
        'lblChemical = New UILabel(oUILib)
        'With lblChemical
        '    .ControlName = "lblChemical"
        '    .Left = 30
        '    .Top = 168
        '    .Width = 184
        '    .Height = 24
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Desired Chemical Resist:"
        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblChemical, UIControl))
        tpChemical = New ctlTechProp(oUILib)
        With tpChemical
            .ControlName = "tpChemical"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 100
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Chemical Resist:"
            .PropertyValue = 0
            .Top = 170
            .Visible = True
            '.Width = 512
            .ToolTipText = "This percentage value of damage will be completely mitigated before being applied to armor." & vbCrLf & _
                           "This value requires that the armor pass the Armor Integrity test before any reduction."
        End With
        Me.AddChild(CType(tpChemical, UIControl))

        ''lblBurn initial props
        'lblBurn = New UILabel(oUILib)
        'With lblBurn
        '    .ControlName = "lblBurn"
        '    .Left = 30
        '    .Top = 193
        '    .Width = 184
        '    .Height = 24
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Desired Burn Resist:"
        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblBurn, UIControl))
        tpBurn = New ctlTechProp(oUILib)
        With tpBurn
            .ControlName = "tpBurn"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 100
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Burn Resist:"
            .PropertyValue = 0
            .Top = 195
            .Visible = True
            '.Width = 512
            .ToolTipText = "This percentage value of damage will be completely mitigated before being applied to armor." & vbCrLf & _
                           "This value requires that the armor pass the Armor Integrity test before any reduction."
        End With
        Me.AddChild(CType(tpBurn, UIControl))

        ''lblRadar initial props
        'lblRadar = New UILabel(oUILib)
        'With lblRadar
        '    .ControlName = "lblRadar"
        '    .Left = 30
        '    .Top = 218
        '    .Width = 184
        '    .Height = 24
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Desired Radar Resist:"
        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblRadar, UIControl))

        '====================
        'tpRadar = New ctlTechProp(oUILib)
        'With tpRadar
        '    .ControlName = "tpRadar"
        '    .Enabled = True
        '    .Height = 18
        '    .Left = 15
        '    .MaxValue = 100
        '    .MinValue = 0
        '    .PropertyLocked = False
        '    .PropertyName = "Radar Resist:"
        '    .PropertyValue = 0
        '    .Top = 220
        '    .Visible = True
        '    '.Width = 512
        '    .ToolTipText = "Indicates how much easy this armor is to being detected."
        'End With
        'Me.AddChild(CType(tpRadar, UIControl))

        ''lblHull initial props
        'lblHull = New UILabel(oUILib)
        'With lblHull
        '    .ControlName = "lblHull"
        '    .Left = 30
        '    .Top = 252
        '    .Width = 184
        '    .Height = 24
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Hull Usage Per Plate:"
        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblHull, UIControl))

        ''lblHP initial props
        'lblHP = New UILabel(oUILib)
        'With lblHP
        '    .ControlName = "lblHP"
        '    .Left = 30
        '    .Top = 277
        '    .Width = 184
        '    .Height = 24
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Hit Points Per Plate:"
        '    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter
        'End With
        'Me.AddChild(CType(lblHP, UIControl))

        'lblArmorDesigner initial props
        lblArmorDesigner = New UILabel(oUILib)
        With lblArmorDesigner
            .ControlName = "lblArmorDesigner"
            .Left = 5
            .Top = 5
            .Width = 182
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Armor Designer"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Arial", 18, FontStyle.Bold Or FontStyle.Italic, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        End With
        Me.AddChild(CType(lblArmorDesigner, UIControl))
 

        ''chkBeamHilite initial props
        'chkBeamHilite = New UICheckBox(oUILib)
        'With chkBeamHilite
        '    .ControlName = "chkBeamHilite"
        '    .Left = tpBeam.Left + tpBeam.Width + 20 '290
        '    .Top = 70
        '    .Width = 69
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = ""
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(6, DrawTextFormat)
        '    .Value = False
        '    .ToolTipText = "Indicates to focus on this resistance when highlighting mineral properties."
        '    .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        'End With
        'Me.AddChild(CType(chkBeamHilite, UIControl))
 

        ''chkImpactHiLite initial props
        'chkImpactHiLite = New UICheckBox(oUILib)
        'With chkImpactHiLite
        '    .ControlName = "chkImpactHiLite"
        '    .Left = chkBeamHilite.Left '290
        '    .Top = 95
        '    .Width = 69
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = ""
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(6, DrawTextFormat)
        '    .Value = False
        '    .ToolTipText = "Indicates to focus on this resistance when highlighting mineral properties."
        '    .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        'End With
        'Me.AddChild(CType(chkImpactHiLite, UIControl))
 

        ''chkPiercingHiLite initial props
        'chkPiercingHiLite = New UICheckBox(oUILib)
        'With chkPiercingHiLite
        '    .ControlName = "chkPiercingHiLite"
        '    .Left = chkBeamHilite.Left '290
        '    .Top = 120
        '    .Width = 69
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = ""
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(6, DrawTextFormat)
        '    .Value = False
        '    .ToolTipText = "Indicates to focus on this resistance when highlighting mineral properties."
        '    .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        'End With
        'Me.AddChild(CType(chkPiercingHiLite, UIControl))

 

        ''chkMagneticHiLite initial props
        'chkMagneticHiLite = New UICheckBox(oUILib)
        'With chkMagneticHiLite
        '    .ControlName = "chkMagneticHiLite"
        '    .Left = chkBeamHilite.Left '290
        '    .Top = 145
        '    .Width = 69
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = ""
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(6, DrawTextFormat)
        '    .Value = False
        '    .ToolTipText = "Indicates to focus on this resistance when highlighting mineral properties."
        '    .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        'End With
        'Me.AddChild(CType(chkMagneticHiLite, UIControl))

 
        ''chkChemicalHiLite initial props
        'chkChemicalHiLite = New UICheckBox(oUILib)
        'With chkChemicalHiLite
        '    .ControlName = "chkChemicalHiLite"
        '    .Left = chkBeamHilite.Left '290
        '    .Top = 170
        '    .Width = 69
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = ""
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(6, DrawTextFormat)
        '    .Value = False
        '    .ToolTipText = "Indicates to focus on this resistance when highlighting mineral properties."
        '    .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        'End With
        'Me.AddChild(CType(chkChemicalHiLite, UIControl))


        ''chkBurnHiLite initial props
        'chkBurnHiLite = New UICheckBox(oUILib)
        'With chkBurnHiLite
        '    .ControlName = "chkBurnHiLite"
        '    .Left = chkBeamHilite.Left '290
        '    .Top = 195
        '    .Width = 69
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = ""
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(6, DrawTextFormat)
        '    .Value = False
        '    .ToolTipText = "Indicates to focus on this resistance when highlighting mineral properties."
        '    .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        'End With
        'Me.AddChild(CType(chkBurnHiLite, UIControl))

         
        tpHullUsage = New ctlTechProp(oUILib)
        With tpHullUsage
            .ControlName = "tpHullUsage"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Hull Usage Per Plate:"
            .PropertyValue = 0
            .Top = 255
            .Visible = True
            '.Width = 512
            .ToolTipText = "Indicates how much hull each plate of armor will take up."
        End With
        Me.AddChild(CType(tpHullUsage, UIControl))

        ''txtHPPerPlate initial props
        'txtHPPerPlate = New UITextBox(oUILib)
        'With txtHPPerPlate
        '    .ControlName = "txtHPPerPlate"
        '    .Left = 225
        '    .Top = 279
        '    .Width = 100
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "1"
        '    '.ForeColor = muSettings.InterfaceBorderColor
        '    .ForeColor = muSettings.InterfaceTextBoxForeColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Left
        '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
        '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
        '    .MaxLength = 9
        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .bNumericOnly = True
        'End With
        'Me.AddChild(CType(txtHPPerPlate, UIControl))
        tpHPPerPlate = New ctlTechProp(oUILib)
        With tpHPPerPlate
            .ControlName = "tpHPPerPlate"
            .Enabled = True
            .Height = 18
            .Left = 15
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .blAbsoluteMaximum = 1000000
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Hit Points Per Plate:"
            .PropertyValue = 0
            .Top = 280
            .Visible = True
            '.Width = 512
            .ToolTipText = "Indicates how many hitpoints of armor each plate of armor will have."
        End With
        Me.AddChild(CType(tpHPPerPlate, UIControl))

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
            .Left = (Me.Width \ 2) - 50 '100
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
            .Left = Me.Width - 110
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
            .Left = 10 ' btnDesign.Left - 105
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

        'lblDesignFlaw initial props
        lblDesignFlaw = New UILabel(oUILib)
        With lblDesignFlaw
            .ControlName = "lblDesignFlaw"
            .Left = 15
            .Top = 400
            .Width = 420
            .Height = 20
            .Enabled = True
            .Visible = False
            .Caption = ""
            .ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            '.BackImageColor = Color.Bisque
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(lblDesignFlaw, UIControl))

        'lblArmorIntegrity initial props
        lblArmorIntegrity = New UILabel(oUILib)
        With lblArmorIntegrity
            .ControlName = "lblArmorIntegrity"
            .Left = 15
            .Top = lblDesignFlaw.Top + lblDesignFlaw.Height + 1
            .Width = 420
            .Height = 20
            .Enabled = True
            .Visible = False
            .Caption = ""
            .ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(lblArmorIntegrity, UIControl))

        'lblOuter initial props
        lblOuter = New UILabel(oUILib)
        With lblOuter
            .ControlName = "lblOuter"
            .Left = 15
            .Top = 310
            .Width = 100 '184
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Outer Layer:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblOuter, UIControl))

        'lblMiddle initial props
        lblMiddle = New UILabel(oUILib)
        With lblMiddle
            .ControlName = "lblMiddle"
            .Left = 15
            .Top = 335
            .Width = 100 '184
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Middle Layer:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblMiddle, UIControl))

        'lblInner initial props
        lblInner = New UILabel(oUILib)
        With lblInner
            .ControlName = "lblInner"
            .Left = 15
            .Top = 360
            .Width = 100 ' 184
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Inner Layer:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblInner, UIControl))

        tpInner = New ctlTechProp(oUILib)
        With tpInner
            .ControlName = "tpInner"
            .Enabled = True
            .Height = 18
            .Left = lblInner.Left + lblInner.Width + 160    '10 for left and right of combo and then 150 for combo
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Inner Material"
            .PropertyValue = 0
            .Top = 362
            .Visible = True
            .Width = 512 - .Left - 10
            .SetNoPropDisplay(True)
        End With
        Me.AddChild(CType(tpInner, UIControl))

        tpMiddle = New ctlTechProp(oUILib)
        With tpMiddle
            .ControlName = "tpMiddle"
            .Enabled = True
            .Height = 18
            .Left = tpInner.Left
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Middle Material"
            .PropertyValue = 0
            .Top = 337
            .Visible = True
            .Width = tpInner.Width
            .SetNoPropDisplay(True)
        End With
        Me.AddChild(CType(tpMiddle, UIControl))

        tpOuter = New ctlTechProp(oUILib)
        With tpOuter
            .ControlName = "tpOuter"
            .Enabled = True
            .Height = 18
            .Left = tpInner.Left
            .MaxValue = 1000 '0000
            .bNoMaxValue = True
            .MinValue = 0
            .PropertyLocked = False
            .PropertyName = "Outer Material"
            .PropertyValue = 0
            .Top = 312
            .Visible = True
            .Width = tpInner.Width
            .SetNoPropDisplay(True)
        End With
        Me.AddChild(CType(tpOuter, UIControl))

        ''cboInner initial props
        'cboInner = New UIComboBox(oUILib)
        'With cboInner
        '    .ControlName = "cboInner"
        '    .Left = lblInner.Left + lblInner.Width + 5
        '    .Top = 362
        '    .Width = 150
        '    .Height = 20
        '    .Enabled = True
        '    .Visible = True
        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .FillColor = muSettings.InterfaceTextBoxFillColor
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    .DropDownListBorderColor = muSettings.InterfaceBorderColor
        '    .mbAcceptReprocessEvents = True
        'End With
        'Me.AddChild(CType(cboInner, UIControl))
        stmInner = New ctlSetTechMineral(oUILib)
        With stmInner
            .ControlName = "stmInner"
            .Left = lblInner.Left + lblInner.Width + 5
            .Top = 362
            .Width = 150 '175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 2
        End With
        Me.AddChild(CType(stmInner, UIControl))

        ''cboMiddle initial props
        'cboMiddle = New UIComboBox(oUILib)
        'With cboMiddle
        '    .ControlName = "cboMiddle"
        '    .Left = cboInner.Left
        '    .Top = 337
        '    .Width = 150
        '    .Height = 20
        '    .Enabled = True
        '    .Visible = True
        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .FillColor = muSettings.InterfaceTextBoxFillColor
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    .DropDownListBorderColor = muSettings.InterfaceBorderColor
        '    .mbAcceptReprocessEvents = True
        'End With
        'Me.AddChild(CType(cboMiddle, UIControl))
        stmMiddle = New ctlSetTechMineral(oUILib)
        With stmMiddle
            .ControlName = "stmMiddle"
            .Left = stmInner.Left
            .Top = 337
            .Width = 150 '175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 1
        End With
        Me.AddChild(CType(stmMiddle, UIControl))

        ''cboOuter initial props
        'cboOuter = New UIComboBox(oUILib)
        'With cboOuter
        '    .ControlName = "cboOuter"
        '    .Left = cboInner.Left
        '    .Top = 312
        '    .Width = 150
        '    .Height = 20
        '    .Enabled = True
        '    .Visible = True
        '    '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    .BorderColor = muSettings.InterfaceBorderColor
        '    .FillColor = muSettings.InterfaceTextBoxFillColor
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
        '    .DropDownListBorderColor = muSettings.InterfaceBorderColor
        '    .mbAcceptReprocessEvents = True
        'End With
        'Me.AddChild(CType(cboOuter, UIControl))
        stmOuter = New ctlSetTechMineral(oUILib)
        With stmOuter
            .ControlName = "stmOuter"
            .Left = stmInner.Left
            .Top = 312
            .Width = 150 '175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 0
        End With
        Me.AddChild(CType(stmOuter, UIControl))

        'AddHandler cboOuter.ItemSelected, AddressOf ComboBoxItemSelected
        'AddHandler cboMiddle.ItemSelected, AddressOf ComboBoxItemSelected
        'AddHandler cboInner.ItemSelected, AddressOf ComboBoxItemSelected
        AddHandler stmOuter.SetButtonClicked, AddressOf SetButtonClicked
        AddHandler stmMiddle.SetButtonClicked, AddressOf SetButtonClicked
        AddHandler stmInner.SetButtonClicked, AddressOf SetButtonClicked

        'Private tpBeam As ctlTechProp
        'Private tpImpact As ctlTechProp
        'Private tpPiercing As ctlTechProp
        'Private tpMagnetic As ctlTechProp
        'Private tpChemical As ctlTechProp
        'Private tpBurn As ctlTechProp
        'Private tpRadar As ctlTechProp
        'Private tpHullUsage As ctlTechProp
        'Private tpHPPerPlate As ctlTechProp

        AddHandler tpBeam.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpImpact.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpPiercing.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpMagnetic.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpChemical.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpBurn.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        'AddHandler tpRadar.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpHullUsage.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpHPPerPlate.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpOuter.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpMiddle.PropertyValueChanged, AddressOf tp_PropertyValueChanged
        AddHandler tpInner.PropertyValueChanged, AddressOf tp_PropertyValueChanged

        FillValues()

        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
        ofrm.ShowMineralDetail(Me.Left + Me.Width + 5, Me.Top, Me.Height, -1)

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        glCurrentEnvirView = CurrentView.eArmorResearch

        MyBase.mfrmBuilderCost.SetAsArmorBuilder()
        SetBuildPhase()

        If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
        goFullScreenBackground = Nothing
        goFullScreenBackground = goResMgr.GetTexture("armorbuilder.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "TechBacks.pak")
    End Sub

    Private Sub ComboBoxItemSelected(ByVal lItemIndex As Int32)
        BuilderCostValueChange(False)
    End Sub
    Private Sub tp_PropertyValueChanged(ByVal sPropName As String, ByRef ctl As ctlTechProp)
        BuilderCostValueChange(False)
    End Sub

    Public Sub FillValues()
        'Dim lSorted() As Int32 = GetSortedMineralIdxArray(False)

        ''Fill our minerals/alloys
        'cboOuter.Clear() : cboMiddle.Clear() : cboInner.Clear()
        'If lSorted Is Nothing = False Then
        '    For X As Int32 = 0 To lSorted.GetUpperBound(0)
        '        'If goMinerals(lSorted(X)).IsFullyStudied = False Then Continue For
        '        'If glMineralIdx(X) <> -1 AndAlso goMinerals(X).bDiscovered = True Then
        '        If lSorted(X) <> -1 Then
        '            cboOuter.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboOuter.ItemData(cboOuter.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '            cboMiddle.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboMiddle.ItemData(cboMiddle.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '            cboInner.AddItem(goMinerals(lSorted(X)).MineralName)
        '            cboInner.ItemData(cboInner.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '        End If
        '    Next X
        'End If

        SetCboHelpers()
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

        

        'Dim lMinID(2) As Int32
        'If cboOuter.ListIndex > -1 Then lMinID(0) = cboOuter.ItemData(cboOuter.ListIndex)
        'If cboMiddle.ListIndex > -1 Then lMinID(1) = cboMiddle.ItemData(cboMiddle.ListIndex)
        'If cboInner.ListIndex > -1 Then lMinID(2) = cboInner.ItemData(cboInner.ListIndex)

        'Check the Techname
        If txtTechName.Caption.Trim.Length = 0 Then
            MyBase.moUILib.AddNotification("You must specify a name for this Armor.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        For X As Int32 = 0 To 2
            If mlMineralIDs(X) < 1 Then
                MyBase.moUILib.AddNotification("You must select three layers of materials!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                MyBase.moUILib.AddNotification("The layers can be the same material.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
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

        'If IsNumeric(txtBeam.Caption) = False OrElse IsNumeric(txtImpact.Caption) = False OrElse IsNumeric(txtPiercing.Caption) = False _
        '  OrElse IsNumeric(txtMagnetic.Caption) = False OrElse IsNumeric(txtChemical.Caption) = False OrElse IsNumeric(txtBurn.Caption) = False _
        '  OrElse IsNumeric(txtRadar.Caption) = False OrElse IsNumeric(txtHullUsage.Caption) = False OrElse IsNumeric(txtHPPerPlate.Caption) = False Then
        '    MyBase.moUILib.AddNotification("Please enter numeric values!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        '    Return
        'End If

        'Dim sTemp As String = txtBeam.Caption & txtImpact.Caption & txtPiercing.Caption & txtMagnetic.Caption & _
        '  txtChemical.Caption & txtBurn.Caption & txtRadar.Caption & txtHullUsage.Caption & txtHPPerPlate.Caption

        'If sTemp.IndexOf("."c) > -1 Then
        '    MyBase.moUILib.AddNotification("Numeric values must be whole numbers.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        '    Return
        'ElseIf sTemp.IndexOf("-"c) > -1 Then
        '    MyBase.moUILib.AddNotification("Numeric values must be greater than 0.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        '    Return
        'End If

        'If Val(txtBeam.Caption) > 100 OrElse Val(txtImpact.Caption) > 100 OrElse Val(txtPiercing.Caption) > 100 OrElse _
        '  Val(txtMagnetic.Caption) > 100 OrElse Val(txtChemical.Caption) > 100 OrElse Val(txtBurn.Caption) > 100 OrElse _
        '  Val(txtRadar.Caption) > 100 Then
        '    MyBase.moUILib.AddNotification("Resistance values must be less than or equal to 100!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        '    Return
        'End If
        If tpBeam.PropertyValue > 200 OrElse tpImpact.PropertyValue > 200 OrElse tpPiercing.PropertyValue > 200 OrElse _
            tpMagnetic.PropertyValue > 200 OrElse tpChemical.PropertyValue > 200 OrElse tpBurn.PropertyValue > 200 Then
            MyBase.moUILib.AddNotification("Resistance values must be less than or equal to 200!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        Dim yBeam As Byte = CByte(tpBeam.PropertyValue)
        Dim yImpact As Byte = CByte(tpImpact.PropertyValue)
        Dim yPierce As Byte = CByte(tpPiercing.PropertyValue)
        Dim yMagnetic As Byte = CByte(tpMagnetic.PropertyValue)
        Dim yChem As Byte = CByte(tpChemical.PropertyValue)
        Dim yBurn As Byte = CByte(tpBurn.PropertyValue)
        Dim yRadar As Byte = 0 'CByte(tpRadar.PropertyValue)
        Dim lHull As Int32 = CInt(tpHullUsage.PropertyValue)
        Dim lHP As Int32 = CInt(tpHPPerPlate.PropertyValue)

        If lHull < 1 Then
            MyBase.moUILib.AddNotification("Hull Usage Per Plate must be greater than 0!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lHP < 1 Then
            MyBase.moUILib.AddNotification("Hit Points Per Plate must be greater than 0!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        If lHull > Int32.MaxValue Then
            MyBase.moUILib.AddNotification("Hull Usage Per Plate cannot be that high.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        ElseIf lHP > Int32.MaxValue Then
            MyBase.moUILib.AddNotification("Hit Points Per Plate cannot be that high.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
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
        If tpOuter.PropertyLocked = True Then lSpecMin1 = CInt(tpOuter.PropertyValue) Else lSpecMin1 = -1
        Dim lSpecMin2 As Int32 = -1
        If tpMiddle.PropertyLocked = True Then lSpecMin2 = CInt(tpMiddle.PropertyValue) Else lSpecMin2 = -1
        Dim lSpecMin3 As Int32 = -1
        If tpInner.PropertyLocked = True Then lSpecMin3 = CInt(tpInner.PropertyValue) Else lSpecMin3 = -1
        Dim lSpecMin4 As Int32 = -1 
        Dim lSpecMin5 As Int32 = -1 
        Dim lSpecMin6 As Int32 = -1 

        Dim lTechID As Int32 = -1
        If moTech Is Nothing = False Then lTechID = moTech.ObjectID
        MyBase.moUILib.GetMsgSys.SubmitArmorDesign(oResearchGuid, txtTechName.Caption, yBeam, yImpact, yPierce, _
          yMagnetic, yChem, yBurn, yRadar, lHull, lHP, mlMineralIDs(0), mlMineralIDs(1), mlMineralIDs(2), lTechID, lHullReq, lPowerReq, _
          blResCost, blResTime, blProdCost, blProdTime, lColonists, lEnlisted, lOfficers, lSpecMin1, lSpecMin2, _
          lSpecMin3, lSpecMin4, lSpecMin5, lSpecMin6)

        If NewTutorialManager.TutorialOn = True Then
            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eSubmitDesignClick, ObjectType.eArmorTech, -1, -1, "")
        End If

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddNotification("Armor Prototype Design Submitted", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

        ReturnToPreviousViewAndReleaseBackground()
    End Sub

    'Private Sub cboInner_ItemSelected(ByVal lItemIndex As Integer) Handles cboInner.ItemSelected
    '    If cboInner.ListIndex > -1 Then
    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboInner.ItemData(cboInner.ListIndex))
    '    End If
    'End Sub

    'Private Sub cboMiddle_ItemSelected(ByVal lItemIndex As Integer) Handles cboMiddle.ItemSelected
    '    If cboMiddle.ListIndex > -1 Then
    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboMiddle.ItemData(cboMiddle.ListIndex))
    '    End If
    'End Sub

    'Private Sub cboOuter_ItemSelected(ByVal lItemIndex As Integer) Handles cboOuter.ItemSelected
    '    If cboOuter.ListIndex > -1 Then
    '        Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
    '        If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
    '        ofrm.ShowMineralDetail(Me.Left + Me.Width + 25, Me.Top, Me.Height, cboOuter.ItemData(cboOuter.ListIndex))
    '    End If
    'End Sub

    Private Sub SetCboHelpers()
        'Dim bChemReact As Boolean = goCurrentPlayer.PlayerKnowsProperty("CHEMICAL REACTANCE")
        'Dim bCompress As Boolean = goCurrentPlayer.PlayerKnowsProperty("COMPRESSIBILITY")
        'Dim bReflect As Boolean = goCurrentPlayer.PlayerKnowsProperty("REFLECTION")
        'Dim bQuantum As Boolean = goCurrentPlayer.PlayerKnowsProperty("QUANTUM")
        'Dim bDensity As Boolean = goCurrentPlayer.PlayerKnowsProperty("DENSITY")
        'Dim bThermCond As Boolean = goCurrentPlayer.PlayerKnowsProperty("THERMAL CONDUCTION")
        'Dim bRefract As Boolean = goCurrentPlayer.PlayerKnowsProperty("REFRACTION")
        'Dim bHardness As Boolean = goCurrentPlayer.PlayerKnowsProperty("HARDNESS")
        'Dim bMalleable As Boolean = goCurrentPlayer.PlayerKnowsProperty("MALLEABLE")
        'Dim bMagReact As Boolean = goCurrentPlayer.PlayerKnowsProperty("MAGNETIC REACTANCE")
        'Dim bCombust As Boolean = goCurrentPlayer.PlayerKnowsProperty("COMBUSTIVENESS")


        ''Beam: reflect, refract high
        'If bReflect = True AndAlso bRefract = True Then
        tpBeam.ToolTipText &= vbCrLf & "Higher Reflective and Refractive properties aid beam resistance."
        'ElseIf bReflect = True Then
        '    txtBeam.ToolTipText = "Higher Reflective properties aid beam resistance."
        'ElseIf bRefract = True Then
        '    txtBeam.ToolTipText = "Higher Refractive properties aid beam resistance."
        'Else : txtBeam.ToolTipText = "We are not quite sure of the best way to engineer this."
        'End If

        ''Impact: density, hardness high
        'If bDensity = True AndAlso bHardness = True Then
        tpImpact.ToolTipText &= vbCrLf & "Harder materials that have high Density help Impact resistance."
        'ElseIf bDensity = True Then
        '    txtImpact.ToolTipText = "Denser materials help Impact resistance."
        'ElseIf bHardness = True Then
        '    txtImpact.ToolTipText = "Harder materials help impact resistance."
        'Else : txtImpact.ToolTipText = "We are not quite sure of the best way to engineer this."
        'End If

        ''Piercing: hardness, malleable high
        'If bHardness = True AndAlso bMalleable = True Then
        tpPiercing.ToolTipText &= vbCrLf & "Hard materials that are Malleable improve Piercing resistance."
        'ElseIf bHardness = True Then
        '    txtPiercing.ToolTipText = "Hard materials improve Piercing resistance."
        'ElseIf bMalleable = True Then
        '    txtPiercing.ToolTipText = "Malleable materials improve Piercing resistance."
        'Else : txtPiercing.ToolTipText = "We are not quite sure of the best way to engineer this."
        'End If

        ''Magnetic: mag react, quantum high
        'If bMagReact = True AndAlso bQuantum = True Then
        tpMagnetic.ToolTipText &= vbCrLf & "Quantum-sensitive materials with high Magnetic Reactance aid in Magnetic resistance."
        'ElseIf bMagReact = True Then
        '    txtMagnetic.ToolTipText = "Higher Magnetic Reactance properties are preferred."
        'ElseIf bQuantum = True Then
        '    txtMagnetic.ToolTipText = "Quantum-sensitive materials are preferred."
        'Else : txtMagnetic.ToolTipText = "We are not quite sure of the best way to engineer this."
        'End If

        ''Chemical: chem react, compress high
        'If bChemReact = True AndAlso bCompress = True Then
        tpChemical.ToolTipText &= vbCrLf & "Highly Compressible materials that are very Chemically Reactive are better."
        'ElseIf bChemReact = True Then
        '    txtChemical.ToolTipText = "Chemically Reactive materials are preferred."
        'ElseIf bCompress = True Then
        '    txtChemical.ToolTipText = "Compressible materials are better."
        'Else : txtChemical.ToolTipText = "We are not quite sure of the best way to engineer this."
        'End If

        ''Burn: combust low, density high
        'If bCombust = True AndAlso bDensity = True Then
        tpBurn.ToolTipText &= vbCrLf & "Dense materials that are not Combustive are preferred."
        'ElseIf bCombust = True Then
        '    txtBurn.ToolTipText = "Materials that are not combustive are preferred."
        'ElseIf bDensity = True Then
        '    txtBurn.ToolTipText = "Dense materials will aid in burn resistance."
        'Else : txtBurn.ToolTipText = "We are not quite sure of the best way to engineer this."
        'End If

        ''Radar: refract high
        'If bRefract = True Then
        'tpRadar.ToolTipText &= vbCrLf & "Higher Refractive properties are better."
        'Else : txtRadar.ToolTipText = "We are not quite sure of the best way to engineer this."
        'End If

        stmInner.ToolTipText = "Depends on the Resistance Values used."
        stmMiddle.ToolTipText = "Depends on the Resistance Values used."
        stmOuter.ToolTipText = "Depends on the Resistance Values used."
    End Sub

    Private Sub btnDelete_Click(ByVal sName As String) Handles btnDelete.Click
        If moTech Is Nothing Then Return
        If btnDelete.Caption = "Delete Design" Then
            btnDelete.Caption = "CONFIRM"
        Else
            'Delete the design
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

    Private Sub frmArmorBuilder_OnRender() Handles Me.OnRender
        'Ok, first, is moTech nothing?
        If moTech Is Nothing = False Then
            'Ok, now, do the values used on the form currently match the tech?
            Dim bChanged As Boolean = False
            With moTech

                bChanged = tpBeam.PropertyValue <> .BeamResist OrElse tpImpact.PropertyValue <> .ImpactResist OrElse _
                  tpPiercing.PropertyValue <> .PiercingResist OrElse tpMagnetic.PropertyValue <> .MagneticResist OrElse _
                  tpChemical.PropertyValue <> .ChemicalResist OrElse tpBurn.PropertyValue <> .BurnResist OrElse _
                  tpHullUsage.PropertyValue <> .HullUsagePerPlate OrElse tpHPPerPlate.PropertyValue <> .HPPerPlate

                If bChanged = False Then
                    If txtTechName.Caption <> .sArmorName AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eInvalidDesign AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
                        bChanged = True
                    End If
                End If

                If bChanged = False Then
                    'Dim lID1 As Int32 = -1
                    'Dim lID2 As Int32 = -1
                    'Dim lID3 As Int32 = -1

                    'If cboOuter.ListIndex <> -1 Then lID1 = cboOuter.ItemData(cboOuter.ListIndex)
                    'If cboMiddle.ListIndex <> -1 Then lID2 = cboMiddle.ItemData(cboMiddle.ListIndex)
                    'If cboInner.ListIndex <> -1 Then lID3 = cboInner.ItemData(cboInner.ListIndex)

                    bChanged = .OuterLayerMineralID <> mlMineralIDs(0) OrElse .MiddleLayerMineralID <> mlMineralIDs(1) OrElse _
                        .InnerLayerMineralID <> mlMineralIDs(2) OrElse (tpOuter.PropertyLocked <> (.lSpecifiedMin1 <> -1)) OrElse _
                          (tpMiddle.PropertyLocked <> (.lSpecifiedMin2 <> -1)) OrElse (tpInner.PropertyLocked <> (.lSpecifiedMin3 <> -1))

                    If bChanged = False Then

                        bChanged = (.lSpecifiedMin1 <> -1 AndAlso .lSpecifiedMin1 <> tpOuter.PropertyValue) OrElse _
                                    (.lSpecifiedMin2 <> -1 AndAlso .lSpecifiedMin2 <> tpMiddle.PropertyValue) OrElse _
                                    (.lSpecifiedMin3 <> -1 AndAlso .lSpecifiedMin3 <> tpInner.PropertyValue)

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
        goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eArmorBuilder)
    End Sub

    'Private Sub cboInner_DropDownExpanded(ByVal sComboBoxName As String) Handles cboInner.DropDownExpanded

    '    Dim lResistValues(5) As Int32       '0 to 5 for 6 total, 0 = combust, 1 = density, 2 = electres, 3 = malleable, 4 = thermcond, 5 = thermexp
    '    Dim bValueUsed(5) As Boolean

    '    For X As Int32 = 0 To 5
    '        lResistValues(X) = 0
    '        bValueUsed(X) = False
    '    Next X

    '    If chkBeamHilite.Value = True Then
    '        lResistValues(2) += 1 : bValueUsed(2) = True
    '        lResistValues(5) -= 1 : bValueUsed(5) = True
    '    End If
    '    If chkImpactHiLite.Value = True Then
    '        lResistValues(3) += 1 : bValueUsed(3) = True
    '    End If
    '    If chkMagneticHiLite.Value = True Then
    '        lResistValues(1) += 1 : bValueUsed(1) = True
    '    End If
    '    If chkBurnHiLite.Value = True Then
    '        lResistValues(0) -= 1 : bValueUsed(0) = True
    '        lResistValues(4) -= 1 : bValueUsed(4) = True
    '        lResistValues(5) -= 1 : bValueUsed(5) = True
    '    End If

    '    Try
    '        Dim ofrmMin As frmMinDetail = CType(MyBase.moUILib.GetWindow("frmMinDetail"), frmMinDetail)
    '        If ofrmMin Is Nothing Then Return

    '        ofrmMin.ClearHighlights()

    '        For X As Int32 = 0 To 5
    '            If bValueUsed(X) = True Then

    '                Dim sPropName As String = ""
    '                Select Case X
    '                    Case 0 : sPropName = "COMBUSTIVENESS"
    '                    Case 1 : sPropName = "DENSITY"
    '                    Case 2 : sPropName = "ELECTRICAL RESISTANCE"
    '                    Case 3 : sPropName = "MALLEABLE"
    '                    Case 4 : sPropName = "THERMAL CONDUCTION"
    '                    Case 5 : sPropName = "THERMAL EXPANSION"
    '                End Select
    '                If sPropName = "" Then Continue For

    '                For Y As Int32 = 0 To glMineralPropertyUB
    '                    If glMineralPropertyIdx(Y) <> -1 Then
    '                        If goMineralProperty(Y).MineralPropertyName.ToUpper = sPropName Then
    '                            If lResistValues(X) = 0 Then
    '                                ofrmMin.HighlightProperty(goMineralProperty(Y).ObjectID, 1, 5)
    '                            ElseIf lResistValues(X) > 0 Then
    '                                ofrmMin.HighlightProperty(goMineralProperty(Y).ObjectID, 2, 10)
    '                            Else
    '                                ofrmMin.HighlightProperty(goMineralProperty(Y).ObjectID, 3, 0)
    '                            End If
    '                            Exit For
    '                        End If
    '                    End If
    '                Next Y
    '            End If
    '        Next X
    '    Catch
    '    End Try
    'End Sub

    'Private Sub cboMiddle_DropDownExpanded(ByVal sComboBoxName As String) Handles cboMiddle.DropDownExpanded
    '    Dim lResistValues(6) As Int32       '0 to 6 for 7 total
    '    Dim bValueUsed(6) As Boolean        '0 = chemreact, 1 = combust, 2 = compress, 3 = density, 4 = magprod, 5 = thermcond, 6= thermexp

    '    For X As Int32 = 0 To 6
    '        lResistValues(X) = 0
    '        bValueUsed(X) = False
    '    Next X

    '    If chkBeamHilite.Value = True Then
    '        lResistValues(5) -= 1 : bValueUsed(5) = True
    '        lResistValues(6) -= 1 : bValueUsed(6) = True
    '    End If
    '    If chkImpactHiLite.Value = True Then
    '        lResistValues(3) += 1 : bValueUsed(3) = True
    '    End If
    '    If chkPiercingHiLite.Value = True Then
    '        lResistValues(2) -= 1 : bValueUsed(2) = True
    '    End If
    '    If chkMagneticHiLite.Value = True Then
    '        lResistValues(4) += 1 : bValueUsed(4) = True
    '    End If
    '    If chkChemicalHiLite.Value = True Then
    '        lResistValues(0) -= 1 : bValueUsed(0) = True
    '        lResistValues(3) += 1 : bValueUsed(3) = True
    '    End If
    '    If chkBurnHiLite.Value = True Then
    '        lResistValues(1) -= 1 : bValueUsed(1) = True
    '        lResistValues(3) += 1 : bValueUsed(3) = True
    '        lResistValues(5) -= 1 : bValueUsed(5) = True
    '        lResistValues(6) += 1 : bValueUsed(6) = True
    '    End If

    '    Try
    '        Dim ofrmMin As frmMinDetail = CType(MyBase.moUILib.GetWindow("frmMinDetail"), frmMinDetail)
    '        If ofrmMin Is Nothing Then Return

    '        ofrmMin.ClearHighlights()

    '        For X As Int32 = 0 To 6
    '            If bValueUsed(X) = True Then

    '                Dim sPropName As String = ""
    '                Select Case X
    '                    Case 0 : sPropName = "CHEMICAL REACTANCE"
    '                    Case 1 : sPropName = "COMBUSTIVENESS"
    '                    Case 2 : sPropName = "COMPRESSIBILITY"
    '                    Case 3 : sPropName = "DENSITY"
    '                    Case 4 : sPropName = "MAGNETIC PRODUCTION"
    '                    Case 5 : sPropName = "THERMAL CONDUCTION"
    '                    Case 6 : sPropName = "THERMAL EXPANSION"
    '                End Select
    '                If sPropName = "" Then Continue For

    '                For Y As Int32 = 0 To glMineralPropertyUB
    '                    If glMineralPropertyIdx(Y) <> -1 Then
    '                        If goMineralProperty(Y).MineralPropertyName.ToUpper = sPropName Then
    '                            If lResistValues(X) = 0 Then
    '                                ofrmMin.HighlightProperty(goMineralProperty(Y).ObjectID, 1, 5)
    '                            ElseIf lResistValues(X) > 0 Then
    '                                ofrmMin.HighlightProperty(goMineralProperty(Y).ObjectID, 2, 10)
    '                            Else
    '                                ofrmMin.HighlightProperty(goMineralProperty(Y).ObjectID, 3, 0)
    '                            End If
    '                            Exit For
    '                        End If
    '                    End If
    '                Next Y
    '            End If
    '        Next X
    '    Catch
    '    End Try
    'End Sub

    'Private Sub cboOuter_DropDownExpanded(ByVal sComboBoxName As String) Handles cboOuter.DropDownExpanded
    '    Dim lResistValues(9) As Int32       '0 to 9 for 10 total.. 0 = chemreact, 1 = combust, 2 = compress, 3 = density, 4 = hardness
    '    Dim bValueUsed(9) As Boolean        '5= magprod, 6 = magreact, 7 = malleable, 8 = reflect, 9 = refract

    '    For X As Int32 = 0 To 9
    '        lResistValues(X) = 0
    '        bValueUsed(X) = False
    '    Next X

    '    If chkBeamHilite.Value = True Then
    '        lResistValues(5) -= 1 : bValueUsed(5) = True
    '        lResistValues(8) += 1 : bValueUsed(8) = True
    '        lResistValues(9) += 1 : bValueUsed(9) = True
    '    End If
    '    If chkImpactHiLite.Value = True Then
    '        lResistValues(2) += 1 : bValueUsed(2) = True
    '        lResistValues(3) -= 1 : bValueUsed(3) = True
    '        lResistValues(4) += 1 : bValueUsed(4) = True
    '        lResistValues(6) -= 1 : bValueUsed(6) = True
    '    End If
    '    If chkPiercingHiLite.Value = True Then
    '        lResistValues(3) += 1 : bValueUsed(3) = True
    '        lResistValues(4) += 1 : bValueUsed(4) = True
    '        lResistValues(7) -= 1 : bValueUsed(7) = True
    '    End If
    '    If chkMagneticHiLite.Value = True Then
    '        lResistValues(3) -= 1 : bValueUsed(3) = True
    '        lResistValues(5) += 1 : bValueUsed(5) = True
    '        lResistValues(6) -= 1 : bValueUsed(6) = True
    '        lResistValues(7) -= 1 : bValueUsed(7) = True
    '    End If
    '    If chkChemicalHiLite.Value = True Then
    '        lResistValues(0) -= 1 : bValueUsed(0) = True
    '        lResistValues(3) += 1 : bValueUsed(3) = True
    '    End If
    '    If chkBurnHiLite.Value = True Then
    '        lResistValues(1) -= 1 : bValueUsed(1) = True
    '        lResistValues(3) += 1 : bValueUsed(3) = True
    '    End If

    '    Try
    '        Dim ofrmMin As frmMinDetail = CType(MyBase.moUILib.GetWindow("frmMinDetail"), frmMinDetail)
    '        If ofrmMin Is Nothing Then Return

    '        ofrmMin.ClearHighlights()

    '        For X As Int32 = 0 To 9
    '            If bValueUsed(X) = True Then

    '                Dim sPropName As String = ""
    '                Select Case X
    '                    Case 0 : sPropName = "CHEMICAL REACTANCE"
    '                    Case 1 : sPropName = "COMBUSTIVENESS"
    '                    Case 2 : sPropName = "COMPRESSIBILITY"
    '                    Case 3 : sPropName = "DENSITY"
    '                    Case 4 : sPropName = "HARDNESS"
    '                    Case 5 : sPropName = "MAGNETIC PRODUCTION"
    '                    Case 6 : sPropName = "MAGNETIC REACTANCE"
    '                    Case 7 : sPropName = "MALLEABLE"
    '                    Case 8 : sPropName = "REFLECTION"
    '                    Case 9 : sPropName = "REFRACTION"
    '                End Select
    '                If sPropName = "" Then Continue For

    '                For Y As Int32 = 0 To glMineralPropertyUB
    '                    If glMineralPropertyIdx(Y) <> -1 Then
    '                        If goMineralProperty(Y).MineralPropertyName.ToUpper = sPropName Then
    '                            If lResistValues(X) = 0 Then
    '                                ofrmMin.HighlightProperty(goMineralProperty(Y).ObjectID, 1, 5)
    '                            ElseIf lResistValues(X) > 0 Then
    '                                ofrmMin.HighlightProperty(goMineralProperty(Y).ObjectID, 2, 10)
    '                            Else
    '                                ofrmMin.HighlightProperty(goMineralProperty(Y).ObjectID, 3, 0)
    '                            End If
    '                            Exit For
    '                        End If
    '                    End If
    '                Next Y
    '            End If
    '        Next X
    '    Catch
    '    End Try
    'End Sub

    'Private Sub ComboExpandeded(ByVal sComboBoxName As String) Handles cboInner.DropDownExpanded, cboOuter.DropDownExpanded, cboMiddle.DropDownExpanded
    '    If sComboBoxName = "" Then Return

    '    Dim oTech As New ArmorTechComputer
    '    With oTech
    '        .blBeam = tpBeam.PropertyValue
    '        .blBurn = tpBurn.PropertyValue
    '        .blHPPerPlate = tpHPPerPlate.PropertyValue
    '        .blHullUsagePerPlate = tpHullUsage.PropertyValue
    '        .blImpact = tpImpact.PropertyValue
    '        .blMagnetic = tpMagnetic.PropertyValue
    '        .blPiercing = tpPiercing.PropertyValue
    '        .blToxic = tpChemical.PropertyValue
    '    End With

    '    Dim lMinIdx As Int32 = -1
    '    Select Case sComboBoxName
    '        Case cboOuter.ControlName
    '            lMinIdx = 0
    '        Case cboMiddle.ControlName
    '            lMinIdx = 1
    '        Case cboInner.ControlName
    '            lMinIdx = 2
    '    End Select
    '    If lMinIdx = -1 Then Return
    '    oTech.MineralCBOExpanded(lMinIdx, ObjectType.eArmorTech)
    'End Sub

    Public Function ValidateTutorialConfig() As Boolean
        If tpBeam.PropertyValue <> 1 Then Return False
        If tpBurn.PropertyValue <> 1 Then Return False
        If tpChemical.PropertyValue <> 1 Then Return False
        If tpHPPerPlate.PropertyValue <> 10 Then Return False
        If tpHullUsage.PropertyValue <> 1 Then Return False
        If tpImpact.PropertyValue <> 1 Then Return False
        If tpMagnetic.PropertyValue <> 1 Then Return False
        If tpPiercing.PropertyValue <> 1 Then Return False
        'If tpRadar.PropertyValue <> 1 Then Return False
        If txtTechName.Caption.Trim.Length = 0 Then Return False

        If tpOuter.PropertyLocked = True OrElse tpMiddle.PropertyLocked = True OrElse tpInner.PropertyLocked = True Then
            MyBase.moUILib.AddNotification("Do NOT lock the Outer, Middle and Inner properties.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If

        Dim blResTime As Int64 = -1
        Dim blProdCost As Int64 = -1
        Dim blProdTime As Int64 = -1
        MyBase.mfrmBuilderCost.GetLockedValues(0, 0, 0, blResTime, blProdCost, blProdTime, 0, 0, 0)

        If blResTime <> -1 Then
            MyBase.moUILib.AddNotification("Do NOT lock the Research Time Property.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If
        If blProdCost <> -1 Then
            MyBase.moUILib.AddNotification("Do NOT lock the Production Cost Property.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If
        If blProdTime <> -1 Then
            MyBase.moUILib.AddNotification("Do NOT lock the Production Time Property.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End If

        For X As Int32 = 0 To mlMineralIDs.GetUpperBound(0)
            If mlMineralIDs(X) <> 157 Then Return False
        Next

        Return True
    End Function

    Protected Overrides Sub BuilderCostValueChange(ByVal bAutoFill As Boolean)
        If mbIgnoreValueChange = True Then Return
        mbIgnoreValueChange = True
        mbImpossibleDesign = True
        SetBuildPhase()

        Dim oArmorTechComputer As New ArmorTechComputer
        With oArmorTechComputer
            .blBeam = tpBeam.PropertyValue
            .blBurn = tpBurn.PropertyValue
            .blHPPerPlate = tpHPPerPlate.PropertyValue
            .blHullUsagePerPlate = tpHullUsage.PropertyValue
            .blImpact = tpImpact.PropertyValue
            .blMagnetic = tpMagnetic.PropertyValue
            .blPiercing = tpPiercing.PropertyValue
            .blToxic = tpChemical.PropertyValue

            .lHullTypeID = 0

            'If cboOuter.ListIndex > -1 Then
            '    .lMineral1ID = cboOuter.ItemData(cboOuter.ListIndex)
            'Else : .lMineral1ID = -1
            'End If
            'If cboMiddle.ListIndex > -1 Then
            '    .lMineral2ID = cboMiddle.ItemData(cboMiddle.ListIndex)
            'Else : .lMineral2ID = -1
            'End If
            'If cboInner.ListIndex > -1 Then
            '    .lMineral3ID = cboInner.ItemData(cboInner.ListIndex)
            'Else : .lMineral3ID = -1
            'End If
            .lMineral1ID = mlMineralIDs(0)
            .lMineral2ID = mlMineralIDs(1)
            .lMineral3ID = mlMineralIDs(2)
            .lMineral4ID = -1 : .lMineral5ID = -1 : .lMineral6ID = -1


            If .lMineral1ID < 1 OrElse .lMineral2ID < 1 OrElse .lMineral3ID < 1 Then
                lblDesignFlaw.Caption = "All properties and materials need to be defined."
                mbIgnoreValueChange = False
                mbImpossibleDesign = True
                Return
            End If

            If .blHullUsagePerPlate > 0 Then
                If .blHPPerPlate / .blHullUsagePerPlate > 30 Then
                    lblDesignFlaw.Caption = "Your scientists tell you that this is an impossible design."
                    mbIgnoreValueChange = False
                    mbImpossibleDesign = True
                    Return
                End If
            End If


            .msMin1Name = "Outer"
            .msMin2Name = "Middle"
            .msMin3Name = "Inner"
            .msMin4Name = ""
            .msMin5Name = ""
            .msMin6Name = ""

            .ArmorBuilderCostValueChange(lblDesignFlaw, MyBase.mfrmBuilderCost, tpOuter, tpMiddle, tpInner, Nothing, Nothing, Nothing, ObjectType.eArmorTech, 50D)
            mbImpossibleDesign = .bImpossibleDesign
            Dim lIntegrity As Int32 = .GetEstimatedIntegrity()
            If lIntegrity = 0 Then
                lblArmorIntegrity.Caption = "This armor is too brittle to work correctly."
                'mbImpossibleDesign = true
                'lblArmorIntegrity.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            Else
                lblArmorIntegrity.Caption = "Estimated Armor Integrity: " & lIntegrity.ToString & "%"
            End If
            If lIntegrity = 100 Then
                lblArmorIntegrity.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            Else
                lblArmorIntegrity.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            End If
            If lblArmorIntegrity.Caption <> "" AndAlso lblArmorIntegrity.Visible = False Then lblArmorIntegrity.Visible = True

        End With
        mbIgnoreValueChange = False
    End Sub

    Public Overloads Overrides Sub ViewResults(ByRef oTech As Base_Tech, ByVal lProdFactor As Integer)
        If oTech Is Nothing Then Return

        moTech = CType(oTech, ArmorTech)

        Me.btnDesign.Enabled = False

        With moTech
            mbIgnoreValueChange = True
            tpBeam.PropertyValue = .BeamResist
            If tpBeam.PropertyValue <> 0 Then tpBeam.PropertyLocked = True
            tpBurn.PropertyValue = .BurnResist
            If tpBurn.PropertyValue <> 0 Then tpBurn.PropertyLocked = True
            tpChemical.PropertyValue = .ChemicalResist
            If tpChemical.PropertyValue <> 0 Then tpChemical.PropertyLocked = True
            tpHPPerPlate.PropertyValue = .HPPerPlate
            If tpHPPerPlate.PropertyValue <> 0 Then tpHPPerPlate.PropertyLocked = True
            tpHullUsage.PropertyValue = .HullUsagePerPlate
            If tpHullUsage.PropertyValue <> 0 Then tpHullUsage.PropertyLocked = True
            tpImpact.PropertyValue = .ImpactResist
            If tpImpact.PropertyValue <> 0 Then tpImpact.PropertyLocked = True
            tpMagnetic.PropertyValue = .MagneticResist
            If tpMagnetic.PropertyValue <> 0 Then tpMagnetic.PropertyLocked = True
            tpPiercing.PropertyValue = .PiercingResist
            If tpPiercing.PropertyValue <> 0 Then tpPiercing.PropertyLocked = True
            'tpRadar.PropertyValue = .RadarResist
            'If tpRadar.PropertyValue <> 0 Then tpRadar.PropertyLocked = True

            Me.txtTechName.Caption = .sArmorName

            mlSelectedMineralIdx = 0 : Mineral_Selected(.OuterLayerMineralID)
            mlSelectedMineralIdx = 1 : Mineral_Selected(.MiddleLayerMineralID)
            mlSelectedMineralIdx = 2 : Mineral_Selected(.InnerLayerMineralID)
            'If .OuterLayerMineralID > 0 Then Me.cboOuter.FindComboItemData(.OuterLayerMineralID)
            'If .MiddleLayerMineralID > 0 Then Me.cboMiddle.FindComboItemData(.MiddleLayerMineralID)
            'If .InnerLayerMineralID > 0 Then Me.cboInner.FindComboItemData(.InnerLayerMineralID)

            tpOuter.PropertyValue = .lSpecifiedMin1
            If .lSpecifiedMin1 <> -1 Then tpOuter.PropertyLocked = True
            tpMiddle.PropertyValue = .lSpecifiedMin2
            If .lSpecifiedMin2 <> -1 Then tpMiddle.PropertyLocked = True
            tpInner.PropertyValue = .lSpecifiedMin3
            If .lSpecifiedMin3 <> -1 Then tpInner.PropertyLocked = True

            MyBase.mfrmBuilderCost.SetAndLockValues(oTech)

            mbIgnoreValueChange = False
            BuilderCostValueChange(False)

            'lblDesignFlaw.Caption = .GetDesignFlawText() & vbCrLf & "Armor Integrity: " & Math.Max(.lIntegrity, 0).ToString & "%"
            Dim sFlaw As String = .GetDesignFlawText()
            If sFlaw = "" Then
                lblDesignFlaw.Caption = ""
            Else
                lblDesignFlaw.Caption = sFlaw
            End If

            'lblArmorIntegrity.Caption = "Armor Integrity: " & Math.Max(.lIntegrity, 0).ToString & "%"
            'If .lIntegrity = 100 Then
            '    lblArmorIntegrity.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
            'Else
            '    lblArmorIntegrity.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
            'End If

            lblDesignFlaw.Visible = True
        End With

        If oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eComponentResearching AndAlso HasAliasedRights(AliasingRights.eAddResearch) = True Then
            btnDesign.Caption = "Research"
            btnDesign.Enabled = True
        End If
        If gbAliased = False Then
            btnDelete.Visible = True 'oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched
            btnRename.Visible = oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched
            If btnDelete.Visible = True Then
                'Ok, set up the buttons better
                btnDesign.Left = (Me.Width \ 2) - (btnDesign.Width \ 2)
                btnCancel.Left = Me.Width - btnCancel.Width - 5
                btnCancel.Left = Me.Width - 110
                btnDelete.Left = 5
            End If
        End If

        'Now... what state is the tech in?
        If oTech.ComponentDevelopmentPhase > Base_Tech.eComponentDevelopmentPhase.eComponentDesign Then
            'Ok, show it's research and production cost
            If mfrmResCost Is Nothing Then mfrmResCost = New frmProdCost(MyBase.moUILib, "frmResCost")
            With mfrmResCost
                .Visible = True
                .Left = Me.Left + Me.Width + 5
                .Top = Me.Top + Me.Height + 10
                .SetFromProdCost(oTech.oResearchCost, lProdFactor, True, 0, 0)
            End With

            If mfrmProdCost Is Nothing Then mfrmProdCost = New frmProdCost(MyBase.moUILib, "frmProdCost")
            With mfrmProdCost
                .Visible = True
                .Left = mfrmResCost.Left + mfrmResCost.Width + 10
                .Top = mfrmResCost.Top
                .SetFromProdCost(oTech.oProductionCost, 1000, False, 0, 0)
            End With
        ElseIf oTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
            If mfrmResCost Is Nothing Then mfrmResCost = New frmProdCost(MyBase.moUILib, "frmResCost")
            With mfrmResCost
                .Visible = True
                .Left = Me.Left + Me.Width + 5
                .Top = Me.Top + Me.Height + 10
                .SetFromFailureCode(oTech.ErrorCodeReason)
            End With
        End If
    End Sub

    Private Sub SetBuildPhase()
        Dim lPhase As Int32 = 0

        'only thing required for phase 1 is hp and hull
        If tpHullUsage.PropertyValue > 0 AndAlso tpHPPerPlate.PropertyValue > 0 Then lPhase = 1
        If lPhase > 0 Then
            'AndAlso cboOuter.ListIndex > -1 AndAlso cboMiddle.ListIndex > -1 AndAlso cboInner.ListIndex > -1 Then lPhase = 2
            lPhase = 2
            For X As Int32 = 0 To 2
                If mlMineralIDs(X) < 1 Then
                    lPhase = 1
                    Exit For
                End If
            Next X
        End If

        Dim bPhase1 As Boolean = lPhase > 0
        Dim bPhase2 As Boolean = lPhase > 1
        If moTech Is Nothing = False Then lPhase = 2

        If stmOuter.Visible <> bPhase1 Then stmOuter.Visible = bPhase1
        If stmMiddle.Visible <> bPhase1 Then stmMiddle.Visible = bPhase1
        If stmInner.Visible <> bPhase1 Then stmInner.Visible = bPhase1
        If lblOuter.Visible <> bPhase1 Then lblOuter.Visible = bPhase1
        If lblMiddle.Visible <> bPhase1 Then lblMiddle.Visible = bPhase1
        If lblInner.Visible <> bPhase1 Then lblInner.Visible = bPhase1
        If tpOuter.Visible <> bPhase1 Then tpOuter.Visible = bPhase1
        If tpMiddle.Visible <> bPhase1 Then tpMiddle.Visible = bPhase1
        If tpInner.Visible <> bPhase1 Then tpInner.Visible = bPhase1

        If mfrmBuilderCost.Visible <> bPhase2 Then mfrmBuilderCost.Visible = bPhase2

    End Sub


    Private mlMineralIDs(2) As Int32
    Private Sub SetButtonClicked(ByVal lMinIdx As Int32)
        Dim oTech As New ArmorTechComputer

        With oTech
            .blBeam = tpBeam.PropertyValue
            .blBurn = tpBurn.PropertyValue
            .blHPPerPlate = tpHPPerPlate.PropertyValue
            .blHullUsagePerPlate = tpHullUsage.PropertyValue
            .blImpact = tpImpact.PropertyValue
            .blMagnetic = tpMagnetic.PropertyValue
            .blPiercing = tpPiercing.PropertyValue
            .blToxic = tpChemical.PropertyValue
        End With

        If lMinIdx = -1 Then Return

        mlSelectedMineralIdx = -1
        oTech.MineralCBOExpanded(lMinIdx, ObjectType.eArmorTech)
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
                    stmOuter.SetMineralName(sMinName)
                Case 1
                    stmMiddle.SetMineralName(sMinName)
                Case 2
                    stmInner.SetMineralName(sMinName)
            End Select
        End If
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
End Class