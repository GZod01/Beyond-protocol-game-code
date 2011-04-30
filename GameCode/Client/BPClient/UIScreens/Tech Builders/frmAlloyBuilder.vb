Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmAlloyBuilder
	Inherits UIWindow

	Private Enum ePresetSelect As Int32
		BeamCasing = 0
		BeamCoil
		BeamCoupler
		BeamFocuser
		BeamMedium
		EngineDriveBody
		EngineDriveFrame
		EngineDriveMeld
		EngineStructureBody
		EngineStructureFrame
		EngineStructureMeld
		HullStructural
		ProjectileBarrel
		ProjectileCasing
		ProjectilePayload1
		PulseAccelerator
		PulseCasing
		PulseChamber
		PulseCoil
		PulseFocuser
		RadarCasing
		RadarCollection
		RadarDetection
		RadarEmitter
		ShieldAccelerator
		ShieldCasing
        ShieldCoil
        MissileBody
        MissileNose
        MissileFuel
        MissileFlaps
        BombCasing
        BombGuidance
        BombPayload
    End Enum

    Private sPrefix() As String = {"Alu", "Arg", "Ber", "Bor", "Calc", "Carb", "Chlor", "Chro", "Copp", _
        "Enoch", "Flo", "Floyl", "Flour", "Hel", "Hyyl", "Irin", "Iryl", "Jan", "Lith", "Mag", "Man", _
        "Net", "Nit", "Nyl", "Oxy", "Phos", "Plat", "Po", "Rak", "Sil", "Sod", "Sulf", "Tam", "Ti"}
    Private sSuffix() As String = {"ganon", "inese", "mer", "miner", "on", "um", "ine", "tassum", _
        "ese", "inur", "rour", "taner", "ylese", "minur", "ous", "roum", "ic", "idum", "tanous", _
        "tassur", "manur", "moer", "inine", "ron", "phoron", "tanine", "er", "ider", "gano", "minous", _
        "ganon", "gen", "icese", "idine", "tasser", "uron", "urur", "ylium", "assum", "ganine", "inum", _
        "ganium", "roon", "droum", "icine", "idous", "inium", "nesum", "nesur", "tanine", "tassine", _
        "urn", "minine", "neson", "inon", "tanur", "idine", "idium", "idur", "phorum", "roium", "ganer", _
        "inous", "minn", "phorn", "roese", "icium", "phorine", "kat", "tum"}

	Private lblSelMats As UILabel
    Private lblProp1 As UILabel
    Private lblProp2 As UILabel
    Private lblProp3 As UILabel
    Private lblResLvl As UILabel
    Private lblProp1Act As UILabel
    Private cboResLvl As UIComboBox
    Private WithEvents cboMinProp3 As UIComboBox
    Private WithEvents cboMinProp2 As UIComboBox
    Private WithEvents cboMinProp1 As UIComboBox
    Private stmMin1 As ctlSetTechMineral
    Private stmMin2 As ctlSetTechMineral
    Private stmMin3 As ctlSetTechMineral
    Private stmMin4 As ctlSetTechMineral
    'Private cboMin4 As UIComboBox
    'Private cboMin3 As UIComboBox
    'Private cboMin2 As UIComboBox
    'private cboMin1 As UIComboBox
    Private lblTitle As UILabel
	'Private WithEvents optHigh1 As UIOption
	'Private WithEvents optHigh2 As UIOption
	'Private WithEvents optHigh3 As UIOption
	'Private WithEvents optLow1 As UIOption
	'Private WithEvents optLow2 As UIOption
	'Private WithEvents optLow3 As UIOption
    Private WithEvents btnDesign As UIButton
    Private WithEvents btnCancel As UIButton
    Private WithEvents btnDelete As UIButton
    Private WithEvents btnRename As UIButton
    Private WithEvents btnRandom As UIButton
    Private WithEvents lblAlloyName As UILabel
	Private WithEvents txtAlloyName As UITextBox
	Private cboMinPropVal1 As UIComboBox
	Private cboMinPropVal2 As UIComboBox
	Private cboMinPropVal3 As UIComboBox

    Private lblDefaultSelect As UILabel
    Private WithEvents cboDefaultHullType As UIComboBox
	Private WithEvents cboDefaultSelect As UIComboBox
	Private WithEvents cboDefaultTypeSelect As UIComboBox

    Private mlEntityIndex As Int32 = -1

    Private moTech As AlloyTech = Nothing
    Private mfrmResCost As frmProdCost = Nothing
    Private mfrmProdCost As frmProdCost = Nothing
    Private mbIgnoreCboEvents As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmAlloyDesigner initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eAlloyBuilder
            .ControlName = "frmAlloyBuilder"
            .Left = 10
            .Top = 10
            .Width = 512 '640 '783
            .Height = 480 '380
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

        'lblSelMats initial props
        lblSelMats = New UILabel(oUILib)
        With lblSelMats
            .ControlName = "lblSelMats"
            .Left = 42
            .Top = 77
            .Width = 201
            .Height = 17
            .Enabled = True
            .Visible = True
            .Caption = "Select 2 to 4 Minerals/Alloys:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblSelMats, UIControl))

        'lblProp1 initial props
        lblProp1 = New UILabel(oUILib)
        With lblProp1
            .ControlName = "lblProp1"
            .Left = 42
            .Top = 171
            .Width = 181
            .Height = 17
            .Enabled = True
            .Visible = True
            .Caption = "Select a Mineral Property:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblProp1, UIControl))

        'lblProp2 initial props
        lblProp2 = New UILabel(oUILib)
        With lblProp2
            .ControlName = "lblProp2"
            .Left = 42
            .Top = 206
            .Width = 181
            .Height = 17
            .Enabled = True
            .Visible = True
            .Caption = "Select a Mineral Property:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblProp2, UIControl))

        'lblProp3 initial props
        lblProp3 = New UILabel(oUILib)
        With lblProp3
            .ControlName = "lblProp3"
            .Left = 42
            .Top = 241
            .Width = 181
            .Height = 17
            .Enabled = True
            .Visible = True
            .Caption = "Select a Mineral Property:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblProp3, UIControl))

        'lblResLvl initial props
        lblResLvl = New UILabel(oUILib)
        With lblResLvl
            .ControlName = "lblResLvl"
            .Left = 42
            .Top = 296
            .Width = 181
            .Height = 17
            .Enabled = True
            .Visible = True
            .Caption = "Select a Research Level:"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblResLvl, UIControl))

        'lblProp1Act initial props
        lblProp1Act = New UILabel(oUILib)
        With lblProp1Act
            .ControlName = "lblProp1Act"
            .Left = 386
            .Top = 151
			.Width = 100
            .Height = 17
            .Enabled = True
            .Visible = True
			.Caption = "Desired Value"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblProp1Act, UIControl))

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 15
            .Top = 15
            .Width = 201
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = "Alloy Designer"
            '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 20.25, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblTitle, UIControl))

		''optHigh1 initial props
		'optHigh1 = New UIOption(oUILib)
		'With optHigh1
		'    .ControlName = "optHigh1"
		'    .Left = 420
		'    .Top = 171
		'    .Width = 64
		'    .Height = 18
        '    .Enabled = True
		'    .Visible = True
		'    .Caption = "Higher"
		'    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
		'    .ForeColor = muSettings.InterfaceBorderColor
		'    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
		'    .DrawBackImage = False
		'    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
		'    .Value = False
		'    .DisplayType = UIOption.eOptionTypes.eSmallOption
		'End With
		'Me.AddChild(CType(optHigh1, UIControl))

		''optHigh2 initial props
		'optHigh2 = New UIOption(oUILib)
		'With optHigh2
		'    .ControlName = "optHigh2"
		'    .Left = 420
		'    .Top = 205
		'    .Width = 64
		'    .Height = 18
        '    .Enabled = True 'True
		'    .Visible = True
		'    .Caption = "Higher"
		'    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
		'    .ForeColor = muSettings.InterfaceBorderColor
		'    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
		'    .DrawBackImage = False
		'    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
		'    .Value = False
		'    .DisplayType = UIOption.eOptionTypes.eSmallOption
		'End With
		'Me.AddChild(CType(optHigh2, UIControl))

		''optHigh3 initial props
		'optHigh3 = New UIOption(oUILib)
		'With optHigh3
		'    .ControlName = "optHigh3"
		'    .Left = 420
		'    .Top = 241
		'    .Width = 64
		'    .Height = 18
        '    .Enabled = True 'True
		'    .Visible = True
		'    .Caption = "Higher"
		'    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
		'    .ForeColor = muSettings.InterfaceBorderColor
		'    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
		'    .DrawBackImage = False
		'    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
		'    .Value = False
		'    .DisplayType = UIOption.eOptionTypes.eSmallOption
		'End With
		'Me.AddChild(CType(optHigh3, UIControl))

		''optLow1 initial props
		'optLow1 = New UIOption(oUILib)
		'With optLow1
		'    .ControlName = "optLow1"
		'    .Left = 520
		'    .Top = 171
		'    .Width = 56
		'    .Height = 18
        '    .Enabled = True 'True
		'    .Visible = True
		'    .Caption = "Lower"
		'    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
		'    .ForeColor = muSettings.InterfaceBorderColor
		'    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
		'    .DrawBackImage = False
		'    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
		'    .Value = True
		'    .DisplayType = UIOption.eOptionTypes.eSmallOption
		'End With
		'Me.AddChild(CType(optLow1, UIControl))

		''optLow2 initial props
		'optLow2 = New UIOption(oUILib)
		'With optLow2
		'    .ControlName = "optLow2"
		'    .Left = 520
		'    .Top = 205
		'    .Width = 56
		'    .Height = 18
        '    .Enabled = True 'True
		'    .Visible = True
		'    .Caption = "Lower"
		'    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
		'    .ForeColor = muSettings.InterfaceBorderColor
		'    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
		'    .DrawBackImage = False
		'    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
		'    .Value = True
		'    .DisplayType = UIOption.eOptionTypes.eSmallOption
		'End With
		'Me.AddChild(CType(optLow2, UIControl))

		''optLow3 initial props
		'optLow3 = New UIOption(oUILib)
		'With optLow3
		'    .ControlName = "optLow3"
		'    .Left = 520
		'    .Top = 241
		'    .Width = 56
		'    .Height = 18
        '    .Enabled = True 'True
		'    .Visible = True
		'    .Caption = "Lower"
		'    '.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 255)
		'    .ForeColor = muSettings.InterfaceBorderColor
		'    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
		'    .DrawBackImage = False
		'    .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
		'    .Value = True
		'    .DisplayType = UIOption.eOptionTypes.eSmallOption
		'End With
		'Me.AddChild(CType(optLow3, UIControl))

        'lblAlloyName initial props
        lblAlloyName = New UILabel(oUILib)
        With lblAlloyName
            .ControlName = "lblAlloyName"
            .Left = 42
            .Top = 52
            .Width = 85
            .Height = 17
            .Enabled = True
            .Visible = True
            .Caption = "Alloy Name:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblAlloyName, UIControl))

        'txtAlloyName initial props
        txtAlloyName = New UITextBox(oUILib)
        With txtAlloyName
            .ControlName = "txtAlloyName"
            .Left = 130
            .Top = 52
            .Width = 245
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
        Me.AddChild(CType(txtAlloyName, UIControl))

        'btnCancel initial props
        btnCancel = New UIButton(oUILib)
        With btnCancel
            .ControlName = "btnCancel"
            .Left = Me.Width - 105
            .Top = Me.Height - 37
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
            .Left = (Me.Width \ 2) - 50 'btnCancel.Left - 105
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

        'btnRandom initial props
        btnRandom = New UIButton(oUILib)
        With btnRandom
            .ControlName = "btnRandom"
            .Left = txtAlloyName.Left + txtAlloyName.Width + 5
            .Top = txtAlloyName.Top - 1
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Random"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnRandom, UIControl))

        'btnRename initial props
        btnRename = New UIButton(oUILib)
        With btnRename
            .ControlName = "btnRename"
            .Left = txtAlloyName.Left + txtAlloyName.Width + 5
            .Top = btnRandom.Top + 26
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

		'lblDefaultSelect initial props
		lblDefaultSelect = New UILabel(oUILib)
		With lblDefaultSelect
			.ControlName = "lblDefaultSelect"
			.Left = 42
			.Top = 326
			.Width = 150
			.Height = 17
			.Enabled = True
			.Visible = True
			.Caption = "Component Presets:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter
		End With
		Me.AddChild(CType(lblDefaultSelect, UIControl))

        'cboDefaultHullType initial props
        cboDefaultHullType = New UIComboBox(oUILib)
        With cboDefaultHullType
            .ControlName = "cboDefaultHullType"
            .Left = 85
            .Top = 345 '325
            .Width = 125
            .Height = 20
            .Enabled = True
            .Visible = True 'True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(cboDefaultHullType, UIControl))

		'cboDefaultTypeSelect initial props
		cboDefaultTypeSelect = New UIComboBox(oUILib)
		With cboDefaultTypeSelect
			.ControlName = "cboDefaultTypeSelect"
			.Left = 225
            .Top = 345 '325
			.Width = 125
			.Height = 20
            .Enabled = False
            .Visible = True 'True
			.BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(cboDefaultTypeSelect, UIControl))

        'cboDefaultSelect initial props
        cboDefaultSelect = New UIComboBox(oUILib)
        With cboDefaultSelect
            .ControlName = "cboDefaultSelect"
            .Left = 365 '375
            .Top = 345 '325
            .Width = 125
            .Height = 20
            .Enabled = False
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(cboDefaultSelect, UIControl))

        'cboResLvl initial props
        cboResLvl = New UIComboBox(oUILib)
        With cboResLvl
            .ControlName = "cboResLvl"
            .Left = 225
            .Top = 295
            .Width = 125 '150
            .Height = 20
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(cboResLvl, UIControl))

        'cboMinProp3 initial props
        cboMinProp3 = New UIComboBox(oUILib)
        With cboMinProp3
            .ControlName = "cboMinProp3"
            .Left = 225
            .Top = 240
            .Width = 150
            .Height = 20
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .l_ListBoxHeight = 150
        End With
        Me.AddChild(CType(cboMinProp3, UIControl))

        'cboMinPropVal3 initial props
        cboMinPropVal3 = New UIComboBox(oUILib)
        With cboMinPropVal3
            .ControlName = "cboMinPropVal3"
            .Left = cboMinProp3.Left + cboMinProp3.Width + 15
            .Top = cboMinProp3.Top
            .Width = 100 '150
            .Height = 20
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .l_ListBoxHeight = 150
        End With
        Me.AddChild(CType(cboMinPropVal3, UIControl))

        'cboMinProp2 initial props
        cboMinProp2 = New UIComboBox(oUILib)
        With cboMinProp2
            .ControlName = "cboMinProp2"
            .Left = 225
            .Top = 205
            .Width = 150
            .Height = 20
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .l_ListBoxHeight = 150
        End With
        Me.AddChild(CType(cboMinProp2, UIControl))

        'cboMinPropVal2 initial props
        cboMinPropVal2 = New UIComboBox(oUILib)
        With cboMinPropVal2
            .ControlName = "cboMinPropVal2"
            .Left = cboMinProp2.Left + cboMinProp2.Width + 15
            .Top = cboMinProp2.Top
            .Width = 100 '150
            .Height = 20
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .l_ListBoxHeight = 150
        End With
        Me.AddChild(CType(cboMinPropVal2, UIControl))

        'cboMinProp1 initial props
        cboMinProp1 = New UIComboBox(oUILib)
        With cboMinProp1
            .ControlName = "cboMinProp1"
            .Left = 225
            .Top = 170
            .Width = 150
            .Height = 20
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .l_ListBoxHeight = 150
        End With
        Me.AddChild(CType(cboMinProp1, UIControl))

        'cboMinPropVal1 initial props
        cboMinPropVal1 = New UIComboBox(oUILib)
        With cboMinPropVal1
            .ControlName = "cboMinPropVal1"
            .Left = cboMinProp1.Left + cboMinProp1.Width + 15
            .Top = cboMinProp1.Top
            .Width = 100 '150
            .Height = 20
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceTextBoxFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .DropDownListBorderColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .l_ListBoxHeight = 150
        End With
        Me.AddChild(CType(cboMinPropVal1, UIControl))

        ''cboMin4 initial props
        'cboMin4 = New UIComboBox(oUILib)
        'With cboMin4
        '    .ControlName = "cboMin4"
        '    .Left = 225
        '    .Top = 125
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
        'Me.AddChild(CType(cboMin4, UIControl))
        stmMin4 = New ctlSetTechMineral(oUILib)
        With stmMin4
            .ControlName = "stmMin4"
            .Left = 250
            .Top = 125
            .Width = 175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 3
        End With
        Me.AddChild(CType(stmMin4, UIControl))

        ''cboMin3 initial props
        'cboMin3 = New UIComboBox(oUILib)
        'With cboMin3
        '    .ControlName = "cboMin3"
        '    .Left = 65
        '    .Top = 125
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
        'Me.AddChild(CType(cboMin3, UIControl))
        stmMin3 = New ctlSetTechMineral(oUILib)
        With stmMin3
            .ControlName = "stmMin3"
            .Left = 65
            .Top = 125
            .Width = 175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 2
        End With
        Me.AddChild(CType(stmMin3, UIControl))

        ''cboMin2 initial props
        'cboMin2 = New UIComboBox(oUILib)
        'With cboMin2
        '    .ControlName = "cboMin2"
        '    .Left = 225
        '    .Top = 100
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
        'Me.AddChild(CType(cboMin2, UIControl))
        stmMin2 = New ctlSetTechMineral(oUILib)
        With stmMin2
            .ControlName = "stmMin2"
            .Left = 250
            .Top = 100
            .Width = 175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 1
        End With
        Me.AddChild(CType(stmMin2, UIControl))

        ''cboMin1 initial props
        'cboMin1 = New UIComboBox(oUILib)
        'With cboMin1
        '    .ControlName = "cboMin1"
        '    .Left = 65
        '    .Top = 100
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
        'Me.AddChild(CType(cboMin1, UIControl))
        stmMin1 = New ctlSetTechMineral(oUILib)
        With stmMin1
            .ControlName = "stmMin1"
            .Left = 65
            .Top = 100
            .Width = 175
            .Height = 18
            .Enabled = True
            .Visible = True
            .mlMineralIndex = 0
        End With
        Me.AddChild(CType(stmMin1, UIControl))

        FillValues()

        AddHandler stmMin1.SetButtonClicked, AddressOf SetButtonClicked
        AddHandler stmMin2.SetButtonClicked, AddressOf SetButtonClicked
        AddHandler stmMin3.SetButtonClicked, AddressOf SetButtonClicked
        AddHandler stmMin4.SetButtonClicked, AddressOf SetButtonClicked

        ShowMinDetail(-1)

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        glCurrentEnvirView = CurrentView.eAlloyResearch

        If goFullScreenBackground Is Nothing = False Then goFullScreenBackground.Dispose()
        goFullScreenBackground = Nothing
        goFullScreenBackground = goResMgr.GetTexture("alloybuilder.bmp", GFXResourceManager.eGetTextureType.NoSpecifics, "TechBacks.pak")
    End Sub

    Private Sub FillValues()

        Dim bValue As Boolean = False

        Dim lAlloyPropCnt As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eAlloyBuilderImprovements)
        'Look at our special techs for what properties to display
        bValue = True 'goCurrentPlayer.SpecialTechResearched(33)
        cboMinProp1.Visible = bValue
        lblProp1.Visible = bValue
		'optHigh1.Visible = bValue
		'optLow1.Visible = bValue
		cboMinPropVal1.Visible = bValue
        lblProp1Act.Visible = bValue

		'bValue = lAlloyPropCnt > 1
        cboMinProp2.Visible = bValue
		lblProp2.Visible = bValue
		cboMinPropVal2.Visible = bValue
		'optHigh2.Visible = bValue
		'optLow2.Visible = bValue

		'bValue = lAlloyPropCnt > 2
        cboMinProp3.Visible = bValue
		lblProp3.Visible = bValue
		cboMinPropVal3.Visible = bValue
		'optHigh3.Visible = bValue
		'optLow3.Visible = bValue

		cboMinPropVal1.Clear() : cboMinPropVal2.Clear() : cboMinPropVal3.Clear()
		For X As Int32 = 0 To 10
			cboMinPropVal1.AddItem(X.ToString) : cboMinPropVal1.ItemData(cboMinPropVal1.NewIndex) = X
			cboMinPropVal2.AddItem(X.ToString) : cboMinPropVal2.ItemData(cboMinPropVal2.NewIndex) = X
			cboMinPropVal3.AddItem(X.ToString) : cboMinPropVal3.ItemData(cboMinPropVal3.NewIndex) = X
		Next X

        '    Dim lSorted() As Int32 = GetSortedMineralIdxArray(False)

        '    'Fill our minerals/alloys
        '    cboMin1.Clear() : cboMin2.Clear() : cboMin3.Clear() : cboMin4.Clear()
        '    cboMin1.AddItem("None") : cboMin1.ItemData(cboMin1.NewIndex) = -1
        '    cboMin2.AddItem("None") : cboMin2.ItemData(cboMin2.NewIndex) = -1
        '    cboMin3.AddItem("None") : cboMin3.ItemData(cboMin3.NewIndex) = -1
        '    cboMin4.AddItem("None") : cboMin4.ItemData(cboMin4.NewIndex) = -1
        '    If lSorted Is Nothing = False Then
        '        For X As Int32 = 0 To lSorted.GetUpperBound(0)
        ''If goMinerals(lSorted(X)).IsFullyStudied = False Then Continue For
        '            If lSorted(X) <> -1 Then
        '                cboMin1.AddItem(goMinerals(lSorted(X)).MineralName)
        '                cboMin1.ItemData(cboMin1.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '                cboMin2.AddItem(goMinerals(lSorted(X)).MineralName)
        '                cboMin2.ItemData(cboMin2.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '                cboMin3.AddItem(goMinerals(lSorted(X)).MineralName)
        '                cboMin3.ItemData(cboMin3.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '                cboMin4.AddItem(goMinerals(lSorted(X)).MineralName)
        '                cboMin4.ItemData(cboMin4.NewIndex) = goMinerals(lSorted(X)).ObjectID
        '            End If
        '        Next X
        '    End If

        'Fill our mineral properties
        cboMinProp1.Clear() : cboMinProp2.Clear() : cboMinProp3.Clear()
        cboMinProp1.AddItem("None") : cboMinProp1.ItemData(cboMinProp1.NewIndex) = 0
        cboMinProp2.AddItem("None") : cboMinProp2.ItemData(cboMinProp2.NewIndex) = 0
        cboMinProp3.AddItem("None") : cboMinProp3.ItemData(cboMinProp3.NewIndex) = 0
        For X As Int32 = 0 To glMineralPropertyUB
            If glMineralPropertyIdx(X) <> -1 AndAlso goMineralProperty(X).yKnowledgeLevel = eMinPropKnowledgeLevel.eExpertKnowledge Then
                cboMinProp1.AddItem(goMineralProperty(X).MineralPropertyName)
                cboMinProp1.ItemData(cboMinProp1.NewIndex) = goMineralProperty(X).ObjectID
                cboMinProp2.AddItem(goMineralProperty(X).MineralPropertyName)
                cboMinProp2.ItemData(cboMinProp2.NewIndex) = goMineralProperty(X).ObjectID
                cboMinProp3.AddItem(goMineralProperty(X).MineralPropertyName)
                cboMinProp3.ItemData(cboMinProp3.NewIndex) = goMineralProperty(X).ObjectID
            End If
        Next X
        cboMinProp1.ListIndex = 0
        cboMinProp2.ListIndex = 0
        cboMinProp3.ListIndex = 0

		'Fill the research levels
		Dim lMaxLvl As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eMaxAlloyImprovement)

        cboResLvl.Clear()
        cboResLvl.AddItem("Gradient")
		cboResLvl.ItemData(cboResLvl.NewIndex) = 0
		If lMaxLvl >= 1 Then
			cboResLvl.AddItem("Cutting Edge")
			cboResLvl.ItemData(cboResLvl.NewIndex) = 1
		End If
		If lMaxLvl >= 2 Then
			cboResLvl.AddItem("Revolutionary")
			cboResLvl.ItemData(cboResLvl.NewIndex) = 2
		End If
		If lMaxLvl >= 3 Then
			cboResLvl.AddItem("Epic")
			cboResLvl.ItemData(cboResLvl.NewIndex) = 3
		End If
		If lMaxLvl >= 4 Then
			cboResLvl.AddItem("Empire")
			cboResLvl.ItemData(cboResLvl.NewIndex) = 4
		End If
		cboResLvl.FindComboItemData(0)

		cboDefaultSelect.Clear()
		cboDefaultTypeSelect.Clear()
		With cboDefaultTypeSelect
			.AddItem("Solid Beam") : .ItemData(.NewIndex) = 0
			.AddItem("Engine") : .ItemData(.NewIndex) = 1
            .AddItem("Hull") : .ItemData(.NewIndex) = 2
            .AddItem("Missile") : .ItemData(.NewIndex) = 7
			.AddItem("Projectile") : .ItemData(.NewIndex) = 3
			.AddItem("Pulse Beam") : .ItemData(.NewIndex) = 4
			.AddItem("Radar") : .ItemData(.NewIndex) = 5
            .AddItem("Shield") : .ItemData(.NewIndex) = 6
        End With

        Dim lMaxPowerThrust As Int32 = 10000
        If goCurrentPlayer Is Nothing = False Then
            lMaxPowerThrust = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eEnginePowerThrustLimit)
        End If

        cboDefaultHullType.Clear()
        If lMaxPowerThrust > 110000 Then
            cboDefaultHullType.AddItem("Battlecruiser") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.BattleCruiser
        End If
        If lMaxPowerThrust > 400000 Then
            cboDefaultHullType.AddItem("Battleship") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.Battleship
        End If
        cboDefaultHullType.AddItem("Corvette") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.Corvette
        If lMaxPowerThrust > 57000 Then
            cboDefaultHullType.AddItem("Cruiser") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.Cruiser
        End If
        If lMaxPowerThrust > 32000 Then
            cboDefaultHullType.AddItem("Destroyer") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.Destroyer
        End If
        cboDefaultHullType.AddItem("Escort") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.Escort
        cboDefaultHullType.AddItem("Facility") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.Facility
        cboDefaultHullType.AddItem("Frigate") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.Frigate
        cboDefaultHullType.AddItem("Fighter (Light)") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.LightFighter
        cboDefaultHullType.AddItem("Fighter (Medium)") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.MediumFighter
        cboDefaultHullType.AddItem("Fighter (Heavy)") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.HeavyFighter
        Dim bHasNavalUnit As Boolean = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eNavalAbility) > 0
        If bHasNavalUnit = True Then
            If lMaxPowerThrust > 181000 Then
                cboDefaultHullType.AddItem("Naval Battleship") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.NavalBattleship
                cboDefaultHullType.AddItem("Naval Carrier") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.NavalCarrier
            End If
            If lMaxPowerThrust > 83000 Then
                cboDefaultHullType.AddItem("Naval Cruiser") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.NavalCruiser
            End If
            If lMaxPowerThrust > 31000 Then
                cboDefaultHullType.AddItem("Naval Destroyer") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.NavalDestroyer
            End If
            If lMaxPowerThrust > 15000 Then
                cboDefaultHullType.AddItem("Naval Frigate") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.NavalFrigate
                If lMaxPowerThrust > 42000 Then
                    cboDefaultHullType.AddItem("Naval Submarine") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.NavalSub
                End If
            End If
            cboDefaultHullType.AddItem("Naval Utility") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.Utility
        End If
        cboDefaultHullType.AddItem("Small Vehicle") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.SmallVehicle
        cboDefaultHullType.AddItem("Space Station") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.SpaceStation
        cboDefaultHullType.AddItem("Tank") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.Tank
        cboDefaultHullType.AddItem("Utility") : cboDefaultHullType.ItemData(cboDefaultHullType.NewIndex) = EngineTech.eyHullType.Utility
    End Sub

	Private Sub SendAlloyBuilderDesign(ByVal sName As String) Handles btnDesign.Click
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

        'Dim lMinID(3) As Int32
		Dim yProp(2) As Byte
        'lMinID(0) = -1 : lMinID(1) = -1 : lMinID(2) = -1 : lMinID(3) = -1
		yProp(0) = 0 : yProp(1) = 0 : yProp(2) = 0

        'If cboMin1.ListIndex > -1 Then lMinID(0) = cboMin1.ItemData(cboMin1.ListIndex)
        'If cboMin2.ListIndex > -1 Then lMinID(1) = cboMin2.ItemData(cboMin2.ListIndex)
        'If cboMin3.ListIndex > -1 Then lMinID(2) = cboMin3.ItemData(cboMin3.ListIndex)
        'If cboMin4.ListIndex > -1 Then lMinID(3) = cboMin4.ItemData(cboMin4.ListIndex)

		'Check the alloyname
		If txtAlloyName.Caption.Trim.Length = 0 Then
			MyBase.moUILib.AddNotification("You must specify a name for this alloy.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		'Now... compact and check our indices
		For X As Int32 = 0 To 3
            If mlMineralIDs(X) < 1 Then
                For Y As Int32 = X + 1 To 3
                    If mlMineralIDs(Y) > 0 Then
                        mlMineralIDs(X) = mlMineralIDs(Y)
                        mlMineralIDs(Y) = -1
                        Exit For
                    End If
                Next Y
            End If
		Next X
		Dim bGood As Boolean = True
		For X As Int32 = 0 To 3
            If mlMineralIDs(X) <> -1 Then
                If mlMineralIDs(X) < 1 Then
                    mlMineralIDs(X) = -1
                    Continue For
                End If
                For Y As Int32 = 0 To 3
                    If X <> Y AndAlso mlMineralIDs(X) = mlMineralIDs(Y) AndAlso mlMineralIDs(Y) > 0 Then
                        bGood = False
                        Exit For
                    End If
                Next Y
            End If
        Next X

		If bGood = False Then
			MyBase.moUILib.AddNotification("Minerals/Alloys used must be different!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
        End If

        Dim bRes As Boolean = True
        For X As Int32 = 0 To 3
            If mlMineralIDs(X) > -1 Then
                For Y As Int32 = 0 To glMineralUB
                    If glMineralIdx(Y) = mlMineralIDs(X) Then
                        If goMinerals(Y).bDiscovered = False Then
                            bRes = False
                        End If
                        Exit For
                    End If
                Next Y
                If bRes = False Then Exit For
            End If
        Next X
        If bRes = False Then
            MyBase.moUILib.AddNotification("You must select minerals that you have researched.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

		'Ok, now... the properties
		If cboMinProp1.ListIndex > -1 Then yProp(0) = CByte(cboMinProp1.ItemData(cboMinProp1.ListIndex))
		If cboMinProp2.ListIndex > -1 Then yProp(1) = CByte(cboMinProp2.ItemData(cboMinProp2.ListIndex))
		If cboMinProp3.ListIndex > -1 Then yProp(2) = CByte(cboMinProp3.ItemData(cboMinProp3.ListIndex))

        Dim bBadProps As Boolean = True
        For X As Int32 = 0 To 2
            If yProp(X) <> 0 Then
                bBadProps = False
                Exit For
            End If
        Next X
        If bBadProps = True Then
            MyBase.moUILib.AddNotification("The desired value must be selected.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

		'Compact our properties
		For X As Int32 = 0 To 2
			If yProp(X) < 1 Then
				For Y As Int32 = X + 1 To 2
					If yProp(Y) > 0 Then
						yProp(X) = yProp(Y)
						yProp(Y) = 0
						Exit For
					End If
				Next Y
			End If
		Next X
		bGood = True
		For X As Int32 = 0 To 2
			If yProp(X) > 0 Then
				For Y As Int32 = 0 To 2
					If X <> Y AndAlso yProp(X) = yProp(Y) Then
						bGood = False
						Exit For
					End If
				Next Y
			End If
		Next X
		If bGood = False Then
			MyBase.moUILib.AddNotification("You must select different mineral properties!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		Dim lMinCnt As Int32 = 0
		Dim lPropCnt As Int32 = 0
		For X As Int32 = 0 To 3
            If mlMineralIDs(X) <> -1 Then lMinCnt += 1
		Next X
		If lMinCnt < 2 Then
			MyBase.moUILib.AddNotification("You must select at least two different minerals.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If
		For X As Int32 = 0 To 2
			If yProp(X) <> 0 Then lPropCnt += 1
		Next X
		If lPropCnt < 1 Then
			MyBase.moUILib.AddNotification("You must select at least one property to manipulate.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
		End If

		Dim yVal1 As Byte = 255
		Dim yVal2 As Byte = 255
		Dim yVal3 As Byte = 255

		If cboMinPropVal1.ListIndex > -1 Then yVal1 = CByte(cboMinPropVal1.ItemData(cboMinPropVal1.ListIndex))
		If cboMinPropVal2.ListIndex > -1 Then yVal2 = CByte(cboMinPropVal2.ItemData(cboMinPropVal2.ListIndex))
		If cboMinPropVal3.ListIndex > -1 Then yVal3 = CByte(cboMinPropVal3.ItemData(cboMinPropVal3.ListIndex))

		Dim yResLvl As Byte
		If cboResLvl.ListIndex > -1 Then yResLvl = CByte(cboResLvl.ItemData(cboResLvl.ListIndex)) Else yResLvl = CByte(0)

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

        MyBase.moUILib.GetMsgSys().SubmitAlloyDesign(oResearchGuid, txtAlloyName.Caption, mlMineralIDs(0), mlMineralIDs(1), mlMineralIDs(2), mlMineralIDs(3), yProp(0), yProp(1), yProp(2), yVal1, yVal2, yVal3, yResLvl, lTechID)

		If NewTutorialManager.TutorialOn = True Then
			NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eSubmitDesignClick, ObjectType.eAlloyTech, -1, -1, "")
		End If

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddNotification("Alloy Prototype Design Submitted", Color.Green, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

		ReturnToPreviousViewAndReleaseBackground()
	End Sub

	'Private Sub optHigh1_Click() Handles optHigh1.Click
	'	optLow1.Value = False
	'End Sub

	'Private Sub optHigh2_Click() Handles optHigh2.Click
	'	optLow2.Value = False
	'End Sub

	'Private Sub optHigh3_Click() Handles optHigh3.Click
	'	optLow3.Value = False
	'End Sub

	'Private Sub optLow1_Click() Handles optLow1.Click
	'	optHigh1.Value = False
	'End Sub

	'Private Sub optLow2_Click() Handles optLow2.Click
	'	optHigh2.Value = False
	'End Sub

	'Private Sub optLow3_Click() Handles optLow3.Click
	'	optHigh3.Value = False
	'End Sub

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

	'Private Sub cboMin1_ItemSelected(ByVal lItemIndex As Integer) Handles cboMin1.ItemSelected
	'    If mbIgnoreCboEvents = True Then Return
	'    If cboMin1.ListIndex > -1 Then
	'        ShowMinDetail(cboMin1.ItemData(cboMin1.ListIndex))
	'    End If
	'End Sub

	'Private Sub cboMin2_ItemSelected(ByVal lItemIndex As Integer) Handles cboMin2.ItemSelected
	'    If mbIgnoreCboEvents = True Then Return
	'    If cboMin2.ListIndex > -1 Then
	'        ShowMinDetail(cboMin2.ItemData(cboMin2.ListIndex))
	'    End If
	'End Sub

	'Private Sub cboMin3_ItemSelected(ByVal lItemIndex As Integer) Handles cboMin3.ItemSelected
	'    If mbIgnoreCboEvents = True Then Return
	'    If cboMin3.ListIndex > -1 Then
	'        ShowMinDetail(cboMin3.ItemData(cboMin3.ListIndex))
	'    End If
	'End Sub

	'Private Sub cboMin4_ItemSelected(ByVal lItemIndex As Integer) Handles cboMin4.ItemSelected
	'    If mbIgnoreCboEvents = True Then Return
	'    If cboMin4.ListIndex > -1 Then
	'        ShowMinDetail(cboMin4.ItemData(cboMin4.ListIndex))
	'    End If
	'End Sub

	Private Sub ShowMinDetail(ByVal lMineralID As Int32)
		Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
		If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
		ofrm.ShowMineralDetail(Me.Left + Me.Width + 10, Me.Top, Me.Height, lMineralID)
	End Sub

	Private Sub ShowMinDetail(ByRef oMineral As Mineral)
		Dim ofrm As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
		If ofrm Is Nothing Then ofrm = New frmMinDetail(goUILib)
		ofrm.ShowMineralDetail(Me.Left + Me.Width + 10, Me.Top, Me.Height, oMineral)
	End Sub

	Public Sub ViewResults(ByRef oTech As AlloyTech, ByVal lProdFactor As Int32)
		If oTech Is Nothing Then Return

		moTech = oTech

		Me.btnDesign.Enabled = False

		mbIgnoreCboEvents = True

		Me.txtAlloyName.Caption = oTech.sAlloyName

		'optHigh1.Value = False : optLow1.Value = False
		'optHigh2.Value = False : optLow2.Value = False
		'optHigh3.Value = False : optLow3.Value = False
		'If oTech.bHigher1 = True Then optHigh1.Value = True Else optLow1.Value = True
		'If oTech.bHigher2 = True Then optHigh2.Value = True Else optLow2.Value = True
		'If oTech.bHigher3 = True Then optHigh3.Value = True Else optLow3.Value = True
		If oTech.yValue1 <> 255 Then cboMinPropVal1.FindComboItemData(oTech.yValue1)
		If oTech.yValue2 <> 255 Then cboMinPropVal2.FindComboItemData(oTech.yValue2)
		If oTech.yValue3 <> 255 Then cboMinPropVal3.FindComboItemData(oTech.yValue3)

		'now, find our minerals
        'If oTech.Mineral1ID > 0 Then cboMin1.FindComboItemData(oTech.Mineral1ID)
        'If oTech.Mineral2ID > 0 Then cboMin2.FindComboItemData(oTech.Mineral2ID)
        'If oTech.Mineral3ID > 0 Then cboMin3.FindComboItemData(oTech.Mineral3ID)
        'If oTech.Mineral4ID > 0 Then cboMin4.FindComboItemData(oTech.Mineral4ID)
        mlSelectedMineralIdx = 0 : Mineral_Selected(oTech.Mineral1ID)
        mlSelectedMineralIdx = 1 : Mineral_Selected(oTech.Mineral2ID)
        mlSelectedMineralIdx = 2 : Mineral_Selected(oTech.Mineral3ID)
        mlSelectedMineralIdx = 3 : Mineral_Selected(oTech.Mineral4ID)

		'Now, find our props
		If oTech.PropertyID1 > 0 Then cboMinProp1.FindComboItemData(oTech.PropertyID1)
		If oTech.PropertyID2 > 0 Then cboMinProp2.FindComboItemData(oTech.PropertyID2)
		If oTech.PropertyID3 > 0 Then cboMinProp3.FindComboItemData(oTech.PropertyID3)

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
				btnDelete.Left = 5
			End If
		End If

		'And the research level...
		cboResLvl.FindComboItemData(oTech.ResearchLevel)

		If oTech.oExpectedResult Is Nothing = False Then
			ShowMinDetail(oTech.oExpectedResult)
		ElseIf oTech.AlloyResultID > 0 Then
			ShowMinDetail(oTech.AlloyResultID)
		End If

		'Now... what state is the tech in?
		If oTech.ComponentDevelopmentPhase > Base_Tech.eComponentDevelopmentPhase.eComponentDesign Then
			'Ok, show it's research and production cost
			If mfrmResCost Is Nothing Then mfrmResCost = New frmProdCost(MyBase.moUILib, "frmResCost")
			With mfrmResCost
				.Visible = True
				.Left = Me.Left
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
				.Left = Me.Left
				.Top = Me.Top + Me.Height + 10
				.SetFromFailureCode(oTech.ErrorCodeReason)
			End With
		End If

		mbIgnoreCboEvents = False

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

                'ok, see if it has a mineral with it
                Dim lMinResult As Int32 = moTech.AlloyResultID
                If lMinResult > -1 Then
                    For X As Int32 = 0 To glMineralUB
                        If glMineralIdx(X) = lMinResult Then
                            goMinerals(X).bArchived = True
                            Exit For
                        End If
                    Next X
                End If
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

    Private Sub frmAlloyBuilder_OnRender() Handles Me.OnRender
        'Ok, first, is moTech nothing?
        If moTech Is Nothing = False Then
            'Ok, now, do the values used on the form currently match the tech?
            Dim bChanged As Boolean = False
            With moTech
				'bChanged = (.bHigher1 <> optHigh1.Value) OrElse (.bHigher2 <> optHigh2.Value) OrElse (.bHigher3 <> optHigh3.Value)
				If cboMinPropVal1.ListIndex > -1 Then
					If .yValue1 <> cboMinPropVal1.ItemData(cboMinPropVal1.ListIndex) Then bChanged = True
				ElseIf .yValue1 <> 255 Then
					bChanged = True
				End If
				If cboMinPropVal2.ListIndex > -1 Then
					If .yValue2 <> cboMinPropVal2.ItemData(cboMinPropVal2.ListIndex) Then bChanged = True
				ElseIf .yValue2 <> 255 Then
					bChanged = True
				End If
				If cboMinPropVal3.ListIndex > -1 Then
					If .yValue3 <> cboMinPropVal3.ItemData(cboMinPropVal3.ListIndex) Then bChanged = True
				ElseIf .yValue3 <> 255 Then
					bChanged = True
                End If

                If cboResLvl.ListIndex > -1 Then
                    If .ResearchLevel <> cboResLvl.ItemData(cboResLvl.ListIndex) Then bChanged = True
                End If
                If txtAlloyName.Caption <> .sAlloyName AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eInvalidDesign AndAlso .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
                    bChanged = True
                End If
                If bChanged = False Then
                    Dim lID1 As Int32 = mlMineralIDs(0)
                    Dim lID2 As Int32 = mlMineralIDs(1)
                    Dim lID3 As Int32 = mlMineralIDs(2)
                    Dim lID4 As Int32 = mlMineralIDs(3)

                    'If cboMin1.ListIndex <> -1 Then lID1 = cboMin1.ItemData(cboMin1.ListIndex)
                    'If cboMin2.ListIndex <> -1 Then lID2 = cboMin2.ItemData(cboMin2.ListIndex)
                    'If cboMin3.ListIndex <> -1 Then lID3 = cboMin3.ItemData(cboMin3.ListIndex)
                    'If cboMin4.ListIndex <> -1 Then lID4 = cboMin4.ItemData(cboMin4.ListIndex)

                    bChanged = .Mineral1ID <> lID1 OrElse .Mineral2ID <> lID2 OrElse .Mineral3ID <> lID3 OrElse .Mineral4ID <> lID4

                    If bChanged = False Then
                        Dim lP1 As Int32 = 0
                        Dim lP2 As Int32 = 0
                        Dim lP3 As Int32 = 0

                        If cboMinProp1.ListIndex <> -1 Then lP1 = cboMinProp1.ItemData(cboMinProp1.ListIndex)
                        If cboMinProp2.ListIndex <> -1 Then lP2 = cboMinProp2.ItemData(cboMinProp2.ListIndex)
                        If cboMinProp3.ListIndex <> -1 Then lP3 = cboMinProp3.ItemData(cboMinProp3.ListIndex)

                        bChanged = .PropertyID1 <> lP1 OrElse .PropertyID2 <> lP2 OrElse .PropertyID3 <> lP3
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

	Private Sub cboDefaultSelect_ItemSelected(ByVal lItemIndex As Integer) Handles cboDefaultSelect.ItemSelected

        Dim oTech As TechBuilderComputer = Nothing
        Dim lMineralIdx As Int32 = -1
        Dim iTechTypeID As Int16 = -1
        Select Case CType(cboDefaultSelect.ItemData(cboDefaultSelect.ListIndex), ePresetSelect)
            Case ePresetSelect.BeamCasing
                oTech = New SolidTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 2
            Case ePresetSelect.BeamCoil
                oTech = New SolidTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 0
            Case ePresetSelect.BeamCoupler
                oTech = New SolidTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 1
            Case ePresetSelect.BeamFocuser
                oTech = New SolidTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 3
            Case ePresetSelect.BeamMedium
                oTech = New SolidTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 4
            Case ePresetSelect.BombCasing
                oTech = New BombTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 0
            Case ePresetSelect.BombGuidance
                oTech = New BombTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 1
            Case ePresetSelect.BombPayload
                oTech = New BombTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 2
            Case ePresetSelect.EngineDriveBody
                oTech = New EngineTechComputer()
                iTechTypeID = ObjectType.eEngineTech
                lMineralIdx = 3
            Case ePresetSelect.EngineDriveFrame
                oTech = New EngineTechComputer()
                iTechTypeID = ObjectType.eEngineTech
                lMineralIdx = 4
            Case ePresetSelect.EngineDriveMeld
                oTech = New EngineTechComputer()
                iTechTypeID = ObjectType.eEngineTech
                lMineralIdx = 5
            Case ePresetSelect.EngineStructureBody
                oTech = New EngineTechComputer()
                iTechTypeID = ObjectType.eEngineTech
                lMineralIdx = 0
            Case ePresetSelect.EngineStructureFrame
                oTech = New EngineTechComputer()
                iTechTypeID = ObjectType.eEngineTech
                lMineralIdx = 1
            Case ePresetSelect.EngineStructureMeld
                oTech = New EngineTechComputer()
                iTechTypeID = ObjectType.eEngineTech
                lMineralIdx = 2
            Case ePresetSelect.HullStructural
                Dim ofrmMin As frmMinDetail = CType(MyBase.moUILib.GetWindow("frmMinDetail"), frmMinDetail)
                If ofrmMin Is Nothing Then Return
                ofrmMin.ClearHighlights()
                Dim sNames(0) As String : Dim yHighlights(0) As Byte : Dim yValues(0) As Byte
                sNames(0) = "HARDNESS" : yHighlights(0) = 2 : yValues(0) = 10
                Try
                    If sNames Is Nothing = False Then
                        For X As Int32 = 0 To sNames.GetUpperBound(0)
                            For Y As Int32 = 0 To glMineralPropertyUB
                                If glMineralPropertyIdx(Y) <> -1 Then
                                    If sNames(X) = goMineralProperty(Y).MineralPropertyName.ToUpper Then
                                        ofrmMin.HighlightProperty(goMineralProperty(Y).ObjectID, yHighlights(X), yValues(X))
                                        Exit For
                                    End If
                                End If
                            Next Y
                        Next X
                    End If
                Catch
                End Try
                Return
            Case ePresetSelect.ProjectileBarrel
                oTech = New ProjectileTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 0
            Case ePresetSelect.ProjectileCasing
                oTech = New ProjectileTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 1
            Case ePresetSelect.ProjectilePayload1
                oTech = New ProjectileTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 2
            Case ePresetSelect.PulseAccelerator
                oTech = New PulseTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 1
            Case ePresetSelect.PulseCasing
                oTech = New PulseTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 2
            Case ePresetSelect.PulseChamber
                oTech = New PulseTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 4
            Case ePresetSelect.PulseCoil
                oTech = New PulseTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 0
            Case ePresetSelect.PulseFocuser
                oTech = New PulseTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 3
            Case ePresetSelect.RadarCasing
                oTech = New RadarTechComputer()
                iTechTypeID = ObjectType.eRadarTech
                lMineralIdx = 3
            Case ePresetSelect.RadarCollection
                oTech = New RadarTechComputer()
                iTechTypeID = ObjectType.eRadarTech
                lMineralIdx = 2
            Case ePresetSelect.RadarDetection
                oTech = New RadarTechComputer()
                iTechTypeID = ObjectType.eRadarTech
                lMineralIdx = 1
            Case ePresetSelect.RadarEmitter
                oTech = New RadarTechComputer()
                iTechTypeID = ObjectType.eRadarTech
                lMineralIdx = 0
            Case ePresetSelect.ShieldAccelerator
                oTech = New ShieldTechComputer()
                iTechTypeID = ObjectType.eShieldTech
                lMineralIdx = 1
            Case ePresetSelect.ShieldCasing
                oTech = New ShieldTechComputer()
                iTechTypeID = ObjectType.eShieldTech
                lMineralIdx = 2
            Case ePresetSelect.ShieldCoil
                oTech = New ShieldTechComputer()
                iTechTypeID = ObjectType.eShieldTech
                lMineralIdx = 0
            Case ePresetSelect.MissileBody
                oTech = New MissileTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 0
            Case ePresetSelect.MissileFlaps
                oTech = New MissileTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 2
            Case ePresetSelect.MissileFuel
                oTech = New MissileTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 3
            Case ePresetSelect.MissileNose
                oTech = New MissileTechComputer()
                iTechTypeID = ObjectType.eWeaponTech
                lMineralIdx = 1
        End Select

        If oTech Is Nothing = False AndAlso lMineralIdx > -1 AndAlso iTechTypeID > -1 Then
            If cboDefaultHullType.ListIndex > -1 Then
                oTech.lHullTypeID = cboDefaultHullType.ItemData(cboDefaultHullType.ListIndex)
            End If
            oTech.MineralCBOExpanded(lMineralIdx, iTechTypeID)
        End If



        'Dim ofrmMin As frmMinDetail = CType(MyBase.moUILib.GetWindow("frmMinDetail"), frmMinDetail)
        'If ofrmMin Is Nothing Then Return
        'ofrmMin.ClearHighlights()

        'If cboDefaultSelect.ListIndex = -1 Then Return

        'Dim sNames() As String = Nothing
        'Dim yHighlights() As Byte = Nothing
        'Dim yValues() As Byte = Nothing

        'Select Case CType(cboDefaultSelect.ItemData(cboDefaultSelect.ListIndex), ePresetSelect)
        '	Case ePresetSelect.BeamCasing
        '		ReDim sNames(3) : ReDim yHighlights(3) : ReDim yValues(3)
        '		sNames(0) = "DENSITY" : sNames(1) = "TEMPERATURE SENSITIVITY" : sNames(2) = "MALLEABLE" : sNames(3) = "THERMAL CONDUCTION"
        '		yHighlights(0) = 2 : yHighlights(1) = 3 : yHighlights(2) = 3 : yHighlights(3) = 2
        '		yValues(0) = 6 : yValues(1) = 0 : yValues(2) = 0 : yValues(3) = 5
        '	Case ePresetSelect.BeamCoil
        '		ReDim sNames(3) : ReDim yHighlights(3) : ReDim yValues(3)
        '		sNames(0) = "DENSITY" : sNames(1) = "SUPERCONDUCTIVE POINT" : sNames(2) = "MAGNETIC REACTANCE" : sNames(3) = "MAGNETIC PRODUCTION"
        '		yHighlights(0) = 2 : yHighlights(1) = 2 : yHighlights(2) = 2 : yHighlights(3) = 3
        '		yValues(0) = 6 : yValues(1) = 7 : yValues(2) = 7 : yValues(3) = 0
        '	Case ePresetSelect.BeamCoupler
        '		ReDim sNames(3) : ReDim yHighlights(3) : ReDim yValues(3)
        '		sNames(0) = "QUANTUM" : sNames(1) = "REFLECTION" : sNames(2) = "TEMPERATURE SENSITIVITY" : sNames(3) = "THERMAL EXPANSION"
        '		yHighlights(0) = 2 : yHighlights(1) = 2 : yHighlights(2) = 3 : yHighlights(3) = 3
        '		yValues(0) = 4 : yValues(1) = 5 : yValues(2) = 0 : yValues(3) = 0
        '	Case ePresetSelect.BeamFocuser
        '		ReDim sNames(3) : ReDim yHighlights(3) : ReDim yValues(3)
        '		sNames(0) = "QUANTUM" : sNames(1) = "REFRACTION" : sNames(2) = "TEMPERATURE SENSITIVITY" : sNames(3) = "THERMAL EXPANSION"
        '		yHighlights(0) = 2 : yHighlights(1) = 2 : yHighlights(2) = 3 : yHighlights(3) = 3
        '		yValues(0) = 4 : yValues(1) = 6 : yValues(2) = 0 : yValues(3) = 0 
        '	Case ePresetSelect.BeamMedium
        '		ReDim sNames(3) : ReDim yHighlights(3) : ReDim yValues(3)
        '		sNames(0) = "QUANTUM" : sNames(1) = "REFLECTION" : sNames(2) = "REFRACTION" : sNames(3) = "BOILING POINT"
        '		yHighlights(0) = 2 : yHighlights(1) = 3 : yHighlights(2) = 3 : yHighlights(3) = 2
        '		yValues(0) = 5 : yValues(1) = 4 : yValues(2) = 0 : yValues(3) = 0
        '	Case ePresetSelect.EngineDriveBody
        '		ReDim sNames(2) : ReDim yHighlights(2) : ReDim yValues(2)
        '		sNames(0) = "DENSITY" : sNames(1) = "COMPRESSIBILITY" : sNames(2) = "MAGNETIC REACTANCE"
        '		yHighlights(0) = 3 : yHighlights(1) = 3 : yHighlights(2) = 2
        '		yValues(0) = 0 : yValues(1) = 0 : yValues(2) = 7
        '	Case ePresetSelect.EngineDriveFrame
        '		ReDim sNames(2) : ReDim yHighlights(2) : ReDim yValues(2)
        '		sNames(0) = "SUPERCONDUCTIVE POINT" : sNames(1) = "COMPRESSIBILITY" : sNames(2) = "HARDNESS"
        '		yHighlights(0) = 2 : yHighlights(1) = 3 : yHighlights(2) = 2
        '		yValues(0) = 7 : yValues(1) = 0 : yValues(2) = 6 
        '	Case ePresetSelect.EngineDriveMeld
        '		ReDim sNames(2) : ReDim yHighlights(2) : ReDim yValues(2)
        '		sNames(0) = "CHEMICAL REACTANCE" : sNames(1) = "COMBUSTIVENESS" : sNames(2) = "MALLEABLE"
        '		yHighlights(0) = 2 : yHighlights(1) = 3 : yHighlights(2) = 3
        '		yValues(0) = 6 : yValues(1) = 0 : yValues(2) = 0 
        '	Case ePresetSelect.EngineStructureBody
        '		ReDim sNames(2) : ReDim yHighlights(2) : ReDim yValues(2)
        '		sNames(0) = "DENSITY" : sNames(1) = "MAGNETIC PRODUCTION" : sNames(2) = "MELTING POINT"
        '		yHighlights(0) = 2 : yHighlights(1) = 3 : yHighlights(2) = 2
        '		yValues(0) = 6 : yValues(1) = 0 : yValues(2) = 6 
        '	Case ePresetSelect.EngineStructureFrame
        '		ReDim sNames(2) : ReDim yHighlights(2) : ReDim yValues(2)
        '		sNames(0) = "HARDNESS" : sNames(1) = "MALLEABLE" : sNames(2) = "COMPRESSIBILITY"
        '		yHighlights(0) = 2 : yHighlights(1) = 2 : yHighlights(2) = 3
        '		yValues(0) = 6 : yValues(1) = 5 : yValues(2) = 0
        '	Case ePresetSelect.EngineStructureMeld
        '		ReDim sNames(2) : ReDim yHighlights(2) : ReDim yValues(2)
        '		sNames(0) = "HARDNESS" : sNames(1) = "MALLEABLE" : sNames(2) = "MELTING POINT"
        '		yHighlights(0) = 2 : yHighlights(1) = 3 : yHighlights(2) = 2
        '		yValues(0) = 6 : yValues(1) = 0 : yValues(2) = 6 
        '	Case ePresetSelect.HullStructural
        '		ReDim sNames(0) : ReDim yHighlights(0) : ReDim yValues(0)
        '		sNames(0) = "HARDNESS" : yHighlights(0) = 2 : yValues(0) = 10
        '	Case ePresetSelect.ProjectileBarrel
        '		ReDim sNames(2) : ReDim yHighlights(2) : ReDim yValues(2)
        '		sNames(0) = "TEMPERATURE SENSITIVITY" : sNames(1) = "THERMAL EXPANSION" : sNames(2) = "THERMAL CONDUCTION"
        '		yHighlights(0) = 3 : yHighlights(1) = 3 : yHighlights(2) = 2
        '		yValues(0) = 0 : yValues(1) = 0 : yValues(2) = 4
        '	Case ePresetSelect.ProjectileCasing
        '		ReDim sNames(2) : ReDim yHighlights(2) : ReDim yValues(2)
        '		sNames(0) = "DENSITY" : sNames(1) = "TEMPERATURE SENSITIVITY" : sNames(2) = "THERMAL CONDUCTION"
        '		yHighlights(0) = 3 : yHighlights(1) = 2 : yHighlights(2) = 3
        '		yValues(0) = 1 : yValues(1) = 5 : yValues(2) = 0
        '	Case ePresetSelect.ProjectilePayload1
        '		ReDim sNames(2) : ReDim yHighlights(2) : ReDim yValues(2)
        '		sNames(0) = "DENSITY" : sNames(1) = "HARDNESS" : sNames(2) = "MALLEABLE"
        '		yHighlights(0) = 2 : yHighlights(1) = 2 : yHighlights(2) = 3
        '		yValues(0) = 6 : yValues(1) = 4 : yValues(2) = 0
        '	Case ePresetSelect.PulseAccelerator
        '		ReDim sNames(3) : ReDim yHighlights(3) : ReDim yValues(3)
        '		sNames(0) = "QUANTUM" : sNames(1) = "SUPERCONDUCTIVE POINT" : sNames(2) = "MAGNETIC REACTANCE" : sNames(3) = "MAGNETIC PRODUCTION"
        '		yHighlights(0) = 2 : yHighlights(1) = 2 : yHighlights(2) = 2 : yHighlights(3) = 3
        '		yValues(0) = 4 : yValues(1) = 6 : yValues(2) = 6 : yValues(3) = 0
        '	Case ePresetSelect.PulseCasing
        '		ReDim sNames(3) : ReDim yHighlights(3) : ReDim yValues(3)
        '		sNames(0) = "DENSITY" : sNames(1) = "TEMPERATURE SENSITIVITY" : sNames(2) = "MALLEABLE" : sNames(3) = "THERMAL CONDUCTION"
        '		yHighlights(0) = 2 : yHighlights(1) = 3 : yHighlights(2) = 3 : yHighlights(3) = 2
        '		yValues(0) = 5 : yValues(1) = 0 : yValues(2) = 0 : yValues(3) = 4
        '	Case ePresetSelect.PulseChamber
        '		ReDim sNames(3) : ReDim yHighlights(3) : ReDim yValues(3)
        '		sNames(0) = "DENSITY" : sNames(1) = "ELECTRICAL RESISTANCE" : sNames(2) = "COMPRESSIBILITY" : sNames(3) = "MALLEABLE"
        '		yHighlights(0) = 2 : yHighlights(1) = 2 : yHighlights(2) = 3 : yHighlights(3) = 3
        '		yValues(0) = 4 : yValues(1) = 5 : yValues(2) = 0 : yValues(3) = 0
        '	Case ePresetSelect.PulseCoil
        '		ReDim sNames(3) : ReDim yHighlights(3) : ReDim yValues(3)
        '		sNames(0) = "DENSITY" : sNames(1) = "SUPERCONDUCTIVE POINT" : sNames(2) = "MAGNETIC REACTANCE" : sNames(3) = "MAGNETIC PRODUCTION"
        '		yHighlights(0) = 2 : yHighlights(1) = 2 : yHighlights(2) = 2 : yHighlights(3) = 3
        '		yValues(0) = 5 : yValues(1) = 6 : yValues(2) = 0 : yValues(3) = 0
        '	Case ePresetSelect.PulseFocuser
        '		ReDim sNames(3) : ReDim yHighlights(3) : ReDim yValues(3)
        '		sNames(0) = "QUANTUM" : sNames(1) = "REFRACTION" : sNames(2) = "TEMPERATURE SENSITIVITY" : sNames(3) = "THERMAL EXPANSION"
        '		yHighlights(0) = 2 : yHighlights(1) = 2 : yHighlights(2) = 3 : yHighlights(3) = 3
        '		yValues(0) = 4 : yValues(1) = 6 : yValues(2) = 0 : yValues(3) = 0
        '	Case ePresetSelect.RadarCasing
        '		ReDim sNames(4) : ReDim yHighlights(4) : ReDim yValues(4)
        '		sNames(0) = "DENSITY" : sNames(1) = "THERMAL CONDUCTION" : sNames(2) = "THERMAL EXPANSION" : sNames(3) = "MAGNETIC PRODUCTION" : sNames(4) = "MAGNETIC REACTANCE"
        '		yHighlights(0) = 2 : yHighlights(1) = 2 : yHighlights(2) = 3 : yHighlights(3) = 3 : yHighlights(4) = 3
        '		yValues(0) = 4 : yValues(1) = 4 : yValues(2) = 0 : yValues(3) = 0 : yValues(4) = 0
        '	Case ePresetSelect.RadarCollection
        '		ReDim sNames(2) : ReDim yHighlights(2) : ReDim yValues(2)
        '		sNames(0) = "MALLEABLE" : sNames(1) = "ELECTRICAL RESISTANCE" : sNames(2) = "MAGNETIC PRODUCTION"
        '		yHighlights(0) = 2 : yHighlights(1) = 3 : yHighlights(2) = 3
        '		yValues(0) = 4 : yValues(1) = 0 : yValues(2) = 0
        '	Case ePresetSelect.RadarDetection
        '		ReDim sNames(2) : ReDim yHighlights(2) : ReDim yValues(2)
        '		sNames(0) = "SUPERCONDUCTIVE POINT" : sNames(1) = "MAGNETIC REACTANCE" : sNames(2) = "MAGNETIC PRODUCTION"
        '		yHighlights(0) = 2 : yHighlights(1) = 2 : yHighlights(2) = 3
        '		yValues(0) = 5 : yValues(1) = 0 : yValues(2) = 6
        '	Case ePresetSelect.RadarEmitter
        '		ReDim sNames(2) : ReDim yHighlights(2) : ReDim yValues(2)
        '		sNames(0) = "ELECTRICAL RESISTANCE" : sNames(1) = "MAGNETIC PRODUCTION" : sNames(2) = "MAGNETIC REACTANCE"
        '		yHighlights(0) = 3 : yHighlights(1) = 2 : yHighlights(2) = 2
        '		yValues(0) = 0 : yValues(1) = 5 : yValues(2) = 5
        '	Case ePresetSelect.ShieldAccelerator
        '		ReDim sNames(2) : ReDim yHighlights(2) : ReDim yValues(2)
        '		sNames(0) = "QUANTUM" : sNames(1) = "SUPERCONDUCTIVE POINT" : sNames(2) = "MAGNETIC REACTANCE"
        '		yHighlights(0) = 2 : yHighlights(1) = 3 : yHighlights(2) = 2
        '		yValues(0) = 5 : yValues(1) = 0 : yValues(2) = 6 
        '	Case ePresetSelect.ShieldCasing
        '		ReDim sNames(2) : ReDim yHighlights(2) : ReDim yValues(2)
        '		sNames(0) = "DENSITY" : sNames(1) = "THERMAL CONDUCTION" : sNames(2) = "TEMPERATURE SENSITIVITY"
        '		yHighlights(0) = 2 : yHighlights(1) = 2 : yHighlights(2) = 3
        '		yValues(0) = 6 : yValues(1) = 4 : yValues(2) = 0
        '	Case ePresetSelect.ShieldCoil
        '		ReDim sNames(2) : ReDim yHighlights(2) : ReDim yValues(2)
        '		sNames(0) = "DENSITY" : sNames(1) = "SUPERCONDUCTIVE POINT" : sNames(2) = "MAGNETIC PRODUCTION"
        '		yHighlights(0) = 2 : yHighlights(1) = 3 : yHighlights(2) = 2
        '		yValues(0) = 5 : yValues(1) = 0 : yValues(2) = 6 
        'End Select

        'Try
        '	If sNames Is Nothing = False Then
        '		For X As Int32 = 0 To sNames.GetUpperBound(0)
        '			For Y As Int32 = 0 To glMineralPropertyUB
        '				If glMineralPropertyIdx(Y) <> -1 Then
        '					If sNames(X) = goMineralProperty(Y).MineralPropertyName.ToUpper Then
        '						ofrmMin.HighlightProperty(goMineralProperty(Y).ObjectID, yHighlights(X), yValues(X))
        '						Exit For
        '					End If
        '				End If
        '			Next Y
        '		Next X
        '	End If
        'Catch

        'End Try
 
	End Sub

    Private Sub cboDefaultTypeSelect_ItemSelected(ByVal lItemIndex As Integer) Handles cboDefaultTypeSelect.ItemSelected
        Dim lSelect As Int32 = -1
        If cboDefaultSelect.ListIndex > -1 Then lSelect = cboDefaultSelect.ItemData(cboDefaultSelect.ListIndex)

        cboDefaultSelect.Clear()

        If cboDefaultTypeSelect.ListIndex > -1 Then
            cboDefaultSelect.Enabled = True
            With cboDefaultSelect
                Select Case cboDefaultTypeSelect.ItemData(cboDefaultTypeSelect.ListIndex)
                    Case 0
                        .AddItem("Casing") : .ItemData(.NewIndex) = ePresetSelect.BeamCasing
                        .AddItem("Coil") : .ItemData(.NewIndex) = ePresetSelect.BeamCoil
                        .AddItem("Coupler") : .ItemData(.NewIndex) = ePresetSelect.BeamCoupler
                        .AddItem("Focuser") : .ItemData(.NewIndex) = ePresetSelect.BeamFocuser
                        .AddItem("Medium") : .ItemData(.NewIndex) = ePresetSelect.BeamMedium
                    Case 1
                        .AddItem("Drive Body") : .ItemData(.NewIndex) = ePresetSelect.EngineDriveBody
                        .AddItem("Drive Frame") : .ItemData(.NewIndex) = ePresetSelect.EngineDriveFrame
                        .AddItem("Drive Meld") : .ItemData(.NewIndex) = ePresetSelect.EngineDriveMeld
                        .AddItem("Structure Body") : .ItemData(.NewIndex) = ePresetSelect.EngineStructureBody
                        .AddItem("Structure Frame") : .ItemData(.NewIndex) = ePresetSelect.EngineStructureFrame
                        .AddItem("Structure Meld") : .ItemData(.NewIndex) = ePresetSelect.EngineStructureMeld
                    Case 2
                        .AddItem("Structure") : .ItemData(.NewIndex) = ePresetSelect.HullStructural
                    Case 3
                        .AddItem("Barrel") : .ItemData(.NewIndex) = ePresetSelect.ProjectileBarrel
                        .AddItem("Casing") : .ItemData(.NewIndex) = ePresetSelect.ProjectileCasing
                        .AddItem("Payload 1") : .ItemData(.NewIndex) = ePresetSelect.ProjectilePayload1
                    Case 4
                        .AddItem("Accelerator") : .ItemData(.NewIndex) = ePresetSelect.PulseAccelerator
                        .AddItem("Casing") : .ItemData(.NewIndex) = ePresetSelect.PulseCasing
                        .AddItem("Chamber") : .ItemData(.NewIndex) = ePresetSelect.PulseChamber
                        .AddItem("Coil") : .ItemData(.NewIndex) = ePresetSelect.PulseCoil
                        .AddItem("Focuser") : .ItemData(.NewIndex) = ePresetSelect.PulseFocuser
                    Case 5
                        .AddItem("Casing") : .ItemData(.NewIndex) = ePresetSelect.RadarCasing
                        .AddItem("Collector") : .ItemData(.NewIndex) = ePresetSelect.RadarCollection
                        .AddItem("Detector") : .ItemData(.NewIndex) = ePresetSelect.RadarDetection
                        .AddItem("Emitter") : .ItemData(.NewIndex) = ePresetSelect.RadarEmitter
                    Case 6
                        .AddItem("Accelerator") : .ItemData(.NewIndex) = ePresetSelect.ShieldAccelerator
                        .AddItem("Casing") : .ItemData(.NewIndex) = ePresetSelect.ShieldCasing
                        .AddItem("Coil") : .ItemData(.NewIndex) = ePresetSelect.ShieldCoil
                    Case 7
                        .AddItem("Body") : .ItemData(.NewIndex) = ePresetSelect.MissileBody
                        .AddItem("Flaps") : .ItemData(.NewIndex) = ePresetSelect.MissileFlaps
                        .AddItem("Fuel") : .ItemData(.NewIndex) = ePresetSelect.MissileFuel
                        .AddItem("Nose") : .ItemData(.NewIndex) = ePresetSelect.MissileNose
                    Case 8
                        .AddItem("Casing") : .ItemData(.NewIndex) = ePresetSelect.BombCasing
                        .AddItem("Guidance") : .ItemData(.NewIndex) = ePresetSelect.BombGuidance
                        .AddItem("Payload") : .ItemData(.NewIndex) = ePresetSelect.BombPayload
                End Select

                If lSelect = -1 OrElse .FindComboItemData(lSelect) = False Then
                    .ListIndex = -1
                    Dim ofrmMin As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
                    If ofrmMin Is Nothing Then Return
                    ofrmMin.ClearHighlights()
                    ofrmMin = Nothing
                End If
            End With
        End If
    End Sub

    Private Sub cboDefaultHullType_ItemSelected(ByVal lItemIndex As Integer) Handles cboDefaultHullType.ItemSelected
        If lItemIndex > -1 Then
            With cboDefaultTypeSelect
                Dim lSelect As Int32 = -1
                If .ListIndex <> -1 Then lSelect = .ItemData(.ListIndex)

                .Enabled = True
                .Clear()
                If cboDefaultHullType.ItemData(cboDefaultHullType.ListIndex) = EngineTech.eyHullType.Frigate Then
                    .AddItem("Bomb") : .ItemData(.NewIndex) = 8
                End If
                .AddItem("Engine") : .ItemData(.NewIndex) = 1
                .AddItem("Hull") : .ItemData(.NewIndex) = 2
                .AddItem("Missile") : .ItemData(.NewIndex) = 7
                .AddItem("Projectile") : .ItemData(.NewIndex) = 3
                .AddItem("Pulse Beam") : .ItemData(.NewIndex) = 4
                .AddItem("Radar") : .ItemData(.NewIndex) = 5
                .AddItem("Shield") : .ItemData(.NewIndex) = 6
                .AddItem("Solid Beam") : .ItemData(.NewIndex) = 0

                If lSelect = -1 OrElse .FindComboItemData(lSelect) = False Then
                    .ListIndex = -1
                    cboDefaultSelect.Enabled = False
                    cboDefaultSelect.ListIndex = -1
                    Dim ofrmMin As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
                    If ofrmMin Is Nothing Then Return
                    ofrmMin.ClearHighlights()
                    ofrmMin = Nothing
                End If
            End With
        End If
    End Sub


    Private mlMineralIDs(3) As Int32
    Private Sub SetButtonClicked(ByVal lMinIdx As Int32)
        If lMinIdx = -1 Then Return
        Dim lPrevIdx As Int32 = mlSelectedMineralIdx
        mlSelectedMineralIdx = lMinIdx
        Dim ofrmMin As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
        If ofrmMin Is Nothing Then Return

        If ofrmMin.bRaiseSelectEvent = True AndAlso mlSelectedMineralIdx = lPrevIdx Then
            ofrmMin.bRaiseSelectEvent = False
            mlMineralIDs(mlSelectedMineralIdx) = -1
            Dim sMinName As String = "Unselected"
            Select Case mlSelectedMineralIdx
                Case 0
                    stmMin1.SetMineralName(sMinName)
                Case 1
                    stmMin2.SetMineralName(sMinName)
                Case 2
                    stmMin3.SetMineralName(sMinName)
                Case 3
                    stmMin4.SetMineralName(sMinName)
            End Select
            Return
        End If

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
                    stmMin1.SetMineralName(sMinName)
                Case 1
                    stmMin2.SetMineralName(sMinName)
                Case 2
                    stmMin3.SetMineralName(sMinName)
                Case 3
                    stmMin4.SetMineralName(sMinName)
            End Select
        End If
        Me.IsDirty = True
        'BuilderCostValueChange()
    End Sub

    Private msRenameVal As String = ""
    Private Sub btnRename_Click(ByVal sName As String) Handles btnRename.Click
        If moTech Is Nothing = False Then
            If moTech.OwnerID = glPlayerID Then
                If goCurrentPlayer.blCredits < 10000000 Then
                    MyBase.moUILib.AddNotification("You require 10,000,000 credits to rename a tech.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If
                Dim sVal As String = txtAlloyName.Caption.Trim
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
            If moTech.ObjTypeID = ObjectType.eAlloyTech AndAlso goMinerals Is Nothing = False AndAlso moTech.AlloyResultID > 0 Then
                For X As Int32 = 0 To goMinerals.GetUpperBound(0)
                    If goMinerals(X).ObjectID = moTech.AlloyResultID Then
                        goMinerals(X).MineralName = msRenameVal
                        Exit Sub
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub cboMinProp1_ItemSelected(ByVal lItemIndex As Integer) Handles cboMinProp1.ItemSelected
        If cboMinProp1.ItemData(cboMinProp1.ListIndex) = 0 Then
            cboMinPropVal1.ListIndex = 0
        End If
    End Sub

    Private Sub cboMinProp2_ItemSelected(ByVal lItemIndex As Integer) Handles cboMinProp2.ItemSelected
        If cboMinProp2.ItemData(cboMinProp2.ListIndex) = 0 Then
            cboMinPropVal2.ListIndex = 0
        End If
    End Sub

    Private Sub cboMinProp3_ItemSelected(ByVal lItemIndex As Integer) Handles cboMinProp3.ItemSelected
        If cboMinProp3.ItemData(cboMinProp3.ListIndex) = 0 Then
            cboMinPropVal3.ListIndex = 0
        End If
    End Sub

    Private Sub btnRandom_Click(ByVal sName As String) Handles btnRandom.Click
        If sPrefix Is Nothing = True OrElse sSuffix Is Nothing = True OrElse goMinerals Is Nothing = True Then Exit Sub
        Dim oRandom As New Random
        Dim sResult As String = ""
        For lAttempt As Int32 = 1 To 1000
            Dim lPrefix As Int32 = oRandom.Next(sPrefix.GetUpperBound(0))
            Dim lSuffix As Int32 = oRandom.Next(sSuffix.GetUpperBound(0))
            sResult = sPrefix(lPrefix) & sSuffix(lSuffix)
            For X As Int32 = 0 To goMinerals.GetUpperBound(0)
                If goMinerals(X).MineralName = sResult Then
                    If lAttempt < 1000 Then sResult = ""
                    Exit For
                End If
            Next
            If sResult <> "" Then Exit For
        Next
        txtAlloyName.Caption = sResult
    End Sub
End Class

