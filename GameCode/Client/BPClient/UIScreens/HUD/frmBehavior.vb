Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Entity Behavior Window
Public Class frmBehavior
    Inherits UIWindow

    Private lblTargetingPref As UILabel
    Private chkTPFighter As UICheckBox
    Private chkTPCargo As UICheckBox
	Private chkTPTroop As UICheckBox
	Private chkTPEscort As UICheckBox
	Private chkTPCarrier As UICheckBox
	Private chkTPCapital As UICheckBox
	Private chkTPArmed As UICheckBox
	Private chkTPUnarmed As UICheckBox
	Private chkTPFacility As UICheckBox
	'Private lnDiv1 As UILine
	Private lblEngagement As UILabel
	Private WithEvents optEPHoldFire As UIOption
	Private WithEvents optEPStandGround As UIOption
	Private WithEvents optEPPursue As UIOption
	Private WithEvents optEPEvade As UIOption
	Private WithEvents optEPEngage As UIOption
	Private WithEvents optEPDock As UIOption
	'Private lnDiv2 As UILine
	Private lblTactic As UILabel
	Private WithEvents optCTMinimize As UIOption
	Private WithEvents optCTNormal As UIOption
	Private WithEvents optCTMaximize As UIOption
	Private chkCTManeuver As UICheckBox
    Private chkCTLaunchAll As UICheckBox
    Private chkCTDockDuringBattle As UICheckBox
    Private chkStopToEngage As UICheckBox
    Private chkAssistOrbit As UICheckBox
	Private WithEvents btnMinMax As UIButton
	Private WithEvents btnRoutes As UIButton

	Private lblTargetSubsystem As UILabel
	Private WithEvents optEngines As UIOption
	Private WithEvents optHangar As UIOption
	Private WithEvents optRadar As UIOption
	Private WithEvents optShield As UIOption
	Private WithEvents optWeapon As UIOption
	Private WithEvents optNone As UIOption
    Private lblOrders As UILabel

    Private lnWpnGrp1 As UILine
    Private lnWpnGroup2 As UILine
    Private lblWpnGrp As UILabel
    Private chkPrimWpnGrp As UICheckBox
    Private chkSecWpnGrp As UICheckBox
    Private chkPDWpnGrp As UICheckBox

	Private mbIgnoreOptEvents As Boolean
	Private mbIgnoreChkEvents As Boolean
	Private mlEntityIndex As Int32

    Private miLastCT As Int32 = -1
	Private miLastTP As Int16 = -1

	Public lExtendVal1 As Int32
	Public lExtendVal2 As Int32

	Private mbMinimized As Boolean
	Private mlStartX As Int32 = 0
	Private mlStartY As Int32 = 0

	Private mbMultiSelect As Boolean = False

	Private mbTPCheckBoxSet As Boolean = False
	Private mbCTCheckBoxSet As Boolean = False
	Private mbIgnoreOptSubsystmeClick As Boolean = True

	Public Property MultiSelect() As Boolean
		Get
			Return mbMultiSelect
		End Get
		Set(ByVal value As Boolean)
			If mbMultiSelect <> value Then
				Dim ofrm As UIWindow
				If value = True Then
					ofrm = goUILib.GetWindow("frmMultiDisplay")
				Else : ofrm = goUILib.GetWindow("frmSingleSelect")
				End If
				If ofrm Is Nothing = False Then
                    Me.Left = ofrm.Left + ofrm.Width + 1
                    If Me.Left + Me.Width > goUILib.oDevice.PresentationParameters.BackBufferWidth Then
                        Me.Left = ofrm.Left - Me.Width
                    End If
					If value = True Then
						Me.Top = goUILib.oDevice.PresentationParameters.BackBufferHeight - Me.Height
					Else : Me.Top = moUILib.oDevice.PresentationParameters.BackBufferHeight - 260
					End If
				End If
				ofrm = Nothing

				Me.IsDirty = True

				btnMinMax.Visible = value
			End If
			mbMultiSelect = value
		End Set
	End Property

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'frmBehavior initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eBehavior
            .ControlName = "frmBehavior"
            If muSettings.BehaviorLocX = -1 OrElse NewTutorialManager.TutorialOn = True Then .Left = 128 Else .Left = muSettings.BehaviorLocX
            If muSettings.BehaviorLocY = -1 OrElse NewTutorialManager.TutorialOn = True Then .Top = moUILib.oDevice.PresentationParameters.BackBufferHeight - 260 Else .Top = muSettings.BehaviorLocY
            .Width = 512 '312
            .Height = 232

            If .Left + .Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If .Top + .Height > oUILib.oDevice.PresentationParameters.BackBufferHeight Then .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight - .Height

            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(255, 0, 255, 255)
            .BorderColor = muSettings.InterfaceBorderColor
            '.FillColor = System.Drawing.Color.FromArgb(128, 64, 128, 192)
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .BorderLineWidth = 2
        End With
		mlStartX = Me.Left
		mlStartY = Me.Top

		'lblTargetingPref initial props
		lblTargetingPref = New UILabel(oUILib)
		With lblTargetingPref
			.ControlName = "lblTargetingPref"
			.Left = 5
			.Top = 30 '5
			.Width = 153
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Targeting Preference"
			'.ForeColor = System.Drawing.Color.FromArgb(-16711681)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.ToolTipText = "Targeting Preferences Define First Choice in Target Acquisition." & vbCrLf & "  Targets that meet this criteria will be targeted before targets that do not."
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter
		End With
		Me.AddChild(CType(lblTargetingPref, UIControl))

		'chkTPFighter initial props
		chkTPFighter = New UICheckBox(oUILib)
		With chkTPFighter
			.ControlName = "chkTPFighter"
			.Left = 16
			.Top = 50 '29
			.Width = 49
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Fighter"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Any aerial or space-bourne unit that has" & vbCrLf & "a hull size less than or equal to 300."
			.DisplayType = UICheckBox.eCheckTypes.eSmallX
		End With
		Me.AddChild(CType(chkTPFighter, UIControl))

		'chkTPCargo initial props
		chkTPCargo = New UICheckBox(oUILib)
		With chkTPCargo
			.ControlName = "chkTPCargo"
			.Left = 15 '110
			.Top = 170 '29
			.Width = 46	'46
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Cargo"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Any entity that has a third of its hull configured for cargo bay."
			.DisplayType = UICheckBox.eCheckTypes.eSmallX
		End With
		Me.AddChild(CType(chkTPCargo, UIControl))

		'chkTPTroopTrans initial props
		chkTPTroop = New UICheckBox(oUILib)
		With chkTPTroop
			.ControlName = "chkTPTroop"
			.Left = 15 '205
			.Top = 110 '29
			.Width = 50
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Troop Transport"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Troopers, Infantry, Mechanized Infantry and Walkers"
			.DisplayType = UICheckBox.eCheckTypes.eSmallX
		End With
		Me.AddChild(CType(chkTPTroop, UIControl))

		'chkTPEscort initial props
		chkTPEscort = New UICheckBox(oUILib)
		With chkTPEscort
			.ControlName = "chkTPEscort"
			.Left = 16
			.Top = 90 '65
			.Width = 46
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Escort"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Any unit that has a hull size greater than 300 but less than 30,000."
			.DisplayType = UICheckBox.eCheckTypes.eSmallX
		End With
		Me.AddChild(CType(chkTPEscort, UIControl))

		'chkTPCarrier initial props
		chkTPCarrier = New UICheckBox(oUILib)
		With chkTPCarrier
			.ControlName = "chkTPCarrier"
			.Left = 15 '110
			.Top = 150 '65
			.Width = 48
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Carrier"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
            .ToolTipText = "Any entity that has a third of its hull configured for hangar bay."
			.DisplayType = UICheckBox.eCheckTypes.eSmallX
		End With
		Me.AddChild(CType(chkTPCarrier, UIControl))

		'chkTPCapital initial props
		chkTPCapital = New UICheckBox(oUILib)
		With chkTPCapital
			.ControlName = "chkTPCapital"
			.Left = 15 '205
			.Top = 210 '65
			.Width = 90
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Capital"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Massive warships"
			.DisplayType = UICheckBox.eCheckTypes.eSmallX
		End With
		Me.AddChild(CType(chkTPCapital, UIControl))

		'chkTPArmed initial props
		chkTPArmed = New UICheckBox(oUILib)
		With chkTPArmed
			.ControlName = "chkTPArmed"
			.Left = 16
			.Top = 70 '47
			.Width = 47
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Armed"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Has at least 1 weapon on-board. Mutually exclusive with Unarmed."
			.DisplayType = UICheckBox.eCheckTypes.eSmallX
		End With
		Me.AddChild(CType(chkTPArmed, UIControl))

		'chkTPUnarmed initial props
		chkTPUnarmed = New UICheckBox(oUILib)
		With chkTPUnarmed
			.ControlName = "chkTPUnarmed"
			.Left = 15 '110
			.Top = 130 '47
			.Width = 61
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Unarmed"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "No weapons on-board. Mutually exclusive with Armed."
			.DisplayType = UICheckBox.eCheckTypes.eSmallX
		End With
		Me.AddChild(CType(chkTPUnarmed, UIControl))

		'chkTPFacility initial props
		chkTPFacility = New UICheckBox(oUILib)
		With chkTPFacility
			.ControlName = "chkTPFacility"
			.Left = 15 '205
			.Top = 190 '47
			.Width = 50
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Facility"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Facilities and Space Stations"
			.DisplayType = UICheckBox.eCheckTypes.eSmallX
		End With
		Me.AddChild(CType(chkTPFacility, UIControl))

		''lnDiv1 initial props
		'lnDiv1 = New UILine(oUILib)
		'With lnDiv1
		'	.ControlName = "lnDiv1"
		'	.Left = 5
		'	.Top = 85
		'	.Width = 300
		'	.Height = 0
		'	.Enabled = True
		'	.Visible = True
		'	'.BorderColor = System.Drawing.Color.FromArgb(-16711681)
		'	.BorderColor = muSettings.InterfaceBorderColor
		'End With
		'Me.AddChild(CType(lnDiv1, UIControl))

		'lblEngagement initial props
		lblEngagement = New UILabel(oUILib)
		With lblEngagement
			.ControlName = "lblEngagement"
			.Left = 145 '5
			.Top = 30 '90
			.Width = 153
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Engagement Pattern"
			'.ForeColor = System.Drawing.Color.FromArgb(-16711681)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter
			.ToolTipText = "Defines how this object reacts to opposition"
		End With
		Me.AddChild(CType(lblEngagement, UIControl))

		'optEPHoldFire initial props
		optEPHoldFire = New UIOption(oUILib)
		With optEPHoldFire
			.ControlName = "optEPHoldFire"
			.Left = 160 '16
			.Top = 50 '110
			.Width = 59
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Hold Fire"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Will not change behavior upon engagement and will not fight back"
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optEPHoldFire, UIControl))

		'optEPStandGround initial props
		optEPStandGround = New UIOption(oUILib)
		With optEPStandGround
			.ControlName = "optEPStandGround"
			.Left = 160 '203
			.Top = 70 '110
			.Width = 84
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Stand Ground"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Will fight back but will not change current destination or activity"
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optEPStandGround, UIControl))

		'optEPPursue initial props
		optEPPursue = New UIOption(oUILib)
		With optEPPursue
			.ControlName = "optEPPursue"
			.Left = 160
			.Top = 90 '110
			.Width = 50
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Pursue"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Does not move unless target attempts to escape"
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optEPPursue, UIControl))

		'optEPEvade initial props
		optEPEvade = New UIOption(oUILib)
		With optEPEvade
			.ControlName = "optEPEvade"
			.Left = 160 '16
			.Top = 110 '128
			.Width = 48
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Evade"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Will attempt to avoid conflict by moving away"
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optEPEvade, UIControl))

		'optEPEngage initial props
		optEPEngage = New UIOption(oUILib)
		With optEPEngage
			.ControlName = "optEPEngage"
			.Left = 160 '111
			.Top = 130 '128
			.Width = 54
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Engage"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Will openly move and follow targets of opportunity"
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optEPEngage, UIControl))

		'optEPDock initial props
		optEPDock = New UIOption(oUILib)
		With optEPDock
			.ControlName = "optEPDock"
			.Left = 160 '203
			.Top = 150 '128
			.Width = 104
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Dock With Target"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Will attempt to dock with the specified target (if target still exists)"
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optEPDock, UIControl))

		''lnDiv2 initial props
		'lnDiv2 = New UILine(oUILib)
		'With lnDiv2
		'	.ControlName = "lnDiv2"
		'	.Left = 5
		'	.Top = 150
		'	.Width = 300
		'	.Height = 0
		'	.Enabled = True
		'	.Visible = True
		'	'.BorderColor = System.Drawing.Color.FromArgb(-16711681)
		'	.BorderColor = muSettings.InterfaceBorderColor
		'End With
		'Me.AddChild(CType(lnDiv2, UIControl))

		'lblTactic initial props
		lblTactic = New UILabel(oUILib)
		With lblTactic
			.ControlName = "lblTactic"
			.Left = 280 '5
			.Top = 30 '155
			.Width = 153
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Combat Tactic"
			'.ForeColor = System.Drawing.Color.FromArgb(-16711681)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter
			.ToolTipText = "Defines in-combat activities, positioning and facing patterns"
		End With
		Me.AddChild(CType(lblTactic, UIControl))

		'optCTMinimize initial props
		optCTMinimize = New UIOption(oUILib)
		With optCTMinimize
			.ControlName = "optCTMinimize"
			.Left = 290 '16
            .Top = 45 '175
			.Width = 100
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Minimize Damage"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Attempt to face the target with this object's most defended side and keep optimal range"
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optCTMinimize, UIControl))

		'optCTNormal initial props
		optCTNormal = New UIOption(oUILib)
		With optCTNormal
			.ControlName = "optCTNormal"
			.Left = 290 '16
            .Top = 65 '193
			.Width = 51
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Normal"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Will do as much damage as possible from the side that gives a good defense"
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optCTNormal, UIControl))

		'optCTMaximize initial props
		optCTMaximize = New UIOption(oUILib)
		With optCTMaximize
			.ControlName = "optCTMaximize"
			.Left = 290 '16
            .Top = 85 '211
			.Width = 104
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Maximize Damage"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
			.ToolTipText = "Do as much damage as possible to target with no regard to self-preservation"
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optCTMaximize, UIControl))

		'chkCTManeuver initial props
		chkCTManeuver = New UICheckBox(oUILib)
		With chkCTManeuver
			.ControlName = "chkCTManeuver"
			.Left = 290 '205
            .Top = 100 '175
			.Width = 65
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Maneuver"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
            .ToolTipText = "When checked, object will attempt to 'dance' around target to avoid being hit"
			.DisplayType = UICheckBox.eCheckTypes.eSmallX
		End With
		Me.AddChild(CType(chkCTManeuver, UIControl))

		'chkCTLaunchAll initial props
		chkCTLaunchAll = New UICheckBox(oUILib)
		With chkCTLaunchAll
			.ControlName = "chkCTLaunchAll"
			.Left = 290 '205
            .Top = 115 '140 '193
            .Width = 94
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Launch All Units"
			'.ForeColor = System.Drawing.Color.FromArgb(-1)
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
			.Value = False
            .ToolTipText = "When checked, attempts to launch all objects in hangar upon engagement"
			.DisplayType = UICheckBox.eCheckTypes.eSmallX
		End With
        Me.AddChild(CType(chkCTLaunchAll, UIControl))

        'chkCTDockDuringBattle
        chkCTDockDuringBattle = New UICheckBox(oUILib)
        With chkCTDockDuringBattle
            .ControlName = "chkCTDockDuringBattle"
            .Left = 290 '205
            .Top = 130 '160 '160
            .Width = 79 '107
            .Height = 32 '16
            .Enabled = True
            .Visible = True
            .Caption = "Stay Docked" & vbCrLf & "During Battle"
            '.ForeColor = System.Drawing.Color.FromArgb(-1)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
            .Value = False
            .ToolTipText = "When checked, will remain docked in facilities until the hostile alert in the budget window is overriden"
            .DisplayType = UICheckBox.eCheckTypes.eSmallX
        End With
        Me.AddChild(CType(chkCTDockDuringBattle, UIControl))

        'chkStopToEngage
        chkStopToEngage = New UICheckBox(oUILib)
        With chkStopToEngage
            .ControlName = "chkStopToEngage"
            .Left = 290 '205
            .Top = 155 '160 '160
            .Width = 91
            .Height = 32 '16
            .Enabled = True
            .Visible = isAdmin()
            .Caption = "Stop to Engage"
            '.ForeColor = System.Drawing.Color.FromArgb(-1)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
            .Value = False
            .ToolTipText = "When checked, any player initiated unit moves will stop when they detect an enemy.  Upon completion the unit will continue the move."
            .DisplayType = UICheckBox.eCheckTypes.eSmallX
        End With
        Me.AddChild(CType(chkStopToEngage, UIControl))

        'chkAssistOrbit
        chkAssistOrbit = New UICheckBox(oUILib)
        With chkAssistOrbit
            .ControlName = "chkAssistOrbit"
            .Left = 290 '205
            .Top = 175 '160 '160
            .Width = 70 '107
            .Height = 32 '16
            .Enabled = True
            .Visible = isAdmin()
            .Caption = "Assist Orbit"
            '.ForeColor = System.Drawing.Color.FromArgb(-1)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Right
            .Value = False
            .ToolTipText = "Will cause mayhem."
            .DisplayType = UICheckBox.eCheckTypes.eSmallX
        End With
        Me.AddChild(CType(chkAssistOrbit, UIControl))

        'btnMinMax initial props
		btnMinMax = New UIButton(oUILib)
		With btnMinMax
            .ControlImageRect_Normal = grc_UI(elInterfaceRectangle.eDownArrow_Button_Normal)
            .ControlImageRect_Pressed = grc_UI(elInterfaceRectangle.eDownArrow_Button_Down)
            .ControlImageRect_Disabled = grc_UI(elInterfaceRectangle.eDownArrow_Button_Disabled)
			.ControlName = "btnMinMax"
			.Left = Me.Width - 26
			.Top = 2
			.Width = 24
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = ""
			.ForeColor = System.Drawing.Color.FromArgb(-1)
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.ToolTipText = "Click to show/hide AI settings screen"
			.FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
		End With
		Me.AddChild(CType(btnMinMax, UIControl))

		'btnRoutes initial props
		btnRoutes = New UIButton(oUILib)
		With btnRoutes
			.ControlName = "btnRoutes"
			.Left = Me.Width - 126
			.Top = 2
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Routes"
            .ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.ToolTipText = "Click to configure the routes for the selection"
			.FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
		End With
		Me.AddChild(CType(btnRoutes, UIControl))

		'lblTargetSubsystem initial props
		lblTargetSubsystem = New UILabel(oUILib)
		With lblTargetSubsystem
			.ControlName = "lblTargetSubsystem"
			.Left = 400
			.Top = 30
			.Width = 153
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Target Subsystem"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(0, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTargetSubsystem, UIControl))

		'optEngines initial props
		optEngines = New UIOption(oUILib)
		With optEngines
			.ControlName = "optEngines"
			.Left = 415
			.Top = 70
			.Width = 55
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Engines"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(2, DrawTextFormat)
			.Value = False
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optEngines, UIControl))

		'optHangar initial props
		optHangar = New UIOption(oUILib)
		With optHangar
			.ControlName = "optHangar"
			.Left = 415
			.Top = 90
			.Width = 57
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Hangars"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(2, DrawTextFormat)
			.Value = False
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optHangar, UIControl))

		'optRadar initial props
		optRadar = New UIOption(oUILib)
		With optRadar
			.ControlName = "optRadar"
			.Left = 415
			.Top = 110
			.Width = 46
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Radar"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(2, DrawTextFormat)
			.Value = False
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optRadar, UIControl))

		'optShield initial props
		optShield = New UIOption(oUILib)
		With optShield
			.ControlName = "optShield"
			.Left = 415
			.Top = 130
			.Width = 51
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Shields"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(2, DrawTextFormat)
			.Value = False
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optShield, UIControl))

		'optWeapon initial props
		optWeapon = New UIOption(oUILib)
		With optWeapon
			.ControlName = "optWeapon"
			.Left = 415
			.Top = 150
			.Width = 64
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Weapons"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(2, DrawTextFormat)
			.Value = False
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optWeapon, UIControl))

		'optNone initial props
		optNone = New UIOption(oUILib)
		With optNone
			.ControlName = "optNone"
			.Left = 415
			.Top = 50
			.Width = 43
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "None"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(2, DrawTextFormat)
			.Value = True
			.DisplayType = UIOption.eOptionTypes.eSmallOption
		End With
		Me.AddChild(CType(optNone, UIControl))

		'lblOrders initial props
		lblOrders = New UILabel(oUILib)
		With lblOrders
			.ControlName = "lblOrders"
			.Left = 5
			.Top = 5
			.Width = 145
			.Height = 16
			.Enabled = True
			.Visible = True
			.Caption = "Orders Management"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
        Me.AddChild(CType(lblOrders, UIControl))

        'lnWpnGrp1 initial props
        lnWpnGrp1 = New UILine(oUILib)
        With lnWpnGrp1
            .ControlName = "lnWpnGrp1"
            .Left = 145
            .Top = 200
            .Width = 366
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnWpnGrp1, UIControl))

        'lnWpnGroup2 initial props
        lnWpnGroup2 = New UILine(oUILib)
        With lnWpnGroup2
            .ControlName = "lnWpnGroup2"
            .Left = 145
            .Top = 200
            .Width = 0
            .Height = 32
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnWpnGroup2, UIControl))

        'lblWpnGrp initial props
        lblWpnGrp = New UILabel(oUILib)
        With lblWpnGrp
            .ControlName = "lblWpnGrp"
            .Left = 155
            .Top = 206
            .Width = 32
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Use:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblWpnGrp, UIControl))

        'chkPrimWpnGrp initial props
        chkPrimWpnGrp = New UICheckBox(oUILib)
        With chkPrimWpnGrp
            .ControlName = "chkPrimWpnGrp"
            .Left = 200
            .Top = 206
            .Width = 63
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Primary"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
            .ToolTipText = "If checked, weapons in the primary group will fire normally. If unchecked, weapons in the primary group will hold fire."
        End With
        Me.AddChild(CType(chkPrimWpnGrp, UIControl))

        'chkSecWpnGrp initial props
        chkSecWpnGrp = New UICheckBox(oUILib)
        With chkSecWpnGrp
            .ControlName = "chkSecWpnGrp"
            .Left = 290
            .Top = 206
            .Width = 83
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Secondary"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
            .ToolTipText = "If checked, weapons in the secondary group will fire normally. If unchecked, weapons in the secondary group will hold fire."
        End With
        Me.AddChild(CType(chkSecWpnGrp, UIControl))

        'chkPDWpnGrp initial props
        chkPDWpnGrp = New UICheckBox(oUILib)
        With chkPDWpnGrp
            .ControlName = "chkPDWpnGrp"
            .Left = 400
            .Top = 206
            .Width = 101
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Point Defense"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
            .ToolTipText = "If checked, weapons in the point defense group will fire normally. If unchecked, weapons in the point defense group will hold fire."
        End With
        Me.AddChild(CType(chkPDWpnGrp, UIControl))


		'All Checkboxes use the same event
		AddHandler chkTPFighter.Click, AddressOf CheckBox_Click
		AddHandler chkTPCargo.Click, AddressOf CheckBox_Click
		AddHandler chkTPTroop.Click, AddressOf CheckBox_Click
		AddHandler chkTPEscort.Click, AddressOf CheckBox_Click
		AddHandler chkTPCarrier.Click, AddressOf CheckBox_Click
		AddHandler chkTPCapital.Click, AddressOf CheckBox_Click
		AddHandler chkTPArmed.Click, AddressOf ArmedCheck_Click
		AddHandler chkTPUnarmed.Click, AddressOf UnarmedCheck_Click
		AddHandler chkTPFacility.Click, AddressOf CheckBox_Click
		AddHandler chkCTManeuver.Click, AddressOf CT_CheckBox_Click
        AddHandler chkCTLaunchAll.Click, AddressOf CT_CheckBox_Click
        AddHandler chkCTDockDuringBattle.Click, AddressOf CT_CheckBox_Click
        AddHandler chkPrimWpnGrp.Click, AddressOf CT_CheckBox_Click
        AddHandler chkSecWpnGrp.Click, AddressOf CT_CheckBox_Click
        AddHandler chkPDWpnGrp.Click, AddressOf CT_CheckBox_Click
        AddHandler chkStopToEngage.Click, AddressOf CT_CheckBox_Click
        AddHandler chkAssistOrbit.Click, AddressOf CT_CheckBox_Click

		'Ensure I am unique
		oUILib.RemoveWindow(Me.ControlName)
		'Now, add me
		If HasAliasedRights(AliasingRights.eChangeBehavior) = True Then oUILib.AddWindow(Me)

		'Default to minimized, so set this to false and call the btn's click
		mbMinimized = False
		btnMinMax_Click(btnMinMax.ControlName)

		mbIgnoreOptSubsystmeClick = False
	End Sub

	Public Sub SetFromEntity(ByVal lEntityIndex As Int32)
		If goCurrentEnvir Is Nothing Then Return
		If mbMultiSelect = False Then
			If lEntityIndex < 0 OrElse lEntityIndex > goCurrentEnvir.lEntityUB Then Return
			If goCurrentEnvir.oEntity(lEntityIndex) Is Nothing Then Return
			If goCurrentEnvir.oEntity(lEntityIndex).oUnitDef Is Nothing Then Return
		End If

		If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
			chkTPCapital.Caption = "Ground Assault"
			chkTPCapital.ToolTipText = "Land-based assets such as tanks and AT Jeeps"
			chkTPTroop.Caption = "Troops"
			chkTPTroop.ToolTipText = "Troopers, Infantry, Mechanized Infantry and Walkers"
			chkTPTroop.Enabled = False
			chkTPCapital.Width = 90
			chkTPTroop.Width = 50
		Else
			chkTPCapital.Caption = "Capital"
			chkTPCapital.ToolTipText = "Capital ships include battlecruisers and battleships"
			chkTPTroop.Enabled = True
			chkTPTroop.Caption = "Heavy Escorts"
			chkTPTroop.ToolTipText = "Hulls at least 30,000 in size but less than 110,000." & vbCrLf & "This mainly covers Destroyers and Cruisers."
			chkTPCapital.Width = 50
			chkTPTroop.Width = 87
		End If

		optEngines.Enabled = False : optEngines.Value = False
		optHangar.Enabled = False : optHangar.Value = False
		optRadar.Enabled = False : optRadar.Value = False
		optShield.Enabled = False : optShield.Value = False
		optWeapon.Enabled = False : optWeapon.Value = False
		optNone.Value = True

		If goCurrentEnvir.oEntity(lEntityIndex).ObjTypeID = ObjectType.eUnit Then
			optEPDock.Enabled = True
			optEPEvade.Enabled = True

			Dim oModelDef As ModelDef = goModelDefs.GetModelDef(goCurrentEnvir.oEntity(lEntityIndex).oUnitDef.ModelID)
			If oModelDef Is Nothing = False Then
				If oModelDef.TypeID = 3 Then
					optEngines.Enabled = True
					optHangar.Enabled = True
					optRadar.Enabled = True
					optShield.Enabled = True
					optWeapon.Enabled = True
					optNone.Value = True
				End If
			End If
			oModelDef = Nothing
		Else
			optEPDock.Enabled = False
			'optEPEvade.Enabled = False
		End If

		If mbMultiSelect = True Then Return

		'If goCurrentEnvir.oEntity(mlEntityIndex).oMesh.bLandBased = True Then
		'    chkCTManeuver.Value = False
		'    chkCTManeuver.Enabled = False
		'Else : chkCTManeuver.Enabled = True
		'End If

		mlEntityIndex = lEntityIndex

		'Send a request for the entity's details
		Dim yMsg(13) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eGetEntityDetails).CopyTo(yMsg, 0)
		goCurrentEnvir.oEntity(lEntityIndex).GetGUIDAsString.CopyTo(yMsg, 2)
		goCurrentEnvir.GetGUIDAsString.CopyTo(yMsg, 8)
		MyBase.moUILib.SendMsgToRegion(yMsg)

		Dim ofrmRoute As frmRouteConfig = CType(MyBase.moUILib.GetWindow("frmRouteConfig"), frmRouteConfig)
		If ofrmRoute Is Nothing = False Then ofrmRoute.SetFromEntity(mlEntityIndex)
		ofrmRoute = Nothing

	End Sub

	Private Sub ClearEngagementPatterns()
		mbIgnoreOptEvents = True
		optEPHoldFire.Value = False
		optEPStandGround.Value = False
		optEPPursue.Value = False
		optEPEvade.Value = False
		optEPEngage.Value = False
		optEPDock.Value = False
		mbIgnoreOptEvents = False
	End Sub

	Private Sub optEPHoldFire_Click() Handles optEPHoldFire.Click
		If mbIgnoreOptEvents = True Then Exit Sub
		ClearEngagementPatterns()
		optEPHoldFire.Value = True
		PrepareAndSendUpdateMsg()
	End Sub

	Private Sub optEPStandGround_Click() Handles optEPStandGround.Click
		If mbIgnoreOptEvents = True Then Exit Sub
		ClearEngagementPatterns()
		optEPStandGround.Value = True
		PrepareAndSendUpdateMsg()
	End Sub

	Private Sub optEPPursue_Click() Handles optEPPursue.Click
		If mbIgnoreOptEvents = True Then Exit Sub
		ClearEngagementPatterns()
		optEPPursue.Value = True
		PrepareAndSendUpdateMsg()
	End Sub

	Private Sub optEPEvade_Click() Handles optEPEvade.Click
		If mbIgnoreOptEvents = True Then Exit Sub
		ClearEngagementPatterns()
		optEPEvade.Value = True
		MyBase.moUILib.lUISelectState = UILib.eSelectState.eEvadeToLocation

		MyBase.moUILib.AddNotification("Left-click a location to evade to...", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		If goSound Is Nothing = False Then
			'TODO: Replace this with a voice-over
			goSound.StartSound("UserInterface\Alert.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
		End If
	End Sub

	Private Sub optEPEngage_Click() Handles optEPEngage.Click
		If mbIgnoreOptEvents = True Then Exit Sub
		ClearEngagementPatterns()
		optEPEngage.Value = True
		PrepareAndSendUpdateMsg()
	End Sub

	Private Sub optEPDock_Click() Handles optEPDock.Click
		If mbIgnoreOptEvents = True Then Exit Sub
		ClearEngagementPatterns()
		optEPDock.Value = True
		MyBase.moUILib.lUISelectState = UILib.eSelectState.eDockWithObjectSelection

		MyBase.moUILib.AddNotification("Right-click a target to dock with...", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		If goSound Is Nothing = False Then
			'TODO: Replace this with a voice-over
			goSound.StartSound("UserInterface\Alert.wav", False, SoundMgr.SoundUsage.eNarrative, Nothing, Nothing)
		End If
	End Sub

	'TP Checkboxes ONLY
	Private Sub CheckBox_Click()
		If mbIgnoreChkEvents = True Then Return
		mbTPCheckBoxSet = True
		PrepareAndSendUpdateMsg()
	End Sub

	Private Sub UnarmedCheck_Click()
		If mbIgnoreChkEvents = True Then Return
		mbTPCheckBoxSet = True
		If chkTPUnarmed.Value = True Then chkTPArmed.Value = False
		PrepareAndSendUpdateMsg()
	End Sub
	Private Sub ArmedCheck_Click()
		If mbIgnoreChkEvents = True Then Return
		mbTPCheckBoxSet = True
		If chkTPArmed.Value = True Then chkTPUnarmed.Value = False
		PrepareAndSendUpdateMsg()
	End Sub

	'CT CHECKBOXES ONLY!!!
	Private Sub CT_CheckBox_Click()
		If mbIgnoreChkEvents = True Then Return
		mbCTCheckBoxSet = True
		PrepareAndSendUpdateMsg()
	End Sub

	Private Sub ClearAllCheckBoxes()
		mbIgnoreChkEvents = True
		chkTPFighter.Value = False
		chkTPCargo.Value = False
		chkTPTroop.Value = False
		chkTPEscort.Value = False
		chkTPCarrier.Value = False
		chkTPCapital.Value = False
		chkTPArmed.Value = False
		chkTPUnarmed.Value = False
		chkTPFacility.Value = False
		chkCTManeuver.Value = False
		chkCTLaunchAll.Value = False
		mbIgnoreChkEvents = False
	End Sub

	Private Sub optCTMinimize_Click() Handles optCTMinimize.Click
		If mbIgnoreOptEvents = True Then Exit Sub

		mbIgnoreOptEvents = True
		optCTNormal.Value = False
		optCTMaximize.Value = False
		optCTMinimize.Value = True
		mbIgnoreOptEvents = False
		PrepareAndSendUpdateMsg()
	End Sub

	Private Sub optCTNormal_Click() Handles optCTNormal.Click
		If mbIgnoreOptEvents = True Then Exit Sub

		mbIgnoreOptEvents = True
		optCTMaximize.Value = False
		optCTMinimize.Value = False
		optCTNormal.Value = True
		mbIgnoreOptEvents = False
		PrepareAndSendUpdateMsg()
	End Sub

	Private Sub optCTMaximize_Click() Handles optCTMaximize.Click
		If mbIgnoreOptEvents = True Then Exit Sub

		mbIgnoreOptEvents = True
		optCTNormal.Value = False
		optCTMinimize.Value = False
		optCTMaximize.Value = True
		mbIgnoreOptEvents = False
		PrepareAndSendUpdateMsg()
	End Sub

	Public Sub PrepareAndSendUpdateMsg()
		If goCurrentEnvir Is Nothing Then Return

		If mbMultiSelect = False Then
			If mlEntityIndex = -1 OrElse goCurrentEnvir.lEntityIdx(mlEntityIndex) = -1 Then Return
			If goCurrentEnvir.oEntity(mlEntityIndex).OwnerID <> glPlayerID Then Return
		End If

		Dim iTargetingTactics As Int16
        Dim iCombatTactics As Int32

		'ok, build our values for targeting pattern
		If chkTPFighter.Value = True Then iTargetingTactics = iTargetingTactics Or eiTacticalAttrs.eFighterClass
		If chkTPCargo.Value = True Then iTargetingTactics = iTargetingTactics Or eiTacticalAttrs.eCargoShipClass
		If chkTPTroop.Value = True Then iTargetingTactics = iTargetingTactics Or eiTacticalAttrs.eTroopClass
		If chkTPEscort.Value = True Then iTargetingTactics = iTargetingTactics Or eiTacticalAttrs.eEscortClass
		If chkTPCarrier.Value = True Then iTargetingTactics = iTargetingTactics Or eiTacticalAttrs.eCarrierClass
		If chkTPCapital.Value = True Then iTargetingTactics = iTargetingTactics Or eiTacticalAttrs.eCapitalClass
		If chkTPArmed.Value = True Then iTargetingTactics = iTargetingTactics Or eiTacticalAttrs.eArmedUnit
		If chkTPUnarmed.Value = True Then iTargetingTactics = iTargetingTactics Or eiTacticalAttrs.eUnarmedUnit
		If chkTPFacility.Value = True Then iTargetingTactics = iTargetingTactics Or eiTacticalAttrs.eFacility

		'Ok build our values for engagement pattern
		Dim bHasCT_Engage As Boolean = False
		If optEPHoldFire.Value = True Then
			iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eEngagement_Hold_Fire
			bHasCT_Engage = True
		ElseIf optEPStandGround.Value = True Then
			iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eEngagement_Stand_Ground
			bHasCT_Engage = True
		ElseIf optEPPursue.Value = True Then
			iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eEngagement_Pursue
			bHasCT_Engage = True
		ElseIf optEPEvade.Value = True Then
			iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eEngagement_Evade
			bHasCT_Engage = True
		ElseIf optEPEngage.Value = True Then
			iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eEngagement_Engage
			bHasCT_Engage = True
		Else
			bHasCT_Engage = optEPDock.Value
			iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eEngagement_Dock_With_Target
		End If

		'Now, build our Combat tactic
		Dim bHasCT_Dmg As Boolean = False
		If optCTMinimize.Value = True Then
			iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eTactics_Minimize_Damage_To_Self
			bHasCT_Dmg = True
		ElseIf optCTMaximize.Value = True Then
			iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eTactics_Maximize_Damage
			bHasCT_Dmg = True
		Else
			iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eTactics_Normal
			bHasCT_Dmg = optCTNormal.Value
		End If
		If chkCTManeuver.Value = True Then iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eTactics_Maneuver
		If chkCTLaunchAll.Value = True Then iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eTactics_LaunchChildren
        If chkCTDockDuringBattle.Value = True Then iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eDockDuringBattle
        If chkStopToEngage.Value = True Then iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eStopToEngage
        If chkAssistOrbit.Value = True Then iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eAssistOrbit

		If optEngines.Value = True AndAlso optEngines.Enabled = True Then
			iTargetingTactics = iTargetingTactics Or eiTacticalAttrs.eFighterTargetEngines
		ElseIf optRadar.Value = True AndAlso optRadar.Enabled = True Then
			iTargetingTactics = iTargetingTactics Or eiTacticalAttrs.eFighterTargetRadar
		ElseIf optShield.Value = True AndAlso optShield.Enabled = True Then
			iTargetingTactics = iTargetingTactics Or eiTacticalAttrs.eFighterTargetShields
		ElseIf optWeapon.Value = True AndAlso optWeapon.Enabled = True Then
			iTargetingTactics = iTargetingTactics Or eiTacticalAttrs.eFighterTargetWeapons
		ElseIf optHangar.Value = True AndAlso optHangar.Enabled = True Then
			iTargetingTactics = iTargetingTactics Or eiTacticalAttrs.eFighterTargetHangar
        End If

        If chkPrimWpnGrp.Value = False Then iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eHoldPrimaryWpnGrp
        If chkSecWpnGrp.Value = False Then iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eHoldSecondaryWpnGrp
        If chkPDWpnGrp.Value = False Then iCombatTactics = iCombatTactics Or eiBehaviorPatterns.eHoldPointDefenseWpnGrp

		'  Evade and Dock With Target are mutually exclusive, so...
		'  If Evade, then the first value is LocX and the second value is LocZ
		'  IF Dock With Target, then the first value is EntityID and the second value is EntityTypeID (as an int32)

		Dim bNotifyNoMan As Boolean = False
		Dim bManSet As Boolean = chkCTManeuver.Value

		'Now, two separate things...
		Dim yMsg() As Byte = Nothing
		If mbMultiSelect = True Then
			'Get our count
			Dim lCnt As Int32 = 0
			Dim lIdxs() As Int32 = Nothing
			For X As Int32 = 0 To goCurrentEnvir.lEntityUB
				If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then

					If bManSet = True AndAlso bNotifyNoMan = False AndAlso goCurrentEnvir.oEntity(X).oMesh.bLandBased = True Then
						bNotifyNoMan = True
					End If

					lCnt += 1
					ReDim Preserve lIdxs(lCnt - 1)
					lIdxs(lCnt - 1) = X
				End If
			Next X

			'Now, go back through and make one really large message...
            '2 (for msg len) + 2 + 6 + 2 + 4 + 4 + 4 = 22 per item... 
            Dim iLen As Int16 = 22
			Dim lPos As Int32 = 0

            'ReDim yMsg((iLen * lCnt) - 1)
            ReDim yMsg(((iLen + 2) * lCnt) - 1)

			For X As Int32 = 0 To lCnt - 1
				'ok, set our value...
				System.BitConverter.GetBytes(iLen).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityAI).CopyTo(yMsg, lPos) : lPos += 2
				goCurrentEnvir.oEntity(lIdxs(X)).GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
				System.BitConverter.GetBytes(iTargetingTactics).CopyTo(yMsg, lPos) : lPos += 2
				System.BitConverter.GetBytes(iCombatTactics).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(lExtendVal1).CopyTo(yMsg, lPos) : lPos += 4
				System.BitConverter.GetBytes(lExtendVal2).CopyTo(yMsg, lPos) : lPos += 4
			Next X
			If yMsg Is Nothing = False Then MyBase.moUILib.SendLenAppendedMsgToRegion(yMsg)
		Else
			goCurrentEnvir.oEntity(mlEntityIndex).iCombatTactics = iCombatTactics
			goCurrentEnvir.oEntity(mlEntityIndex).iTargetingTactics = iTargetingTactics

			If bManSet = True AndAlso bNotifyNoMan = False AndAlso goCurrentEnvir.oEntity(mlEntityIndex).oMesh.bLandBased = True Then
				bNotifyNoMan = True
			End If

			'Now, send our message
            ReDim yMsg(21)

			System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityAI).CopyTo(yMsg, 0)
			goCurrentEnvir.oEntity(mlEntityIndex).GetGUIDAsString.CopyTo(yMsg, 2)
			System.BitConverter.GetBytes(iTargetingTactics).CopyTo(yMsg, 8)
			System.BitConverter.GetBytes(iCombatTactics).CopyTo(yMsg, 10)
            System.BitConverter.GetBytes(lExtendVal1).CopyTo(yMsg, 14)
            System.BitConverter.GetBytes(lExtendVal2).CopyTo(yMsg, 18)
			If yMsg Is Nothing = False Then MyBase.moUILib.SendMsgToRegion(yMsg)
		End If

		If bNotifyNoMan = True Then
			MyBase.moUILib.AddNotification("Some entities could not be set to Maneuver.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

			If mbMultiSelect = False Then

				mbIgnoreChkEvents = True
				chkCTManeuver.Value = False
				mbIgnoreChkEvents = False
			End If
		End If
	End Sub

	Private Sub frmBehavior_OnNewFrame() Handles MyBase.OnNewFrame
		If mbMultiSelect = True Then Return

        If goCurrentEnvir Is Nothing Then Return
        If mlEntityIndex < 0 OrElse mlEntityIndex > goCurrentEnvir.lEntityUB Then Return
        If goCurrentEnvir.oEntity(mlEntityIndex) Is Nothing Then Return
        If goCurrentEnvir.oEntity(mlEntityIndex).oUnitDef Is Nothing Then Return

		'Now, check if values match
		If goCurrentEnvir.oEntity(mlEntityIndex).iCombatTactics <> miLastCT OrElse _
		   goCurrentEnvir.oEntity(mlEntityIndex).iTargetingTactics <> miLastTP Then
			'set up our values from the new values...
			mbIgnoreOptEvents = True
			ClearEngagementPatterns()
			optCTMinimize.Value = False : optCTNormal.Value = False : optCTMaximize.Value = False
			ClearAllCheckBoxes()

			mbIgnoreOptEvents = True
			mbIgnoreChkEvents = True

			With goCurrentEnvir.oEntity(mlEntityIndex)
				If (.iCombatTactics And eiBehaviorPatterns.eEngagement_Engage) <> 0 Then
					optEPEngage.Value = True
				ElseIf (.iCombatTactics And eiBehaviorPatterns.eEngagement_Evade) <> 0 Then
					optEPEvade.Value = True
				ElseIf (.iCombatTactics And eiBehaviorPatterns.eEngagement_Hold_Fire) <> 0 Then
					optEPHoldFire.Value = True
				ElseIf (.iCombatTactics And eiBehaviorPatterns.eEngagement_Pursue) <> 0 Then
					optEPPursue.Value = True
				ElseIf (.iCombatTactics And eiBehaviorPatterns.eEngagement_Stand_Ground) <> 0 Then
					optEPStandGround.Value = True
				Else
					optEPDock.Value = True
				End If

				If (.iCombatTactics And eiBehaviorPatterns.eTactics_Minimize_Damage_To_Self) <> 0 Then
					optCTMinimize.Value = True
				ElseIf (.iCombatTactics And eiBehaviorPatterns.eTactics_Maximize_Damage) <> 0 Then
					optCTMaximize.Value = True
				Else : optCTNormal.Value = True
				End If

				If (.iCombatTactics And eiBehaviorPatterns.eTactics_LaunchChildren) <> 0 Then chkCTLaunchAll.Value = True
                If (.iCombatTactics And eiBehaviorPatterns.eTactics_Maneuver) <> 0 Then chkCTManeuver.Value = True
                If (.iCombatTactics And eiBehaviorPatterns.eDockDuringBattle) <> 0 Then chkCTDockDuringBattle.Value = True
                If (.iCombatTactics And eiBehaviorPatterns.eStopToEngage) <> 0 Then chkStopToEngage.Value = True
                If (.iCombatTactics And eiBehaviorPatterns.eAssistOrbit) <> 0 Then chkAssistOrbit.Value = True

                If (.iCombatTactics And eiBehaviorPatterns.eHoldPointDefenseWpnGrp) <> 0 Then chkPDWpnGrp.Value = False Else chkPDWpnGrp.Value = True
                If (.iCombatTactics And eiBehaviorPatterns.eHoldPrimaryWpnGrp) <> 0 Then chkPrimWpnGrp.Value = False Else chkPrimWpnGrp.Value = True
                If (.iCombatTactics And eiBehaviorPatterns.eHoldSecondaryWpnGrp) <> 0 Then chkSecWpnGrp.Value = False Else chkSecWpnGrp.Value = True

				'now for targeting patterns...
				If (.iTargetingTactics And eiTacticalAttrs.eFighterClass) <> 0 Then chkTPFighter.Value = True
				If (.iTargetingTactics And eiTacticalAttrs.eCargoShipClass) <> 0 Then chkTPCargo.Value = True
				If (.iTargetingTactics And eiTacticalAttrs.eTroopClass) <> 0 Then chkTPTroop.Value = True
				If (.iTargetingTactics And eiTacticalAttrs.eEscortClass) <> 0 Then chkTPEscort.Value = True
				If (.iTargetingTactics And eiTacticalAttrs.eCarrierClass) <> 0 Then chkTPCarrier.Value = True
				If (.iTargetingTactics And eiTacticalAttrs.eCapitalClass) <> 0 Then chkTPCapital.Value = True
				If (.iTargetingTactics And eiTacticalAttrs.eArmedUnit) <> 0 Then chkTPArmed.Value = True
				If (.iTargetingTactics And eiTacticalAttrs.eUnarmedUnit) <> 0 Then chkTPUnarmed.Value = True
				If (.iTargetingTactics And eiTacticalAttrs.eFacility) <> 0 Then chkTPFacility.Value = True

				optEngines.Value = (.iTargetingTactics And eiTacticalAttrs.eFighterTargetEngines) <> 0
				optHangar.Value = (.iTargetingTactics And eiTacticalAttrs.eFighterTargetHangar) <> 0
				optShield.Value = (.iTargetingTactics And eiTacticalAttrs.eFighterTargetShields) <> 0
				optRadar.Value = (.iTargetingTactics And eiTacticalAttrs.eFighterTargetRadar) <> 0
				optWeapon.Value = (.iTargetingTactics And eiTacticalAttrs.eFighterTargetWeapons) <> 0
				optNone.Value = (.iTargetingTactics And (eiTacticalAttrs.eFighterTargetEngines Or eiTacticalAttrs.eFighterTargetHangar Or eiTacticalAttrs.eFighterTargetRadar Or eiTacticalAttrs.eFighterTargetShields Or eiTacticalAttrs.eFighterTargetWeapons)) = 0

				miLastCT = .iCombatTactics
				miLastTP = .iTargetingTactics
			End With



			mbIgnoreOptEvents = False
			mbIgnoreChkEvents = False
		End If
	End Sub

	Public Sub btnMinMax_Click(ByVal sName As String) Handles btnMinMax.Click
		'Dim rcTemp As System.Drawing.Rectangle
		Dim lNewY As Int32
		Dim lNewWidth As Int32
		Dim lNewHeight As Int32
		Dim X As Int32
		Dim lBtnX As Int32

		mbMinimized = Not mbMinimized

		If mbMinimized = True Then
			'Ok, we are now minimized, let's set up our values
			If mbMultiSelect = True Then
				lNewY = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - btnMinMax.Height - 4
			Else
				lNewY = mlStartY + Me.Height - btnMinMax.Height - 4
			End If
			lBtnX = 2
			lNewWidth = btnMinMax.Width + 4
			lNewHeight = btnMinMax.Height + 4
			Me.bRoundedBorder = False
		Else
			'ok, we are now maximized, let's set up our values
			lNewY = mlStartY
			lNewWidth = 512 '312
			lBtnX = lNewWidth - btnMinMax.Width - 1	'240
			lNewHeight = 232
			Me.bRoundedBorder = True
		End If
		'Now, set the btn's loc and window props
		Me.Top = lNewY
		Me.Width = lNewWidth
		Me.Height = lNewHeight
		btnMinMax.Left = lBtnX

		If Me.Top + Me.Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then
			Dim lTempVal As Int32 = Me.Top + Me.Height - MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight
			Me.Top -= lTempVal
		End If

		For X = 0 To MyBase.ChildrenUB
			MyBase.moChildren(X).Visible = Not mbMinimized
		Next X
		btnMinMax.Visible = True

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

		'rcTemp = btnMinMax.ControlImageRect_Disabled
		'btnMinMax.ControlImageRect_Disabled = btnMinMax.ControlImageRect_Normal
		'btnMinMax.ControlImageRect_Normal = rcTemp
		'btnMinMax.ControlImageRect_Pressed = rcTemp

		''And see if we need to relocate the facility advanced window
		'Dim oTmpWin As frmFacilityAdv = CType(MyBase.moUILib.GetWindow("frmFacilityAdv"), frmFacilityAdv)
		'If oTmpWin Is Nothing = False Then
		'    If mbMinimized = True Then
		'        oTmpWin.Left = Me.Left + Me.Width
		'        oTmpWin.Top = (Me.Top + Me.Height) - oTmpWin.Height
		'    Else
		'        oTmpWin.Left = Me.Left
		'        oTmpWin.Top = Me.Top - oTmpWin.Height
		'    End If
		'End If
		'oTmpWin = Nothing
	End Sub

    Private Sub frmBehavior_WindowMoved() Handles Me.WindowMoved
        muSettings.BehaviorLocX = Me.Left
        muSettings.BehaviorLocY = Me.Top
    End Sub

	Private Sub btnRoutes_Click(ByVal sName As String) Handles btnRoutes.Click
        OpenRouteConfig()
    End Sub

    Public Sub OpenRouteConfig()
        Dim ofrm As frmRouteConfig = CType(goUILib.GetWindow("frmRouteConfig"), frmRouteConfig)
        If ofrm Is Nothing = False AndAlso ofrm.Visible = True Then
            goUILib.RemoveWindow("frmRouteConfig")
        Else
            ofrm = New frmRouteConfig(MyBase.moUILib)
            ofrm.SetFromEntity(Me.mlEntityIndex)
            ofrm.Visible = True
            ofrm.Left = Me.Left
            ofrm.Top = Me.Top - ofrm.Height - 1
        End If
        ofrm = Nothing
    End Sub

	Private Sub optHangar_Click() Handles optHangar.Click
		If mbIgnoreOptSubsystmeClick = True Then Return
		mbIgnoreOptSubsystmeClick = True
		optEngines.Value = False
		optNone.Value = False
		optRadar.Value = False
		optShield.Value = False
		optWeapon.Value = False
		PrepareAndSendUpdateMsg()
		mbIgnoreOptSubsystmeClick = False
	End Sub

	Private Sub optEngines_Click() Handles optEngines.Click
		If mbIgnoreOptSubsystmeClick = True Then Return
		mbIgnoreOptSubsystmeClick = True
		optNone.Value = False
		optHangar.Value = False
		optRadar.Value = False
		optShield.Value = False
		optWeapon.Value = False
		PrepareAndSendUpdateMsg()
		mbIgnoreOptSubsystmeClick = False
	End Sub

	Private Sub optRadar_Click() Handles optRadar.Click
		If mbIgnoreOptSubsystmeClick = True Then Return
		mbIgnoreOptSubsystmeClick = True
		optEngines.Value = False
		optHangar.Value = False
		optNone.Value = False
		optShield.Value = False
		optWeapon.Value = False
		PrepareAndSendUpdateMsg()
		mbIgnoreOptSubsystmeClick = False
	End Sub

	Private Sub optShield_Click() Handles optShield.Click
		If mbIgnoreOptSubsystmeClick = True Then Return
		mbIgnoreOptSubsystmeClick = True
		optEngines.Value = False
		optHangar.Value = False
		optRadar.Value = False
		optNone.Value = False
		optWeapon.Value = False
		PrepareAndSendUpdateMsg()
		mbIgnoreOptSubsystmeClick = False
	End Sub

	Private Sub optWeapon_Click() Handles optWeapon.Click
		If mbIgnoreOptSubsystmeClick = True Then Return
		mbIgnoreOptSubsystmeClick = True
		optEngines.Value = False
		optHangar.Value = False
		optRadar.Value = False
		optShield.Value = False
		optNone.Value = False
		PrepareAndSendUpdateMsg()
		mbIgnoreOptSubsystmeClick = False
	End Sub

	Private Sub optNone_Click() Handles optNone.Click
		If mbIgnoreOptSubsystmeClick = True Then Return
		mbIgnoreOptSubsystmeClick = True
		optEngines.Value = False
		optHangar.Value = False
		optRadar.Value = False
		optShield.Value = False
		optWeapon.Value = False
		PrepareAndSendUpdateMsg()
		mbIgnoreOptSubsystmeClick = False
	End Sub
End Class
