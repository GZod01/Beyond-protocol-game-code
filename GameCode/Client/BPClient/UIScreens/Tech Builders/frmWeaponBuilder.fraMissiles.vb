Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmWeaponBuilder
    'Interface created from Interface Builder
    Private Class fraMissiles
        Inherits fraWeaponBase


        Private tpMaxDmg As ctlTechProp
        Private tpHullSize As ctlTechProp
        Private tpAgility As ctlTechProp
        Private tpROF As ctlTechProp
        Private tpRange As ctlTechProp
        Private tpHoming As ctlTechProp
        Private tpExplosionRadius As ctlTechProp
        Private tpStructHP As ctlTechProp

        'Private lblMaxDmg As UILabel
        'Private lblHullSize As UILabel
        'Private lblMaxSpeed As UILabel
        'Private lblManeuver As UILabel
        'Private lblROF As UILabel
        'Private lblFlightTime As UILabel
        'Private lblHoming As UILabel
        'Private lblExplosionRadius As UILabel
        'Private lblStructHP As UILabel

        'Private txtMaxDamage As UITextBox
        'Private txtHullSize As UITextBox
        'Private txtMaxSpeed As UITextBox
        'Private txtManeuver As UITextBox
        'Private txtFlightTime As UITextBox
        'Private txtHoming As UITextBox
        'Private txtExplosionRadius As UITextBox
        'Private txtStructHP As UITextBox
        'Private txtROF As UITextBox

        Private tpBodyMat As ctlTechProp
        Private tpNoseMat As ctlTechProp
        Private tpFlapsMat As ctlTechProp
        Private tpFuelMat As ctlTechProp
        Private tpPayloadMat As ctlTechProp

        Private lblPayloadType As UILabel
        Private lblBodyMat As UILabel
        Private lblNoseMat As UILabel
        Private lblFlapsMat As UILabel
        Private lblFuelMat As UILabel
        Private lblPayloadMat As UILabel

        'Private WithEvents cboPayloadMat As UIComboBox
        'Private WithEvents cboFuelMat As UIComboBox
        'Private WithEvents cboFlapsMat As UIComboBox
        'Private WithEvents cboNoseMat As UIComboBox
        'Private WithEvents cboBodyMat As UIComboBox
        Private stmPayloadMat As ctlSetTechMineral
        Private stmFuelMat As ctlSetTechMineral
        Private stmFlapsMat As ctlSetTechMineral
        Private stmNoseMat As ctlSetTechMineral
        Private stmBodyMat As ctlSetTechMineral

        Private WithEvents cboPayloadType As UIComboBox

        Private mbLoading As Boolean = True
        Private mlEntityIndex As Int32 = -1

        Private moTech As WeaponTech = Nothing

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraMissiles initial props
            With Me
                .lWindowMetricID = BPMetrics.eWindow.eMissileBuilder
                .ControlName = "fraMissiles"
                .Left = 18
                .Top = 52
                .Width = 470
                .Height = 335
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .Moveable = False
                .BorderLineWidth = 1
                .mbAcceptReprocessEvents = True
            End With

            ''lblMaxDmg initial props
            'lblMaxDmg = New UILabel(oUILib)
            'With lblMaxDmg
            '    .ControlName = "lblMaxDmg"
            '    .Left = 10
            '    .Top = 10
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Maximum Damage:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "The maximum damage expected to inflict with a single missile."
            'End With
            'Me.AddChild(CType(lblMaxDmg, UIControl))

            ''lblHullSize initial props
            'lblHullSize = New UILabel(oUILib)
            'With lblHullSize
            '    .ControlName = "lblHullSize"
            '    .Left = 10
            '    .Top = 35
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Missile Hull Size:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "The hull size of a single missile (directly impacts hull size of the overall system)."
            'End With
            'Me.AddChild(CType(lblHullSize, UIControl))

            ''lblMaxSpeed initial props
            'lblMaxSpeed = New UILabel(oUILib)
            'With lblMaxSpeed
            '    .ControlName = "lblMaxSpeed"
            '    .Left = 10
            '    .Top = 60
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Maximum Speed:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "Maximum speed the missile will move." & vbCrLf & "Faster targets may be able to outrun slower missiles."
            'End With
            'Me.AddChild(CType(lblMaxSpeed, UIControl))

            ''lblManeuver initial props
            'lblManeuver = New UILabel(oUILib)
            'With lblManeuver
            '    .ControlName = "lblManeuver"
            '    .Left = 250
            '    .Top = 60
            '    .Width = 65
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Maneuver:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(6, DrawTextFormat)
            '    .ToolTipText = "Maneuver rating of the missile."
            'End With
            'Me.AddChild(CType(lblManeuver, UIControl))

            ''lblROF initial props
            'lblROF = New UILabel(oUILib)
            'With lblROF
            '    .ControlName = "lblROF"
            '    .Left = 10
            '    .Top = 110
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Rate of Fire:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "Time in seconds for the launcher to fire a new missile." '& vbCrLf & "0 or less indicates single shot launchers."
            'End With
            'Me.AddChild(CType(lblROF, UIControl))

            ''lblFlightTime initial props
            'lblFlightTime = New UILabel(oUILib)
            'With lblFlightTime
            '    .ControlName = "lblFlightTime"
            '    .Left = 245
            '    .Top = 85
            '    .Width = 69
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Flight Time:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(6, DrawTextFormat)
            '    .ToolTipText = "Number of seconds the missile can fly before exhausting its fuel and exploding."
            'End With
            'Me.AddChild(CType(lblFlightTime, UIControl))

            ''lblHoming initial props
            'lblHoming = New UILabel(oUILib)
            'With lblHoming
            '    .ControlName = "lblHoming"
            '    .Left = 10
            '    .Top = 85
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Homing Accuracy:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "Number between 0 and 255 indicating the intelligence" & vbCrLf & "level of the missile's on-board guidance systems."
            'End With
            'Me.AddChild(CType(lblHoming, UIControl))

            ''lblExplosionRadius initial props
            'lblExplosionRadius = New UILabel(oUILib)
            'With lblExplosionRadius
            '    .ControlName = "lblExplosionRadius"
            '    .Left = 10
            '    .Top = 135
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Explosion Radius:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "Radius from the point of detonation that the missile inflicts damage." & vbCrLf & "Area effect damage is based on the payload only."
            'End With
            'Me.AddChild(CType(lblExplosionRadius, UIControl))

            ''lblStructHP initial props
            'lblStructHP = New UILabel(oUILib)
            'With lblStructHP
            '    .ControlName = "lblStructHP"
            '    .Left = 193
            '    .Top = 135
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Structural Hitpoints:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(6, DrawTextFormat)
            '    .ToolTipText = "Amount of damage the missile can sustain before being destroyed."
            'End With
            'Me.AddChild(CType(lblStructHP, UIControl))

            ''txtMaxDamage initial props
            'txtMaxDamage = New UITextBox(oUILib)
            'With txtMaxDamage
            '    .ControlName = "txtMaxDamage"
            '    .Left = 135
            '    .Top = 10
            '    .Width = 100
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = ""
            '    .ForeColor = muSettings.InterfaceTextBoxForeColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            '    .MaxLength = 9
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .ToolTipText = "The maximum damage expected to inflict with a single missile."
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtMaxDamage, UIControl))
            tpMaxDmg = New ctlTechProp(oUILib)
            With tpMaxDmg
                .ControlName = "tpMaxDmg"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 10000 '000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Maximum Damage:"
                .PropertyValue = 0
                .Top = 5 '10
                .Visible = True
                .Width = 430
                .bNoMaxValue = True
                .ToolTipText = "The maximum damage expected to inflict with a single missile."
            End With
            Me.AddChild(CType(tpMaxDmg, UIControl))

            ''txtHullSize initial props
            'txtHullSize = New UITextBox(oUILib)
            'With txtHullSize
            '    .ControlName = "txtHullSize"
            '    .Left = 135
            '    .Top = 35
            '    .Width = 100
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = ""
            '    .ForeColor = muSettings.InterfaceTextBoxForeColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            '    .MaxLength = 5
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .ToolTipText = "The hull size of a single missile (directly impacts hull size of the overall system)."
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtHullSize, UIControl))
            tpHullSize = New ctlTechProp(oUILib)
            With tpHullSize
                .ControlName = "tpHullSize"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 100 '000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Missile Size:"
                .PropertyValue = 0
                .Top = 25 '35
                .Visible = True
                .Width = 430
                .bNoMaxValue = True
                .ToolTipText = "The hull size of a single missile (directly impacts hull size of the overall system)."
            End With
            Me.AddChild(CType(tpHullSize, UIControl))

            ''txtMaxSpeed initial props
            'txtMaxSpeed = New UITextBox(oUILib)
            'With txtMaxSpeed
            '    .ControlName = "txtMaxSpeed"
            '    .Left = 135
            '    .Top = 60
            '    .Width = 50
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = ""
            '    .ForeColor = muSettings.InterfaceTextBoxForeColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            '    .MaxLength = 3
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .ToolTipText = "Maximum speed the missile will move." & vbCrLf & "Faster targets may be able to outrun slower missiles."
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtMaxSpeed, UIControl))
            tpAgility = New ctlTechProp(oUILib)
            With tpAgility
                .ControlName = "tpAgility"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 255
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Agility:"
                .PropertyValue = 0
                .Top = 45 '60
                .Visible = True
                .Width = 430
                .ToolTipText = "Defines the maximum speed and maneuver of the missile."
            End With
            Me.AddChild(CType(tpAgility, UIControl))

            ''txtManeuver initial props
            'txtManeuver = New UITextBox(oUILib)
            'With txtManeuver
            '    .ControlName = "txtManeuver"
            '    .Left = 320
            '    .Top = 60
            '    .Width = 50
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = ""
            '    .ForeColor = muSettings.InterfaceTextBoxForeColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            '    .MaxLength = 3
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .ToolTipText = "Maneuver rating of the missile."
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtManeuver, UIControl))


            ''txtROF initial props
            'txtROF = New UITextBox(oUILib)
            'With txtROF
            '    .ControlName = "txtROF"
            '    .Left = 135
            '    .Top = 110
            '    .Width = 100
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = ""
            '    .ForeColor = muSettings.InterfaceTextBoxForeColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            '    .MaxLength = 7
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .ToolTipText = "Time in seconds for the launcher to fire a new missile." '& vbCrLf & "Values less than 0 indicates single shot launchers."
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtROF, UIControl))
            tpROF = New ctlTechProp(oUILib)
            With tpROF
                .ControlName = "tpROF"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 1500
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Rate of Fire:"
                .PropertyValue = 300
                .Top = 85 '110
                .Visible = True
                .Width = 430
                .bNoMaxValue = True
                .SetDivisor(30)
                .ToolTipText = "Number of seconds between shots fired." & vbCrLf & "Can be any positive decimal number greater than 1/2 a second."
            End With
            Me.AddChild(CType(tpROF, UIControl))

            ''txtFlightTime initial props
            'txtFlightTime = New UITextBox(oUILib)
            'With txtFlightTime
            '    .ControlName = "txtFlightTime"
            '    .Left = 320
            '    .Top = 85
            '    .Width = 50
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = ""
            '    .ForeColor = muSettings.InterfaceTextBoxForeColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            '    .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            '    .MaxLength = 7
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .ToolTipText = "Number of seconds the missile can fly before exhausting its fuel and exploding."
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtFlightTime, UIControl))
            tpRange = New ctlTechProp(oUILib)
            With tpRange
                .ControlName = "tpRange"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 1500
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Range:"
                .PropertyValue = 0
                .Top = 65 '110
                .Visible = True
                .Width = 430
                .SetDivisor(1)
                .ToolTipText = "The maximum range the missile can be fired from."
            End With
            Me.AddChild(CType(tpRange, UIControl))

            ''txtHoming initial props
            'txtHoming = New UITextBox(oUILib)
            'With txtHoming
            '    .ControlName = "txtHoming"
            '    .Left = 135
            '    .Top = 85
            '    .Width = 50
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
            '    .MaxLength = 3
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .ToolTipText = "Number between 0 and 255 indicating the intelligence" & vbCrLf & "level of the missile's on-board guidance systems."
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtHoming, UIControl))
            tpHoming = New ctlTechProp(oUILib)
            With tpHoming
                .ControlName = "tpHoming"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 255
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Guidance:"
                .PropertyValue = 0
                .Top = 105
                .Visible = True
                .Width = 430
                .ToolTipText = "Number between 0 and 255 indicating the intelligence" & vbCrLf & "level of the missile's on-board guidance systems."
            End With
            Me.AddChild(CType(tpHoming, UIControl))

            ''txtExplosionRadius initial props
            'txtExplosionRadius = New UITextBox(oUILib)
            'With txtExplosionRadius
            '    .ControlName = "txtExplosionRadius"
            '    .Left = 135
            '    .Top = 135
            '    .Width = 50
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
            '    .MaxLength = 3
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .ToolTipText = "Radius from the point of detonation that the missile inflicts damage." & vbCrLf & "Area effect damage is based on the payload only."
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtExplosionRadius, UIControl))
            tpExplosionRadius = New ctlTechProp(oUILib)
            With tpExplosionRadius
                .ControlName = "tpExplosionRadius"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 255
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Explosion Radius:"
                .PropertyValue = 0
                .Top = 125
                .Visible = True
                .Width = 430
                .ToolTipText = "Radius from the point of detonation that the missile inflicts damage." & vbCrLf & "Area effect damage is based on the payload only."
            End With
            Me.AddChild(CType(tpExplosionRadius, UIControl))

            ''txtStructHP initial props
            'txtStructHP = New UITextBox(oUILib)
            'With txtStructHP
            '    .ControlName = "txtStructHP"
            '    .Left = 320
            '    .Top = 135
            '    .Width = 50
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
            '    .MaxLength = 5
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .ToolTipText = "Amount of damage the missile can sustain before being destroyed."
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtStructHP, UIControl))
            tpStructHP = New ctlTechProp(oUILib)
            With tpStructHP
                .ControlName = "tpStructHP"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 1000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Structural Hitpoints:"
                .PropertyValue = 0
                .Top = 145
                .Visible = True
                .Width = 430
                .bNoMaxValue = False
                .MaxValue = 30
                .blAbsoluteMaximum = 30
                .ToolTipText = "Amount of damage the missile can sustain before being destroyed."
            End With
            Me.AddChild(CType(tpStructHP, UIControl))


            'lblPayloadType initial props
            lblPayloadType = New UILabel(oUILib)
            With lblPayloadType
                .ControlName = "lblPayloadType"
                .Left = 10
                .Top = 165
                .Width = 120
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Payload Type:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Indication of the type of damage the payload" & vbCrLf & "of the missile inflicts. Missiles always" & vbCrLf & "inflict Impact damage in addition to the payload."
            End With
            Me.AddChild(CType(lblPayloadType, UIControl))


            'lblBodyMat initial props
            lblBodyMat = New UILabel(oUILib)
            With lblBodyMat
                .ControlName = "lblBodyMat"
                .Left = 10
                .Top = 190
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Body:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblBodyMat, UIControl))

            'lblNoseMat initial props
            lblNoseMat = New UILabel(oUILib)
            With lblNoseMat
                .ControlName = "lblNoseMat"
                .Left = 10
                .Top = 215
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Nose:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblNoseMat, UIControl))

            'lblFlapsMat initial props
            lblFlapsMat = New UILabel(oUILib)
            With lblFlapsMat
                .ControlName = "lblFlapsMat"
                .Left = 10
                .Top = 240
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Flaps:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblFlapsMat, UIControl))

            'lblFuelMat initial props
            lblFuelMat = New UILabel(oUILib)
            With lblFuelMat
                .ControlName = "lblFuelMat"
                .Left = 10
                .Top = 265
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Fuel:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblFuelMat, UIControl))

            'lblPayloadMat initial props
            lblPayloadMat = New UILabel(oUILib)
            With lblPayloadMat
                .ControlName = "lblPayloadMat"
                .Left = 10
                .Top = 290
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Payload:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblPayloadMat, UIControl))

            ''chkSingleShot initial props
            'chkSingleShot = New UICheckBox(oUILib)
            'With chkSingleShot
            '    .ControlName = "chkSingleShot"
            '    .Left = 250
            '    .Top = 110
            '    .Width = 89
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = False    'true
            '    .Caption = "Single Shot"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(6, DrawTextFormat)
            '    .Value = False
            '    .ToolTipText = "If checked, the Rate Of Fire is set to -1." & vbCrLf & "A Single Shot missile system is smaller in size than" & vbCrLf & "a multi-shot system making it useful for smaller craft." & vbCrLf & "Single-shot missile systems must be reloaded at a facility."
            '    .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            'End With
            'Me.AddChild(CType(chkSingleShot, UIControl))

            tpPayloadMat = New ctlTechProp(oUILib)
            With tpPayloadMat
                .ControlName = "tpPayloadMat"
                .Enabled = True
                .Height = 18
                .Left = lblPayloadMat.Left + lblPayloadMat.Width + 160        '10 for left/right and 150 for cbo
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Payload:"
                .PropertyValue = 0
                .Top = 290
                .bNoMaxValue = True
                .Visible = True
                .Width = Me.Width - .Left - 10
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpPayloadMat, UIControl))

            tpFuelMat = New ctlTechProp(oUILib)
            With tpFuelMat
                .ControlName = "tpFuelMat"
                .Enabled = True
                .Height = 18
                .Left = tpPayloadMat.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Fuel:"
                .PropertyValue = 0
                .Top = 265
                .bNoMaxValue = True
                .Visible = True
                .Width = Me.Width - .Left - 10
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpFuelMat, UIControl))

            tpFlapsMat = New ctlTechProp(oUILib)
            With tpFlapsMat
                .ControlName = "tpFlapsMat"
                .Enabled = True
                .Height = 18
                .Left = tpPayloadMat.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Flaps:"
                .PropertyValue = 0
                .Top = 240
                .bNoMaxValue = True
                .Visible = True
                .Width = Me.Width - .Left - 10
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpFlapsMat, UIControl))

            tpNoseMat = New ctlTechProp(oUILib)
            With tpNoseMat
                .ControlName = "tpNoseMat"
                .Enabled = True
                .Height = 18
                .Left = tpPayloadMat.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Nose:"
                .PropertyValue = 0
                .Top = 215
                .bNoMaxValue = True
                .Visible = True
                .Width = Me.Width - .Left - 10
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpNoseMat, UIControl))

            tpBodyMat = New ctlTechProp(oUILib)
            With tpBodyMat
                .ControlName = "tpBodyMat"
                .Enabled = True
                .Height = 18
                .Left = tpPayloadMat.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Body:"
                .PropertyValue = 0
                .Top = 190
                .bNoMaxValue = True
                .Visible = True
                .Width = Me.Width - .Left - 10
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpBodyMat, UIControl))

            ''cboPayloadMat initial props
            'cboPayloadMat = New UIComboBox(oUILib)
            'With cboPayloadMat
            '    .ControlName = "cboPayloadMat"
            '    .Left = lblPayloadMat.Left + lblPayloadMat.Width + 5
            '    .Top = 290
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
            '    .mbAcceptReprocessEvents = True
            'End With
            'Me.AddChild(CType(cboPayloadMat, UIControl))
            stmPayloadMat = New ctlSetTechMineral(oUILib)
            With stmPayloadMat
                .ControlName = "stmPayloadMat"
                .Left = lblPayloadMat.Left + lblPayloadMat.Width + 5
                .Top = 290
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 4
            End With
            Me.AddChild(CType(stmPayloadMat, UIControl))

            ''cboFuelMat initial props
            'cboFuelMat = New UIComboBox(oUILib)
            'With cboFuelMat
            '    .ControlName = "cboFuelMat"
            '    .Left = cboPayloadMat.Left
            '    .Top = 265
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
            '    .mbAcceptReprocessEvents = True
            'End With
            'Me.AddChild(CType(cboFuelMat, UIControl))
            stmFuelMat = New ctlSetTechMineral(oUILib)
            With stmFuelMat
                .ControlName = "stmFuelMat"
                .Left = stmPayloadMat.Left
                .Top = 265
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 3
            End With
            Me.AddChild(CType(stmFuelMat, UIControl))

            ''cboFlapsMat initial props
            'cboFlapsMat = New UIComboBox(oUILib)
            'With cboFlapsMat
            '    .ControlName = "cboFlapsMat"
            '    .Left = cboPayloadMat.Left
            '    .Top = 240
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
            '    .mbAcceptReprocessEvents = True
            'End With
            'Me.AddChild(CType(cboFlapsMat, UIControl))
            stmFlapsMat = New ctlSetTechMineral(oUILib)
            With stmFlapsMat
                .ControlName = "stmFlapsMat"
                .Left = stmPayloadMat.Left
                .Top = 240
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 2
            End With
            Me.AddChild(CType(stmFlapsMat, UIControl))

            ''cboNoseMat initial props
            'cboNoseMat = New UIComboBox(oUILib)
            'With cboNoseMat
            '    .ControlName = "cboNoseMat"
            '    .Left = cboPayloadMat.Left
            '    .Top = 215
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
            '    .mbAcceptReprocessEvents = True
            'End With
            'Me.AddChild(CType(cboNoseMat, UIControl))
            stmNoseMat = New ctlSetTechMineral(oUILib)
            With stmNoseMat
                .ControlName = "stmNoseMat"
                .Left = stmPayloadMat.Left
                .Top = 215
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 1
            End With
            Me.AddChild(CType(stmNoseMat, UIControl))

            ''cboBodyMat initial props
            'cboBodyMat = New UIComboBox(oUILib)
            'With cboBodyMat
            '    .ControlName = "cboBodyMat"
            '    .Left = cboPayloadMat.Left
            '    .Top = 190
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
            '    .mbAcceptReprocessEvents = True
            'End With
            'Me.AddChild(CType(cboBodyMat, UIControl))
            stmBodyMat = New ctlSetTechMineral(oUILib)
            With stmBodyMat
                .ControlName = "stmBodyMat"
                .Left = stmPayloadMat.Left
                .Top = 190
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 0
            End With
            Me.AddChild(CType(stmBodyMat, UIControl))

            'cboPayloadType initial props
            cboPayloadType = New UIComboBox(oUILib)
            With cboPayloadType
                .ControlName = "cboPayloadType"
                .Left = 135
                .Top = 165
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
                .ToolTipText = "Indication of the type of damage the payload" & vbCrLf & "of the missile inflicts. Missiles always" & vbCrLf & "inflict Impact damage in addition to the payload."
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(cboPayloadType, UIControl))

            AddHandler tpAgility.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpBodyMat.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpExplosionRadius.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpFlapsMat.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpFuelMat.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpHoming.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpHullSize.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpMaxDmg.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpNoseMat.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpPayloadMat.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpRange.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpROF.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpStructHP.PropertyValueChanged, AddressOf tp_ValueChanged

            AddHandler stmBodyMat.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmNoseMat.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmFlapsMat.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmFuelMat.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmPayloadMat.SetButtonClicked, AddressOf SetButtonClicked

            FillValues()

            mbLoading = False
        End Sub

        Private Sub FillValues()

            'Dim lSorted() As Int32 = GetSortedMineralIdxArray(False)

            'cboPayloadMat.Clear() : cboFuelMat.Clear() : cboFlapsMat.Clear() : cboNoseMat.Clear() : cboBodyMat.Clear() : cboPayloadType.Clear()
            ''Now... loop through our minerals
            'If lSorted Is Nothing = False Then
            '    For X As Int32 = 0 To lSorted.GetUpperBound(0)
            '        If lSorted(X) <> -1 Then
            '            'If goMinerals(lSorted(X)).IsFullyStudied = False Then Continue For
            '            cboPayloadMat.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboPayloadMat.ItemData(cboPayloadMat.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboFuelMat.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboFuelMat.ItemData(cboFuelMat.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboFlapsMat.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboFlapsMat.ItemData(cboFlapsMat.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboNoseMat.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboNoseMat.ItemData(cboNoseMat.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboBodyMat.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboBodyMat.ItemData(cboBodyMat.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '        End If
            '    Next X
            'End If

            cboPayloadType.AddItem("Explosive")
            cboPayloadType.ItemData(cboPayloadType.NewIndex) = 0

            Dim lPayloadAvail As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.ePayloadTypeAvailable)
            If lPayloadAvail > 0 Then
                cboPayloadType.AddItem("Chemical")
                cboPayloadType.ItemData(cboPayloadType.NewIndex) = 1
            End If
            'If lPayloadAvail > 1 Then
            '    cboPayloadType.AddItem("Magnetic")
            '    cboPayloadType.ItemData(cboPayloadType.NewIndex) = 2
            'End If
            cboPayloadType.ListIndex = 0

            SetCboHelpers()
        End Sub

        Public Overrides Function SubmitDesign(ByVal sName As String, ByVal yWeaponTypeID As Byte) As Boolean
            'If True = True Then
            '    MyBase.moUILib.AddNotification("Functionality not yet implemented.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '    Return False
            'End If

            If goCurrentEnvir Is Nothing Then Return False
            If mlEntityIndex = -1 Then
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID AndAlso goCurrentEnvir.oEntity(X).bSelected Then
                        mlEntityIndex = X
                        Exit For
                    End If
                Next X
            End If
            If mlEntityIndex = -1 Then Return False
            If mbImpossibleDesign = True Then
                MyBase.moUILib.AddNotification("You must fix the flaws of this design.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                Return False
            End If

            'If cboPayloadMat.ListIndex = -1 OrElse cboFuelMat.ListIndex = -1 OrElse cboFlapsMat.ListIndex = -1 OrElse cboNoseMat.ListIndex = -1 OrElse cboBodyMat.ListIndex = -1 Then
            '    MyBase.moUILib.AddNotification("Please select a material for all entries provided.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '    Return False
            'End If
            For X As Int32 = 0 To mlMineralIDs.GetUpperBound(0)
                If mlMineralIDs(X) < 1 Then
                    MyBase.moUILib.AddNotification("Please select a material for all entries provided.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                End If
            Next X

            Dim yMsg(153) As Byte '76) As Byte
            Dim lPos As Int32 = 0

            'Same for all techs...
            System.BitConverter.GetBytes(GlobalMessageCode.eSubmitComponentPrototype).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(ObjectType.eWeaponTech).CopyTo(yMsg, lPos) : lPos += 2

            If moTech Is Nothing = False Then
                System.BitConverter.GetBytes(moTech.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
            Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
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
            oResearchGuid.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6

            'Now... send our WeaponClassType...
            yMsg(lPos) = WeaponClassType.eMissile : lPos += 1
            yMsg(lPos) = yWeaponTypeID : lPos += 1

            With CType(Me.ParentControl, frmWeaponBuilder).mfrmBuilderCost
                Dim lHullReq As Int32 = -1
                Dim lPowerReq As Int32 = -1
                Dim blResCost As Int64 = -1
                Dim blResTime As Int64 = -1
                Dim blProdCost As Int64 = -1
                Dim blProdTime As Int64 = -1
                Dim lColonists As Int32 = -1
                Dim lEnlisted As Int32 = -1
                Dim lOfficers As Int32 = -1
                .GetLockedValues(lHullReq, lPowerReq, blResCost, blResTime, blProdCost, blProdTime, lColonists, lEnlisted, lOfficers)

                System.BitConverter.GetBytes(lHullReq).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lPowerReq).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(blResCost).CopyTo(yMsg, lPos) : lPos += 8
                System.BitConverter.GetBytes(blResTime).CopyTo(yMsg, lPos) : lPos += 8
                System.BitConverter.GetBytes(blProdCost).CopyTo(yMsg, lPos) : lPos += 8
                System.BitConverter.GetBytes(blProdTime).CopyTo(yMsg, lPos) : lPos += 8
                System.BitConverter.GetBytes(lColonists).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lEnlisted).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(lOfficers).CopyTo(yMsg, lPos) : lPos += 4

                Dim lTempMin As Int32 = -1
                If tpBodyMat.PropertyLocked = True Then lTempMin = CInt(tpBodyMat.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpNoseMat.PropertyLocked = True Then lTempMin = CInt(tpNoseMat.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpFlapsMat.PropertyLocked = True Then lTempMin = CInt(tpFlapsMat.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpFuelMat.PropertyLocked = True Then lTempMin = CInt(tpFuelMat.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpPayloadMat.PropertyLocked = True Then lTempMin = CInt(tpPayloadMat.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4

                yMsg(lPos) = CByte(Me.lHullTypeID) : lPos += 1
            End With

            'Now... for the data specific this weapon type...
            'Need to take the value...
            Dim blValue As Int64 = tpROF.PropertyValue
            'If chkSingleShot.Value = True Then blValue = -1
            If blValue <> -1 Then
                'fValue *= 30.0F
                If blValue < 30 Then
                    MyBase.moUILib.AddNotification("Minimum ROF is 1 second", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                ElseIf blValue > Int16.MaxValue Then
                    MyBase.moUILib.AddNotification("Maximum value for ROF is: " & (Int16.MaxValue / 30.0F).ToString("#,###.#0"), Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                End If
            End If

            Dim lMaxDmg As Int32 = CInt(tpMaxDmg.PropertyValue)
            If lMaxDmg < 1 Then
                MyBase.moUILib.AddNotification("Please enter a valid Max Damage entry.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Dim lHullSize As Int32 = CInt(tpHullSize.PropertyValue)
            If lHullSize < 1 Then
                MyBase.moUILib.AddNotification("Please enter a valid Hull Size.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            ElseIf lHullSize > Int16.MaxValue Then
                MyBase.moUILib.AddNotification("Please enter a smaller Hull Size.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Dim lMaxSpeed As Int32 = CInt(tpAgility.PropertyValue)
            If lMaxSpeed < 1 OrElse lMaxSpeed > 255 Then
                MyBase.moUILib.AddNotification("Max Speed is limited to a number between 1 and 255.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Dim lManeuver As Int32 = lMaxSpeed \ 2
            If lManeuver < 1 Then lManeuver = 1
            If lManeuver < 1 OrElse lManeuver > 255 Then
                MyBase.moUILib.AddNotification("Maneuver is limited to a number between 1 and 255.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Dim blRange As Int64 = tpRange.PropertyValue
            If blRange < 1 Then
                MyBase.moUILib.AddNotification("Please enter a valid Range.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            Else
                'blFlightTime *= 30.0F
                If blRange > 255 Then
                    MyBase.moUILib.AddNotification("Range is too high.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                End If
            End If
            Dim lHoming As Int32 = CInt(tpHoming.PropertyValue)
            If lHoming < 0 OrElse lHoming > 255 Then
                MyBase.moUILib.AddNotification("Valid range of values for Homing Accuracy is 0 to 255.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If cboPayloadType.ListIndex = -1 Then
                MyBase.moUILib.AddNotification("Please select a Payload type for this missile.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Dim lExpRadius As Int32 = CInt(tpExplosionRadius.PropertyValue)
            If lExpRadius < 0 OrElse lExpRadius > 255 Then
                MyBase.moUILib.AddNotification("Explosion Radius values must be between 0 and 255.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Dim lStructureHP As Int32 = CInt(tpStructHP.PropertyValue)
            If lStructureHP < 0 Then
                MyBase.moUILib.AddNotification("Please enter a valid Structure Hitpoints (must be a positive number).", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            'Now, translate it
            Dim iActualROF As Int16 = CShort(blValue)
            'ROF (2)
            System.BitConverter.GetBytes(iActualROF).CopyTo(yMsg, lPos) : lPos += 2
            'MaxDmg (4)
            System.BitConverter.GetBytes(lMaxDmg).CopyTo(yMsg, lPos) : lPos += 4
            'HullSize (4)
            System.BitConverter.GetBytes(lHullSize).CopyTo(yMsg, lPos) : lPos += 4
            'MaxSpeed (1)
            yMsg(lPos) = CByte(lMaxSpeed) : lPos += 1
            'Maneuver (1)
            yMsg(lPos) = CByte(lManeuver) : lPos += 1
            'FlightTime (2)
            System.BitConverter.GetBytes(CShort(blRange)).CopyTo(yMsg, lPos) : lPos += 2
            'Homing Accuracy (1)
            yMsg(lPos) = CByte(lHoming) : lPos += 1
            'Payload Type (1)
            yMsg(lPos) = CByte(cboPayloadType.ItemData(cboPayloadType.ListIndex)) : lPos += 1
            'Exp Radius (1)
            yMsg(lPos) = CByte(lExpRadius) : lPos += 1
            'StructHP
            System.BitConverter.GetBytes(lStructureHP).CopyTo(yMsg, lPos) : lPos += 4

            'Now, the 5 minerals
            'System.BitConverter.GetBytes(cboBodyMat.ItemData(cboBodyMat.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
            'System.BitConverter.GetBytes(cboNoseMat.ItemData(cboNoseMat.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
            'System.BitConverter.GetBytes(cboFlapsMat.ItemData(cboFlapsMat.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
            'System.BitConverter.GetBytes(cboFuelMat.ItemData(cboFuelMat.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
            'System.BitConverter.GetBytes(cboPayloadMat.ItemData(cboPayloadMat.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
            For X As Int32 = 0 To mlMineralIDs.GetUpperBound(0)
                System.BitConverter.GetBytes(mlMineralIDs(X)).CopyTo(yMsg, lPos) : lPos += 4
            Next X

            Dim sTemp As String = Mid$(sName, 1, 20)
            System.Text.ASCIIEncoding.ASCII.GetBytes(sTemp).CopyTo(yMsg, lPos) : lPos += 20

            MyBase.moUILib.SendMsgToPrimary(yMsg)

            Return True
        End Function

        Public Overrides Function ValidateData() As Boolean

            For X As Int32 = 0 To mlMineralIDs.GetUpperBound(0)
                If mlMineralIDs(X) < 1 Then
                    MyBase.moUILib.AddNotification("Please select a material for all entries provided.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                Else
                    For Y As Int32 = 0 To glMineralUB
                        If glMineralIdx(Y) = mlMineralIDs(X) Then
                            If goMinerals(Y).bDiscovered = False Then
                                MyBase.moUILib.AddNotification("You must select minerals that you have researched.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                                Return False
                            End If
                            Exit For
                        End If
                    Next Y
                End If
            Next X

            Dim fValue As Single = CSng(tpROF.PropertyValue) ' * 30.0F
            'If chkSingleShot.Value = True Then fValue = -1

            If fValue <> -1 Then
                'fValue *= 30.0F
                If fValue <= 0 Then
                    MyBase.moUILib.AddNotification("Please enter a valid ROF", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                ElseIf fValue > Int16.MaxValue Then
                    MyBase.moUILib.AddNotification("Maximum value for ROF is: " & (Int16.MaxValue / 30.0F).ToString("#,###.#0"), Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                End If
            End If

            Dim lMaxDmg As Int32 = CInt(tpMaxDmg.PropertyValue)
            If lMaxDmg < 1 Then
                MyBase.moUILib.AddNotification("Please enter a valid Max Damage entry.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Dim lHullSize As Int32 = CInt(tpHullSize.PropertyValue)
            'If txtHullSize.Caption.IndexOf(".") <> -1 Then
            '    MyBase.moUILib.AddNotification("Hull Size must be a whole number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '    Return False
            'End If
            If lHullSize < 1 Then
                MyBase.moUILib.AddNotification("Please enter a valid Hull Size.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            ElseIf lHullSize > Int16.MaxValue Then
                MyBase.moUILib.AddNotification("Please enter a smaller Hull Size.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Dim lMaxSpeed As Int32 = CInt(tpAgility.PropertyValue)
            'If txtMaxSpeed.Caption.IndexOf(".") <> -1 Then
            '    MyBase.moUILib.AddNotification("Max Speed must be a whole number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '    Return False
            'End If
            If lMaxSpeed < 1 OrElse lMaxSpeed > 255 Then
                MyBase.moUILib.AddNotification("Max Speed is limited to a number between 1 and 255.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            'Dim lManeuver As Int32 = CInt(tpManeuver.PropertyValue)
            ''If txtManeuver.Caption.IndexOf(".") <> -1 Then
            ''    MyBase.moUILib.AddNotification("Maneuver must be a whole number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            ''    Return False
            ''End If
            'If lManeuver < 1 OrElse lManeuver > 255 Then
            '    MyBase.moUILib.AddNotification("Maneuver is limited to a number between 1 and 255.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '    Return False
            'End If
            Dim blRange As Int64 = tpRange.PropertyValue
            If blRange < 1 Then
                MyBase.moUILib.AddNotification("Please enter a valid Range.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Dim lHoming As Int32 = CInt(tpHoming.PropertyValue)
            'If txtHoming.Caption.IndexOf(".") <> -1 Then
            '    MyBase.moUILib.AddNotification("Homing Accuracy must be a whole number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '    Return False
            'End If
            If lHoming < 0 OrElse lHoming > 255 Then
                MyBase.moUILib.AddNotification("Valid range of values for Homing Accuracy is 0 to 255.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If cboPayloadType.ListIndex = -1 Then
                MyBase.moUILib.AddNotification("Please select a Payload type for this missile.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Dim lExpRadius As Int32 = CInt(tpExplosionRadius.PropertyValue)
            'If txtExplosionRadius.Caption.IndexOf(".") <> -1 Then
            '    MyBase.moUILib.AddNotification("Explosion Radius must be a whole number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '    Return False
            'End If
            If lExpRadius < 0 OrElse lExpRadius > 255 Then
                MyBase.moUILib.AddNotification("Explosion Radius values must be between 0 and 255.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Dim lStructureHP As Int32 = CInt(tpStructHP.PropertyValue)
            'If txtStructHP.Caption.IndexOf(".") <> -1 Then
            '    MyBase.moUILib.AddNotification("Structure Hitpoints must be a whole number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '    Return False
            'End If
            If lStructureHP < 0 Then
                MyBase.moUILib.AddNotification("Please enter a valid Structure Hitpoints (must be a positive number).", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Return True
        End Function

        Public Overrides Sub ViewResults(ByRef oTech As WeaponTech, ByVal lProdFactor As Integer)
            If oTech Is Nothing Then Return

            moTech = oTech

            With CType(oTech, MissileWeaponTech)
                mbIgnoreValueChange = True
                Me.cboPayloadType.FindComboItemData(.PayloadType)
                mbLoading = True

                MyBase.lHullTypeID = .HullTypeID

                tpMaxDmg.PropertyValue = .MaximumDamage
                If .MaximumDamage > 0 Then tpMaxDmg.PropertyLocked = True
                tpHullSize.PropertyValue = .MissileHullSize
                If .MissileHullSize > 0 Then tpHullSize.PropertyLocked = True
                tpAgility.PropertyValue = .MaxSpeed
                If .MaxSpeed > 0 Then tpAgility.PropertyLocked = True
                tpRange.PropertyValue = .FlightTime
                If .FlightTime > 0 Then tpRange.PropertyLocked = True
                tpHoming.PropertyValue = .HomingAccuracy
                If .HomingAccuracy > 0 Then tpHoming.PropertyLocked = True
                tpExplosionRadius.PropertyValue = .ExplosionRadius
                If .ExplosionRadius > 0 Then tpExplosionRadius.PropertyLocked = True
                tpStructHP.PropertyValue = .StructuralHP
                If .StructuralHP > 0 Then tpStructHP.PropertyLocked = True
                If tpROF.MaxValue < .ROF Then tpROF.MaxValue = .ROF
                tpROF.PropertyValue = .ROF
                If .ROF > 0 Then tpROF.PropertyLocked = True

                mlSelectedMineralIdx = 0 : Mineral_Selected(.BodyMaterialID)
                mlSelectedMineralIdx = 1 : Mineral_Selected(.NoseMaterialID)
                mlSelectedMineralIdx = 2 : Mineral_Selected(.FlapsMaterialID)
                mlSelectedMineralIdx = 3 : Mineral_Selected(.FuelMaterialID)
                mlSelectedMineralIdx = 4 : Mineral_Selected(.PayloadMaterialID)
                'If .BodyMaterialID > 0 Then Me.cboBodyMat.FindComboItemData(.BodyMaterialID)
                'If .NoseMaterialID > 0 Then Me.cboNoseMat.FindComboItemData(.NoseMaterialID)
                'If .FlapsMaterialID > 0 Then Me.cboFlapsMat.FindComboItemData(.FlapsMaterialID)
                'If .FuelMaterialID > 0 Then Me.cboFuelMat.FindComboItemData(.FuelMaterialID)
                'If .PayloadMaterialID > 0 Then Me.cboPayloadMat.FindComboItemData(.PayloadMaterialID)

                tpBodyMat.PropertyValue = .lSpecifiedMin1
                tpNoseMat.PropertyValue = .lSpecifiedMin2
                tpFlapsMat.PropertyValue = .lSpecifiedMin3
                tpFuelMat.PropertyValue = .lSpecifiedMin4
                tpPayloadMat.PropertyValue = .lSpecifiedMin5
                If .lSpecifiedMin1 <> -1 Then tpBodyMat.PropertyLocked = True
                If .lSpecifiedMin2 <> -1 Then tpNoseMat.PropertyLocked = True
                If .lSpecifiedMin3 <> -1 Then tpFlapsMat.PropertyLocked = True
                If .lSpecifiedMin4 <> -1 Then tpFuelMat.PropertyLocked = True
                If .lSpecifiedMin5 <> -1 Then tpPayloadMat.PropertyLocked = True

                MyBase.mfrmBuilderCost.SetAndLockValues(oTech)

                mbLoading = False
                mbIgnoreValueChange = False
                BuilderCostValueChange(False)
            End With

        End Sub

        Private Sub SetCboHelpers()
            Dim oSB As New System.Text.StringBuilder()

            Dim bHardness As Boolean = goCurrentPlayer.PlayerKnowsProperty("HARDNESS")
            Dim bDensity As Boolean = goCurrentPlayer.PlayerKnowsProperty("DENSITY")
            Dim bMalleable As Boolean = goCurrentPlayer.PlayerKnowsProperty("MALLEABLE")
            Dim bCompress As Boolean = goCurrentPlayer.PlayerKnowsProperty("COMPRESSIBILITY")
            Dim bThermExp As Boolean = goCurrentPlayer.PlayerKnowsProperty("THERMAL EXPANSION")
            Dim bTempSens As Boolean = goCurrentPlayer.PlayerKnowsProperty("TEMPERATURE SENSITIVITY")
            Dim bChemReact As Boolean = goCurrentPlayer.PlayerKnowsProperty("CHEMICAL REACTANCE")
            Dim bBoilingPt As Boolean = goCurrentPlayer.PlayerKnowsProperty("BOILING POINT")
            Dim bCombust As Boolean = goCurrentPlayer.PlayerKnowsProperty("COMBUSTIVENESS")

            ''Body
            oSB.Length = 0
            'If bDensity = True Then 
            oSB.Append("Low Density materials ") 'Else oSB.Append("Materials ")
            'If bHardness = True Then 
            oSB.Append("with high Hardness properties ")
            'If bMalleable = True Then 
            oSB.Append("that are highly Malleable ")
            oSB.Append("work best.")
            'If bDensity = True OrElse bHardness = True OrElse bMalleable = True Then 
            stmBodyMat.ToolTipText = oSB.ToString 'Else cboBodyMat.ToolTipText = "We are not quite sure of the best way to engineer this."

            ''Nose
            oSB.Length = 0
            'If bHardness = True Then 
            oSB.Append("Hard materials ") 'Else oSB.Append("Materials ")
            'If bCompress = True Then 
            oSB.Append("with low Compressibility properties ")
            oSB.Append("are preferred.")
            'If bHardness = True OrElse bCompress = True Then 
            stmNoseMat.ToolTipText = oSB.ToString 'Else cboNoseMat.ToolTipText = "We are not quite sure of the best way to engineer this."

            ''Flaps
            oSB.Length = 0
            'If bDensity = True Then 
            oSB.Append("High Density materials ") 'Else oSB.Append("Materials ")
            'If bTempSens = True Then 
            oSB.Append("that are not sensitive to temperature ")
            'If bThermExp = True Then 
            oSB.Append("with low Thermal Expansion attributes ")
            oSB.Append("provide the best results.")
            'If bDensity = True OrElse bTempSens = True OrElse bThermExp = True Then 
            stmFlapsMat.ToolTipText = oSB.ToString 'Else cboFlapsMat.ToolTipText = "We are not quite sure of the best way to engineer this."

            ''Fuel
            oSB.Length = 0
            'If bDensity = True Then 
            oSB.Append("Low Density materials ") 'Else oSB.Append("Materials ")
            'If bBoilingPt = True Then 
            oSB.Append("with a low Boiling Point ")
            'If bChemReact = True Then 
            oSB.Append("that are very Chemically Reactive ")
            oSB.Append("work best.")
            'If bDensity = True OrElse bBoilingPt = True OrElse bChemReact = True Then cboFuelMat.ToolTipText = oSB.ToString Else cboFuelMat.ToolTipText = "We are not quite sure of the best way to engineer this."
            stmFuelMat.ToolTipText = oSB.ToString

            'tpROF.ToolTipText = "Minimum Rate of Fire is 1 second"
            tpRange.MaxValue = 255
            tpRange.MinValue = 10
            tpRange.PropertyValue = tpRange.MinValue

            tpROF.MinValue = 900 '240
            tpROF.PropertyValue = tpROF.MinValue
            tpROF.MaxValue = 2000
            tpROF.blAbsoluteMaximum = 30000
            tpROF.ToolTipText = "Number of seconds between shots. Minimum for this is 30 seconds."
            tpHullSize.MaxValue = 10

            tpRange.ToolTipText = "Range for the missile. Value can be between 10 and 255."

            ''Payload
            If cboPayloadType.ListIndex <> -1 Then
                If cboPayloadType.ItemData(cboPayloadType.ListIndex) = 0 Then
                    'explosive
                    'If bCombust = True Then 
                    stmPayloadMat.ToolTipText = "Highly Combustiveness materials are needed." 'Else cboPayloadMat.ToolTipText = "We are not sure of the best way to engineer this."
                Else
                    'chemical
                    'If bChemReact = True Then 
                    stmPayloadMat.ToolTipText = "Materials that are very Chemically Reactive are needed." 'Else cboPayloadMat.ToolTipText = "We are not sure of the best way to engineer this."
                    'NOTE: Add more if needed...
                End If
            End If

            tpAgility.MinValue = 20
            tpHullSize.MinValue = 1
            tpAgility.SetToInitialDefault()
            tpHoming.SetToInitialDefault()
            tpMaxDmg.SetToInitialDefault()
            tpHullSize.SetToInitialDefault()
            tpRange.SetToInitialDefault()
            tpROF.PropertyValue = 1800
            'tpStructHP.SetToInitialDefault()


        End Sub

        Private Sub cboPayloadType_ItemSelected(ByVal lItemIndex As Integer) Handles cboPayloadType.ItemSelected
            If mbLoading = True Then Return
            SetCboHelpers()
            CheckForDARequest()
        End Sub

        'Private Sub chkSingleShot_Click() Handles chkSingleShot.Click
        '    If mbLoading = True Then Return
        '    mbLoading = True
        '    If chkSingleShot.Value = True Then
        '        txtROF.Caption = "-1"
        '        txtROF.Enabled = False
        '    Else
        '        txtROF.Caption = "1"
        '        txtROF.Enabled = True
        '    End If
        '    mbLoading = False
        'End Sub

        '     Private Sub txtROF_TextChanged() Handles txtROF.TextChanged
        '         If mbLoading = True Then Return
        'If (Val(txtROF.Caption)) < 1.0F Then
        '	If chkSingleShot.Value = False Then chkSingleShot.Value = True
        'Else
        '	If chkSingleShot.Value = True Then chkSingleShot.Value = False
        'End If
        '     End Sub

        Private Sub fraMissiles_OnRender() Handles Me.OnRender
            'Ok, first, is moTech nothing?
            If moTech Is Nothing = False Then
                'Ok, now, do the values used on the form currently match the tech?
                Dim bChanged As Boolean = False
                With CType(moTech, MissileWeaponTech)
                    bChanged = .MaximumDamage <> tpMaxDmg.PropertyValue OrElse .MissileHullSize <> tpHullSize.PropertyValue OrElse _
                      .MaxSpeed <> tpAgility.PropertyValue OrElse tpRange.PropertyValue <> .FlightTime OrElse .ExplosionRadius <> tpExplosionRadius.PropertyValue OrElse _
                      tpHoming.PropertyValue <> .HomingAccuracy OrElse tpStructHP.PropertyValue <> .StructuralHP OrElse tpROF.PropertyValue <> .ROF

                    'If bChanged = False Then
                    '    Dim fROF As Single = 0
                    '    If chkSingleShot.Value = True Then
                    '        fROF = -1
                    '    Else : fROF = CSng(Val(txtROF.Caption))
                    '    End If

                    '    bChanged = fROF.ToString("#,###.#0") <> (.ROF / 30.0F).ToString("#,###.#0")
                    'End If
                    If bChanged = False Then
                        Dim lPayload As Int32 = mlMineralIDs(4)
                        Dim lFuel As Int32 = mlMineralIDs(3)
                        Dim lFlaps As Int32 = mlMineralIDs(2)
                        Dim lNose As Int32 = mlMineralIDs(1)
                        Dim lBody As Int32 = mlMineralIDs(0)
                        Dim lPayloadtype As Int32 = -1

                        'If cboPayloadMat.ListIndex <> -1 Then lPayload = cboPayloadMat.ItemData(cboPayloadMat.ListIndex)
                        'If cboFuelMat.ListIndex <> -1 Then lFuel = cboFuelMat.ItemData(cboFuelMat.ListIndex)
                        'If cboFlapsMat.ListIndex <> -1 Then lFlaps = cboFlapsMat.ItemData(cboFlapsMat.ListIndex)
                        'If cboNoseMat.ListIndex <> -1 Then lNose = cboNoseMat.ItemData(cboNoseMat.ListIndex)
                        'If cboBodyMat.ListIndex <> -1 Then lBody = cboBodyMat.ItemData(cboBodyMat.ListIndex)
                        If cboPayloadType.ListIndex <> -1 Then lPayloadtype = cboPayloadType.ItemData(cboPayloadType.ListIndex)

                        bChanged = .PayloadMaterialID <> lPayload OrElse .FuelMaterialID <> lFuel OrElse lFlaps <> .FlapsMaterialID OrElse _
                          lNose <> .NoseMaterialID OrElse lBody <> .BodyMaterialID OrElse lPayloadtype <> .PayloadType OrElse (tpBodyMat.PropertyLocked <> (.lSpecifiedMin1 <> -1)) OrElse _
                          (tpNoseMat.PropertyLocked <> (.lSpecifiedMin2 <> -1)) OrElse (tpFlapsMat.PropertyLocked <> (.lSpecifiedMin3 <> -1)) OrElse _
                          (tpFuelMat.PropertyLocked <> (.lSpecifiedMin4 <> -1)) OrElse (tpPayloadMat.PropertyLocked <> (.lSpecifiedMin5 <> -1))

                        If bChanged = False Then

                            bChanged = (.lSpecifiedMin1 <> -1 AndAlso .lSpecifiedMin1 <> tpBodyMat.PropertyValue) OrElse _
                                        (.lSpecifiedMin2 <> -1 AndAlso .lSpecifiedMin2 <> tpNoseMat.PropertyValue) OrElse _
                                        (.lSpecifiedMin3 <> -1 AndAlso .lSpecifiedMin3 <> tpFlapsMat.PropertyValue) OrElse _
                                        (.lSpecifiedMin4 <> -1 AndAlso .lSpecifiedMin4 <> tpFuelMat.PropertyValue) OrElse _
                                        (.lSpecifiedMin5 <> -1 AndAlso .lSpecifiedMin5 <> tpPayloadMat.PropertyValue)
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


                    Dim oFrm As frmWeaponBuilder = CType(MyBase.moUILib.GetWindow("frmWeaponBuilder"), frmWeaponBuilder)
                    If oFrm Is Nothing = False Then
                        oFrm.SetDataChanged(bChanged)
                    End If
                End With
            End If
        End Sub

        'Private Sub COMBO_DropDownExpanded(ByVal sComboBoxName As String) Handles cboBodyMat.DropDownExpanded, cboFlapsMat.DropDownExpanded, cboFuelMat.DropDownExpanded, cboNoseMat.DropDownExpanded, cboPayloadMat.DropDownExpanded
        '    If sComboBoxName = "" Then Return

        '    Dim oTech As New MissileTechComputer
        '    oTech.lHullTypeID = Me.lHullTypeID

        '    If cboPayloadType.ListIndex > -1 Then
        '        oTech.yPayloadType = CByte(cboPayloadType.ItemData(cboPayloadType.ListIndex))
        '    End If

        '    Dim lMinIdx As Int32 = -1
        '    Select Case sComboBoxName
        '        Case cboBodyMat.ControlName
        '            lMinIdx = 0
        '        Case cboNoseMat.ControlName
        '            lMinIdx = 1
        '        Case cboFlapsMat.ControlName
        '            lMinIdx = 2
        '        Case cboFuelMat.ControlName
        '            lMinIdx = 3
        '        Case cboPayloadMat.ControlName
        '            lMinIdx = 4
        '    End Select
        '    If lMinIdx = -1 Then Return
        '    oTech.MineralCBOExpanded(lMinIdx, ObjectType.eWeaponTech)
        'End Sub

        Private Sub tp_ValueChanged(ByVal sPropName As String, ByRef ctl As ctlTechProp)
            tpROF.MaxValue = Math.Min(Int16.MaxValue, Math.Max(tpROF.MaxValue, tpROF.PropertyValue + 500))
            BuilderCostValueChange(False)
        End Sub
        'Private Sub ComboValueChanged(ByVal lIndex As Int32) Handles cboPayloadType.ItemSelected
        '    BuilderCostValueChange(False)
        'End Sub

        Public Overrides Sub BuilderCostValueChange(ByVal bAutoFill As Boolean)
            If mbIgnoreValueChange = True Then Return
            mbIgnoreValueChange = True
            mbImpossibleDesign = True
            Dim oTechComputer As New MissileTechComputer
            With oTechComputer

                Dim decNormalizer As Decimal = 1D
                Dim lMaxGuns As Int32 = 1
                Dim lMaxDPS As Int32 = 1
                Dim lMaxHullSize As Int32 = 1
                Dim lHullAvail As Int32 = 1
                Dim lMaxPower As Int32 = 1

                If MyBase.lHullTypeID = -1 Then
                    SetDesignFlaw("Invalid Design: a Hull Type Must be selected!")
                    mbIgnoreValueChange = False
                    Return
                End If

                TechBuilderComputer.GetTypeValues(MyBase.lHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)
                AddToTestString("MaxDPS: " & lMaxDPS.ToString)

                tpMaxDmg.MaxValue = lMaxDPS * 10 * CLng(tpROF.PropertyValue / 30D)
                If tpMaxDmg.PropertyValue > tpMaxDmg.MaxValue Then tpMaxDmg.PropertyValue = tpMaxDmg.MaxValue
                tpHullSize.MaxValue = 10
                If tpHullSize.PropertyValue > tpHullSize.MaxValue Then tpHullSize.PropertyValue = tpHullSize.MaxValue
                tpStructHP.MaxValue = 30
                If tpStructHP.PropertyValue > tpStructHP.MaxValue Then tpStructHP.PropertyValue = tpStructHP.MaxValue

                With CType(Me.ParentControl, frmWeaponBuilder)
                    oTechComputer.SetMinDAValues(.ml_Min1DA, .ml_Min2DA, .ml_Min3DA, .ml_Min4DA, .ml_Min5DA, .ml_Min6DA)
                End With

                Dim fROF As Single = Math.Max(1, tpROF.PropertyValue) / 30.0F
                .decDPS = CDec(tpMaxDmg.PropertyValue) / CDec(fROF)
                .decDPS = Math.Max(.decDPS, (.decDPS + tpMaxDmg.PropertyValue) / 2D)
                'Dim decNominal As Decimal = (CDec(lMaxPower / lMaxDPS) * 3D) * .decDPS

                .lHullTypeID = Me.lHullTypeID
                .blExplosionRadius = tpExplosionRadius.PropertyValue
                .blHomingAccuracy = tpHoming.PropertyValue
                .blHullSize = tpHullSize.PropertyValue
                .blManeuver = tpAgility.PropertyValue \ 2
                If tpAgility.PropertyValue > 0 AndAlso .blManeuver = 0 Then .blManeuver = 1
                .blMaxDamage = tpMaxDmg.PropertyValue
                .blMaxSpeed = tpAgility.PropertyValue
                .blRange = tpRange.PropertyValue
                .blROF = tpROF.PropertyValue
                .blStructHP = tpStructHP.PropertyValue

                If cboPayloadType.ListIndex > -1 Then
                    .yPayloadType = CByte(cboPayloadType.ItemData(cboPayloadType.ListIndex))
                End If

                'If cboBodyMat.ListIndex > -1 Then
                '    .lMineral1ID = cboBodyMat.ItemData(cboBodyMat.ListIndex)
                'Else : .lMineral1ID = -1
                'End If
                'If cboNoseMat.ListIndex > -1 Then
                '    .lMineral2ID = cboNoseMat.ItemData(cboNoseMat.ListIndex)
                'Else : .lMineral2ID = -1
                'End If
                'If cboFlapsMat.ListIndex > -1 Then
                '    .lMineral3ID = cboFlapsMat.ItemData(cboFlapsMat.ListIndex)
                'Else : .lMineral3ID = -1
                'End If
                'If cboFuelMat.ListIndex > -1 Then
                '    .lMineral4ID = cboFuelMat.ItemData(cboFuelMat.ListIndex)
                'Else : .lMineral4ID = -1
                'End If
                'If cboPayloadMat.ListIndex > -1 Then
                '    .lMineral5ID = cboPayloadMat.ItemData(cboPayloadMat.ListIndex)
                'Else : .lMineral5ID = -1
                'End If
                .lMineral1ID = mlMineralIDs(0)
                .lMineral2ID = mlMineralIDs(1)
                .lMineral3ID = mlMineralIDs(2)
                .lMineral4ID = mlMineralIDs(3)
                .lMineral5ID = mlMineralIDs(4)
                .lMineral6ID = -1

                .msMin1Name = "Body"
                .msMin2Name = "Nose"
                .msMin3Name = "Flaps"
                .msMin4Name = "Fuel"
                .msMin5Name = "Payload"
                .msMin6Name = ""

                Dim lblDesignFlaw As UILabel = CType(MyBase.ParentControl, frmWeaponBuilder).lblDesignFlaw

                If .lMineral1ID < 1 OrElse .lMineral2ID < 1 OrElse .lMineral3ID < 1 OrElse .lMineral4ID < 1 OrElse .lMineral5ID < 1 Then
                    lblDesignFlaw.Caption = "All properties and materials need to be defined."
                    mbIgnoreValueChange = False
                    mbImpossibleDesign = True
                    Return
                End If

                If bAutoFill = True Then
                    .DoBuilderCostPreconfigure(lblDesignFlaw, MyBase.mfrmBuilderCost, tpBodyMat, tpNoseMat, tpFlapsMat, tpFuelMat, tpPayloadMat, Nothing, ObjectType.eWeaponTech, 0D, False)
                Else
                    .BuilderCostValueChange(lblDesignFlaw, MyBase.mfrmBuilderCost, tpBodyMat, tpNoseMat, tpFlapsMat, tpFuelMat, tpPayloadMat, Nothing, ObjectType.eWeaponTech, 0D)
                End If

                mbImpossibleDesign = .bImpossibleDesign
                If Me.ParentControl Is Nothing = False Then

                    Dim lMinFlame As Int32
                    Dim lMaxFlame As Int32
                    Dim lMinChemical As Int32
                    Dim lMaxChemical As Int32
                    Dim lMinImpact As Int32 = 0
                    Dim lMaxImpact As Int32 = 0

                    Dim lMaxDmg As Int32 = CInt(tpMaxDmg.PropertyValue)
                    lMaxImpact = lMaxDmg \ 10
                    lMinImpact = lMaxFlame \ 10
                    If .yPayloadType = 0 Then
                        lMaxFlame = lMaxDmg - lMaxImpact
                        lMinFlame = lMaxFlame \ 10
                    Else
                        lMaxChemical = lMaxDmg - lMaxImpact
                        lMinChemical = lMaxChemical \ 10
                    End If

                    CType(Me.ParentControl, frmWeaponBuilder).SetWeaponDetailedStats(lMinFlame, lMaxFlame, 0, 0, lMinChemical, lMaxChemical, 0, 0, lMinImpact, lMaxImpact, 0, 0, fROF)
                End If
            End With

            mbIgnoreValueChange = False
        End Sub

        Private mlMineralIDs(4) As Int32
        Private Sub SetButtonClicked(ByVal lMinIdx As Int32)

            Dim oTech As New MissileTechComputer
            oTech.lHullTypeID = Me.lHullTypeID

            If cboPayloadType.ListIndex > -1 Then
                oTech.yPayloadType = CByte(cboPayloadType.ItemData(cboPayloadType.ListIndex))
            End If

            If lMinIdx = -1 Then Return

            mlSelectedMineralIdx = -1
            oTech.MineralCBOExpanded(lMinIdx, ObjectType.eWeaponTech)
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
                        stmBodyMat.SetMineralName(sMinName)
                    Case 1
                        stmNoseMat.SetMineralName(sMinName)
                    Case 2
                        stmFlapsMat.SetMineralName(sMinName)
                    Case 3
                        stmFuelMat.SetMineralName(sMinName)
                    Case 4
                        stmPayloadMat.SetMineralName(sMinName)
                End Select
            End If

            CheckForDARequest()

        End Sub

        Public Overrides Sub CheckForDARequest()
            Dim yPayload As Byte = 0
            If cboPayloadType.ListIndex > -1 Then
                yPayload = CByte(cboPayloadType.ItemData(cboPayloadType.ListIndex))
            End If

            'BuilderCostValueChange(False)
            Dim bRequestDA As Boolean = True
            For X As Int32 = 0 To mlMineralIDs.GetUpperBound(0)
                If mlMineralIDs(X) < 1 Then
                    bRequestDA = False
                    Exit For
                End If
            Next X
            If bRequestDA = True Then CType(Me.ParentControl, frmWeaponBuilder).RequestDAValues(ObjectType.eWeaponTech, WeaponClassType.eMissile, CByte(lHullTypeID), mlMineralIDs(0), mlMineralIDs(1), mlMineralIDs(2), mlMineralIDs(3), mlMineralIDs(4), -1, yPayload, 0)
        End Sub

        Public Overrides Sub CloseFrame()
            Dim ofrmMin As frmMinDetail = CType(goUILib.GetWindow("frmMinDetail"), frmMinDetail)
            If ofrmMin Is Nothing Then Return
            ofrmMin.bRaiseSelectEvent = True
            Try
                RemoveHandler ofrmMin.MineralSelected, AddressOf Mineral_Selected
            Catch
            End Try
        End Sub
    End Class
End Class
