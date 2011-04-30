Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Partial Class frmWeaponBuilder
    Private Class fraSolidBeam
        Inherits fraWeaponBase

        'Private lblMaxDmg As UILabel
        'Private lblMaxRng As UILabel
        'Private lblROF As UILabel
        'Private lblAccuracy As UILabel
        'Private txtMaxDamage As UITextBox
        'Private txtMaxRange As UITextBox
        'Private txtROF As UITextBox
        'Private txtAccuracy As UITextBox

        Private tpMaxDmg As ctlTechProp
        Private tpMaxRng As ctlTechProp
        Private tpROF As ctlTechProp
        Private tpAccuracy As ctlTechProp
        Private tpMedium As ctlTechProp
        Private tpFocuser As ctlTechProp
        Private tpCasing As ctlTechProp
        Private tpCoupler As ctlTechProp
        Private tpCoil As ctlTechProp

        Private lblDmgType As UILabel
        Private lblCoilMat As UILabel
        Private lblCoupler As UILabel
        Private lblCasing As UILabel
        Private lblFocuserMat As UILabel
        Private lblMediumMat As UILabel

        'Private WithEvents cboMedium As UIComboBox
        'Private WithEvents cboFocuser As UIComboBox
        'Private WithEvents cboCasing As UIComboBox
        'Private WithEvents cboCoupler As UIComboBox
        'Private WithEvents cboCoil As UIComboBox
        Private stmMedium As ctlSetTechMineral
        Private stmFocuser As ctlSetTechMineral
        Private stmCasing As ctlSetTechMineral
        Private stmCoupler As ctlSetTechMineral
        Private stmCoil As ctlSetTechMineral

        Private WithEvents cboDmgType As UIComboBox

        Private mlEntityIndex As Int32 = -1
        Private moTech As WeaponTech = Nothing

        Private mlMaxOptRange As Int32 = 40
        Private mlMaxWpnAcc As Int32 = 25
        Private mlMaxDmg As Int32 = 40
        Private mlMinROF As Int32 = 360

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraSolidBeam initial props
            With Me
                .lWindowMetricID = BPMetrics.eWindow.eSolidBeamBuilder
                .ControlName = "fraSolidBeam"
                .Left = 18
                .Top = 52
                .Width = 470
                .Height = 335
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .mbAcceptReprocessEvents = True
                .Moveable = False
                .BorderLineWidth = 1
            End With

            'lblDmgType initial props
            lblDmgType = New UILabel(oUILib)
            With lblDmgType
                .ControlName = "lblDmgType"
                .Left = 10
                .Top = 110
                .Width = 180
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Beam Damage Type:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "Indicates the type of beam used and whether it is a cutting beam (Piercing)" & vbCrLf & "or a burning beam (Thermal). This also impacts the appearance of the beam."
            End With
            Me.AddChild(CType(lblDmgType, UIControl))

            'lblCoilMat initial props
            lblCoilMat = New UILabel(oUILib)
            With lblCoilMat
                .ControlName = "lblCoilMat"
                .Left = 10
                .Top = 150
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Coil:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblCoilMat, UIControl))

            'lblCoupler initial props
            lblCoupler = New UILabel(oUILib)
            With lblCoupler
                .ControlName = "lblCoupler"
                .Left = 10
                .Top = 175
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Coupler:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblCoupler, UIControl))

            'lblCasing initial props
            lblCasing = New UILabel(oUILib)
            With lblCasing
                .ControlName = "lblCasing"
                .Left = 10
                .Top = 200
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Casing:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblCasing, UIControl))

            'lblFocuserMat initial props
            lblFocuserMat = New UILabel(oUILib)
            With lblFocuserMat
                .ControlName = "lblFocuserMat"
                .Left = 10
                .Top = 225
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Focuser:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblFocuserMat, UIControl))

            'lblMediumMat initial props
            lblMediumMat = New UILabel(oUILib)
            With lblMediumMat
                .ControlName = "lblMediumMat"
                .Left = 10
                .Top = 250
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Medium:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblMediumMat, UIControl))

            ''lblMaxDmg initial props
            'lblMaxDmg = New UILabel(oUILib)
            'With lblMaxDmg
            '    .ControlName = "lblMaxDmg"
            '    .Left = 10
            '    .Top = 10
            '    .Width = 180
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Maximum Damage:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "The maximum damage this beam weapon will inflict" & vbCrLf & "Solid beam weapons differ from their Pulse beam" & vbCrLf & "couterparts in that solid beam weapons can" & vbCrLf & "inflict far greater amounts of damage." & vbCrLf & "Any positive whole number is a valid range for this value."
            'End With
            'Me.AddChild(CType(lblMaxDmg, UIControl))

            ''lblMaxRng initial props
            'lblMaxRng = New UILabel(oUILib)
            'With lblMaxRng
            '    .ControlName = "lblMaxRng"
            '    .Left = 10
            '    .Top = 35
            '    .Width = 180
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Maximum Range:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "The maximum range the beam can travel." & vbCrLf & "Valid value ranges are 0 to 255."
            'End With
            'Me.AddChild(CType(lblMaxRng, UIControl))

            ''lblROF initial props
            'lblROF = New UILabel(oUILib)
            'With lblROF
            '    .ControlName = "lblROF"
            '    .Left = 10
            '    .Top = 60
            '    .Width = 180
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Rate of Fire:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "A positive decimal number representing the number of seconds" & vbCrLf & "between shots. Solid beam weapons generate an enormous" & vbCrLf & "amount of heat and require large quantities of" & vbCrLf & "power to fire. Having a fast rate of fire greatly increases the costs."
            'End With
            'Me.AddChild(CType(lblROF, UIControl))

            ''lblAccuracy initial props
            'lblAccuracy = New UILabel(oUILib)
            'With lblAccuracy
            '    .ControlName = "lblAccuracy"
            '    .Left = 10
            '    .Top = 85
            '    .Width = 180
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Target Alignment Efficiency:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "A score representing the accuracy of the weapon." & vbCrLf & "Solid beam weapons have a harder time hitting smaller" & vbCrLf & "moving targets. Higher values for this score indicate" & vbCrLf & "an increased ability for the weapon to acquire the target." & vbCrLf & "Valid value range is 0 to 255."
            'End With
            'Me.AddChild(CType(lblAccuracy, UIControl))

            ''txtMaxDamage initial props
            'txtMaxDamage = New UITextBox(oUILib)
            'With txtMaxDamage
            '    .ControlName = "txtMaxDamage"
            '    .Left = 190
            '    .Top = 10
            '    .Width = 100
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
            '    .ToolTipText = lblMaxDmg.ToolTipText
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtMaxDamage, UIControl))
            tpMaxDmg = New ctlTechProp(oUILib)
            With tpMaxDmg
                .ControlName = "tpMaxDmg"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 255
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Maximum Damage:"
                .PropertyValue = 0
                .Top = 10
                .Visible = True
                .Width = 430
                .ToolTipText = "The maximum damage this beam weapon will inflict" & vbCrLf & "Solid beam weapons differ from their Pulse beam" & vbCrLf & "couterparts in that solid beam weapons can" & vbCrLf & "inflict far greater amounts of damage." & vbCrLf & "Any positive whole number is a valid range for this value."
            End With
            Me.AddChild(CType(tpMaxDmg, UIControl))

            ''txtMaxRange initial props
            'txtMaxRange = New UITextBox(oUILib)
            'With txtMaxRange
            '    .ControlName = "txtMaxRange"
            '    .Left = 190
            '    .Top = 35
            '    .Width = 100
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
            '    .ToolTipText = lblMaxRng.ToolTipText
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtMaxRange, UIControl))
            tpMaxRng = New ctlTechProp(oUILib)
            With tpMaxRng
                .ControlName = "tpMaxRng"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 255
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Maximum Range:"
                .PropertyValue = 0
                .Top = 35
                .Visible = True
                .Width = 430
                .ToolTipText = "The maximum range the beam can travel." & vbCrLf & "Valid value ranges are 0 to 255."
            End With
            Me.AddChild(CType(tpMaxRng, UIControl))

            ''txtROF initial props
            'txtROF = New UITextBox(oUILib)
            'With txtROF
            '    .ControlName = "txtROF"
            '    .Left = 190
            '    .Top = 60
            '    .Width = 100
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
            '    .ToolTipText = lblROF.ToolTipText
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
                .PropertyValue = 0
                .Top = 60
                .Visible = True
                .Width = 430
                .SetDivisor(30)
                .bNoMaxValue = True
                .ToolTipText = "A positive decimal number representing the number of seconds" & vbCrLf & "between shots. Solid beam weapons generate an enormous" & vbCrLf & "amount of heat and require large quantities of" & vbCrLf & "power to fire. Having a fast rate of fire greatly increases the costs."
            End With
            Me.AddChild(CType(tpROF, UIControl))

            ''txtAccuracy initial props
            'txtAccuracy = New UITextBox(oUILib)
            'With txtAccuracy
            '    .ControlName = "txtAccuracy"
            '    .Left = 190
            '    .Top = 85
            '    .Width = 100
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
            '    .ToolTipText = lblAccuracy.ToolTipText
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtAccuracy, UIControl))
            tpAccuracy = New ctlTechProp(oUILib)
            With tpAccuracy
                .ControlName = "tpAccuracy"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 255
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Accuracy:"
                .PropertyValue = 0
                .Top = 85
                .Visible = True
                .Width = 430
                .ToolTipText = "A positive decimal number representing the accuracy of the weapon."
            End With
            Me.AddChild(CType(tpAccuracy, UIControl))

            tpMedium = New ctlTechProp(oUILib)
            With tpMedium
                .ControlName = "tpMedium"
                .Enabled = True
                .Height = 18
                .Left = lblMediumMat.Left + lblMediumMat.Width + 160    '10 for left/right and 150 for cbo
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Medium:"
                .PropertyValue = 0
                .Top = 250
                .Visible = True
                .bNoMaxValue = True
                .Width = Me.Width - .Left - 10
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpMedium, UIControl))

            tpFocuser = New ctlTechProp(oUILib)
            With tpFocuser
                .ControlName = "tpFocuser"
                .Enabled = True
                .Height = 18
                .Left = tpMedium.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Focuser:"
                .PropertyValue = 0
                .Top = 225
                .Visible = True
                .Width = tpMedium.Width
                .bNoMaxValue = True
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpFocuser, UIControl))

            tpCasing = New ctlTechProp(oUILib)
            With tpCasing
                .ControlName = "tpCasing"
                .Enabled = True
                .Height = 18
                .Left = tpMedium.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Casing:"
                .PropertyValue = 0
                .Top = 200
                .Visible = True
                .Width = tpMedium.Width
                .bNoMaxValue = True
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpCasing, UIControl))

            tpCoupler = New ctlTechProp(oUILib)
            With tpCoupler
                .ControlName = "tpCoupler"
                .Enabled = True
                .Height = 18
                .Left = tpMedium.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Coupler:"
                .PropertyValue = 0
                .Top = 175
                .Visible = True
                .Width = tpMedium.Width
                .bNoMaxValue = True
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpCoupler, UIControl))

            tpCoil = New ctlTechProp(oUILib)
            With tpCoil
                .ControlName = "tpCoil"
                .Enabled = True
                .Height = 18
                .Left = tpMedium.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Coil:"
                .PropertyValue = 0
                .Top = 150
                .Visible = True
                .Width = tpMedium.Width
                .bNoMaxValue = True
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpCoil, UIControl))

            ''cboMedium initial props
            'cboMedium = New UIComboBox(oUILib)
            'With cboMedium
            '    .ControlName = "cboMedium"
            '    .Left = lblMediumMat.Left + lblMediumMat.Width + 5
            '    .Top = 250
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
            'End With
            'Me.AddChild(CType(cboMedium, UIControl))
            stmMedium = New ctlSetTechMineral(oUILib)
            With stmMedium
                .ControlName = "stmMedium"
                .Left = lblMediumMat.Left + lblMediumMat.Width + 5
                .Top = 250
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 4
            End With
            Me.AddChild(CType(stmMedium, UIControl))

            ''cboFocuser initial props
            'cboFocuser = New UIComboBox(oUILib)
            'With cboFocuser
            '    .ControlName = "cboFocuser"
            '    .Left = cboMedium.Left
            '    .Top = 225
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
            'End With
            'Me.AddChild(CType(cboFocuser, UIControl))
            stmFocuser = New ctlSetTechMineral(oUILib)
            With stmFocuser
                .ControlName = "stmFocuser"
                .Left = stmMedium.Left
                .Top = 225
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 3
            End With
            Me.AddChild(CType(stmFocuser, UIControl))

            ''cboCasing initial props
            'cboCasing = New UIComboBox(oUILib)
            'With cboCasing
            '    .ControlName = "cboCasing"
            '    .Left = cboMedium.Left
            '    .Top = 200
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
            'End With
            'Me.AddChild(CType(cboCasing, UIControl))
            stmCasing = New ctlSetTechMineral(oUILib)
            With stmCasing
                .ControlName = "stmCasing"
                .Left = stmMedium.Left
                .Top = 200
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 2
            End With
            Me.AddChild(CType(stmCasing, UIControl))

            ''cboCoupler initial props
            'cboCoupler = New UIComboBox(oUILib)
            'With cboCoupler
            '    .ControlName = "cboCoupler"
            '    .Left = cboMedium.Left
            '    .Top = 175
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
            'End With
            'Me.AddChild(CType(cboCoupler, UIControl))
            stmCoupler = New ctlSetTechMineral(oUILib)
            With stmCoupler
                .ControlName = "stmCoupler"
                .Left = stmMedium.Left
                .Top = 175
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 1
            End With
            Me.AddChild(CType(stmCoupler, UIControl))

            ''cboCoil initial props
            'cboCoil = New UIComboBox(oUILib)
            'With cboCoil
            '    .ControlName = "cboCoil"
            '    .Left = cboMedium.Left
            '    .Top = 150
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
            'End With
            'Me.AddChild(CType(cboCoil, UIControl))
            stmCoil = New ctlSetTechMineral(oUILib)
            With stmCoil
                .ControlName = "stmCoil"
                .Left = stmMedium.Left
                .Top = 150
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 0
            End With
            Me.AddChild(CType(stmCoil, UIControl))

            'cboDmgType initial props
            cboDmgType = New UIComboBox(oUILib)
            With cboDmgType
                .ControlName = "cboDmgType"
                .Left = 190
                .Top = 110
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
                .ToolTipText = lblDmgType.ToolTipText
            End With
            Me.AddChild(CType(cboDmgType, UIControl))

            AddHandler tpAccuracy.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpCasing.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpCoil.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpCoupler.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpFocuser.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpMaxDmg.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpMaxRng.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpMedium.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpROF.PropertyValueChanged, AddressOf tp_ValueChanged

            AddHandler stmCasing.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmCoil.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmCoupler.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmFocuser.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmMedium.SetButtonClicked, AddressOf SetButtonClicked

            FillValues()

        End Sub

        Private Sub FillValues()

            'Dim lSorted() As Int32 = GetSortedMineralIdxArray(False)

            'cboMedium.Clear() : cboFocuser.Clear() : cboCasing.Clear() : cboCoupler.Clear() : cboCoil.Clear()
            ''Now... loop through our minerals
            'If lSorted Is Nothing = False Then
            '    For X As Int32 = 0 To lSorted.GetUpperBound(0)
            '        If lSorted(X) <> -1 Then
            '            'If goMinerals(lSorted(X)).IsFullyStudied = False Then Continue For
            '            cboMedium.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboMedium.ItemData(cboMedium.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboFocuser.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboFocuser.ItemData(cboFocuser.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboCasing.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboCasing.ItemData(cboCasing.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboCoupler.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboCoupler.ItemData(cboCoupler.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboCoil.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboCoil.ItemData(cboCoil.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '        End If
            '    Next X
            'End If

            cboDmgType.Clear()
            cboDmgType.AddItem("Cutting Beam (Pierce)")
            cboDmgType.ItemData(cboDmgType.NewIndex) = 0
            cboDmgType.AddItem("Burning Beam (Thermal)")
            cboDmgType.ItemData(cboDmgType.NewIndex) = 1
            cboDmgType.ListIndex = 0

            SetCboHelpers()
        End Sub

        Public Overrides Function SubmitDesign(ByVal sName As String, ByVal yWeaponTypeID As Byte) As Boolean
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

            If ValidateData() = False Then Return False
            If mbImpossibleDesign = True Then
                MyBase.moUILib.AddNotification("You must fix the flaws of this design.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eUserInterface, Nothing, Nothing)
                Return False
            End If

            Dim yMsg(142) As Byte '65) As Byte
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
            yMsg(lPos) = WeaponClassType.eEnergyBeam : lPos += 1
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
                If tpCoil.PropertyLocked = True Then lTempMin = CInt(tpCoil.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpCoupler.PropertyLocked = True Then lTempMin = CInt(tpCoupler.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpCasing.PropertyLocked = True Then lTempMin = CInt(tpCasing.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpFocuser.PropertyLocked = True Then lTempMin = CInt(tpFocuser.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpMedium.PropertyLocked = True Then lTempMin = CInt(tpMedium.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4

                yMsg(lPos) = CByte(Me.lHullTypeID) : lPos += 1
            End With

            'Now... for the data specific this weapon type...
            'Need to take the value...
            Dim fValue As Single = CSng(tpROF.PropertyValue) ' * 30.0F

            If fValue <> -1 Then
                'fValue *= 30.0F
                If fValue = 0 Then
                    MyBase.moUILib.AddNotification("Please enter a valid ROF", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                ElseIf fValue > Int16.MaxValue Then
                    MyBase.moUILib.AddNotification("Maximum value for ROF is: " & (Int16.MaxValue / 30.0F).ToString("#,###.#0"), Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                End If
            End If

            'Now, translate it
            Dim iActualROF As Int16 = CShort(fValue)
            'ROF (2)
            System.BitConverter.GetBytes(iActualROF).CopyTo(yMsg, lPos) : lPos += 2
            'MaxDmg (4)
            System.BitConverter.GetBytes(CInt(tpMaxDmg.PropertyValue)).CopyTo(yMsg, lPos) : lPos += 4
            'MaxRng (2)
            System.BitConverter.GetBytes(CShort(tpMaxRng.PropertyValue)).CopyTo(yMsg, lPos) : lPos += 2
            'Accuracy (1)
            yMsg(lPos) = CByte(tpAccuracy.PropertyValue) : lPos += 1
            'DmgType (1)
            yMsg(lPos) = CByte(cboDmgType.ItemData(cboDmgType.ListIndex)) : lPos += 1

            'Now, the 5 minerals 
            System.BitConverter.GetBytes(mlMineralIDs(4)).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(mlMineralIDs(3)).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(mlMineralIDs(2)).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(mlMineralIDs(1)).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(mlMineralIDs(0)).CopyTo(yMsg, lPos) : lPos += 4

            Dim sTemp As String = Mid$(sName, 1, 20)
            System.Text.ASCIIEncoding.ASCII.GetBytes(sTemp).CopyTo(yMsg, lPos) : lPos += 20

            MyBase.moUILib.SendMsgToPrimary(yMsg)

            Return True
        End Function

        Public Overrides Function ValidateData() As Boolean

            'Dim sTemp As String = txtMaxDamage.Caption & txtMaxRange.Caption & txtAccuracy.Caption
            'If sTemp.IndexOf(".") <> -1 Then
            '    MyBase.moUILib.AddNotification("All numerical entries must be whole numbers!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '    Return False
            'End If
            'sTemp &= txtROF.Caption
            'If sTemp.IndexOf("-") <> -1 Then
            '    MyBase.moUILib.AddNotification("All numerical entries mut be non-negative!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '    Return False
            'End If

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
            If cboDmgType.ListIndex = -1 Then
                MyBase.moUILib.AddNotification("Please select a Damage Type.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            'Dim fValue As Single = CSng(tpROF.PropertyValue) ' * 30.0F
            'fValue *= 30.0F
            If tpROF.PropertyValue < mlMinROF Then
                MyBase.moUILib.AddNotification("Please enter a valid ROF that is at least " & (mlMinROF / 30.0F).ToString("#,##0.#0"), Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            ElseIf tpROF.PropertyValue > Int16.MaxValue Then
                MyBase.moUILib.AddNotification("Maximum value for ROF is: " & (Int16.MaxValue / 30.0F).ToString("#,###.#0"), Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            If tpMaxDmg.PropertyValue < 1 OrElse tpMaxDmg.PropertyValue > mlMaxDmg Then
                MyBase.moUILib.AddNotification("Please enter a valid maximum damage value (1 - " & mlMaxDmg & ").", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If tpMaxRng.PropertyValue < 1 OrElse tpMaxRng.PropertyValue > mlMaxOptRange Then
                MyBase.moUILib.AddNotification("Please enter a valid Maximum Range value (1 - " & mlMaxOptRange & ").", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If tpAccuracy.PropertyValue < 0 OrElse tpAccuracy.PropertyValue > mlMaxWpnAcc Then
                MyBase.moUILib.AddNotification("Please enter a valid Target Alignment Efficiency value (0 - " & mlMaxWpnAcc & ").", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            Return True
        End Function

        Public Overrides Sub ViewResults(ByRef oTech As WeaponTech, ByVal lProdFactor As Integer)
            If oTech Is Nothing Then Return

            moTech = oTech

            With CType(oTech, SolidBeamWeaponTech)
                mbIgnoreValueChange = True

                MyBase.lHullTypeID = .HullTypeID

                tpMaxDmg.PropertyValue = .MaxDamage
                If .MaxDamage > 0 Then tpMaxDmg.PropertyLocked = True
                tpMaxRng.PropertyValue = .MaxRange
                If .MaxRange > 0 Then tpMaxRng.PropertyLocked = True
                If tpROF.MaxValue < .ROF Then tpROF.MaxValue = .ROF
                tpROF.PropertyValue = .ROF
                If .ROF > 0 Then tpROF.PropertyLocked = True
                tpAccuracy.PropertyValue = .Accuracy
                If .Accuracy > 0 Then tpAccuracy.PropertyLocked = True

                If cboDmgType.FindComboItemData(.yDmgType) = True Then
                    Dim oWpnApp As frmWeaponAppearance = CType(MyBase.moUILib.GetWindow("frmWeaponAppearance"), frmWeaponAppearance)
                    If oWpnApp Is Nothing = False Then
                        oWpnApp.SetSolidBeamType(cboDmgType.ItemData(cboDmgType.ListIndex))
                        oWpnApp.SetWeaponTypeID(CByte(.WeaponTypeID))
                    End If
                    oWpnApp = Nothing
                End If
                mlSelectedMineralIdx = 0 : Mineral_Selected(.CoilID)
                mlSelectedMineralIdx = 1 : Mineral_Selected(.CouplerID)
                mlSelectedMineralIdx = 2 : Mineral_Selected(.CasingID)
                mlSelectedMineralIdx = 3 : Mineral_Selected(.FocuserID)
                mlSelectedMineralIdx = 4 : Mineral_Selected(.MediumID)
                'If .MediumID > 0 Then cboMedium.FindComboItemData(.MediumID)
                'If .FocuserID > 0 Then cboFocuser.FindComboItemData(.FocuserID)
                'If .CasingID > 0 Then cboCasing.FindComboItemData(.CasingID)
                'If .CouplerID > 0 Then cboCoupler.FindComboItemData(.CouplerID)
                'If .CoilID > 0 Then cboCoil.FindComboItemData(.CoilID)

                tpCoil.PropertyValue = .lSpecifiedMin1
                If .lSpecifiedMin1 <> -1 Then tpCoil.PropertyLocked = True
                tpCoupler.PropertyValue = .lSpecifiedMin2
                If .lSpecifiedMin2 <> -1 Then tpCoupler.PropertyLocked = True
                tpCasing.PropertyValue = .lSpecifiedMin3
                If .lSpecifiedMin3 <> -1 Then tpCasing.PropertyLocked = True
                tpFocuser.PropertyValue = .lSpecifiedMin4
                If .lSpecifiedMin4 <> -1 Then tpFocuser.PropertyLocked = True
                tpMedium.PropertyValue = .lSpecifiedMin5
                If .lSpecifiedMin5 <> -1 Then tpMedium.PropertyLocked = True

                MyBase.mfrmBuilderCost.SetAndLockValues(oTech)
                mbIgnoreValueChange = False
                BuilderCostValueChange(False)
            End With
        End Sub

        Private Sub SetCboHelpers()
            Dim oSB As New System.Text.StringBuilder()

            'Coil
            oSB.Length = 0
            oSB.Append("High Density materials ")
            oSB.Append("with a high Superconductive point ")
            oSB.Append("and high Magnetic Reactance with a low Magnetic Production ")
            oSB.Append("work best.")
            stmCoil.ToolTipText = oSB.ToString

            'Coupler
            stmCoupler.ToolTipText = "Quantum-sensitive, highly Reflective materials with low Thermal Expansion and Temperature Sensitivity properties are a must."

            'Casing
            stmCasing.ToolTipText = "High Density materials with a high Thermal Conductance and low Temperature Sensitivity and Malleable traits work best."

            'Focuser
            stmFocuser.ToolTipText = "Quantum-sensitive, highly Refractive materials with low Thermal Expansion and Temperature Sensitivity properties are a must."

            'Medium
            stmMedium.ToolTipText = "Quantum-sensitive materials with low Reflection and Refraction properties and a high Boiling Point are essential."

            If goCurrentPlayer Is Nothing = False Then
                mlMaxOptRange = Math.Min(goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eSolidBeamOptimumRange), 255)
                mlMaxWpnAcc = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eSolidBeamWeaponMaxAccuracy)
                mlMaxDmg = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eSolidBeamWpnMaxDmg)
                mlMinROF = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eSolidBeamWpnROF)
            End If

            tpMaxDmg.ToolTipText = "The maximum damage this beam weapon will inflict" & vbCrLf & "Solid beam weapons differ from their Pulse beam" & vbCrLf & "couterparts in that solid beam weapons can" & vbCrLf & "inflict far greater amounts of damage." & vbCrLf & "Any positive whole number less than " & mlMaxDmg & " is a valid range for this value."
            'txtMaxDamage.ToolTipText = lblMaxDmg.ToolTipText
            tpMaxDmg.MaxValue = mlMaxDmg
            tpMaxRng.ToolTipText = "The maximum range the beam can travel." & vbCrLf & "Valid value ranges are 10 to " & mlMaxOptRange & "."
            'txtMaxRange.ToolTipText = lblMaxRng.ToolTipText
            tpMaxRng.MaxValue = mlMaxOptRange
            tpMaxRng.MinValue = 10
            tpROF.ToolTipText = "A positive decimal number representing the number of seconds" & vbCrLf & "between shots. Solid beam weapons generate an enormous" & vbCrLf & "amount of heat and require large quantities of" & vbCrLf & "power to fire. Having a fast rate of fire greatly increases the costs." & vbCrLf & "This value cannot be less than " & (mlMinROF / 30.0F).ToString("#,##0.0###") & "."
            'txtROF.ToolTipText = lblROF.ToolTipText
            tpROF.MinValue = mlMinROF
            tpROF.PropertyValue = tpROF.MinValue
            tpROF.MaxValue = 2000
            tpROF.blAbsoluteMaximum = 30000

            tpAccuracy.ToolTipText = "A score representing the accuracy of the weapon." & vbCrLf & "Solid beam weapons have a harder time hitting smaller" & vbCrLf & "moving targets. Higher values for this score indicate" & vbCrLf & "an increased ability for the weapon to acquire the target." & vbCrLf & "Valid value range is 0 to " & mlMaxWpnAcc & "."
            'txtAccuracy.ToolTipText = lblAccuracy.ToolTipText
            tpAccuracy.MaxValue = mlMaxWpnAcc


            'set our defaults
            tpAccuracy.SetToInitialDefault()
            tpMaxDmg.SetToInitialDefault()
            tpMaxRng.SetToInitialDefault()

            tpROF.PropertyValue = 1800

        End Sub

        Private Sub fraSolidBeam_OnRender() Handles Me.OnRender
            'Ok, first, is moTech nothing?
            If moTech Is Nothing = False Then
                'Ok, now, do the values used on the form currently match the tech?
                Dim bChanged As Boolean = False
                With CType(moTech, SolidBeamWeaponTech)
                    bChanged = tpMaxDmg.PropertyValue <> .MaxDamage OrElse tpMaxRng.PropertyValue <> .MaxRange OrElse _
                      tpROF.PropertyValue <> .ROF OrElse tpAccuracy.PropertyValue <> .Accuracy

                    If bChanged = False Then
                        Dim lID1 As Int32 = mlMineralIDs(0)
                        Dim lID2 As Int32 = mlMineralIDs(1)
                        Dim lID3 As Int32 = mlMineralIDs(2)
                        Dim lID4 As Int32 = mlMineralIDs(3)
                        Dim lID5 As Int32 = mlMineralIDs(4)

                        Dim lID6 As Int32 = -1

                        'If cboCoil.ListIndex <> -1 Then lID1 = cboCoil.ItemData(cboCoil.ListIndex)
                        'If cboCoupler.ListIndex <> -1 Then lID2 = cboCoupler.ItemData(cboCoupler.ListIndex)
                        'If cboCasing.ListIndex <> -1 Then lID3 = cboCasing.ItemData(cboCasing.ListIndex)
                        'If cboFocuser.ListIndex <> -1 Then lID4 = cboFocuser.ItemData(cboFocuser.ListIndex)
                        'If cboMedium.ListIndex <> -1 Then lID5 = cboMedium.ItemData(cboMedium.ListIndex)
                        If cboDmgType.ListIndex <> -1 Then lID6 = cboDmgType.ItemData(cboDmgType.ListIndex)

                        bChanged = .CoilID <> lID1 OrElse .CouplerID <> lID2 OrElse .CasingID <> lID3 OrElse .FocuserID <> lID4 OrElse _
                          .MediumID <> lID5 OrElse .yDmgType <> lID6 OrElse (tpCoil.PropertyLocked <> (.lSpecifiedMin1 <> -1)) OrElse _
                          (tpCoupler.PropertyLocked <> (.lSpecifiedMin2 <> -1)) OrElse (tpCasing.PropertyLocked <> (.lSpecifiedMin3 <> -1)) OrElse _
                          (tpFocuser.PropertyLocked <> (.lSpecifiedMin4 <> -1)) OrElse (tpMedium.PropertyLocked <> (.lSpecifiedMin5 <> -1))

                        If bChanged = False Then

                            bChanged = (.lSpecifiedMin1 <> -1 AndAlso .lSpecifiedMin1 <> tpCoil.PropertyValue) OrElse _
                                        (.lSpecifiedMin2 <> -1 AndAlso .lSpecifiedMin2 <> tpCoupler.PropertyValue) OrElse _
                                        (.lSpecifiedMin3 <> -1 AndAlso .lSpecifiedMin3 <> tpCasing.PropertyValue) OrElse _
                                        (.lSpecifiedMin4 <> -1 AndAlso .lSpecifiedMin4 <> tpFocuser.PropertyValue) OrElse _
                                        (.lSpecifiedMin5 <> -1 AndAlso .lSpecifiedMin5 <> tpMedium.PropertyValue)
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

        Private Sub cboDmgType_ItemSelected(ByVal lItemIndex As Integer) Handles cboDmgType.ItemSelected
            If lItemIndex = -1 Then Return
            Dim oWpnApp As frmWeaponAppearance = CType(MyBase.moUILib.GetWindow("frmWeaponAppearance"), frmWeaponAppearance)
            If oWpnApp Is Nothing = False Then
                oWpnApp.SetSolidBeamType(cboDmgType.ItemData(lItemIndex))
            End If
            oWpnApp = Nothing
            BuilderCostValueChange(False)
        End Sub

        'Private Sub COMBO_DropDownExpanded(ByVal sComboBoxName As String) Handles cboMedium.DropDownExpanded, cboCasing.DropDownExpanded, cboCoupler.DropDownExpanded, cboCoil.DropDownExpanded, cboFocuser.DropDownExpanded
        '    Try

        '        If sComboBoxName = "" Then Return

        '        Dim oTech As New SolidTechComputer
        '        oTech.lHullTypeID = Me.lHullTypeID

        '        Dim lMinIdx As Int32 = -1
        '        Select Case sComboBoxName
        '            Case cboCoil.ControlName
        '                lMinIdx = 0
        '            Case cboCoupler.ControlName
        '                lMinIdx = 1
        '            Case cboCasing.ControlName
        '                lMinIdx = 2
        '            Case cboFocuser.ControlName
        '                lMinIdx = 3
        '            Case cboMedium.ControlName
        '                lMinIdx = 4
        '        End Select
        '        If lMinIdx = -1 Then Return
        '        oTech.MineralCBOExpanded(lMinIdx, ObjectType.eWeaponTech)
        '    Catch
        '    End Try
        'End Sub

        Private Sub tp_ValueChanged(ByVal sPropName As String, ByRef ctl As ctlTechProp)
            tpROF.MaxValue = Math.Min(Int16.MaxValue, Math.Max(tpROF.MaxValue, tpROF.PropertyValue + 500))
            BuilderCostValueChange(False)
        End Sub
        'Private Sub ComboValueChanged(ByVal lIndex As Int32) Handles cboCasing.ItemSelected, cboCoil.ItemSelected, cboCoupler.ItemSelected, cboFocuser.ItemSelected, cboMedium.ItemSelected
        '    BuilderCostValueChange()
        'End Sub
        Public Overrides Sub BuilderCostValueChange(ByVal bAutoFill As Boolean)
            If mbIgnoreValueChange = True Then Return
            mbIgnoreValueChange = True
            mbImpossibleDesign = True
            Dim oSolidTechComputer As New SolidTechComputer
            With oSolidTechComputer
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

                tpMaxDmg.MaxValue = Math.Min(mlMaxDmg, lMaxDPS * 100)
                If tpMaxDmg.PropertyValue > tpMaxDmg.MaxValue Then tpMaxDmg.PropertyValue = tpMaxDmg.MaxValue
                tpMaxDmg.ToolTipText = "The maximum damage this beam weapon will inflict" & vbCrLf & "Solid beam weapons differ from their Pulse beam" & vbCrLf & "couterparts in that solid beam weapons can" & vbCrLf & "inflict far greater amounts of damage." & vbCrLf & "Any positive whole number less than " & tpMaxDmg.MaxValue & " is a valid range for this value."

                Dim fROF As Single = Math.Max(1, tpROF.PropertyValue) / 30.0F
                Dim yDmgType As Byte = 0
                If cboDmgType.ListIndex > -1 Then
                    yDmgType = CByte(cboDmgType.ItemData(cboDmgType.ListIndex))
                End If
                Dim lMinPierce As Int32 = 0
                Dim lMaxPierce As Int32 = 0
                Dim lMinBurn As Int32 = 0
                Dim lMaxBurn As Int32 = 0
                Dim lMinBeam As Int32 = 0
                Dim lMaxBeam As Int32 = 0

                Dim lMaxDmg As Int32 = CInt(tpMaxDmg.PropertyValue)
                If yDmgType = 0 Then
                    'Piercing
                    Dim lVal As Int32 = CInt(lMaxDmg * 0.3F)
                    lMaxPierce = lVal
                    lVal = lMaxDmg - lVal
                    lMaxBeam = lVal
                Else
                    'thermal
                    Dim lVal As Int32 = CInt(lMaxDmg * 0.3F)
                    lMaxBurn = lVal
                    lVal = lMaxDmg - lVal
                    lMaxBeam = lVal
                End If
                lMinBeam = lMaxBeam \ 10

                With CType(Me.ParentControl, frmWeaponBuilder)
                    oSolidTechComputer.SetMinDAValues(.ml_Min1DA, .ml_Min2DA, .ml_Min3DA, .ml_Min4DA, .ml_Min5DA, .ml_Min6DA)
                End With

                .decDPS = tpMaxDmg.PropertyValue / CDec(fROF)
                .decDPS = Math.Max(.decDPS, (.decDPS + tpMaxDmg.PropertyValue) / 2D)
                'Dim decNominal As Decimal = (CDec(lMaxPower / lMaxDPS) * 3D) * .decDPS

                .blAccuracy = tpAccuracy.PropertyValue
                .blMaxDmg = tpMaxDmg.PropertyValue
                .blMaxRng = tpMaxRng.PropertyValue
                .lHullTypeID = MyBase.lHullTypeID

                'If cboCoil.ListIndex > -1 Then
                '    .lMineral1ID = cboCoil.ItemData(cboCoil.ListIndex)
                'Else : .lMineral1ID = -1
                'End If
                'If cboCoupler.ListIndex > -1 Then
                '    .lMineral2ID = cboCoupler.ItemData(cboCoupler.ListIndex)
                'Else : .lMineral2ID = -1
                'End If
                'If cboCasing.ListIndex > -1 Then
                '    .lMineral3ID = cboCasing.ItemData(cboCasing.ListIndex)
                'Else : .lMineral3ID = -1
                'End If
                'If cboFocuser.ListIndex > -1 Then
                '    .lMineral4ID = cboFocuser.ItemData(cboFocuser.ListIndex)
                'Else : .lMineral4ID = -1
                'End If
                'If cboMedium.ListIndex > -1 Then
                '    .lMineral5ID = cboMedium.ItemData(cboMedium.ListIndex)
                'Else : .lMineral5ID = -1
                'End If
                .lMineral1ID = mlMineralIDs(0)
                .lMineral2ID = mlMineralIDs(1)
                .lMineral3ID = mlMineralIDs(2)
                .lMineral4ID = mlMineralIDs(3)
                .lMineral5ID = mlMineralIDs(4)
                .lMineral6ID = -1

                .msMin1Name = "Coil"
                .msMin2Name = "Coupler"
                .msMin3Name = "Casing"
                .msMin4Name = "Focuser"
                .msMin5Name = "Medium"
                .msMin6Name = ""

                Dim lblDesignFlaw As UILabel = CType(MyBase.ParentControl, frmWeaponBuilder).lblDesignFlaw

                If .lMineral1ID < 1 OrElse .lMineral2ID < 1 OrElse .lMineral3ID < 1 OrElse .lMineral4ID < 1 OrElse .lMineral5ID < 1 Then
                    lblDesignFlaw.Caption = "All properties and materials need to be defined."
                    mbIgnoreValueChange = False
                    mbImpossibleDesign = True
                    Return
                End If

                If bAutoFill = True Then
                    .DoBuilderCostPreconfigure(lblDesignFlaw, MyBase.mfrmBuilderCost, tpCoil, tpCoupler, tpCasing, tpFocuser, tpMedium, Nothing, ObjectType.eWeaponTech, 0D, False)
                Else
                    .BuilderCostValueChange(lblDesignFlaw, MyBase.mfrmBuilderCost, tpCoil, tpCoupler, tpCasing, tpFocuser, tpMedium, Nothing, ObjectType.eWeaponTech, 0D)
                End If

                mbImpossibleDesign = .bImpossibleDesign
                If Me.ParentControl Is Nothing = False Then
                    CType(Me.ParentControl, frmWeaponBuilder).SetWeaponDetailedStats(lMinBurn, lMaxBurn, lMinBeam, lMaxBeam, 0, 0, 0, 0, 0, 0, lMinPierce, lMaxPierce, fROF)
                End If

            End With

            mbIgnoreValueChange = False
        End Sub

        Private mlMineralIDs(4) As Int32
        Private Sub SetButtonClicked(ByVal lMinIdx As Int32)

            Dim oTech As New SolidTechComputer
            oTech.lHullTypeID = Me.lHullTypeID

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
                        stmCoil.SetMineralName(sMinName)
                    Case 1
                        stmCoupler.SetMineralName(sMinName)
                    Case 2
                        stmCasing.SetMineralName(sMinName)
                    Case 3
                        stmFocuser.SetMineralName(sMinName)
                    Case 4
                        stmMedium.SetMineralName(sMinName)
                End Select
            End If

            'BuilderCostValueChange(False)
            CheckForDARequest()
        End Sub

        Public Overrides Sub CheckForDARequest()
            Dim bRequestDA As Boolean = True
            For X As Int32 = 0 To 4
                If mlMineralIDs(X) < 1 Then
                    bRequestDA = False
                    Exit For
                End If
            Next X
            If bRequestDA = True Then CType(Me.ParentControl, frmWeaponBuilder).RequestDAValues(ObjectType.eWeaponTech, WeaponClassType.eEnergyBeam, CByte(lHullTypeID), mlMineralIDs(0), mlMineralIDs(1), mlMineralIDs(2), mlMineralIDs(3), mlMineralIDs(4), -1, 0, 0)
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
