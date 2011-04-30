Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D
Partial Class frmWeaponBuilder

    'Interface created from Interface Builder
    Private Class fraPulse
        Inherits fraWeaponBase

        'Private lblInputEnergy As UILabel
        'Private lblCompress As UILabel
        'Private lblMaxRange As UILabel
        'Private lblROF As UILabel
        'Private lblScatterRadius As UILabel
        'Private txtInputEnergy As UITextBox
        'Private txtCompress As UITextBox
        'Private txtMaxRange As UITextBox
        'Private txtROF As UITextBox
        'Private txtScatterRadius As UITextBox
        Private tpInputEnergy As ctlTechProp
        Private tpCompress As ctlTechProp
        Private tpMaxRange As ctlTechProp
        Private tpROF As ctlTechProp
        Private tpScatterRadius As ctlTechProp

        Private lblCoil As UILabel
        Private lblAccMat As UILabel
        Private lblCasing As UILabel
        Private lblFocuser As UILabel
        Private lblCompChmbr As UILabel

        'Private WithEvents cboCoilMat As UIComboBox
        'Private WithEvents cboAccelerator As UIComboBox
        'Private WithEvents cboCasing As UIComboBox
        'Private WithEvents cboFocuser As UIComboBox
        'Private WithEvents cboCompressChmbr As UIComboBox
        Private stmCoilMat As ctlSetTechMineral
        Private stmAccelerator As ctlSetTechMineral
        Private stmCasing As ctlSetTechMineral
        Private stmFocuser As ctlSetTechMineral
        Private stmCompressChmbr As ctlSetTechMineral

        Private tpCoil As ctlTechProp
        Private tpAccelerator As ctlTechProp
        Private tpCasing As ctlTechProp
        Private tpFocuser As ctlTechProp
        Private tpCompressChmbr As ctlTechProp

        Private mlEntityIndex As Int32 = -1
        Private moTech As WeaponTech = Nothing

        Private mlMaxRange As Int32 = 30
        Private mlMaxCompress As Int32 = 2
        Private mlMinWpnROF As Int32 = 240
        Private mlMaxInputEnergy As Int32 = 30

        Private mbDoneInitialConfig As Boolean = False

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'frmWeaponBuilder.fraPulse initial props
            With Me
                .lWindowMetricID = BPMetrics.eWindow.ePulseBuilder
                .ControlName = "fraPulse"
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

            ''lblInputEnergy initial props
            'lblInputEnergy = New UILabel(oUILib)
            'With lblInputEnergy
            '    .ControlName = "lblInputEnergy"
            '    .Left = 10
            '    .Top = 10
            '    .Width = 100
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Input Energy:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "Indicates the power injected into the system." & vbCrLf & "This directly indicates the power requirement for the system."
            'End With
            'Me.AddChild(CType(lblInputEnergy, UIControl))

            ''txtInputEnergy initial props
            'txtInputEnergy = New UITextBox(oUILib)
            'With txtInputEnergy
            '    .ControlName = "txtInputEnergy"
            '    .Left = 150
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
            '    .ToolTipText = lblInputEnergy.ToolTipText
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtInputEnergy, UIControl))
            tpInputEnergy = New ctlTechProp(oUILib)
            With tpInputEnergy
                .ControlName = "tpInputEnergy"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 255
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Input Energy:"
                .PropertyValue = 0
                .Top = 10
                .Visible = True
                .Width = 430
                .ToolTipText = "Indicates the power injected into the system." & vbCrLf & "This directly indicates the power requirement for the system."
            End With
            Me.AddChild(CType(tpInputEnergy, UIControl))

            ''lblCompress initial props
            'lblCompress = New UILabel(oUILib)
            'With lblCompress
            '    .ControlName = "lblCompress"
            '    .Left = 10
            '    .Top = 35
            '    .Width = 130
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Compression Factor:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "The multiplier of power from the compression chamber." & vbCrLf & "Higher values here require purer materials for the Compression Chamber." & vbCrLf & "Value can be a decimal number but must be positive."
            'End With
            'Me.AddChild(CType(lblCompress, UIControl))

            ''txtCompress initial props
            'txtCompress = New UITextBox(oUILib)
            'With txtCompress
            '    .ControlName = "txtCompress"
            '    .Left = 150
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
            '    .MaxLength = 9
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .ToolTipText = lblCompress.ToolTipText
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtCompress, UIControl))
            tpCompress = New ctlTechProp(oUILib)
            With tpCompress
                .ControlName = "tpCompress"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 255
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Compression Factor:"
                .PropertyValue = 0
                .Top = 35
                .Visible = True
                .Width = 430
                .SetDivisor(10)
                .ToolTipText = "The multiplier of power from the compression chamber." & vbCrLf & "Higher values here require purer materials for the Compression Chamber." & vbCrLf & "Value can be a decimal number but must be positive."
            End With
            Me.AddChild(CType(tpCompress, UIControl))

            ''lblMaxRange initial props
            'lblMaxRange = New UILabel(oUILib)
            'With lblMaxRange
            '    .ControlName = "lblMaxRange"
            '    .Left = 10
            '    .Top = 60
            '    .Width = 130
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Optimum Range:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "The maximum range of the system. Pulse weapons " & vbCrLf & "are generally considered short ranged weapons" & vbCrLf & "as it becomes increasingly more difficult to keep the energy" & vbCrLf & "pulse together as it travels longer distances." & vbCrLf & "Valid value range is 0 to 255."
            'End With
            'Me.AddChild(CType(lblMaxRange, UIControl))

            ''txtMaxRange initial props
            'txtMaxRange = New UITextBox(oUILib)
            'With txtMaxRange
            '    .ControlName = "txtMaxRange"
            '    .Left = 150
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
            '    .MaxLength = 5
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .ToolTipText = lblMaxRange.ToolTipText
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtMaxRange, UIControl))
            tpMaxRange = New ctlTechProp(oUILib)
            With tpMaxRange
                .ControlName = "tpMaxRange"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 255
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Optimum Range:"
                .PropertyValue = 0
                .Top = 60
                .Visible = True
                .Width = 430
                .ToolTipText = "The maximum range of the system. Pulse weapons " & vbCrLf & "are generally considered short ranged weapons" & vbCrLf & "as it becomes increasingly more difficult to keep the energy" & vbCrLf & "pulse together as it travels longer distances." & vbCrLf & "Valid value range is 0 to 255."
            End With
            Me.AddChild(CType(tpMaxRange, UIControl))

            ''lblROF initial props
            'lblROF = New UILabel(oUILib)
            'With lblROF
            '    .ControlName = "lblROF"
            '    .Left = 10
            '    .Top = 85
            '    .Width = 130
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Rate of Fire:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "Number of seconds between shots." & vbCrLf & "It is much easier to make a pulse" & vbCrLf & "weapon fire faster than it is to fire farther." & vbCrLf & "Valid value range is any positive decimal number."
            'End With
            'Me.AddChild(CType(lblROF, UIControl))

            ''txtROF initial props
            'txtROF = New UITextBox(oUILib)
            'With txtROF
            '    .ControlName = "txtROF"
            '    .Left = 150
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
                .Top = 85
                .Visible = True
                .Width = 430
                .SetDivisor(30)
                .bNoMaxValue = True
                .ToolTipText = "Number of seconds between shots." & vbCrLf & "It is much easier to make a pulse" & vbCrLf & "weapon fire faster than it is to fire farther." & vbCrLf & "Valid value range is any positive decimal number."
            End With
            Me.AddChild(CType(tpROF, UIControl))

            ''lblScatterRadius initial props
            'lblScatterRadius = New UILabel(oUILib)
            'With lblScatterRadius
            '    .ControlName = "lblScatterRadius"
            '    .Left = 10
            '    .Top = 110
            '    .Width = 130
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Scatter Radius:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "An area effect damage component for the system." & vbCrLf & "Pulse weapons are the hardest to add an area effect" & vbCrLf & "damage element to as it directly contradicts the essence of the" & vbCrLf & "weapon. Increasing this value makes obtaining higher" & vbCrLf & "range even more difficult. Valid value Range is 0 to 255."
            'End With
            'Me.AddChild(CType(lblScatterRadius, UIControl))

            ''txtScatterRadois initial props
            'txtScatterRadius = New UITextBox(oUILib)
            'With txtScatterRadius
            '    .ControlName = "txtScatterRadius"
            '    .Left = 150
            '    .Top = 110
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
            '    .ToolTipText = lblScatterRadius.ToolTipText
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtScatterRadius, UIControl))
            tpScatterRadius = New ctlTechProp(oUILib)
            With tpScatterRadius
                .ControlName = "tpScatterRadius"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 255
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Scatter Radius:"
                .PropertyValue = 0
                .Top = 110
                .Visible = True
                .Width = 430
                .ToolTipText = "An area effect damage component for the system." & vbCrLf & "Pulse weapons are the hardest to add an area effect" & vbCrLf & "damage element to as it directly contradicts the essence of the" & vbCrLf & "weapon. Increasing this value makes obtaining higher" & vbCrLf & "range even more difficult. Valid value Range is 0 to 255."
            End With
            Me.AddChild(CType(tpScatterRadius, UIControl))

            'lblCoil initial props
            lblCoil = New UILabel(oUILib)
            With lblCoil
                .ControlName = "lblCoil"
                .Left = 10
                .Top = 155
                .Width = 70
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Coil:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblCoil, UIControl))

            'lblAccMat initial props
            lblAccMat = New UILabel(oUILib)
            With lblAccMat
                .ControlName = "lblAccMat"
                .Left = 10
                .Top = 180
                .Width = 70
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Accelerator:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblAccMat, UIControl))

            'lblCasing initial props
            lblCasing = New UILabel(oUILib)
            With lblCasing
                .ControlName = "lblCasing"
                .Left = 10
                .Top = 205
                .Width = 70
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

            'lblFocuser initial props
            lblFocuser = New UILabel(oUILib)
            With lblFocuser
                .ControlName = "lblFocuser"
                .Left = 10
                .Top = 230
                .Width = 70
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Focuser:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblFocuser, UIControl))

            'lblCompChmbr initial props
            lblCompChmbr = New UILabel(oUILib)
            With lblCompChmbr
                .ControlName = "lblCompChmbr"
                .Left = 10
                .Top = 255
                .Width = 70
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Chamber:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblCompChmbr, UIControl))

            tpCoil = New ctlTechProp(oUILib)
            With tpCoil
                .ControlName = "tpCoil"
                .Enabled = True
                .Height = 18
                .Left = lblCoil.Left + lblCoil.Width + 160      '10 for left/right and 150 for cbo
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Coil:"
                .PropertyValue = 0
                .Top = lblCoil.Top
                .Visible = True
                .Width = Me.Width - .Left - 10
                .bNoMaxValue = True
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpCoil, UIControl))

            tpAccelerator = New ctlTechProp(oUILib)
            With tpAccelerator
                .ControlName = "tpAccelerator"
                .Enabled = True
                .Height = 18
                .Left = tpCoil.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Accelerator:"
                .PropertyValue = 0
                .Top = lblAccMat.Top
                .Visible = True
                .Width = tpCoil.Width
                .bNoMaxValue = True
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpAccelerator, UIControl))

            tpCasing = New ctlTechProp(oUILib)
            With tpCasing
                .ControlName = "tpCasing"
                .Enabled = True
                .Height = 18
                .Left = tpCoil.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Casing:"
                .PropertyValue = 0
                .Top = lblCasing.Top
                .Visible = True
                .Width = tpCoil.Width
                .bNoMaxValue = True
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpCasing, UIControl))

            tpFocuser = New ctlTechProp(oUILib)
            With tpFocuser
                .ControlName = "tpFocuser"
                .Enabled = True
                .Height = 18
                .Left = tpCoil.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Focuser:"
                .PropertyValue = 0
                .Top = lblFocuser.Top
                .Visible = True
                .Width = tpCoil.Width
                .bNoMaxValue = True
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpFocuser, UIControl))

            tpCompressChmbr = New ctlTechProp(oUILib)
            With tpCompressChmbr
                .ControlName = "tpCompressChmbr"
                .Enabled = True
                .Height = 18
                .Left = tpCoil.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "CompressChamber:"
                .PropertyValue = 0
                .Top = lblCompChmbr.Top
                .Visible = True
                .Width = tpCoil.Width
                .bNoMaxValue = True
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpCompressChmbr, UIControl))

            '================= combo boxes =======================

            ''cboCompressChmbr initial props
            'cboCompressChmbr = New UIComboBox(oUILib)
            'With cboCompressChmbr
            '    .ControlName = "cboCompressChmbr"
            '    .Left = lblCompChmbr.Left + lblCompChmbr.Width + 5
            '    .Top = 255
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
            'Me.AddChild(CType(cboCompressChmbr, UIControl))
            stmCompressChmbr = New ctlSetTechMineral(oUILib)
            With stmCompressChmbr
                .ControlName = "stmCompressChmbr"
                .Left = lblCompChmbr.Left + lblCompChmbr.Width + 5
                .Top = 255
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 4
            End With
            Me.AddChild(CType(stmCompressChmbr, UIControl))

            ''cboFocuser initial props
            'cboFocuser = New UIComboBox(oUILib)
            'With cboFocuser
            '    .ControlName = "cboFocuser"
            '    .Left = cboCompressChmbr.Left
            '    .Top = 230
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
                .Left = stmCompressChmbr.Left
                .Top = 230
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
            '    .Left = cboCompressChmbr.Left
            '    .Top = 205
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
                .Left = stmCompressChmbr.Left
                .Top = 205
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 2
            End With
            Me.AddChild(CType(stmCasing, UIControl))

            ''cboAccelerator initial props
            'cboAccelerator = New UIComboBox(oUILib)
            'With cboAccelerator
            '    .ControlName = "cboAccelerator"
            '    .Left = cboCompressChmbr.Left
            '    .Top = 180
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
            'Me.AddChild(CType(cboAccelerator, UIControl))
            stmAccelerator = New ctlSetTechMineral(oUILib)
            With stmAccelerator
                .ControlName = "stmAccelerator"
                .Left = stmCompressChmbr.Left
                .Top = 180
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 1
            End With
            Me.AddChild(CType(stmAccelerator, UIControl))

            ''cboCoilMat initial props
            'cboCoilMat = New UIComboBox(oUILib)
            'With cboCoilMat
            '    .ControlName = "cboCoilMat"
            '    .Left = cboCompressChmbr.Left
            '    .Top = 155
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
            'Me.AddChild(CType(cboCoilMat, UIControl))
            stmCoilMat = New ctlSetTechMineral(oUILib)
            With stmCoilMat
                .ControlName = "stmCoilMat"
                .Left = stmCompressChmbr.Left
                .Top = 155
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 0
            End With
            Me.AddChild(CType(stmCoilMat, UIControl))

            AddHandler tpAccelerator.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpCasing.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpCoil.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpCompress.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpCompressChmbr.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpFocuser.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpInputEnergy.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpMaxRange.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpROF.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpScatterRadius.PropertyValueChanged, AddressOf tp_ValueChanged

            AddHandler stmAccelerator.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmCasing.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmCoilMat.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmCompressChmbr.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmFocuser.SetButtonClicked, AddressOf SetButtonClicked

            FillValues()

        End Sub

        Private Sub FillValues()

            'Dim lSorted() As Int32 = GetSortedMineralIdxArray(False)

            'cboCoilMat.Clear() : cboAccelerator.Clear() : cboCasing.Clear() : cboFocuser.Clear() : cboCompressChmbr.Clear()
            ''Now... loop through our minerals
            'If lSorted Is Nothing = False Then
            '    Dim lSortedUB As Int32 = lSorted.GetUpperBound(0)
            '    For X As Int32 = 0 To lSortedUB
            '        If lSorted(X) <> -1 Then
            '            'If goMinerals(lSorted(X)).IsFullyStudied = False Then Continue For
            '            cboCoilMat.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboCoilMat.ItemData(cboCoilMat.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboAccelerator.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboAccelerator.ItemData(cboAccelerator.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboCasing.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboCasing.ItemData(cboCasing.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboFocuser.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboFocuser.ItemData(cboFocuser.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboCompressChmbr.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboCompressChmbr.ItemData(cboCompressChmbr.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '        End If
            '    Next X
            'End If

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

            Dim yMsg(144) As Byte '67) As Byte
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
            yMsg(lPos) = WeaponClassType.eEnergyPulse : lPos += 1
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
                If tpAccelerator.PropertyLocked = True Then lTempMin = CInt(tpAccelerator.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpCasing.PropertyLocked = True Then lTempMin = CInt(tpCasing.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpFocuser.PropertyLocked = True Then lTempMin = CInt(tpFocuser.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpCompressChmbr.PropertyLocked = True Then lTempMin = CInt(tpCompressChmbr.PropertyValue) Else lTempMin = -1
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
                If fValue < goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.ePulseBeamWpnROF) Then
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
            'InputEnergy (4)
            System.BitConverter.GetBytes(CInt(tpInputEnergy.PropertyValue)).CopyTo(yMsg, lPos) : lPos += 4
            'Compression Factor (4)
            System.BitConverter.GetBytes(CSng(tpCompress.PropertyValue / 10.0F)).CopyTo(yMsg, lPos) : lPos += 4
            'MaxRng (1)
            yMsg(lPos) = CByte(tpMaxRange.PropertyValue) : lPos += 1
            'ScatterRadius (1)
            yMsg(lPos) = CByte(tpScatterRadius.PropertyValue) : lPos += 1

            'Now, the 5 minerals
            'System.BitConverter.GetBytes(cboCoilMat.ItemData(cboCoilMat.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
            'System.BitConverter.GetBytes(cboAccelerator.ItemData(cboAccelerator.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
            'System.BitConverter.GetBytes(cboCasing.ItemData(cboCasing.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
            'System.BitConverter.GetBytes(cboFocuser.ItemData(cboFocuser.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
            'System.BitConverter.GetBytes(cboCompressChmbr.ItemData(cboCompressChmbr.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
            For X As Int32 = 0 To mlMineralIDs.GetUpperBound(0)
                System.BitConverter.GetBytes(mlMineralIDs(X)).CopyTo(yMsg, lPos) : lPos += 4
            Next X

            Dim sTemp As String = Mid$(sName, 1, 20)
            System.Text.ASCIIEncoding.ASCII.GetBytes(sTemp).CopyTo(yMsg, lPos) : lPos += 20

            MyBase.moUILib.SendMsgToPrimary(yMsg)

            Return True
        End Function

        Public Overrides Function ValidateData() As Boolean

            'Dim sTemp As String = txtInputEnergy.Caption & txtCompress.Caption & txtMaxRange.Caption & txtScatterRadius.Caption & txtROF.Caption
            'If sTemp.IndexOf("-") <> -1 Then
            '    MyBase.moUILib.AddNotification("All numerical values must be non-negative numbers.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
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
            'If txtInputEnergy.Caption.IndexOf(".") <> -1 Then
            '    MyBase.moUILib.AddNotification("Input Energy must be a whole number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '    Return False
            'End If
            If tpInputEnergy.PropertyValue > mlMaxInputEnergy OrElse tpInputEnergy.PropertyValue < 1 Then
                MyBase.moUILib.AddNotification("Please enter a valid Input Energy value (1 - " & mlMaxInputEnergy & ").", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If tpCompress.PropertyValue / 10.0F > mlMaxCompress OrElse tpCompress.PropertyValue < 1 Then
                MyBase.moUILib.AddNotification("Please enter a valid Compression Factor value (1 - " & mlMaxCompress & ").", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            'If txtMaxRange.Caption.IndexOf(".") <> -1 Then
            '    MyBase.moUILib.AddNotification("Maximum Range must be a whole number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '    Return False
            'End If
            If tpMaxRange.PropertyValue > mlMaxRange OrElse tpMaxRange.PropertyValue > 255 OrElse tpMaxRange.PropertyValue < 1 Then
                MyBase.moUILib.AddNotification("Please enter a valid Optimum Range (1 to " & Math.Min(255, mlMaxRange) & ").", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            'If txtScatterRadius.Caption.IndexOf(".") <> -1 Then
            '    MyBase.moUILib.AddNotification("Scatter Radius must be a whole number.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '    Return False
            'End If
            If tpScatterRadius.PropertyValue > 255 OrElse tpScatterRadius.PropertyValue < 0 Then
                MyBase.moUILib.AddNotification("Please enter a valid Scatter Radius (0 to 255).", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If tpROF.PropertyValue < goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.ePulseBeamWpnROF) Then
                MyBase.moUILib.AddNotification("Please enter a valid value for Rate of Fire.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            Return True
        End Function

        Public Overrides Sub ViewResults(ByRef oTech As WeaponTech, ByVal lProdFactor As Integer)
            If oTech Is Nothing Then Return

            moTech = oTech

            With CType(oTech, PulseWeaponTech)
                mbIgnoreValueChange = True

                MyBase.lHullTypeID = .HullTypeID

                If mbDoneInitialConfig = False Then
                    mbDoneInitialConfig = True
                    tpCompress.PropertyValue = 10
                    tpInputEnergy.SetToInitialDefault()
                    tpMaxRange.SetToInitialDefault()
                End If

                tpInputEnergy.PropertyValue = .InputEnergy
                If .InputEnergy > 0 Then tpInputEnergy.PropertyLocked = True
                tpCompress.PropertyValue = CLng(.CompressFactor * 10)
                If tpCompress.PropertyValue > 0 Then tpCompress.PropertyLocked = True
                tpMaxRange.PropertyValue = .MaxRange
                If .MaxRange > 0 Then tpMaxRange.PropertyLocked = True
                If tpROF.MaxValue < .ROF Then tpROF.MaxValue = .ROF
                tpROF.PropertyValue = .ROF
                If .ROF > 0 Then tpROF.PropertyLocked = True
                tpScatterRadius.PropertyValue = .ScatterRadius
                If .ScatterRadius > 0 Then tpScatterRadius.PropertyLocked = True

                mlSelectedMineralIdx = 0 : Mineral_Selected(.CoilMatID)
                mlSelectedMineralIdx = 1 : Mineral_Selected(.AcceleratorMatID)
                mlSelectedMineralIdx = 2 : Mineral_Selected(.CasingMatID)
                mlSelectedMineralIdx = 3 : Mineral_Selected(.FocuserMatID)
                mlSelectedMineralIdx = 4 : Mineral_Selected(.CompressChmbrMatID)
                'If .CoilMatID > 0 Then cboCoilMat.FindComboItemData(.CoilMatID)
                'If .AcceleratorMatID > 0 Then cboAccelerator.FindComboItemData(.AcceleratorMatID)
                'If .CasingMatID > 0 Then cboCasing.FindComboItemData(.CasingMatID)
                'If .FocuserMatID > 0 Then cboFocuser.FindComboItemData(.FocuserMatID)
                'If .CompressChmbrMatID > 0 Then cboCompressChmbr.FindComboItemData(.CompressChmbrMatID)

                tpCoil.PropertyValue = .lSpecifiedMin1
                If .lSpecifiedMin1 <> -1 Then tpCoil.PropertyLocked = True
                tpAccelerator.PropertyValue = .lSpecifiedMin2
                If .lSpecifiedMin2 <> -1 Then tpAccelerator.PropertyLocked = True
                tpCasing.PropertyValue = .lSpecifiedMin3
                If .lSpecifiedMin3 <> -1 Then tpCasing.PropertyLocked = True
                tpFocuser.PropertyValue = .lSpecifiedMin4
                If .lSpecifiedMin4 <> -1 Then tpFocuser.PropertyLocked = True
                tpCompressChmbr.PropertyValue = .lSpecifiedMin5
                If .lSpecifiedMin5 <> -1 Then tpCompressChmbr.PropertyLocked = True
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
            oSB.Append("and low Magnetic Reactance with a low Magnetic Production ")
            oSB.Append("work best.")
            stmCoilMat.ToolTipText = oSB.ToString

            'Accelerator
            oSB.Length = 0
            oSB.Append("Quantum-sensitive materials ")
            oSB.Append("with a high Superconductive Point that have High Magnetic Reactance and Low Magnetic Production properties ")
            oSB.Append("are preferred.")
            stmAccelerator.ToolTipText = oSB.ToString 'Else cboNoseMat.ToolTipText = "We are not quite sure of the best way to engineer this."

            'Casing
            oSB.Length = 0
            oSB.Append("High Density materials ") 'Else oSB.Append("Materials ")
            oSB.Append("that are not sensitive to temperature or Malleable ")
            oSB.Append("with high Thermal Conduction attributes ")
            oSB.Append("provide the best results.")
            stmCasing.ToolTipText = oSB.ToString 'Else cboFlapsMat.ToolTipText = "We are not quite sure of the best way to engineer this."

            'Focuser
            oSB.Length = 0
            oSB.Append("Quantum-sensitive materials ")
            oSB.Append("with a high Refraction property ")
            oSB.Append("that have a low Temperature Sensitivity and a very low Thermal Expansion property ")
            oSB.Append("work best.")
            stmFocuser.ToolTipText = oSB.ToString

            'Compression Chamber
            oSB.Length = 0
            oSB.Append("High Density materials ")
            oSB.Append("with a high Electrical Resistance ")
            oSB.Append("that are not compressible ")
            oSB.Append("with a very low Malleable property ")
            oSB.Append("are required.")
            stmCompressChmbr.ToolTipText = oSB.ToString

            If goCurrentPlayer Is Nothing = False Then
                mlMaxRange = Math.Min(255, goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.ePulseBeamOptimumRange))
                mlMaxCompress = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.ePulseBeamMaxCompressFactor)
                mlMaxInputEnergy = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.ePulseBeamInputEnergyMax)
                mlMinWpnROF = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.ePulseBeamWpnROF)
            End If

            tpMaxRange.ToolTipText = "The maximum range of the system. Pulse weapons " & vbCrLf & "are generally considered short ranged weapons" & vbCrLf & "as it becomes increasingly more difficult to keep the energy" & vbCrLf & "pulse together as it travels longer distances." & vbCrLf & "Valid value range is 10 to " & mlMaxRange & "."
            'txtMaxRange.ToolTipText = lblMaxRange.ToolTipText
            tpMaxRange.MaxValue = mlMaxRange
            tpMaxRange.MinValue = 10


            tpInputEnergy.ToolTipText = "Indicates the power injected into the system." & vbCrLf & "This directly indicates the power requirement for the system." & vbCrLf & "Maximum Value for this is based on the hull."
            'txtInputEnergy.ToolTipText = lblInputEnergy.ToolTipText

            tpCompress.ToolTipText = "The multiplier of power from the compression chamber." & vbCrLf & "Higher values here require purer materials for the Compression Chamber." & vbCrLf & "Value can be a decimal number but must be positive" & vbCrLf & "and cannot exceed " & mlMaxCompress & "."
            'txtCompress.ToolTipText = lblCompress.ToolTipText
            tpCompress.MaxValue = CLng(mlMaxCompress) * 10L

            tpROF.ToolTipText = "Number of seconds between shots." & vbCrLf & "It is much easier to make a pulse" & vbCrLf & "weapon fire faster than it is to fire farther." & vbCrLf & "Valid value range is any positive decimal" & vbCrLf & "number at least " & (mlMinWpnROF / 30.0F).ToString("#,##0.0##") & " or higher."
            'txtROF.ToolTipText = lblROF.ToolTipText
            tpROF.MinValue = mlMinWpnROF
            tpROF.MaxValue = 2000
            tpROF.PropertyValue = 1800 'mlMinWpnROF
            tpROF.blAbsoluteMaximum = 30000


            'tpCompress.SetToInitialDefault()
            tpCompress.PropertyValue = 10
            tpInputEnergy.SetToInitialDefault()
            tpMaxRange.SetToInitialDefault()
        End Sub

        Private Sub fraPulse_OnRender() Handles Me.OnRender
            'Ok, first, is moTech nothing?
            If moTech Is Nothing = False Then
                'Ok, now, do the values used on the form currently match the tech?
                Dim bChanged As Boolean = False
                With CType(moTech, PulseWeaponTech)
                    bChanged = tpInputEnergy.PropertyValue <> .InputEnergy OrElse tpCompress.PropertyValue / 10.0F <> .CompressFactor OrElse _
                      tpMaxRange.PropertyValue <> .MaxRange OrElse tpROF.PropertyValue <> .ROF OrElse _
                      tpScatterRadius.PropertyValue <> .ScatterRadius

                    If bChanged = False Then
                        Dim lID1 As Int32 = mlMineralIDs(0)
                        Dim lID2 As Int32 = mlMineralIDs(1)
                        Dim lID3 As Int32 = mlMineralIDs(2)
                        Dim lID4 As Int32 = mlMineralIDs(3)
                        Dim lID5 As Int32 = mlMineralIDs(4)

                        'If cboCoilMat.ListIndex <> -1 Then lID1 = cboCoilMat.ItemData(cboCoilMat.ListIndex)
                        'If cboAccelerator.ListIndex <> -1 Then lID2 = cboAccelerator.ItemData(cboAccelerator.ListIndex)
                        'If cboCasing.ListIndex <> -1 Then lID3 = cboCasing.ItemData(cboCasing.ListIndex)
                        'If cboFocuser.ListIndex <> -1 Then lID4 = cboFocuser.ItemData(cboFocuser.ListIndex)
                        'If cboCompressChmbr.ListIndex <> -1 Then lID5 = cboCompressChmbr.ItemData(cboCompressChmbr.ListIndex)

                        bChanged = .CoilMatID <> lID1 OrElse .AcceleratorMatID <> lID2 OrElse .CasingMatID <> lID3 OrElse _
                          .FocuserMatID <> lID4 OrElse .CompressChmbrMatID <> lID5 OrElse (tpCoil.PropertyLocked <> (.lSpecifiedMin1 <> -1)) OrElse _
                          (tpAccelerator.PropertyLocked <> (.lSpecifiedMin2 <> -1)) OrElse (tpCasing.PropertyLocked <> (.lSpecifiedMin3 <> -1)) OrElse _
                          (tpFocuser.PropertyLocked <> (.lSpecifiedMin4 <> -1)) OrElse (tpCompressChmbr.PropertyLocked <> (.lSpecifiedMin5 <> -1))

                        If bChanged = False Then

                            bChanged = (.lSpecifiedMin1 <> -1 AndAlso .lSpecifiedMin1 <> tpCoil.PropertyValue) OrElse _
                                        (.lSpecifiedMin2 <> -1 AndAlso .lSpecifiedMin2 <> tpAccelerator.PropertyValue) OrElse _
                                        (.lSpecifiedMin3 <> -1 AndAlso .lSpecifiedMin3 <> tpCasing.PropertyValue) OrElse _
                                        (.lSpecifiedMin4 <> -1 AndAlso .lSpecifiedMin4 <> tpFocuser.PropertyValue) OrElse _
                                        (.lSpecifiedMin5 <> -1 AndAlso .lSpecifiedMin5 <> tpCompressChmbr.PropertyValue)

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

        'Private Sub COMBO_DropDownExpanded(ByVal sComboBoxName As String) Handles cboAccelerator.DropDownExpanded, cboCasing.DropDownExpanded, cboCoilMat.DropDownExpanded, cboCompressChmbr.DropDownExpanded, cboFocuser.DropDownExpanded
        '    If sComboBoxName = "" Then Return

        '    Dim oTech As New PulseTechComputer
        '    oTech.lHullTypeID = Me.lHullTypeID

        '    Dim lMinIdx As Int32 = -1
        '    Select Case sComboBoxName
        '        Case cboCoilMat.ControlName
        '            lMinIdx = 0
        '        Case cboAccelerator.ControlName
        '            lMinIdx = 1
        '        Case cboCasing.ControlName
        '            lMinIdx = 2
        '        Case cboFocuser.ControlName
        '            lMinIdx = 3
        '        Case cboCompressChmbr.ControlName
        '            lMinIdx = 4 
        '    End Select
        '    If lMinIdx = -1 Then Return
        '    oTech.MineralCBOExpanded(lMinIdx, ObjectType.eWeaponTech)

        'End Sub

        'Private Sub ComboValueChanged(ByVal lIndex As Int32) Handles cboAccelerator.ItemSelected, cboCasing.ItemSelected, cboCoilMat.ItemSelected, cboCompressChmbr.ItemSelected, cboFocuser.ItemSelected
        '    BuilderCostValueChange()
        'End Sub
        Private Sub tp_ValueChanged(ByVal sPropName As String, ByRef ctl As ctlTechProp)
            tpROF.MaxValue = Math.Min(Int16.MaxValue, Math.Max(tpROF.MaxValue, tpROF.PropertyValue + 500))
            BuilderCostValueChange(False)
        End Sub

        'Public Overrides Sub BuilderCostValueChange()
        '    If mbIgnoreValueChange = True Then Return
        '    mbIgnoreValueChange = True

        '    If moSB Is Nothing = False Then moSB.Length = 0

        '    'TODO: REMOVE ME WHEN DONE!!!
        '    Dim sPath As String = AppDomain.CurrentDomain.BaseDirectory
        '    If sPath.EndsWith("\") = False Then sPath &= "\"
        '    sPath &= "TECH.dat"
        '    Dim oINI As New InitFile(sPath)
        '    Dim sHeader As String = "PULSE"
        '    Dim ml_POINT_PER_HULL As Int32 = CInt(Val(oINI.GetString(sHeader, "PointPerHull", "600000")))
        '    Dim ml_POINT_PER_POWER As Int32 = CInt(Val(oINI.GetString(sHeader, "PointPerPower", "600000")))
        '    Dim ml_POINT_PER_RES_TIME As Int32 = CInt(Val(oINI.GetString(sHeader, "PointPerResTime", "5")))
        '    Dim ml_POINT_PER_RES_COST As Int32 = CInt(Val(oINI.GetString(sHeader, "PointPerResCost", "1")))
        '    Dim ml_POINT_PER_PROD_TIME As Int32 = CInt(Val(oINI.GetString(sHeader, "PointPerProdTime", "1")))
        '    Dim ml_POINT_PER_PROD_COST As Int32 = CInt(Val(oINI.GetString(sHeader, "PointPerProdCost", "72")))
        '    Dim ml_POINT_PER_COLONIST As Int32 = CInt(Val(oINI.GetString(sHeader, "PointPerColonist", "6")))
        '    Dim ml_POINT_PER_ENLISTED As Int32 = CInt(Val(oINI.GetString(sHeader, "PointPerEnlisted", "60")))
        '    Dim ml_POINT_PER_OFFICER As Int32 = CInt(Val(oINI.GetString(sHeader, "PointPerOfficer", "240")))

        '    Dim ml_POINT_PER_COIL As Int32 = CInt(Val(oINI.GetString(sHeader, "PointPerCoil", "600")))
        '    Dim ml_POINT_PER_ACC As Int32 = CInt(Val(oINI.GetString(sHeader, "PointPerAcc", "600")))
        '    Dim ml_POINT_PER_CASING As Int32 = CInt(Val(oINI.GetString(sHeader, "PointPerCasing", "600")))
        '    Dim ml_POINT_PER_FOCUSER As Int32 = CInt(Val(oINI.GetString(sHeader, "PointPerFocuser", "600")))
        '    Dim ml_POINT_PER_COMPCHMBR As Int32 = CInt(Val(oINI.GetString(sHeader, "PointPerCompChmbr", "600")))

        '    oINI.WriteString(sHeader, "PointPerHull", ml_POINT_PER_HULL.ToString)
        '    oINI.WriteString(sHeader, "PointPerPower", ml_POINT_PER_POWER.ToString)
        '    oINI.WriteString(sHeader, "PointPerResTime", ml_POINT_PER_RES_TIME.ToString)
        '    oINI.WriteString(sHeader, "PointPerResCost", ml_POINT_PER_RES_COST.ToString)
        '    oINI.WriteString(sHeader, "PointPerProdTime", ml_POINT_PER_PROD_TIME.ToString)
        '    oINI.WriteString(sHeader, "PointPerProdCost", ml_POINT_PER_PROD_COST.ToString)
        '    oINI.WriteString(sHeader, "PointPerColonist", ml_POINT_PER_COLONIST.ToString)
        '    oINI.WriteString(sHeader, "PointPerEnlisted", ml_POINT_PER_ENLISTED.ToString)
        '    oINI.WriteString(sHeader, "PointPerOfficer", ml_POINT_PER_OFFICER.ToString)

        '    oINI.WriteString(sHeader, "PointPerCoil", ml_POINT_PER_COIL.ToString)
        '    oINI.WriteString(sHeader, "PointPerAcc", ml_POINT_PER_ACC.ToString)
        '    oINI.WriteString(sHeader, "PointPerCasing", ml_POINT_PER_CASING.ToString)
        '    oINI.WriteString(sHeader, "PointPerFocuser", ml_POINT_PER_FOCUSER.ToString)
        '    oINI.WriteString(sHeader, "PointPerCompChmbr", ml_POINT_PER_COMPCHMBR.ToString)
        '    oINI = Nothing

        '    Try
        '        Dim fROF As Single = Math.Max(1, tpROF.PropertyValue) / 30.0F
        '        Dim lMaxBeam As Int32 = GetMaxBeamDamage()
        '        Dim lMinBeam As Int32 = GetMinBeamDamage()
        '        Dim lMaxImpact As Int32 = GetMaxImpactDamage()
        '        Dim lMinImpact As Int32 = GetMinImpactDamage()

        '        Dim decNormalizer As Decimal = 1D
        '        Dim lMaxGuns As Int32 = 1
        '        Dim lMaxDPS As Int32 = 1
        '        Dim lMaxHullSize As Int32 = 1
        '        Dim lHullAvail As Int32 = 1
        '        Dim lMaxPower As Int32 = 1

        '        If MyBase.lHullTypeID = -1 Then
        '            SetDesignFlaw("Invalid Design: a Hull Type Must be selected!")
        '            mbIgnoreValueChange = False
        '            Return
        '        End If

        '        frmTechBuilder.GetTypeValues(MyBase.lHullTypeID, decNormalizer, lMaxGuns, lMaxDPS, lMaxHullSize, lHullAvail, lMaxPower)

        '        AddToTestString("MaxDPS: " & lMaxDPS.ToString)

        '        Dim decDPS As Decimal = (CDec(lMaxBeam) + CDec(lMaxImpact)) / CDec(fROF)
        '        'Ok, something has changed, let's calculate our values
        '        'This occurs in the following steps/order
        '        ' 1) Determine the required values lookups for the mineral properties
        '        ' 2) Calculate all base values for the costs
        '        '   a) Hull
        '        Dim lBaseHull As Int32 = CalculateBaseHull(decDPS)
        '        AddToTestString("Basehull: " & lBaseHull.ToString)
        '        '   b) Power - we do not use power here
        '        Dim lBasePower As Int32 = CalculateBasePower()
        '        AddToTestString("BasePower: " & lBasePower.ToString)
        '        '   c) Colonists
        '        Dim lBaseColonists As Int32 = CalculateBaseColonists()
        '        AddToTestString("BaseColonists: " & lBaseColonists.ToString)
        '        '   d) Enlisted
        '        Dim lBaseEnlisted As Int32 = CalculateBaseEnlisted()
        '        AddToTestString("BaseEnlisted: " & lBaseEnlisted.ToString)
        '        '   e) Officers
        '        Dim lBaseOfficers As Int32 = CalculateBaseOfficers()
        '        AddToTestString("BaseOfficers: " & lBaseOfficers.ToString)
        '        '   f) Research Cost
        '        Dim blBaseResCost As Int64 = CalculateBaseResCost(decDPS)
        '        AddToTestString("BaseResCost: " & blBaseResCost.ToString)
        '        '   g) Research Time
        '        Dim blBaseResTime As Int64 = CalculateBaseResTime()
        '        AddToTestString("BaseResTime: " & blBaseResTime.ToString)
        '        '   h) Production Cost
        '        Dim blBaseProdCost As Int64 = CalculateBaseProdCost(decDPS)
        '        AddToTestString("BaseProdCost: " & blBaseProdCost.ToString)
        '        '   i) Production Time
        '        Dim blBaseProdTime As Int64 = CalculateBaseProdTime(decDPS)
        '        AddToTestString("BaseProdTime: " & blBaseProdTime.ToString)
        '        '   j) Material Costs
        '        Dim lBaseCoil As Int32 = CalculateBaseMineralCost(cboCoilMat, decDPS, 1, 5, 10, 6, 15, 0, 14, 0)
        '        Dim lBaseAcc As Int32 = CalculateBaseMineralCost(cboAccelerator, decDPS, 21, 4, 10, 6, 14, 6, 15, 0)
        '        Dim lBaseCase As Int32 = CalculateBaseMineralCost(cboCasing, decDPS, 1, 5, 12, 4, 11, 0, 3, 0)
        '        Dim lBaseFocuser As Int32 = CalculateBaseMineralCost(cboFocuser, decDPS, 7, 6, 21, 4, 11, 0, 13, 0)
        '        Dim lBaseCompChmbr As Int32 = CalculateBaseMineralCost(cboCompressChmbr, decDPS, 1, 4, 5, 0, 3, 0, 4, 5)

        '        AddToTestString("BaseCoil: " & lBaseCoil.ToString)
        '        AddToTestString("BaseAcc: " & lBaseAcc.ToString)
        '        AddToTestString("BaseCase: " & lBaseCase.ToString)
        '        AddToTestString("BaseFocuser: " & lBaseFocuser.ToString)
        '        AddToTestString("BaseCompChmbr: " & lBaseCompChmbr.ToString)

        '        ' 3) If the value is locked for the cost, determine the difference between the locked value and the calculated value
        '        Dim lLockedHull As Int32 = -1
        '        Dim lLockedPower As Int32 = -1
        '        Dim blLockedResCost As Int64 = -1
        '        Dim blLockedResTime As Int64 = -1
        '        Dim blLockedProdCost As Int64 = -1
        '        Dim blLockedProdTime As Int64 = -1
        '        Dim lLockedColonists As Int32 = -1
        '        Dim lLockedEnlisted As Int32 = -1
        '        Dim lLockedOfficers As Int32 = -1
        '        MyBase.mfrmBuilderCost.GetLockedValues(lLockedHull, lLockedPower, blLockedResCost, blLockedResTime, blLockedProdCost, blLockedProdTime, lLockedColonists, lLockedEnlisted, lLockedOfficers)

        '        Dim lLockedCoil As Int32 = -1
        '        Dim lLockedAcc As Int32 = -1
        '        Dim lLockedCase As Int32 = -1
        '        Dim lLockedFocuser As Int32 = -1
        '        Dim lLockedCompChmbr As Int32 = -1
        '        If tpCoil.PropertyLocked = True Then lLockedCoil = CInt(tpCoil.PropertyValue)
        '        If tpAccelerator.PropertyLocked = True Then lLockedAcc = CInt(tpAccelerator.PropertyValue)
        '        If tpCasing.PropertyLocked = True Then lLockedCase = CInt(tpCasing.PropertyValue)
        '        If tpFocuser.PropertyLocked = True Then lLockedFocuser = CInt(tpFocuser.PropertyValue)
        '        If tpCompressChmbr.PropertyLocked = True Then lLockedCompChmbr = CInt(tpCompressChmbr.PropertyValue)

        '        Dim bHullLocked As Boolean = lLockedHull <> -1
        '        Dim bPowerLocked As Boolean = lLockedPower <> -1
        '        Dim bResCostLocked As Boolean = blLockedResCost <> -1
        '        Dim bResTimeLocked As Boolean = blLockedResTime <> -1
        '        Dim bProdCostLocked As Boolean = blLockedProdCost <> -1
        '        Dim bProdTimeLocked As Boolean = blLockedProdTime <> -1
        '        Dim bColonistLocked As Boolean = lLockedColonists <> -1
        '        Dim bEnlistedLocked As Boolean = lLockedEnlisted <> -1
        '        Dim bOfficersLocked As Boolean = lLockedOfficers <> -1

        '        Dim blOrigLockedResTime As Int64 = blLockedResTime
        '        Dim blOrigLockedProdTime As Int64 = blLockedProdTime
        '        'If bResTimeLocked = True Then blLockedResTime \= 180000
        '        'If bProdTimeLocked = True Then blLockedProdTime \= 900000
        '        'blBaseResTime \= 180000
        '        'blBaseProdTime \= 900000

        '        Dim lResultHull As Int32 = 0 'lBaseHull
        '        If bHullLocked = True Then lResultHull = lLockedHull Else lResultHull = lBaseHull
        '        Dim lResultPower As Int32 = 0
        '        If bPowerLocked = True Then lResultPower = lLockedPower Else lResultPower = lBasePower
        '        Dim blResultResCost As Int64 = 0
        '        If bResCostLocked = True Then blResultResCost = blLockedResCost Else blResultResCost = blBaseResCost
        '        Dim blResultResTime As Int64 = 0
        '        If bResTimeLocked = True Then blResultResTime = blLockedResTime Else blResultResTime = blBaseResTime
        '        Dim blResultProdCost As Int64 = 0
        '        If bProdCostLocked = True Then blResultProdCost = blLockedProdCost Else blResultProdCost = blBaseProdCost
        '        Dim blResultProdTime As Int64 = 0
        '        If bProdTimeLocked = True Then blResultProdTime = blLockedProdTime Else blResultProdTime = blBaseProdTime
        '        Dim lResultColonists As Int32 = 0
        '        If bColonistLocked = True Then lResultColonists = lLockedColonists Else lResultColonists = lBaseColonists
        '        Dim lResultEnlisted As Int32 = 0
        '        If bEnlistedLocked = True Then lResultEnlisted = lLockedEnlisted Else lResultEnlisted = lBaseEnlisted
        '        Dim lResultOfficers As Int32 = 0
        '        If bOfficersLocked = True Then lResultOfficers = lLockedOfficers Else lResultOfficers = lBaseOfficers
        '        Dim lResultCoil As Int32 = 0
        '        If tpCoil.PropertyLocked = True Then lResultCoil = lLockedCoil Else lResultCoil = lBaseCoil
        '        Dim lResultAcc As Int32 = 0
        '        If tpAccelerator.PropertyLocked = True Then lResultAcc = lLockedAcc Else lResultAcc = lBaseAcc
        '        Dim lResultCasing As Int32 = 0
        '        If tpCasing.PropertyLocked = True Then lResultCasing = lLockedCase Else lResultCasing = lBaseCase
        '        Dim lResultFocuser As Int32 = 0
        '        If tpFocuser.PropertyLocked = True Then lResultFocuser = lLockedFocuser Else lResultFocuser = lBaseFocuser
        '        Dim lResultCompChmbr As Int32 = 0
        '        If tpCompressChmbr.PropertyLocked = True Then lResultCompChmbr = lLockedCompChmbr Else lResultCompChmbr = lBaseCompChmbr

        '        Dim lblDesignFlaw As UILabel = CType(MyBase.ParentControl, frmWeaponBuilder).lblDesignFlaw
        '        If bHullLocked = True AndAlso InvalidHalfBaseCheck(lBaseHull, lLockedHull, "Hull", False, lblDesignFlaw, mbIgnoreValueChange) = True Then Return
        '        If bResCostLocked = True AndAlso InvalidHalfBaseCheck(blBaseResCost, blLockedResCost, "research cost", False, lblDesignFlaw, mbIgnoreValueChange) = True Then Return
        '        If bProdCostLocked = True AndAlso InvalidHalfBaseCheck(blBaseProdCost, blLockedProdCost, "production cost", False, lblDesignFlaw, mbIgnoreValueChange) = True Then Return
        '        If bColonistLocked = True AndAlso InvalidHalfBaseCheck(lBaseColonists, lLockedColonists, "Colonists", True, lblDesignFlaw, mbIgnoreValueChange) = True Then Return
        '        If bEnlistedLocked = True AndAlso InvalidHalfBaseCheck(lBaseEnlisted, lLockedEnlisted, "Enlisted", True, lblDesignFlaw, mbIgnoreValueChange) = True Then Return
        '        If bOfficersLocked = True AndAlso InvalidHalfBaseCheck(lBaseOfficers, lLockedOfficers, "Officers", True, lblDesignFlaw, mbIgnoreValueChange) = True Then Return
        '        If tpCoil.PropertyLocked = True AndAlso InvalidHalfBaseCheck(lBaseCoil, lLockedCoil, "Coil usage", False, lblDesignFlaw, mbIgnoreValueChange) = True Then Return
        '        If tpAccelerator.PropertyLocked = True AndAlso InvalidHalfBaseCheck(lBaseAcc, lLockedAcc, "Accelerator usage", False, lblDesignFlaw, mbIgnoreValueChange) = True Then Return
        '        If tpCasing.PropertyLocked = True AndAlso InvalidHalfBaseCheck(lBaseCase, lLockedCase, "Casing usage", False, lblDesignFlaw, mbIgnoreValueChange) = True Then Return
        '        If tpFocuser.PropertyLocked = True AndAlso InvalidHalfBaseCheck(lBaseFocuser, lLockedFocuser, "Focuser usage", False, lblDesignFlaw, mbIgnoreValueChange) = True Then Return
        '        If tpCompressChmbr.PropertyLocked = True AndAlso InvalidHalfBaseCheck(lBaseCompChmbr, lLockedCompChmbr, "Compress Chamber usage", False, lblDesignFlaw, mbIgnoreValueChange) = True Then Return
        '        If (bResTimeLocked = True AndAlso blLockedResTime < 1) OrElse (bProdTimeLocked = True AndAlso blLockedProdTime < 1) Then
        '            lblDesignFlaw.Caption = "Your scientists cannot get the time that low."
        '            lblDesignFlaw.Visible = True
        '            mbIgnoreValueChange = False
        '            Return
        '        End If

        '        Dim lHullDiff As Int32 = DoLockedToBaseCheck(lLockedHull, lBaseHull)
        '        Dim lPowerDiff As Int32 = DoLockedToBaseCheck(lLockedPower, lBasePower)
        '        Dim blResCostDiff As Int64 = DoLockedToBaseCheck(blLockedResCost, blBaseResCost)
        '        Dim blResTimeDiff As Int64 = DoLockedToBaseCheck(blLockedResTime, blBaseResTime)
        '        Dim blProdCostDiff As Int64 = DoLockedToBaseCheck(blLockedProdCost, blBaseProdCost)
        '        Dim blProdTimeDiff As Int64 = DoLockedToBaseCheck(blLockedProdTime, blBaseProdTime)
        '        Dim lColonistDiff As Int32 = DoLockedToBaseCheck(lLockedColonists, lBaseColonists)
        '        Dim lEnlistedDiff As Int32 = DoLockedToBaseCheck(lLockedEnlisted, lBaseEnlisted)
        '        Dim lOfficerDiff As Int32 = DoLockedToBaseCheck(lLockedOfficers, lBaseOfficers)
        '        Dim lCoilDiff As Int32 = DoLockedToBaseCheck(lLockedCoil, lBaseCoil)
        '        Dim lAccDiff As Int32 = DoLockedToBaseCheck(lLockedAcc, lBaseAcc)
        '        Dim lCaseDiff As Int32 = DoLockedToBaseCheck(lLockedCase, lBaseCase)
        '        Dim lFocuserDiff As Int32 = DoLockedToBaseCheck(lLockedFocuser, lBaseFocuser)
        '        Dim lCompChmbrDiff As Int32 = DoLockedToBaseCheck(lLockedCompChmbr, lBaseCompChmbr)

        '        ' 4) Determine the point costs for each item
        '        'Const ml_POINT_PER_HULL As Int32 = 1000
        '        'Const ml_POINT_PER_POWER As Int32 = 1000
        '        'Const ml_POINT_PER_RES_TIME As Int32 = 100
        '        'Const ml_POINT_PER_RES_COST As Int32 = 230
        '        'Const ml_POINT_PER_PROD_TIME As Int32 = 300
        '        'Const ml_POINT_PER_PROD_COST As Int32 = 180
        '        'Const ml_POINT_PER_COLONIST As Int32 = 1
        '        'Const ml_POINT_PER_ENLISTED As Int32 = 10
        '        'Const ml_POINT_PER_OFFICER As Int32 = 40

        '        'Const ml_POINT_PER_COIL As Int32 = 70
        '        'Const ml_POINT_PER_ACC As Int32 = 120
        '        'Const ml_POINT_PER_CASING As Int32 = 100
        '        'Const ml_POINT_PER_FOCUSER As Int32 = 110
        '        'Const ml_POINT_PER_COMPCHMBR As Int32 = 90

        '        Dim blPoints As Int64 = 0
        '        blPoints += CLng(lHullDiff) * ml_POINT_PER_HULL
        '        blPoints += CLng(lPowerDiff) * ml_POINT_PER_POWER
        '        blPoints += blResTimeDiff * ml_POINT_PER_RES_TIME
        '        blPoints += blResCostDiff * ml_POINT_PER_RES_COST
        '        blPoints += blProdTimeDiff * ml_POINT_PER_PROD_TIME
        '        blPoints += blProdCostDiff * ml_POINT_PER_PROD_COST
        '        blPoints += lColonistDiff * ml_POINT_PER_COLONIST
        '        blPoints += lEnlistedDiff * ml_POINT_PER_ENLISTED
        '        blPoints += lOfficerDiff * ml_POINT_PER_OFFICER
        '        blPoints += lCoilDiff * ml_POINT_PER_COIL
        '        blPoints += lAccDiff * ml_POINT_PER_ACC
        '        blPoints += lCaseDiff * ml_POINT_PER_CASING
        '        blPoints += lFocuserDiff * ml_POINT_PER_FOCUSER
        '        blPoints += lCompChmbrDiff * ml_POINT_PER_COMPCHMBR

        '        AddToTestString("BasePoints: " & blPoints.ToString)
        '        AddToTestString("DPS: " & decDPS.ToString)
        '        If decDPS > lMaxDPS Then
        '            Dim decAdd As Decimal = decDPS - CDec(lMaxDPS)
        '            decAdd *= 1000000D
        '            Dim blAllConstants As Int64 = ml_POINT_PER_HULL
        '            blAllConstants += ml_POINT_PER_POWER
        '            blAllConstants += ml_POINT_PER_RES_TIME
        '            blAllConstants += ml_POINT_PER_RES_COST
        '            blAllConstants += ml_POINT_PER_PROD_TIME
        '            blAllConstants += ml_POINT_PER_PROD_COST
        '            blAllConstants += ml_POINT_PER_COLONIST
        '            blAllConstants += ml_POINT_PER_ENLISTED
        '            blAllConstants += ml_POINT_PER_OFFICER
        '            blAllConstants += ml_POINT_PER_ACC
        '            blAllConstants += ml_POINT_PER_CASING
        '            blAllConstants += ml_POINT_PER_COIL
        '            blAllConstants += ml_POINT_PER_COMPCHMBR
        '            blAllConstants += ml_POINT_PER_FOCUSER
        '            decAdd *= CDec(blAllConstants)
        '            If CDec(blPoints + decAdd) > Int64.MaxValue Then blPoints = Int64.MaxValue Else blPoints = CLng(blPoints + decAdd)
        '        End If

        '        AddToTestString("Points: " & blPoints.ToString)

        '        ' 5) Take the sum of the remaining points and divide it by 1 point of each type not locked summed together rounded down
        '        Dim blPointSum As Int64 = 0

        '        Dim lModByVal As Int32 = 0

        '        If blPoints < 0 Then
        '            If bHullLocked = False AndAlso lResultHull > 0 Then blPointSum += ml_POINT_PER_HULL
        '            If bPowerLocked = False AndAlso lResultPower > 0 Then blPointSum += ml_POINT_PER_POWER
        '            If bResCostLocked = False AndAlso blResultResCost > 0 Then blPointSum += ml_POINT_PER_RES_COST
        '            If bResTimeLocked = False AndAlso blResultResTime > 0 Then blPointSum += ml_POINT_PER_RES_TIME
        '            If bProdCostLocked = False AndAlso blResultProdCost > 0 Then blPointSum += ml_POINT_PER_PROD_COST
        '            If bProdTimeLocked = False AndAlso blResultProdTime > 0 Then blPointSum += ml_POINT_PER_PROD_TIME
        '            If bColonistLocked = False AndAlso lResultColonists > 0 Then blPointSum += ml_POINT_PER_COLONIST
        '            If bEnlistedLocked = False AndAlso lResultEnlisted > 0 Then blPointSum += ml_POINT_PER_ENLISTED
        '            If bOfficersLocked = False AndAlso lResultOfficers > 0 Then blPointSum += ml_POINT_PER_OFFICER
        '            If tpCoil.PropertyLocked = False AndAlso lResultCoil > 0 Then blPointSum += ml_POINT_PER_COIL
        '            If tpAccelerator.PropertyLocked = False AndAlso lResultAcc > 0 Then blPointSum += ml_POINT_PER_ACC
        '            If tpCasing.PropertyLocked = False AndAlso lResultCasing > 0 Then blPointSum += ml_POINT_PER_CASING
        '            If tpFocuser.PropertyLocked = False AndAlso lResultFocuser > 0 Then blPointSum += ml_POINT_PER_FOCUSER
        '            If tpCompressChmbr.PropertyLocked = False AndAlso lResultCompChmbr > 0 Then blPointSum += ml_POINT_PER_COMPCHMBR

        '            If blPointSum > 0 Then
        '                AddToTestString("PointSum: " & blPointSum.ToString)

        '                ' 6) The result is applied to all non-locked values
        '                Dim blTemp As Int64 = Math.Abs(blPoints) \ blPointSum
        '                Dim lAddPts As Int32 = 0
        '                If blTemp > Int32.MaxValue Then blTemp = Int32.MaxValue
        '                If blTemp < 0 Then blTemp = 0
        '                lAddPts = CInt(blTemp)
        '                If lAddPts > 0 Then
        '                    If bHullLocked = False AndAlso lResultHull * ml_POINT_PER_HULL > lAddPts Then lResultHull -= lAddPts
        '                    If bPowerLocked = False AndAlso lResultPower * ml_POINT_PER_POWER > lAddPts Then lResultPower -= lAddPts
        '                    If bResCostLocked = False AndAlso blResultResCost * ml_POINT_PER_RES_COST > lAddPts Then blResultResCost -= lAddPts
        '                    If bResTimeLocked = False AndAlso blResultResTime * ml_POINT_PER_RES_TIME > lAddPts Then blResultResTime -= lAddPts
        '                    If bProdCostLocked = False AndAlso blResultProdCost * ml_POINT_PER_PROD_COST > lAddPts Then blResultProdCost -= lAddPts
        '                    If bProdTimeLocked = False AndAlso blResultProdTime * ml_POINT_PER_PROD_TIME > lAddPts Then blResultProdTime -= lAddPts
        '                    If bColonistLocked = False AndAlso lResultColonists * ml_POINT_PER_COLONIST > lAddPts Then lResultColonists -= lAddPts
        '                    If bEnlistedLocked = False AndAlso lResultEnlisted * ml_POINT_PER_ENLISTED > lAddPts Then lResultEnlisted -= lAddPts
        '                    If bOfficersLocked = False AndAlso lResultOfficers * ml_POINT_PER_OFFICER > lAddPts Then lResultOfficers -= lAddPts

        '                    If tpCoil.PropertyLocked = False AndAlso lResultCoil * ml_POINT_PER_COIL > lAddPts Then lResultCoil -= lAddPts
        '                    If tpAccelerator.PropertyLocked = False AndAlso lResultAcc * ml_POINT_PER_ACC > lAddPts Then lResultAcc -= lAddPts
        '                    If tpCasing.PropertyLocked = False AndAlso lResultCasing * ml_POINT_PER_CASING > lAddPts Then lResultCasing -= lAddPts
        '                    If tpFocuser.PropertyLocked = False AndAlso lResultFocuser * ml_POINT_PER_FOCUSER > lAddPts Then lResultFocuser -= lAddPts
        '                    If tpCompressChmbr.PropertyLocked = False AndAlso lResultCompChmbr * ml_POINT_PER_COMPCHMBR > lAddPts Then lResultCompChmbr -= lAddPts

        '                    blPoints += CLng(lAddPts) * blPointSum
        '                End If
        '            End If

        '            lModByVal = -1
        '        ElseIf blPoints > 0 Then
        '            If bHullLocked = False Then blPointSum += ml_POINT_PER_HULL
        '            If bResCostLocked = False Then blPointSum += ml_POINT_PER_RES_COST
        '            If bResTimeLocked = False Then blPointSum += ml_POINT_PER_RES_TIME
        '            If bProdCostLocked = False Then blPointSum += ml_POINT_PER_PROD_COST
        '            If bProdTimeLocked = False Then blPointSum += ml_POINT_PER_PROD_TIME
        '            'If bColonistLocked = False Then blPointSum += ml_POINT_PER_COLONIST
        '            'If bEnlistedLocked = False Then blPointSum += ml_POINT_PER_ENLISTED
        '            'If bOfficersLocked = False Then blPointSum += ml_POINT_PER_OFFICER
        '            If tpCoil.PropertyLocked = False Then blPointSum += ml_POINT_PER_COIL
        '            If tpAccelerator.PropertyLocked = False Then blPointSum += ml_POINT_PER_ACC
        '            If tpCasing.PropertyLocked = False Then blPointSum += ml_POINT_PER_CASING
        '            If tpFocuser.PropertyLocked = False Then blPointSum += ml_POINT_PER_FOCUSER
        '            If tpCompressChmbr.PropertyLocked = False Then blPointSum += ml_POINT_PER_COMPCHMBR

        '            If blPointSum > 0 Then
        '                AddToTestString("PointSum: " & blPointSum.ToString)

        '                ' 6) The result is applied to all non-locked values
        '                Dim blTemp As Int64 = blPoints \ blPointSum
        '                Dim lAddPts As Int32 = 0
        '                'If blTemp > Int32.MaxValue Then lAddPts = Int32.MaxValue
        '                'If blTemp < 0 Then lAddPts = 0
        '                If blTemp > 0 Then
        '                    If bHullLocked = False Then
        '                        Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_HULL
        '                        If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
        '                        lResultHull += lAddPts
        '                    End If
        '                    If bPowerLocked = False Then
        '                        Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_POWER
        '                        If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
        '                        lResultPower += lAddPts
        '                    End If
        '                    If bResCostLocked = False Then
        '                        blResultResCost += blTemp '\ ml_POINT_PER_RES_COST
        '                    End If
        '                    If bResTimeLocked = False Then blResultResTime += blTemp '\ ml_POINT_PER_RES_TIME
        '                    If bProdCostLocked = False Then blResultProdCost += blTemp '\ ml_POINT_PER_PROD_COST
        '                    If bProdTimeLocked = False Then blResultProdTime += blTemp '\ ml_POINT_PER_PROD_TIME

        '                    If tpCoil.PropertyLocked = False Then
        '                        Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_COIL
        '                        If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
        '                        lResultCoil += lAddPts
        '                    End If
        '                    If tpAccelerator.PropertyLocked = False Then
        '                        Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_ACC
        '                        If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
        '                        lResultAcc += lAddPts
        '                    End If
        '                    If tpCasing.PropertyLocked = False Then
        '                        Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_CASING
        '                        If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
        '                        lResultCasing += lAddPts
        '                    End If
        '                    If tpFocuser.PropertyLocked = False Then
        '                        Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_FOCUSER
        '                        If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
        '                        lResultFocuser += lAddPts
        '                    End If
        '                    If tpCompressChmbr.PropertyLocked = False Then
        '                        Dim blAdd As Int64 = blTemp '\ ml_POINT_PER_COMPCHMBR
        '                        If blAdd > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blAdd)
        '                        lResultCompChmbr += lAddPts
        '                    End If

        '                    blPoints -= CLng(blTemp) * blPointSum
        '                End If
        '            End If

        '            lModByVal = 1
        '        End If

        '        If blPoints > 0 Then
        '            blPointSum = 0
        '            If bColonistLocked = False Then blPointSum += ml_POINT_PER_COLONIST
        '            If bEnlistedLocked = False Then blPointSum += ml_POINT_PER_ENLISTED
        '            If bOfficersLocked = False Then blPointSum += ml_POINT_PER_OFFICER

        '            If blPointSum > 0 Then
        '                Dim blTemp As Int64 = blPoints \ blPointSum
        '                Dim lAddPts As Int32 = 0
        '                If blTemp > Int32.MaxValue Then lAddPts = Int32.MaxValue Else lAddPts = CInt(blTemp)
        '                If bColonistLocked = False Then lResultColonists += lAddPts
        '                If bEnlistedLocked = False Then lResultEnlisted += lAddPts
        '                If bOfficersLocked = False Then lResultOfficers += lAddPts

        '                blPoints -= CLng(blTemp) * blPointSum
        '            End If
        '        End If

        '        ' 7) The remainder points are applied to the non-locked items until the remainder < 0
        '        Dim lCnt As Int32 = 0
        '        While blPoints > 0
        '            Dim blPrevPoints As Int64 = blPoints
        '            If bHullLocked = False Then AffectValueByPoints(blPoints, ml_POINT_PER_HULL, lResultHull, lModByVal)
        '            If bPowerLocked = False Then AffectValueByPoints(blPoints, ml_POINT_PER_POWER, lResultPower, lModByVal)
        '            If bResCostLocked = False Then AffectValueByPoints(blPoints, ml_POINT_PER_RES_COST, blResultResCost, lModByVal)
        '            If bResTimeLocked = False Then AffectValueByPoints(blPoints, ml_POINT_PER_RES_TIME, blResultResTime, lModByVal)
        '            If bProdCostLocked = False Then AffectValueByPoints(blPoints, ml_POINT_PER_PROD_COST, blResultProdCost, lModByVal)
        '            If bProdTimeLocked = False Then AffectValueByPoints(blPoints, ml_POINT_PER_PROD_TIME, blResultProdTime, lModByVal)
        '            If bColonistLocked = False Then AffectValueByPoints(blPoints, ml_POINT_PER_COLONIST, lResultColonists, lModByVal)
        '            If bEnlistedLocked = False Then AffectValueByPoints(blPoints, ml_POINT_PER_ENLISTED, lResultEnlisted, lModByVal)
        '            If bOfficersLocked = False Then AffectValueByPoints(blPoints, ml_POINT_PER_OFFICER, lResultOfficers, lModByVal)
        '            If tpCoil.PropertyLocked = False Then AffectValueByPoints(blPoints, ml_POINT_PER_COIL, lResultCoil, lModByVal)
        '            If tpAccelerator.PropertyLocked = False Then AffectValueByPoints(blPoints, ml_POINT_PER_ACC, lResultAcc, lModByVal)
        '            If tpCasing.PropertyLocked = False Then AffectValueByPoints(blPoints, ml_POINT_PER_CASING, lResultCasing, lModByVal)
        '            If tpFocuser.PropertyLocked = False Then AffectValueByPoints(blPoints, ml_POINT_PER_FOCUSER, lResultFocuser, lModByVal)
        '            If tpCompressChmbr.PropertyLocked = False Then AffectValueByPoints(blPoints, ml_POINT_PER_COMPCHMBR, lResultCompChmbr, lModByVal)
        '            ' 8) If the remainder does not reach 0 after several attempts, break out
        '            If blPrevPoints = blPoints Then Exit While
        '            lCnt += 1
        '            If lCnt > 100 Then Exit While
        '        End While

        '        ' 9) finally, set our prod cost items to their proper values
        '        'If bHullLocked = False Then lLockedHull = lBaseHull
        '        'If bPowerLocked = False Then lLockedPower = lBasePower
        '        'If bResCostLocked = False Then blLockedResCost = blBaseResCost
        '        'If bResTimeLocked = False Then blLockedResTime = blBaseResTime
        '        'If bProdCostLocked = False Then blLockedProdCost = blBaseProdCost
        '        'If bProdTimeLocked = False Then blLockedProdTime = blBaseProdTime
        '        'If bColonistLocked = False Then lLockedColonists = lBaseColonists
        '        'If bEnlistedLocked = False Then lLockedEnlisted = lBaseEnlisted
        '        'If bOfficersLocked = False Then lLockedOfficers = lBaseOfficers

        '        'If bResTimeLocked = False Then
        '        '    blLockedResTime *= 180000L
        '        'Else : blLockedResTime = blOrigLockedResTime
        '        'End If
        '        'If bProdTimeLocked = False Then
        '        '    blLockedProdTime *= 900000L
        '        'Else : blLockedProdTime = blOrigLockedProdTime
        '        'End If

        '        MyBase.mfrmBuilderCost.SetBaseValues(lResultHull, lResultPower, blResultResCost, blResultResTime, blResultProdCost, blResultProdTime, lResultColonists, lResultEnlisted, lResultOfficers)
        '        If tpCoil.PropertyLocked = False Then
        '            tpCoil.MaxValue = Math.Max(tpCoil.MaxValue, lResultCoil)
        '            tpCoil.PropertyValue = lResultCoil
        '        End If
        '        If tpAccelerator.PropertyLocked = False Then
        '            tpAccelerator.MaxValue = Math.Max(tpAccelerator.MaxValue, lResultAcc)
        '            tpAccelerator.PropertyValue = lResultAcc
        '        End If
        '        If tpCasing.PropertyLocked = False Then
        '            tpCasing.MaxValue = Math.Max(tpCasing.MaxValue, lResultCasing)
        '            tpCasing.PropertyValue = lResultCasing
        '        End If
        '        If tpFocuser.PropertyLocked = False Then
        '            tpFocuser.MaxValue = Math.Max(tpFocuser.MaxValue, lResultFocuser)
        '            tpFocuser.PropertyValue = lResultFocuser
        '        End If
        '        If tpCompressChmbr.PropertyLocked = False Then
        '            tpCompressChmbr.MaxValue = Math.Max(tpCompressChmbr.MaxValue, lResultCompChmbr)
        '            tpCompressChmbr.PropertyValue = lResultCompChmbr
        '        End If

        '        AddToTestString(" Remaining Points: " & blPoints.ToString)
        '        AddToTestString(" Hull: " & lResultHull.ToString)
        '        AddToTestString(" ResCost: " & blResultResCost.ToString)
        '        AddToTestString(" ResTime: " & blResultResTime.ToString)
        '        AddToTestString(" ProdCost: " & blResultProdCost.ToString)
        '        AddToTestString(" ProdTime: " & blResultProdTime.ToString)
        '        AddToTestString(" Colonist: " & lResultColonists.ToString)
        '        AddToTestString(" Enlisted: " & lResultEnlisted.ToString)
        '        AddToTestString(" Officer: " & lResultOfficers.ToString)
        '        AddToTestString(" Acc: " & lResultAcc.ToString)
        '        AddToTestString(" Casing: " & lResultCasing.ToString)
        '        AddToTestString(" Coil: " & lResultCoil.ToString)
        '        AddToTestString(" CompChmbr: " & lResultCompChmbr.ToString)
        '        AddToTestString(" Focuser: " & lResultFocuser.ToString)

        '        If blPoints > 1000 Then
        '            SetDesignFlaw("This design is impossible to comprehend.")
        '        Else
        '            SetDesignFlaw("Your scientists believe this design to be possible.")
        '        End If

        '        If Me.ParentControl Is Nothing = False Then
        '            CType(Me.ParentControl, frmWeaponBuilder).SetWeaponDetailedStats(0, 0, lMinBeam, lMaxBeam, 0, 0, 0, 0, lMinImpact, lMaxImpact, 0, 0, fROF)
        '        End If
        '    Catch
        '        SetDesignFlaw("This design is impossible to comprehend.")
        '    End Try

        '    Dim oResFrm As frmTechBuilderDisplay = CType(MyBase.moUILib.GetWindow("frmTechBuilderDisplay"), frmTechBuilderDisplay)
        '    If oResFrm Is Nothing Then oResFrm = New frmTechBuilderDisplay(MyBase.moUILib)
        '    oResFrm.SetValues(moSB.ToString)
        '    moSB.Length = 0

        '    mbIgnoreValueChange = False

        'End Sub
        Public Overrides Sub BuilderCostValueChange(ByVal bAutoFill As Boolean)
            If mbIgnoreValueChange = True Then Return
            mbIgnoreValueChange = True
            mbImpossibleDesign = True

            Dim oPulseTechComputer As New PulseTechComputer
            With oPulseTechComputer
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

                tpInputEnergy.MaxValue = Math.Min(mlMaxInputEnergy, lMaxPower)
                If tpInputEnergy.PropertyValue > tpInputEnergy.MaxValue Then tpInputEnergy.PropertyValue = tpInputEnergy.MaxValue
                tpInputEnergy.ToolTipText = "Indicates the power injected into the system." & vbCrLf & "This directly indicates the power requirement for the system." & vbCrLf & "Maximum Value for this is " & tpInputEnergy.MaxValue & "."
                tpCompress.MaxValue = Math.Min(mlMaxCompress * 10L, (lMaxHullSize \ 2) * 10)
                If tpCompress.PropertyValue > tpCompress.MaxValue Then tpCompress.PropertyValue = tpCompress.MaxValue
                tpCompress.ToolTipText = "The multiplier of power from the compression chamber." & vbCrLf & "Higher values here require purer materials for the Compression Chamber." & vbCrLf & "Value can be a decimal number but must be positive" & vbCrLf & "and cannot exceed " & tpCompress.MaxValue \ 10 & "."

                Dim fROF As Single = Math.Max(1, tpROF.PropertyValue) / 30.0F
                Dim lMaxBeam As Int32 = GetMaxBeamDamage()
                Dim lMinBeam As Int32 = GetMinBeamDamage()
                Dim lMaxImpact As Int32 = GetMaxImpactDamage()
                Dim lMinImpact As Int32 = GetMinImpactDamage()

                With CType(Me.ParentControl, frmWeaponBuilder)
                    oPulseTechComputer.SetMinDAValues(.ml_Min1DA, .ml_Min2DA, .ml_Min3DA, .ml_Min4DA, .ml_Min5DA, .ml_Min6DA)
                End With

                .decDPS = (CDec(lMaxBeam) + CDec(lMaxImpact)) / CDec(fROF)
                .decDPS = Math.Max(.decDPS, (.decDPS + (CDec(lMaxBeam) + CDec(lMaxImpact))) / 2D)
                'Dim decNominal As Decimal = (CDec(lMaxPower / lMaxDPS) * 3D) * .decDPS

                .blCompress = tpCompress.PropertyValue
                .blInputEnergy = tpInputEnergy.PropertyValue
                .blMaxRange = tpMaxRange.PropertyValue
                .blROF = tpROF.PropertyValue
                .blScatterRadius = tpScatterRadius.PropertyValue
                .lHullTypeID = Me.lHullTypeID

                'If cboCoilMat.ListIndex > -1 Then
                '    .lMineral1ID = cboCoilMat.ItemData(cboCoilMat.ListIndex)
                'Else : .lMineral1ID = -1
                'End If
                'If cboAccelerator.ListIndex > -1 Then
                '    .lMineral2ID = cboAccelerator.ItemData(cboAccelerator.ListIndex)
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
                'If cboCompressChmbr.ListIndex > -1 Then
                '    .lMineral5ID = cboCompressChmbr.ItemData(cboCompressChmbr.ListIndex)
                'Else : .lMineral5ID = -1
                'End If
                .lMineral1ID = mlMineralIDs(0)
                .lMineral2ID = mlMineralIDs(1)
                .lMineral3ID = mlMineralIDs(2)
                .lMineral4ID = mlMineralIDs(3)
                .lMineral5ID = mlMineralIDs(4)
                .lMineral6ID = -1

                .msMin1Name = "Coil"
                .msMin2Name = "Accelerator"
                .msMin3Name = "Casing"
                .msMin4Name = "Focuser"
                .msMin5Name = "Compression Chamber"
                .msMin6Name = ""

                Dim lblDesignFlaw As UILabel = CType(MyBase.ParentControl, frmWeaponBuilder).lblDesignFlaw

                If .lMineral1ID < 1 OrElse .lMineral2ID < 1 OrElse .lMineral3ID < 1 OrElse .lMineral4ID < 1 OrElse .lMineral5ID < 1 Then
                    lblDesignFlaw.Caption = "All properties and materials need to be defined."
                    mbIgnoreValueChange = False
                    mbImpossibleDesign = True
                    Return
                End If

                If bAutoFill = True Then
                    .DoBuilderCostPreconfigure(lblDesignFlaw, MyBase.mfrmBuilderCost, tpCoil, tpAccelerator, tpCasing, tpFocuser, tpCompressChmbr, Nothing, ObjectType.eWeaponTech, 0D, False)
                Else
                    .BuilderCostValueChange(lblDesignFlaw, MyBase.mfrmBuilderCost, tpCoil, tpAccelerator, tpCasing, tpFocuser, tpCompressChmbr, Nothing, ObjectType.eWeaponTech, 0D)
                End If

                mbImpossibleDesign = .bImpossibleDesign
                If Me.ParentControl Is Nothing = False Then
                    CType(Me.ParentControl, frmWeaponBuilder).SetWeaponDetailedStats(0, 0, lMinBeam, lMaxBeam, 0, 0, 0, 0, lMinImpact, lMaxImpact, 0, 0, fROF)
                End If

            End With

            mbIgnoreValueChange = False
        End Sub
        '#Region "  Base Calculation Methods  "
        '        Private Function CalculateBaseHull(ByVal mdecDPS As Decimal) As Int32
        '            Dim decTemp As Decimal = tpCompress.PropertyValue + (mdecDPS * 4D)
        '            If decTemp > Int32.MaxValue Then Return Int32.MaxValue
        '            If decTemp < 0 Then Return 0
        '            Return CInt(decTemp)
        '        End Function
        '        Private Function CalculateBasePower() As Int32
        '            Dim decTemp As Decimal = (tpInputEnergy.PropertyValue * 10) + (tpScatterRadius.PropertyValue * 20)
        '            If decTemp > Int32.MaxValue Then Return Int32.MaxValue
        '            If decTemp < 0 Then Return 0
        '            Return CInt(decTemp)
        '        End Function
        '        Private Function CalculateBaseResTime() As Int64
        '            Dim decTemp As Decimal = tpInputEnergy.PropertyValue * 2000000D
        '            decTemp += tpCompress.PropertyValue * 1000000D
        '            Const fMinROF As Single = 1 / 30.0F
        '            Dim fROF As Single = Math.Min(fMinROF, tpROF.PropertyValue / 30.0F)
        '            decTemp += CDec(1.0F / fROF) * 10000000D
        '            decTemp += tpMaxRange.PropertyValue * 700000D
        '            decTemp += tpScatterRadius.PropertyValue * 10000000
        '            If decTemp > Int64.MaxValue Then Return Int64.MaxValue
        '            If decTemp < 0 Then Return 0
        '            Return CLng(decTemp)
        '        End Function
        '        Private Function CalculateBaseResCost(ByVal decDPS As Decimal) As Int64
        '            Dim decTemp As Decimal = decDPS * 540000D
        '            decTemp += tpMaxRange.PropertyValue * 31000D
        '            decTemp += tpScatterRadius.PropertyValue * 161000D
        '            If decTemp > Int64.MaxValue Then Return Int64.MaxValue
        '            If decTemp < 0 Then Return 0
        '            Return CLng(decTemp)
        '        End Function
        '        Private Function CalculateBaseProdTime(ByVal decDPS As Decimal) As Int64
        '            Dim decTemp As Decimal = decDPS * 1400000D
        '            decTemp += tpScatterRadius.PropertyValue * 500000D
        '            If decTemp > Int64.MaxValue Then Return Int64.MaxValue
        '            If decTemp < 0 Then Return 0
        '            Return CLng(decTemp)
        '        End Function
        '        Private Function CalculateBaseProdCost(ByVal decDPS As Decimal) As Int64
        '            Dim decTemp As Decimal = decDPS * 300000D
        '            decTemp += tpMaxRange.PropertyValue * 21000D
        '            decTemp += tpScatterRadius.PropertyValue * 121000D
        '            If decTemp > Int64.MaxValue Then Return Int64.MaxValue
        '            If decTemp < 0 Then Return 0
        '            Return CLng(decTemp)
        '        End Function
        '        Private Function CalculateBaseColonists() As Int32
        '            Dim decTemp As Decimal = tpMaxRange.PropertyValue / 2D
        '            decTemp += tpScatterRadius.PropertyValue
        '            If decTemp < 0 Then Return 0
        '            If decTemp > Int32.MaxValue Then Return Int32.MaxValue
        '            Return CInt(decTemp)
        '        End Function
        '        Private Function CalculateBaseEnlisted() As Int32
        '            Dim decTemp As Decimal = (tpInputEnergy.PropertyValue / 1000D) + 1D
        '            Const fMinROF As Single = 1 / 30.0F
        '            Dim fROF As Single = Math.Min(fMinROF, tpROF.PropertyValue / 30.0F)
        '            decTemp += CDec(10D / fROF)
        '            If decTemp > Int32.MaxValue Then Return Int32.MaxValue
        '            If decTemp < 0 Then Return 0
        '            Return CInt(decTemp)
        '        End Function
        '        Private Function CalculateBaseOfficers() As Int32
        '            Dim decTemp As Decimal = tpCompress.PropertyValue / 10D
        '            decTemp *= 1.5D
        '            If decTemp > Int32.MaxValue Then Return Int32.MaxValue
        '            If decTemp < 0 Then Return 0
        '            Return CInt(decTemp)
        '        End Function
        '        Private Function CalculateBaseMineralCost(ByRef cbo As UIComboBox, ByVal decDPS As Decimal, ByVal lProp1ID As Int32, ByVal lProp1Desired As Int32, ByVal lProp2ID As Int32, ByVal lProp2Desired As Int32, ByVal lProp3ID As Int32, ByVal lProp3Desired As Int32, ByVal lProp4ID As Int32, ByVal lProp4Desired As Int32) As Int32
        '            Dim oMineral As Mineral = Nothing
        '            If cbo.ListIndex > -1 Then
        '                Dim lID As Int32 = cbo.ItemData(cbo.ListIndex)
        '                For X As Int32 = 0 To glMineralUB
        '                    If glMineralIdx(X) = lID Then
        '                        oMineral = goMinerals(X)
        '                        Exit For
        '                    End If
        '                Next X
        '            End If
        '            If oMineral Is Nothing Then Return 0

        '            Dim lTemp As Int32 = oMineral.GetPropertyIndex(lProp1ID)
        '            Dim lProp1Actual As Int32 = oMineral.MinPropValueScore(lTemp)
        '            lTemp = oMineral.GetPropertyIndex(lProp2ID)
        '            Dim lProp2Actual As Int32 = oMineral.MinPropValueScore(lTemp)
        '            lTemp = oMineral.GetPropertyIndex(lProp3ID)
        '            Dim lProp3Actual As Int32 = oMineral.MinPropValueScore(lTemp)
        '            lTemp = oMineral.GetPropertyIndex(lProp4ID)
        '            Dim lProp4Actual As Int32 = oMineral.MinPropValueScore(lTemp)

        '            lProp1Actual = Math.Abs(lProp1Desired - lProp1Actual)
        '            lProp2Actual = Math.Abs(lProp2Desired - lProp2Actual)
        '            lProp3Actual = Math.Abs(lProp3Desired - lProp3Actual)
        '            lProp4Actual = Math.Abs(lProp4Desired - lProp4Actual)

        '            Dim lMaxDA As Int32 = Math.Max(lProp4Actual, Math.Max(Math.Max(lProp2Actual, lProp3Actual), lProp1Actual))

        '            'TODO: Remove me when ready!!!
        '            If oMineral.ObjectID = 1 Then
        '                lMaxDA = 5 * 23
        '            ElseIf oMineral.ObjectID = 5 Then
        '                lMaxDA = 23
        '            ElseIf oMineral.ObjectID = 9 Then
        '                lMaxDA = 255
        '            End If

        '            '(1+abs(Desired-Actual))*DPS*(Range+Scatter+1)/19
        '            Dim decTemp As Decimal = (1D + lMaxDA) * decDPS * (tpMaxRange.PropertyValue + tpScatterRadius.PropertyValue + 1D) / 19D
        '            If decTemp > Int32.MaxValue Then Return Int32.MaxValue
        '            If decTemp < 0 Then Return 0
        '            Return CInt(decTemp)
        '        End Function
        '#End Region

        Private Function GetMaxBeamDamage() As Int32
            'Dim fMult As Single = 1.0F + (Owner.oSpecials.yPulseMaxDmgBonus / 100.0F)
            'Return CInt(Math.Ceiling((InputEnergy / 10.0F) * CompressFactor * fMult))
            Return CInt(Math.Ceiling(GetMaxDamage() * 0.7F))
        End Function
        Private Function GetMinBeamDamage() As Int32
            'Dim fMult As Single = 1.0F + (Owner.oSpecials.yPulseMinDmgBonus / 100.0F)
            Return 0 'CInt(Math.Ceiling((lMaxBeamDmg \ 10) * fMult))
        End Function
        Private Function GetMaxImpactDamage() As Int32
            Return CInt(Math.Floor(GetMaxDamage() * 0.3F))
            'Return lMinBeamDmg
        End Function
        Private Function GetMinImpactDamage() As Int32
            Return 0 'lMaxImpactDmg \ 10
        End Function
        Public Function GetMaxDamage() As Int32
            'Dim lMaxBeam As Int32 = GetMaxBeamDamage()
            'Dim lMinBeam As Int32 = GetMinBeamDamage(lMaxBeam)
            'Return lMaxBeam + GetMaxImpactDamage(lMinBeam)
            Dim dblTemp As Double = (tpInputEnergy.PropertyValue / 10.0F) * (tpCompress.PropertyValue / 10.0F)
            If dblTemp > Int32.MaxValue Then Return Int32.MaxValue Else Return CInt(Math.Ceiling(dblTemp))
        End Function
        Public Function GetMinDamage() As Int32
            'Dim lMaxBeam As Int32 = GetMaxBeamDamage()
            'Dim lMinBeam As Int32 = GetMinBeamDamage(lMaxBeam)
            'Dim lMaxImpact As Int32 = GetMaxImpactDamage(lMinBeam)
            'Return lMinBeam + GetMinImpactDamage(lMaxImpact)
            Return 0
        End Function

        Private mlMineralIDs(4) As Int32
        Private Sub SetButtonClicked(ByVal lMinIdx As Int32)

            Dim oTech As New PulseTechComputer
            oTech.lHullTypeID = Me.lHullTypeID

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
                        stmCoilMat.SetMineralName(sMinName)
                    Case 1
                        stmAccelerator.SetMineralName(sMinName)
                    Case 2
                        stmCasing.SetMineralName(sMinName)
                    Case 3
                        stmFocuser.SetMineralName(sMinName)
                    Case 4
                        stmCompressChmbr.SetMineralName(sMinName)
                End Select
            End If

            'BuilderCostValueChange(False)
            CheckForDARequest()
        End Sub

        Public Overrides Sub CheckForDARequest()
            Dim bRequestDA As Boolean = True
            For X As Int32 = 0 To mlMineralIDs.GetUpperBound(0)
                If mlMineralIDs(X) < 1 Then
                    bRequestDA = False
                    Exit For
                End If
            Next X
            If bRequestDA = True Then CType(Me.ParentControl, frmWeaponBuilder).RequestDAValues(ObjectType.eWeaponTech, WeaponClassType.eEnergyPulse, CByte(lHullTypeID), mlMineralIDs(0), mlMineralIDs(1), mlMineralIDs(2), mlMineralIDs(3), mlMineralIDs(4), -1, 0, 0)
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