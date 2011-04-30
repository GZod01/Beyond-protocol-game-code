Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmWeaponBuilder
    'Interface created from Interface Builder
    Private Class fraProjectile
        Inherits fraWeaponBase

        'Private lblSize As UILabel
        'Private lbPierce As UILabel
        'Private lblROF As UILabel
        'Private lblMaxRange As UILabel
        'Private lblExplosionRadius As UILabel
        'Private txtCartridgeSize As UITextBox
        'Private txtPierceRatio As UITextBox
        'Private txtROF As UITextBox
        'Private txtMaxRange As UITextBox
        'Private txtExplosionRadius As UITextBox
        Private tpCartridgeSize As ctlTechProp
        Private tpPierceRatio As ctlTechProp
        Private tpROF As ctlTechProp
        Private tpMaxRange As ctlTechProp
        Private tpExplosionRadius As ctlTechProp

        Private lblPayloadType As UILabel
        Private lblProjectionType As UILabel
        Private lblBarrelMat As UILabel
        Private lblCasingMat As UILabel
        Private lblPayload1Mat As UILabel
        Private lblPayload2Mat As UILabel
        Private lblProjectionMat As UILabel

        'Private WithEvents cboProjectionMat As UIComboBox
        'Private WithEvents cboPayload2Mat As UIComboBox
        'Private WithEvents cboPayload1Mat As UIComboBox
        'Private WithEvents cboCasingMat As UIComboBox
        'Private WithEvents cboBarrelMat As UIComboBox
        Private stmProjectionMat As ctlSetTechMineral
        Private stmPayload2Mat As ctlSetTechMineral
        Private stmPayload1Mat As ctlSetTechMineral
        Private stmCasingMat As ctlSetTechMineral
        Private stmBarrelMat As ctlSetTechMineral

        Private WithEvents cboPayloadType As UIComboBox
        Private WithEvents cboProjectionType As UIComboBox

        Private tpProjectionMat As ctlTechProp
        Private tpPayload2Mat As ctlTechProp
        Private tpPayload1Mat As ctlTechProp
        Private tpCasingMat As ctlTechProp
        Private tpBarrelMat As ctlTechProp

        Private mbLoading As Boolean = True
        Private mlEntityIndex As Int32 = -1

        Private moTech As WeaponTech = Nothing

        Private mlMaxExpRadius As Int32 = 10
        Private mlMaxRng As Int32 = 50

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraProjectile initial props
            With Me
                .lWindowMetricID = BPMetrics.eWindow.eProjectileBuilder
                .ControlName = "fraProjectile"
                .Left = 150
                .Top = 101
                .Width = 470
                .Height = 335
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 1
                .Moveable = False
                .mbAcceptReprocessEvents = True
            End With

            'lblProjectionType initial props
            lblProjectionType = New UILabel(oUILib)
            With lblProjectionType
                .ControlName = "lblProjectionType"
                .Left = 10
                .Top = 10
                .Width = 120
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Projection Type:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblProjectionType, UIControl))

            ''lblSize initial props
            'lblSize = New UILabel(oUILib)
            'With lblSize
            '    .ControlName = "lblSize"
            '    .Left = 10
            '    .Top = 35
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Cartridge Size:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "Directly impacts damage, costs and size of the weapon." & vbCrLf & "Can be any whole number greater than 0."
            'End With
            'Me.AddChild(CType(lblSize, UIControl))

            ''lbPierce initial props
            'lbPierce = New UILabel(oUILib)
            'With lbPierce
            '    .ControlName = "lbPierce"
            '    .Left = 10
            '    .Top = 60
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Pierce Ratio:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "Indicates the ""Bluntness"" of the cartridge." & vbCrLf & "Higher values indicate more piercing damage" & vbCrLf & "while lower values produce impact damage." & vbCrLf & "Acceptable value is between 0 and 100."
            'End With
            'Me.AddChild(CType(lbPierce, UIControl))

            ''lblROF initial props
            'lblROF = New UILabel(oUILib)
            'With lblROF
            '    .ControlName = "lblROF"
            '    .Left = 10
            '    .Top = 85
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Rate of Fire:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            '    .ToolTipText = "Number of seconds between shots fired." & vbCrLf & "Can be any positive decimal number greater than 1/2 a second."
            'End With
            'Me.AddChild(CType(lblROF, UIControl))

            ''lblMaxRange initial props
            'lblMaxRange = New UILabel(oUILib)
            'With lblMaxRange
            '    .ControlName = "lblMaxRange"
            '    .Left = 10
            '    .Top = 110
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Maximum Range:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            'End With
            'Me.AddChild(CType(lblMaxRange, UIControl))

            'lblPayloadType initial props
            lblPayloadType = New UILabel(oUILib)
            With lblPayloadType
                .ControlName = "lblPayloadType"
                .Left = 10
                .Top = 135
                .Width = 120
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Payload Type:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
                .ToolTipText = "The type of additional damage the cartridge will inflict upon impact."
            End With
            Me.AddChild(CType(lblPayloadType, UIControl))

            ''lblExplosionRadius initial props
            'lblExplosionRadius = New UILabel(oUILib)
            'With lblExplosionRadius
            '    .ControlName = "lblExplosionRadius"
            '    .Left = 10
            '    .Top = 160
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Explosion Radius:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            'End With
            'Me.AddChild(CType(lblExplosionRadius, UIControl))

            ''txtCartridgeSize initial props
            'txtCartridgeSize = New UITextBox(oUILib)
            'With txtCartridgeSize
            '    .ControlName = "txtCartridgeSize"
            '    .Left = 140
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
            '    .MaxLength = 8
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .ToolTipText = lblSize.ToolTipText
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtCartridgeSize, UIControl))
            tpCartridgeSize = New ctlTechProp(oUILib)
            With tpCartridgeSize
                .ControlName = "tpCartridgeSize"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Cartridge Size:"
                .PropertyValue = 0
                .Top = 35
                .Visible = True
                .Width = 430
                .bNoMaxValue = True
                .ToolTipText = "Directly impacts damage, costs and size of the weapon." & vbCrLf & "Can be any whole number greater than 0."
            End With
            Me.AddChild(CType(tpCartridgeSize, UIControl))

            ''txtPierceRatio initial props
            'txtPierceRatio = New UITextBox(oUILib)
            'With txtPierceRatio
            '    .ControlName = "txtPierceRatio"
            '    .Left = 140
            '    .Top = 60
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
            '    .MaxLength = 0
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .ToolTipText = lbPierce.ToolTipText
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtPierceRatio, UIControl))
            tpPierceRatio = New ctlTechProp(oUILib)
            With tpPierceRatio
                .ControlName = "tpPierceRatio"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 100
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Pierce Ratio:"
                .PropertyValue = 0
                .Top = 60
                .Visible = True
                .Width = 430
                .ToolTipText = "Indicates the ""Bluntness"" of the cartridge." & vbCrLf & "Higher values indicate more piercing damage" & vbCrLf & "while lower values produce impact damage." & vbCrLf & "Acceptable value is between 0 and 100."
            End With
            Me.AddChild(CType(tpPierceRatio, UIControl))

            ''txtROF initial props
            'txtROF = New UITextBox(oUILib)
            'With txtROF
            '    .ControlName = "txtROF"
            '    .Left = 140
            '    .Top = 85
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
            '    .MaxLength = 0
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
                .bNoMaxValue = True
                .SetDivisor(30)
                .ToolTipText = "Number of seconds between shots fired." & vbCrLf & "Can be any positive decimal number greater than a second."
            End With
            Me.AddChild(CType(tpROF, UIControl))

            ''txtMaxRange initial props
            'txtMaxRange = New UITextBox(oUILib)
            'With txtMaxRange
            '    .ControlName = "txtMaxRange"
            '    .Left = 140
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
            '    .MaxLength = 0
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtMaxRange, UIControl))
            tpMaxRange = New ctlTechProp(oUILib)
            With tpMaxRange
                .ControlName = "tpMaxRange"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 100
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Maximum Range:"
                .PropertyValue = 0
                .Top = 110
                .Visible = True
                .Width = 430
            End With
            Me.AddChild(CType(tpMaxRange, UIControl))

            ''txtExplosionRadius initial props
            'txtExplosionRadius = New UITextBox(oUILib)
            'With txtExplosionRadius
            '    .ControlName = "txtExplosionRadius"
            '    .Left = 140
            '    .Top = 160
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
            '    .MaxLength = 0
            '    .BorderColor = muSettings.InterfaceBorderColor
            '    .bNumericOnly = True
            'End With
            'Me.AddChild(CType(txtExplosionRadius, UIControl))
            tpExplosionRadius = New ctlTechProp(oUILib)
            With tpExplosionRadius
                .ControlName = "tpExplosionRadius"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 100
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Explosion Radius:"
                .PropertyValue = 0
                .Top = 160
                .Visible = True
                .Width = 430
            End With
            Me.AddChild(CType(tpExplosionRadius, UIControl))

            'lblBarrelMat initial props
            lblBarrelMat = New UILabel(oUILib)
            With lblBarrelMat
                .ControlName = "lblBarrelMat"
                .Left = 10
                .Top = 200
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Barrel:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblBarrelMat, UIControl))

            tpBarrelMat = New ctlTechProp(oUILib)
            With tpBarrelMat
                .ControlName = "tpBarrelMat"
                .Enabled = True
                .Height = 18
                .Left = lblBarrelMat.Left + lblBarrelMat.Width + 160        '10 for left/right and 150 for cbo
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Barrel:"
                .PropertyValue = 0
                .Top = 200
                .Visible = True
                .bNoMaxValue = True
                .Width = Me.Width - .Left - 10
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpBarrelMat, UIControl))

            'lblCasingMat initial props
            lblCasingMat = New UILabel(oUILib)
            With lblCasingMat
                .ControlName = "lblCasingMat"
                .Left = 10
                .Top = 225
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
            Me.AddChild(CType(lblCasingMat, UIControl))

            tpCasingMat = New ctlTechProp(oUILib)
            With tpCasingMat
                .ControlName = "tpCasingMat"
                .Enabled = True
                .Height = 18
                .Left = tpBarrelMat.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Casing:"
                .PropertyValue = 0
                .Top = 225
                .Visible = True
                .bNoMaxValue = True
                .Width = tpBarrelMat.Width
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpCasingMat, UIControl))

            'lblPayload1Mat initial props
            lblPayload1Mat = New UILabel(oUILib)
            With lblPayload1Mat
                .ControlName = "lblPayload1Mat"
                .Left = 10
                .Top = 250
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Payload 1:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblPayload1Mat, UIControl))

            tpPayload1Mat = New ctlTechProp(oUILib)
            With tpPayload1Mat
                .ControlName = "tpPayload1Mat"
                .Enabled = True
                .Height = 18
                .Left = tpBarrelMat.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Payload1:"
                .PropertyValue = 0
                .Top = 250
                .Visible = True
                .bNoMaxValue = True
                .Width = tpBarrelMat.Width
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpPayload1Mat, UIControl))

            'lblPayload2Mat initial props
            lblPayload2Mat = New UILabel(oUILib)
            With lblPayload2Mat
                .ControlName = "lblPayload2Mat"
                .Left = 10
                .Top = 275
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Payload 2:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblPayload2Mat, UIControl))

            tpPayload2Mat = New ctlTechProp(oUILib)
            With tpPayload2Mat
                .ControlName = "tpPayload2Mat"
                .Enabled = True
                .Height = 18
                .Left = tpBarrelMat.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Payload2:"
                .PropertyValue = 0
                .Top = 275
                .Visible = True
                .bNoMaxValue = True
                .Width = tpBarrelMat.Width
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpPayload2Mat, UIControl))

            'lblProjectionMat initial props
            lblProjectionMat = New UILabel(oUILib)
            With lblProjectionMat
                .ControlName = "lblProjectionMat"
                .Left = 10
                .Top = 300
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Projection:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblProjectionMat, UIControl))

            tpProjectionMat = New ctlTechProp(oUILib)
            With tpProjectionMat
                .ControlName = "tpProjectionMat"
                .Enabled = True
                .Height = 18
                .Left = tpBarrelMat.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Projection:"
                .PropertyValue = 0
                .Top = 300
                .Visible = True
                .bNoMaxValue = True
                .Width = tpBarrelMat.Width
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpProjectionMat, UIControl))

            ''cboProjectionMat initial props
            'cboProjectionMat = New UIComboBox(oUILib)
            'With cboProjectionMat
            '    .ControlName = "cboProjectionMat"
            '    .Left = lblProjectionMat.Left + lblProjectionMat.Width + 5
            '    .Top = 300
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
            '    .ToolTipText = "Select a Projection Type First."
            '    .mbAcceptReprocessEvents = True
            'End With
            'Me.AddChild(CType(cboProjectionMat, UIControl))
            stmProjectionMat = New ctlSetTechMineral(oUILib)
            With stmProjectionMat
                .ControlName = "stmProjectionMat"
                .Left = lblProjectionMat.Left + lblProjectionMat.Width + 5
                .Top = 300
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 4
            End With
            Me.AddChild(CType(stmProjectionMat, UIControl))

            ''cboPayload2Mat initial props
            'cboPayload2Mat = New UIComboBox(oUILib)
            'With cboPayload2Mat
            '    .ControlName = "cboPayload2Mat"
            '    .Left = cboProjectionMat.Left
            '    .Top = 275
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
            'Me.AddChild(CType(cboPayload2Mat, UIControl))
            stmPayload2Mat = New ctlSetTechMineral(oUILib)
            With stmPayload2Mat
                .ControlName = "stmPayload2Mat"
                .Left = stmProjectionMat.Left
                .Top = 275
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 3
            End With
            Me.AddChild(CType(stmPayload2Mat, UIControl))

            ''cboPayload1Mat initial props
            'cboPayload1Mat = New UIComboBox(oUILib)
            'With cboPayload1Mat
            '    .ControlName = "cboPayload1Mat"
            '    .Left = cboProjectionMat.Left
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
            '    .mbAcceptReprocessEvents = True
            'End With
            'Me.AddChild(CType(cboPayload1Mat, UIControl))
            stmPayload1Mat = New ctlSetTechMineral(oUILib)
            With stmPayload1Mat
                .ControlName = "stmPayload1Mat"
                .Left = stmProjectionMat.Left
                .Top = 250
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 2
            End With
            Me.AddChild(CType(stmPayload1Mat, UIControl))

            ''cboCasingMat initial props
            'cboCasingMat = New UIComboBox(oUILib)
            'With cboCasingMat
            '    .ControlName = "cboCasingMat"
            '    .Left = cboProjectionMat.Left
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
            '    .mbAcceptReprocessEvents = True
            'End With
            'Me.AddChild(CType(cboCasingMat, UIControl))
            stmCasingMat = New ctlSetTechMineral(oUILib)
            With stmCasingMat
                .ControlName = "stmCasingMat"
                .Left = stmProjectionMat.Left
                .Top = 225
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 1
            End With
            Me.AddChild(CType(stmCasingMat, UIControl))

            ''cboBarrelMat initial props
            'cboBarrelMat = New UIComboBox(oUILib)
            'With cboBarrelMat
            '    .ControlName = "cboBarrelMat"
            '    .Left = cboProjectionMat.Left
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
            '    .mbAcceptReprocessEvents = True
            'End With
            'Me.AddChild(CType(cboBarrelMat, UIControl))
            stmBarrelMat = New ctlSetTechMineral(oUILib)
            With stmBarrelMat
                .ControlName = "stmBarrelMat"
                .Left = stmProjectionMat.Left
                .Top = 200
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 0
            End With
            Me.AddChild(CType(stmBarrelMat, UIControl))

            'cboPayloadType initial props
            cboPayloadType = New UIComboBox(oUILib)
            With cboPayloadType
                .ControlName = "cboPayloadType"
                .Left = 165
                .Top = 135
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
                .mbAcceptReprocessEvents = True
                .ToolTipText = lblPayloadType.ToolTipText
            End With
            Me.AddChild(CType(cboPayloadType, UIControl))

            'cboProjectionType initial props
            cboProjectionType = New UIComboBox(oUILib)
            With cboProjectionType
                .ControlName = "cboProjectionType"
                .Left = 165
                .Top = 10
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
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(cboProjectionType, UIControl))

            AddHandler stmBarrelMat.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmCasingMat.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmPayload1Mat.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmPayload2Mat.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmProjectionMat.SetButtonClicked, AddressOf SetButtonClicked


            AddHandler tpBarrelMat.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpCartridgeSize.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpCasingMat.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpExplosionRadius.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpMaxRange.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpPayload1Mat.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpPayload2Mat.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpPierceRatio.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpProjectionMat.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpROF.PropertyValueChanged, AddressOf tp_ValueChanged

            mbLoading = False

            FillValues()
        End Sub

        Private Sub FillValues()
            'cboProjectionMat.Clear() : cboPayload2Mat.Clear() : cboPayload1Mat.Clear() : cboCasingMat.Clear() : cboBarrelMat.Clear()
            cboPayloadType.Clear() : cboProjectionType.Clear()

            'Dim lSorted() As Int32 = GetSortedMineralIdxArray(False)

            ''Now... loop through our minerals
            'If lSorted Is Nothing = False Then
            '    For X As Int32 = 0 To lSorted.GetUpperBound(0)
            '        If lSorted(X) <> -1 Then
            '            'If goMinerals(lSorted(X)).IsFullyStudied = False Then Continue For
            '            cboProjectionMat.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboProjectionMat.ItemData(cboProjectionMat.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboPayload2Mat.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboPayload2Mat.ItemData(cboPayload2Mat.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboPayload1Mat.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboPayload1Mat.ItemData(cboPayload1Mat.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboCasingMat.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboCasingMat.ItemData(cboCasingMat.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboBarrelMat.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboBarrelMat.ItemData(cboBarrelMat.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '        End If
            '    Next X
            'End If

            Dim lPayloadAvail As Int32 = 0
            If goCurrentPlayer Is Nothing = False Then lPayloadAvail = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.ePayloadTypeAvailable)
            cboPayloadType.AddItem("No Extra Payload")
            cboPayloadType.ItemData(cboPayloadType.NewIndex) = 0
            cboPayloadType.AddItem("Explosive")
            cboPayloadType.ItemData(cboPayloadType.NewIndex) = 1
            If lPayloadAvail > 1 Then
                cboPayloadType.AddItem("Chemical")
                cboPayloadType.ItemData(cboPayloadType.NewIndex) = 2
            End If
            If lPayloadAvail > 2 Then
                cboPayloadType.AddItem("Magnetic")
                cboPayloadType.ItemData(cboPayloadType.NewIndex) = 3
            End If
            cboPayloadType.ListIndex = 0

            cboProjectionType.AddItem("Explosive")
            cboProjectionType.ItemData(cboProjectionType.NewIndex) = 0
            cboProjectionType.AddItem("Magnetic")
            cboProjectionType.ItemData(cboProjectionType.NewIndex) = 1

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
            yMsg(lPos) = WeaponClassType.eProjectile : lPos += 1
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
                If tpBarrelMat.PropertyLocked = True Then lTempMin = CInt(tpBarrelMat.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpCasingMat.PropertyLocked = True Then lTempMin = CInt(tpCasingMat.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpPayload1Mat.PropertyLocked = True Then lTempMin = CInt(tpPayload1Mat.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpPayload2Mat.PropertyLocked = True Then lTempMin = CInt(tpPayload2Mat.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpProjectionMat.PropertyLocked = True Then lTempMin = CInt(tpProjectionMat.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4

                yMsg(lPos) = CByte(Me.lHullTypeID) : lPos += 1
            End With

            'Now... for the data specific this weapon type...
            yMsg(lPos) = CByte(cboProjectionType.ItemData(cboProjectionType.ListIndex)) : lPos += 1

            Dim lTempVal As Int32
            lTempVal = CInt(tpCartridgeSize.PropertyValue)
            System.BitConverter.GetBytes(lTempVal).CopyTo(yMsg, lPos) : lPos += 4

            yMsg(lPos) = CByte(tpPierceRatio.PropertyValue) : lPos += 1

            Dim fROF As Single = CSng(tpROF.PropertyValue)
            'fROF *= 30.0F
            Dim iROF As Int16 = CShort(fROF)
            System.BitConverter.GetBytes(iROF).CopyTo(yMsg, lPos) : lPos += 2

            iROF = CShort(tpMaxRange.PropertyValue)
            System.BitConverter.GetBytes(iROF).CopyTo(yMsg, lPos) : lPos += 2

            yMsg(lPos) = CByte(cboPayloadType.ItemData(cboPayloadType.ListIndex)) : lPos += 1
            yMsg(lPos) = CByte(tpExplosionRadius.PropertyValue) : lPos += 1

            For X As Int32 = 0 To mlMineralIDs.GetUpperBound(0)
                System.BitConverter.GetBytes(mlMineralIDs(X)).CopyTo(yMsg, lPos) : lPos += 4
            Next X

            'System.BitConverter.GetBytes(cboBarrelMat.ItemData(cboBarrelMat.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
            'System.BitConverter.GetBytes(cboCasingMat.ItemData(cboCasingMat.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
            'lTempVal = cboPayload1Mat.ItemData(cboPayload1Mat.ListIndex)
            'System.BitConverter.GetBytes(lTempVal).CopyTo(yMsg, lPos) : lPos += 4
            'If cboPayload2Mat.ListIndex <> -1 Then lTempVal = cboPayload2Mat.ItemData(cboPayload2Mat.ListIndex)
            'System.BitConverter.GetBytes(lTempVal).CopyTo(yMsg, lPos) : lPos += 4
            'System.BitConverter.GetBytes(cboProjectionMat.ItemData(cboProjectionMat.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4

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

            If cboPayloadType.ListIndex = -1 Then
                MyBase.moUILib.AddNotification("Please select a Payload Type.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            If cboProjectionType.ListIndex = -1 Then
                MyBase.moUILib.AddNotification("Please select a Projection Type.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            If cboPayloadType.ListIndex < 1 Then
                If tpExplosionRadius.PropertyValue <> 0 Then tpExplosionRadius.PropertyValue = 0
            End If

            If tpCartridgeSize.PropertyValue > Int32.MaxValue Then
                MyBase.moUILib.AddNotification("Cartridge Size is too large!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            Dim lSize As Int32 = CInt(tpCartridgeSize.PropertyValue)
            If lSize < 1 Then
                MyBase.moUILib.AddNotification("Cartridge Size must be a whole number greater than 0.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            Dim lPierce As Int32 = CInt(tpPierceRatio.PropertyValue)
            If lPierce < 0 OrElse lPierce > 100 Then
                MyBase.moUILib.AddNotification("Pierce Ratio must be a whole number between 0 and 100.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            Dim fROF As Single = CSng(tpROF.PropertyValue)

            If fROF < 0.5F Then
                MyBase.moUILib.AddNotification("Please enter a ROF of at least 1/2 a second.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            'fROF *= 30.0F
            If fROF < 30 Then
                MyBase.moUILib.AddNotification("Please enter a ROF of at least a second.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            ElseIf fROF > Int16.MaxValue Then
                MyBase.moUILib.AddNotification("Maximum value for ROF is: " & (Int16.MaxValue / 30.0F).ToString("#,###.#0"), Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            Dim lMaxRng As Int32 = CInt(tpMaxRange.PropertyValue)
            If lMaxRng < 1 OrElse lMaxRng > mlMaxRng Then
                MyBase.moUILib.AddNotification("Maximum Range must be a whole number between 1 and " & mlMaxRng & ".", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            Dim lExpRad As Int32 = CInt(tpExplosionRadius.PropertyValue)
            If lExpRad < 0 OrElse lExpRad > mlMaxExpRadius Then
                MyBase.moUILib.AddNotification("Explosion Radius must be a whole number between 0 and " & mlMaxExpRadius & ".", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            Return True
        End Function

        Public Overrides Sub ViewResults(ByRef oTech As WeaponTech, ByVal lProdFactor As Integer)
            If oTech Is Nothing Then Return

            moTech = oTech

            With CType(oTech, ProjectileWeaponTech)
                mbIgnoreValueChange = True
                Me.cboProjectionType.FindComboItemData(.ProjectionType)
                Me.cboPayloadType.FindComboItemData(.PayloadType)
                mbLoading = True

                MyBase.lHullTypeID = .HullTypeID

                tpCartridgeSize.PropertyValue = .CartridgeSize
                If .CartridgeSize > 0 Then tpCartridgeSize.PropertyLocked = True
                tpExplosionRadius.PropertyValue = .ExplosionRadius
                If .ExplosionRadius > 0 Then tpExplosionRadius.PropertyLocked = True
                tpMaxRange.PropertyValue = .MaxRange
                If .MaxRange > 0 Then tpMaxRange.PropertyLocked = True
                tpPierceRatio.PropertyValue = .PierceRatio
                tpPierceRatio.PropertyLocked = True
                If tpROF.MaxValue < .ROF Then tpROF.MaxValue = .ROF
                tpROF.PropertyValue = .ROF '(.ROF / 30.0F)
                tpROF.PropertyLocked = True

                mlSelectedMineralIdx = 0 : Mineral_Selected(.BarrelMineralID)
                mlSelectedMineralIdx = 1 : Mineral_Selected(.CasingMineralID)
                mlSelectedMineralIdx = 2 : Mineral_Selected(.Payload1MineralID)
                mlSelectedMineralIdx = 3 : Mineral_Selected(.Payload2MineralID)
                mlSelectedMineralIdx = 4 : Mineral_Selected(.ProjectionMineralID)
                'Me.cboBarrelMat.FindComboItemData(.BarrelMineralID)
                'Me.cboCasingMat.FindComboItemData(.CasingMineralID)
                'Me.cboPayload1Mat.FindComboItemData(.Payload1MineralID)
                'Me.cboPayload2Mat.FindComboItemData(.Payload2MineralID)
                'Me.cboProjectionMat.FindComboItemData(.ProjectionMineralID)

                tpBarrelMat.PropertyValue = .lSpecifiedMin1
                tpCasingMat.PropertyValue = .lSpecifiedMin2
                tpPayload1Mat.PropertyValue = .lSpecifiedMin3
                tpPayload2Mat.PropertyValue = .lSpecifiedMin4
                tpProjectionMat.PropertyValue = .lSpecifiedMin5
                If .lSpecifiedMin1 <> -1 Then tpBarrelMat.PropertyLocked = True
                If .lSpecifiedMin2 <> -1 Then tpCasingMat.PropertyLocked = True
                If .lSpecifiedMin3 <> -1 Then tpPayload1Mat.PropertyLocked = True
                If .lSpecifiedMin4 <> -1 Then tpPayload2Mat.PropertyLocked = True
                If .lSpecifiedMin5 <> -1 Then tpProjectionMat.PropertyLocked = True

                MyBase.mfrmBuilderCost.SetAndLockValues(oTech)

                mbLoading = False
                mbIgnoreValueChange = False
                BuilderCostValueChange(False)
            End With

        End Sub

        Private Sub SetCboHelpers()
            'Dim bTempSens As Boolean = goCurrentPlayer.PlayerKnowsProperty("TEMPERATURE SENSITIVITY")
            'Dim bThermCond As Boolean = goCurrentPlayer.PlayerKnowsProperty("THERMAL CONDUCTANCE")
            'Dim bThermExp As Boolean = goCurrentPlayer.PlayerKnowsProperty("THERMAL EXPANSION")
            'Dim bDensity As Boolean = goCurrentPlayer.PlayerKnowsProperty("DENSITY")
            'Dim bMalleable As Boolean = goCurrentPlayer.PlayerKnowsProperty("MALLEABLE")
            'Dim bHardness As Boolean = goCurrentPlayer.PlayerKnowsProperty("HARDNESS")

            Dim oSB As New System.Text.StringBuilder()

            ''Barrel
            oSB.Length = 0
            oSB.Append("Materials ")
            'If bTempSens = True Then 
            oSB.Append("that are not Sensitive to Temperature ")
            'If bThermExp = True Then 
            oSB.Append("with low Thermal Expansion ")
            'If bThermCond = True Then 
            oSB.Append("that are highly Thermal Conductive ")
            oSB.Append("work best.")
            'If bTempSens = True OrElse bThermExp = True OrElse bThermCond = True Then 
            stmBarrelMat.ToolTipText = oSB.ToString 'Else cboBarrelMat.ToolTipText = "We are not quite sure of the best way to engineer this."

            ''Casing
            oSB.Length = 0
            'If bDensity = True Then 
            oSB.Append("Low Density materials ") 'Else oSB.Append("Materials ")
            'If bTempSens = True Then 
            oSB.Append("that are very Sensitive to Temperature ")
            'If bThermCond = True Then 
            oSB.Append("with a low Thermal Conductance ")
            oSB.Append("are preferred.")
            'If bTempSens = True OrElse bDensity = True OrElse bThermCond = True Then 
            stmCasingMat.ToolTipText = oSB.ToString 'Else cboCasingMat.ToolTipText = "We are not quite sure of the best way to engineer this."

            ''Payload1
            oSB.Length = 0
            'If bDensity = True Then 
            oSB.Append("High Density materials ") 'Else oSB.Append("Materials ")
            'If bHardness = True Then 
            oSB.Append("with high Hardness ratings ")
            'If bMalleable = True Then 
            oSB.Append("that have low Malleable characteristics ")
            oSB.Append("would do the most damage.")
            'If bMalleable = True OrElse bDensity = True OrElse bHardness = True Then 
            stmPayload1Mat.ToolTipText = oSB.ToString 'Else cboPayload1Mat.ToolTipText = "We are not quite sure of the best way to engineer this."

            If goCurrentPlayer Is Nothing = False Then
                mlMaxExpRadius = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eProjectileExplosionRadius)
                mlMaxRng = Math.Min(255, goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eProjectileMaxRng))
            End If

            tpMaxRange.ToolTipText = "Maximum range this weapon will fire." & vbCrLf & "Valid values are between 10 and " & mlMaxRng & "."
            tpExplosionRadius.ToolTipText = "Area of effect for this weapon's damage." & vbCrLf & "Requires a Payload Type to be selected." & vbCrLf & "Valid values are between 0 and " & mlMaxExpRadius & "."

            tpMaxRange.MaxValue = mlMaxRng
            tpMaxRange.MinValue = 10
            tpExplosionRadius.MaxValue = mlMaxExpRadius
            tpCartridgeSize.MaxValue = 1000 '0000
            tpCartridgeSize.MinValue = 1
            tpROF.MinValue = 30
            tpROF.MaxValue = 2000
            tpROF.PropertyValue = 1800 'tpROF.MinValue
            tpROF.blAbsoluteMaximum = 30000

            tpMaxRange.SetToInitialDefault()
            tpPierceRatio.SetToInitialDefault()
        End Sub

        Private Sub cboProjectionType_ItemSelected(ByVal lItemIndex As Integer) Handles cboProjectionType.ItemSelected
            If mbLoading = True Then Return
            Dim sToolTip As String = "We are not quite sure of the best way to engineer this."
            If lItemIndex = 0 Then
                'If goCurrentPlayer.PlayerKnowsProperty("COMBUSTIVENESS") = True Then
                sToolTip = "Materials with a high Combustiveness property work best."
                'End If
            ElseIf lItemIndex = 1 Then
                'If goCurrentPlayer.PlayerKnowsProperty("MAGNETIC PRODUCTION") = True Then
                sToolTip = "Materials with a high Magnetic Production property work best."
                'End If
            End If
            stmProjectionMat.ToolTipText = sToolTip
            CheckForDARequest()
        End Sub

        Private Sub cboPayloadType_ItemSelected(ByVal lItemIndex As Integer) Handles cboPayloadType.ItemSelected
            If mbLoading = True Then Return
            Dim sToolTip As String = "We are not quite sure of the best way to engineer this."

            tpExplosionRadius.Enabled = True

            If lItemIndex = 0 Then
                sToolTip = "Materials with a high Density are perfect for this."
                tpExplosionRadius.Enabled = False
                tpExplosionRadius.PropertyValue = 0
            ElseIf lItemIndex = 1 Then
                'If goCurrentPlayer.PlayerKnowsProperty("COMBUSTIVENESS") = True Then
                sToolTip = "Materials with high Combustiveness are perfect for this."
                'End If
            ElseIf lItemIndex = 2 Then
                'If goCurrentPlayer.PlayerKnowsProperty("CHEMICAL REACTANCE") = True Then
                sToolTip = "Materials with high Chemical Reactance are perfect for this."
                'End If
            ElseIf lItemIndex = 3 Then
                'If goCurrentPlayer.PlayerKnowsProperty("MAGNETIC PRODUCTION") = True Then
                sToolTip = "Materials with high Magnetic Production are perfect for this."
                'End If
            End If
            stmPayload2Mat.ToolTipText = sToolTip
            CheckForDARequest()
        End Sub

        Private Sub fraProjectile_OnRender() Handles Me.OnRender
            'Ok, first, is moTech nothing?
            If moTech Is Nothing = False Then
                'Ok, now, do the values used on the form currently match the tech?
                Dim bChanged As Boolean = False
                With CType(moTech, ProjectileWeaponTech)
                    bChanged = tpCartridgeSize.PropertyValue <> .CartridgeSize OrElse tpPierceRatio.PropertyValue <> .PierceRatio OrElse _
                      tpMaxRange.PropertyValue <> .MaxRange OrElse tpExplosionRadius.PropertyValue <> .ExplosionRadius

                    If bChanged = False Then
                        bChanged = tpROF.PropertyValue <> .ROF
                    End If
                    If bChanged = False Then
                        Dim lProjMat As Int32 = mlMineralIDs(4)
                        Dim lPay1 As Int32 = mlMineralIDs(2)
                        Dim lPay2 As Int32 = mlMineralIDs(3)
                        Dim lCaseID As Int32 = mlMineralIDs(1)
                        Dim lBarrel As Int32 = mlMineralIDs(0)
                        Dim lPType As Int32 = -1
                        Dim lProjType As Int32 = -1

                        'If cboProjectionMat.ListIndex <> -1 Then lProjMat = cboProjectionMat.ItemData(cboProjectionMat.ListIndex)
                        'If cboPayload1Mat.ListIndex <> -1 Then lPay1 = cboPayload1Mat.ItemData(cboPayload1Mat.ListIndex)
                        'If cboPayload2Mat.ListIndex <> -1 Then lPay2 = cboPayload2Mat.ItemData(cboPayload2Mat.ListIndex)
                        'If cboCasingMat.ListIndex <> -1 Then lCaseID = cboCasingMat.ItemData(cboCasingMat.ListIndex)
                        'If cboBarrelMat.ListIndex <> -1 Then lBarrel = cboBarrelMat.ItemData(cboBarrelMat.ListIndex)
                        If cboPayloadType.ListIndex <> -1 Then lPType = cboPayloadType.ItemData(cboPayloadType.ListIndex)
                        If cboProjectionType.ListIndex <> -1 Then lProjType = cboProjectionType.ItemData(cboProjectionType.ListIndex)

                        bChanged = .ProjectionMineralID <> lProjMat OrElse .Payload1MineralID <> lPay1 OrElse .Payload2MineralID <> lPay2 OrElse _
                          .CasingMineralID <> lCaseID OrElse .BarrelMineralID <> lBarrel OrElse .PayloadType <> lPType OrElse lProjType <> .ProjectionType OrElse _
                          (tpBarrelMat.PropertyLocked <> (.lSpecifiedMin1 <> -1)) OrElse (tpCasingMat.PropertyLocked <> (.lSpecifiedMin2 <> -1)) OrElse _
                          (tpPayload1Mat.PropertyLocked <> (.lSpecifiedMin3 <> -1)) OrElse (tpPayload2Mat.PropertyLocked <> (.lSpecifiedMin4 <> -1)) OrElse _
                          (tpProjectionMat.PropertyLocked <> (.lSpecifiedMin5 <> -1))

                        If bChanged = False Then

                            bChanged = (.lSpecifiedMin1 <> -1 AndAlso .lSpecifiedMin1 <> tpBarrelMat.PropertyValue) OrElse _
                                        (.lSpecifiedMin2 <> -1 AndAlso .lSpecifiedMin2 <> tpCasingMat.PropertyValue) OrElse _
                                        (.lSpecifiedMin3 <> -1 AndAlso .lSpecifiedMin3 <> tpPayload1Mat.PropertyValue) OrElse _
                                        (.lSpecifiedMin4 <> -1 AndAlso .lSpecifiedMin4 <> tpPayload2Mat.PropertyValue) OrElse _
                                        (.lSpecifiedMin5 <> -1 AndAlso .lSpecifiedMin5 <> tpProjectionMat.PropertyValue)
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

        'Try
        '    Dim ofrmMin As frmMinDetail = CType(MyBase.moUILib.GetWindow("frmMinDetail"), frmMinDetail)
        '    If ofrmMin Is Nothing Then Return

        '    Select Case sComboBoxName
        '        Case cboBarrelMat.ControlName
        '            ofrmMin.ClearHighlights()
        '            For X As Int32 = 0 To glMineralPropertyUB
        '                If glMineralPropertyIdx(X) <> -1 Then
        '                    Select Case goMineralProperty(X).MineralPropertyName.ToUpper
        '                        Case "TEMPERATURE SENSITIVITY"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
        '                        Case "THERMAL EXPANSION"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
        '                        Case "THERMAL CONDUCTION"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 4)
        '                    End Select
        '                End If
        '            Next X
        '        Case cboCasingMat.ControlName
        '            ofrmMin.ClearHighlights()
        '            For X As Int32 = 0 To glMineralPropertyUB
        '                If glMineralPropertyIdx(X) <> -1 Then
        '                    Select Case goMineralProperty(X).MineralPropertyName.ToUpper
        '                        Case "DENSITY"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 1)
        '                        Case "TEMPERATURE SENSITIVITY"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 5)
        '                        Case "THERMAL CONDUCTION"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
        '                    End Select
        '                End If
        '            Next X
        '        Case cboPayload1Mat.ControlName
        '            ofrmMin.ClearHighlights()
        '            For X As Int32 = 0 To glMineralPropertyUB
        '                If glMineralPropertyIdx(X) <> -1 Then
        '                    Select Case goMineralProperty(X).MineralPropertyName.ToUpper
        '                        Case "DENSITY"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 6)
        '                        Case "HARDNESS"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 4)
        '                        Case "MALLEABLE"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
        '                    End Select
        '                End If
        '            Next X
        '        Case cboPayload2Mat.ControlName
        '            ofrmMin.ClearHighlights()

        '            Dim sProp As String = ""
        '            Dim yPV As Byte = 10
        '            Select Case cboPayloadType.ListIndex
        '                Case 1
        '                    sProp = "COMBUSTIVENESS"
        '                    yPV = 6
        '                Case 2
        '                    sProp = "CHEMICAL REACTANCE"
        '                    yPV = 5
        '                Case 3
        '                    sProp = "MAGNETIC PRODUCTION"
        '                    yPV = 6
        '                Case Else
        '                    sProp = "DENSITY"
        '                    yPV = 6
        '            End Select
        '            If sProp <> "" Then
        '                For X As Int32 = 0 To glMineralPropertyUB
        '                    If glMineralPropertyIdx(X) <> -1 Then
        '                        If goMineralProperty(X).MineralPropertyName.ToUpper = sProp Then
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, yPV)
        '                            Exit For
        '                        End If
        '                    End If
        '                Next X
        '            End If
        '        Case cboProjectionMat.ControlName
        '            ofrmMin.ClearHighlights()
        '            Dim sProp As String = ""
        '            Dim yPV As Byte = 10
        '            If cboProjectionType.ListIndex = 0 Then
        '                sProp = "COMBUSTIVENESS"
        '                yPV = 5
        '            ElseIf cboProjectionType.ListIndex = 1 Then
        '                sProp = "MAGNETIC PRODUCTION"
        '                yPV = 6
        '            End If
        '            If sProp <> "" Then
        '                For X As Int32 = 0 To glMineralPropertyUB
        '                    If glMineralPropertyIdx(X) <> -1 Then
        '                        If goMineralProperty(X).MineralPropertyName.ToUpper = sProp Then
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, yPV)
        '                            Exit For
        '                        End If
        '                    End If
        '                Next X
        '            End If
        '    End Select
        'Catch
        'End Try

        Private Sub tp_ValueChanged(ByVal sPropName As String, ByRef ctl As ctlTechProp)
            tpROF.MaxValue = Math.Min(Int16.MaxValue, Math.Max(tpROF.MaxValue, tpROF.PropertyValue + 500))
            BuilderCostValueChange(False)
        End Sub
        'Private Sub ComboValueChanged(ByVal lIndex As Int32) Handles cboBarrelMat.ItemSelected, cboCasingMat.ItemSelected, cboPayload1Mat.ItemSelected, cboPayload2Mat.ItemSelected, cboProjectionMat.ItemSelected, cboPayloadType.ItemSelected, cboProjectionType.ItemSelected
        '    BuilderCostValueChange()
        'End Sub

        Public Overrides Sub BuilderCostValueChange(ByVal bAutoFill As Boolean)
            If mbIgnoreValueChange = True Then Return
            mbIgnoreValueChange = True
            mbImpossibleDesign = True

            Dim oProjTechComputer As New ProjectileTechComputer
            With oProjTechComputer

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

                tpCartridgeSize.MaxValue = lMaxHullSize
                If tpCartridgeSize.MaxValue < tpCartridgeSize.PropertyValue Then tpCartridgeSize.PropertyValue = tpCartridgeSize.MaxValue

                With CType(Me.ParentControl, frmWeaponBuilder)
                    oProjTechComputer.SetMinDAValues(.ml_Min1DA, .ml_Min2DA, .ml_Min3DA, .ml_Min4DA, .ml_Min5DA, .ml_Min6DA)
                End With

                Dim fROF As Single = Math.Max(1, tpROF.PropertyValue) / 30.0F
                Dim lPayloadDmg As Int32 = CInt(tpCartridgeSize.PropertyValue \ 2)
                Dim lPIDmg As Int32 = lPayloadDmg
                Dim lMagDmg As Int32 = 0

                If cboProjectionType.ListIndex > -1 Then
                    .yProjectionType = CByte(cboProjectionType.ItemData(cboProjectionType.ListIndex))
                End If

                If .yProjectionType = 1 Then
                    lMagDmg = lPIDmg \ 2
                    lPIDmg = lMagDmg
                End If

                Dim lMinPierce As Int32 = 0
                Dim lMinImpact As Int32 = 0
                Dim lMaxPierce As Int32 = CInt(lPIDmg * (tpPierceRatio.PropertyValue / 100.0F))
                Dim lMaxImpact As Int32 = CInt(lPIDmg * ((100 - tpPierceRatio.PropertyValue) / 100.0F))
                Dim lMinPayload As Int32 = 0
                Dim lMaxPayload As Int32 = lPayloadDmg

                Dim lMaxFlame As Int32 = 0
                Dim lMinFlame As Int32 = 0
                Dim lMaxChemical As Int32 = 0
                Dim lMinChemical As Int32 = 0
                Dim lMaxECM As Int32 = lMagDmg
                Dim lMinECM As Int32 = 0

                Dim yPayloadType As Byte = 0
                If cboPayloadType.ListIndex > -1 Then
                    yPayloadType = CByte(cboPayloadType.ItemData(cboPayloadType.ListIndex))
                    Select Case yPayloadType
                        Case 1
                            lMaxFlame = lMaxPayload : lMinFlame = lMinPayload
                        Case 2
                            lMaxChemical = lMaxPayload : lMinChemical = lMinPayload
                        Case 3
                            lMaxECM += lMaxPayload : lMinECM = lMinPayload
                        Case Else
                            lMaxImpact += lMaxPayload
                            lMinImpact += lMinPayload
                    End Select
                Else
                    lMaxImpact += lMaxPayload
                    lMinImpact += lMinPayload
                End If
                .yPayloadType = yPayloadType

                .decDPS = (CDec(lMaxPierce) + CDec(lMaxImpact) + CDec(lMaxFlame) + CDec(lMaxChemical) + CDec(lMaxECM)) / CDec(fROF)
                .decDPS = Math.Max(.decDPS, (.decDPS + CDec(lMaxPierce) + CDec(lMaxImpact) + CDec(lMaxFlame) + CDec(lMaxChemical) + CDec(lMaxECM)) / 2D)

                'Dim decNominal As Decimal = (CDec(lMaxPower / lMaxDPS) * 3D) * .decDPS

                .blCartridgeSize = tpCartridgeSize.PropertyValue
                .blExplosionRadius = tpExplosionRadius.PropertyValue
                .blMaxRange = tpMaxRange.PropertyValue
                .blPierceRatio = tpPierceRatio.PropertyValue
                .blROF = tpROF.PropertyValue


                .lHullTypeID = Me.lHullTypeID

                'If cboBarrelMat.ListIndex > -1 Then
                '    .lMineral1ID = cboBarrelMat.ItemData(cboBarrelMat.ListIndex)
                'Else : .lMineral1ID = -1
                'End If
                'If cboCasingMat.ListIndex > -1 Then
                '    .lMineral2ID = cboCasingMat.ItemData(cboCasingMat.ListIndex)
                'Else : .lMineral2ID = -1
                'End If
                'If cboPayload1Mat.ListIndex > -1 Then
                '    .lMineral3ID = cboPayload1Mat.ItemData(cboPayload1Mat.ListIndex)
                'Else : .lMineral3ID = -1
                'End If
                'If cboPayload2Mat.ListIndex > -1 Then
                '    .lMineral4ID = cboPayload2Mat.ItemData(cboPayload2Mat.ListIndex)
                'Else : .lMineral4ID = -1
                'End If
                'If cboProjectionMat.ListIndex > -1 Then
                '    .lMineral5ID = cboProjectionMat.ItemData(cboProjectionMat.ListIndex)
                'Else : .lMineral5ID = -1
                'End If
                .lMineral1ID = mlMineralIDs(0)
                .lMineral2ID = mlMineralIDs(1)
                .lMineral3ID = mlMineralIDs(2)
                .lMineral4ID = mlMineralIDs(3)
                .lMineral5ID = mlMineralIDs(4)
                .lMineral6ID = -1

                .msMin1Name = "Barrel"
                .msMin2Name = "Casing"
                .msMin3Name = "Payload 1"
                .msMin4Name = "Payload 2"
                .msMin5Name = "Projection"
                .msMin6Name = ""

                Dim lblDesignFlaw As UILabel = CType(MyBase.ParentControl, frmWeaponBuilder).lblDesignFlaw

                If .lMineral1ID < 1 OrElse .lMineral2ID < 1 OrElse .lMineral3ID < 1 OrElse .lMineral4ID < 1 OrElse .lMineral5ID < 1 Then
                    lblDesignFlaw.Caption = "All properties and materials need to be defined."
                    mbIgnoreValueChange = False
                    mbImpossibleDesign = True
                    Return
                End If

                If bAutoFill = True Then
                    .DoBuilderCostPreconfigure(lblDesignFlaw, MyBase.mfrmBuilderCost, tpBarrelMat, tpCasingMat, tpPayload1Mat, tpPayload2Mat, tpProjectionMat, Nothing, ObjectType.eWeaponTech, 0D, False)
                Else
                    .BuilderCostValueChange(lblDesignFlaw, MyBase.mfrmBuilderCost, tpBarrelMat, tpCasingMat, tpPayload1Mat, tpPayload2Mat, tpProjectionMat, Nothing, ObjectType.eWeaponTech, 0D)
                End If

                mbImpossibleDesign = .bImpossibleDesign
                If Me.ParentControl Is Nothing = False Then
                    CType(Me.ParentControl, frmWeaponBuilder).SetWeaponDetailedStats(lMinFlame, lMaxFlame, 0, 0, lMinChemical, lMaxChemical, lMinECM, lMaxECM, lMinImpact, lMaxImpact, lMinPierce, lMaxPierce, fROF)
                End If
            End With

            mbIgnoreValueChange = False
        End Sub

        Private mlMineralIDs(4) As Int32
        Private Sub SetButtonClicked(ByVal lMinIdx As Int32)

            Dim lHullTechID As Int32 = -1
            Dim oTech As New ProjectileTechComputer
            oTech.lHullTypeID = Me.lHullTypeID

            If cboProjectionType.ListIndex > -1 Then
                oTech.yProjectionType = CByte(cboProjectionType.ItemData(cboProjectionType.ListIndex))
            End If
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
                        stmBarrelMat.SetMineralName(sMinName)
                    Case 1
                        stmCasingMat.SetMineralName(sMinName)
                    Case 2
                        stmPayload1Mat.SetMineralName(sMinName)
                    Case 3
                        stmPayload2Mat.SetMineralName(sMinName)
                    Case 4
                        stmProjectionMat.SetMineralName(sMinName)
                End Select
            End If

            CheckForDARequest()
        End Sub

        Public Overrides Sub CheckForDARequest()
            Dim yPayloadVal As Byte = 0
            Dim yProjVal As Byte = 0

            If cboProjectionType.ListIndex > -1 Then
                yProjVal = CByte(cboProjectionType.ItemData(cboProjectionType.ListIndex))
            End If
            If cboPayloadType.ListIndex > -1 Then
                yPayloadVal = CByte(cboPayloadType.ItemData(cboPayloadType.ListIndex))
            End If

            'BuilderCostValueChange(False)
            Dim bRequestDA As Boolean = True
            For X As Int32 = 0 To mlMineralIDs.GetUpperBound(0)
                If mlMineralIDs(X) < 1 Then
                    bRequestDA = False
                    Exit For
                End If
            Next X
            If bRequestDA = True Then CType(Me.ParentControl, frmWeaponBuilder).RequestDAValues(ObjectType.eWeaponTech, WeaponClassType.eProjectile, CByte(lHullTypeID), mlMineralIDs(0), mlMineralIDs(1), mlMineralIDs(2), mlMineralIDs(3), mlMineralIDs(4), -1, yPayloadVal, yProjVal)
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