Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmWeaponBuilder

	'Interface created from Interface Builder
	Private Class fraBomb
        Inherits fraWeaponBase

        Private tpPayloadSize As ctlTechProp
        Private tpAOE As ctlTechProp
        Private tpGuidance As ctlTechProp
        Private tpROF As ctlTechProp
        Private tpRange As ctlTechProp

        Private lblPayloadMat As UILabel
        Private lblPayloadType As UILabel
        Private lblCasingMat As UILabel
        Private lblGuidanceMat As UILabel

        Private tpPayloadMat As ctlTechProp
        Private tpGuidanceMat As ctlTechProp
        Private tpCasingMat As ctlTechProp

        'Private WithEvents cboPayloadMat As UIComboBox
        'Private WithEvents cboGuidanceMat As UIComboBox
        'Private WithEvents cboCasingMat As UIComboBox
        Private WithEvents cboPayloadType As UIComboBox
        Private stmPayloadMat As ctlSetTechMineral
        Private stmGuidanceMat As ctlSetTechMineral
        Private stmCasingMat As ctlSetTechMineral

        Private mbLoading As Boolean = True
        Private mlEntityIndex As Int32 = -1

        Private moTech As WeaponTech = Nothing

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraBomb initial props
            With Me
                .lWindowMetricID = BPMetrics.eWindow.eBombBuilder
                .ControlName = "fraBomb"
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

            ''lblPayloadSize initial props
            'lblPayloadSize = New UILabel(oUILib)
            'With lblPayloadSize
            '    .ControlName = "lblPayloadSize"
            '    .Left = 10
            '    .Top = 10
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Payload Size:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            'End With
            'Me.AddChild(CType(lblPayloadSize, UIControl))

            ''txtPayloadSize initial props
            'txtPayloadSize = New UITextBox(oUILib)
            'With txtPayloadSize
            '    .ControlName = "txtPayloadSize"
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
            '    .MaxLength = 7
            '    .BorderColor = muSettings.InterfaceBorderColor
            'End With
            'Me.AddChild(CType(txtPayloadSize, UIControl))
            tpPayloadSize = New ctlTechProp(oUILib)
            With tpPayloadSize
                .ControlName = "tpPayloadSize"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Payload:"
                .PropertyValue = 0
                .Top = 10
                .Visible = True
                .Width = 430
            End With
            Me.AddChild(CType(tpPayloadSize, UIControl))

            ''lblAOE initial props
            'lblAOE = New UILabel(oUILib)
            'With lblAOE
            '    .ControlName = "lblAOE"
            '    .Left = 10
            '    .Top = 35
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Area of Effect:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            'End With
            'Me.AddChild(CType(lblAOE, UIControl))

            ''lblGuidance initial props
            'lblGuidance = New UILabel(oUILib)
            'With lblGuidance
            '    .ControlName = "lblGuidance"
            '    .Left = 10
            '    .Top = 60
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Self Guidance:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            'End With
            'Me.AddChild(CType(lblGuidance, UIControl))

            ''lblRange initial props
            'lblRange = New UILabel(oUILib)
            'With lblRange
            '    .ControlName = "lblRange"
            '    .Left = 10
            '    .Top = 85
            '    .Width = 120
            '    .Height = 18
            '    .Enabled = True
            '    .Visible = True
            '    .Caption = "Range:"
            '    .ForeColor = muSettings.InterfaceBorderColor
            '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            '    .DrawBackImage = False
            '    .FontFormat = CType(4, DrawTextFormat)
            'End With
            'Me.AddChild(CType(lblRange, UIControl))

            'lblPayloadType initial props
            lblPayloadType = New UILabel(oUILib)
            With lblPayloadType
                .ControlName = "lblPayloadType"
                .Left = 10
                .Top = 150
                .Width = 120
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Payload Type:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblPayloadType, UIControl))

            'lblCasingMat initial props
            lblCasingMat = New UILabel(oUILib)
            With lblCasingMat
                .ControlName = "lblCasingMat"
                .Left = 10
                .Top = 185
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

            'lblGuidanceMat initial props
            lblGuidanceMat = New UILabel(oUILib)
            With lblGuidanceMat
                .ControlName = "lblGuidanceMat"
                .Left = 10
                .Top = 210
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Guidance:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblGuidanceMat, UIControl))

            ''txtAOE initial props
            'txtAOE = New UITextBox(oUILib)
            'With txtAOE
            '    .ControlName = "txtAOE"
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
            '    .MaxLength = 3
            '    .BorderColor = muSettings.InterfaceBorderColor
            'End With
            'Me.AddChild(CType(txtAOE, UIControl))
            tpAOE = New ctlTechProp(oUILib)
            With tpAOE
                .ControlName = "tpAOE"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 255
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Area of Effect:"
                .PropertyValue = 0
                .Top = 35
                .Visible = True
                .Width = 430
            End With
            Me.AddChild(CType(tpAOE, UIControl))

            ''txtGuidance initial props
            'txtGuidance = New UITextBox(oUILib)
            'With txtGuidance
            '    .ControlName = "txtGuidance"
            '    .Left = 135
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
            '    .MaxLength = 3
            '    .BorderColor = muSettings.InterfaceBorderColor
            'End With
            'Me.AddChild(CType(txtGuidance, UIControl))
            tpGuidance = New ctlTechProp(oUILib)
            With tpGuidance
                .ControlName = "tpGuidance"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 255
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Self Guidance:"
                .PropertyValue = 0
                .Top = 60
                .Visible = True
                .Width = 430
            End With
            Me.AddChild(CType(tpGuidance, UIControl))

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
            '    .MaxLength = 6
            '    .BorderColor = muSettings.InterfaceBorderColor
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
                .Top = 110
                .bNoMaxValue = True
                .Visible = True
                .Width = 430
                .SetDivisor(30)
            End With
            Me.AddChild(CType(tpROF, UIControl))

            ''txtRange initial props
            'txtRange = New UITextBox(oUILib)
            'With txtRange
            '    .ControlName = "txtRange"
            '    .Left = 135
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
            '    .MaxLength = 3
            '    .BorderColor = muSettings.InterfaceBorderColor
            'End With
            'Me.AddChild(CType(txtRange, UIControl))
            tpRange = New ctlTechProp(oUILib)
            With tpRange
                .ControlName = "tpRange"
                .Enabled = True
                .Height = 18
                .Left = 10
                .MaxValue = 255
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Range:"
                .PropertyValue = 0
                .Top = 85
                .Visible = True
                .Width = 430
            End With
            Me.AddChild(CType(tpRange, UIControl))

            'lblPayloadMat initial props
            lblPayloadMat = New UILabel(oUILib)
            With lblPayloadMat
                .ControlName = "lblPayloadMat"
                .Left = 10
                .Top = 235
                .Width = 60
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Payload:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblPayloadMat, UIControl))

            tpPayloadMat = New ctlTechProp(oUILib)
            With tpPayloadMat
                .ControlName = "tpPayloadMat"
                .Enabled = True
                .Height = 18
                .Left = lblPayloadMat.Left + lblPayloadMat.Width + 160      '10 left/right, 150
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Payload Mat:"
                .PropertyValue = 0
                .Top = 235
                .bNoMaxValue = True
                .Visible = True
                .Width = Me.Width - .Left - 10
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpPayloadMat, UIControl))
            tpCasingMat = New ctlTechProp(oUILib)
            With tpCasingMat
                .ControlName = "tpCasingMat"
                .Enabled = True
                .Height = 18
                .Left = tpPayloadMat.Left
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Casing Mat:"
                .PropertyValue = 0
                .Top = 185
                .bNoMaxValue = True
                .Visible = True
                .Width = tpPayloadMat.Width
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpCasingMat, UIControl))
            tpGuidanceMat = New ctlTechProp(oUILib)
            With tpGuidanceMat
                .ControlName = "tpGuidanceMat"
                .Enabled = True
                .Height = 18
                .Left = lblPayloadMat.Left + lblPayloadMat.Width + 160      '10 left/right, 150
                .MaxValue = 1000 '0000
                .MinValue = 0
                .PropertyLocked = False
                .PropertyName = "Guidance Mat:"
                .PropertyValue = 0
                .Top = 210
                .bNoMaxValue = True
                .Visible = True
                .Width = tpPayloadMat.Width
                .SetNoPropDisplay(True)
                .SetPercOfPaymentVisibility(True)
            End With
            Me.AddChild(CType(tpGuidanceMat, UIControl))

            ''cboPayloadMat initial props
            'cboPayloadMat = New UIComboBox(oUILib)
            'With cboPayloadMat
            '    .ControlName = "cboPayloadMat"
            '    .Left = lblPayloadMat.Left + lblPayloadMat.Width + 5
            '    .Top = 235
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
            'Me.AddChild(CType(cboPayloadMat, UIControl))
            stmPayloadMat = New ctlSetTechMineral(oUILib)
            With stmPayloadMat
                .ControlName = "stmPayloadMat"
                .Left = lblPayloadMat.Left + lblPayloadMat.Width + 5
                .Top = 235
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 2
            End With
            Me.AddChild(CType(stmPayloadMat, UIControl))

            ''cboGuidanceMat initial props
            'cboGuidanceMat = New UIComboBox(oUILib)
            'With cboGuidanceMat
            '    .ControlName = "cboGuidanceMat"
            '    .Left = cboPayloadMat.Left
            '    .Top = 210
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
            'Me.AddChild(CType(cboGuidanceMat, UIControl))
            stmGuidanceMat = New ctlSetTechMineral(oUILib)
            With stmGuidanceMat
                .ControlName = "stmGuidanceMat"
                .Left = stmPayloadMat.Left
                .Top = 210
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 1
            End With
            Me.AddChild(CType(stmGuidanceMat, UIControl))

            ''cboCasingMat initial props
            'cboCasingMat = New UIComboBox(oUILib)
            'With cboCasingMat
            '    .ControlName = "cboCasingMat"
            '    .Left = cboPayloadMat.Left
            '    .Top = 185
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
            'Me.AddChild(CType(cboCasingMat, UIControl))
            stmCasingMat = New ctlSetTechMineral(oUILib)
            With stmCasingMat
                .ControlName = "stmCasingMat"
                .Left = stmPayloadMat.Left
                .Top = 185
                .Width = 150
                .Height = 18
                .Enabled = True
                .Visible = True
                .mlMineralIndex = 0
            End With
            Me.AddChild(CType(stmCasingMat, UIControl))

            'cboPayloadType initial props
            cboPayloadType = New UIComboBox(oUILib)
            With cboPayloadType
                .ControlName = "cboPayloadType"
                .Left = 135
                .Top = 150
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
            End With
            Me.AddChild(CType(cboPayloadType, UIControl))

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
            'End With
            'Me.AddChild(CType(lblROF, UIControl))

            AddHandler tpAOE.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpCasingMat.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpGuidance.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpGuidanceMat.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpPayloadMat.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpPayloadSize.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpRange.PropertyValueChanged, AddressOf tp_ValueChanged
            AddHandler tpROF.PropertyValueChanged, AddressOf tp_ValueChanged

            AddHandler stmCasingMat.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmGuidanceMat.SetButtonClicked, AddressOf SetButtonClicked
            AddHandler stmPayloadMat.SetButtonClicked, AddressOf SetButtonClicked

            FillValues()

            mbLoading = False
        End Sub

        Private Sub FillValues()
            'Dim lSorted() As Int32 = GetSortedMineralIdxArray(False)

            'cboPayloadMat.Clear() : cboGuidanceMat.Clear() : cboCasingMat.Clear()
            cboPayloadType.Clear()
            ''Now... loop through our minerals
            'If lSorted Is Nothing = False Then
            '    For X As Int32 = 0 To lSorted.GetUpperBound(0)
            '        If lSorted(X) <> -1 Then
            '            cboPayloadMat.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboPayloadMat.ItemData(cboPayloadMat.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboGuidanceMat.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboGuidanceMat.ItemData(cboGuidanceMat.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '            cboCasingMat.AddItem(goMinerals(lSorted(X)).MineralName)
            '            cboCasingMat.ItemData(cboCasingMat.NewIndex) = goMinerals(lSorted(X)).ObjectID
            '        End If
            '    Next X
            'End If

            cboPayloadType.AddItem("Explosive")
            cboPayloadType.ItemData(cboPayloadType.NewIndex) = 0
            Dim lPayloadTypeAvail As Int32 = 0
            If goCurrentPlayer Is Nothing = False Then
                lPayloadTypeAvail = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBombPayloadType)
            End If
            If (lPayloadTypeAvail And 1) <> 0 Then
                cboPayloadType.AddItem("Toxic - Kills Colonists")
                cboPayloadType.ItemData(cboPayloadType.NewIndex) = 1
            End If
            If (lPayloadTypeAvail And 2) <> 0 Then
                cboPayloadType.AddItem("Concussion - Decrease Morale")
                cboPayloadType.ItemData(cboPayloadType.NewIndex) = 2
            End If
            If (lPayloadTypeAvail And 4) <> 0 Then
                cboPayloadType.AddItem("EMP - Leaves targets powerless")
                cboPayloadType.ItemData(cboPayloadType.NewIndex) = 3
            End If
            cboPayloadType.ListIndex = 0

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

            Dim yMsg(134) As Byte '57) As Byte
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
            yMsg(lPos) = WeaponClassType.eBomb : lPos += 1
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
                If tpCasingMat.PropertyLocked = True Then lTempMin = CInt(tpCasingMat.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpGuidanceMat.PropertyLocked = True Then lTempMin = CInt(tpGuidanceMat.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                If tpPayloadMat.PropertyLocked = True Then lTempMin = CInt(tpPayloadMat.PropertyValue) Else lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4
                lTempMin = -1
                System.BitConverter.GetBytes(lTempMin).CopyTo(yMsg, lPos) : lPos += 4

                yMsg(lPos) = CByte(Me.lHullTypeID) : lPos += 1
            End With

            'Now... for the data specific this weapon type...
            Dim lPayloadSize As Int32 = CInt(tpPayloadSize.PropertyValue)
            Dim yAOE As Byte = CByte(tpAOE.PropertyValue)
            Dim yGuidance As Byte = CByte(tpGuidance.PropertyValue)
            Dim fROF As Single = CSng(tpROF.PropertyValue)
            'fROF *= 30.0F
            Dim iROF As Int16 = CShort(fROF)
            Dim yRange As Byte = CByte(tpRange.PropertyValue)
            Dim yPayloadType As Byte = CByte(cboPayloadType.ItemData(cboPayloadType.ListIndex))
            Dim lPayloadID As Int32 = mlMineralIDs(2)
            Dim lGuidanceID As Int32 = mlMineralIDs(1)
            Dim lCasingID As Int32 = mlMineralIDs(0)

            System.BitConverter.GetBytes(lPayloadSize).CopyTo(yMsg, lPos) : lPos += 4
            yMsg(lPos) = yAOE : lPos += 1
            yMsg(lPos) = yGuidance : lPos += 1
            System.BitConverter.GetBytes(iROF).CopyTo(yMsg, lPos) : lPos += 2
            yMsg(lPos) = yRange : lPos += 1
            yMsg(lPos) = yPayloadType : lPos += 1
            System.BitConverter.GetBytes(lPayloadID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lGuidanceID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(lCasingID).CopyTo(yMsg, lPos) : lPos += 4

            Dim sTemp As String = Mid$(sName, 1, 20)
            System.Text.ASCIIEncoding.ASCII.GetBytes(sTemp).CopyTo(yMsg, lPos) : lPos += 20

            MyBase.moUILib.SendMsgToPrimary(yMsg)

            Return True
        End Function

        Public Overrides Function ValidateData() As Boolean
            If cboPayloadType.ListIndex < 0 Then
                MyBase.moUILib.AddNotification("You must select a Payload Type.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            For X As Int32 = 0 To mlMineralIDs.GetUpperBound(0)
                If mlMineralIDs(X) < 1 Then
                    MyBase.moUILib.AddNotification("You must select a Material for each entry.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
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
            'If IsNumeric(txtPayloadSize.Caption) = False OrElse Val(txtPayloadSize.Caption) < 1 OrElse txtPayloadSize.Caption.Contains(".") = True OrElse Val(txtPayloadSize.Caption) > Int32.MaxValue Then
            '    MyBase.moUILib.AddNotification("Payload Size must be a non-negative whole number greater than 0.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            '    Return False
            'End If
            Dim lMaxVal As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBombMaxAOE)
            If tpAOE.PropertyValue > lMaxVal Then
                MyBase.moUILib.AddNotification("Area of Effect must be a whole number between 0 and " & lMaxVal & ".", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            lMaxVal = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBombMaxGuidance)
            If tpGuidance.PropertyValue > lMaxVal Then
                MyBase.moUILib.AddNotification("Self Guidance must be a whole number between 0 and " & lMaxVal & ".", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            'Dim fROF As Single = CSng(tpROF.PropertyValue)
            lMaxVal = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBombMinROF)
            'fROF *= 30.0F
            If tpROF.PropertyValue < lMaxVal Then
                MyBase.moUILib.AddNotification("Please enter a ROF of at least " & (lMaxVal / 30.0F).ToString("#,###.#0") & " seconds.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            ElseIf tpROF.PropertyValue > Int16.MaxValue Then
                MyBase.moUILib.AddNotification("Maximum value for ROF is: " & (Int16.MaxValue / 30.0F).ToString("#,###.#0"), Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            lMaxVal = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBombMaxRange)
            If tpRange.PropertyValue > lMaxVal Then
                MyBase.moUILib.AddNotification("Range must be a whole number between 0 and " & lMaxVal & ".", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            Return True
        End Function

        Public Overrides Sub ViewResults(ByRef oTech As WeaponTech, ByVal lProdFactor As Integer)
            If oTech Is Nothing Then Return

            moTech = oTech
            With CType(oTech, BombWeaponTech)
                mbIgnoreValueChange = True
                cboPayloadType.FindComboItemData(.yPayloadType)
                mbLoading = True

                MyBase.lHullTypeID = .HullTypeID

                If tpROF.MaxValue < .iROF Then tpROF.MaxValue = .iROF
                tpROF.PropertyValue = .iROF '(.iROF / 30.0F)
                If .iROF > 0 Then tpROF.PropertyLocked = True
                tpPayloadSize.PropertyValue = .lPayloadSize
                If .lPayloadSize > 0 Then tpPayloadSize.PropertyLocked = True
                tpAOE.PropertyValue = .yAOE
                If .yAOE > 0 Then tpAOE.PropertyLocked = True
                tpGuidance.PropertyValue = .yGuidance
                If .yGuidance > 0 Then tpGuidance.PropertyLocked = True
                tpRange.PropertyValue = .yRange
                If .yRange > 0 Then tpRange.PropertyLocked = True

                mlSelectedMineralIdx = 0 : Mineral_Selected(.lCasingMat)
                mlSelectedMineralIdx = 1 : Mineral_Selected(.lGuidanceMat)
                mlSelectedMineralIdx = 2 : Mineral_Selected(.lPayloadMat)
                'Me.cboCasingMat.FindComboItemData(.lCasingMat)
                'Me.cboGuidanceMat.FindComboItemData(.lGuidanceMat)
                'Me.cboPayloadMat.FindComboItemData(.lPayloadMat)

                tpCasingMat.PropertyValue = .lSpecifiedMin1
                tpGuidanceMat.PropertyValue = .lSpecifiedMin2
                tpPayloadMat.PropertyValue = .lSpecifiedMin3
                If .lSpecifiedMin1 <> -1 Then tpCasingMat.PropertyLocked = True
                If .lSpecifiedMin2 <> -1 Then tpGuidanceMat.PropertyLocked = True
                If .lSpecifiedMin3 <> -1 Then tpPayloadMat.PropertyLocked = True

                MyBase.mfrmBuilderCost.SetAndLockValues(oTech)

                mbLoading = False

                mbIgnoreValueChange = False
                BuilderCostValueChange(False)
            End With
        End Sub

        Private Sub SetCboHelpers()
            If goCurrentPlayer Is Nothing Then Return
            Dim oSB As New System.Text.StringBuilder()

            Dim bDensity As Boolean = goCurrentPlayer.PlayerKnowsProperty("DENSITY")
            Dim bMalleable As Boolean = goCurrentPlayer.PlayerKnowsProperty("MALLEABLE")
            Dim bCompress As Boolean = goCurrentPlayer.PlayerKnowsProperty("COMPRESSIBILITY")
            Dim bThermExp As Boolean = goCurrentPlayer.PlayerKnowsProperty("THERMAL EXPANSION")
            Dim bThermCond As Boolean = goCurrentPlayer.PlayerKnowsProperty("THERMAL CONDUCTION")
            Dim bTempSens As Boolean = goCurrentPlayer.PlayerKnowsProperty("TEMPERATURE SENSITIVITY")
            Dim bChemReact As Boolean = goCurrentPlayer.PlayerKnowsProperty("CHEMICAL REACTANCE")
            Dim bBoilingPt As Boolean = goCurrentPlayer.PlayerKnowsProperty("BOILING POINT")
            Dim bCombust As Boolean = goCurrentPlayer.PlayerKnowsProperty("COMBUSTIVENESS")
            Dim bElectRes As Boolean = goCurrentPlayer.PlayerKnowsProperty("ELECTRICAL RESISTANCE")
            Dim bMagProd As Boolean = goCurrentPlayer.PlayerKnowsProperty("MAGNETIC PRODUCTION")
            Dim bSuperC As Boolean = goCurrentPlayer.PlayerKnowsProperty("SUPERCONDUCTIVE POINT")

            oSB.Length = 0
            If bDensity = True Then oSB.Append("High Density materials with low ") Else oSB.Append("Materials with low ")
            If bCompress = True Then oSB.Append("Compressibility")
            If bThermCond = True Then
                If oSB.ToString.EndsWith(" ") = False Then oSB.Append(", ")
                oSB.Append("Thermal Conduction")
            End If
            If bThermExp = True Then
                If oSB.ToString.EndsWith(" ") = False Then oSB.Append(", ")
                oSB.Append("Thermal Expansion")
            End If
            If bCombust = True Then
                If oSB.ToString.EndsWith(" ") = False Then oSB.Append(", ")
                oSB.Append("Combustion")
            End If
            If oSB.ToString.EndsWith(" ") = False Then oSB.Append(" ")
            oSB.Append("are best for this material")
            stmCasingMat.ToolTipText = oSB.ToString

            oSB.Length = 0
            If bMalleable = True AndAlso bSuperC = True Then
                oSB.Append("Malleable materials with high Superconductive properties and low ")
            ElseIf bSuperC = True Then
                oSB.Append("Materials with high Superconductive properties and low ")
            ElseIf bMalleable = True Then
                oSB.Append("Malleable materials with low ")
            Else
                oSB.Append("Materials with low ")
            End If
            If bElectRes = True And bMagProd = True Then
                oSB.Append("Electrical Resistance and Magnetic Production are best for this material.")
            ElseIf bElectRes = True Then
                oSB.Append("Electrical Resistance work good here.")
            ElseIf bMagProd = True Then
                oSB.Append("Magnetic Production are valuable here.")
            Else
                oSB.Append("electrical properties work best.")
            End If
            stmGuidanceMat.ToolTipText = oSB.ToString


            oSB.Length = 0
            If cboPayloadType.ListIndex > -1 Then
                If bDensity = True Then
                    oSB.Append("Low Density materials ")
                Else : oSB.Append("Materials ")
                End If
                If bCompress = True Then
                    oSB.Append("with high compressibility and ")
                Else
                    oSB.Append("with high ")
                End If
                Select Case cboPayloadType.ItemData(cboPayloadType.ListIndex)
                    Case 0
                        oSB.Append("Combustiveness")
                    Case 1
                        oSB.Append("Toxicity")
                    Case 2
                        oSB.Append("Thermal Expansion")
                    Case 3
                        oSB.Append("Magnetic Production")
                End Select
                oSB.Append(" are best for the payload material.")
            Else
                oSB.Append("Select a payload type to see the best properties.")
            End If
            stmPayloadMat.ToolTipText = oSB.ToString

            Dim lMaxRange As Int32 = 30
            Dim lMaxAOE As Int32 = 10
            Dim lMaxGuidance As Int32 = 10
            Dim lMinROF As Int32 = 900
            If goCurrentPlayer Is Nothing = False Then
                lMaxRange = Math.Min(255, goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBombMaxRange))
                lMaxAOE = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBombMaxAOE)
                lMaxGuidance = Math.Min(255, goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBombMaxGuidance))
                lMinROF = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eBombMinROF)
            End If
            tpRange.MinValue = 10
            tpRange.MaxValue = lMaxRange
            tpAOE.MaxValue = lMaxAOE
            tpGuidance.MaxValue = lMaxGuidance
            tpROF.MinValue = lMinROF
            tpROF.PropertyValue = tpROF.MinValue
            tpROF.MaxValue = 2000
            tpROF.blAbsoluteMaximum = 30000

            tpPayloadSize.MinValue = 2 : tpPayloadSize.MaxValue = 90000 '3000

            tpPayloadSize.ToolTipText = "The size of the bomb's payload inside of the casing. Directly impacts damage effect as well as size of weapon."
            tpAOE.ToolTipText = "Area effect radius. For orbital bombardment, this radius is amplified. Max value is " & lMaxAOE
            tpGuidance.ToolTipText = "Guidance of the bomb for making the bomb 'smarter'. Max value is " & lMaxGuidance
            tpROF.ToolTipText = "Rate of fire at which the bombs are dropped. Minimum value is " & (lMinROF / 30.0F).ToString("####.##") & " seconds."
            tpRange.ToolTipText = "Range of the bomb. Max value is " & lMaxRange


            tpGuidance.SetToInitialDefault()
            tpRange.SetToInitialDefault()
            tpROF.PropertyValue = 1800

        End Sub

        Private Sub cboPayloadType_ItemSelected(ByVal lItemIndex As Integer) Handles cboPayloadType.ItemSelected
            If mbLoading = True Then Return
            SetCboHelpers()
            CheckForDARequest()
        End Sub

        'Private Sub COMBO_DropDownExpanded(ByVal sComboBoxName As String) Handles cboCasingMat.DropDownExpanded, cboGuidanceMat.DropDownExpanded, cboPayloadMat.DropDownExpanded
        '    If sComboBoxName = "" Then Return

        '    Dim oTech As New BombTechComputer
        '    oTech.lHullTypeID = Me.lHullTypeID
        '    If cboPayloadType.ListIndex > -1 Then
        '        oTech.yPayloadType = CByte(cboPayloadType.ItemData(cboPayloadType.ListIndex))
        '    End If

        '    Dim lMinIdx As Int32 = -1
        '    Select Case sComboBoxName
        '        Case cboCasingMat.ControlName
        '            lMinIdx = 0
        '        Case cboGuidanceMat.ControlName
        '            lMinIdx = 1
        '        Case cboPayloadMat.ControlName
        '            lMinIdx = 2
        '    End Select
        '    If lMinIdx = -1 Then Return
        '    oTech.MineralCBOExpanded(lMinIdx, ObjectType.eWeaponTech)
        'End Sub
        'Try
        '    Dim ofrmMin As frmMinDetail = CType(MyBase.moUILib.GetWindow("frmMinDetail"), frmMinDetail)
        '    If ofrmMin Is Nothing Then Return
        '    Select Case sComboBoxName
        '        Case cboCasingMat.ControlName
        '            ofrmMin.ClearHighlights()
        '            For X As Int32 = 0 To glMineralPropertyUB
        '                If glMineralPropertyIdx(X) <> -1 Then
        '                    Select Case goMineralProperty(X).MineralPropertyName.ToUpper
        '                        Case "DENSITY"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 2, 8)
        '                        Case "COMPRESSIBILITY"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
        '                        Case "THERMAL CONDUCTION"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
        '                        Case "THERMAL EXPANSION"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
        '                        Case "COMBUSTIVENESS"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
        '                    End Select
        '                End If
        '            Next X
        '        Case cboGuidanceMat.ControlName
        '            ofrmMin.ClearHighlights()
        '            For X As Int32 = 0 To glMineralPropertyUB
        '                If glMineralPropertyIdx(X) <> -1 Then
        '                    Select Case goMineralProperty(X).MineralPropertyName.ToUpper
        '                        Case "ELECTRICAL RESISTANCE"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
        '                        Case "MALLEABLE"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 4, 6)
        '                        Case "MAGNETIC PRODUCTION"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
        '                        Case "SUPERCONDUCTIVE POINT"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 2, 7)
        '                    End Select
        '                End If
        '            Next X
        '        Case cboPayloadMat.ControlName
        '            ofrmMin.ClearHighlights()

        '            Dim sProperty As String = ""
        '            If cboPayloadType.ListIndex > -1 Then
        '                Select Case cboPayloadType.ItemData(cboPayloadType.ListIndex)
        '                    Case 0
        '                        sProperty = "COMBUSTIVENESS"
        '                    Case 1
        '                        sProperty = "TOXICITY"
        '                    Case 2
        '                        sProperty = "THERMAL EXPANSION"
        '                    Case 3
        '                        sProperty = "MAGNETIC PRODUCTION"
        '                End Select
        '            End If

        '            For X As Int32 = 0 To glMineralPropertyUB
        '                If glMineralPropertyIdx(X) <> -1 Then
        '                    Select Case goMineralProperty(X).MineralPropertyName.ToUpper
        '                        Case "DENSITY"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 3, 0)
        '                        Case "COMPRESSIBILITY"
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 2, 8)
        '                        Case sProperty
        '                            ofrmMin.HighlightProperty(goMineralProperty(X).ObjectID, 2, 10)
        '                    End Select
        '                End If
        '            Next X
        '    End Select
        'Catch
        'End Try

        Private Sub fraBomb_OnRenderEnd() Handles Me.OnRenderEnd
            'Ok, first, is moTech nothing?
            If moTech Is Nothing = False Then
                'Ok, now, do the values used on the form currently match the tech?
                Dim bChanged As Boolean = False
                With CType(moTech, BombWeaponTech)
                    bChanged = tpAOE.PropertyValue <> .yAOE OrElse tpGuidance.PropertyValue <> .yGuidance OrElse _
                     tpPayloadSize.PropertyValue <> .lPayloadSize OrElse tpRange.PropertyValue <> .yRange

                    If bChanged = False Then
                        bChanged = tpROF.PropertyValue <> .iROF
                    End If
                    If bChanged = False Then

                        Dim lCaseID As Int32 = mlMineralIDs(0)
                        Dim lGuidance As Int32 = mlMineralIDs(1)
                        Dim lPayload As Int32 = mlMineralIDs(2)
                        Dim lPType As Int32 = -1

                        'If cboCasingMat.ListIndex <> -1 Then lCaseID = cboCasingMat.ItemData(cboCasingMat.ListIndex)
                        'If cboGuidanceMat.ListIndex > -1 Then lGuidance = cboGuidanceMat.ItemData(cboGuidanceMat.ListIndex)
                        'If cboPayloadMat.ListIndex > -1 Then lPayload = cboPayloadMat.ItemData(cboPayloadMat.ListIndex)
                        If cboPayloadType.ListIndex <> -1 Then lPType = cboPayloadType.ItemData(cboPayloadType.ListIndex)

                        bChanged = .lPayloadMat <> lPayload OrElse .lGuidanceMat <> lGuidance OrElse .lCasingMat <> lCaseID OrElse _
                            (tpCasingMat.PropertyLocked <> (.lSpecifiedMin1 <> -1)) OrElse _
                            (tpGuidance.PropertyLocked <> (.lSpecifiedMin2 <> -1)) OrElse (tpPayloadMat.PropertyLocked <> (.lSpecifiedMin3 <> -1))

                        If bChanged = False Then

                            bChanged = (.lSpecifiedMin1 <> -1 AndAlso .lSpecifiedMin1 <> tpCasingMat.PropertyValue) OrElse _
                                        (.lSpecifiedMin2 <> -1 AndAlso .lSpecifiedMin2 <> tpGuidance.PropertyValue) OrElse _
                                        (.lSpecifiedMin3 <> -1 AndAlso .lSpecifiedMin3 <> tpPayloadMat.PropertyValue)
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

        Private Sub tp_ValueChanged(ByVal sPropName As String, ByRef ctl As ctlTechProp)
            tpROF.MaxValue = Math.Min(Int16.MaxValue, Math.Max(tpROF.MaxValue, tpROF.PropertyValue + 500))
            BuilderCostValueChange(False)
        End Sub 

        Public Overrides Sub BuilderCostValueChange(ByVal bAutoFill As Boolean)
            If mbIgnoreValueChange = True Then Return
            If Me.ParentControl Is Nothing Then Return
            mbIgnoreValueChange = True
            mbImpossibleDesign = True
            Dim oBombTechComputer As New BombTechComputer
            With oBombTechComputer
                Dim fROF As Single = Math.Max(1, tpROF.PropertyValue) / 30.0F
                Dim lImpactMin As Int32 = 0
                Dim lImpactMax As Int32 = 0
                Dim lFlameMin As Int32 = 0
                Dim lFlameMax As Int32 = 0
                If cboPayloadType.ListIndex = 0 Then
                    Dim lHalfPayload As Int32 = CInt(tpPayloadSize.PropertyValue \ 2)
                    Dim lQtrPayload As Int32 = lHalfPayload \ 2
                    lImpactMax = lHalfPayload : lImpactMin = lQtrPayload
                    lFlameMax = lHalfPayload : lFlameMin = lQtrPayload
                End If

                If cboPayloadType.ListIndex > -1 Then
                    .yPayloadType = CByte(cboPayloadType.ItemData(cboPayloadType.ListIndex))
                End If

                With CType(Me.ParentControl, frmWeaponBuilder)
                    oBombTechComputer.SetMinDAValues(.ml_Min1DA, .ml_Min2DA, .ml_Min3DA, .ml_Min4DA, .ml_Min5DA, .ml_Min6DA)
                End With

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


                .decDPS = CDec((CDec(lImpactMax) + CDec(lFlameMax)) / fROF)
                .decDPS = Math.Max(.decDPS, (.decDPS + CDec(lImpactMax) + CDec(lFlameMax)) / 2D)
                'Dim decNominal As Decimal = (CDec(lMaxPower / lMaxDPS) * 3D) * .decDPS

                .blAOE = tpAOE.PropertyValue
                .blGuidance = tpGuidance.PropertyValue
                .blPayloadSize = tpPayloadSize.PropertyValue
                .blRange = tpRange.PropertyValue
                .blROF = tpROF.PropertyValue

                .lHullTypeID = Me.lHullTypeID

                'If cboCasingMat.ListIndex > -1 Then
                '    .lMineral1ID = cboCasingMat.ItemData(cboCasingMat.ListIndex)
                'Else : .lMineral1ID = -1
                'End If
                'If cboGuidanceMat.ListIndex > -1 Then
                '    .lMineral2ID = cboGuidanceMat.ItemData(cboGuidanceMat.ListIndex)
                'Else : .lMineral2ID = -1
                'End If
                'If cboPayloadMat.ListIndex > -1 Then
                '    .lMineral3ID = cboPayloadMat.ItemData(cboPayloadMat.ListIndex)
                'Else : .lMineral3ID = -1
                'End If
                .lMineral1ID = mlMineralIDs(0)
                .lMineral2ID = mlMineralIDs(1)
                .lMineral3ID = mlMineralIDs(2)
                .lMineral4ID = -1
                .lMineral5ID = -1
                .lMineral6ID = -1

                .msMin1Name = "Casing"
                .msMin2Name = "Guidance"
                .msMin3Name = "Payload"
                .msMin4Name = ""
                .msMin5Name = ""
                .msMin6Name = ""

                Dim lblDesignFlaw As UILabel = CType(MyBase.ParentControl, frmWeaponBuilder).lblDesignFlaw

                If .lMineral1ID < 1 OrElse .lMineral2ID < 1 OrElse .lMineral3ID < 1 Then
                    lblDesignFlaw.Caption = "All properties and materials need to be defined."
                    mbIgnoreValueChange = False
                    mbImpossibleDesign = True
                    Return
                End If

                If bAutoFill = True Then
                    .DoBuilderCostPreconfigure(lblDesignFlaw, MyBase.mfrmBuilderCost, tpCasingMat, tpGuidanceMat, tpPayloadMat, Nothing, Nothing, Nothing, ObjectType.eWeaponTech, 0D, False)
                Else
                    .BuilderCostValueChange(lblDesignFlaw, MyBase.mfrmBuilderCost, tpCasingMat, tpGuidanceMat, tpPayloadMat, Nothing, Nothing, Nothing, ObjectType.eWeaponTech, 0D)
                End If

                mbImpossibleDesign = .bImpossibleDesign
                If Me.ParentControl Is Nothing = False Then
                    CType(Me.ParentControl, frmWeaponBuilder).SetWeaponDetailedStats(lFlameMin, lFlameMax, 0, 0, 0, 0, 0, 0, lImpactMin, lImpactMax, 0, 0, fROF)
                End If
            End With

            mbIgnoreValueChange = False
        End Sub

        Private mlMineralIDs(2) As Int32
        Private Sub SetButtonClicked(ByVal lMinIdx As Int32)

            Dim oTech As New BombTechComputer
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
                        stmCasingMat.SetMineralName(sMinName)
                    Case 1
                        stmGuidanceMat.SetMineralName(sMinName)
                    Case 2
                        stmPayloadMat.SetMineralName(sMinName)
                End Select
            End If

            CheckForDARequest()
        End Sub

        Public Overrides Sub CheckForDARequest()
            Dim yPayload As Byte = 0
            If cboPayloadType Is Nothing = False AndAlso cboPayloadType.ListIndex > -1 Then
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
            If bRequestDA = True Then CType(Me.ParentControl, frmWeaponBuilder).RequestDAValues(ObjectType.eWeaponTech, WeaponClassType.eBomb, CByte(lHullTypeID), mlMineralIDs(0), mlMineralIDs(1), mlMineralIDs(2), -1, -1, -1, yPayload, 0)
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
