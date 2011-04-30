Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

#Region "  Old Code?  "
''Research Main
'Public Class frmResearchMain
'    Inherits UIWindow

'    Private WithEvents btnAlloy As UIButton
'    Private lblSelect As UILabel
'    Private lblFilterSetting As UILabel
'    Private WithEvents btnMineral As UIButton
'    Private WithEvents btnArmor As UIButton
'    Private WithEvents btnEngine As UIButton
'    Private WithEvents btnRadar As UIButton
'    Private WithEvents btnShield As UIButton
'    Private WithEvents btnWeapon As UIButton
'    Private WithEvents btnSpecials As UIButton
'    Private WithEvents btnHull As UIButton
'    Private WithEvents btnPrototype As UIButton
'    Private lnDiv1 As UILine
'    Private lblMaterials As UILabel
'    Private lnDiv2 As UILine
'    Private lblComponent As UILabel
'    Private lnDiv3 As UILine
'    Private lblDesigns As UILabel
'    Private lnDiv4 As UILine
'    Private lblResExist As UILabel
'    Private lnDiv5 As UILine
'    Private WithEvents lstTechs As UIListBox
'    Private WithEvents btnClose As UIButton
'    Private lnDiv6 As UILine
'    Private WithEvents btnResearch As UIButton
'    Private lblComponentName As UILabel
'    Private lblComponentType As UILabel
'    Private lblResPhase As UILabel
'    Private WithEvents btnViewResults As UIButton

'    Private WithEvents chkFilterArchived As UICheckBox
'    Private WithEvents chkSetFilter As UICheckBox

'    Private WithEvents cboFilter As UIComboBox

'    Private WithEvents btnHelp As UIButton

'    Private mlLastCheck As Int32

'    Private mbIgnoreSetFilter As Boolean = False

'    Private mbLoading As Boolean = True

'    Public Sub New(ByRef oUILib As UILib)
'        MyBase.New(oUILib)

'        'frmResearchMain initial props
'        With Me
'            .lWindowMetricID = BPMetrics.eWindow.eResearchMain
'            .ControlName = "frmResearchMain"
'            .Left = 158
'            .Top = 144
'            .Width = 375
'            .Height = 511 '411
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
'            .BorderColor = muSettings.InterfaceBorderColor
'            '.FillColor = System.Drawing.Color.FromArgb(128, 64, 128, 192)
'            .FillColor = muSettings.InterfaceFillColor
'            .FullScreen = False
'            .mbAcceptReprocessEvents = True
'        End With

'        'btnAlloy initial props
'        btnAlloy = New UIButton(oUILib)
'        With btnAlloy
'            .ControlName = "btnAlloy"
'            .Left = 250
'            .Top = 302 '202
'            .Width = 119
'            .Height = 23
'            .Enabled = True
'            .Visible = True
'            .Caption = "Alloy"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
'        End With
'        Me.AddChild(CType(btnAlloy, UIControl))

'        'lblSelect initial props
'        lblSelect = New UILabel(oUILib)
'        With lblSelect
'            .ControlName = "lblSelect"
'            .Left = 5
'            .Top = 275 '175
'            .Width = 317
'            .Height = 16
'            .Enabled = True
'            .Visible = True
'            .Caption = "OR SELECT A RESEARCH TYPE TO DESIGN"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblSelect, UIControl))

'        'lblFilterSetting initial props
'        lblFilterSetting = New UILabel(oUILib)
'        With lblFilterSetting
'            .ControlName = "lblFilterSetting"
'            .Left = 10
'            .Top = 28
'            .Width = 90
'            .Height = 16
'            .Enabled = True
'            .Visible = True
'            .Caption = "Filter Setting:"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblFilterSetting, UIControl))

'        'btnMineral initial props
'        btnMineral = New UIButton(oUILib)
'        With btnMineral
'            .ControlName = "btnMineral"
'            .Left = 105
'            .Top = 302 '202
'            .Width = 119
'            .Height = 23
'            .Enabled = True
'            .Visible = True
'            .Caption = "Mineral"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
'        End With
'        Me.AddChild(CType(btnMineral, UIControl))

'        'btnArmor initial props
'        btnArmor = New UIButton(oUILib)
'        With btnArmor
'            .ControlName = "btnArmor"
'            .Left = 105
'            .Top = 340 '240
'            .Width = 119
'            .Height = 23
'            .Enabled = True
'            .Visible = True
'            .Caption = "Armor"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
'        End With
'        Me.AddChild(CType(btnArmor, UIControl))

'        'btnEngine initial props
'        btnEngine = New UIButton(oUILib)
'        With btnEngine
'            .ControlName = "btnEngine"
'            .Left = 105
'            .Top = 370 '270
'            .Width = 119
'            .Height = 23
'            .Enabled = True
'            .Visible = True
'            .Caption = "Engine"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
'        End With
'        Me.AddChild(CType(btnEngine, UIControl))

'        'btnRadar initial props
'        btnRadar = New UIButton(oUILib)
'        With btnRadar
'            .ControlName = "btnRadar"
'            .Left = 250
'            .Top = 340 ' 240
'            .Width = 119
'            .Height = 23
'            .Enabled = True
'            .Visible = True
'            .Caption = "Radar"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
'        End With
'        Me.AddChild(CType(btnRadar, UIControl))

'        'btnShield initial props
'        btnShield = New UIButton(oUILib)
'        With btnShield
'            .ControlName = "btnShield"
'            .Left = 250
'            .Top = 370 '270
'            .Width = 119
'            .Height = 23
'            .Enabled = True
'            .Visible = True
'            .Caption = "Shield"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
'        End With
'        Me.AddChild(CType(btnShield, UIControl))

'        'btnSpecials initial props
'        btnSpecials = New UIButton(oUILib)
'        With btnSpecials
'            .ControlName = "btnSpecials"
'            .Left = 105
'            '.Left = 250
'            .Top = 400 '300
'            .Width = 119
'            .Height = 23
'            .Enabled = True
'            .Visible = True
'            .Caption = "Specials"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
'        End With
'        Me.AddChild(CType(btnSpecials, UIControl))

'        'btnWeapon initial props
'        btnWeapon = New UIButton(oUILib)
'        With btnWeapon
'            .ControlName = "btnWeapon"
'            .Left = 250
'            .Top = 400 ' 300
'            .Width = 119
'            .Height = 23
'            .Enabled = True
'            .Visible = True
'            .Caption = "Weapon"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
'        End With
'        Me.AddChild(CType(btnWeapon, UIControl))

'        'btnHull initial props
'        btnHull = New UIButton(oUILib)
'        With btnHull
'            .ControlName = "btnHull"
'            .Left = 105
'            .Top = 440 '340
'            .Width = 119
'            .Height = 23
'            .Enabled = True
'            .Visible = True
'            .Caption = "Hull"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
'        End With
'        Me.AddChild(CType(btnHull, UIControl))

'        'btnPrototype initial props
'        btnPrototype = New UIButton(oUILib)
'        With btnPrototype
'            .ControlName = "btnPrototype"
'            .Left = 250
'            .Top = 440 ' 340
'            .Width = 119
'            .Height = 23
'            .Enabled = True
'            .Visible = True
'            .Caption = "Prototype"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
'        End With
'        Me.AddChild(CType(btnPrototype, UIControl))

'        'lnDiv1 initial props
'        lnDiv1 = New UILine(oUILib)
'        With lnDiv1
'            .ControlName = "lnDiv1"
'            .Left = 1
'            .Top = 294 '194
'            .Width = 374
'            .Height = 0
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(lnDiv1, UIControl))

'        'lblMaterials initial props
'        lblMaterials = New UILabel(oUILib)
'        With lblMaterials
'            .ControlName = "lblMaterials"
'            .Left = 5
'            .Top = 303 '203
'            .Width = 69
'            .Height = 16
'            .Enabled = True
'            .Visible = True
'            .Caption = "Materials:"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblMaterials, UIControl))

'        'lnDiv2 initial props
'        lnDiv2 = New UILine(oUILib)
'        With lnDiv2
'            .ControlName = "lnDiv2"
'            .Left = 1
'            .Top = 330 '230
'            .Width = 374
'            .Height = 0
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(lnDiv2, UIControl))

'        'lblComponent initial props
'        lblComponent = New UILabel(oUILib)
'        With lblComponent
'            .ControlName = "lblComponent"
'            .Left = 5
'            .Top = 340 '240
'            .Width = 92
'            .Height = 16
'            .Enabled = True
'            .Visible = True
'            .Caption = "Components:"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblComponent, UIControl))

'        'lnDiv3 initial props
'        lnDiv3 = New UILine(oUILib)
'        With lnDiv3
'            .ControlName = "lnDiv3"
'            .Left = 1
'            .Top = 430 '330
'            .Width = 374
'            .Height = 0
'            .Enabled = True
'            .Visible = True
'            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(lnDiv3, UIControl))

'        'lblDesigns initial props
'        lblDesigns = New UILabel(oUILib)
'        With lblDesigns
'            .ControlName = "lblDesigns"
'            .Left = 5
'            .Top = 440 ' 340
'            .Width = 92
'            .Height = 16
'            .Enabled = True
'            .Visible = True
'            .Caption = "Designs:"
'            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblDesigns, UIControl))

'        'lnDiv4 initial props
'        lnDiv4 = New UILine(oUILib)
'        With lnDiv4
'            .ControlName = "lnDiv4"
'            .Left = 1
'            .Top = 270 '170
'            .Width = 374
'            .Height = 0
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(lnDiv4, UIControl))

'        'lblResExist initial props
'        lblResExist = New UILabel(oUILib)
'        With lblResExist
'            .ControlName = "lblResExist"
'            .Left = 5
'            .Top = 5
'            .Width = 317
'            .Height = 16
'            .Enabled = True
'            .Visible = True
'            .Caption = "RESEARCH AN EXISTING DESIGN"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblResExist, UIControl))

'        'lnDiv5 initial props
'        lnDiv5 = New UILine(oUILib)
'        With lnDiv5
'            .ControlName = "lnDiv5"
'            .Left = 1
'            .Top = 25
'            .Width = 374
'            .Height = 0
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(lnDiv5, UIControl))

'        'lstTechs initial props
'        lstTechs = New UIListBox(oUILib)
'        With lstTechs
'            .ControlName = "lstTechs"
'            .Left = 5
'            .Top = 65 '45
'            .Width = 365
'            .Height = 176 '196
'            .Enabled = True
'            .Visible = True
'            .SetFont(New System.Drawing.Font("Courier New", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(lstTechs, UIControl))


'        'btnClose initial props
'        btnClose = New UIButton(oUILib)
'        With btnClose
'            .ControlName = "btnClose"
'            .Left = 135
'            .Top = 480 ' 380
'            .Width = 119
'            .Height = 23
'            .Enabled = True
'            .Visible = True
'            .Caption = "Close"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(btnClose, UIControl))

'        'lnDiv6 initial props
'        lnDiv6 = New UILine(oUILib)
'        With lnDiv6
'            .ControlName = "lnDiv6"
'            .Left = 1
'            .Top = 470 '370
'            .Width = 374
'            .Height = 0
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(lnDiv6, UIControl))

'        'btnResearch initial props
'        btnResearch = New UIButton(oUILib)
'        With btnResearch
'            .ControlName = "btnResearch"
'            .Left = 252
'            .Top = 245 '145
'            .Width = 119
'            .Height = 23
'            .Enabled = True
'            .Visible = True
'            .Caption = "Research"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(btnResearch, UIControl))

'        'lblComponentName initial props
'        lblComponentName = New UILabel(oUILib)
'        With lblComponentName
'            .ControlName = "lblComponentName"
'            .Left = 5
'            .Top = 48 '28
'            .Width = 101
'            .Height = 16
'            .Enabled = True
'            .Visible = True
'            .Caption = "Component Name"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblComponentName, UIControl))

'        'lblComponentType initial props
'        lblComponentType = New UILabel(oUILib)
'        With lblComponentType
'            .ControlName = "lblComponentType"
'            .Left = 160
'            .Top = 48 '28
'            .Width = 32
'            .Height = 16
'            .Enabled = True
'            .Visible = True
'            .Caption = "Type"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblComponentType, UIControl))

'        'lblResPhase initial props
'        lblResPhase = New UILabel(oUILib)
'        With lblResPhase
'            .ControlName = "lblResPhase"
'            .Left = 245
'            .Top = 48 '28
'            .Width = 100
'            .Height = 16
'            .Enabled = True
'            .Visible = True
'            .Caption = "Research Phase"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblResPhase, UIControl))

'        'btnViewResults initial props
'        btnViewResults = New UIButton(oUILib)
'        With btnViewResults
'            .ControlName = "btnViewResults"
'            .Left = btnResearch.Left - 125
'            .Top = 245 '145
'            .Width = 119
'            .Height = 23
'            .Enabled = True
'            .Visible = True
'            .Caption = "View Results"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(btnViewResults, UIControl))

'        'chkFilterArchived initial props
'        chkFilterArchived = New UICheckBox(oUILib)
'        With chkFilterArchived
'            .ControlName = "chkFilterArchived"
'            .Left = Me.Width - 110
'            .Top = 30
'            .Width = 100
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Filter Archived"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(6, DrawTextFormat)
'            .Value = True
'            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
'        End With
'        Me.AddChild(CType(chkFilterArchived, UIControl))

'        'chkSetFilter initial props
'        chkSetFilter = New UICheckBox(oUILib)
'        With chkSetFilter
'            .ControlName = "chkSetFilter"
'            .Left = 10
'            .Top = btnViewResults.Top + 2
'            .Width = 100
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Set Archived"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(6, DrawTextFormat)
'            .Value = False
'            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
'        End With
'        Me.AddChild(CType(chkSetFilter, UIControl))

'        'btnHelp initial props
'        btnHelp = New UIButton(oUILib)
'        With btnHelp
'            .ControlName = "btnHelp"
'            .Left = Me.Width - 24 - Me.BorderLineWidth
'            .Top = Me.Height - 28
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

'        'cboFilter initial props
'        cboFilter = New UIComboBox(oUILib)
'        With cboFilter
'            .ControlName = "cboFilter"
'            .Left = lblFilterSetting.Left + lblFilterSetting.Width + 10  'Me.Width - btnResearch.Width - 3
'            .Top = 28
'            .Width = btnResearch.Width
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceTextBoxFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .HighlightColor = System.Drawing.Color.FromArgb(255, muSettings.InterfaceFillColor.R, muSettings.InterfaceFillColor.G, muSettings.InterfaceFillColor.B)
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(cboFilter, UIControl))

'        'Now, position me better
'        Dim lFinalLeft As Int32
'        Dim lFinalTop As Int32
'        If (muSettings.ResearchWindowLeft < 0 AndAlso muSettings.ResearchWindowTop < 0) OrElse NewTutorialManager.TutorialOn = True Then
'            'Ok, check if the lower right is taken...
'            Dim ofrmChat As UIWindow = oUILib.GetWindow("frmChat")
'            Dim bTopSet As Boolean = False
'            lFinalLeft = oUILib.oDevice.PresentationParameters.BackBufferWidth - Me.Width
'            If ofrmChat Is Nothing = False Then
'                If ofrmChat.Left < lFinalLeft AndAlso ofrmChat.Left + ofrmChat.Width > lFinalLeft Then
'                    lFinalLeft = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (Me.Width \ 2)
'                    lFinalTop = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (Me.Height \ 2)
'                    bTopSet = True
'                End If
'            End If

'            If lFinalLeft + Me.Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then
'                'we exceed it... so center the form in the screen
'                lFinalLeft = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (Me.Width \ 2)
'                'Center the top too
'                lFinalTop = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (Me.Height \ 2)
'            ElseIf bTopSet = False Then
'                lFinalTop = oUILib.oDevice.PresentationParameters.BackBufferHeight - Me.Height
'            End If
'        Else
'            If muSettings.ResearchWindowTop < 0 OrElse (muSettings.ResearchWindowTop + Me.Height) > oUILib.oDevice.PresentationParameters.BackBufferHeight Then
'                'Ok, center vertically
'                lFinalTop = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (Me.Height \ 2)
'            Else : lFinalTop = muSettings.ResearchWindowTop
'            End If
'            If muSettings.ResearchWindowLeft < 0 OrElse (muSettings.ResearchWindowLeft + Me.Width) > oUILib.oDevice.PresentationParameters.BackBufferWidth Then
'                'Center horizontally
'                lFinalLeft = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (Me.Width \ 2)
'            Else : lFinalLeft = muSettings.ResearchWindowLeft
'            End If
'        End If

'        Me.Left = lFinalLeft
'        Me.Top = lFinalTop

'        muSettings.ResearchWindowLeft = lFinalLeft
'        muSettings.ResearchWindowTop = lFinalTop

'        FillFilterList()
'        cboFilter.FindComboItemData(muSettings.lCurrentResearchFilter)

'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'        If HasAliasedRights(AliasingRights.eViewResearch Or AliasingRights.eCreateDesigns Or AliasingRights.eViewTechDesigns) = True Then
'            MyBase.moUILib.AddWindow(Me)
'        Else
'            MyBase.moUILib.AddNotification("You lack rights to view the research interface.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
'            Return
'        End If


'        If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.ResearchWindow)

'        mbLoading = False
'    End Sub

'    Private Sub FillFilterList()
'        cboFilter.Clear()
'        cboFilter.AddItem("All Types") : cboFilter.ItemData(cboFilter.NewIndex) = -1
'        cboFilter.AddItem("Components Only") : cboFilter.ItemData(cboFilter.NewIndex) = -2
'        cboFilter.AddItem("Specials Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eSpecialTech
'        'If goCurrentPlayer.SpecialTechResearched(33) = True Then
'        cboFilter.AddItem("Alloys Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eAlloyTech
'        'End If
'        cboFilter.AddItem("Armors Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eArmorTech
'        cboFilter.AddItem("Engines Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eEngineTech
'        'cboFilter.AddItem("Hangars Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eHangarTech
'        cboFilter.AddItem("Hulls Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eHullTech
'        cboFilter.AddItem("Prototypes Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.ePrototype
'        cboFilter.AddItem("Radars Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eRadarTech
'        cboFilter.AddItem("Shields Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eShieldTech
'        cboFilter.AddItem("Weapons Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eWeaponTech
'    End Sub

'    Private Sub HideMeIfNeeded()
'        'MyBase.moUILib.RemoveWindow(Me.ControlName)
'    End Sub

'    Private Sub btnAlloy_Click(ByVal sName As String) Handles btnAlloy.Click
'        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
'            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
'            Return
'        End If
'        Dim ofrmAlloy As frmAlloyBuilder = New frmAlloyBuilder(MyBase.moUILib)
'        ofrmAlloy = Nothing
'        HideMeIfNeeded()
'    End Sub

'    Private Sub btnMineral_Click(ByVal sName As String) Handles btnMineral.Click
'        If HasAliasedRights(AliasingRights.eViewMining) = False Then
'            MyBase.moUILib.AddNotification("You lack rights to view mineral properties.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
'            Return
'        End If
'        Dim ofrmMatRes As frmMatRes = New frmMatRes(MyBase.moUILib)
'        ofrmMatRes = Nothing
'        HideMeIfNeeded()
'    End Sub

'    Private Sub btnArmor_Click(ByVal sName As String) Handles btnArmor.Click
'        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
'            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
'            Return
'        End If
'        Dim ofrmArmor As frmArmorBuilder = New frmArmorBuilder(MyBase.moUILib)
'        ofrmArmor = Nothing
'        HideMeIfNeeded()
'    End Sub

'    Private Sub btnEngine_Click(ByVal sName As String) Handles btnEngine.Click
'        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
'            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
'            Return
'        End If
'        Dim ofrmEngine As frmEngineBuilder = New frmEngineBuilder(MyBase.moUILib)
'        ofrmEngine = Nothing
'        HideMeIfNeeded()
'    End Sub

'    Private Sub btnRadar_Click(ByVal sName As String) Handles btnRadar.Click
'        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
'            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
'            Return
'        End If
'        Dim ofrmRadar As frmRadarBuilder = CType(goUILib.GetWindow("frmRadarBuilder"), frmRadarBuilder) ' New frmRadarBuilder(MyBase.moUILib)
'        If ofrmRadar Is Nothing Then ofrmRadar = New frmRadarBuilder(goUILib)
'        ofrmRadar = Nothing
'        HideMeIfNeeded()
'    End Sub

'    Private Sub btnShield_Click(ByVal sName As String) Handles btnShield.Click
'        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
'            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
'            Return
'        End If
'        Dim ofrmShield As frmShieldBuilder = New frmShieldBuilder(MyBase.moUILib)
'        ofrmShield = Nothing
'        HideMeIfNeeded()
'    End Sub

'    Public Sub RefreshComponentList()
'        Dim lFilterID As Int32 = -1

'        Dim lScrlBarVal As Int32 = lstTechs.ScrollBarValue

'        btnAlloy.Visible = True 'goCurrentPlayer.SpecialTechResearched(33)

'        If cboFilter.ListIndex > -1 Then lFilterID = cboFilter.ItemData(cboFilter.ListIndex)

'        Dim bFilterArchived As Boolean = chkFilterArchived.Value

'        lstTechs.Clear()
'        If goCurrentPlayer Is Nothing = False Then

'            For X As Int32 = 0 To goCurrentPlayer.mlTechUB
'                If goCurrentPlayer.moTechs(X) Is Nothing = False Then
'                    If lFilterID = -1 OrElse (lFilterID = -2 AndAlso goCurrentPlayer.moTechs(X).ObjTypeID <> ObjectType.eSpecialTech) OrElse goCurrentPlayer.moTechs(X).ObjTypeID = lFilterID Then

'                        If goCurrentPlayer.moTechs(X).bArchived = False OrElse bFilterArchived = False Then
'                            Dim sValue As String = goCurrentPlayer.moTechs(X).GetComponentName

'                            If sValue.Length > 20 Then sValue = sValue.Substring(0, 20)
'                            sValue = sValue.PadRight(22, " "c)

'                            sValue &= (Base_Tech.GetComponentTypeName(goCurrentPlayer.moTechs(X).ObjTypeID)).PadRight(12, " "c)

'                            If goCurrentPlayer.moTechs(X).Researchers <> 0 Then
'                                sValue &= "Assigned (" & goCurrentPlayer.moTechs(X).Researchers & ")"
'                            Else
'                                Select Case goCurrentPlayer.moTechs(X).ComponentDevelopmentPhase
'                                    Case Base_Tech.eComponentDevelopmentPhase.eComponentDesign
'                                        If goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eSpecialTech Then sValue &= "Unresearched" Else sValue &= "In Design"
'                                    Case Base_Tech.eComponentDevelopmentPhase.eComponentResearching
'                                        If goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eSpecialTech Then sValue &= "Unresearched" Else sValue &= "Designed"
'                                    Case Base_Tech.eComponentDevelopmentPhase.eResearched
'                                        sValue &= "Researched"
'                                    Case Base_Tech.eComponentDevelopmentPhase.eInvalidDesign
'                                        sValue &= "Poor Design" ' Design"
'                                End Select
'                            End If

'                            lstTechs.AddItem(sValue)
'                            lstTechs.ItemData(lstTechs.NewIndex) = goCurrentPlayer.moTechs(X).ObjectID
'                            lstTechs.ItemData2(lstTechs.NewIndex) = goCurrentPlayer.moTechs(X).ObjTypeID
'                        End If
'                    End If

'                End If
'            Next X
'        End If

'        lstTechs.ScrollBarValue = lScrlBarVal
'    End Sub

'    Private Sub btnResearch_Click(ByVal sName As String) Handles btnResearch.Click
'        If HasAliasedRights(AliasingRights.eAddResearch) = False Then
'            MyBase.moUILib.AddNotification("You lack rights to set research projects.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
'            Return
'        End If

'        If lstTechs.ListIndex > -1 Then
'            Dim lID As Int32 = lstTechs.ItemData(lstTechs.ListIndex)
'            Dim iTypeID As Int16 = CShort(lstTechs.ItemData2(lstTechs.ListIndex))

'            Dim oTmpTech As Base_Tech = goCurrentPlayer.GetTech(lID, iTypeID)
'            If oTmpTech Is Nothing OrElse oTmpTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
'                goUILib.AddNotification("Unable to Research selected component!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                Return
'            End If

'            If oTmpTech.ObjTypeID = ObjectType.eSpecialTech AndAlso oTmpTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
'                goUILib.AddNotification("That special project has already been researched.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                Return
'            End If
'            If oTmpTech.ObjTypeID = ObjectType.eSpecialTech AndAlso CType(oTmpTech, SpecialTech).bInTheTank = True Then
'                goUILib.AddNotification("Unable to research a special project that has been bypassed.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                Return
'            End If

'            Dim yData(13) As Byte

'            If goCurrentEnvir Is Nothing Then Exit Sub

'            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
'                If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID AndAlso goCurrentEnvir.oEntity(X).bSelected Then
'                    'ok, if the entity production is station...
'                    If goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eSpaceStationSpecial Then
'                        'Ok, need to find research lab... try our lstFac...
'                        Dim frmSelFac As frmSelectFac = CType(MyBase.moUILib.GetWindow("frmSelectFac"), frmSelectFac)
'                        If frmSelFac Is Nothing = False Then
'                            Dim oChild As StationChild = frmSelFac.GetCurrentChild()
'                            If oChild Is Nothing = False Then
'                                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yData, 0)
'                                System.BitConverter.GetBytes(oChild.lChildID).CopyTo(yData, 2)
'                                System.BitConverter.GetBytes(oChild.iChildTypeID).CopyTo(yData, 6)
'                                System.BitConverter.GetBytes(lID).CopyTo(yData, 8)
'                                System.BitConverter.GetBytes(iTypeID).CopyTo(yData, 12)

'                                MyBase.moUILib.GetMsgSys.SendToPrimary(yData)
'                                MyBase.moUILib.RemoveWindow(Me.ControlName)

'                                Return
'                            End If
'                        End If
'                    End If

'                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yData, 0)
'                    goCurrentEnvir.oEntity(X).GetGUIDAsString.CopyTo(yData, 2)
'                    System.BitConverter.GetBytes(lID).CopyTo(yData, 8)
'                    System.BitConverter.GetBytes(iTypeID).CopyTo(yData, 12)

'                    MyBase.moUILib.GetMsgSys.SendToPrimary(yData)

'                    If NewTutorialManager.TutorialOn = True Then
'                        NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eBuildOrderSent, lID, iTypeID, 1, "")
'                    End If

'                    HideMeIfNeeded()

'                    Exit For
'                End If
'            Next X

'        End If

'    End Sub

'    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'    End Sub

'    Private Sub btnHull_Click(ByVal sName As String) Handles btnHull.Click
'        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
'            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
'            Return
'        End If
'        Dim ofrmHull As frmHullBuilder = New frmHullBuilder(MyBase.moUILib)
'        ofrmHull = Nothing
'        HideMeIfNeeded()
'    End Sub

'    Private Sub btnPrototype_Click(ByVal sName As String) Handles btnPrototype.Click
'        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
'            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
'            Return
'        End If
'        Dim ofrmPrototype As frmPrototypeBuilder = New frmPrototypeBuilder(MyBase.moUILib)
'        ofrmPrototype = Nothing
'        HideMeIfNeeded()
'    End Sub

'    Private Sub btnWeapon_Click(ByVal sName As String) Handles btnWeapon.Click
'        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
'            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
'            Return
'        End If
'        Dim ofrmWeapon As frmWeaponBuilder = New frmWeaponBuilder(MyBase.moUILib)
'        ofrmWeapon = Nothing
'        HideMeIfNeeded()
'    End Sub

'    Private Sub btnViewResults_Click(ByVal sName As String) Handles btnViewResults.Click
'        If HasAliasedRights(AliasingRights.eViewTechDesigns) = False Then
'            MyBase.moUILib.AddNotification("You lack rights to view technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
'            Return
'        End If

'        If lstTechs.ListIndex > -1 Then
'            Dim lTechID As Int32 = lstTechs.ItemData(lstTechs.ListIndex)
'            Dim iTechTypeID As Int16 = CShort(lstTechs.ItemData2(lstTechs.ListIndex))

'            If NewTutorialManager.TutorialOn = True Then
'                Dim sParms() As String = {lTechID.ToString, iTechTypeID.ToString}
'                If goUILib.CommandAllowedWithParms(True, btnViewResults.GetFullName, sParms, False) = False Then Return
'                'If goUILib.CommandAllowed(True, btnViewResults.GetFullName) = False Then Return
'                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eViewResultsClicked, lTechID, iTechTypeID, -1, "")
'            End If

'            'TODO: Figure out how to get the prod factor...
'            Dim lProdFactor As Int32 = 100

'            Select Case iTechTypeID
'                Case ObjectType.eAlloyTech
'                    Dim ofrmAlloy As frmAlloyBuilder = New frmAlloyBuilder(MyBase.moUILib)
'                    ofrmAlloy.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), AlloyTech), lProdFactor)
'                    ofrmAlloy = Nothing
'                    'MyBase.moUILib.RemoveWindow(Me.ControlName)
'                Case ObjectType.eArmorTech
'                    Dim ofrmArmor As frmArmorBuilder = New frmArmorBuilder(MyBase.moUILib)
'                    ofrmArmor.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), ArmorTech), lProdFactor)
'                    ofrmArmor = Nothing
'                Case ObjectType.eEngineTech
'                    Dim ofrmEngine As frmEngineBuilder = New frmEngineBuilder(MyBase.moUILib)
'                    ofrmEngine.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), EngineTech), lProdFactor)
'                    ofrmEngine = Nothing
'                    'Case ObjectType.eHangarTech
'                    '    MyBase.moUILib.AddNotification("Not implemented yet!", Color.Yellow, -1, -1, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                    '    Return
'                Case ObjectType.eHullTech
'                    Dim ofrmHull As frmHullBuilder = New frmHullBuilder(MyBase.moUILib)
'                    ofrmHull.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), HullTech), lProdFactor)
'                    ofrmHull = Nothing
'                Case ObjectType.ePrototype
'                    Dim ofrmPrototype As frmPrototypeBuilder = New frmPrototypeBuilder(MyBase.moUILib)
'                    ofrmPrototype.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), PrototypeTech), lProdFactor)
'                    ofrmPrototype = Nothing
'                Case ObjectType.eRadarTech
'                    Dim ofrmRadar As frmRadarBuilder = New frmRadarBuilder(MyBase.moUILib)
'                    ofrmRadar.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), RadarTech), lProdFactor)
'                    ofrmRadar = Nothing
'                Case ObjectType.eShieldTech
'                    Dim ofrmShield As frmShieldBuilder = New frmShieldBuilder(MyBase.moUILib)
'                    ofrmShield.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), ShieldTech), lProdFactor)
'                    ofrmShield = Nothing
'                Case ObjectType.eWeaponTech
'                    Dim ofrmWeapon As frmWeaponBuilder = New frmWeaponBuilder(MyBase.moUILib)
'                    ofrmWeapon.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), WeaponTech), lProdFactor)
'                    ofrmWeapon = Nothing
'                Case ObjectType.eSpecialTech
'                    Dim ofrmSpecTech As frmSpecTech = New frmSpecTech(MyBase.moUILib)
'                    ofrmSpecTech.SetFromSpecialTech(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), SpecialTech))
'                    ofrmSpecTech = Nothing
'            End Select

'            If iTechTypeID <> ObjectType.eSpecialTech Then HideMeIfNeeded()
'        End If

'    End Sub

'    Private Sub frmResearchMain_OnNewFrame() Handles Me.OnNewFrame

'        If glCurrentCycle - mlLastCheck > 30 Then
'            mlLastCheck = glCurrentCycle
'            'Ok, go through the list and check... this basically is exactly like RefreshComponentList
'            Dim lFilterID As Int32 = -1

'            Dim bFilter As Boolean = chkFilterArchived.Value

'            If goCurrentPlayer Is Nothing = False Then
'                'If btnAlloy.Visible <> goCurrentPlayer.SpecialTechResearched(33) Then btnAlloy.Visible = Not btnAlloy.Visible
'                If cboFilter.ListIndex > -1 Then lFilterID = cboFilter.ItemData(cboFilter.ListIndex)

'                For X As Int32 = 0 To goCurrentPlayer.mlTechUB
'                    If goCurrentPlayer.moTechs(X) Is Nothing = False Then
'                        If (goCurrentPlayer.moTechs(X).bArchived = False OrElse bFilter = False) AndAlso (lFilterID = -1 OrElse (lFilterID = -2 AndAlso goCurrentPlayer.moTechs(X).ObjTypeID <> ObjectType.eSpecialTech) OrElse goCurrentPlayer.moTechs(X).ObjTypeID = lFilterID) Then
'                            With goCurrentPlayer.moTechs(X)
'                                Dim lIdx As Int32 = -1
'                                For lTmpIdx As Int32 = 0 To lstTechs.ListCount - 1
'                                    If lstTechs.ItemData(lTmpIdx) = .ObjectID AndAlso lstTechs.ItemData2(lTmpIdx) = .ObjTypeID Then
'                                        lIdx = lTmpIdx
'                                        Exit For
'                                    End If
'                                Next lTmpIdx

'                                If lIdx = -1 Then
'                                    'It should be there but it isn't, so force a full refresh
'                                    RefreshComponentList()
'                                    Return
'                                End If

'                                Dim sValue As String = .GetComponentName
'                                If sValue.Length > 20 Then sValue = sValue.Substring(0, 20)
'                                sValue = sValue.PadRight(22, " "c)
'                                sValue &= (Base_Tech.GetComponentTypeName(.ObjTypeID)).PadRight(12, " "c)

'                                If goCurrentPlayer.moTechs(X).Researchers <> 0 Then
'                                    sValue &= "Assigned (" & goCurrentPlayer.moTechs(X).Researchers & ")"
'                                Else
'                                    Select Case goCurrentPlayer.moTechs(X).ComponentDevelopmentPhase
'                                        Case Base_Tech.eComponentDevelopmentPhase.eComponentDesign
'                                            If goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eSpecialTech Then sValue &= "Unresearched" Else sValue &= "In Design"
'                                        Case Base_Tech.eComponentDevelopmentPhase.eComponentResearching
'                                            If goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eSpecialTech Then sValue &= "Unresearched" Else sValue &= "Designed"
'                                        Case Base_Tech.eComponentDevelopmentPhase.eResearched
'                                            sValue &= "Researched"
'                                        Case Base_Tech.eComponentDevelopmentPhase.eInvalidDesign
'                                            sValue &= "Impossible Design"
'                                    End Select
'                                End If

'                                If lstTechs.List(lIdx) <> sValue Then
'                                    lstTechs.List(lIdx) = sValue
'                                End If
'                            End With
'                        End If

'                    End If
'                Next X
'            End If

'        End If
'    End Sub

'    Private Sub frmResearchMain_OnResize() Handles Me.OnResize
'        If mbLoading = True Then Return
'        muSettings.ResearchWindowLeft = Me.Left
'        muSettings.ResearchWindowTop = Me.Top
'    End Sub

'    Private Sub cboFilter_ItemSelected(ByVal lItemIndex As Integer) Handles cboFilter.ItemSelected
'        If lItemIndex > -1 Then muSettings.lCurrentResearchFilter = cboFilter.ItemData(lItemIndex)
'        If NewTutorialManager.TutorialOn = True Then
'            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eResearchWindowFilterSelected, muSettings.lCurrentResearchFilter, -1, -1, "")
'        End If
'        RefreshComponentList()
'    End Sub

'    Private Sub btnSpecials_Click(ByVal sName As String) Handles btnSpecials.Click
'        If HasAliasedRights(AliasingRights.eViewTechDesigns) = False Then
'            MyBase.moUILib.AddNotification("You lack rights to view technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
'            Return
'        End If
'        Dim ofrmSpecial As frmSpecTech = New frmSpecTech(MyBase.moUILib)
'        ofrmSpecial = Nothing
'        HideMeIfNeeded()
'    End Sub

'    Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
'        If goTutorial Is Nothing Then goTutorial = New TutorialManager()
'        goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eResearchMain)
'    End Sub

'    Private Sub chkFilterArchived_Click() Handles chkFilterArchived.Click
'        MyBase.moUILib.GetMsgSys.LoadArchived()
'        RefreshComponentList()
'    End Sub

'    Private Sub lstTechs_ItemClick(ByVal lIndex As Integer) Handles lstTechs.ItemClick
'        mbIgnoreSetFilter = True
'        chkSetFilter.Value = False
'        mbIgnoreSetFilter = False
'        If lstTechs.ListIndex > -1 Then
'            Dim lID As Int32 = lstTechs.ItemData(lstTechs.ListIndex)
'            Dim iTypeID As Int16 = CShort(lstTechs.ItemData2(lstTechs.ListIndex))

'            Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lID, iTypeID)
'            Dim sToolTip As String = ""
'            If oTech Is Nothing = False Then
'                If oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched AndAlso oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
'                    If oTech.oResearchCost Is Nothing = False Then
'                        sToolTip = oTech.oResearchCost.GetBuildCostText("")
'                    End If
'                End If
'            End If
'            btnResearch.ToolTipText = sToolTip

'            mbIgnoreSetFilter = True
'            chkSetFilter.Value = oTech.bArchived
'            mbIgnoreSetFilter = False
'        End If
'    End Sub

'    Private Sub chkSetFilter_Click() Handles chkSetFilter.Click
'        If mbIgnoreSetFilter = True Then Return
'        If lstTechs.ListIndex > -1 Then
'            Dim lID As Int32 = lstTechs.ItemData(lstTechs.ListIndex)
'            Dim iTypeID As Int16 = CShort(lstTechs.ItemData2(lstTechs.ListIndex))

'            Dim yMsg(8) As Byte
'            Dim lPos As Int32 = 0
'            System.BitConverter.GetBytes(GlobalMessageCode.eSetArchiveState).CopyTo(yMsg, lPos) : lPos += 2
'            System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
'            System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, lPos) : lPos += 2
'            If chkSetFilter.Value = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 0
'            lPos += 1

'            Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lID, iTypeID)
'            If oTech Is Nothing = False Then
'                oTech.bArchived = chkSetFilter.Value
'                If oTech.ObjTypeID = ObjectType.eAlloyTech Then
'                    'ok, see if it has a mineral with it
'                    Dim lMinResult As Int32 = CType(oTech, AlloyTech).AlloyResultID
'                    If lMinResult > -1 Then
'                        For X As Int32 = 0 To glMineralUB
'                            If glMineralIdx(X) = lMinResult Then
'                                goMinerals(X).bArchived = chkSetFilter.Value
'                                Exit For
'                            End If
'                        Next X
'                    End If
'                End If
'            End If

'            MyBase.moUILib.SendMsgToPrimary(yMsg)

'            lstTechs.ListIndex = -1
'            RefreshComponentList()
'        End If
'    End Sub

'    Private Sub frmResearchMain_WindowMoved() Handles Me.WindowMoved
'        muSettings.ResearchWindowLeft = Me.Left
'        muSettings.ResearchWindowTop = Me.Top
'    End Sub
'End Class
#End Region

Public Class frmResearchMain
    Inherits UIWindow

    Private WithEvents btnAlloy As UIButton
    Private lblNewResearch As UILabel
    Private WithEvents btnMineral As UIButton
    Private WithEvents btnArmor As UIButton
    Private WithEvents btnEngine As UIButton
    Private WithEvents btnRadar As UIButton
    Private WithEvents btnShield As UIButton
    Private WithEvents btnWeapon As UIButton
    Private WithEvents btnSpecials As UIButton
    Private WithEvents btnHull As UIButton
    Private WithEvents btnPrototype As UIButton
    Private WithEvents btnWorkSheet As UIButton
    Private lnDiv1 As UILine
    Private lnDiv4 As UILine
    Private lblTitle As UILabel
    Private lnDiv5 As UILine
    Private WithEvents tvwTechs As UITreeView
    Private WithEvents btnClose As UIButton
    Private lnDiv6 As UILine
    Private WithEvents btnResearch As UIButton

    Private WithEvents btnViewResults As UIButton

    Private WithEvents chkFilterArchived As UICheckBox
    Private WithEvents chkSetFilter As UICheckBox

    Private WithEvents btnHelp As UIButton

    Private mlLastCheck As Int32

    Private mbIgnoreSetFilter As Boolean = False

    Private mbLoading As Boolean = True
    Private mbFromStation As Boolean = False

    'Private Shared mbExp(33, 256) As Boolean

    Private mbExp(33, 256) As Boolean

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenResearchWindow)

        'frmResearchMain initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eResearchMain
            .ControlName = "frmResearchMain"
            .Left = 158
            .Top = 144
            .Width = 375
            .Height = 511 '411
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
            .BorderColor = muSettings.InterfaceBorderColor
            '.FillColor = System.Drawing.Color.FromArgb(128, 64, 128, 192)
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .mbAcceptReprocessEvents = True
        End With

        'lstTechs initial props
        tvwTechs = New UITreeView(oUILib) 'UIListBox(oUILib)
        With tvwTechs
            .ControlName = "tvwTechs"
            .Left = 3
            .Top = 31
            .Width = Me.Width - 6 '365
            .Height = 305 '272
            .Enabled = True
            .Visible = True
            .SetFont(New System.Drawing.Font("Courier New", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .mbAcceptReprocessEvents = True
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(tvwTechs, UIControl))

        'btnResearch initial props
        btnResearch = New UIButton(oUILib)
        With btnResearch
            .ControlName = "btnResearch"
            .Left = 252
            .Top = 340 ' 305
            .Width = 119
            .Height = 23
            .Enabled = NewTutorialManager.TutorialOn = True Or glMineralUB > -1
            .Visible = True
            .Caption = "Research"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(btnResearch, UIControl))

        'btnViewResults initial props
        btnViewResults = New UIButton(oUILib)
        With btnViewResults
            .ControlName = "btnViewResults"
            .Left = btnResearch.Left - 125
            .Top = btnResearch.Top
            .Width = 119
            .Height = 23
            .Enabled = NewTutorialManager.TutorialOn = True Or glMineralUB > -1
            .Visible = True
            .Caption = "View Results"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(btnViewResults, UIControl))

        'chkSetFilter initial props
        chkSetFilter = New UICheckBox(oUILib)
        With chkSetFilter
            .ControlName = "chkSetFilter"
            .Left = 10
            .Top = btnViewResults.Top + 2
            .Width = 100
            .Height = 18
            .Enabled = Not gbAliased AndAlso (NewTutorialManager.TutorialOn = True Or glMineralUB > -1)
            .Visible = True
            .Caption = "Set Archived"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkSetFilter, UIControl))

        'lnDiv4 initial props
        lnDiv4 = New UILine(oUILib)
        With lnDiv4
            .ControlName = "lnDiv4"
            .Left = Me.BorderLineWidth \ 2
            .Top = 365 '330
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv4, UIControl))

        'lblSelect initial props
        lblNewResearch = New UILabel(oUILib)
        With lblNewResearch
            .ControlName = "lblNewResearch"
            .Left = 5
            .Top = 370 '335 '175
            .Width = 317
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "New Research:"
            '.ForeColor = System.Drawing.Color.FromArgb(-16711681)
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblNewResearch, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = Me.BorderLineWidth \ 2
            .Top = 390 '355 '194
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            '.BorderColor = System.Drawing.Color.FromArgb(-16711681)
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'btnAlloy initial props
        btnAlloy = New UIButton(oUILib)
        With btnAlloy
            .ControlName = "btnAlloy"
            .Left = 5
            .Top = 395 '360 '300
            .Width = 119
            .Height = 23
            .Enabled = NewTutorialManager.TutorialOn = True Or glMineralUB > -1
            .Visible = True
            .Caption = "Alloy"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnAlloy, UIControl))

        'btnArmor initial props
        btnArmor = New UIButton(oUILib)
        With btnArmor
            .ControlName = "btnArmor"
            .Left = btnAlloy.Left + btnAlloy.Width + 5
            .Top = btnAlloy.Top
            .Width = 119
            .Height = 23
            .Enabled = NewTutorialManager.TutorialOn = True Or glMineralUB > -1
            .Visible = True
            .Caption = "Armor"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnArmor, UIControl))

        'btnEngine initial props
        btnEngine = New UIButton(oUILib)
        With btnEngine
            .ControlName = "btnEngine"
            .Left = btnArmor.Left + btnArmor.Width + 5
            .Top = btnArmor.Top
            .Width = 119
            .Height = 23
            .Enabled = NewTutorialManager.TutorialOn = True Or glMineralUB > -1
            .Visible = True
            .Caption = "Engine"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnEngine, UIControl))

        'btnHull initial props
        btnHull = New UIButton(oUILib)
        With btnHull
            .ControlName = "btnHull"
            .Left = 5
            .Top = btnAlloy.Top + btnAlloy.Height + 5
            .Width = 119
            .Height = 23
            .Enabled = NewTutorialManager.TutorialOn = True Or glMineralUB > -1
            .Visible = True
            .Caption = "Hull"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnHull, UIControl))

        'btnMineral initial props
        btnMineral = New UIButton(oUILib)
        With btnMineral
            .ControlName = "btnMineral"
            .Left = btnHull.Left + btnHull.Width + 5
            .Top = btnHull.Top
            .Width = 119
            .Height = 23
            .Enabled = True
            .Visible = True
            .Caption = "Mineral"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnMineral, UIControl))

        'btnPrototype initial props
        btnPrototype = New UIButton(oUILib)
        With btnPrototype
            .ControlName = "btnPrototype"
            .Left = btnMineral.Left + btnMineral.Width + 5
            .Top = btnMineral.Top
            .Width = 119
            .Height = 23
            .Enabled = NewTutorialManager.TutorialOn = True Or glMineralUB > -1
            .Visible = True
            .Caption = "Prototype"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnPrototype, UIControl))

        'btnRadar initial props
        btnRadar = New UIButton(oUILib)
        With btnRadar
            .ControlName = "btnRadar"
            .Left = 5
            .Top = btnHull.Top + btnHull.Height + 5
            .Width = 119
            .Height = 23
            .Enabled = NewTutorialManager.TutorialOn = True Or glMineralUB > -1
            .Visible = True
            .Caption = "Radar"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnRadar, UIControl))

        'btnShield initial props
        btnShield = New UIButton(oUILib)
        With btnShield
            .ControlName = "btnShield"
            .Left = btnRadar.Left + btnRadar.Width + 5
            .Top = btnRadar.Top
            .Width = 119
            .Height = 23
            .Enabled = NewTutorialManager.TutorialOn = True Or glMineralUB > -1
            .Visible = True
            .Caption = "Shield"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnShield, UIControl))

        'btnSpecials initial props
        btnSpecials = New UIButton(oUILib)
        With btnSpecials
            .ControlName = "btnSpecials"
            .Left = btnShield.Left + btnShield.Width + 5
            .Top = btnShield.Top
            .Width = 119
            .Height = 23
            .Enabled = True
            .Visible = True
            .Caption = "Specials"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnSpecials, UIControl))

        'btnWeapon initial props
        btnWeapon = New UIButton(oUILib)
        With btnWeapon
            .ControlName = "btnWeapon"
            .Left = btnShield.Left
            .Top = btnShield.Top + btnShield.Height + 5
            .Width = 119
            .Height = 23
            .Enabled = NewTutorialManager.TutorialOn = True Or glMineralUB > -1
            .Visible = True
            .Caption = "Weapon"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnWeapon, UIControl))

        'btnWorkSheet initial props
        btnWorkSheet = New UIButton(oUILib)
        With btnWorkSheet
            .ControlName = "btnWorkSheet"
            .Left = btnRadar.Left
            .Top = btnRadar.Top + btnRadar.Height + 1
            .Width = 32
            .Height = 32
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = Color.Cyan 'muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .ControlImageRect = grc_UI(elInterfaceRectangle.eQuickbar_Command)
            .ControlImageRect_Disabled = .ControlImageRect
            .ControlImageRect_Normal = .ControlImageRect
            .ControlImageRect_Pressed = .ControlImageRect_Pressed
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
            .ToolTipText = "Click to open the Research Worksheet window."
        End With
        Me.AddChild(CType(btnWorkSheet, UIControl))

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 317
            .Height = 16
            .Enabled = True
            .Visible = True
            .Caption = "Research Management"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 24 - Me.BorderLineWidth
            .Top = Me.BorderLineWidth
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "X"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'btnHelp initial props
        btnHelp = New UIButton(oUILib)
        With btnHelp
            .ControlName = "btnHelp"
            .Left = btnClose.Left - 25
            .Top = Me.BorderLineWidth
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "?"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
            .ToolTipText = "Click to begin the tutorial for this window."
        End With
        Me.AddChild(CType(btnHelp, UIControl))

        'lnDiv5 initial props
        lnDiv5 = New UILine(oUILib)
        With lnDiv5
            .ControlName = "lnDiv5"
            .Left = Me.BorderLineWidth \ 2
            .Top = btnClose.Top + btnClose.Height + 2
            .Width = Me.Width - Me.BorderLineWidth
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv5, UIControl))

        'chkFilterArchived initial props
        chkFilterArchived = New UICheckBox(oUILib)
        With chkFilterArchived
            .ControlName = "chkFilterArchived"
            .Left = Me.Width \ 2
            .Top = btnClose.Top + 3
            .Width = 100
            .Height = 18
            .Enabled = NewTutorialManager.TutorialOn = True Or glMineralUB > -1
            .Visible = True
            .Caption = "Filter Archived"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        End With
        Me.AddChild(CType(chkFilterArchived, UIControl))

        'Now, position me better
        Dim lFinalLeft As Int32
        Dim lFinalTop As Int32

        Dim lDesiredLeft As Int32 = muSettings.ResearchWindowLeft
        Dim lDesiredTop As Int32 = muSettings.ResearchWindowTop
        Dim ofrmSelectFac As UIWindow = oUILib.GetWindow("frmSelectFac")
        If ofrmSelectFac Is Nothing = False Then
            lDesiredTop = ofrmSelectFac.Top
            lDesiredLeft = ofrmSelectFac.Left + ofrmSelectFac.Width
            mbFromStation = True
        End If

        If (lDesiredLeft < 0 AndAlso muSettings.ResearchWindowTop < 0) OrElse NewTutorialManager.TutorialOn = True Then
            'Ok, check if the lower right is taken...
            Dim bTopSet As Boolean = False
            lFinalLeft = oUILib.oDevice.PresentationParameters.BackBufferWidth - Me.Width
            Dim ofrmChat As UIWindow = oUILib.GetWindow("frmChat")
            If ofrmChat Is Nothing = False Then
                If ofrmChat.Left < lFinalLeft AndAlso ofrmChat.Left + ofrmChat.Width > lFinalLeft Then
                    lFinalLeft = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (Me.Width \ 2)
                    lFinalTop = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (Me.Height \ 2)
                    bTopSet = True
                End If
            End If

            If lFinalLeft + Me.Width > oUILib.oDevice.PresentationParameters.BackBufferWidth Then
                'we exceed it... so center the form in the screen
                lFinalLeft = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (Me.Width \ 2)
                'Center the top too
                lFinalTop = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (Me.Height \ 2)
            ElseIf bTopSet = False Then
                lFinalTop = oUILib.oDevice.PresentationParameters.BackBufferHeight - Me.Height
            End If
        Else
            If lDesiredTop < 0 OrElse (lDesiredTop + Me.Height) > oUILib.oDevice.PresentationParameters.BackBufferHeight Then
                'Ok, center vertically
                lFinalTop = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - (Me.Height \ 2)
            Else : lFinalTop = lDesiredTop
            End If
            If lDesiredLeft < 0 OrElse (lDesiredLeft + Me.Width) > oUILib.oDevice.PresentationParameters.BackBufferWidth Then
                'Center horizontally
                lFinalLeft = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - (Me.Width \ 2)
            Else : lFinalLeft = lDesiredLeft
            End If
        End If

        Me.Left = lFinalLeft
        Me.Top = lFinalTop

        If Not mbFromStation Then
            muSettings.ResearchWindowLeft = lFinalLeft
            muSettings.ResearchWindowTop = lFinalTop
        End If

        FillFilterList()
        'cboFilter.FindComboItemData(muSettings.lCurrentResearchFilter)

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        If HasAliasedRights(AliasingRights.eViewResearch Or AliasingRights.eCreateDesigns Or AliasingRights.eViewTechDesigns) = True Then
            MyBase.moUILib.AddWindow(Me)
        Else
            MyBase.moUILib.AddNotification("You lack rights to view the research interface.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If


        If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.ResearchWindow)
        If NewTutorialManager.TutorialOn = False AndAlso glMineralUB = -1 Then
            MyBase.moUILib.AddNotification("You havn't discoverd any minerals!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            MyBase.moUILib.AddNotification("You must first mine and research a resource", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            MyBase.moUILib.AddNotification("before you can create alloys, design components, or fabricate prototypes.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
        End If

        mbLoading = False
    End Sub

    Private Sub FillFilterList()
        RefreshComponentList()
        'cboFilter.Clear()
        'cboFilter.AddItem("All Types") : cboFilter.ItemData(cboFilter.NewIndex) = -1
        'cboFilter.AddItem("Components Only") : cboFilter.ItemData(cboFilter.NewIndex) = -2
        'cboFilter.AddItem("Specials Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eSpecialTech
        ''If goCurrentPlayer.SpecialTechResearched(33) = True Then
        'cboFilter.AddItem("Alloys Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eAlloyTech
        ''End If
        'cboFilter.AddItem("Armors Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eArmorTech
        'cboFilter.AddItem("Engines Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eEngineTech
        ''cboFilter.AddItem("Hangars Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eHangarTech
        'cboFilter.AddItem("Hulls Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eHullTech
        'cboFilter.AddItem("Prototypes Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.ePrototype
        'cboFilter.AddItem("Radars Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eRadarTech
        'cboFilter.AddItem("Shields Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eShieldTech
        'cboFilter.AddItem("Weapons Only") : cboFilter.ItemData(cboFilter.NewIndex) = ObjectType.eWeaponTech
    End Sub

    Private Sub HideMeIfNeeded()
        'MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

#Region "  Create New Buttons  "
    Private Sub btnAlloy_Click(ByVal sName As String) Handles btnAlloy.Click
        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        Dim ofrmAlloy As frmAlloyBuilder = New frmAlloyBuilder(MyBase.moUILib)
        ofrmAlloy = Nothing
        HideMeIfNeeded()
    End Sub

    Private Sub btnMineral_Click(ByVal sName As String) Handles btnMineral.Click
        If HasAliasedRights(AliasingRights.eViewMining) = False Then
            MyBase.moUILib.AddNotification("You lack rights to view mineral properties.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        Dim ofrmMatRes As frmMatRes = New frmMatRes(MyBase.moUILib)
        ofrmMatRes = Nothing
        HideMeIfNeeded()
    End Sub

    Private Sub btnArmor_Click(ByVal sName As String) Handles btnArmor.Click
        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        Dim ofrmArmor As frmArmorBuilder = New frmArmorBuilder(MyBase.moUILib)
        ofrmArmor = Nothing
        HideMeIfNeeded()
    End Sub

    Private Sub btnEngine_Click(ByVal sName As String) Handles btnEngine.Click
        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        Dim ofrmEngine As frmEngineBuilder = New frmEngineBuilder(MyBase.moUILib)
        ofrmEngine = Nothing
        HideMeIfNeeded()
    End Sub

    Private Sub btnRadar_Click(ByVal sName As String) Handles btnRadar.Click
        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        Dim ofrmRadar As frmRadarBuilder = CType(goUILib.GetWindow("frmRadarBuilder"), frmRadarBuilder) ' New frmRadarBuilder(MyBase.moUILib)
        If ofrmRadar Is Nothing Then ofrmRadar = New frmRadarBuilder(goUILib)
        ofrmRadar = Nothing
        HideMeIfNeeded()
    End Sub

    Private Sub btnShield_Click(ByVal sName As String) Handles btnShield.Click
        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        Dim ofrmShield As frmShieldBuilder = New frmShieldBuilder(MyBase.moUILib)
        ofrmShield = Nothing
        HideMeIfNeeded()
    End Sub

    Private Sub btnHull_Click(ByVal sName As String) Handles btnHull.Click
        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        Dim ofrmHull As frmHullBuilder = New frmHullBuilder(MyBase.moUILib)
        ofrmHull = Nothing
        HideMeIfNeeded()
    End Sub

    Private Sub btnPrototype_Click(ByVal sName As String) Handles btnPrototype.Click
        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        Dim ofrmPrototype As frmPrototypeBuilder = New frmPrototypeBuilder(MyBase.moUILib)
        ofrmPrototype = Nothing
        HideMeIfNeeded()
    End Sub

    Private Sub btnWeapon_Click(ByVal sName As String) Handles btnWeapon.Click
        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eCreateDesigns) = False Then
            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        Dim ofrmWeapon As frmWeaponBuilder = New frmWeaponBuilder(MyBase.moUILib)
        ofrmWeapon = Nothing
        HideMeIfNeeded()
    End Sub

    Private Sub btnSpecials_Click(ByVal sName As String) Handles btnSpecials.Click
        If HasAliasedRights(AliasingRights.eViewTechDesigns) = False Then
            MyBase.moUILib.AddNotification("You lack rights to view technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        Dim ofrmSpecial As frmSpecTech = New frmSpecTech(MyBase.moUILib)
        ofrmSpecial = Nothing
        HideMeIfNeeded()
    End Sub

#End Region

    Private Sub btnResearch_Click(ByVal sName As String) Handles btnResearch.Click
        If HasAliasedRights(AliasingRights.eAddResearch) = False Then
            MyBase.moUILib.AddNotification("You lack rights to set research projects.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If

        Dim oNode As UITreeView.UITreeViewItem = tvwTechs.oSelectedNode

        If oNode Is Nothing = False Then
            Dim lID As Int32 = oNode.lItemData
            Dim iTypeID As Int16 = CShort(oNode.lItemData2)

            Dim oTmpTech As Base_Tech = goCurrentPlayer.GetTech(lID, iTypeID)
            If oTmpTech Is Nothing OrElse oTmpTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eInvalidDesign OrElse (oTmpTech.MajorDesignFlaw And 31) <> 0 Then
                goUILib.AddNotification("Unable to Research selected component!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            If oTmpTech.ObjTypeID = ObjectType.eSpecialTech AndAlso oTmpTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
                goUILib.AddNotification("That special project has already been researched.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
            If oTmpTech.ObjTypeID = ObjectType.eSpecialTech AndAlso CType(oTmpTech, SpecialTech).bInTheTank = True Then
                goUILib.AddNotification("Unable to research a special project that has been bypassed.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            Dim yData(13) As Byte

            If goCurrentEnvir Is Nothing Then Exit Sub

            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID AndAlso goCurrentEnvir.oEntity(X).bSelected Then
                    'ok, if the entity production is station...
                    If goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eSpaceStationSpecial Then
                        'Ok, need to find research lab... try our lstFac...
                        Dim frmSelFac As frmSelectFac = CType(MyBase.moUILib.GetWindow("frmSelectFac"), frmSelectFac)
                        If frmSelFac Is Nothing = False Then
                            Dim oChild As StationChild = frmSelFac.GetCurrentChild()
                            If oChild Is Nothing = False Then
                                System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yData, 0)
                                System.BitConverter.GetBytes(oChild.lChildID).CopyTo(yData, 2)
                                System.BitConverter.GetBytes(oChild.iChildTypeID).CopyTo(yData, 6)
                                System.BitConverter.GetBytes(lID).CopyTo(yData, 8)
                                System.BitConverter.GetBytes(iTypeID).CopyTo(yData, 12)

                                MyBase.moUILib.GetMsgSys.SendToPrimary(yData)
                                MyBase.moUILib.RemoveWindow(Me.ControlName)

                                Return
                            End If
                        End If
                    End If

                    System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yData, 0)
                    goCurrentEnvir.oEntity(X).GetGUIDAsString.CopyTo(yData, 2)
                    System.BitConverter.GetBytes(lID).CopyTo(yData, 8)
                    System.BitConverter.GetBytes(iTypeID).CopyTo(yData, 12)

                    MyBase.moUILib.GetMsgSys.SendToPrimary(yData)

                    If NewTutorialManager.TutorialOn = True Then
                        NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eBuildOrderSent, lID, iTypeID, 1, "")
                    End If

                    HideMeIfNeeded()

                    Exit For
                End If
            Next X

        End If

    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
        Dim ofrmSelectFac As UIWindow = MyBase.moUILib.GetWindow("frmSelectFac")
        If ofrmSelectFac Is Nothing = False Then
            MyBase.moUILib.RemoveWindow("frmSelectFac")
        End If
        ofrmSelectFac = Nothing
    End Sub

    Private Function GetRootNode(ByVal iTypeID As Int16, ByVal lHullTypeID As Int32) As UITreeView.UITreeViewItem

        Dim oNode As UITreeView.UITreeViewItem = tvwTechs.GetNodeByItemData2(-(lHullTypeID + 2), iTypeID)
        If oNode Is Nothing Then
            Dim oRoot As UITreeView.UITreeViewItem = tvwTechs.GetNodeByItemData2(Int32.MinValue, iTypeID)
            If oRoot Is Nothing Then Return Nothing
            If iTypeID = ObjectType.eArmorTech OrElse iTypeID = ObjectType.eAlloyTech OrElse iTypeID = ObjectType.eSpecialTech Then Return oRoot
            oNode = tvwTechs.AddNode(Base_Tech.GetHullTypeName(CByte(lHullTypeID)), -(lHullTypeID + 2), iTypeID, 0, oRoot, Nothing)
            If oNode.sItem = "Universal" AndAlso iTypeID = ObjectType.ePrototype Then oNode.sItem = "Hull Deleted"
            oNode.bItemBold = True : oNode.bUseFillColor = True : oNode.clrFillColor = System.Drawing.Color.FromArgb(255, 64, 64, 64)
        End If
        Return oNode
    End Function

    Public Sub RefreshComponentList()
        btnAlloy.Visible = True

        Dim bFilterArchived As Boolean = chkFilterArchived.Value
        Dim lTypeID() As Int32 = {ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eSpecialTech, ObjectType.eWeaponTech}

        SaveListExpansion()
        tvwTechs.Clear()

        Dim oAlloy As UITreeView.UITreeViewItem = tvwTechs.AddNode("Alloys", Int32.MinValue, ObjectType.eAlloyTech, 0, Nothing, Nothing)
        Dim oArmor As UITreeView.UITreeViewItem = tvwTechs.AddNode("Armor", Int32.MinValue, ObjectType.eArmorTech, 0, Nothing, Nothing)
        Dim oEngine As UITreeView.UITreeViewItem = tvwTechs.AddNode("Engines", Int32.MinValue, ObjectType.eEngineTech, 0, Nothing, Nothing)
        Dim oHull As UITreeView.UITreeViewItem = tvwTechs.AddNode("Hulls", Int32.MinValue, ObjectType.eHullTech, 0, Nothing, Nothing)
        Dim oPrototypes As UITreeView.UITreeViewItem = tvwTechs.AddNode("Prototypes", Int32.MinValue, ObjectType.ePrototype, 0, Nothing, Nothing)
        Dim oRadar As UITreeView.UITreeViewItem = tvwTechs.AddNode("Radar", Int32.MinValue, ObjectType.eRadarTech, 0, Nothing, Nothing)
        Dim oShield As UITreeView.UITreeViewItem = tvwTechs.AddNode("Shields", Int32.MinValue, ObjectType.eShieldTech, 0, Nothing, Nothing)
        Dim oSpecials As UITreeView.UITreeViewItem = tvwTechs.AddNode("Special Projects", Int32.MinValue, ObjectType.eSpecialTech, 0, Nothing, Nothing)
        Dim oWeapon As UITreeView.UITreeViewItem = tvwTechs.AddNode("Weapons", Int32.MinValue, ObjectType.eWeaponTech, 0, Nothing, Nothing)

        oAlloy.bItemBold = True : oAlloy.bUseFillColor = True
        oArmor.bItemBold = True : oArmor.bUseFillColor = True
        oEngine.bItemBold = True : oEngine.bUseFillColor = True
        oHull.bItemBold = True : oHull.bUseFillColor = True
        oPrototypes.bItemBold = True : oPrototypes.bUseFillColor = True
        oRadar.bItemBold = True : oRadar.bUseFillColor = True
        oShield.bItemBold = True : oShield.bUseFillColor = True
        oSpecials.bItemBold = True : oSpecials.bUseFillColor = True
        oWeapon.bItemBold = True : oWeapon.bUseFillColor = True

        Dim oOther As UITreeView.UITreeViewItem = Nothing

        If goCurrentPlayer Is Nothing = False Then

            For X As Int32 = 0 To goCurrentPlayer.mlTechUB
                If goCurrentPlayer.moTechs(X) Is Nothing = False Then

                    ' If goCurrentPlayer.moTechs(X).GetComponentName.ToUpper = "TEST DESTRUCTION" AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.ePrototype Then Stop
                    If goCurrentPlayer.moTechs(X).bArchived = False OrElse bFilterArchived = False Then
                        Dim sValue As String = goCurrentPlayer.moTechs(X).GetResearchMainText()

                        'Dim oRoot As UITreeView.UITreeViewItem = tvwTechs.GetNodeByItemData2(-1, goCurrentPlayer.moTechs(X).ObjTypeID)
                        'If oRoot Is Nothing Then
                        '    If oOther Is Nothing Then oOther = tvwTechs.AddNode("Other", -1, -1, 0, Nothing, Nothing)
                        '    oRoot = oOther
                        'End If
                        Dim oRoot As UITreeView.UITreeViewItem = GetRootNode(goCurrentPlayer.moTechs(X).ObjTypeID, goCurrentPlayer.moTechs(X).GetTechHullTypeID())
                        Dim oTech As UITreeView.UITreeViewItem = tvwTechs.AddNode(sValue, goCurrentPlayer.moTechs(X).ObjectID, goCurrentPlayer.moTechs(X).ObjTypeID, -1, oRoot, Nothing)
                        If oRoot.oParentNode Is Nothing = False Then oRoot.oParentNode.lItemData3 += 1
                        oRoot.lItemData3 += 1
                        'lstTechs.AddItem(sValue)
                        'lstTechs.ItemData(lstTechs.NewIndex) = goCurrentPlayer.moTechs(X).ObjectID
                        'lstTechs.ItemData2(lstTechs.NewIndex) = goCurrentPlayer.moTechs(X).ObjTypeID
                    End If

                End If
            Next X
        End If

        For X As Int32 = 0 To lTypeID.GetUpperBound(0)
            Dim oNode As UITreeView.UITreeViewItem = tvwTechs.GetNodeByItemData2(Int32.MinValue, lTypeID(X))
            If oNode Is Nothing = False Then
                oNode.bExpanded = mbExp(X, 0)
                oNode.sItem = oNode.sItem & " (" & oNode.lItemData3 & ")"
                oNode.clrFillColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                If lTypeID(X) <> ObjectType.eAlloyTech AndAlso lTypeID(X) <> ObjectType.eArmorTech AndAlso lTypeID(X) <> ObjectType.eSpecialTech Then
                    Dim oTmpNode As UITreeView.UITreeViewItem = oNode.oFirstChild
                    While oTmpNode Is Nothing = False
                        oTmpNode.sItem = oTmpNode.sItem & " (" & oTmpNode.lItemData3 & ")"
                        oTmpNode.bExpanded = mbExp(X, -(oTmpNode.lItemData) - 2 + 1)
                        oTmpNode = oTmpNode.oNextSibling
                    End While
                End If
            End If
        Next X

        'lstTechs.ScrollBarValue = lScrlBarVal
    End Sub

    Private Sub SaveListExpansion()
        Dim lTypeID() As Int32 = {ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eSpecialTech, ObjectType.eWeaponTech}
        'Dim bExp(lTypeID.GetUpperBound(0), 256) As Boolean
        For X As Int32 = 0 To lTypeID.GetUpperBound(0)
            'mbExp(X, 0) = False
            If muSettings.ResearchWindowStaticExpand = False Then
                For Y As Int32 = 0 To 256
                    mbExp(X, Y) = False
                Next Y
            End If

            Dim oNode As UITreeView.UITreeViewItem = tvwTechs.GetNodeByItemData2(Int32.MinValue, lTypeID(X))
            If oNode Is Nothing = False Then
                mbExp(X, 0) = oNode.bExpanded
                If oNode.lItemData2 <> ObjectType.eAlloyTech AndAlso oNode.lItemData2 <> ObjectType.eArmorTech AndAlso oNode.lItemData2 <> ObjectType.eSpecialTech Then
                    Dim oTmpNode As UITreeView.UITreeViewItem = oNode.oFirstChild
                    While oTmpNode Is Nothing = False
                        mbExp(X, -(oTmpNode.lItemData) - 2 + 1) = oTmpNode.bExpanded
                        oTmpNode = oTmpNode.oNextSibling
                    End While
                End If
            End If
        Next X

    End Sub

    Private Sub btnViewResults_Click(ByVal sName As String) Handles btnViewResults.Click
        If HasAliasedRights(AliasingRights.eViewTechDesigns) = False Then
            MyBase.moUILib.AddNotification("You lack rights to view technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If

        Dim oNode As UITreeView.UITreeViewItem = tvwTechs.oSelectedNode
        If oNode Is Nothing = False Then
            Dim lTechID As Int32 = oNode.lItemData
            Dim iTechTypeID As Int16 = CShort(oNode.lItemData2)

            If lTechID < 1 Then Return

            If NewTutorialManager.TutorialOn = True Then
                Dim sParms() As String = {lTechID.ToString, iTechTypeID.ToString}
                If goUILib.CommandAllowedWithParms(True, btnViewResults.GetFullName, sParms, False) = False Then Return
                'If goUILib.CommandAllowed(True, btnViewResults.GetFullName) = False Then Return
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eViewResultsClicked, lTechID, iTechTypeID, -1, "")
            End If

            'TODO: Figure out how to get the prod factor...
            Dim lProdFactor As Int32 = 100

            Select Case iTechTypeID
                Case ObjectType.eAlloyTech
                    Dim ofrmAlloy As frmAlloyBuilder = New frmAlloyBuilder(MyBase.moUILib)
                    ofrmAlloy.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), AlloyTech), lProdFactor)
                    ofrmAlloy = Nothing
                    'MyBase.moUILib.RemoveWindow(Me.ControlName)
                Case ObjectType.eArmorTech
                    Dim ofrmArmor As frmArmorBuilder = New frmArmorBuilder(MyBase.moUILib)
                    ofrmArmor.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), ArmorTech), lProdFactor)
                    ofrmArmor = Nothing
                Case ObjectType.eEngineTech
                    Dim ofrmEngine As frmEngineBuilder = New frmEngineBuilder(MyBase.moUILib)
                    ofrmEngine.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), EngineTech), lProdFactor)
                    ofrmEngine = Nothing
                    'Case ObjectType.eHangarTech
                    '    MyBase.moUILib.AddNotification("Not implemented yet!", Color.Yellow, -1, -1, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    '    Return
                Case ObjectType.eHullTech
                    Dim ofrmHull As frmHullBuilder = New frmHullBuilder(MyBase.moUILib)
                    ofrmHull.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), HullTech), lProdFactor)
                    ofrmHull = Nothing
                Case ObjectType.ePrototype
                    Dim ofrmPrototype As frmPrototypeBuilder = New frmPrototypeBuilder(MyBase.moUILib)
                    ofrmPrototype.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), PrototypeTech), lProdFactor)
                    ofrmPrototype = Nothing
                Case ObjectType.eRadarTech
                    Dim ofrmRadar As frmRadarBuilder = New frmRadarBuilder(MyBase.moUILib)
                    ofrmRadar.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), RadarTech), lProdFactor)
                    ofrmRadar = Nothing
                Case ObjectType.eShieldTech
                    Dim ofrmShield As frmShieldBuilder = New frmShieldBuilder(MyBase.moUILib)
                    ofrmShield.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), ShieldTech), lProdFactor)
                    ofrmShield = Nothing
                Case ObjectType.eWeaponTech
                    Dim ofrmWeapon As frmWeaponBuilder = New frmWeaponBuilder(MyBase.moUILib)
                    ofrmWeapon.ViewResults(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), WeaponTech), lProdFactor)
                    ofrmWeapon = Nothing
                Case ObjectType.eSpecialTech
                    Dim ofrmSpecTech As frmSpecTech = New frmSpecTech(MyBase.moUILib)
                    ofrmSpecTech.SetFromSpecialTech(CType(goCurrentPlayer.GetTech(lTechID, iTechTypeID), SpecialTech))
                    ofrmSpecTech = Nothing
            End Select

            If iTechTypeID <> ObjectType.eSpecialTech Then HideMeIfNeeded()
        End If
    End Sub

    Private Sub tvwTechs_NodeDoubleClicked() Handles tvwTechs.NodeDoubleClicked
        If btnViewResults.Enabled = False Then Exit Sub
        Dim oNode As UITreeView.UITreeViewItem = tvwTechs.oSelectedNode
        If oNode Is Nothing = True OrElse oNode.lItemData < 0 Then Exit Sub
        btnViewResults_Click(Nothing)
        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
    End Sub

    Private Sub frmResearchMain_OnNewFrame() Handles Me.OnNewFrame
        'If True = True Then Return
        If glCurrentCycle - mlLastCheck > 30 Then
            mlLastCheck = glCurrentCycle
            'Ok, go through the list and check... this basically is exactly like RefreshComponentList
            'Dim lFilterID As Int32 = -1

            Dim bFilter As Boolean = chkFilterArchived.Value

            If goCurrentPlayer Is Nothing = False Then
                'If btnAlloy.Visible <> goCurrentPlayer.SpecialTechResearched(33) Then btnAlloy.Visible = Not btnAlloy.Visible
                'If cboFilter.ListIndex > -1 Then lFilterID = cboFilter.ItemData(cboFilter.ListIndex)

                For X As Int32 = 0 To goCurrentPlayer.mlTechUB
                    If goCurrentPlayer.moTechs(X) Is Nothing = False Then
                        If (goCurrentPlayer.moTechs(X).bArchived = False OrElse bFilter = False) Then 'AndAlso (lFilterID = -1 OrElse (lFilterID = -2 AndAlso goCurrentPlayer.moTechs(X).ObjTypeID <> ObjectType.eSpecialTech) OrElse goCurrentPlayer.moTechs(X).ObjTypeID = lFilterID) Then
                            With goCurrentPlayer.moTechs(X)
                                'Dim lIdx As Int32 = -1
                                'For lTmpIdx As Int32 = 0 To lstTechs.ListCount - 1
                                '    If lstTechs.ItemData(lTmpIdx) = .ObjectID AndAlso lstTechs.ItemData2(lTmpIdx) = .ObjTypeID Then
                                '        lIdx = lTmpIdx
                                '        Exit For
                                '    End If
                                'Next lTmpIdx
                                Dim sValue As String = .GetResearchMainText()
                                Dim oNode As UITreeView.UITreeViewItem = tvwTechs.GetNodeByItemData2(.ObjectID, .ObjTypeID)
                                If oNode Is Nothing Then
                                    Dim oRoot As UITreeView.UITreeViewItem = GetRootNode(.ObjTypeID, .GetTechHullTypeID()) 'tvwTechs.GetNodeByItemData2(-1, .ObjTypeID)
                                    If oRoot Is Nothing Then oRoot = tvwTechs.AddNode("Other", -1, -1, -1, Nothing, Nothing)
                                    oNode = tvwTechs.AddNode(sValue, .ObjectID, .ObjTypeID, -1, oRoot, Nothing)
                                    Me.IsDirty = True
                                    Dim lIdx As Int32 = -1
                                    If oRoot.oParentNode Is Nothing = False Then
                                        oRoot.oParentNode.lItemData3 += 1
                                        lIdx = oRoot.oParentNode.sItem.LastIndexOf("("c)
                                        If lIdx > -1 Then
                                            oRoot.oParentNode.sItem = oRoot.oParentNode.sItem.Substring(0, lIdx) & "(" & oRoot.oParentNode.lItemData3 & ")"
                                        Else
                                            oRoot.oParentNode.sItem &= " (" & oRoot.oParentNode.lItemData3 & ")"
                                        End If
                                    End If
                                    oRoot.lItemData3 += 1
                                    lIdx = oRoot.sItem.LastIndexOf("("c)
                                    If lIdx > -1 Then
                                        oRoot.sItem = oRoot.sItem.Substring(0, lIdx) & "(" & oRoot.lItemData3 & ")"
                                    Else
                                        oRoot.sItem &= " (" & oRoot.lItemData3 & ")"
                                    End If

                                End If
                                If oNode Is Nothing = False Then
                                    If oNode.sItem <> sValue Then
                                        oNode.sItem = sValue
                                        Me.IsDirty = True
                                    End If
                                End If



                            End With
                        End If

                    End If
                Next X
            End If

        End If
    End Sub

    Private Sub frmResearchMain_OnResize() Handles Me.OnResize
        If mbLoading = True Then Return
        muSettings.ResearchWindowLeft = Me.Left
        muSettings.ResearchWindowTop = Me.Top
    End Sub

    'Private Sub cboFilter_ItemSelected(ByVal lItemIndex As Integer) Handles cboFilter.ItemSelected
    '    If lItemIndex > -1 Then muSettings.lCurrentResearchFilter = cboFilter.ItemData(lItemIndex)
    '    If NewTutorialManager.TutorialOn = True Then
    '        NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eResearchWindowFilterSelected, muSettings.lCurrentResearchFilter, -1, -1, "")
    '    End If
    '    RefreshComponentList()
    'End Sub

    Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
        If goTutorial Is Nothing Then goTutorial = New TutorialManager()
        goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eResearchMain)
    End Sub

    Private Sub chkFilterArchived_Click() Handles chkFilterArchived.Click
        MyBase.moUILib.GetMsgSys.LoadArchived()
        RefreshComponentList()
    End Sub


    Private Sub chkSetFilter_Click() Handles chkSetFilter.Click
        If mbIgnoreSetFilter = True Then Return

        Dim oNode As UITreeView.UITreeViewItem = tvwTechs.oSelectedNode
        If oNode Is Nothing = False Then
            Dim lID As Int32 = oNode.lItemData
            Dim iTypeID As Int16 = CShort(oNode.lItemData2)

            If lID < 1 Then Return

            Dim yMsg(8) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetArchiveState).CopyTo(yMsg, lPos) : lPos += 2
            System.BitConverter.GetBytes(lID).CopyTo(yMsg, lPos) : lPos += 4
            System.BitConverter.GetBytes(iTypeID).CopyTo(yMsg, lPos) : lPos += 2
            If chkSetFilter.Value = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 0
            lPos += 1

            Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lID, iTypeID)
            If oTech Is Nothing = False Then
                oTech.bArchived = chkSetFilter.Value
                If oTech.ObjTypeID = ObjectType.eAlloyTech Then
                    'ok, see if it has a mineral with it
                    Dim lMinResult As Int32 = CType(oTech, AlloyTech).AlloyResultID
                    If lMinResult > -1 Then
                        For X As Int32 = 0 To glMineralUB
                            If glMineralIdx(X) = lMinResult Then
                                goMinerals(X).bArchived = chkSetFilter.Value
                                Exit For
                            End If
                        Next X
                    End If
                End If
            End If

            MyBase.moUILib.SendMsgToPrimary(yMsg)

            'lstTechs.ListIndex = -1
            RefreshComponentList()
        End If
    End Sub

    Private Sub frmResearchMain_WindowClosed() Handles Me.WindowClosed
        SaveListExpansion()
    End Sub

    Private Sub frmResearchMain_WindowMoved() Handles Me.WindowMoved
        If Not mbFromStation Then
            muSettings.ResearchWindowLeft = Me.Left
            muSettings.ResearchWindowTop = Me.Top
        Else
            Dim ofrmSelectFac As UIWindow = MyBase.moUILib.GetWindow("frmSelectFac")
            If ofrmSelectFac Is Nothing = False Then
                ofrmSelectFac.Top = Me.Top
                ofrmSelectFac.Left = Me.Left - ofrmSelectFac.Width
            End If
        End If
    End Sub

    Private Sub tvwTechs_NodeExpanded(ByVal oNode As UITreeView.UITreeViewItem) Handles tvwTechs.NodeExpanded
        If oNode Is Nothing = False Then
            If NewTutorialManager.TutorialOn = True Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eResearchWindowFilterSelected, oNode.lItemData, oNode.lItemData2, -1, "")
            End If
        End If
    End Sub

    Private Sub tvwTechs_NodeSelected(ByRef oNode As UITreeView.UITreeViewItem) Handles tvwTechs.NodeSelected
        mbIgnoreSetFilter = True
        chkSetFilter.Value = False
        mbIgnoreSetFilter = False
        If oNode Is Nothing = False Then
            Dim lID As Int32 = oNode.lItemData
            Dim iTypeID As Int16 = CShort(oNode.lItemData2)

            If lID < 1 Then Return

            Dim oTech As Base_Tech = goCurrentPlayer.GetTech(lID, iTypeID)
            If oTech Is Nothing Then Return
            Dim sToolTip As String = ""
            If oTech Is Nothing = False Then
                If oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched AndAlso oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
                    If oTech.oResearchCost Is Nothing = False Then
                        sToolTip = oTech.oResearchCost.GetBuildCostText("")
                    End If
                End If
            End If
            btnResearch.ToolTipText = sToolTip

            mbIgnoreSetFilter = True
            chkSetFilter.Value = oTech.bArchived
            mbIgnoreSetFilter = False
        End If
    End Sub

    Private Sub btnWorkSheet_Click(ByVal sName As String) Handles btnWorkSheet.Click
        If gbAliased = True AndAlso HasAliasedRights(AliasingRights.eViewTechDesigns) = False Then
            MyBase.moUILib.AddNotification("You lack rights to create technology designs.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If
        Dim ofrmWorkSheet As frmBuilderWrksht = New frmBuilderWrksht(MyBase.moUILib)
        ofrmWorkSheet = Nothing
        HideMeIfNeeded()
    End Sub
End Class
