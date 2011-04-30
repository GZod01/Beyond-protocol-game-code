Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmRepair
    Inherits UIWindow

    Private Const ml_FULL_HP_HEIGHT As Int32 = 100
    Private Const ml_HP_BAR_BOTTOM As Int32 = 125

    Private fraBack As UIWindow
    Private shpLeft As UIWindow
    Private shpFront As UIWindow
    Private shpStructure As UIWindow
    Private shpBack As UIWindow
    Private shpRight As UIWindow

    'Defense Repair Buttons
    Private WithEvents btnLeft As UIButton
    Private WithEvents btnFront As UIButton
    Private WithEvents btnStructure As UIButton
    Private WithEvents btnBack As UIButton
    Private WithEvents btnRight As UIButton
    Private WithEvents btnRepairAll As UIButton

    Private lblLeft As UILabel
    Private lblFront As UILabel
    Private lblStructure As UILabel
    Private lblBack As UILabel
    Private lblRight As UILabel

    Private WithEvents btnCancelRepair As UIButton
    Private WithEvents btnClose As UIButton
    Private WithEvents btnHelp As UIButton

    Private lnDiv1 As UILine
    Private lnDiv2 As UILine

    'Component Repair items
    Private lblEngines As UILabel
    Private lblRadar As UILabel
    Private lblShields As UILabel
    Private lblHangar As UILabel
    Private lblCargo As UILabel
    Private lblWeapons As UILabel
    Private lblFore1 As UILabel
    Private lblFore2 As UILabel
    Private lblLeft1 As UILabel
    Private lblLeft2 As UILabel
    Private lblRear1 As UILabel
    Private lblRear2 As UILabel
    Private lblRight1 As UILabel
    Private lblRight2 As UILabel
    'Private lblFuelBay As UILabel

    Private WithEvents btnEngines As UIButton
    Private WithEvents btnRepairAllComps As UIButton
    Private WithEvents btnRadar As UIButton
    Private WithEvents btnShield As UIButton
    Private WithEvents btnHangar As UIButton
    Private WithEvents btnCargo As UIButton
    'Private WithEvents btnFuelBay As UIButton
    Private WithEvents btnFore1 As UIButton
    Private WithEvents btnFore2 As UIButton
    Private WithEvents btnLeft1 As UIButton
    Private WithEvents btnLeft2 As UIButton
    Private WithEvents btnRear1 As UIButton
    Private WithEvents btnRear2 As UIButton
    Private WithEvents btnRight1 As UIButton
    Private WithEvents btnRight2 As UIButton

    Private mlParentID As Int32
    Private miParentTypeID As Int16
    Private mlEntityID As Int32
    Private miEntityTypeID As Int16

    Private mlParentEntityIdx As Int32 = -1

    Private mlLastDefenseRequest As Int32 = -1

    Shared mbRequestedArchived As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmRepair initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eRepairWindow
            .ControlName = "frmRepair"
            Dim oContents As frmContents = CType(MyBase.moUILib.GetWindow("frmContents"), frmContents)
            If oContents Is Nothing = False Then
                .Left = oContents.Left + oContents.Width + 2
                .Top = oContents.Top
            Else
                .Left = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - 256
                .Top = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 128
            End If
            oContents = Nothing
            .Width = 512 '340
            .Height = 225 '345
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 3
            .Moveable = True
        End With

        'fraBack initial props
        fraBack = New UIWindow(oUILib)
        With fraBack
            .ControlName = "fraBack"
            .Left = 20
            .Top = 25
            .Width = 105
            .Height = 100
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = Color.FromArgb(255, 255, 0, 0)
            .FullScreen = False
            .BorderLineWidth = 1
            .Moveable = False
            .bAcceptEvents = False
        End With
        Me.AddChild(CType(fraBack, UIControl))

        'shpLeft initial props
        shpLeft = New UIWindow(oUILib)
        With shpLeft
            .ControlName = "shpLeft"
            .Left = 21
            .Top = 65
            .Width = 20
            .Height = 60
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = Color.FromArgb(255, 0, 255, 0)
            .FullScreen = False
            .DrawBorder = False
            .Moveable = False
            .bRoundedBorder = False
        End With
        Me.AddChild(CType(shpLeft, UIControl))

        'shpFront initial props
        shpFront = New UIWindow(oUILib)
        With shpFront
            .ControlName = "shpFront"
            .Left = 42
            .Top = 55
            .Width = 20
            .Height = 70
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = Color.FromArgb(255, 0, 255, 0)
            .FullScreen = False
            .DrawBorder = False
            .Moveable = False
            .bRoundedBorder = False
        End With
        Me.AddChild(CType(shpFront, UIControl))

        'shpStructure initial props
        shpStructure = New UIWindow(oUILib)
        With shpStructure
            .ControlName = "shpStructure"
            .Left = 63
            .Top = 25
            .Width = 20
            .Height = 100
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = Color.FromArgb(255, 0, 255, 0)
            .FullScreen = False
            .DrawBorder = False
            .Moveable = False
            .bRoundedBorder = False
        End With
        Me.AddChild(CType(shpStructure, UIControl))

        'shpBack initial props
        shpBack = New UIWindow(oUILib)
        With shpBack
            .ControlName = "shpBack"
            .Left = 84
            .Top = 70
            .Width = 20
            .Height = 55
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = Color.FromArgb(255, 0, 255, 0)
            .FullScreen = False
            .DrawBorder = False
            .Moveable = False
            .bRoundedBorder = False
        End With
        Me.AddChild(CType(shpBack, UIControl))

        'shpRight initial props
        shpRight = New UIWindow(oUILib)
        With shpRight
            .ControlName = "shpRight"
            .Left = 105
            .Top = 75
            .Width = 20
            .Height = 50
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = Color.FromArgb(255, 0, 255, 0)
            .FullScreen = False
            .DrawBorder = False
            .Moveable = False
            .bRoundedBorder = False
        End With
        Me.AddChild(CType(shpRight, UIControl))

        'btnLeft initial props
        btnLeft = New UIButton(oUILib)
        With btnLeft
            .ControlName = "btnLeft"
            .Left = 20
            .Top = 127
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to Repair the Armor on the Left Side"
        End With
        Me.AddChild(CType(btnLeft, UIControl))

        'btnFront initial props
        btnFront = New UIButton(oUILib)
        With btnFront
            .ControlName = "btnFront"
            .Left = 41
            .Top = 127
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to Repair the Armor on the Front Side"
        End With
        Me.AddChild(CType(btnFront, UIControl))

        'btnStructure initial props
        btnStructure = New UIButton(oUILib)
        With btnStructure
            .ControlName = "btnStructure"
            .Left = 62
            .Top = 127
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to Repair the Structural Integrity"
        End With
        Me.AddChild(CType(btnStructure, UIControl))

        'btnBack initial props
        btnBack = New UIButton(oUILib)
        With btnBack
            .ControlName = "btnBack"
            .Left = 84
            .Top = 127
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to Repair the Armor on the Back Side"
        End With
        Me.AddChild(CType(btnBack, UIControl))

        'btnRight initial props
        btnRight = New UIButton(oUILib)
        With btnRight
            .ControlName = "btnRight"
            .Left = 105
            .Top = 127
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to Repair the Armor on the Right Side"
        End With
        Me.AddChild(CType(btnRight, UIControl))

        'btnRepairAll initial props
        btnRepairAll = New UIButton(oUILib)
        With btnRepairAll
            .ControlName = "btnRepairAll"
            .Left = 20
            .Top = 150 '278
            .Width = 110
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Repair All"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to Repair the Armor on all damaged sides"
        End With
        Me.AddChild(CType(btnRepairAll, UIControl))

        'lblLeft initial props
        lblLeft = New UILabel(oUILib)
        With lblLeft
            .ControlName = "lblLeft"
            .Left = 21
            .Top = 4
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "L"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblLeft, UIControl))

        'lblFront initial props
        lblFront = New UILabel(oUILib)
        With lblFront
            .ControlName = "lblFront"
            .Left = 42
            .Top = 4
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "F"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblFront, UIControl))

        'lblStructure initial props
        lblStructure = New UILabel(oUILib)
        With lblStructure
            .ControlName = "lblStructure"
            .Left = 63
            .Top = 4
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "S"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblStructure, UIControl))

        'lblBack initial props
        lblBack = New UILabel(oUILib)
        With lblBack
            .ControlName = "lblBack"
            .Left = 84
            .Top = 4
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "B"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblBack, UIControl))

        'lblRight initial props
        lblRight = New UILabel(oUILib)
        With lblRight
            .ControlName = "lblRight"
            .Left = 105
            .Top = 4
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(lblRight, UIControl))

        'btnCancelRepair initial props
        btnCancelRepair = New UIButton(oUILib)
        With btnCancelRepair
            .ControlName = "btnCancelRepair"
            .Left = Me.Width \ 2 - 55
            .Top = Me.Height - 30
            .Width = 110
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Cancel Repair"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to cancel all repair efforts." & vbCrLf & "Clearing the production queue has the same effect."
        End With
        Me.AddChild(CType(btnCancelRepair, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = Me.Width - 25
            .Top = 1
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "X"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to close this window"
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'btnHelp initial props
        btnHelp = New UIButton(oUILib)
        With btnHelp
            .ControlName = "btnHelp"
            .Left = btnClose.Left - 25
            .Top = btnClose.Top
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "?"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to start the tutorial for this window"
        End With
        Me.AddChild(CType(btnHelp, UIControl))

        'lblEngines initial props
        lblEngines = New UILabel(oUILib)
        With lblEngines
            .ControlName = "lblEngines"
            .Left = 150
            .Top = 5
            .Width = 80
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Engines"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Status of the Engines." & vbCrLf & "Green indicates Operational." & vbCrLf & "Red indicates Destroyed." & vbCrLf & "Gray indicates that the component" & vbCrLf & "is not on this unit."
        End With
        Me.AddChild(CType(lblEngines, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 140
            .Top = 1
            .Width = 0
            .Height = 185
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = 1
            .Top = 185
            .Width = Me.Width - 1
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        'btnEngines initial props
        btnEngines = New UIButton(oUILib)
        With btnEngines
            .ControlName = "btnEngines"
            .Left = 250
            .Top = 4
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to repair the Engines."
        End With
        Me.AddChild(CType(btnEngines, UIControl))

        'lblRadar initial props
        lblRadar = New UILabel(oUILib)
        With lblRadar
            .ControlName = "lblRadar"
            .Left = 150
            .Top = 30
            .Width = 70
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Radar"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Status of the Radar." & vbCrLf & "Green indicates Operational." & vbCrLf & "Red indicates Destroyed." & vbCrLf & "Gray indicates that the component" & vbCrLf & "is not on this unit."
        End With
        Me.AddChild(CType(lblRadar, UIControl))

        'lblShields initial props
        lblShields = New UILabel(oUILib)
        With lblShields
            .ControlName = "lblShields"
            .Left = 150
            .Top = 55
            .Width = 70
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Shield"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Status of the Shields." & vbCrLf & "Green indicates Operational." & vbCrLf & "Red indicates Destroyed." & vbCrLf & "Gray indicates that the component" & vbCrLf & "is not on this unit."
        End With
        Me.AddChild(CType(lblShields, UIControl))

        'lblHangar initial props
        lblHangar = New UILabel(oUILib)
        With lblHangar
            .ControlName = "lblHangar"
            .Left = 150
            .Top = 80
            .Width = 90
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Hangar Bay"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Status of the Hangar Bay." & vbCrLf & "Green indicates Operational." & vbCrLf & "Red indicates Destroyed." & vbCrLf & "Gray indicates that the component" & vbCrLf & "is not on this unit."
        End With
        Me.AddChild(CType(lblHangar, UIControl))

        'lblCargo initial props
        lblCargo = New UILabel(oUILib)
        With lblCargo
            .ControlName = "lblCargo"
            .Left = 150
            .Top = 105
            .Width = 80
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Cargo Bay"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Status of the Cargo Bay." & vbCrLf & "Green indicates Operational." & vbCrLf & "Red indicates Destroyed." & vbCrLf & "Gray indicates that the component" & vbCrLf & "is not on this unit."
        End With
        Me.AddChild(CType(lblCargo, UIControl))

        'lblWeapons initial props
        lblWeapons = New UILabel(oUILib)
        With lblWeapons
            .ControlName = "lblWeapons"
            .Left = 295
            .Top = 10
            .Width = 83
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "WEAPONS"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblWeapons, UIControl))

        'lblFore1 initial props
        lblFore1 = New UILabel(oUILib)
        With lblFore1
            .ControlName = "lblFore1"
            .Left = 300
            .Top = 30
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Forward 1"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Status of the Forward Weapons (group 1)." & vbCrLf & "Green indicates Operational." & vbCrLf & "Red indicates Destroyed." & vbCrLf & "Gray indicates that the component" & vbCrLf & "is not on this unit."
        End With
        Me.AddChild(CType(lblFore1, UIControl))

        'lblFore2 initial props
        lblFore2 = New UILabel(oUILib)
        With lblFore2
            .ControlName = "lblFore2"
            .Left = 410
            .Top = 30
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Forward 2"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Status of the Forward Weapons (group 2)." & vbCrLf & "Green indicates Operational." & vbCrLf & "Red indicates Destroyed." & vbCrLf & "Gray indicates that the component" & vbCrLf & "is not on this unit."
        End With
        Me.AddChild(CType(lblFore2, UIControl))

        'lblLeft1 initial props
        lblLeft1 = New UILabel(oUILib)
        With lblLeft1
            .ControlName = "lblLeft1"
            .Left = 300
            .Top = 55
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Left 1"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Status of the Left Weapons (group 1)." & vbCrLf & "Green indicates Operational." & vbCrLf & "Red indicates Destroyed." & vbCrLf & "Gray indicates that the component" & vbCrLf & "is not on this unit."
        End With
        Me.AddChild(CType(lblLeft1, UIControl))

        'lblLeft2 initial props
        lblLeft2 = New UILabel(oUILib)
        With lblLeft2
            .ControlName = "lblLeft2"
            .Left = 411
            .Top = 55
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Left 2"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Status of the Left Weapons (group 2)." & vbCrLf & "Green indicates Operational." & vbCrLf & "Red indicates Destroyed." & vbCrLf & "Gray indicates that the component" & vbCrLf & "is not on this unit."
        End With
        Me.AddChild(CType(lblLeft2, UIControl))

        'lblRear1 initial props
        lblRear1 = New UILabel(oUILib)
        With lblRear1
            .ControlName = "lblRear1"
            .Left = 300
            .Top = 80
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Rear 1"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Status of the Rear Weapons (group 1)." & vbCrLf & "Green indicates Operational." & vbCrLf & "Red indicates Destroyed." & vbCrLf & "Gray indicates that the component" & vbCrLf & "is not on this unit."
        End With
        Me.AddChild(CType(lblRear1, UIControl))

        'lblRear2 initial props
        lblRear2 = New UILabel(oUILib)
        With lblRear2
            .ControlName = "lblRear2"
            .Left = 410
            .Top = 80
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Rear 2"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Status of the Rear Weapons (group 2)." & vbCrLf & "Green indicates Operational." & vbCrLf & "Red indicates Destroyed." & vbCrLf & "Gray indicates that the component" & vbCrLf & "is not on this unit."
        End With
        Me.AddChild(CType(lblRear2, UIControl))

        'lblRight1 initial props
        lblRight1 = New UILabel(oUILib)
        With lblRight1
            .ControlName = "lblRight1"
            .Left = 300
            .Top = 105
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Right 1"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Status of the Right Weapons (group 1)." & vbCrLf & "Green indicates Operational." & vbCrLf & "Red indicates Destroyed." & vbCrLf & "Gray indicates that the component" & vbCrLf & "is not on this unit."
        End With
        Me.AddChild(CType(lblRight1, UIControl))

        'lblRight2 initial props
        lblRight2 = New UILabel(oUILib)
        With lblRight2
            .ControlName = "lblRight2"
            .Left = 409
            .Top = 105
            .Width = 60
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Right 2"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .ToolTipText = "Status of the Right Weapons (group 2)." & vbCrLf & "Green indicates Operational." & vbCrLf & "Red indicates Destroyed." & vbCrLf & "Gray indicates that the component" & vbCrLf & "is not on this unit."
        End With
        Me.AddChild(CType(lblRight2, UIControl))

        ''lblFuelBay initial props
        'lblFuelBay = New UILabel(oUILib)
        'With lblFuelBay
        '    .ControlName = "lblFuelBay"
        '    .Left = 150
        '    .Top = 130
        '    .Width = 70
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Fuel Bay"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(4, DrawTextFormat)
        '    .ToolTipText = "Status of the Fuel Bay." & vbCrLf & "Green indicates Operational." & vbCrLf & "Red indicates Destroyed." & vbCrLf & "Gray indicates that the component" & vbCrLf & "is not on this unit."
        'End With
        'Me.AddChild(CType(lblFuelBay, UIControl))

        'btnRepairAllComps initial props
        btnRepairAllComps = New UIButton(oUILib)
        With btnRepairAllComps
            .ControlName = "btnRepairAllComps"
            .Left = 280
            .Top = 145
            .Width = 110
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Repair All"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to repair all damaged components."
        End With
        Me.AddChild(CType(btnRepairAllComps, UIControl))

        'btnRadar initial props
        btnRadar = New UIButton(oUILib)
        With btnRadar
            .ControlName = "btnRadar"
            .Left = 250
            .Top = 29
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to repair the Radar."
        End With
        Me.AddChild(CType(btnRadar, UIControl))

        'btnShield initial props
        btnShield = New UIButton(oUILib)
        With btnShield
            .ControlName = "btnShield"
            .Left = 250
            .Top = 54
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to repair the Shields."
        End With
        Me.AddChild(CType(btnShield, UIControl))

        'btnHangar initial props
        btnHangar = New UIButton(oUILib)
        With btnHangar
            .ControlName = "btnHangar"
            .Left = 250
            .Top = 79
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to repair the Hangar Bay."
        End With
        Me.AddChild(CType(btnHangar, UIControl))

        'btnCargo initial props
        btnCargo = New UIButton(oUILib)
        With btnCargo
            .ControlName = "btnCargo"
            .Left = 250
            .Top = 104
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to repair the Cargo Bay."
        End With
        Me.AddChild(CType(btnCargo, UIControl))

        ''btnFuelBay initial props
        'btnFuelBay = New UIButton(oUILib)
        'With btnFuelBay
        '    .ControlName = "btnFuelBay"
        '    .Left = 250
        '    .Top = 129
        '    .Width = 23
        '    .Height = 24
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "R"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = True
        '    .FontFormat = CType(5, DrawTextFormat)
        '    .ControlImageRect = New Rectangle(0, 0, 120, 32)
        '    .ToolTipText = "Click to repair the Fuel Bay."
        'End With
        'Me.AddChild(CType(btnFuelBay, UIControl))

        'btnFore1 initial props
        btnFore1 = New UIButton(oUILib)
        With btnFore1
            .ControlName = "btnFore1"
            .Left = 375
            .Top = 28
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to repair the Forward Weapons (group 1)."
        End With
        Me.AddChild(CType(btnFore1, UIControl))

        'btnFore2 initial props
        btnFore2 = New UIButton(oUILib)
        With btnFore2
            .ControlName = "btnFore2"
            .Left = 480
            .Top = 28
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to repair the Forward Weapons (group 2)."
        End With
        Me.AddChild(CType(btnFore2, UIControl))

        'btnLeft1 initial props
        btnLeft1 = New UIButton(oUILib)
        With btnLeft1
            .ControlName = "btnLeft1"
            .Left = 375
            .Top = 53
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to repair the Left Weapons (group 1)."
        End With
        Me.AddChild(CType(btnLeft1, UIControl))

        'btnLeft2 initial props
        btnLeft2 = New UIButton(oUILib)
        With btnLeft2
            .ControlName = "btnLeft2"
            .Left = 480
            .Top = 53
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to repair the Left Weapons (group 2)."
        End With
        Me.AddChild(CType(btnLeft2, UIControl))

        'btnRear1 initial props
        btnRear1 = New UIButton(oUILib)
        With btnRear1
            .ControlName = "btnRear1"
            .Left = 375
            .Top = 78
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to repair the Rear Weapons (group 1)."
        End With
        Me.AddChild(CType(btnRear1, UIControl))

        'btnRear2 initial props
        btnRear2 = New UIButton(oUILib)
        With btnRear2
            .ControlName = "btnRear2"
            .Left = 480
            .Top = 78
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to repair the Rear Weapons (group 2)."
        End With
        Me.AddChild(CType(btnRear2, UIControl))

        'btnRight1 initial props
        btnRight1 = New UIButton(oUILib)
        With btnRight1
            .ControlName = "btnRight1"
            .Left = 375
            .Top = 103
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to repair the Right Weapons (group 1)."
        End With
        Me.AddChild(CType(btnRight1, UIControl))

        'btnRight2 initial props
        btnRight2 = New UIButton(oUILib)
        With btnRight2
            .ControlName = "btnRight2"
            .Left = 480
            .Top = 103
            .Width = 23
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "R"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Click to repair the Right Weapons (group 2)."
        End With
        Me.AddChild(CType(btnRight2, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)
    End Sub

    Public Sub SetFromEntity(ByVal lParentID As Int32, ByVal iParentTypeID As Int16, ByVal lEntityID As Int32, ByVal iEntityTypeID As Int16)

        mlParentEntityIdx = -1

        mlParentID = lParentID
        miParentTypeID = iParentTypeID
        mlEntityID = lEntityID
        miEntityTypeID = iEntityTypeID

        If goCurrentEnvir Is Nothing = False Then
            For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                If goCurrentEnvir.lEntityIdx(X) = mlParentID AndAlso goCurrentEnvir.oEntity(X).ObjTypeID = iParentTypeID Then
                    mlParentEntityIdx = X
                    Exit For
                End If
            Next X
        End If

        Dim yMsg(7) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eRequestEntityDefenses).CopyTo(yMsg, 0)
        System.BitConverter.GetBytes(lEntityID).CopyTo(yMsg, 2)
        System.BitConverter.GetBytes(iEntityTypeID).CopyTo(yMsg, 6)
        MyBase.moUILib.SendMsgToPrimary(yMsg)
        mlLastDefenseRequest = glCurrentCycle

    End Sub

    Public Sub HandleDefensesMsg(ByRef yData() As Byte)
        Dim lPos As Int32 = 2
        Dim lObjID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
        Dim iObjTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

        If mlEntityID = lObjID AndAlso miEntityTypeID = iObjTypeID Then
            Dim lArmorHP(3) As Int32
            For X As Int32 = 0 To 3
                lArmorHP(X) = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Next X
            Dim lStructureHP As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lShieldHP As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lStatus As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lDefStatus As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim lDefID As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
            Dim iDefTypeID As Int16 = System.BitConverter.ToInt16(yData, lPos) : lPos += 2

            Dim oEntityDef As EntityDef = Nothing
            For X As Int32 = 0 To glEntityDefUB
                If glEntityDefIdx(X) = lDefID AndAlso goEntityDefs(X).ObjTypeID = iDefTypeID Then
                    oEntityDef = goEntityDefs(X)
                    Exit For
                End If
            Next X

            If mbRequestedArchived = False Then
                If oEntityDef Is Nothing Then
                    MyBase.moUILib.GetMsgSys.LoadArchived()
                    mbRequestedArchived = True
                    Return
                Else
                    For X As Int32 = 0 To oEntityDef.ProductionCost.ItemCostUB
                        If oEntityDef.ProductionCost.ItemCosts(X).ItemTypeID = ObjectType.eArmorTech Then
                            Dim oTech As Base_Tech = goCurrentPlayer.GetTech(oEntityDef.ProductionCost.ItemCosts(X).ItemID, ObjectType.eArmorTech)
                            If oTech Is Nothing = True Then
                                MyBase.moUILib.GetMsgSys.LoadArchived()
                                mbRequestedArchived = True
                                Return
                            End If
                        End If
                    Next X
                End If
            ElseIf oEntityDef Is Nothing Then
                Return
            End If

            Dim fMult As Single

            If oEntityDef.Armor_MaxHP(UnitArcs.eForwardArc) = 0 Then
                fMult = 0
            Else
                fMult = CSng(lArmorHP(UnitArcs.eForwardArc) / oEntityDef.Armor_MaxHP(UnitArcs.eForwardArc))
            End If
            shpFront.Height = CInt(fMult * ml_FULL_HP_HEIGHT)
            shpFront.Top = ml_HP_BAR_BOTTOM - shpFront.Height
            shpFront.ToolTipText = "Front: " & lArmorHP(UnitArcs.eForwardArc) & " / " & oEntityDef.Armor_MaxHP(UnitArcs.eForwardArc)

            btnFront.Visible = lArmorHP(UnitArcs.eForwardArc) <> oEntityDef.Armor_MaxHP(UnitArcs.eForwardArc)
            If btnFront.Visible = True Then
                btnFront.ToolTipText = GetArmorRepairCostsToolTip(oEntityDef, oEntityDef.Armor_MaxHP(UnitArcs.eForwardArc), lArmorHP(UnitArcs.eForwardArc), "Click to Repair the Armor on the Front Side")
            End If

            If oEntityDef.Armor_MaxHP(UnitArcs.eLeftArc) = 0 Then
                fMult = 0
            Else
                fMult = CSng(lArmorHP(UnitArcs.eLeftArc) / oEntityDef.Armor_MaxHP(UnitArcs.eLeftArc))
            End If
            shpLeft.Height = CInt(fMult * ml_FULL_HP_HEIGHT)
            shpLeft.Top = ml_HP_BAR_BOTTOM - shpLeft.Height
            shpLeft.ToolTipText = "Left: " & lArmorHP(UnitArcs.eLeftArc) & " / " & oEntityDef.Armor_MaxHP(UnitArcs.eLeftArc)

            btnLeft.Visible = lArmorHP(UnitArcs.eLeftArc) <> oEntityDef.Armor_MaxHP(UnitArcs.eLeftArc)
            If btnLeft.Visible = True Then
                btnLeft.ToolTipText = GetArmorRepairCostsToolTip(oEntityDef, oEntityDef.Armor_MaxHP(UnitArcs.eLeftArc), lArmorHP(UnitArcs.eLeftArc), "Click to Repair the Armor on the Left Side")
            End If

            If oEntityDef.Armor_MaxHP(UnitArcs.eRightArc) = 0 Then
                fMult = 0
            Else
                fMult = CSng(lArmorHP(UnitArcs.eRightArc) / oEntityDef.Armor_MaxHP(UnitArcs.eRightArc))
            End If
            shpRight.Height = CInt(fMult * ml_FULL_HP_HEIGHT)
            shpRight.Top = ml_HP_BAR_BOTTOM - shpRight.Height
            shpRight.ToolTipText = "Right: " & lArmorHP(UnitArcs.eRightArc) & " / " & oEntityDef.Armor_MaxHP(UnitArcs.eRightArc)

            btnRight.Visible = lArmorHP(UnitArcs.eRightArc) <> oEntityDef.Armor_MaxHP(UnitArcs.eRightArc)
            If btnRight.Visible = True Then
                btnRight.ToolTipText = GetArmorRepairCostsToolTip(oEntityDef, oEntityDef.Armor_MaxHP(UnitArcs.eRightArc), lArmorHP(UnitArcs.eRightArc), "Click to Repair the Armor on the Right Side")
            End If

            If oEntityDef.Armor_MaxHP(UnitArcs.eBackArc) = 0 Then
                fMult = 0
            Else
                fMult = CSng(lArmorHP(UnitArcs.eBackArc) / oEntityDef.Armor_MaxHP(UnitArcs.eBackArc))
            End If
            shpBack.Height = CInt(fMult * ml_FULL_HP_HEIGHT)
            shpBack.Top = ml_HP_BAR_BOTTOM - shpBack.Height
            shpBack.ToolTipText = "Back: " & lArmorHP(UnitArcs.eBackArc) & " / " & oEntityDef.Armor_MaxHP(UnitArcs.eBackArc)

            btnBack.Visible = lArmorHP(UnitArcs.eBackArc) <> oEntityDef.Armor_MaxHP(UnitArcs.eBackArc)
            If btnBack.Visible = True Then
                btnBack.ToolTipText = GetArmorRepairCostsToolTip(oEntityDef, oEntityDef.Armor_MaxHP(UnitArcs.eBackArc), lArmorHP(UnitArcs.eBackArc), "Click to Repair the Armor on the Back Side")
            End If

            If oEntityDef.Structure_MaxHP = 0 Then
                fMult = 0
            Else
                fMult = CSng(lStructureHP / oEntityDef.Structure_MaxHP)
            End If
            fMult = CSng(lStructureHP / oEntityDef.Structure_MaxHP)
            shpStructure.Height = CInt(fMult * ml_FULL_HP_HEIGHT)
            shpStructure.Top = ml_HP_BAR_BOTTOM - shpStructure.Height
            shpStructure.ToolTipText = "Structure: " & lStructureHP & " / " & oEntityDef.Structure_MaxHP

            btnStructure.Visible = lStructureHP <> oEntityDef.Structure_MaxHP
            If btnStructure.Visible = True Then
                btnStructure.ToolTipText = "Click to Repair the Structural Integrity" & vbCrLf & vbCrLf & "Estimated costs for the repairs:" & vbCrLf & _
                  ((oEntityDef.Structure_MaxHP - lStructureHP) * 1000).ToString("#,##0") & " credits" & vbCrLf & "Materials required to replace shattered hull frame"
            End If

            If btnFront.Visible = False AndAlso btnLeft.Visible = False AndAlso btnRight.Visible = False AndAlso btnBack.Visible = False AndAlso btnStructure.Visible = False Then
                btnRepairAll.Visible = False
            Else
                btnRepairAll.Visible = True
                Dim lCurrentArmor As Int32 = 0
                Dim lMaxArmor As Int32 = 0
                For X As Int32 = 0 To oEntityDef.Armor_MaxHP.GetUpperBound(0)
                    lMaxArmor += oEntityDef.Armor_MaxHP(X)
                Next X
                For X As Int32 = 0 To lArmorHP.GetUpperBound(0)
                    lCurrentArmor += lArmorHP(X)
                Next X

                Dim sArmorMessage As String = GetArmorRepairCostsToolTip(oEntityDef, lMaxArmor, lCurrentArmor, "Click to Repair the Armor and Structure")

                Dim lDiff As Int32 = oEntityDef.Structure_MaxHP - lStructureHP
                If lDiff <> 0 Then
                    Dim lCreditStart As Int32 = sArmorMessage.IndexOf("credits")
                    If lCreditStart <> -1 Then
                        Dim lPrevVBCRLF As Int32 = sArmorMessage.LastIndexOf(vbCrLf, lCreditStart)
                        If lPrevVBCRLF <> -1 Then
                            Dim lCredits As Int32 = CInt(Val(Replace$(Replace$(sArmorMessage.Substring(lPrevVBCRLF, lCreditStart - lPrevVBCRLF), vbCrLf, ""), ",", "")))
                            Dim sReplaceStr As String = lCredits.ToString("#,##0") & " credits"

                            lCredits += (lDiff * 1000)
                            sArmorMessage = Replace$(sArmorMessage, sReplaceStr, lCredits.ToString("#,##0") & " credits")
                            sArmorMessage &= vbCrLf & "Materials required to replace shattered hull frame"
                        End If
                    End If
                End If
                btnRepairAll.ToolTipText = sArmorMessage
            End If

            'Now, for the statuses
            Dim lTemp As Int32
            Dim oLbl As UILabel
            Dim oBtn As UIButton

            'Engine
            lTemp = elUnitStatus.eEngineOperational
            oLbl = lblEngines
            oBtn = btnEngines

            If (lDefStatus And lTemp) <> 0 Then
                If (lStatus And lTemp) <> 0 Then
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)      'operational
                    oBtn.Visible = False
                Else
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)      'damaged
                    oBtn.Visible = True
                End If
            Else
                oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                oBtn.Visible = False
            End If
            If oBtn.Visible = True Then
                oBtn.ToolTipText = GetComponentRepairCostsToolTip(oEntityDef, ObjectType.eEngineTech)
            End If

            'Radar
            lTemp = elUnitStatus.eRadarOperational
            oLbl = lblRadar
            oBtn = btnRadar

            If (lDefStatus And lTemp) <> 0 Then
                If (lStatus And lTemp) <> 0 Then
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)      'operational
                    oBtn.Visible = False
                Else
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)      'damaged
                    oBtn.Visible = True
                End If
            Else
                oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                oBtn.Visible = False
            End If
            If oBtn.Visible = True Then
                oBtn.ToolTipText = GetComponentRepairCostsToolTip(oEntityDef, ObjectType.eRadarTech)
            End If

            'Shields
            lTemp = elUnitStatus.eShieldOperational
            oLbl = lblShields
            oBtn = btnShield

            If (lDefStatus And lTemp) <> 0 Then
                If (lStatus And lTemp) <> 0 Then
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)      'operational
                    oBtn.Visible = False
                Else
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)      'damaged
                    oBtn.Visible = True
                End If
            Else
                oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                oBtn.Visible = False
            End If
            If oBtn.Visible = True Then
                oBtn.ToolTipText = GetComponentRepairCostsToolTip(oEntityDef, ObjectType.eShieldTech)
            End If

            'Cargo Bay
            lTemp = elUnitStatus.eCargoBayOperational
            oLbl = lblCargo
            oBtn = btnCargo

            If (lDefStatus And lTemp) <> 0 Then
                If (lStatus And lTemp) <> 0 Then
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)      'operational
                    oBtn.Visible = False
                Else
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)      'damaged
                    oBtn.Visible = True
                End If
            Else
                oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                oBtn.Visible = False
            End If
            If oBtn.Visible = True Then
                oBtn.ToolTipText = "Click to repair the Cargo Bay" & vbCrLf & vbCrLf & "Estimated cost for the repairs:" & vbCrLf & "100000 credits"
            End If

            'Hangar
            lTemp = elUnitStatus.eHangarOperational
            oLbl = lblHangar
            oBtn = btnHangar

            If (lDefStatus And lTemp) <> 0 Then
                If (lStatus And lTemp) <> 0 Then
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)      'operational
                    oBtn.Visible = False
                Else
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)      'damaged
                    oBtn.Visible = True
                End If
            Else
                oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                oBtn.Visible = False
            End If
            If oBtn.Visible = True Then
                oBtn.ToolTipText = "Click to repair the Cargo Bay" & vbCrLf & vbCrLf & "Estimated cost for the repairs:" & vbCrLf & "100000 credits"
            End If

            'Aft Weapon 1
            lTemp = elUnitStatus.eAftWeapon1
            oLbl = lblRear1
            oBtn = btnRear1

            If (lDefStatus And lTemp) <> 0 Then
                If (lStatus And lTemp) <> 0 Then
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)      'operational
                    oBtn.Visible = False
                Else
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)      'damaged
                    oBtn.Visible = True
                End If
            Else
                oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                oBtn.Visible = False
            End If
            If oBtn.Visible = True Then
                oBtn.ToolTipText = GetWeaponSectionRepairCostsToolTip(oEntityDef, lTemp)
            End If

            'Aft weapon 2
            lTemp = elUnitStatus.eAftWeapon2
            oLbl = lblRear2
            oBtn = btnRear2

            If (lDefStatus And lTemp) <> 0 Then
                If (lStatus And lTemp) <> 0 Then
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)      'operational
                    oBtn.Visible = False
                Else
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)      'damaged
                    oBtn.Visible = True
                End If
            Else
                oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                oBtn.Visible = False
            End If
            If oBtn.Visible = True Then
                oBtn.ToolTipText = GetWeaponSectionRepairCostsToolTip(oEntityDef, lTemp)
            End If

            'Front Weapon 1
            lTemp = elUnitStatus.eForwardWeapon1
            oLbl = lblFore1
            oBtn = btnFore1

            If (lDefStatus And lTemp) <> 0 Then
                If (lStatus And lTemp) <> 0 Then
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)      'operational
                    oBtn.Visible = False
                Else
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)      'damaged
                    oBtn.Visible = True
                End If
            Else
                oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                oBtn.Visible = False
            End If
            If oBtn.Visible = True Then
                oBtn.ToolTipText = GetWeaponSectionRepairCostsToolTip(oEntityDef, lTemp)
            End If

            'Front 2
            lTemp = elUnitStatus.eForwardWeapon2
            oLbl = lblFore2
            oBtn = btnFore2

            If (lDefStatus And lTemp) <> 0 Then
                If (lStatus And lTemp) <> 0 Then
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)      'operational
                    oBtn.Visible = False
                Else
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)      'damaged
                    oBtn.Visible = True
                End If
            Else
                oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                oBtn.Visible = False
            End If
            If oBtn.Visible = True Then
                oBtn.ToolTipText = GetWeaponSectionRepairCostsToolTip(oEntityDef, lTemp)
            End If

            'Left 1
            lTemp = elUnitStatus.eLeftWeapon1
            oLbl = lblLeft1
            oBtn = btnLeft1

            If (lDefStatus And lTemp) <> 0 Then
                If (lStatus And lTemp) <> 0 Then
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)      'operational
                    oBtn.Visible = False
                Else
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)      'damaged
                    oBtn.Visible = True
                End If
            Else
                oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                oBtn.Visible = False
            End If
            If oBtn.Visible = True Then
                oBtn.ToolTipText = GetWeaponSectionRepairCostsToolTip(oEntityDef, lTemp)
            End If

            'Left 2
            lTemp = elUnitStatus.eLeftWeapon2
            oLbl = lblLeft2
            oBtn = btnLeft2

            If (lDefStatus And lTemp) <> 0 Then
                If (lStatus And lTemp) <> 0 Then
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)      'operational
                    oBtn.Visible = False
                Else
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)      'damaged
                    oBtn.Visible = True
                End If
            Else
                oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                oBtn.Visible = False
            End If
            If oBtn.Visible = True Then
                oBtn.ToolTipText = GetWeaponSectionRepairCostsToolTip(oEntityDef, lTemp)
            End If

            'Right 1
            lTemp = elUnitStatus.eRightWeapon1
            oLbl = lblRight1
            oBtn = btnRight1

            If (lDefStatus And lTemp) <> 0 Then
                If (lStatus And lTemp) <> 0 Then
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)      'operational
                    oBtn.Visible = False
                Else
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)      'damaged
                    oBtn.Visible = True
                End If
            Else
                oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                oBtn.Visible = False
            End If
            If oBtn.Visible = True Then
                oBtn.ToolTipText = GetWeaponSectionRepairCostsToolTip(oEntityDef, lTemp)
            End If

            'Right 2
            lTemp = elUnitStatus.eRightWeapon2
            oLbl = lblRight2
            oBtn = btnRight2

            If (lDefStatus And lTemp) <> 0 Then
                If (lStatus And lTemp) <> 0 Then
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)      'operational
                    oBtn.Visible = False
                Else
                    oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)      'damaged
                    oBtn.Visible = True
                End If
            Else
                oLbl.ForeColor = System.Drawing.Color.FromArgb(255, 128, 128, 128)
                oBtn.Visible = False
            End If
            If oBtn.Visible = True Then
                oBtn.ToolTipText = GetWeaponSectionRepairCostsToolTip(oEntityDef, lTemp)
            End If

            'Now, for the final repair all
            btnRepairAllComps.Visible = btnEngines.Visible OrElse btnShield.Visible OrElse btnRadar.Visible OrElse btnHangar.Visible OrElse _
              btnCargo.Visible OrElse btnFore1.Visible OrElse btnFore2.Visible OrElse btnLeft1.Visible OrElse btnLeft2.Visible OrElse _
              btnRear1.Visible OrElse btnRear2.Visible OrElse btnRight1.Visible OrElse btnRight2.Visible

        End If
    End Sub

    Private Sub CreateAndSubmitRepairMessage(ByVal lRepairItem As Int32)
        Dim yMsg(17) As Byte
        Dim lPos As Int32 = 0

        System.BitConverter.GetBytes(GlobalMessageCode.eRepairOrder).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(mlParentID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(miParentTypeID).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(mlEntityID).CopyTo(yMsg, lPos) : lPos += 4
        System.BitConverter.GetBytes(miEntityTypeID).CopyTo(yMsg, lPos) : lPos += 2
        System.BitConverter.GetBytes(lRepairItem).CopyTo(yMsg, lPos) : lPos += 4

        MyBase.moUILib.SendMsgToPrimary(yMsg)

    End Sub

    Private Function GetArmorRepairCostsToolTip(ByRef oEntityDef As EntityDef, ByVal lMaxValue As Int32, ByVal lCurrentValue As Int32, ByVal sRepairWhat As String) As String
        'Ok, let's get our costs
        Dim lDiff As Int32 = lMaxValue - lCurrentValue
        Dim sTemp As String = sRepairWhat & vbCrLf & vbCrLf & "Estimated costs for repairs:" & vbCrLf

        '1000 is the default cost per plate
        sTemp &= (lDiff * 1000).ToString("#,##0") & " credits" & vbCrLf

        If oEntityDef.ProductionCost Is Nothing = False Then
            Dim bFound As Boolean = False
            For X As Int32 = 0 To oEntityDef.ProductionCost.ItemCostUB
                If oEntityDef.ProductionCost.ItemCosts(X).ItemTypeID = ObjectType.eArmorTech Then
                    Dim oTech As Base_Tech = goCurrentPlayer.GetTech(oEntityDef.ProductionCost.ItemCosts(X).ItemID, ObjectType.eArmorTech)
                    If oTech Is Nothing = False Then
                        bFound = True
                        With CType(oTech, ArmorTech)
                            Dim lPlatesNeeded As Int32 = CInt(Math.Ceiling(lDiff / .HPPerPlate))
                            sTemp &= lPlatesNeeded.ToString("#,##0") & " " & .sArmorName & " replacement armor plates"
                        End With
                    End If
                    Exit For
                End If
            Next X
            If bFound = False Then sTemp &= "Any Armor Plates that need to be replaced"
        Else
            sTemp &= "Any Armor Plates that need to be replaced"
        End If
        Return sTemp
    End Function

    Private Function GetComponentRepairCostsToolTip(ByRef oEntityDef As EntityDef, ByVal iComponentTypeID As Int16) As String
        Dim lCreditCost As Int32 = 0
        Dim sComponentName As String = "Replacement Components"

        Select Case iComponentTypeID
            Case ObjectType.eEngineTech
                lCreditCost = 100000
            Case ObjectType.eShieldTech
                lCreditCost = 160000
            Case Else
                lCreditCost = 100000
        End Select
        If oEntityDef.ProductionCost Is Nothing = False Then
            For X As Int32 = 0 To oEntityDef.ProductionCost.ItemCostUB
                If oEntityDef.ProductionCost.ItemCosts(X).ItemTypeID = iComponentTypeID Then
                    Dim oTech As Base_Tech = goCurrentPlayer.GetTech(oEntityDef.ProductionCost.ItemCosts(X).ItemID, oEntityDef.ProductionCost.ItemCosts(X).ItemTypeID)
                    If oTech Is Nothing = False Then
                        sComponentName = "Replacement " & oTech.GetComponentName() & " (1)"
                    End If
                    Exit For
                End If
            Next X
        End If

        Dim sWhatToReplace As String = "Click to repair the " & Base_Tech.GetComponentTypeName(iComponentTypeID) & vbCrLf & vbCrLf


        'epica_tech.GetComponentTypeName(
        Return sWhatToReplace & "Estimated Costs for Repairs:" & vbCrLf & lCreditCost.ToString("#,##0") & " credits " & vbCrLf & sComponentName
    End Function

    Private Function GetWeaponSectionRepairCostsToolTip(ByRef oEntityDef As EntityDef, ByVal lStatus As Int32) As String
        Dim sFinal As String = "Click to repair this weapon group" & vbCrLf & vbCrLf & _
                               "Estimated costs for the repairs:" & vbCrLf

        Dim lTotalCredits As Int32 = 0

        Dim lWpnUB As Int32 = -1
        Dim lWpnDefID() As Int32 = Nothing
        Dim sWpnDefName() As String = Nothing
        Dim lWpnCnt() As Int32 = Nothing

        For X As Int32 = 0 To oEntityDef.WeaponDefUB
            If oEntityDef.WeaponDefs(X).lEntityStatusGroup = lStatus Then
                lTotalCredits += 500000
                Dim bFound As Boolean = False
                For Y As Int32 = 0 To lWpnUB
                    If lWpnDefID(Y) = oEntityDef.WeaponDefs(X).WpnDefID Then
                        lWpnCnt(Y) += 1
                        bFound = True
                        Exit For
                    End If
                Next Y
                If bFound = False Then
                    lWpnUB += 1
                    ReDim Preserve lWpnDefID(lWpnUB)
                    ReDim Preserve sWpnDefName(lWpnUB)
                    ReDim Preserve lWpnCnt(lWpnUB)
                    lWpnDefID(lWpnUB) = oEntityDef.WeaponDefs(X).WpnDefID
                    sWpnDefName(lWpnUB) = oEntityDef.WeaponDefs(X).WeaponName
                    lWpnCnt(lWpnUB) = 1
                End If
            End If
        Next X

        'Now, append our costs
        sFinal &= lTotalCredits.ToString("#,##0") & " credits" & vbCrLf
        For X As Int32 = 0 To lWpnUB
            If lWpnCnt(X) = 1 Then
                sFinal &= sWpnDefName(X) & " Replacement Weapon" & vbCrLf
            Else
                sFinal &= lWpnCnt(X).ToString("#,##0") & " " & sWpnDefName(X) & " Replacements" & vbCrLf
            End If
        Next X
        Return sFinal
    End Function

#Region " Button Clicks "
	Private Sub btnBack_Click(ByVal sName As String) Handles btnBack.Click
		CreateAndSubmitRepairMessage(-(CInt(UnitArcs.eBackArc) + 2))
	End Sub

	Private Sub btnCancelRepair_Click(ByVal sName As String) Handles btnCancelRepair.Click
        Dim yData(17) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityProduction).CopyTo(yData, 0)
        System.BitConverter.GetBytes(mlParentID).CopyTo(yData, 2)
        System.BitConverter.GetBytes(miParentTypeID).CopyTo(yData, 6)
        System.BitConverter.GetBytes(-1I).CopyTo(yData, 8)
        System.BitConverter.GetBytes(-1S).CopyTo(yData, 12)
        System.BitConverter.GetBytes(-1I).CopyTo(yData, 14)

        MyBase.moUILib.SendMsgToPrimary(yData)
	End Sub

	Private Sub btnCargo_Click(ByVal sName As String) Handles btnCargo.Click
		CreateAndSubmitRepairMessage(elUnitStatus.eCargoBayOperational)
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnEngines_Click(ByVal sName As String) Handles btnEngines.Click
		CreateAndSubmitRepairMessage(elUnitStatus.eEngineOperational)
	End Sub

	Private Sub btnFore1_Click(ByVal sName As String) Handles btnFore1.Click
		CreateAndSubmitRepairMessage(elUnitStatus.eForwardWeapon1)
	End Sub

	Private Sub btnFore2_Click(ByVal sName As String) Handles btnFore2.Click
		CreateAndSubmitRepairMessage(elUnitStatus.eForwardWeapon2)
	End Sub

	Private Sub btnFront_Click(ByVal sName As String) Handles btnFront.Click
		CreateAndSubmitRepairMessage(-(CInt(UnitArcs.eForwardArc) + 2))
	End Sub

	'Private Sub btnFuelBay_Click(ByVal sName As String) Handles btnFuelBay.Click
    '    CreateAndSubmitRepairMessage(elUnitStatus.eFuelBayOperational)
    'End Sub

	Private Sub btnHangar_Click(ByVal sName As String) Handles btnHangar.Click
		CreateAndSubmitRepairMessage(elUnitStatus.eHangarOperational)
	End Sub

	Private Sub btnLeft_Click(ByVal sName As String) Handles btnLeft.Click
		CreateAndSubmitRepairMessage(-(CInt(UnitArcs.eLeftArc) + 2))
	End Sub

	Private Sub btnLeft1_Click(ByVal sName As String) Handles btnLeft1.Click
		CreateAndSubmitRepairMessage(elUnitStatus.eLeftWeapon1)
	End Sub

	Private Sub btnLeft2_Click(ByVal sName As String) Handles btnLeft2.Click
		CreateAndSubmitRepairMessage(elUnitStatus.eLeftWeapon2)
	End Sub

	Private Sub btnRadar_Click(ByVal sName As String) Handles btnRadar.Click
		CreateAndSubmitRepairMessage(elUnitStatus.eRadarOperational)
	End Sub

	Private Sub btnRear1_Click(ByVal sName As String) Handles btnRear1.Click
		CreateAndSubmitRepairMessage(elUnitStatus.eAftWeapon1)
	End Sub

	Private Sub btnRear2_Click(ByVal sName As String) Handles btnRear2.Click
		CreateAndSubmitRepairMessage(elUnitStatus.eAftWeapon2)
	End Sub

	Private Sub btnRepairAll_Click(ByVal sName As String) Handles btnRepairAll.Click
		CreateAndSubmitRepairMessage(-7I)
	End Sub

	Private Sub btnRepairAllComps_Click(ByVal sName As String) Handles btnRepairAllComps.Click
		CreateAndSubmitRepairMessage(Int32.MaxValue)
	End Sub

	Private Sub btnRight_Click(ByVal sName As String) Handles btnRight.Click
		CreateAndSubmitRepairMessage(-(CInt(UnitArcs.eRightArc) + 2))
	End Sub

	Private Sub btnRight1_Click(ByVal sName As String) Handles btnRight1.Click
		CreateAndSubmitRepairMessage(elUnitStatus.eRightWeapon1)
	End Sub

	Private Sub btnRight2_Click(ByVal sName As String) Handles btnRight2.Click
		CreateAndSubmitRepairMessage(elUnitStatus.eRightWeapon2)
	End Sub

	Private Sub btnShield_Click(ByVal sName As String) Handles btnShield.Click
		CreateAndSubmitRepairMessage(elUnitStatus.eShieldOperational)
	End Sub

	Private Sub btnStructure_Click(ByVal sName As String) Handles btnStructure.Click
		CreateAndSubmitRepairMessage(-6I)
	End Sub
#End Region

    Private Sub frmRepair_OnNewFrameEnd() Handles Me.OnNewFrameEnd
        If glCurrentCycle - mlLastDefenseRequest > 90 Then
            Dim yMsg(7) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eRequestEntityDefenses).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(mlEntityID).CopyTo(yMsg, 2)
            System.BitConverter.GetBytes(miEntityTypeID).CopyTo(yMsg, 6)
            MyBase.moUILib.SendMsgToPrimary(yMsg)
            mlLastDefenseRequest = glCurrentCycle
        End If
    End Sub

	Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
		If goTutorial Is Nothing Then goTutorial = New TutorialManager()
		goTutorial.BeginNewStepType(TutorialManager.TutorialStepType.eRepairing)
	End Sub
End Class