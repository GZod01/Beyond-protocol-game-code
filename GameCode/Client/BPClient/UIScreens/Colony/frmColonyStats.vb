Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmColonyStats
    Inherits UIWindow

    Private Const ml_REQUEST_DELAY As Int32 = 5000

    Private WithEvents lblPopulation As UILabel
    Private WithEvents lblJobs As UILabel
    Private WithEvents lblUnemployment As UILabel
    Private WithEvents lblPowered As UILabel
    Private WithEvents lblHousing As UILabel
    Private WithEvents lblUnpowered As UILabel
    Private WithEvents lblTotalHouse As UILabel
    Private WithEvents lblMorale As UILabel
    Private WithEvents lblGrowthRate As UILabel
    Private WithEvents txtName As UITextBox
    Private lnDiv1 As UILine
    Private lnDiv2 As UILine
    Private lnDiv3 As UILine
    Private lnDiv4 As UILine
    Private lnDiv5 As UILine
    Private lnDiv6 As UILine
    Private lnDiv7 As UILine
    Private lnDiv8 As UILine
    Private WithEvents lblTaxRate As UILabel
    Private WithEvents txtTaxRate As UITextBox
    Private WithEvents txtControlGrowth As UITextBox
    Private WithEvents txtControlMorale As UITextBox
    Private WithEvents btnClose As UIButton
    Private WithEvents btnSetTax As UIButton
    Private WithEvents lblEnlisted As UILabel
    Private WithEvents lblOfficers As UILabel
    Private WithEvents lblPowerGen As UILabel
    Private WithEvents lblPowerNeed As UILabel
    Private WithEvents lblExpand As UILabel
    Private lblTotalPowerNeed As UILabel
    Private lblPopIntel As UILabel
    Private lblEfficiency As UILabel

    Private msw_Delay As Stopwatch

    Private mlColonyID As Int32 = -1
    Private mbIgnoreEvents As Boolean = False

    Private mlPopulation As Int32 = 0
    Private mlJobs As Int32 = 0
    Private myUnemployment As Byte = 0
    Private mlPoweredHomes As Int32 = 0
    Private mlUnpoweredHomes As Int32 = 0
    Private mlMorale As Int32
    Private mlGrowthRate As Int32
    Private mlEnlisted As Int32 = 0
    Private mlOfficers As Int32 = 0
    Private mlPowerGen As Int32 = 0
    Private mlPowerNeed As Int32 = 0
    Private mlTotalPowerNeed As Int32 = -1
    Private msColonyName As String = ""
    Private myTaxRate As Byte
    Private myPopIntel As Byte
    Private mlScienceJobs As Int32
    Private miControlledGrowth As Int16
    Private miControlledMorale As Int16

    Private msMoraleRollOver As String = ""
    Private msHousingRollOver As String = ""
    Private mbLoading As Boolean = True

    Private mlResetColonyTaxRateCycle As Int32 = Int32.MinValue

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenColonyStatsWindow)

        Dim lTempTop As Int32
        Dim oWin As UIWindow = oUILib.GetWindow("frmEnvirDisplay")

        If oWin Is Nothing = False Then
            lTempTop = oWin.Top + oWin.Height + 1
        Else : lTempTop = CInt(oUILib.oDevice.PresentationParameters.BackBufferHeight / 2) - 85
        End If
        oWin = Nothing

        If lTempTop + 440 > oUILib.oDevice.PresentationParameters.BackBufferHeight Then lTempTop = oUILib.oDevice.PresentationParameters.BackBufferHeight - 420

        'frmColonyStats initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eColonyStats
            .ControlName = "frmColonyStats"
            '.Left = oUILib.oDevice.PresentationParameters.BackBufferWidth - 170
            '.Top = lTempTop

            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1

            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.ColonyStatsX
                lTop = muSettings.ColonyStatsY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 170
            If lTop < 0 Then lTop = lTempTop
            If lLeft + 170 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - 170
            If lTop + 440 > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - 440

            .Left = lLeft
            .Top = lTop

            .Width = 170
            .Height = 440 '420 '387 '375
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
        End With

        'lblExpand initial props
        lblExpand = New UILabel(oUILib)
        With lblExpand
            .ControlName = "lblExpand"
            .Left = 2
            .Top = 2
            .Width = 166
            .Height = 10
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = grc_UI(elInterfaceRectangle.eColonyStatsMinimizer)
            .bAcceptEvents = True
        End With
        Me.AddChild(CType(lblExpand, UIControl))

        'lblPop initial props
        lblPopulation = New UILabel(oUILib)
        With lblPopulation
            .ControlName = "lblPop"
            .Left = 5
            .Top = 42 '30
            .Width = 141
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Population: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblPopulation, UIControl))

        'lblJobs initial props
        lblJobs = New UILabel(oUILib)
        With lblJobs
            .ControlName = "lblJobs"
            .Left = 5
            .Top = 62 '50
            .Width = 141
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Jobs: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblJobs, UIControl))

        'lblUnemployment initial props
        lblUnemployment = New UILabel(oUILib)
        With lblUnemployment
            .ControlName = "lblUnemployment"
            .Left = 5
            .Top = 82 '70
            .Width = 141
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Unemployment: 0%"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblUnemployment, UIControl))

        'lblPopIntel initial props
        lblPopIntel = New UILabel(oUILib)
        With lblPopIntel
            .ControlName = "lblPopIntel"
            .Left = 5
            .Top = 102
            .Width = 141
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Intelligence: 100"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .ToolTipText = "Range of values is 100 to 160"
        End With
        Me.AddChild(CType(lblPopIntel, UIControl))

        'lblEfficiency initial props
        lblEfficiency = New UILabel(oUILib)
        With lblEfficiency
            .ControlName = "lblEfficiency"
            .Left = 5
            .Top = 122
            .Width = 141
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Efficiency: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .ToolTipText = "Represents the production efficiency of the workforce population" & vbCrLf & _
                           "before any morale penalties or other negative effects are applied." & vbCrLf & _
                           "This number goes less than 100% when the number of jobs supplied" & vbCrLf & _
                           "exceeds the population of the colony meaning that jobs are left unfilled."
        End With
        Me.AddChild(CType(lblEfficiency, UIControl))

        'lblPowered initial props
        lblPowered = New UILabel(oUILib)
        With lblPowered
            .ControlName = "lblPowered"
            .Left = 15
            .Top = 162 '110
            .Width = 141
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Powered: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblPowered, UIControl))

        'lblHousing initial props
        lblHousing = New UILabel(oUILib)
        With lblHousing
            .ControlName = "lblHousing"
            .Left = 5
            .Top = 142 '90
            .Width = 141
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Colony Housing"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .bAcceptEvents = True
        End With
        Me.AddChild(CType(lblHousing, UIControl))

        'lblUnpowered initial props
        lblUnpowered = New UILabel(oUILib)
        With lblUnpowered
            .ControlName = "lblUnpowered"
            .Left = 15
            .Top = 182 '130
            .Width = 150
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Unpowered: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblUnpowered, UIControl))

        'lblTotalHouse initial props
        lblTotalHouse = New UILabel(oUILib)
        With lblTotalHouse
            .ControlName = "lblTotalHouse"
            .Left = 15
            .Top = 202 '150
            .Width = 141
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Total: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblTotalHouse, UIControl))

        'lblMorale initial props
        lblMorale = New UILabel(oUILib)
        With lblMorale
            .ControlName = "lblMorale"
            .Left = 5
            .Top = 232 '180
            .Width = 120
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Morale: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .bAcceptEvents = True
        End With
        Me.AddChild(CType(lblMorale, UIControl))

        'lblGrowthRate initial props
        lblGrowthRate = New UILabel(oUILib)
        With lblGrowthRate
            .ControlName = "lblGrowthRate"
            .Left = 5
            .Top = 252 '200
            .Width = 120
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Growth: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblGrowthRate, UIControl))

        'txtColonyName initial props
        txtName = New UITextBox(oUILib)
        With txtName
            .ControlName = "txtColonyName"
            .Left = 5
            .Top = 17 '5
            .Width = 161
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Unknown Colony"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 20
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtName, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 2
            .Top = 39 '27
            .Width = 167
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lnDiv2 initial props
        lnDiv2 = New UILine(oUILib)
        With lnDiv2
            .ControlName = "lnDiv2"
            .Left = 2
            .Top = 142 '90
            .Width = 167
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv2, UIControl))

        'lnDiv3 initial props
        lnDiv3 = New UILine(oUILib)
        With lnDiv3
            .ControlName = "lnDiv3"
            .Left = 6
            .Top = 158 '106
            .Width = 99
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv3, UIControl))

        'lnDiv4 initial props
        lnDiv4 = New UILine(oUILib)
        With lnDiv4
            .ControlName = "lnDiv4"
            .Left = 2
            .Top = 227 '175
            .Width = 167
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv4, UIControl))

        'lnDiv5 initial props
        lnDiv5 = New UILine(oUILib)
        With lnDiv5
            .ControlName = "lnDiv5"
            .Left = 2
            .Top = 272 '220
            .Width = 167
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv5, UIControl))

        'lblTaxRate initial props
        lblTaxRate = New UILabel(oUILib)
        With lblTaxRate
            .ControlName = "lblTaxRate"
            .Left = 5
            .Top = 277 '225
            .Width = 58
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Tax Rate:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblTaxRate, UIControl))

        'txtTaxRate initial props
        txtTaxRate = New UITextBox(oUILib)
        With txtTaxRate
            .ControlName = "txtTaxRate"
            .Left = 74
            .Top = 277 '225
            .Width = 43
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "0%"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 4
            .BorderColor = muSettings.InterfaceBorderColor
            .bNumericOnly = True
        End With
        Me.AddChild(CType(txtTaxRate, UIControl))

        'txtControlGrowth initial props
        txtControlGrowth = New UITextBox(oUILib)
        With txtControlGrowth
            .ControlName = "txtControlGrowth"
            .Left = 130
            .Top = 252
            .Width = 33
            .Height = 18
            .Enabled = True
            .Visible = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eControlledGrowth) <> 0
            .Caption = "0"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 4
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtControlGrowth, UIControl))

        'txtControlMorale initial props
        txtControlMorale = New UITextBox(oUILib)
        With txtControlMorale
            .ControlName = "txtControlMorale"
            .Left = 130
            .Top = 232
            .Width = 33
            .Height = 18
            .Enabled = True
            .Visible = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eControlledMorale) <> 0
            .Caption = "0"
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.VerticalCenter
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 4
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtControlMorale, UIControl))

        'lnDiv6 initial props
        lnDiv6 = New UILine(oUILib)
        With lnDiv6
            .ControlName = "lnDiv6"
            .Left = 2
            .Top = 300 '248
            .Width = 167
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv6, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 35
            .Top = 412 '357 '345
            .Width = 100
            .Height = 25
            .Enabled = True
            .Visible = True
            .Caption = "Close"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = DrawTextFormat.VerticalCenter Or DrawTextFormat.Center
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'btnSetTax initial props
        btnSetTax = New UIButton(oUILib)
        With btnSetTax
            .ControlName = "btnSetTax"
            .Left = 121
            .Top = 277 '225
            .Width = 45
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Set"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
        End With
        Me.AddChild(CType(btnSetTax, UIControl))

        'lblEnlisted initial props
        lblEnlisted = New UILabel(oUILib)
        With lblEnlisted
            .ControlName = "lblEnlisted"
            .Left = 5
            .Top = 302 '250
            .Width = 141
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Enlisted: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblEnlisted, UIControl))

        'lblOfficers initial props
        lblOfficers = New UILabel(oUILib)
        With lblOfficers
            .ControlName = "lblOfficers"
            .Left = 5
            .Top = 322 '270
            .Width = 141
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Officers: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblOfficers, UIControl))

        'lnDiv7 initial props
        lnDiv7 = New UILine(oUILib)
        With lnDiv7
            .ControlName = "lnDiv7"
            .Left = 2
            .Top = 342 '290
            .Width = 167
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv7, UIControl))

        'lblPowerGen initial props
        lblPowerGen = New UILabel(oUILib)
        With lblPowerGen
            .ControlName = "lblPowerGen"
            .Left = 5
            .Top = 347 '295
            .Width = 141
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Power Gen: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblPowerGen, UIControl))

        'lblPowerNeed initial props
        lblPowerNeed = New UILabel(oUILib)
        With lblPowerNeed
            .ControlName = "lblPowerNeed"
            .Left = 5
            .Top = 367 '315
            .Width = 154
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Power Need: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblPowerNeed, UIControl))

        'lblTotalPowerNeed initial props
        lblTotalPowerNeed = New UILabel(oUILib)
        With lblTotalPowerNeed
            .ControlName = "lblTotalPowerNeed"
            .Left = 5
            .Top = 387
            .Width = 154
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Total Need: 0"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTotalPowerNeed, UIControl))

        'lnDiv8 initial props
        lnDiv8 = New UILine(oUILib)
        With lnDiv8
            .ControlName = "lnDiv8"
            .Left = 2
            .Top = 407 '352 '340
            .Width = 167
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv8, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        If HasAliasedRights(AliasingRights.eViewColonyStats) = True Then
            MyBase.moUILib.AddWindow(Me)
        Else
            MyBase.moUILib.AddNotification("You lack the rights to view colony statistics.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
            Return
        End If

        muSettings.ExpandedColonyStatsScreen = True
        RefreshDetails()

        'msw_Delay = Stopwatch.StartNew
        mbLoading = False
        If goTutorial Is Nothing = False AndAlso goTutorial.TutorialOn = True Then goTutorial.HandleTutorialTrigger(TutorialManager.TutorialTriggerType.ColonyDetailsFullShown)
    End Sub

    Private Sub RefreshDetails()
        Dim yMsg(7) As Byte
        Try
            System.BitConverter.GetBytes(GlobalMessageCode.eGetColonyDetails).CopyTo(yMsg, 0)
            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                goCurrentEnvir.GetGUIDAsString.CopyTo(yMsg, 2)
            Else
                Dim bFound As Boolean = False
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).bSelected = True AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID Then
                        If goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eSpaceStationSpecial Then
                            goCurrentEnvir.oEntity(X).GetGUIDAsString.CopyTo(yMsg, 2)
                            bFound = True
                        End If
                        Exit For
                    End If
                Next X
                If bFound = False Then
                    Me.Visible = False
                    MyBase.moUILib.RemoveWindow(Me.ControlName)
                    Exit Sub
                End If
            End If

            If msw_Delay Is Nothing Then
                msw_Delay = Stopwatch.StartNew()
            Else
                msw_Delay.Reset()
                msw_Delay.Start()
            End If

            MyBase.moUILib.SendMsgToPrimary(yMsg)
        Catch
            'do nothing
        End Try

    End Sub

    Public Sub HandleColonyMsg(ByVal yData() As Byte)
        Dim lTemp As Int32
        Dim sName As String

        mbIgnoreEvents = True
        Try
            mlColonyID = System.BitConverter.ToInt32(yData, 2)
            'typeid is 2 bytes

            'Population
            Dim bResetEff As Boolean = False
            lTemp = System.BitConverter.ToInt32(yData, 8)
            If lTemp <> mlPopulation Then
                bResetEff = True
                mlPopulation = lTemp
                lblPopulation.Caption = "Population: " & mlPopulation.ToString("#,###,###,##0")
            End If

            'Jobs
            lTemp = System.BitConverter.ToInt32(yData, 12)
            If lTemp <> mlJobs Then
                bResetEff = True
                mlJobs = lTemp
                lblJobs.Caption = "Jobs: " & mlJobs.ToString("#,###,###,##0")
            End If

            If bResetEff = True Then
                Dim sTemp As String
                If mlJobs > 0 Then
                    If mlJobs < mlPopulation Then
                        sTemp = "Efficiency: 100%"
                        If lblEfficiency.ForeColor <> muSettings.InterfaceBorderColor Then lblEfficiency.ForeColor = muSettings.InterfaceBorderColor
                    Else
                        sTemp = "Efficiency: " & ((mlPopulation / mlJobs) * 100).ToString("##0.#") & "%"
                        If lblEfficiency.ForeColor <> System.Drawing.Color.FromArgb(255, 255, 0, 0) Then lblEfficiency.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                    End If
                Else
                    sTemp = "Efficiency: 0"
                End If
                If lblEfficiency.Caption <> sTemp Then lblEfficiency.Caption = sTemp
            End If

            'Enlisted
            lTemp = System.BitConverter.ToInt32(yData, 16)
            If lTemp <> mlEnlisted Then
                mlEnlisted = lTemp
                lblEnlisted.Caption = "Enlisted: " & mlEnlisted.ToString("#,###,###,##0")
            End If

            'Officers
            lTemp = System.BitConverter.ToInt32(yData, 20)
            If lTemp <> mlOfficers Then
                mlOfficers = lTemp
                lblOfficers.Caption = "Officers: " & mlOfficers.ToString("#,###,###,##0")
            End If

            'Power Generated
            lTemp = System.BitConverter.ToInt32(yData, 24)
            If lTemp <> mlPowerGen Then
                mlPowerGen = lTemp
                lblPowerGen.Caption = "Power Gen: " & mlPowerGen.ToString("#,###,###,##0")
            End If

            'Power Need
            lTemp = System.BitConverter.ToInt32(yData, 28)
            If lTemp <> mlPowerNeed Then
                mlPowerNeed = lTemp
                lblPowerNeed.Caption = "Power Need: " & mlPowerNeed.ToString("#,###,###,##0")
            End If

            'Colony Name
            sName = GetStringFromBytes(yData, 32, 20)

            If sName <> msColonyName Then
                txtName.Caption = sName
                msColonyName = sName
            End If
            If txtName.Enabled = False Then
                If NewTutorialManager.TutorialOn = True Then
                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eColonyNameChanged, -1, -1, -1, "")
                End If
            End If
            txtName.Enabled = True

            'Unemployment Rate
            lTemp = yData(52)
            If lTemp <> myUnemployment Then
                myUnemployment = CByte(lTemp)
                lblUnemployment.Caption = "Unemployment: " & myUnemployment & "%"
            End If
            Dim lUnemployedTemp As Int32 = Math.Max(0, mlPopulation - mlJobs)
            If myUnemployment = 0 AndAlso lUnemployedTemp > 0 Then
                Dim sTmpVal As String = "Unemployment: " & CSng(lUnemployedTemp / mlPopulation).ToString("0.0#")
                If lblUnemployment.Caption <> sTmpVal Then lblUnemployment.Caption = sTmpVal
            End If

            'Powered Housing
            Dim bHousing As Boolean = False
            lTemp = System.BitConverter.ToInt32(yData, 53)
            If lTemp <> mlPoweredHomes Then
                mlPoweredHomes = lTemp
                lblPowered.Caption = "Powered: " & mlPoweredHomes.ToString("#,###,###,##0")
                bHousing = True
            End If

            'Unpowered Housing
            lTemp = System.BitConverter.ToInt32(yData, 57)
            If lTemp <> mlUnpoweredHomes Then
                mlUnpoweredHomes = lTemp
                lblUnpowered.Caption = "Unpowered: " & mlUnpoweredHomes.ToString("#,###,###,##0")
                bHousing = True
            End If

            'Controlled Growth
            lTemp = System.BitConverter.ToInt16(yData, 81)
            If lTemp <> miControlledGrowth Then
                miControlledGrowth = CShort(lTemp)
                txtControlGrowth.Caption = miControlledGrowth.ToString
            End If
            'Controlled Morale
            lTemp = System.BitConverter.ToInt16(yData, 83)
            If lTemp <> miControlledMorale Then
                miControlledMorale = CShort(lTemp)
                txtControlMorale.Caption = miControlledMorale.ToString
            End If

            If bHousing = True Then
                lblTotalHouse.Caption = "Total: " & (mlPoweredHomes + mlUnpoweredHomes).ToString("#,###,###,##0")

                Dim clrTemp As System.Drawing.Color = muSettings.InterfaceBorderColor
                Dim sTool As String = ""
                If mlPopulation > 0 AndAlso (mlPoweredHomes > 0 OrElse mlUnpoweredHomes > 0) Then
                    If mlPopulation / (mlPoweredHomes + mlUnpoweredHomes) > 1.14F Then
                        clrTemp = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                        sTool = "You are approaching maximum population" & vbCrLf & _
                                "for the housing you have available."
                    End If
                End If

                Dim lMaxPop As Int32 = mlPoweredHomes + mlUnpoweredHomes
                If sTool <> "" Then sTool &= vbCrLf & vbCrLf
                sTool &= "Max population based on residence: " & CInt(lMaxPop * 1.15F).ToString("#,###,###,##0")

                lblTotalHouse.ToolTipText = sTool
                lblTotalHouse.ForeColor = clrTemp
            End If

            'Colony Morale
            lTemp = System.BitConverter.ToInt32(yData, 61)
            If lTemp <> mlMorale Then
                mlMorale = lTemp
                lblMorale.Caption = "Morale: " & mlMorale
            End If

            'Colony Growth Rate
            lTemp = System.BitConverter.ToInt16(yData, 65)
            If lTemp <> mlGrowthRate Then
                mlGrowthRate = lTemp
                If mlGrowthRate = 0 Then
                    lblGrowthRate.Caption = "Growth: 0"
                    lblGrowthRate.ForeColor = Color.Yellow
                ElseIf mlGrowthRate > 0 Then
                    lblGrowthRate.Caption = "Growth: +" & mlGrowthRate
                    lblGrowthRate.ForeColor = Color.Green
                Else
                    lblGrowthRate.Caption = "Growth: " & mlGrowthRate
                    lblGrowthRate.ForeColor = Color.Red
                End If
            End If

            'Tax Rate
            lTemp = yData(67)
            If lTemp <> myTaxRate Then
                myTaxRate = yData(67)
                txtTaxRate.Caption = myTaxRate & "%"
            End If

            'Fill our roll-over displays...
            Dim bHasCC As Boolean = False
            If goCurrentEnvir.ObjTypeID = ObjectType.ePlanet Then
                For X As Int32 = 0 To goCurrentEnvir.lEntityUB
                    If goCurrentEnvir.lEntityIdx(X) <> -1 AndAlso goCurrentEnvir.oEntity(X).OwnerID = glPlayerID AndAlso _
                      goCurrentEnvir.oEntity(X).ObjTypeID = ObjectType.eFacility AndAlso goCurrentEnvir.oEntity(X).yProductionType = ProductionType.eCommandCenterSpecial Then
                        bHasCC = True
                        Exit For
                    End If
                Next X
            Else : bHasCC = True
            End If

            'Ok, calculate our numbers
            Dim lHomeless As Int32 = Math.Max(0, mlPopulation - (mlPoweredHomes + mlUnpoweredHomes))
            Dim lPowerless As Int32 = Math.Max(0, Math.Min((mlPopulation - mlPoweredHomes), mlUnpoweredHomes))
            Dim lUnemployed As Int32 = Math.Max(0, mlPopulation - mlJobs)
            Dim sFinal As String = ""

            'Housing rollover
            msHousingRollOver = "Housing Breakdown" & vbCrLf & "Powered Residence: " & mlPoweredHomes

            If lPowerless <> 0 Then
                sFinal &= vbCrLf & "Unpowered Residence: " & Math.Round(CSng(lPowerless / mlPopulation) * -50).ToString
                msHousingRollOver &= vbCrLf & "Unpowered Residence: " & lPowerless
            End If
            If lHomeless <> 0 Then
                Dim fHomelessPerc As Single = CSng(lHomeless / mlPopulation)
                If fHomelessPerc * 100 < goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eHomelessAllowedBeforePenalty) Then fHomelessPerc = 0
                sFinal &= vbCrLf & "Homeless: " & Math.Round(fHomelessPerc * -100).ToString
                msHousingRollOver &= vbCrLf & "Homeless Population: " & lHomeless
            End If
            If lUnemployed <> 0 Then
                Dim fUnemploymentPerc As Single = CSng(lUnemployed / mlPopulation)
                If fUnemploymentPerc * 100 < goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eUnemploymentWithoutPenalty) Then fUnemploymentPerc = 0
                sFinal &= vbCrLf & "Unemployment: " & Math.Round(fUnemploymentPerc * -200).ToString
            End If
            If myTaxRate <> 0 Then
                Dim fAdjustedTaxRate As Single = CSng(myTaxRate / 100.0F)
                If myTaxRate < goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eTaxRateWithoutPenalty) Then fAdjustedTaxRate = 0

                Dim lTaxRateMult As Int32 = 10
                Dim lTaxRateMax As Int32 = 80
                If mlPopulation > 2000000 Then
                    lTaxRateMult = 40
                    lTaxRateMax = 20
                ElseIf mlPopulation > 1200000 Then
                    lTaxRateMult = 38
                    lTaxRateMax = 22
                ElseIf mlPopulation > 900000 Then
                    lTaxRateMult = 36
                    lTaxRateMax = 24
                ElseIf mlPopulation > 800000 Then
                    lTaxRateMult = 34
                    lTaxRateMax = 26
                ElseIf mlPopulation > 700000 Then
                    lTaxRateMult = 32
                    lTaxRateMax = 28
                ElseIf mlPopulation > 600000 Then
                    lTaxRateMult = 30
                    lTaxRateMax = 30
                ElseIf mlPopulation > 500000 Then
                    lTaxRateMult = 28
                    lTaxRateMax = 32
                ElseIf mlPopulation > 400000 Then
                    lTaxRateMult = 26
                    lTaxRateMax = 34
                ElseIf mlPopulation > 300000 Then
                    lTaxRateMult = 24
                    lTaxRateMax = 36
                ElseIf mlPopulation > 200000 Then
                    lTaxRateMult = 22
                    lTaxRateMax = 38
                ElseIf mlPopulation > 100000 Then
                    lTaxRateMult = 20
                    lTaxRateMax = 40
                End If

                Dim fNegMod As Single = ((fAdjustedTaxRate * lTaxRateMult) + 1) * (fAdjustedTaxRate * -lTaxRateMult)
                If fNegMod < -72 Then
                    lTaxRateMax = CInt(myTaxRate) - lTaxRateMax
                    If lTaxRateMax > 0 Then
                        Dim fTemp As Single = CSng(Math.Floor(1 + (goCurrentPlayer.blCredits / 1000000000))) * -1.0F
                        fNegMod += (fTemp * lTaxRateMax)
                    End If
                End If
                If fNegMod < 0 Then sFinal &= vbCrLf & "Taxes: " & Math.Round(fNegMod).ToString
            End If
            If bHasCC = False Then sFinal &= vbCrLf & "No Command Center: -30"

            'War Sentiment
            Dim lWarSentiment As Int32 = System.BitConverter.ToInt32(yData, 77)

            If lWarSentiment <> 0 Then
                If lWarSentiment > 0 Then
                    sFinal &= vbCrLf & "Colonists Desire Conquest: -" & Math.Abs(lWarSentiment)
                Else : sFinal &= vbCrLf & "Colonists Desire Peace: -" & Math.Abs(lWarSentiment)
                End If
            End If
            If goCurrentPlayer.BadWarDecMoralePenalty <> 0 Then
                sFinal &= vbCrLf & "Offline War Dec Outrage: -" & Math.Abs(goCurrentPlayer.BadWarDecMoralePenalty)
            End If

            If NewTutorialManager.TutorialOn = True AndAlso goCurrentPlayer.yPlayerPhase = 0 Then
                sFinal &= vbCrLf & "Phase 1 Tutorial: +50"
            End If

            If sFinal.Length <> 0 Then
                msMoraleRollOver = "Negative Morale Effects:" & sFinal
            Else
                msMoraleRollOver = "Negative Morale Effects:" & vbCrLf & "None Reported!"
            End If

            'Power Generated
            lTemp = System.BitConverter.ToInt32(yData, 68)
            If lTemp <> mlTotalPowerNeed Then
                mlTotalPowerNeed = lTemp
                lblTotalPowerNeed.Caption = "Total Need: " & mlTotalPowerNeed.ToString("#,###,###,##0")
            End If
            If mlTotalPowerNeed > mlPowerGen Then
                If lblTotalPowerNeed.ForeColor <> System.Drawing.Color.FromArgb(255, 255, 0, 0) Then
                    lblTotalPowerNeed.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                    Me.IsDirty = True
                End If
            Else
                If lblTotalPowerNeed.ForeColor <> muSettings.InterfaceBorderColor Then
                    lblTotalPowerNeed.ForeColor = muSettings.InterfaceBorderColor
                    Me.IsDirty = True
                End If
            End If

            'Population Intelligence
            Dim yPopIntel As Byte = yData(72)
            If myPopIntel <> yPopIntel Then
                myPopIntel = yPopIntel
                lblPopIntel.Caption = "Intelligence: " & myPopIntel
            End If
            'Research Job Cnt
            Dim lResearchJobCnt As Int32 = System.BitConverter.ToInt32(yData, 73)
            If mlScienceJobs <> lResearchJobCnt Then
                lblPopIntel.ToolTipText = "Indicates the average intelligence of this" & vbCrLf & "colony's population. This value directly" & vbCrLf & _
                  "impacts all research projects and is determined" & vbCrLf & "by the number of scientists in research facilities." & vbCrLf & _
                  "There are currently " & lResearchJobCnt & " scientist jobs."
            End If
        Catch
        End Try
        mbIgnoreEvents = False
    End Sub

    Private Sub frmColonyStats_OnNewFrame() Handles Me.OnNewFrame
        If mlResetColonyTaxRateCycle <> Int32.MinValue AndAlso glCurrentCycle > mlResetColonyTaxRateCycle Then
            myTaxRate = 255
            msw_Delay = Nothing
            mlResetColonyTaxRateCycle = Int32.MinValue
        End If
        If msw_Delay Is Nothing OrElse msw_Delay.ElapsedMilliseconds > ml_REQUEST_DELAY Then
            Me.IsDirty = True
            RefreshDetails()
        End If
    End Sub

    Private Sub txtName_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtName.OnKeyDown
        If e.KeyCode = Keys.Enter Then
            If mlColonyID = -1 Then Exit Sub

            txtName.Enabled = False

            'Submit a name change...
            Dim yMsg(27) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eSetEntityName).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(mlColonyID).CopyTo(yMsg, 2)
            System.BitConverter.GetBytes(ObjectType.eColony).CopyTo(yMsg, 6)
            System.Text.ASCIIEncoding.ASCII.GetBytes(Mid$(txtName.Caption, 1, 20)).CopyTo(yMsg, 8)

            SetCacheObjectValue(mlColonyID, ObjectType.eColony, Mid$(txtName.Caption, 1, 20))
            BPMetrics.MetricMgr.AddActivity(BPMetrics.eInterface.eEntityChangeName, mlColonyID, ObjectType.eColony)
            MyBase.moUILib.SendMsgToPrimary(yMsg)
		End If
		Me.IsDirty = True
    End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

    Private Sub MsgBoxResult(ByVal lResult As MsgBoxResult)
        If lResult = Microsoft.VisualBasic.MsgBoxResult.Yes Then
            Dim sTemp As String = Replace$(txtTaxRate.Caption, "%", "")
            Dim yNewRate As Byte = CByte(Val(sTemp))

            Dim yMsg(8) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.eSetColonyTaxRate).CopyTo(yMsg, 0)
            System.BitConverter.GetBytes(mlColonyID).CopyTo(yMsg, 2)
            yMsg(6) = yNewRate
            yMsg(7) = CByte(CInt(Val(txtControlGrowth.Caption)) + 128I)
            yMsg(8) = CByte(CInt(Val(txtControlMorale.Caption)) + 128I)
            MyBase.moUILib.SendMsgToPrimary(yMsg)

            If NewTutorialManager.TutorialOn = True Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eTaxRateSet, yNewRate, -1, -1, "")
            End If

            msw_Delay = Nothing
        Else
            mlResetColonyTaxRateCycle = glCurrentCycle + 1
            Me.IsDirty = True
            RefreshDetails()
        End If
        
    End Sub

	Private Sub btnSetTax_Click(ByVal sName As String) Handles btnSetTax.Click
		Dim yNewRate As Byte
		Dim yControlledMorale As Byte = 0
		Dim yControlledGrowth As Byte = 0

		If mlColonyID = -1 Then Return

		If HasAliasedRights(AliasingRights.eAlterColonyStats) = False Then
			MyBase.moUILib.AddNotification("You lack the rights to alter tax rates.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If

		Dim sTemp As String = Replace$(txtTaxRate.Caption, "%", "")

		If (IsNumeric(sTemp) = False) OrElse Val(sTemp) < 0 OrElse Val(sTemp) > 100 Then
			MyBase.moUILib.AddNotification("Tax Rate must be a numerical value from 0 to 100!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
        End If

        Dim lTaxRateMax As Int32 = 80
        If mlPopulation > 2000000 Then
            lTaxRateMax = 20
        ElseIf mlPopulation > 1200000 Then
            lTaxRateMax = 22
        ElseIf mlPopulation > 900000 Then
            lTaxRateMax = 24
        ElseIf mlPopulation > 800000 Then
            lTaxRateMax = 26
        ElseIf mlPopulation > 700000 Then
            lTaxRateMax = 28
        ElseIf mlPopulation > 600000 Then
            lTaxRateMax = 30
        ElseIf mlPopulation > 500000 Then
            lTaxRateMax = 32
        ElseIf mlPopulation > 400000 Then
            lTaxRateMax = 34
        ElseIf mlPopulation > 300000 Then
            lTaxRateMax = 36
        ElseIf mlPopulation > 200000 Then
            lTaxRateMax = 38
        ElseIf mlPopulation > 100000 Then
            lTaxRateMax = 40
        End If

        If Val(sTemp) >= lTaxRateMax Then
            Dim ofrm As New frmMsgBox(goUILib, "Setting the tax rate to values greater than " & lTaxRateMax & " could destroy the colony. Are you sure?", MsgBoxStyle.YesNo, "Confirm Tax Rate")
            ofrm.Visible = True
            AddHandler ofrm.DialogClosed, AddressOf MsgBoxResult
            Return
        End If

        If txtControlGrowth.Visible = True Then
            sTemp = Replace$(txtControlGrowth.Caption, "%", "")
            Dim lLimit As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eControlledGrowth)
            If IsNumeric(sTemp) = False OrElse Val(sTemp) < -lLimit OrElse Val(sTemp) > lLimit Then
                MyBase.moUILib.AddNotification("Controlled Growth value must be a numerical value between -" & lLimit & " and " & lLimit & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
        End If

        If txtControlMorale.Visible = True Then
            sTemp = Replace$(txtControlMorale.Caption, "%", "")
            Dim lLimit As Int32 = goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eControlledMorale)
            If IsNumeric(sTemp) = False OrElse Val(sTemp) < -lLimit OrElse Val(sTemp) > lLimit Then
                MyBase.moUILib.AddNotification("Controlled Morale value must be a numerical value between -" & lLimit & " and " & lLimit & "!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
        End If

        sTemp = Replace$(txtTaxRate.Caption, "%", "")
		yNewRate = CByte(Val(sTemp))

		Dim yMsg(8) As Byte
		System.BitConverter.GetBytes(GlobalMessageCode.eSetColonyTaxRate).CopyTo(yMsg, 0)
		System.BitConverter.GetBytes(mlColonyID).CopyTo(yMsg, 2)
		yMsg(6) = yNewRate
        yMsg(7) = CByte(CInt(Val(txtControlGrowth.Caption)) + 128I)
        yMsg(8) = CByte(CInt(Val(txtControlMorale.Caption)) + 128I)
		MyBase.moUILib.SendMsgToPrimary(yMsg)

		If NewTutorialManager.TutorialOn = True Then
			NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eTaxRateSet, yNewRate, -1, -1, "")
		End If

		msw_Delay = Nothing
	End Sub

    Protected Overrides Sub Finalize()
        If msw_Delay Is Nothing = False Then msw_Delay.Stop()
        msw_Delay = Nothing
        MyBase.Finalize()
    End Sub

	Private Sub lblExpand_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles lblExpand.OnMouseDown

		If NewTutorialManager.TutorialOn = True Then
			If MyBase.moUILib.CommandAllowed(True, "ToggleColonyStatsView") = False Then Return
		End If

		Dim frmTemp As frmColonyStatsSmall = CType(MyBase.moUILib.GetWindow("frmColonyStatsSmall"), frmColonyStatsSmall)
		If frmTemp Is Nothing Then
			frmTemp = New frmColonyStatsSmall(MyBase.moUILib)
		Else : frmTemp.Visible = True
		End If
		MyBase.moUILib.RemoveWindow(Me.ControlName)
		muSettings.ExpandedColonyStatsScreen = False
        frmTemp = Nothing

    End Sub

    Private Sub lblMorale_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles lblMorale.OnMouseMove
        If msMoraleRollOver <> "" Then MyBase.moUILib.SetToolTip(msMoraleRollOver, Me.Left - 150, lMouseY)
    End Sub

    Private Sub lblHousing_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles lblHousing.OnMouseDown
        frmMain.SelectAndGotoNextUnit(frmMain.eSelectNextType.eUnpoweredResidence)
    End Sub

    Private Sub lblHousing_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles lblHousing.OnMouseMove
        If msHousingRollOver <> "" Then MyBase.moUILib.SetToolTip(msHousingRollOver, Me.Left - 150, lMouseY)
	End Sub

	Private Sub txtTaxRate_OnGotFocus() Handles txtTaxRate.OnGotFocus
		txtTaxRate.Caption = txtTaxRate.Caption.Replace("%", "")
	End Sub

	Private Sub txtTaxRate_OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs) Handles txtTaxRate.OnKeyDown
		If e.KeyCode = Keys.Enter Then
			If NewTutorialManager.TutorialOn = True Then
				If MyBase.moUILib.CommandAllowed(True, "frmColonyStats.btnSetTax") = False Then Return
			End If
			btnSetTax_Click("btnSetTax")
		End If
	End Sub

    Private Sub txtTaxRate_OnLostFocus() Handles txtTaxRate.OnLostFocus
        mlResetColonyTaxRateCycle = glCurrentCycle + 450
    End Sub

    Private Sub frmColonyStats_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.ColonyStatsX = Me.Left
            muSettings.ColonyStatsY = Me.Top
        End If
    End Sub
End Class