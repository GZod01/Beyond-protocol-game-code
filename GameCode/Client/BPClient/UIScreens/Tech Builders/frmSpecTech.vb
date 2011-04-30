
Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'#Region "  Interface created from Interface Builder  "
'Public Class frmSpecTech
'    Inherits UIWindow

'    Private lblTitle As UILabel
'    Private lblName As UILabel
'    Private txtRPDesc As UITextBox
'    Private lblRPDesc As UILabel
'    Private lblBenefits As UILabel
'    Private txtBenefits As UITextBox
'    Private lnDiv1 As UILine
'    Private WithEvents btnClose As UIButton

'    Private mfrmResCost As frmProdCost

'    Public Sub New(ByRef oUILib As UILib)
'        MyBase.New(oUILib)

'        'frmSpecTech initial props
'        With Me

'            Dim lTempX As Int32 = -1
'            Dim lTempY As Int32 = -1
'            Dim oTmpWin As UIWindow = oUILib.GetWindow("frmResearchMain")
'            If oTmpWin Is Nothing = False Then
'                lTempX = oTmpWin.Left + oTmpWin.Width
'                If lTempX + 350 > oUILib.oDevice.PresentationParameters.BackBufferWidth Then lTempX = -1
'                lTempY = oTmpWin.Top
'            End If
'            If lTempX = -1 OrElse lTempY = -1 Then
'                lTempX = (oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - 275
'                lTempY = (oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 132
'            End If

'            .ControlName = "frmSpecTech"
'            .Left = lTempX
'            .Top = lTempY
'            .Width = 555 '350
'            .Height = 265
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceFillColor
'            .FullScreen = False
'            .BorderLineWidth = 1
'        End With

'        'lblTitle initial props
'        lblTitle = New UILabel(oUILib)
'        With lblTitle
'            .ControlName = "lblTitle"
'            .Left = 5
'            .Top = 5
'            .Width = 210
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Special Technology Research"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblTitle, UIControl))

'        'lblName initial props
'        lblName = New UILabel(oUILib)
'        With lblName
'            .ControlName = "lblName"
'            .Left = 5
'            .Top = 30
'            .Width = 698
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Name: Dynamic Field Induction Principles"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblName, UIControl))

'        'txtRPDesc initial props
'        txtRPDesc = New UITextBox(oUILib)
'        With txtRPDesc
'            .ControlName = "txtRPDesc"
'            .Left = 5
'            .Top = 70
'            .Width = 340
'            .Height = 80
'            .Enabled = True
'            .Visible = True
'            .Caption = ""
'            .ForeColor = muSettings.InterfaceTextBoxForeColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.Left Or DrawTextFormat.Top
'            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'            .MaxLength = 0
'            .BorderColor = muSettings.InterfaceBorderColor
'            .Locked = True
'            .MultiLine = True
'        End With
'        Me.AddChild(CType(txtRPDesc, UIControl))

'        'lblRPDesc initial props
'        lblRPDesc = New UILabel(oUILib)
'        With lblRPDesc
'            .ControlName = "lblRPDesc"
'            .Left = 5
'            .Top = 50
'            .Width = 145
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "From the Head Scientist"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblRPDesc, UIControl))

'        'lblBenefits initial props
'        lblBenefits = New UILabel(oUILib)
'        With lblBenefits
'            .ControlName = "lblBenefits"
'            .Left = 5
'            .Top = 160
'            .Width = 108
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Expected Benefits"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblBenefits, UIControl))

'        'txtBenefits initial props
'        txtBenefits = New UITextBox(oUILib)
'        With txtBenefits
'            .ControlName = "txtBenefits"
'            .Left = 5
'            .Top = 180
'            .Width = 340
'            .Height = 80
'            .Enabled = True
'            .Visible = True
'            .Caption = ""
'            .ForeColor = muSettings.InterfaceTextBoxForeColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.Left Or DrawTextFormat.Top
'            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
'            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
'            .MaxLength = 0
'            .BorderColor = muSettings.InterfaceBorderColor
'            .Locked = True
'            .MultiLine = True
'        End With
'        Me.AddChild(CType(txtBenefits, UIControl))

'        'lnDiv1 initial props
'        lnDiv1 = New UILine(oUILib)
'        With lnDiv1
'            .ControlName = "lnDiv1"
'            .Left = 0
'            .Top = 27
'            .Width = 555
'            .Height = 0
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(lnDiv1, UIControl))

'        'btnClose initial props
'        btnClose = New UIButton(oUILib)
'        With btnClose
'            .ControlName = "btnClose"
'            .Left = Me.Width - 25
'            .Top = 1
'            .Width = 24
'            .Height = 24
'            .Enabled = True
'            .Visible = True
'            .Caption = "X"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'        End With
'        Me.AddChild(CType(btnClose, UIControl))

'        'mfrmResCost Initial Props
'        mfrmResCost = New frmProdCost(goUILib, "Research Cost", True)
'        With mfrmResCost
'            '.Visible = False
'            .Left = 350
'            .Top = 70
'            .Height = 190
'        End With
'        Me.AddChild(CType(mfrmResCost, UIControl))

'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'        MyBase.moUILib.AddWindow(Me)
'    End Sub

'    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'    End Sub

'    Private Sub frmSpecTech_OnNewFrame() Handles Me.OnNewFrame
'        Dim oTmpWin As UIWindow = MyBase.moUILib.GetWindow("frmResearchMain")
'        If oTmpWin Is Nothing OrElse oTmpWin.Visible = False Then MyBase.moUILib.RemoveWindow(Me.ControlName)
'        oTmpWin = Nothing
'    End Sub

'    Public Sub SetFromSpecialTech(ByRef oSpecTech As SpecialTech)
'        If oSpecTech Is Nothing = False Then
'            With oSpecTech
'                lblName.Caption = "Name: " & .sName
'                txtRPDesc.Caption = .sDesc
'                txtBenefits.Caption = .sBenefits

'                mfrmResCost.SetFromProdCost(.oResearchCost, 100, True, 0, 0)
'            End With
'        Else
'            MyBase.moUILib.RemoveWindow(Me.ControlName)
'        End If
'    End Sub
'End Class
'#End Region

Public Class frmSpecTech
	Inherits UIWindow

    Private Enum eyThrowBackType As Byte
        ePushToBypass = 1
        eSwapFromBypass = 2
    End Enum

    'Private WithEvents lstTechs As UIListBox
    'Private WithEvents chkShowUnresearched As UICheckBox
    'Private WithEvents chkShowResearched As UICheckBox
    Private WithEvents tvwTechs As UITreeView
	Private WithEvents btnResearch As UIButton
    Private WithEvents btnClose As UIButton
    Private WithEvents btnHelp As UIButton
    Private WithEvents btnExport As UIButton
    Private WithEvents btnThrowBack As UIButton

	Private lblTitle As UILabel
	Private lnDiv1 As UILine
	Private txtRPDesc As UITextBox
	Private lblTech As UILabel
	Private lblRPDesc As UILabel
	Private lblBenefits As UILabel
	Private txtBenefits As UITextBox

	Private lblCosts As UILabel
    Public mfrmResCost As frmProdCost

	Private mlLastResCnt As Int32 = 0
    Private mlLastUnresCnt As Int32 = 0
    Private mlLastTankCnt As Int32 = 0
	Private mlLastCheckUpdate As Int32 = 0

    Private moSwapBox As frmSwapTech

	Private Shared myLastFilterSettings As Byte = 3

    Private mbLoading As Boolean = True

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

        If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase = 1 Then
            Dim oFrm As New frmSpecTechSelect(goUILib)
            oFrm.Visible = True
            Return
        End If

        'frmSpecTech initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eSpecialTech
            .ControlName = "frmSpecTech"
            .Width = 600
            .Height = 511 '455
            If muSettings.SpecialTechLeft < 0 OrElse muSettings.SpecialTechTop < 0 OrElse muSettings.SpecialTechLeft > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width OrElse muSettings.SpecialTechTop > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height OrElse NewTutorialManager.TutorialOn = True Then
                .Left = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2) - 300
                .Top = (MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2) - 220
                muSettings.SpecialTechLeft = .Left
                muSettings.SpecialTechTop = .Top
            Else
                .Left = muSettings.SpecialTechLeft
                .Top = muSettings.SpecialTechTop
            End If
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = True
            .BorderLineWidth = 1
            .mbAcceptReprocessEvents = True
        End With

        ''lstTechs initial props
        'lstTechs = New UIListBox(oUILib)
        'With lstTechs
        '	.ControlName = "lstTechs"
        '	.Left = 5
        '	.Top = 55
        '	.Width = 335
        '	.Height = 360
        '	.Enabled = True
        '	.Visible = True
        '	.BorderColor = muSettings.InterfaceBorderColor
        '	.FillColor = muSettings.InterfaceFillColor
        '	.ForeColor = muSettings.InterfaceBorderColor
        '	.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '	.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        '	.mbAcceptReprocessEvents = True
        'End With
        'Me.AddChild(CType(lstTechs, UIControl))
        'tvwTechs initial props
        tvwTechs = New UITreeView(oUILib)
        With tvwTechs
            .ControlName = "tvwTechs"
            .Left = 5
            .Top = 55
            .Width = 335
            .Height = 416 '360
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(tvwTechs, UIControl))

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 300
            .Height = 20
            .Enabled = True
            .Visible = True
            .Caption = "Special Project Technology Research"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 0
            .Top = 32
            .Width = 600
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'txtRPDesc initial props
        txtRPDesc = New UITextBox(oUILib)
        With txtRPDesc
            .ControlName = "txtRPDesc"
            .Left = 345
            .Top = 55
            .Width = 250
            .Height = 100
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Top Or DrawTextFormat.Left
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .MultiLine = True
            .Locked = True
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(txtRPDesc, UIControl))

        'lblTech initial props
        lblTech = New UILabel(oUILib)
        With lblTech
            .ControlName = "lblTech"
            .Left = 5
            .Top = 35
            .Width = 198
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Currently Proposed Projects"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblTech, UIControl))

        'lblRPDesc initial props
        lblRPDesc = New UILabel(oUILib)
        With lblRPDesc
            .ControlName = "lblRPDesc"
            .Left = 345
            .Top = 35
            .Width = 133
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Proposal Summary"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblRPDesc, UIControl))

        'lblBenefits initial props
        lblBenefits = New UILabel(oUILib)
        With lblBenefits
            .ControlName = "lblBenefits"
            .Left = 345
            .Top = 155 '165
            .Width = 198
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Expected Project Benefits"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblBenefits, UIControl))

        'txtBenefits initial props
        txtBenefits = New UITextBox(oUILib)
        With txtBenefits
            .ControlName = "txtBenefits"
            .Left = 345
            .Top = 175 '185
            .Width = 250
            .Height = 100
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .MultiLine = True
            .Locked = True
            .mbAcceptReprocessEvents = True
        End With
        Me.AddChild(CType(txtBenefits, UIControl))

        ''chkShowUnresearched initial props
        'chkShowUnresearched = New UICheckBox(oUILib)
        'With chkShowUnresearched
        '	.ControlName = "chkShowUnresearched"
        '	.Left = 20
        '	.Top = 425
        '	.Width = 138
        '	.Height = 18
        '	.Enabled = True
        '	.Visible = True
        '	.Caption = "Show Unresearched"
        '	.ForeColor = muSettings.InterfaceBorderColor
        '	.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '	.DrawBackImage = False
        '	.FontFormat = CType(6, DrawTextFormat)
        '	.Value = True
        '	.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        'End With
        'Me.AddChild(CType(chkShowUnresearched, UIControl))

        ''chkShowResearched initial props
        'chkShowResearched = New UICheckBox(oUILib)
        'With chkShowResearched
        '	.ControlName = "chkShowResearched"
        '	.Left = 185
        '	.Top = 425
        '	.Width = 127
        '	.Height = 18
        '	.Enabled = True
        '	.Visible = True
        '	.Caption = "Show Researched"
        '	.ForeColor = muSettings.InterfaceBorderColor
        '	.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '	.DrawBackImage = False
        '	.FontFormat = CType(6, DrawTextFormat)
        '	.Value = True
        '	.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
        'End With
        'Me.AddChild(CType(chkShowResearched, UIControl))

        'btnResearch initial props
        btnResearch = New UIButton(oUILib)
        With btnResearch
            .ControlName = "btnResearch"
            .Height = 24
            .Width = 100
            .Left = Me.Width - .Width - 5 '380
            .Top = 478 ' 422
            .Enabled = True
            .Visible = True
            .Caption = "Research"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnResearch, UIControl))


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
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
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
        End With
        Me.AddChild(CType(btnHelp, UIControl))

        'btnExport
        btnExport = New UIButton(oUILib)
        With btnExport
            .ControlName = "btnExport"
            .Left = btnHelp.Left - 25
            .Top = btnClose.Top
            .Width = 24
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "E"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            '.ToolTipText = "You can change the format of exported data from CSV to XML " & vbCrLf & "in the general settings window." & vbCrLf & "Exported data will go into your GameDir\ExportedData"
            .ToolTipText = "Click to export data." & vbCrLf & "Exported data will go into your GameDir\ExportedData"
        End With
        Me.AddChild(CType(btnExport, UIControl))

        'btnThrowBack initial props
        btnThrowBack = New UIButton(oUILib)
        With btnThrowBack
            .ControlName = "btnThrowBack"
            .Left = tvwTechs.Left + (tvwTechs.Width \ 2) - 120
            .Top = 478
            .Width = 240
            .Height = 24
            .Enabled = True
            .Visible = False
            .Caption = "What Else Is Available?"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Puts the currently selected tech into the Bypass grouping." & vbCrLf & _
                           "Bypassed techs do not count towards the maximum techs available" & vbCrLf & _
                           "and cannot be researched while bypassed. A tech in Unresearched" & vbCrLf & _
                           "group can be swapped with a tech that has been bypassed for a" & vbCrLf & _
                           "fee of 10% of the research cost of the bypassed tech. When a tech" & vbCrLf & _
                           "is bypassed, your scientists will attempt to come up with a new" & vbCrLf & _
                           "idea for a special project as soon as possible."
        End With
        Me.AddChild(CType(btnThrowBack, UIControl))

        'lblCosts initial props
        lblCosts = New UILabel(oUILib)
        With lblCosts
            .ControlName = "lblCosts"
            .Left = 345
            .Top = 275 '295
            .Width = 198
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Proposal Costs"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblCosts, UIControl))

        'mfrmResCost Initial Props
        mfrmResCost = New frmProdCost(goUILib, "Research Cost", True)
        With mfrmResCost
            '.Visible = False
            .Left = 345
            .Top = 292 ' 315
            .Height = 180 ' 120 '100
            .Width = 250
            .Moveable = False
        End With
        Me.AddChild(CType(mfrmResCost, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        'chkShowResearched.Value = (myLastFilterSettings And 1) <> 0
        'chkShowUnresearched.Value = (myLastFilterSettings And 2) <> 0

        RefreshLists()
        mbLoading = False
    End Sub

    'Private Sub RefreshLists()
    '	lstTechs.Clear()
    '	mlLastResCnt = 0
    '	mlLastUnresCnt = 0

    '	Try

    '		Dim bShowRes As Boolean = chkShowResearched.Value
    '		Dim bShowUnRes As Boolean = chkShowUnresearched.Value

    '		Dim lCreditCostID() As Int32 = Nothing
    '		Dim iCreditCostTypeID() As Int16 = Nothing
    '		Dim blCreditCosts() As Int64 = Nothing
    '		Dim lCreditCostUB As Int32 = -1

    '		For X As Int32 = 0 To goCurrentPlayer.mlTechUB
    '			If goCurrentPlayer.moTechs(X) Is Nothing = False Then
    '				Dim lPhase As Int32 = goCurrentPlayer.moTechs(X).ComponentDevelopmentPhase
    '				If lPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched AndAlso lPhase <> Base_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
    '					lCreditCostUB += 1
    '					ReDim Preserve blCreditCosts(lCreditCostUB)
    '					ReDim Preserve lCreditCostID(lCreditCostUB)
    '					ReDim Preserve iCreditCostTypeID(lCreditCostUB)
    '					If goCurrentPlayer.moTechs(X).oResearchCost Is Nothing = False Then blCreditCosts(lCreditCostUB) = goCurrentPlayer.moTechs(X).oResearchCost.CreditCost Else blCreditCosts(lCreditCostUB) = 0
    '					lCreditCostID(lCreditCostUB) = goCurrentPlayer.moTechs(X).ObjectID
    '					iCreditCostTypeID(lCreditCostUB) = goCurrentPlayer.moTechs(X).ObjTypeID
    '				End If
    '			End If
    '		Next X

    '		For X As Int32 = 0 To goCurrentPlayer.mlTechUB
    '			If goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eSpecialTech Then
    '				With CType(goCurrentPlayer.moTechs(X), SpecialTech)

    '					Dim bAdd As Boolean = False

    '					If .ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
    '						mlLastResCnt += 1
    '						If bShowRes = True Then bAdd = True
    '					End If
    '					If .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
    '						mlLastUnresCnt += 1
    '						If bShowUnRes = True Then bAdd = True
    '					End If

    '					If bAdd = True Then
    '						lstTechs.AddItem(.sName)
    '						lstTechs.ItemData(lstTechs.NewIndex) = .ObjectID
    '						lstTechs.ItemData2(lstTechs.NewIndex) = .ObjTypeID

    '						'Now, determine how I fair against other techs
    '						Dim lIdx As Int32 = -1
    '						For Y As Int32 = 0 To lCreditCostUB
    '							If lCreditCostID(Y) = goCurrentPlayer.moTechs(X).ObjectID AndAlso iCreditCostTypeID(Y) = goCurrentPlayer.moTechs(X).ObjTypeID Then
    '								lIdx = Y
    '								Exit For
    '							End If
    '						Next Y
    '						If lIdx <> -1 Then
    '							Dim lBelowMe As Int32 = 0
    '							Dim lAboveMe As Int32 = 0
    '							For Y As Int32 = 0 To lCreditCostUB
    '								If lIdx <> Y Then
    '									If blCreditCosts(Y) <= blCreditCosts(lIdx) Then lBelowMe += 1 Else lAboveMe += 1
    '								End If
    '							Next Y

    '							Dim lTotal As Int32 = lBelowMe + lAboveMe
    '							If lTotal <> 0 Then
    '								Dim fPerc As Single = CSng(lBelowMe / lTotal)
    '								Dim lG As Int32 = 255 - CInt(fPerc * 250)
    '								Dim lR As Int32 = 255 - lG

    '								lstTechs.ItemCustomColor(lstTechs.NewIndex) = System.Drawing.Color.FromArgb(255, lR, lG, 0)
    '							Else : lstTechs.ItemCustomColor(lstTechs.NewIndex) = System.Drawing.Color.FromArgb(255, 0, 255, 0)
    '							End If
    '						End If

    '					End If
    '				End With
    '			End If
    '		Next X

    '		If lstTechs.ListIndex > -1 AndAlso lstTechs.ListIndex < lstTechs.ListCount Then
    '			lstTechs_ItemClick(lstTechs.ListIndex)
    '		End If
    '	Catch
    '		mlLastResCnt = -1
    '		mlLastUnresCnt = -1
    '	End Try
    'End Sub
    Private Sub RefreshLists()

        Dim oTmpNode As UITreeView.UITreeViewItem = tvwTechs.oRootNode
        Dim bUnresExpanded As Boolean = False
        Dim bResExpanded As Boolean = False
        Dim bTankExpanded As Boolean = False
        If oTmpNode Is Nothing = False Then
            bUnresExpanded = oTmpNode.bExpanded
            oTmpNode = oTmpNode.oNextSibling
            If oTmpNode Is Nothing = False Then
                bResExpanded = oTmpNode.bExpanded
                oTmpNode = oTmpNode.oNextSibling
                If oTmpNode Is Nothing = False Then bTankExpanded = oTmpNode.bExpanded
            End If
        End If


        tvwTechs.Clear()
        tvwTechs.oSelectedNode = Nothing
        mlLastResCnt = 0
        mlLastUnresCnt = 0
        mlLastTankCnt = 0

        Try

            'Dim bShowRes As Boolean = chkShowResearched.Value
            'Dim bShowUnRes As Boolean = chkShowUnresearched.Value

            Dim oUnresNode As UITreeView.UITreeViewItem = tvwTechs.AddNode("Unresearched", -1, -1, -1, Nothing, Nothing)
            Dim oResNode As UITreeView.UITreeViewItem = tvwTechs.AddNode("Researched", -1, -1, -1, Nothing, Nothing)
            Dim oTankNode As UITreeView.UITreeViewItem = tvwTechs.AddNode("Bypassed", -1, -1, -1, Nothing, Nothing)

            oUnresNode.bItemBold = True : oUnresNode.clrFillColor = System.Drawing.Color.FromArgb(255, 128, 128, 128) : oUnresNode.bUseFillColor = True
            oResNode.bItemBold = True : oResNode.clrFillColor = System.Drawing.Color.FromArgb(255, 128, 128, 128) : oResNode.bUseFillColor = True
            oTankNode.bItemBold = True : oTankNode.clrFillColor = System.Drawing.Color.FromArgb(255, 128, 128, 128) : oTankNode.bUseFillColor = True


            Dim oPNodes(elCategory.eLastCategory - 1) As UITreeView.UITreeViewItem
            Dim lPNodeCnt(oPNodes.GetUpperBound(0)) As Int32
            For X As Int32 = 0 To elCategory.eLastCategory - 1
                oPNodes(X) = tvwTechs.AddNode(GetCategoryName(CType(X, elCategory)), -1, -1, -1, oResNode, Nothing)
                oPNodes(X).bItemBold = True : oPNodes(X).clrFillColor = System.Drawing.Color.FromArgb(255, 128, 128, 128) : oPNodes(X).bUseFillColor = True
                lPNodeCnt(X) = 0
            Next X

            Dim lCreditCostID() As Int32 = Nothing
            Dim iCreditCostTypeID() As Int16 = Nothing
            Dim blCreditCosts() As Int64 = Nothing
            Dim lCreditCostUB As Int32 = -1

            For X As Int32 = 0 To goCurrentPlayer.mlTechUB
                If goCurrentPlayer.moTechs(X) Is Nothing = False Then
                    Dim lPhase As Int32 = goCurrentPlayer.moTechs(X).ComponentDevelopmentPhase
                    If lPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched AndAlso lPhase <> Base_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
                        lCreditCostUB += 1
                        ReDim Preserve blCreditCosts(lCreditCostUB)
                        ReDim Preserve lCreditCostID(lCreditCostUB)
                        ReDim Preserve iCreditCostTypeID(lCreditCostUB)
                        If goCurrentPlayer.moTechs(X).oResearchCost Is Nothing = False Then blCreditCosts(lCreditCostUB) = goCurrentPlayer.moTechs(X).oResearchCost.CreditCost Else blCreditCosts(lCreditCostUB) = 0
                        lCreditCostID(lCreditCostUB) = goCurrentPlayer.moTechs(X).ObjectID
                        iCreditCostTypeID(lCreditCostUB) = goCurrentPlayer.moTechs(X).ObjTypeID
                    End If
                End If
            Next X

            For X As Int32 = 0 To goCurrentPlayer.mlTechUB
                If goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eSpecialTech Then
                    With CType(goCurrentPlayer.moTechs(X), SpecialTech)

                        'Dim bAdd As Boolean = True False

                        Dim oNode As UITreeView.UITreeViewItem = Nothing
                        If .ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
                            mlLastResCnt += 1

                            'oNode = tvwTechs.AddNode(.sName, .ObjectID, .ObjTypeID, 0, oResNode, Nothing)
                            Dim lPNode As Int32 = GetBaseTypeFromProgramControl(.lProgramControl)
                            lPNodeCnt(lPNode) += 1
                            oNode = tvwTechs.AddNode(.sName, .ObjectID, .ObjTypeID, 0, oPNodes(lPNode), Nothing)
                        End If
                        If .ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then

                            If .bInTheTank = True Then
                                mlLastTankCnt += 1
                                'If bShowUnRes = True Then bAdd = True
                                oNode = tvwTechs.AddNode(.sName, .ObjectID, .ObjTypeID, 0, oTankNode, Nothing)
                            Else
                                mlLastUnresCnt += 1
                                'If bShowUnRes = True Then bAdd = True
                                oNode = tvwTechs.AddNode(.sName, .ObjectID, .ObjTypeID, 0, oUnresNode, Nothing)
                            End If

                        End If
                        If oNode Is Nothing Then Continue For

                        'oNode.sItem &= " (" & .lProgramControl & ")"


                        'If bAdd = True Then
                        '    lstTechs.AddItem(.sName)
                        '    lstTechs.ItemData(lstTechs.NewIndex) = .ObjectID
                        '    lstTechs.ItemData2(lstTechs.NewIndex) = .ObjTypeID

                        'Now, determine how I fair against other techs
                        Dim lIdx As Int32 = -1
                        For Y As Int32 = 0 To lCreditCostUB
                            If lCreditCostID(Y) = goCurrentPlayer.moTechs(X).ObjectID AndAlso iCreditCostTypeID(Y) = goCurrentPlayer.moTechs(X).ObjTypeID Then
                                lIdx = Y
                                Exit For
                            End If
                        Next Y
                        Dim clrVal As System.Drawing.Color
                        If lIdx <> -1 Then
                            Dim lBelowMe As Int32 = 0
                            Dim lAboveMe As Int32 = 0
                            For Y As Int32 = 0 To lCreditCostUB
                                If lIdx <> Y Then
                                    If blCreditCosts(Y) <= blCreditCosts(lIdx) Then lBelowMe += 1 Else lAboveMe += 1
                                End If
                            Next Y

                            Dim lTotal As Int32 = lBelowMe + lAboveMe
                            If lTotal <> 0 Then
                                Dim fPerc As Single = CSng(lBelowMe / lTotal)
                                Dim lG As Int32 = 255 - CInt(fPerc * 250)
                                Dim lR As Int32 = 255 - lG

                                'lstTechs.ItemCustomColor(lstTechs.NewIndex) = System.Drawing.Color.FromArgb(255, lR, lG, 0)
                                clrVal = System.Drawing.Color.FromArgb(255, lR, lG, 0)
                            Else
                                'lstTechs.ItemCustomColor(lstTechs.NewIndex) = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                                clrVal = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                            End If

                            oNode.clrItemColor = clrVal
                            oNode.bUseItemColor = True
                        End If

                        'End If
                    End With
                End If
            Next X

            For X As Int32 = 0 To oPNodes.GetUpperBound(0)
                If oPNodes(X).oFirstChild Is Nothing Then
                    tvwTechs.RemoveNode(oPNodes(X))
                Else
                    oPNodes(X).sItem &= " (" & lPNodeCnt(X) & ")"
                End If
            Next X

            oUnresNode.sItem = "Unresearched (" & mlLastUnresCnt & ")"
            oResNode.sItem = "Researched (" & mlLastResCnt & ")"
            oTankNode.sItem = "Bypassed (" & mlLastTankCnt & ")"

            oUnresNode.bExpanded = bUnresExpanded
            oResNode.bExpanded = bResExpanded
            oTankNode.bExpanded = bTankExpanded

            'If lstTechs.ListIndex > -1 AndAlso lstTechs.ListIndex < lstTechs.ListCount Then
            '    lstTechs_ItemClick(lstTechs.ListIndex)
            'End If
        Catch
            mlLastResCnt = -1
            mlLastUnresCnt = -1
            mlLastTankCnt = -1
        End Try
    End Sub

	Private Sub frmSpecTech_OnNewFrame() Handles Me.OnNewFrame
		Dim oTmpWin As UIWindow = MyBase.moUILib.GetWindow("frmResearchMain")
		If oTmpWin Is Nothing OrElse oTmpWin.Visible = False Then MyBase.moUILib.RemoveWindow(Me.ControlName)
		oTmpWin = Nothing

		Dim lResCnt As Int32 = 0
        Dim lUnResCnt As Int32 = 0
        Dim lTankCnt As Int32 = 0
		If glCurrentCycle - mlLastCheckUpdate > 30 Then
			mlLastCheckUpdate = glCurrentCycle

            If goCurrentPlayer Is Nothing = False Then
                For X As Int32 = 0 To goCurrentPlayer.mlTechUB
                    Dim oTech As Base_Tech = goCurrentPlayer.moTechs(X)
                    If oTech Is Nothing = False AndAlso oTech.ObjTypeID = ObjectType.eSpecialTech Then
                        With oTech
                            If .ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
                                lResCnt += 1
                            ElseIf CType(oTech, SpecialTech).bInTheTank = True Then
                                lTankCnt += 1
                            Else
                                lUnResCnt += 1
                            End If
                        End With
                    End If
                Next X
            End If

            'If (mlLastResCnt <> lResCnt AndAlso chkShowResearched.Value = True) OrElse (mlLastUnresCnt <> lUnResCnt AndAlso chkShowUnresearched.Value = True) Then RefreshLists()
            If (mlLastResCnt <> lResCnt) OrElse (mlLastUnresCnt <> lUnResCnt) OrElse (mlLastTankCnt <> lTankCnt) Then RefreshLists()
        End If
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

    Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
        If goTutorial Is Nothing Then goTutorial = New TutorialManager()
        goTutorial.ExecuteTutorialStep(120)
    End Sub

	Private Sub btnResearch_Click(ByVal sName As String) Handles btnResearch.Click
		If HasAliasedRights(AliasingRights.eAddResearch) = False Then
			MyBase.moUILib.AddNotification("You lack rights to add research.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			Return
        End If

        If tvwTechs.oSelectedNode Is Nothing = False Then
            Dim oNode As UITreeView.UITreeViewItem = tvwTechs.oSelectedNode
            If oNode Is Nothing Then Return

            Dim lID As Int32 = oNode.lItemData 'lstTechs.ItemData(lstTechs.ListIndex)
            Dim iTypeID As Int16 = ObjectType.eSpecialTech

            Dim oTmpTech As Base_Tech = goCurrentPlayer.GetTech(lID, iTypeID)
            If oTmpTech Is Nothing OrElse oTmpTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eInvalidDesign Then
                goUILib.AddNotification("Unable to Research selected component!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
            If CType(oTmpTech, SpecialTech).bInTheTank = True Then
                goUILib.AddNotification("Unable to research selected item as it has been Bypassed.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            If oTmpTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
                goUILib.AddNotification("Cannot research that technology as it is already researched!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
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

	Private Sub HideMeIfNeeded()
		'MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

    'Private Sub chkFilter_Click() Handles chkShowResearched.Click, chkShowUnresearched.Click
    '	Dim yValue As Byte = 0
    '	If chkShowResearched.Value = True Then yValue = CByte(yValue Or 1)
    '	If chkShowUnresearched.Value = True Then yValue = CByte(yValue Or 2)
    '	myLastFilterSettings = yValue
    '	RefreshLists()
    'End Sub

    'Private Sub lstTechs_ItemClick(ByVal lIndex As Integer) Handles lstTechs.ItemClick
    '	If lIndex > -1 Then
    '		Dim lID As Int32 = lstTechs.ItemData(lIndex)

    '		Dim oSpecTech As SpecialTech = CType(goCurrentPlayer.GetTech(lID, ObjectType.eSpecialTech), SpecialTech)
    '		btnResearch.Enabled = False
    '		If oSpecTech Is Nothing = False Then

    '			If NewTutorialManager.TutorialOn = True Then
    '				NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eSpecialProjectSelected, lID, ObjectType.eSpecialTech, -1, "")
    '			End If

    '			With oSpecTech
    '				txtRPDesc.Caption = .sDesc
    '				txtBenefits.Caption = .sBenefits

    '				If .ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
    '					lblCosts.Visible = False
    '					mfrmResCost.Visible = False
    '					btnResearch.Enabled = False
    '				Else
    '					btnResearch.Enabled = True
    '					lblCosts.Visible = True
    '					mfrmResCost.Visible = True
    '					mfrmResCost.SetFromProdCost(.oResearchCost, 100, True, 0, 0)
    '				End If
    '			End With
    '		End If
    '	End If

    'End Sub

    'Private Sub lstTechs_ItemDblClick(ByVal lIndex As Integer) Handles lstTechs.ItemDblClick
    '	btnResearch_Click(btnResearch.ControlName)
    'End Sub

	Public Sub SetFromSpecialTech(ByRef oSpecTech As SpecialTech)
        Try
            If oSpecTech Is Nothing = False Then
                tvwTechs.oSelectedNode = tvwTechs.GetNodeByItemData2(oSpecTech.ObjectID, oSpecTech.ObjTypeID)

                'For X As Int32 = 0 To lstTechs.ListCount - 1
                '    If lstTechs.ItemData(X) = oSpecTech.ObjectID AndAlso lstTechs.ItemData2(X) = ObjectType.eSpecialTech Then
                '        lstTechs.ListIndex = X
                '        lstTechs_ItemClick(X)
                '        Exit For
                '    End If
                'Next X
            Else
                MyBase.moUILib.RemoveWindow(Me.ControlName)
            End If
        Catch
        End Try
    End Sub

    Private Sub tvwTechs_NodeDoubleClicked() Handles tvwTechs.NodeDoubleClicked
        'btnResearch_Click(btnResearch.ControlName)
    End Sub

    Private Sub tvwTechs_NodeSelected(ByRef oNode As UITreeView.UITreeViewItem) Handles tvwTechs.NodeSelected
        If oNode Is Nothing = False Then
            Dim lID As Int32 = oNode.lItemData '.ItemData(lIndex)

            Dim oSpecTech As SpecialTech = CType(goCurrentPlayer.GetTech(lID, ObjectType.eSpecialTech), SpecialTech)
            btnResearch.Enabled = False
            If oSpecTech Is Nothing = False Then

                If NewTutorialManager.TutorialOn = True Then
                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eSpecialProjectSelected, lID, ObjectType.eSpecialTech, -1, "")
                End If

                With oSpecTech
                    txtRPDesc.Caption = .sDesc
                    txtBenefits.Caption = .sBenefits

                    If .ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
                        lblCosts.Visible = False
                        mfrmResCost.Visible = False
                        btnResearch.Enabled = False

                        btnThrowBack.Visible = False
                    Else
                        btnResearch.Enabled = True
                        lblCosts.Visible = True
                        mfrmResCost.Visible = True
                        mfrmResCost.SetFromProdCost(.oResearchCost, 100, True, 0, 0)


                        btnThrowBack.Visible = True
                        If .bInTheTank = True Then
                            btnThrowBack.Caption = "Bring Back Design"
                            btnThrowBack.ToolTipText = "Brings the selected design back out of the Bypass group." & vbCrLf & _
                                                       "Doing so will cost a fee of 10% of the research cost of" & vbCrLf & _
                                                       "the selected technology. You will be prompted for an idea" & vbCrLf & _
                                                       "currently available in the list to put in the selected" & vbCrLf & _
                                                       "technology's place if you already have three projects to research." & vbCrLf & _
                                                       "Just because the research queue is empty does not mean" & vbCrLf & _
                                                       "that your scientists are out of ideas. They may have some" & vbCrLf & _
                                                       "ideas in the works that they are preparing to propose to you."
                        Else
                            btnThrowBack.Caption = "What Else Is Available?"
                            btnThrowBack.ToolTipText = "Puts the currently selected tech into the Bypass grouping." & vbCrLf & _
                                                       "Bypassed techs do not count towards the maximum techs available" & vbCrLf & _
                                                       "and cannot be researched while bypassed. A tech in Unresearched" & vbCrLf & _
                                                       "group can be swapped with a tech that has been bypassed for a fee" & vbCrLf & _
                                                       "of 10% of the research cost of the bypassed tech. When a tech is" & vbCrLf & _
                                                       "bypassed, your scientists will attempt to come up with a new idea" & vbCrLf & _
                                                       "for a special project as soon as possible, which can take a while."
                        End If
                    End If

                End With
            Else : btnThrowBack.Visible = False
            End If
        End If
    End Sub

    Private Enum elCategory As Int32
        eAgents = 0
        eAlloy '= 0
        eArmor '= 1
        eWeapon_Bomb '= 2
        eColonial
        eCommand '= 3
        eEngine '= 4
        eHull '= 5
        eMiscellaneous '= 6
        eWeapon_Missile '= 7
        eWeapon_Projectile '= 8
        eProduction
        ePrototype '= 9
        eWeapon_Pulse '= 10
        eRadar '= 11
        eShield '= 12
        eWeapon_Solid '= 13
        eSuperTopSecret '= 14
        eTrade '= 15


        eLastCategory
    End Enum
    Private Function GetCategoryName(ByVal lVal As elCategory) As String
        Select Case lVal
            Case elCategory.eAgents
                Return "Agents"
            Case elCategory.eAlloy
                Return "Alloy"
            Case elCategory.eArmor
                Return "Armor"
            Case elCategory.eColonial
                Return "Colonial"
            Case elCategory.eCommand
                Return "Command"
            Case elCategory.eEngine
                Return "Engines"
            Case elCategory.eHull
                Return "Hulls"
            Case elCategory.eProduction
                Return "Production"
            Case elCategory.ePrototype
                Return "Prototypes"
            Case elCategory.eRadar
                Return "Radar"
            Case elCategory.eShield
                Return "Shields"
            Case elCategory.eSuperTopSecret
                Return "Top Secret"
            Case elCategory.eTrade
                Return "Trade"
            Case elCategory.eWeapon_Bomb
                Return "Bomb Weapons"
            Case elCategory.eWeapon_Missile
                Return "Missile Weapons"
            Case elCategory.eWeapon_Projectile
                Return "Projectile Weapons"
            Case elCategory.eWeapon_Pulse
                Return "Pulse Weapons"
            Case elCategory.eWeapon_Solid
                Return "Solid Beam Weapons"
            Case Else
                Return "Miscellaneous"
        End Select
    End Function

    Private Function GetBaseTypeFromProgramControl(ByVal lVal As Int32) As elCategory
        Select Case CType(lVal, PlayerSpecialAttributeID)
            Case PlayerSpecialAttributeID.eAerialUpkeepReduction
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.eAlloyBuilderImprovements
                Return elCategory.eAlloy
            Case PlayerSpecialAttributeID.eAreaEffectJamming
                Return elCategory.eRadar
            Case PlayerSpecialAttributeID.eArmorCheaperHullToHP
                Return elCategory.eArmor
            Case PlayerSpecialAttributeID.eAutomaticMineralMovement
                Return elCategory.eProduction
            Case PlayerSpecialAttributeID.eBaseEnlistedPerTrainer
                Return elCategory.eProduction
            Case PlayerSpecialAttributeID.eBaseOfficerPerTrainer
                Return elCategory.eProduction
            Case PlayerSpecialAttributeID.eBeamResistImprove
                Return elCategory.eArmor
            Case PlayerSpecialAttributeID.eBetterHullToResidence
                Return elCategory.eMiscellaneous
            Case PlayerSpecialAttributeID.eBombMaxAOE
                Return elCategory.eWeapon_Bomb
            Case PlayerSpecialAttributeID.eBombMaxGuidance
                Return elCategory.eWeapon_Bomb
            Case PlayerSpecialAttributeID.eBombMaxRange
                Return elCategory.eWeapon_Bomb
            Case PlayerSpecialAttributeID.eBombMinROF
                Return elCategory.eWeapon_Bomb
            Case PlayerSpecialAttributeID.eBombPayloadType
                Return elCategory.eWeapon_Bomb
            Case PlayerSpecialAttributeID.eBurnResistImprove
                Return elCategory.eArmor
            Case PlayerSpecialAttributeID.eBuyAndSellOrderSlots
                Return elCategory.eTrade
            Case PlayerSpecialAttributeID.eBuyOrderSlots
                Return elCategory.eTrade
            Case PlayerSpecialAttributeID.eChemResistImprove
                Return elCategory.eArmor
            Case PlayerSpecialAttributeID.eColonyMoraleBoost
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.eCommandEspionageAbility
                Return elCategory.eAgents
            Case PlayerSpecialAttributeID.eConstructionEspionageAbility
                Return elCategory.eAgents
            Case PlayerSpecialAttributeID.eControlledGrowth
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.eControlledMorale
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.eCPLimit
                Return elCategory.eCommand
            Case PlayerSpecialAttributeID.eDiplomaticEspionage
                Return elCategory.eAgents
            Case PlayerSpecialAttributeID.eDiscoverTime
                Return elCategory.eMiscellaneous
            Case PlayerSpecialAttributeID.eEngineMaxSpeed
                Return elCategory.eEngine
            Case PlayerSpecialAttributeID.eEnginePowerBonus
                Return elCategory.eEngine
            Case PlayerSpecialAttributeID.eEnginePowerThrustLimit
                Return elCategory.eEngine
            Case PlayerSpecialAttributeID.eEnlistedsColonistsCost
                Return elCategory.eProduction
            Case PlayerSpecialAttributeID.eEnlistedTrainingFactor
                Return elCategory.eProduction
            Case PlayerSpecialAttributeID.eFactoryUpkeepReduction
                Return elCategory.eProduction
            Case PlayerSpecialAttributeID.eGrowthRateBonus
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.eHomelessAllowedBeforePenalty
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.eImpactResistImprove
                Return elCategory.eArmor
            Case PlayerSpecialAttributeID.eJammingEffectAvailable
                Return elCategory.eRadar
            Case PlayerSpecialAttributeID.eJammingImmunityAvailable
                Return elCategory.eRadar
            Case PlayerSpecialAttributeID.eJammingStrength
                Return elCategory.eRadar
            Case PlayerSpecialAttributeID.eJammingTargets
                Return elCategory.eRadar
            Case PlayerSpecialAttributeID.eLinkChanceBonus
                Return elCategory.eSuperTopSecret
            Case PlayerSpecialAttributeID.eMagResistImprove
                Return elCategory.eArmor
            Case PlayerSpecialAttributeID.eManeuverBonus
                Return elCategory.eEngine
            Case PlayerSpecialAttributeID.eManeuverLimit
                Return elCategory.eEngine
            Case PlayerSpecialAttributeID.eMaxAlloyImprovement
                Return elCategory.eAlloy
            Case PlayerSpecialAttributeID.eMaxBattlegroups
                Return elCategory.eCommand
            Case PlayerSpecialAttributeID.eMaxBattlegroupUnits
                Return elCategory.eCommand
            Case PlayerSpecialAttributeID.eMaxnumberofconcurrentLinks
                Return elCategory.eSuperTopSecret
            Case PlayerSpecialAttributeID.eMaxSpeedBonus
                Return elCategory.eEngine
            Case PlayerSpecialAttributeID.eMineralConcentrationBonus
                Return elCategory.eProduction
            Case PlayerSpecialAttributeID.eMiningWorkers
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.eMissileExplosionRadiusBonus
                Return elCategory.eWeapon_Missile
            Case PlayerSpecialAttributeID.eMissileFlightTimeBonus
                Return elCategory.eWeapon_Missile
            Case PlayerSpecialAttributeID.eMissileHullSizeImprove
                Return elCategory.eWeapon_Missile
            Case PlayerSpecialAttributeID.eMissileManeuverBonus
                Return elCategory.eWeapon_Missile
            Case PlayerSpecialAttributeID.eMissileMaxDmgBonus
                Return elCategory.eWeapon_Missile
            Case PlayerSpecialAttributeID.eMissileMaxSpeedImprove
                Return elCategory.eWeapon_Missile
            Case PlayerSpecialAttributeID.eNavalAbility
                Return elCategory.eHull
            Case PlayerSpecialAttributeID.eNonAerialUpkeepReduction
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.eOfficersEnlistedCost
                Return elCategory.eProduction
            Case PlayerSpecialAttributeID.eOfficerTrainingFactor
                Return elCategory.eProduction
            Case PlayerSpecialAttributeID.eOtherFacUpkeepReduction
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.ePayloadTypeAvailable
                Return elCategory.eMiscellaneous
            Case PlayerSpecialAttributeID.ePersonnelCargoUsage
                Return elCategory.eMiscellaneous
            Case PlayerSpecialAttributeID.ePierceResistImprove
                Return elCategory.eArmor
            Case PlayerSpecialAttributeID.ePopulationUpkeepReduction
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.ePowerGenProdToWorker
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.ePrecisionMissiles
                Return elCategory.eWeapon_Missile
            Case PlayerSpecialAttributeID.eProjectileExplosionRadius
                Return elCategory.eWeapon_Projectile
            Case PlayerSpecialAttributeID.eProjectileMaxRng
                Return elCategory.eWeapon_Projectile
            Case PlayerSpecialAttributeID.ePulseBeamInputEnergyMax
                Return elCategory.eWeapon_Pulse
            Case PlayerSpecialAttributeID.ePulseBeamMaxCompressFactor
                Return elCategory.eWeapon_Pulse
            Case PlayerSpecialAttributeID.ePulseBeamMaxDmgImprove
                Return elCategory.eWeapon_Pulse
            Case PlayerSpecialAttributeID.ePulseBeamMinDmgImprove
                Return elCategory.eWeapon_Pulse
            Case PlayerSpecialAttributeID.ePulseBeamOptimumRange
                Return elCategory.eWeapon_Pulse
            Case PlayerSpecialAttributeID.ePulseBeamOptRngImprove
                Return elCategory.eWeapon_Pulse
            Case PlayerSpecialAttributeID.ePulseBeamPowerReduced
                Return elCategory.eWeapon_Pulse
            Case PlayerSpecialAttributeID.ePulseBeamROFImprove
                Return elCategory.eWeapon_Pulse
            Case PlayerSpecialAttributeID.ePulseBeamWpnROF
                Return elCategory.eWeapon_Pulse
            Case PlayerSpecialAttributeID.eRadarDisRes
                Return elCategory.eRadar
            Case PlayerSpecialAttributeID.eRadarMaxRange
                Return elCategory.eRadar
            Case PlayerSpecialAttributeID.eRadarMaxRangeBonus
                Return elCategory.eRadar
            Case PlayerSpecialAttributeID.eRadarOptRange
                Return elCategory.eRadar
            Case PlayerSpecialAttributeID.eRadarOptRangeBonus
                Return elCategory.eRadar
            Case PlayerSpecialAttributeID.eRadarResistImprove
                Return elCategory.eRadar
            Case PlayerSpecialAttributeID.eRadarScanRes
                Return elCategory.eRadar
            Case PlayerSpecialAttributeID.eRadarScanResBonus
                Return elCategory.eRadar
            Case PlayerSpecialAttributeID.eRadarWpnAcc
                Return elCategory.eRadar
            Case PlayerSpecialAttributeID.eRadarWpnAccBonus
                Return elCategory.eRadar
            Case PlayerSpecialAttributeID.eRepairSpeedBonus
                Return elCategory.eProduction
            Case PlayerSpecialAttributeID.eResearchCrewQtrs
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.eResearchEspionage
                Return elCategory.eAgents
            Case PlayerSpecialAttributeID.eResearchFacAssistBonus
                Return elCategory.eProduction
            Case PlayerSpecialAttributeID.eResearchFacUpkeepReduction
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.eResearchProdFactor
                Return elCategory.eProduction
            Case PlayerSpecialAttributeID.eSellOrderSlots
                Return elCategory.eTrade
            Case PlayerSpecialAttributeID.eShieldMaxHP
                Return elCategory.eShield
            Case PlayerSpecialAttributeID.eShieldMaxHPBonus
                Return elCategory.eShield
            Case PlayerSpecialAttributeID.eShieldProjSize
                Return elCategory.eShield
            Case PlayerSpecialAttributeID.eShieldProjSizeBonus
                Return elCategory.eShield
            Case PlayerSpecialAttributeID.eShieldRechargeFreqDecrease
                Return elCategory.eShield
            Case PlayerSpecialAttributeID.eShieldRechargeFreqLow
                Return elCategory.eShield
            Case PlayerSpecialAttributeID.eShieldRechargeRate
                Return elCategory.eShield
            Case PlayerSpecialAttributeID.eShieldRechargeRateBonus
                Return elCategory.eShield
            Case PlayerSpecialAttributeID.eSolidBeamMaxDmgBonus
                Return elCategory.eWeapon_Solid
            Case PlayerSpecialAttributeID.eSolidBeamMinDmgBonus
                Return elCategory.eWeapon_Solid
            Case PlayerSpecialAttributeID.eSolidBeamOptimumRange
                Return elCategory.eWeapon_Solid
            Case PlayerSpecialAttributeID.eSolidBeamOptRangeBonus
                Return elCategory.eWeapon_Solid
            Case PlayerSpecialAttributeID.eSolidBeamPierceRatio
                Return elCategory.eWeapon_Solid
            Case PlayerSpecialAttributeID.eSolidBeamPowerReduced
                Return elCategory.eWeapon_Solid
            Case PlayerSpecialAttributeID.eSolidBeamROFDecrease
                Return elCategory.eWeapon_Solid
            Case PlayerSpecialAttributeID.eSolidBeamWeaponMaxAccuracy
                Return elCategory.eWeapon_Solid
            Case PlayerSpecialAttributeID.eSolidBeamWpnMaxDmg
                Return elCategory.eWeapon_Solid
            Case PlayerSpecialAttributeID.eSolidBeamWpnROF
                Return elCategory.eWeapon_Solid
            Case PlayerSpecialAttributeID.eSpaceportUpkeepReduction
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.eSpecialSpecialAttribute
                Return elCategory.eSuperTopSecret
            Case PlayerSpecialAttributeID.eStudyMineralEffect
                Return elCategory.eAlloy
            Case PlayerSpecialAttributeID.eSuperSpecials
                Return elCategory.eSuperTopSecret
            Case PlayerSpecialAttributeID.eTaxRateWithoutPenalty
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.eTradeBoardRange
                Return elCategory.eTrade
            Case PlayerSpecialAttributeID.eTradeCosts
                Return elCategory.eTrade
            Case PlayerSpecialAttributeID.eTradeGTCSpeed
                Return elCategory.eTrade
            Case PlayerSpecialAttributeID.eTransporters
                Return elCategory.eSuperTopSecret
            Case PlayerSpecialAttributeID.eUnderwaterFacUpkeepReduct
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.eUnemploymentUpkeepReduct
                Return elCategory.eColonial
            Case PlayerSpecialAttributeID.eUnemploymentWithoutPenalty
                Return elCategory.eColonial
        End Select
        Return elCategory.eMiscellaneous
    End Function

    Private Sub btnThrowBack_Click(ByVal sName As String) Handles btnThrowBack.Click
        If gbAliased = True Then
            MyBase.moUILib.AddNotification("Aliases cannot bypass or pull back specials.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If
        If tvwTechs.oSelectedNode Is Nothing = False Then
            Dim lID As Int32 = tvwTechs.oSelectedNode.lItemData
            Dim oTech As SpecialTech = CType(goCurrentPlayer.GetTech(lID, ObjectType.eSpecialTech), SpecialTech)
            If oTech Is Nothing = False Then
                If oTech.bInTheTank = True Then
                    'ok, gotta swap out... are there any that are in the list?
                    Dim lActiveCnt As Int32 = 0
                    Dim lSlotsInUse As Int32 = 0
                    For X As Int32 = 0 To goCurrentPlayer.mlTechUB
                        If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eSpecialTech Then
                            If goCurrentPlayer.moTechs(X).ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched AndAlso CType(goCurrentPlayer.moTechs(X), SpecialTech).bInTheTank = False Then
                                lSlotsInUse += 1
                                If goCurrentPlayer.moTechs(X).Researchers = 0 Then lActiveCnt += 1
                                'Exit For
                            End If
                        End If
                    Next X
                    If lSlotsInUse < 3 Then
                        SendSetSpecialTechThrowback(Nothing, oTech)
                    ElseIf lActiveCnt > 0 Then

                        If moSwapBox Is Nothing Then
                            moSwapBox = New frmSwapTech(MyBase.moUILib)
                            moSwapBox.Visible = False
                            Me.AddChild(CType(moSwapBox, UIControl))
                            AddHandler moSwapBox.SwapButtonClicked, AddressOf swapbox_selected
                            AddHandler moSwapBox.CancelButtonClicked, AddressOf swapbox_cancelled
                        End If
                        moSwapBox.RefreshList()
                        moSwapBox.Left = (Me.Width \ 2) - (moSwapBox.Width \ 2)
                        moSwapBox.Top = (Me.Height \ 2) - (moSwapBox.Height \ 2)
                        moSwapBox.Visible = True

                        For X As Int32 = 0 To Me.ChildrenUB
                            If moChildren(X) Is Nothing = False Then
                                If moChildren(X).ControlName <> moSwapBox.ControlName Then
                                    moChildren(X).Enabled = False
                                End If
                            End If
                        Next X
                    Else
                        MyBase.moUILib.AddNotification("You do not have any projects that can be swapped out at the moment.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    End If
                ElseIf oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
                    moThrowbackTech = oTech

                    Dim oMsgBox As New frmMsgBox(MyBase.moUILib, "This will put the currently selected technology, '" & oTech.sName & "' in Bypass which will make it unavailable until you swap it out with another tech or the end of your cloud has been reached." & vbCrLf & vbCrLf & "Your scientists will then go back to the drawing board and provide you a new possibility as soon as they can come up with one." & vbCrLf & vbCrLf & "Are you sure you wish to Bypass '" & oTech.sName & "'?", MsgBoxStyle.YesNo Or MsgBoxStyle.Critical Or MsgBoxStyle.AbortRetryIgnore, "Confirm Bypass")
                    oMsgBox.Visible = True
                    AddHandler oMsgBox.DialogClosed, AddressOf MsgBoxRes
                Else
                    MyBase.moUILib.AddNotification("Invalid selection.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                End If
            Else
                MyBase.moUILib.AddNotification("Select a technology first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
        Else
            MyBase.moUILib.AddNotification("Select a technology first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Private Sub SendSetSpecialTechThrowback(ByVal oFromQ As SpecialTech, ByVal oToQ As SpecialTech)
        Dim yMsg(12) As Byte
        Dim lPos As Int32 = 0
        System.BitConverter.GetBytes(GlobalMessageCode.eSetSpecialTechThrowback).CopyTo(yMsg, lPos) : lPos += 2

        If oToQ.oResearchCost Is Nothing = False Then
            If goCurrentPlayer.blCredits < (oToQ.oResearchCost.CreditCost \ 10) Then
                MyBase.moUILib.AddNotification("Insufficient credits to bring that project back to the table.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
        End If

        oToQ.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
        If (oToQ.yFlags And 2) <> 0 Then oToQ.yFlags = CByte(oToQ.yFlags Xor 2)

        yMsg(lPos) = eyThrowBackType.eSwapFromBypass : lPos += 1
        If oFromQ Is Nothing = False Then
            System.BitConverter.GetBytes(oFromQ.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
            oFromQ.yFlags = CByte(oFromQ.yFlags Or 2)
        Else
            System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
        End If


        MyBase.moUILib.SendMsgToPrimary(yMsg)

        RefreshLists()

    End Sub

    Private moThrowbackTech As SpecialTech
    Private Sub MsgBoxRes(ByVal lRes As MsgBoxResult)
        If lRes = MsgBoxResult.Yes Then
            'ok, gonna add it to the bypass
            Dim yMsg(8) As Byte
            Dim lPos As Int32 = 0
            System.BitConverter.GetBytes(GlobalMessageCode.eSetSpecialTechThrowback).CopyTo(yMsg, lPos) : lPos += 2
            moThrowbackTech.GetGUIDAsString.CopyTo(yMsg, lPos) : lPos += 6
            yMsg(lPos) = eyThrowBackType.ePushToBypass : lPos += 1
            MyBase.moUILib.SendMsgToPrimary(yMsg)
            moThrowbackTech.yFlags = CByte(moThrowbackTech.yFlags Or 2)

            MyBase.moUILib.AddNotification("Your scientists will attempt to come up with another project.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            RefreshLists()
        End If
    End Sub

    Private Sub swapbox_cancelled()
        moSwapBox.Visible = False
        For X As Int32 = 0 To Me.ChildrenUB
            If moChildren(X) Is Nothing = False Then
                If moChildren(X).ControlName <> moSwapBox.ControlName Then
                    moChildren(X).Enabled = True
                End If
            End If
        Next X
    End Sub
    Private Sub swapbox_selected(ByVal lID As Int32)
        'send our message and then close the box
        If lID > 1 Then
            
            Dim oNode As UITreeView.UITreeViewItem = tvwTechs.oSelectedNode
            If oNode Is Nothing = False Then

                Dim oFromQ As SpecialTech = CType(goCurrentPlayer.GetTech(lID, ObjectType.eSpecialTech), SpecialTech)
                Dim oToQ As SpecialTech = CType(goCurrentPlayer.GetTech(oNode.lItemData, ObjectType.eSpecialTech), SpecialTech)
                If oFromQ Is Nothing = False AndAlso oToQ Is Nothing = False Then
                    SendSetSpecialTechThrowback(oFromQ, oToQ)

                End If
            End If
        End If

        swapbox_cancelled()
    End Sub

    'Interface created from Interface Builder
    Private Class frmSwapTech
        Inherits UIWindow

        Private lstInQueue As UIListBox
        Private WithEvents btnSwap As UIButton
        Private WithEvents btnCancel As UIButton

        Public Event CancelButtonClicked()
        Public Event SwapButtonClicked(ByVal lID As Int32)

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            ' initial props
            With Me
                .ControlName = "frmSwapTech"
                .Left = 275
                .Top = 174
                .Width = 256
                .Height = 145
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
            End With

            'lstInQueue initial props
            lstInQueue = New UIListBox(oUILib)
            With lstInQueue
                .ControlName = "lstInQueue"
                .Left = 5
                .Top = 11
                .Width = 245
                .Height = 100
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            End With
            Me.AddChild(CType(lstInQueue, UIControl))

            'btnSwap initial props
            btnSwap = New UIButton(oUILib)
            With btnSwap
                .ControlName = "btnSwap"
                .Left = 5
                .Top = 116
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Swap"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSwap, UIControl))

            'btnCancel initial props
            btnCancel = New UIButton(oUILib)
            With btnCancel
                .ControlName = "btnCancel"
                .Left = 152
                .Top = 116
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Cancel"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnCancel, UIControl))

        End Sub

        Public Sub RefreshList()
            lstInQueue.Clear()
            For X As Int32 = 0 To goCurrentPlayer.mlTechUB
                If goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eSpecialTech AndAlso goCurrentPlayer.moTechs(X).ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
                    With CType(goCurrentPlayer.moTechs(X), SpecialTech)
                        If .bInTheTank = True Then Continue For
                        If .Researchers <> 0 Then Continue For

                        lstInQueue.AddItem(.sName, False)
                        lstInQueue.ItemData(lstInQueue.NewIndex) = .ObjectID
                        lstInQueue.ItemData2(lstInQueue.NewIndex) = .ObjTypeID
                    End With
                End If
            Next X
        End Sub

        Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
            RaiseEvent CancelButtonClicked()
        End Sub

        Private Sub btnSwap_Click(ByVal sName As String) Handles btnSwap.Click
            If lstInQueue.ListIndex > -1 Then
                RaiseEvent SwapButtonClicked(lstInQueue.ItemData(lstInQueue.ListIndex))
            Else
                MyBase.moUILib.AddNotification("You must select a technology to swap with.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            End If
            Dim oWin As frmColonyResearch = CType(goUILib.GetWindow("frmColonyResearch"), frmColonyResearch)
            If oWin Is Nothing = False Then oWin.ViewState = oWin.ViewState
            oWin = Nothing
        End Sub
    End Class

    Private Sub frmSpecTech_WindowMoved() Handles Me.WindowMoved
        If mbLoading = True Then Return
        muSettings.SpecialTechLeft = Me.Left
        muSettings.SpecialTechTop = Me.Top
    End Sub

    Private Sub btnExport_Click(ByVal sName As String) Handles btnExport.Click
        Export_SpecialTech()
    End Sub

    Private Sub Export_SpecialTech()
        If goCurrentPlayer Is Nothing Then Return
        If muSettings.ExportedDataFormat = 1 Then
            Export_SpecialTech_Csv()
        ElseIf muSettings.ExportedDataFormat = 2 Then
            'Export_Specialech_Xml()
        End If
    End Sub

    Private Sub Export_SpecialTech_Csv()
        Dim sExportData As String = ""
        sExportData &= "Name,Benefits,CreditCost,TimeCost,Status,Assigned" & vbCrLf

        For X As Int32 = 0 To goCurrentPlayer.mlTechUB
            If goCurrentPlayer.moTechs(X) Is Nothing = False AndAlso goCurrentPlayer.moTechs(X).ObjTypeID = ObjectType.eSpecialTech Then
                Dim oSpecTech As SpecialTech = CType(goCurrentPlayer.GetTech(goCurrentPlayer.moTechs(X).ObjectID, ObjectType.eSpecialTech), SpecialTech)
                'sExportData = ""
                'sExportData &= oSpecTech.ObjectID
                'sExportData &= ","
                sExportData &= oSpecTech.sName
                sExportData &= ","
                sExportData &= Replace(oSpecTech.sBenefits, ",", ":")
                sExportData &= ","
                sExportData &= oSpecTech.oResearchCost.CreditCost
                sExportData &= ","
                sExportData &= oSpecTech.oResearchCost.PointsRequired.ToString
                sExportData &= ","
                If oSpecTech.bInTheTank = True Then
                    sExportData &= "Bypassed"
                ElseIf oSpecTech.ComponentDevelopmentPhase = Base_Tech.eComponentDevelopmentPhase.eResearched Then
                    sExportData &= "Researched"
                Else
                    sExportData &= "Unresearched"
                    If oSpecTech.Researchers > 0 Then
                        sExportData &= ","
                        sExportData &= oSpecTech.Researchers
                    End If
                End If
                sExportData &= vbCrLf
            End If
        Next X
        If sExportData = "" Then Return
        Dim sFile As String = AppDomain.CurrentDomain.BaseDirectory
        If sFile.EndsWith("\") = False Then sFile = sFile & "\"
        sFile &= "ExportedData"
        If Exists(sFile) = False Then MkDir(sFile)
        If sFile.EndsWith("\") = False Then sFile &= "\"
        sFile &= "Specials_" & goCurrentPlayer.PlayerName & "_" & Now.ToString("MM_dd_yyyy_HHmmss") & ".csv"

        Dim fsFile As IO.FileStream = New IO.FileStream(sFile, IO.FileMode.Create)

        Dim info As Byte() = New System.Text.UTF8Encoding(True).GetBytes(sExportData)
        fsFile.Write(info, 0, info.Length)
        fsFile.Close()
        fsFile.Dispose()
        If goUILib Is Nothing = False Then
            goUILib.AddNotification("Special Tech Exported.", Color.White, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub
End Class
