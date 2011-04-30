Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmEmailSetup
    Inherits UIWindow

    Private lblTitle As UILabel
    Private lnDiv1 As UILine
    Private lblAddress As UILabel
    Private txtAddress As UITextBox
	Private WithEvents chkPlayerRel As UICheckBox
	Private WithEvents chkLowRes As UICheckBox
	Private WithEvents chkEngaged As UICheckBox
	Private WithEvents chkUnderAttack As UICheckBox
    'Private WithEvents chkBuyOrderAccepted As UICheckBox
    Private chkAllInternal As UICheckBox
	Private WithEvents chkResearchComplete As UICheckBox
	Private WithEvents chkTradeRequests As UICheckBox
	Private WithEvents chkRebuildAI As UICheckBox

	Private lblInt As UILabel
	Private lblExt As UILabel
	Private chkAgentUpdates As UICheckBox
	Private chkGuildNotices As UICheckBox
    Private WithEvents chkMineralOutbid As UICheckBox
    Private chkFacilityLost As UICheckBox
    Private chkColonyLost As UICheckBox
    Private chkTitleChange As UICheckBox
    Private chkFactionUpdates As UICheckBox
    Private WithEvents chkPlayerRelExt As UICheckBox
    Private WithEvents chkLowResExt As UICheckBox
    Private WithEvents chkEngagedExt As UICheckBox
    Private WithEvents chkUnderAttackExt As UICheckBox
    'Private WithEvents chkBuyOrderAcceptExt As UICheckBox
    Private WithEvents chkResearchCompleteExt As UICheckBox
    Private WithEvents chkTradeRequestsExt As UICheckBox
    Private WithEvents chkRebuildAIExt As UICheckBox
    Private WithEvents chkMineralOutbidExt As UICheckBox

    Private WithEvents btnClose As UIButton
    Private WithEvents btnSubmit As UIButton

    Private mlLastForcefulUpdate As Int32 = Int32.MinValue

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmEmailSetup initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eEmailSetup
            .ControlName = "frmEmailSetup"
            .Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            .Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 128
            .Width = 256
            .Height = 415 '256
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 5
            .Width = 158
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Email Alert Setup"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 231
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
        End With
        Me.AddChild(CType(btnClose, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 1
            .Top = 25
            .Width = 255
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'lblAddress initial props
        lblAddress = New UILabel(oUILib)
        With lblAddress
            .ControlName = "lblAddress"
            .Left = 5
            .Top = 30
            .Width = 55
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Address:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblAddress, UIControl))

        'txtAddress initial props
        txtAddress = New UITextBox(oUILib)
        With txtAddress
            .ControlName = "txtAddress"
            .Left = 60
            .Top = 30
            .Width = 190
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 255
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(txtAddress, UIControl))

        'chkAllInternal initial props
        chkAllInternal = New UICheckBox(oUILib)
        With chkAllInternal
            .ControlName = "chkAllInternal"
            .Left = 20
            .Top = 70
            .Width = 160
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "All Internal Emails"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .ToolTipText = "If checked, all emails received will be forwarded to your external address" & vbCrLf & _
                           "except for those specifically indicated as not to send externally below."
        End With
        Me.AddChild(CType(chkAllInternal, UIControl))

        'chkPlayerRel initial props
        chkPlayerRel = New UICheckBox(oUILib)
        With chkPlayerRel
            .ControlName = "chkPlayerRel"
            .Left = 55 '35
            .Top = 90 '70 
            .Width = 192
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Player Relationship Changes"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkPlayerRel, UIControl))

        'chkLowRes initial props
        chkLowRes = New UICheckBox(oUILib)
        With chkLowRes
            .ControlName = "chkLowRes"
            .Left = 55 '35
            .Top = 110 '90 
            .Width = 141
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Low Resource Alerts"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkLowRes, UIControl))

        'chkEngaged initial props
        chkEngaged = New UICheckBox(oUILib)
        With chkEngaged
            .ControlName = "chkEngaged"
            .Left = 55 '35
            .Top = 130 '110 
            .Width = 156
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Engaged Enemy Alerts"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkEngaged, UIControl))

        'chkUnderAttack initial props
        chkUnderAttack = New UICheckBox(oUILib)
        With chkUnderAttack
            .ControlName = "chkUnderAttack"
            .Left = 55 '35
            .Top = 150 '130 
            .Width = 132
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Under Attack Alerts"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkUnderAttack, UIControl))

        ''chkBuyOrderAccepted initial props
        'chkBuyOrderAccepted = New UICheckBox(oUILib)
        'With chkBuyOrderAccepted
        '    .ControlName = "chkBuyOrderAccepted"
        '    .Left = 55 ' 35
        '    .Top = 150 '135
        '    .Width = 139
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = "Buy Order Accepted"
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(6, DrawTextFormat)
        '    .Value = False
        'End With
        'Me.AddChild(CType(chkBuyOrderAccepted, UIControl))


        'chkResearchComplete  initial props
        chkResearchComplete = New UICheckBox(oUILib)
        With chkResearchComplete
            .ControlName = "chkResearchComplete"
            .Left = 55 '35
            .Top = 170 '155
            .Width = 139
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Research Complete"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkResearchComplete, UIControl))

        'chkTradeRequests initial props
        chkTradeRequests = New UICheckBox(oUILib)
        With chkTradeRequests
            .ControlName = "chkTradeRequests"
            .Left = 55 '35
            .Top = 190 '175
            .Width = 117
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Trade Requests"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkTradeRequests, UIControl))

        'chkRebuildAI initial props
        chkRebuildAI = New UICheckBox(oUILib)
        With chkRebuildAI
            .ControlName = "chkRebuildAI"
            .Left = 55 '35
            .Top = 210 '195
            .Width = 81
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Rebuild AI"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkRebuildAI, UIControl))

        'lblInt initial props
        lblInt = New UILabel(oUILib)
        With lblInt
            .ControlName = "lblInt"
            .Left = 55
            .Top = 53
            .Width = 43
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Internal"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblInt, UIControl))

        'lblExt initial props
        lblExt = New UILabel(oUILib)
        With lblExt
            .ControlName = "lblExt"
            .Left = 5
            .Top = 53
            .Width = 41
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "External"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblExt, UIControl))

        'chkAgentUpdates initial props
        chkAgentUpdates = New UICheckBox(oUILib)
        With chkAgentUpdates
            .ControlName = "chkAgentUpdates"
            .Left = 55
            .Top = 230
            .Width = 109
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Agent Updates"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkAgentUpdates, UIControl))

        'chkGuildNotices initial props
        chkGuildNotices = New UICheckBox(oUILib)
        With chkGuildNotices
            .ControlName = "chkGuildNotices"
            .Left = 55
            .Top = 250
            .Width = 177
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Guild Membership Notices"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkGuildNotices, UIControl))

        'chkUnitLost initial props
        chkMineralOutbid = New UICheckBox(oUILib)
        With chkMineralOutbid
            .ControlName = "chkMineralOutbid"
            .Left = 55
            .Top = 270
            .Width = 105
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Mineral Outbid"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkMineralOutbid, UIControl))

        'chkFacilityLost initial props
        chkFacilityLost = New UICheckBox(oUILib)
        With chkFacilityLost
            .ControlName = "chkFacilityLost"
            .Left = 55
            .Top = 290
            .Width = 90
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Facility Lost"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkFacilityLost, UIControl))

        'chkColonyLost initial props
        chkColonyLost = New UICheckBox(oUILib)
        With chkColonyLost
            .ControlName = "chkColonyLost"
            .Left = 55
            .Top = 310
            .Width = 90
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Colony Lost"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkColonyLost, UIControl))

        'chkTitleChange initial props
        chkTitleChange = New UICheckBox(oUILib)
        With chkTitleChange
            .ControlName = "chkTitleChange"
            .Left = 55
            .Top = 330
            .Width = 97
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Title Change"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkTitleChange, UIControl))

        'chkFactionUpdates initial props
        chkFactionUpdates = New UICheckBox(oUILib)
        With chkFactionUpdates
            .ControlName = "chkFactionUpdates"
            .Left = 55
            .Top = 350
            .Width = 120
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Faction Updates"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkFactionUpdates, UIControl))

        'chkPlayerRelExt initial props
        chkPlayerRelExt = New UICheckBox(oUILib)
        With chkPlayerRelExt
            .ControlName = "chkPlayerRelExt"
            .Left = 20
            .Top = chkPlayerRel.Top
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkPlayerRelExt, UIControl))

        'chkLowResExt initial props
        chkLowResExt = New UICheckBox(oUILib)
        With chkLowResExt
            .ControlName = "chkLowResExt"
            .Left = 20
            .Top = chkLowRes.Top
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkLowResExt, UIControl))

        'chkEngagedExt initial props
        chkEngagedExt = New UICheckBox(oUILib)
        With chkEngagedExt
            .ControlName = "chkEngagedExt"
            .Left = 20
            .Top = chkEngaged.Top
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkEngagedExt, UIControl))

        'chkUnderAttackExt initial props
        chkUnderAttackExt = New UICheckBox(oUILib)
        With chkUnderAttackExt
            .ControlName = "chkUnderAttackExt"
            .Left = 20
            .Top = chkUnderAttack.Top
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkUnderAttackExt, UIControl))

        ''chkBuyOrderAcceptExt initial props
        'chkBuyOrderAcceptExt = New UICheckBox(oUILib)
        'With chkBuyOrderAcceptExt
        '    .ControlName = "chkBuyOrderAcceptExt"
        '    .Left = 20
        '    .Top = 150
        '    .Width = 18
        '    .Height = 18
        '    .Enabled = True
        '    .Visible = True
        '    .Caption = ""
        '    .ForeColor = muSettings.InterfaceBorderColor
        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
        '    .DrawBackImage = False
        '    .FontFormat = CType(6, DrawTextFormat)
        '    .Value = False
        'End With
        'Me.AddChild(CType(chkBuyOrderAcceptExt, UIControl))

        'chkResearchCompleteExt initial props
        chkResearchCompleteExt = New UICheckBox(oUILib)
        With chkResearchCompleteExt
            .ControlName = "chkResearchCompleteExt"
            .Left = 20
            .Top = 170
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkResearchCompleteExt, UIControl))

        'chkTradeRequestsExt initial props
        chkTradeRequestsExt = New UICheckBox(oUILib)
        With chkTradeRequestsExt
            .ControlName = "chkTradeRequestsExt"
            .Left = 20
            .Top = 190
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkTradeRequestsExt, UIControl))

        'chkRebuildAIExt initial props
        chkRebuildAIExt = New UICheckBox(oUILib)
        With chkRebuildAIExt
            .ControlName = "chkRebuildAIExt"
            .Left = 20
            .Top = 210
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkRebuildAIExt, UIControl))

        'chkMineralOutbidExt initial props
        chkMineralOutbidExt = New UICheckBox(oUILib)
        With chkMineralOutbidExt
            .ControlName = "chkMineralOutbidExt"
            .Left = 20
            .Top = 270
            .Width = 18
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(chkMineralOutbidExt, UIControl))

        'btnSubmit initial props
        btnSubmit = New UIButton(oUILib)
        With btnSubmit
            .ControlName = "btnSubmit"
            .Left = 80
            .Top = 380 ' 220
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Submit"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnSubmit, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        SendSubmitMsg(-1, -1, "", False)
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnSubmit_Click(ByVal sName As String) Handles btnSubmit.Click
        Dim iVals As Int16 = 0
        If chkPlayerRelExt.Value = True Then iVals = iVals Or 1S
        If chkLowResExt.Value = True Then iVals = iVals Or 2S
        If chkEngagedExt.Value = True Then iVals = iVals Or 4S
        If chkUnderAttackExt.Value = True Then iVals = iVals Or 8S
        'If chkBuyOrderAcceptExt.Value = True Then iVals = iVals Or 16S
        If chkAllInternal.Value = True Then iVals = iVals Or 16S
        If chkResearchCompleteExt.Value = True Then iVals = iVals Or 32S
        If chkTradeRequestsExt.Value = True Then iVals = iVals Or 64S
        If chkRebuildAIExt.Value = True Then iVals = iVals Or 128S
        If chkMineralOutbidExt.Value = True Then iVals = iVals Or 1024S

        Dim iInternalVals As Int16 = 0
        If chkPlayerRel.Value = True Then iInternalVals = iInternalVals Or 1S
        If chkLowRes.Value = True Then iInternalVals = iInternalVals Or 2S
        If chkEngaged.Value = True Then iInternalVals = iInternalVals Or 4S
        If chkUnderAttack.Value = True Then iInternalVals = iInternalVals Or 8S
        'If chkBuyOrderAccepted.Value = True Then iInternalVals = iInternalVals Or 16S
        If chkResearchComplete.Value = True Then iInternalVals = iInternalVals Or 32S
        If chkTradeRequests.Value = True Then iInternalVals = iInternalVals Or 64S
        If chkRebuildAI.Value = True Then iInternalVals = iInternalVals Or 128S
        If chkAgentUpdates.Value = True Then iInternalVals = iInternalVals Or 256S
        If chkGuildNotices.Value = True Then iInternalVals = iInternalVals Or 512S
        If chkMineralOutbid.Value = True Then iInternalVals = iInternalVals Or 1024S
        If chkFacilityLost.Value = True Then iInternalVals = iInternalVals Or 2048S
        If chkColonyLost.Value = True Then iInternalVals = iInternalVals Or 4096S
        If chkTitleChange.Value = True Then iInternalVals = iInternalVals Or 8192S
        If chkFactionUpdates.Value = True Then iInternalVals = iInternalVals Or 16384S

        SendSubmitMsg(iVals, iInternalVals, txtAddress.Caption, True)
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

	Private Sub SendSubmitMsg(ByVal iVals As Int16, ByVal iInternalVals As Int16, ByVal sNewAddress As String, ByVal bSendAddr As Boolean)
		Dim yMsg() As Byte
		Dim lPos As Int32 = 0

        If sNewAddress Is Nothing Then sNewAddress = ""
        sNewAddress = sNewAddress.Trim
        If sNewAddress <> "" Then
            Dim bResult As Boolean = True
            If sNewAddress.Contains("@") = False Then bResult = False
            If sNewAddress.Contains(".") = False Then bResult = False
            If sNewAddress.Contains("<") = True Then bResult = False
            If sNewAddress.Contains(">") = True Then bResult = False
            If sNewAddress.Contains("(") = True Then bResult = False
            If sNewAddress.Contains(")") = True Then bResult = False
            If sNewAddress.Contains("[") = True Then bResult = False
            If sNewAddress.Contains("]") = True Then bResult = False
            If sNewAddress.Contains(";") = True Then bResult = False
            If sNewAddress.Contains(":") = True Then bResult = False
            If sNewAddress.Contains(",") = True Then bResult = False
            If sNewAddress.Contains("\") = True Then bResult = False
            If sNewAddress.Contains(" ") = True Then bResult = False

            'space, pound-sign
            If bResult = False Then
                MyBase.moUILib.AddNotification("Please enter a valid email address.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            Dim lAtIdx As Int32 = sNewAddress.LastIndexOf("@")
            Dim lDotIdx As Int32 = sNewAddress.LastIndexOf(".")
            If lDotIdx < lAtIdx Then bResult = False
            Dim sChrsAtEnd As String = sNewAddress.Substring(lDotIdx)
            If sChrsAtEnd.Length < 3 Then bResult = False
            Dim sChrsBeforeAt As String = sNewAddress.Substring(0, lAtIdx)
            If sChrsBeforeAt.Length = 0 Then bResult = False
            If sChrsBeforeAt.Contains("@") = True Then bResult = False
            If "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789".Contains(sNewAddress.Substring(0, 1)) = False Then bResult = False

            If bResult = False Then
                MyBase.moUILib.AddNotification("Please enter a valid email address.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If
        End If

		If bSendAddr = True Then
			ReDim yMsg(264)
		Else : ReDim yMsg(9)
		End If

		System.BitConverter.GetBytes(GlobalMessageCode.eEmailSettings).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(glPlayerID).CopyTo(yMsg, lPos) : lPos += 4
		System.BitConverter.GetBytes(iVals).CopyTo(yMsg, lPos) : lPos += 2
		System.BitConverter.GetBytes(iInternalVals).CopyTo(yMsg, lPos) : lPos += 2

		If bSendAddr = True Then
			System.Text.ASCIIEncoding.ASCII.GetBytes(sNewAddress.Trim).CopyTo(yMsg, lPos) : lPos += 255
		End If

		MyBase.moUILib.SendMsgToPrimary(yMsg)
	End Sub

    Public Sub HandleEmailSettingsMsg(ByRef yData() As Byte)
        Dim lPlayerID As Int32 = System.BitConverter.ToInt32(yData, 2)
		Dim iVals As Int16 = System.BitConverter.ToInt16(yData, 6)
		Dim iInternalVals As Int16 = System.BitConverter.ToInt16(yData, 8)
		Dim sEmailAddr As String = GetStringFromBytes(yData, 10, 255)

		chkPlayerRelExt.Value = (iVals And 1) <> 0
		chkLowResExt.Value = (iVals And 2) <> 0
		chkEngagedExt.Value = (iVals And 4) <> 0
		chkUnderAttackExt.Value = (iVals And 8) <> 0
        'chkBuyOrderAcceptExt.Value = (iVals And 16) <> 0
        chkAllInternal.Value = (iVals And 16) <> 0
		chkResearchCompleteExt.Value = (iVals And 32) <> 0
		chkTradeRequestsExt.Value = (iVals And 64) <> 0
        chkRebuildAIExt.Value = (iVals And 128) <> 0
        chkMineralOutbidExt.Value = (iVals And 1024) <> 0

		chkPlayerRel.Value = (iInternalVals And 1) <> 0
		chkLowRes.Value = (iInternalVals And 2) <> 0
		chkEngaged.Value = (iInternalVals And 4) <> 0
		chkUnderAttack.Value = (iInternalVals And 8) <> 0
        'chkBuyOrderAccepted.Value = (iInternalVals And 16) <> 0
		chkResearchComplete.Value = (iInternalVals And 32) <> 0
		chkTradeRequests.Value = (iInternalVals And 64) <> 0
		chkRebuildAI.Value = (iInternalVals And 128) <> 0
		chkAgentUpdates.Value = (iInternalVals And 256) <> 0
		chkGuildNotices.Value = (iInternalVals And 512) <> 0
        chkMineralOutbid.Value = (iInternalVals And 1024) <> 0
		chkFacilityLost.Value = (iInternalVals And 2048) <> 0
		chkColonyLost.Value = (iInternalVals And 4096) <> 0
		chkTitleChange.Value = (iInternalVals And 8192) <> 0
		chkFactionUpdates.Value = (iInternalVals And 16384) <> 0

		txtAddress.Caption = sEmailAddr
    End Sub

	Private Sub frmEmailSetup_OnNewFrame() Handles Me.OnNewFrame
		If mlLastForcefulUpdate = Int32.MinValue OrElse glCurrentCycle - mlLastForcefulUpdate > 30 Then
			Me.IsDirty = True
		End If
	End Sub

    'Private Sub chkBuyOrderAccepted_Click() Handles chkBuyOrderAccepted.Click
    '	If chkBuyOrderAccepted.Value = False Then chkBuyOrderAcceptExt.Value = False
    'End Sub

    'Private Sub chkBuyOrderAcceptExt_Click() Handles chkBuyOrderAcceptExt.Click
    '	If chkBuyOrderAcceptExt.Value = True Then chkBuyOrderAccepted.Value = True
    'End Sub

	Private Sub chkEngaged_Click() Handles chkEngaged.Click
		If chkEngaged.Value = False Then chkEngagedExt.Value = False
	End Sub

	Private Sub chkEngagedExt_Click() Handles chkEngagedExt.Click
		If chkEngagedExt.Value = True Then chkEngaged.Value = True
	End Sub

	Private Sub chkLowRes_Click() Handles chkLowRes.Click
		If chkLowRes.Value = False Then chkLowResExt.Value = False
	End Sub

	Private Sub chkLowResExt_Click() Handles chkLowResExt.Click
		If chkLowResExt.Value = True Then chkLowRes.Value = True
	End Sub

	Private Sub chkPlayerRel_Click() Handles chkPlayerRel.Click
		If chkPlayerRel.Value = False Then chkPlayerRelExt.Value = False
	End Sub

	Private Sub chkPlayerRelExt_Click() Handles chkPlayerRelExt.Click
		If chkPlayerRelExt.Value = True Then chkPlayerRel.Value = True
	End Sub

	Private Sub chkRebuildAI_Click() Handles chkRebuildAI.Click
		If chkRebuildAI.Value = False Then chkRebuildAIExt.Value = False
	End Sub

	Private Sub chkRebuildAIExt_Click() Handles chkRebuildAIExt.Click
		If chkRebuildAIExt.Value = True Then chkRebuildAI.Value = True
    End Sub

    Private Sub chkMineralOutbid_Click() Handles chkMineralOutbid.Click
        If chkMineralOutbid.Value = False Then chkMineralOutbidExt.Value = False
    End Sub

    Private Sub chkMineralOutbidExt_Click() Handles chkMineralOutbidExt.Click
        If chkMineralOutbidExt.Value = True Then chkMineralOutbid.Value = True
    End Sub

	Private Sub chkResearchComplete_Click() Handles chkResearchComplete.Click
		If chkResearchComplete.Value = False Then chkResearchCompleteExt.Value = False
	End Sub

	Private Sub chkResearchCompleteExt_Click() Handles chkResearchCompleteExt.Click
		If chkResearchCompleteExt.Value = True Then chkResearchComplete.Value = True
	End Sub

	Private Sub chkTradeRequests_Click() Handles chkTradeRequests.Click
		If chkTradeRequests.Value = False Then chkTradeRequestsExt.Value = False
	End Sub

	Private Sub chkTradeRequestsExt_Click() Handles chkTradeRequestsExt.Click
		If chkTradeRequestsExt.Value = True Then chkTradeRequests.Value = True
	End Sub

	Private Sub chkUnderAttack_Click() Handles chkUnderAttack.Click
		If chkUnderAttack.Value = False Then chkUnderAttackExt.Value = False
	End Sub

	Private Sub chkUnderAttackExt_Click() Handles chkUnderAttackExt.Click
		If chkUnderAttackExt.Value = True Then chkUnderAttack.Value = True
	End Sub
End Class