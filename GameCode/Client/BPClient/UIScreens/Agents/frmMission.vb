Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
#Region "current code"
Public Class frmMission

    Inherits UIWindow

    Private Enum eTargetFilterType As Byte
        eUnusedTarget = 0       'the target is unused
        ''' <summary>
        ''' Component belonging to the targeted player
        ''' </summary>
        ''' <remarks></remarks>
        eTarget_Component = 1
        ''' <summary>
        ''' Agent belonging to the targeted player
        ''' </summary>
        ''' <remarks></remarks>
        eTarget_Agent
        ''' <summary>
        ''' An agent belonging to Current Player that is Captured by Target player
        ''' </summary>
        ''' <remarks></remarks>
        eCaptured_Agent
        ''' <summary>
        ''' A known Mineral
        ''' </summary>
        ''' <remarks></remarks>
        eMineral
        ''' <summary>
        ''' Production Type list
        ''' </summary>
        ''' <remarks></remarks>
        eProductionType
        ''' <summary>
        ''' Known Players
        ''' </summary>
        ''' <remarks></remarks>
        ePlayer
        ''' <summary>
        ''' Player Missions (non-state-specific)
        ''' </summary>
        ''' <remarks></remarks>
        ePlayerMissions
        ''' <summary>
        ''' Will place an item in the list to select a location... which will change UI state to selecting point
        ''' </summary>
        ''' <remarks></remarks>
        ePICK_A_LOCATION
        ''' <summary>
        ''' Known Units or facilities belonging to the target player
        ''' </summary>
        ''' <remarks></remarks>
        eTarget_UnitOrFacility
        ''' <summary>
        ''' Known units belonging to the target player
        ''' </summary>
        ''' <remarks></remarks>
        eTarget_Unit
        ''' <summary>
        ''' Known facilities belonging to the target player
        ''' </summary>
        ''' <remarks></remarks>
        eTarget_Facility
        ''' <summary>
        ''' A solar system
        ''' </summary>
        ''' <remarks></remarks>
        eSolarSystem
        ''' <summary>
        ''' A Planet
        ''' </summary>
        ''' <remarks></remarks>
        ePlanet
        ''' <summary>
        ''' The current system and all planets within that system will be listed
        ''' </summary>
        ''' <remarks></remarks>
        eSystemAndPlanets
        ''' <summary>
        ''' Colony belonging to the target player
        ''' </summary>
        ''' <remarks></remarks>
        eTarget_Colony

        ''' <summary>
        ''' Special tech in the target player's arsenal that has not been fully researched
        ''' </summary>
        ''' <remarks></remarks>
        eTargetSpecialTechUnresearched

        eMineralOrTargetComponent
        eTarget_ComponentOrKnownMineral
        ''' <summary>
        ''' bit-wise (use AND/OR on value to determine) - indicates that the ANY work will be included in the list
        ''' </summary>
        ''' <remarks></remarks>
        eAny = 128
    End Enum

    Private lblTitle As UILabel
    Private lblMission As UILabel
    Private lblMissionDesc As UILabel
    Private lblTarget As UILabel
    Private lblMethod As UILabel
    Private lblMethodDesc As UILabel
    Private fraGoals As UIWindow
    Private lblGoalTime As UILabel
    Private lblGoalRisk As UILabel
    Private lblGoalName As UILabel
    Private lblGoalPhase As UILabel
    Private txtGoalDesc As UITextBox
    Private lnDiv1 As UILine
    Private lblReqInfType As UILabel
    Private WithEvents fraSkillSets As fraSkillSet
    Private WithEvents cboMission As UIComboBox
    Private WithEvents cboTarget1 As UIComboBox
    Private WithEvents cboTarget2 As UIComboBox
    Private WithEvents cboTarget3 As UIComboBox
    Private WithEvents cboMethod As UIComboBox
    Private WithEvents lstGoals As UIListBox
    Private WithEvents btnClose As UIButton
    Private WithEvents btnHelp As UIButton

    Private WithEvents chkSafehouse As UICheckBox

    Private lstCoverAgents As UIListBox
    Private WithEvents btnAssignCover As UIButton
    Private WithEvents btnRemoveCover As UIButton

    Private WithEvents btnExecuteNow As UIButton
    Private WithEvents btnSaveForLater As UIButton
    Private WithEvents btnCancel As UIButton

    Private WithEvents btnAgentLeft As UIButton
    Private WithEvents btnAgentRight As UIButton

    Private mbMouseOver() As Boolean
    Private moMission As PlayerMission

    Private Const ml_MOUSE_OVER_DELAY As Int32 = 45     '1.5 seconds
    Private moAgentPop As frmAgent = Nothing
    Private mlMouseOverEvent As Int32 = -1

    Private mlAgentBarCnt As Int32
    Private mlScrollValue As Int32 = 0
    Private mlScrollMaxValue As Int32 = 0
    Private mlSelectedAgentID As Int32 = -1

    Private mbHasUnknowns As Boolean = False

    Private myTarget1ListType As eTargetFilterType = eTargetFilterType.ePlayer   'default to player
    Private myTarget2ListType As eTargetFilterType = eTargetFilterType.eUnusedTarget
    Private myTarget3ListType As eTargetFilterType = eTargetFilterType.eUnusedTarget

    Private moFilterSkillSet As SkillSet = Nothing
    Private mbLoading As Boolean = True

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenCreateMissionWindow)

        'frmMission initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eCreateMission
            .ControlName = "frmMission"

            .Width = 800
            .Height = 670
            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1
            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.AgentMissionCreateX
                lTop = muSettings.AgentMissionCreateY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 400
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 310
            If lLeft + .Width > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If lTop + .Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Left = lLeft
            .Top = lTop

            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 2
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 5
            .Top = 3
            .Width = 175
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Mission Management"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTitle, UIControl))

        'lblMission initial props
        lblMission = New UILabel(oUILib)
        With lblMission
            .ControlName = "lblMission"
            .Left = 10
            .Top = 35
            .Width = 70
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Mission:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblMission, UIControl))

        'lblMissionDesc initial props
        lblMissionDesc = New UILabel(oUILib)
        With lblMissionDesc
            .ControlName = "lblMissionDesc"
            .Left = 310
            .Top = 35
            .Width = 490
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblMissionDesc, UIControl))

        'lblTarget initial props
        lblTarget = New UILabel(oUILib)
        With lblTarget
            .ControlName = "lblTarget"
            .Left = 10
            .Top = 60
            .Width = 70
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Target:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblTarget, UIControl))

        'lblMethod initial props
        lblMethod = New UILabel(oUILib)
        With lblMethod
            .ControlName = "lblMethod"
            .Left = 10
            .Top = 85
            .Width = 70
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Method:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblMethod, UIControl))

        'lblMethodDesc initial props
        lblMethodDesc = New UILabel(oUILib)
        With lblMethodDesc
            .ControlName = "lblMethodDesc"
            .Left = 310
            .Top = 85
            .Width = 490
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblMethodDesc, UIControl))

        'fraGoals initial props
        fraGoals = New UIWindow(oUILib)
        With fraGoals
            .ControlName = "fraGoals"
            .Left = 10
            .Top = 120
            .Width = 780
            .Height = 125
            .Enabled = False
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
            .Caption = "Mission Goals"
            .BorderLineWidth = 2
        End With
        Me.AddChild(CType(fraGoals, UIControl))

        'chkSafehouse initial props
        chkSafehouse = New UICheckBox(oUILib)
        With chkSafehouse
            .ControlName = "chkSafehouse"
            .Left = 20
            .Top = 127
            .Width = 140
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Create Safehouse"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
            .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
            .ToolTipText = "Indicates whether or not to include the create safehouse goal." & vbCrLf & "Safehouses provide bonuses to preparation and re-infiltration."
        End With
        Me.AddChild(CType(chkSafehouse, UIControl))

        'lstGoals initial props
        lstGoals = New UIListBox(oUILib)
        With lstGoals
            .ControlName = "lstGoals"
            .Left = 20
            .Top = 145
            .Width = 250
            .Height = 90
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstGoals, UIControl))

        'lblGoalTime initial props
        lblGoalTime = New UILabel(oUILib)
        With lblGoalTime
            .ControlName = "lblGoalTime"
            .Left = 280
            .Top = 165
            .Width = 250
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Goal Time: Approx. 3 days"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblGoalTime, UIControl))

        'lblGoalRisk initial props
        lblGoalRisk = New UILabel(oUILib)
        With lblGoalRisk
            .ControlName = "lblGoalRisk"
            .Left = 280
            .Top = 190
            .Width = 250
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Risk Assessment: HIGH"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblGoalRisk, UIControl))

        'lblGoalPhase initial props
        lblGoalPhase = New UILabel(oUILib)
        With lblGoalPhase
            .ControlName = "lblGoalPhase"
            .Left = 280
            .Top = 215
            .Width = 250
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblGoalPhase, UIControl))

        'lblGoalName initial props
        lblGoalName = New UILabel(oUILib)
        With lblGoalName
            .ControlName = "lblGoalName"
            .Left = 280
            .Top = 140
            .Width = 250
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Goal:"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblGoalName, UIControl))

        'txtGoalDesc initial props
        txtGoalDesc = New UITextBox(oUILib)
        With txtGoalDesc
            .ControlName = "txtGoalDesc"
            .Left = 540
            .Top = 135
            .Width = 240
            .Height = 100
            .Enabled = True
            .Visible = True
            .Caption = "Goal Description...."
            .ForeColor = muSettings.InterfaceTextBoxForeColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(0, DrawTextFormat)
            .BackColorEnabled = muSettings.InterfaceTextBoxFillColor
            .BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
            .MaxLength = 0
            .BorderColor = muSettings.InterfaceBorderColor
            .MultiLine = True
            .Locked = True
        End With
        Me.AddChild(CType(txtGoalDesc, UIControl))

        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 1
            .Top = 27
            .Width = 799
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
        Me.AddChild(CType(lnDiv1, UIControl))

        'btnClose initial props
        btnClose = New UIButton(oUILib)
        With btnClose
            .ControlName = "btnClose"
            .Left = 774
            .Top = 2
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
            .Left = 750
            .Top = 2
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

        'btnAgentLeft initial props
        btnAgentLeft = New UIButton(oUILib)
        With btnAgentLeft
            .ControlName = "btnAgentLeft"
            .Left = 15
            .Top = fraGoals.Top + fraGoals.Height + 54      '(btnHeight \ 2) + (ImageHeight \ 2) + 15
            .Width = 21
            .Height = 50
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = grc_UI(elInterfaceRectangle.eAgentScrollLeftButton)
            .ControlImageRect_Disabled = .ControlImageRect
            .ControlImageRect_Normal = .ControlImageRect
            .ControlImageRect_Pressed = .ControlImageRect
        End With
        Me.AddChild(CType(btnAgentLeft, UIControl))

        'btnAgentRight initial props
        btnAgentRight = New UIButton(oUILib)
        With btnAgentRight
            .ControlName = "btnAgentRight"
            .Left = Me.Width - 21 - btnAgentLeft.Left
            .Top = btnAgentLeft.Top
            .Width = 21
            .Height = 50
            .Enabled = True
            .Visible = True
            .Caption = ""
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = grc_UI(elInterfaceRectangle.eAgentScrollRightButton)
            .ControlImageRect_Disabled = .ControlImageRect
            .ControlImageRect_Normal = .ControlImageRect
            .ControlImageRect_Pressed = .ControlImageRect
        End With
        Me.AddChild(CType(btnAgentRight, UIControl))

        'fraSkillSets initial props
        fraSkillSets = New fraSkillSet(oUILib)
        With fraSkillSets
            .ControlName = "fraSkillSets"
            .Left = fraGoals.Left
            .Top = fraGoals.Top + fraGoals.Height + 15 + 150
            '.Width = fraGoals.Width
            .Visible = True
            .Enabled = True
            .Caption = "Skillset Assignments"
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = False
            .BorderLineWidth = 2
            .bRoundedBorder = True
        End With
        Me.AddChild(CType(fraSkillSets, UIControl))

        'btnAssignCover initial props
        btnAssignCover = New UIButton(oUILib)
        With btnAssignCover
            .ControlName = "btnAssignCover"
            .Left = fraSkillSets.Left + fraSkillSets.Width + 5
            .Top = fraSkillSets.Top
            .Width = Me.Width - .Left - 5
            .Height = 28
            .Enabled = True
            .Visible = True
            .Caption = "Assign to Cover"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnAssignCover, UIControl))

        'lstCoverAgents initial props
        lstCoverAgents = New UIListBox(oUILib)
        With lstCoverAgents
            .ControlName = "lstCoverAgents"
            .Left = btnAssignCover.Left
            .Top = btnAssignCover.Top + btnAssignCover.Height + 5
            .Width = btnAssignCover.Width
            .Height = fraSkillSets.Height - (btnAssignCover.Height * 2) - 10
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstCoverAgents, UIControl))

        'btnRemoveCover initial props
        btnRemoveCover = New UIButton(oUILib)
        With btnRemoveCover
            .ControlName = "btnRemoveCover"
            .Left = btnAssignCover.Left
            .Top = lstCoverAgents.Top + lstCoverAgents.Height + 5
            .Width = btnAssignCover.Width
            .Height = btnAssignCover.Height
            .Enabled = True
            .Visible = True
            .Caption = "Remove Assignment"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnRemoveCover, UIControl))

        'btnSaveForLater initial props
        btnSaveForLater = New UIButton(oUILib)
        With btnSaveForLater
            .ControlName = "btnSaveForLater"
            .Left = fraSkillSets.Left + fraSkillSets.Width - 100
            .Top = fraSkillSets.Top + fraSkillSets.Height + 5
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Execute Later"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "The mission is planned but will not execute until Execute Now is pressed."
        End With
        Me.AddChild(CType(btnSaveForLater, UIControl))

        'btnExecuteNow initial props
        btnExecuteNow = New UIButton(oUILib)
        With btnExecuteNow
            .ControlName = "btnExecuteNow"
            .Left = btnSaveForLater.Left - 105
            .Top = btnSaveForLater.Top
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Execute Now"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "The mission will execute as soon as all agents are infiltrated and in place."
        End With
        Me.AddChild(CType(btnExecuteNow, UIControl))

        'btnCancel initial props
        btnCancel = New UIButton(oUILib)
        With btnCancel
            .ControlName = "btnCancel"
            .Left = fraSkillSets.Left
            .Top = btnSaveForLater.Top
            .Width = 130
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Cancel Mission"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            .ToolTipText = "Cancels the mission if the mission can be cancelled." & vbCrLf & _
                           "Once missions reach a certain phase of execution," & vbCrLf & _
                           "they can no longer be cancelled."
        End With
        Me.AddChild(CType(btnCancel, UIControl))

        'lblReqInfType initial props
        lblReqInfType = New UILabel(oUILib)
        With lblReqInfType
            .ControlName = "lblReqInfType"
            .Left = btnCancel.Left + btnCancel.Width + 10
            .Top = btnCancel.Top
            .Width = btnExecuteNow.Left - .Left
            .Height = btnCancel.Height
            .Enabled = True
            .Visible = True
            .Caption = "Required Infiltration: Select Mission"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
        End With
        Me.AddChild(CType(lblReqInfType, UIControl))

        'cboMethod initial props
        cboMethod = New UIComboBox(oUILib)
        With cboMethod
            .ControlName = "cboMethod"
            .Left = 80
            .Top = 85
            .Width = 220
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
        Me.AddChild(CType(cboMethod, UIControl))

        'cboTarget3 initial props
        cboTarget3 = New UIComboBox(oUILib)
        With cboTarget3
            .ControlName = "cboTarget3"
            .Left = 570
            .Top = 60
            .Width = 220
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
        Me.AddChild(CType(cboTarget3, UIControl))

        'cboTarget2 initial props
        cboTarget2 = New UIComboBox(oUILib)
        With cboTarget2
            .ControlName = "cboTarget2"
            .Left = 325
            .Top = 60
            .Width = 220
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
        Me.AddChild(CType(cboTarget2, UIControl))

        'cboTarget1 initial props
        cboTarget1 = New UIComboBox(oUILib)
        With cboTarget1
            .ControlName = "cboTarget1"
            .Left = 80
            .Top = 60
            .Width = 220
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
        Me.AddChild(CType(cboTarget1, UIControl))

        'cboMission initial props
        cboMission = New UIComboBox(oUILib)
        With cboMission
            .ControlName = "cboMission"
            .Left = 80
            .Top = 35
            .Width = 220
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
        Me.AddChild(CType(cboMission, UIControl))

        MyBase.moUILib.RemoveWindow(Me.ControlName)
        MyBase.moUILib.AddWindow(Me)

        FillMissionList()
        mbLoading = False
    End Sub

    Public Sub SetFromMission(ByRef oMission As PlayerMission, ByVal bCreateNew As Boolean)
        If oMission Is Nothing Then
            moMission = New PlayerMission()
            moMission.PM_ID = -1
        Else
            moMission = oMission.CreateClone()           'now, moMission is a workable copy
            moMission.PM_ID = -1
        End If
        If moMission.oMission Is Nothing = False Then cboMission.FindComboItemData(moMission.oMission.ObjectID)
        cboMission_ItemSelected(cboMission.ListIndex)
        cboTarget1.FindComboItemData(moMission.lTargetPlayerID)
        cboTarget1_ItemSelected(cboTarget1.ListIndex)

        If cboTarget2.Visible = True Then
            cboTarget2.ListIndex = -1
            For X As Int32 = 0 To cboTarget2.ListCount - 1
                If cboTarget2.ItemData(X) = moMission.lTargetID AndAlso cboTarget2.ItemData2(X) = moMission.iTargetTypeID Then
                    cboTarget2.ListIndex = X
                    Exit For
                End If
            Next X
        End If
        If cboTarget3.Visible = True Then
            cboTarget3.ListIndex = -1
            For X As Int32 = 0 To cboTarget3.ListCount - 1
                If cboTarget3.ItemData(X) = moMission.lTargetID2 AndAlso cboTarget3.ItemData2(X) = moMission.iTargetTypeID2 Then
                    cboTarget3.ListIndex = X
                    Exit For
                End If
            Next X
        End If

        cboMethod.FindComboItemData(moMission.lMethodID)
        cboMethod_ItemSelected(cboMethod.ListIndex)

        cboMission.Enabled = moMission.PM_ID < 1

    End Sub

    Private Sub FillMissionList()
        Dim lSorted() As Int32 = GetSortedIndexArray(goMissions, glMissionIdx, GetSortedIndexArrayType.eMissionName)

        'Fill our mission list
        cboMission.Clear()
        If lSorted Is Nothing = False Then
            For X As Int32 = 0 To lSorted.GetUpperBound(0)

                If NewTutorialManager.TutorialOn = True AndAlso goCurrentPlayer.yPlayerPhase = 0 Then
                    If glMissionIdx(lSorted(X)) = 999 Then
                        cboMission.AddItem(goMissions(lSorted(X)).sMissionName)
                        cboMission.ItemData(cboMission.NewIndex) = goMissions(lSorted(X)).ObjectID
                    End If
                Else
                    If glMissionIdx(lSorted(X)) <> -1 AndAlso glMissionIdx(lSorted(X)) <> 999 Then
                        If goMissions(lSorted(X)).ProgramControlID = eMissionResult.eLocateWormhole Then
                            If goCurrentPlayer Is Nothing OrElse (goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eSuperSpecials) And 4) = 0 Then
                                Continue For
                            End If
                        End If
                        cboMission.AddItem(goMissions(lSorted(X)).sMissionName)
                        cboMission.ItemData(cboMission.NewIndex) = goMissions(lSorted(X)).ObjectID
                    End If
                End If
            Next X
        End If
    End Sub

    Private Sub frmMission_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
        If mbMouseOver Is Nothing = False Then
            For X As Int32 = 0 To mbMouseOver.GetUpperBound(0)
                If mbMouseOver(X) = True Then
                    'ok, select that item
                    'If mlScrollValue + X > goCurrentPlayer.AgentUB Then Return
                    Dim lIdx As Int32 = GetScrolledSelectionIndex(X)
                    If lIdx > -1 AndAlso lIdx <= goCurrentPlayer.AgentUB Then AgentSelected(goCurrentPlayer.Agents(lIdx))
                    Exit For
                End If
            Next X
        End If
    End Sub

    Private Function GetScrolledSelectionIndex(ByVal lSelectionIndex As Int32) As Int32
        Dim lTempIdx As Int32 = -1

        Dim lAgentListUB As Int32 = -1
        Dim lAgentListIdx() As Int32 = Nothing
        Dim lCnt As Int32 = 0
        Dim lStartIdx As Int32 = -1
        Dim lSorted() As Int32 = GetSortedIndexArray(goCurrentPlayer.Agents, goCurrentPlayer.AgentIdx, GetSortedIndexArrayType.eAgentName)
        If lSorted Is Nothing = False Then
            For X As Int32 = 0 To lSorted.GetUpperBound(0)
                lTempIdx = lSorted(X)
                If goCurrentPlayer.AgentIdx(lTempIdx) <> -1 AndAlso (goCurrentPlayer.Agents(lTempIdx).lAgentStatus And (AgentStatus.NewRecruit Or AgentStatus.IsDead Or AgentStatus.Dismissed Or AgentStatus.HasBeenCaptured)) = 0 Then
                    Dim oAgent As Agent = goCurrentPlayer.Agents(lTempIdx)
                    If oAgent Is Nothing = False Then
                        Dim bExclude As Boolean = False
                        If moFilterSkillSet Is Nothing = False Then
                            bExclude = True
                            For Y As Int32 = 0 To moFilterSkillSet.SkillUB
                                Dim bSkillIsAbility As Boolean = moFilterSkillSet.Skills(Y).oSkill.SkillType <> 0
                                For Z As Int32 = 0 To oAgent.SkillUB
                                    If oAgent.Skills(Z).ObjectID = moFilterSkillSet.Skills(Y).oSkill.ObjectID Then
                                        bExclude = False
                                        Exit For
                                    ElseIf oAgent.Skills(Z).ObjectID = 36 AndAlso bSkillIsAbility = False Then          '36 is Naturally Talented ID
                                        bExclude = False
                                        Exit For
                                    End If
                                Next Z
                                If bExclude = False Then Exit For
                            Next Y
                        End If

                        If bExclude = False Then
                            lAgentListUB += 1
                            ReDim Preserve lAgentListIdx(lAgentListUB)
                            lAgentListIdx(lAgentListUB) = lTempIdx

                            mlScrollMaxValue += 1
                            lCnt += 1
                            If lCnt = mlScrollValue + 1 AndAlso lStartIdx = -1 Then
                                lStartIdx = lAgentListUB
                            End If
                        End If
                    End If
                End If
            Next X
        End If
        lTempIdx = -1

        For X As Int32 = 0 To lAgentListUB
            Dim oAgent As Agent = goCurrentPlayer.Agents(lAgentListIdx(X))
            If oAgent Is Nothing = False Then
                lTempIdx += 1
                If lTempIdx = mlScrollValue + lSelectionIndex Then Return lAgentListIdx(X)
            End If
        Next X
        Return -1
    End Function

    Private Sub frmMission_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        Dim oLoc As Point = Me.GetAbsolutePosition()
        lMouseY -= oLoc.Y
        lMouseX -= oLoc.X
        Dim bMouseOverAgent As Boolean = False
        If lMouseX > 75 Then ' 50 Then
            If lMouseY > fraGoals.Top + fraGoals.Height + 15 AndAlso lMouseY < fraGoals.Top + fraGoals.Height + 164 Then

                'rcDest.X += 130
                lMouseX -= 75 '50
                Dim lItem As Int32 = lMouseX \ 130
                If lItem < mlAgentBarCnt Then
                    For X As Int32 = 0 To mlAgentBarCnt - 1
                        mbMouseOver(X) = False
                    Next X

                    bMouseOverAgent = True

                    If mlMouseOverEvent = -1 Then
                        mlMouseOverEvent = glCurrentCycle
                    ElseIf moAgentPop Is Nothing = False Then
                        Dim lIdx As Int32 = GetScrolledSelectionIndex(lItem)
                        If lIdx > -1 AndAlso lIdx <= goCurrentPlayer.AgentUB Then moAgentPop.SetFromAgent(goCurrentPlayer.Agents(lIdx))
                        'If mlScrollValue + lItem < goCurrentPlayer.AgentUB Then moAgentPop.SetFromAgent(goCurrentPlayer.Agents(mlScrollValue + lItem))
                    End If
                    mbMouseOver(lItem) = True
                    Me.IsDirty = True
                End If
            End If
        End If

        If bMouseOverAgent = False Then
            For X As Int32 = 0 To mlAgentBarCnt - 1
                If mbMouseOver(X) = True Then
                    mbMouseOver(X) = False
                    Me.IsDirty = True
                    mlMouseOverEvent = -1
                End If
            Next X
        End If
    End Sub

    Private Sub AgentSelected(ByRef oAgent As Agent)
        mlSelectedAgentID = -1
        If oAgent Is Nothing = False Then
            mlSelectedAgentID = oAgent.ObjectID
            If NewTutorialManager.TutorialOn = True Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionSelectAgent, oAgent.ObjectID, -1, -1, "")
            End If
            Me.IsDirty = True
        End If
    End Sub

    Private Sub frmMission_OnNewFrame() Handles Me.OnNewFrame

        If cboTarget1 Is Nothing = False Then
            With cboTarget1
                For X As Int32 = 0 To .ListCount - 1
                    If .ItemData2(X) = ObjectType.ePlayer Then
                        Dim sText As String = GetCacheObjectValue(.ItemData(X), ObjectType.ePlayer)
                        If .List(X) <> sText Then .List(X) = sText
                    ElseIf .ItemData2(X) = ObjectType.eAgent Then
                        If .List(X).ToUpper = "UNKNOWN" Then
                            Dim sText As String = GetCacheObjectValue(.ItemData(X), ObjectType.eAgent)
                            If .List(X) <> sText Then .List(X) = sText
                        End If
                    End If
                Next X
            End With
        End If
        If cboTarget2 Is Nothing = False Then
            With cboTarget2
                For X As Int32 = 0 To .ListCount - 1
                    If .ItemData2(X) = ObjectType.ePlayer Then
                        Dim sText As String = GetCacheObjectValue(.ItemData(X), ObjectType.ePlayer)
                        If .List(X) <> sText Then .List(X) = sText
                    ElseIf .ItemData2(X) = ObjectType.eAgent Then
                        If .List(X).ToUpper = "UNKNOWN" Then
                            Dim sText As String = GetCacheObjectValue(.ItemData(X), ObjectType.eAgent)
                            If .List(X) <> sText Then .List(X) = sText
                        End If
                    End If
                Next X
            End With
        End If
        If cboTarget3 Is Nothing = False Then
            With cboTarget3
                For X As Int32 = 0 To .ListCount - 1
                    If .ItemData2(X) = ObjectType.ePlayer Then
                        Dim sText As String = GetCacheObjectValue(.ItemData(X), ObjectType.ePlayer)
                        If .List(X) <> sText Then .List(X) = sText
                    ElseIf .ItemData2(X) = ObjectType.eAgent Then
                        If .List(X).ToUpper = "UNKNOWN" Then
                            Dim sText As String = GetCacheObjectValue(.ItemData(X), ObjectType.eAgent)
                            If .List(X) <> sText Then .List(X) = sText
                        End If
                    End If
                Next X
            End With
        End If

        If mlMouseOverEvent = -1 Then
            If moAgentPop Is Nothing = False Then
                RemoveHandler moAgentPop.AgentSelected, AddressOf AgentSelected
                MyBase.moUILib.RemoveWindow(moAgentPop.ControlName)
                moAgentPop = Nothing
            End If
        ElseIf glCurrentCycle - mlMouseOverEvent > ml_MOUSE_OVER_DELAY Then
            'figure out which item we are hovered over
            Dim lIdx As Int32 = -1
            Dim lAgentIdx As Int32 = -1
            For X As Int32 = 0 To mlAgentBarCnt - 1
                If mbMouseOver(X) = True Then
                    lIdx = X
                    Exit For
                End If
            Next X
            If lIdx = -1 Then
                mlMouseOverEvent = -1
                Return
            Else
                lAgentIdx = GetScrolledSelectionIndex(lIdx)
                If lAgentIdx < 0 OrElse lAgentIdx > goCurrentPlayer.AgentUB Then
                    mlMouseOverEvent = -1
                    Return
                End If
            End If

            'ok, is moAgentPop nothing?
            If moAgentPop Is Nothing = True Then
                moAgentPop = New frmAgent(goUILib, True, False)
                AddHandler moAgentPop.AgentSelected, AddressOf AgentSelected
            End If
            Dim lTemp As Int32 = 29 + ((lIdx) * 130) + Me.Left  '4
            If moAgentPop.Left <> lTemp Then moAgentPop.Left = lTemp
            lTemp = fraGoals.Top + fraGoals.Height + Me.Top + 15 - 249
            If moAgentPop.Top <> lTemp Then moAgentPop.Top = lTemp
            If moAgentPop.Visible = False Then moAgentPop.Visible = True
            moAgentPop.SetAgentIfNeeded(goCurrentPlayer.Agents(lAgentIdx))
        End If

        If mbHasUnknowns = True Then Me.IsDirty = True
    End Sub

    Private Sub frmMission_OnRender() Handles Me.OnRender
        mlAgentBarCnt = (Me.Width - 100) \ 130
        If mbMouseOver Is Nothing OrElse mbMouseOver.GetUpperBound(0) + 1 <> mlAgentBarCnt Then ReDim mbMouseOver(mlAgentBarCnt - 1)

        If goCurrentPlayer Is Nothing = True Then Return
        If goCurrentPlayer.AgentUB = -1 Then Return
        If fraGoals Is Nothing Then Return

        'get our total scroll count
        Dim lStartIdx As Int32 = -1
        Dim lCnt As Int32 = 0

        Dim lAgentListIdx() As Int32 = Nothing
        Dim lAgentListUB As Int32 = -1
        mlScrollMaxValue = 0

        Dim lSorted() As Int32 = GetSortedIndexArray(goCurrentPlayer.Agents, goCurrentPlayer.AgentIdx, GetSortedIndexArrayType.eAgentName)
        If lSorted Is Nothing = False Then
            For X As Int32 = 0 To lSorted.GetUpperBound(0)
                Dim lTempIdx As Int32 = lSorted(X)
                If goCurrentPlayer.AgentIdx(lTempIdx) <> -1 Then
                    Dim oAgent As Agent = goCurrentPlayer.Agents(lTempIdx)
                    If oAgent Is Nothing = False Then
                        If (oAgent.lAgentStatus And (AgentStatus.NewRecruit Or AgentStatus.IsDead Or AgentStatus.Dismissed Or AgentStatus.HasBeenCaptured)) <> 0 Then Continue For

                        Dim bExclude As Boolean = False
                        If moFilterSkillSet Is Nothing = False Then
                            bExclude = True
                            For Y As Int32 = 0 To moFilterSkillSet.SkillUB
                                Dim bSkillIsAbility As Boolean = moFilterSkillSet.Skills(Y).oSkill.SkillType <> 0
                                For Z As Int32 = 0 To oAgent.SkillUB
                                    If oAgent.Skills(Z).ObjectID = moFilterSkillSet.Skills(Y).oSkill.ObjectID Then
                                        bExclude = False
                                        Exit For
                                    ElseIf oAgent.Skills(Z).ObjectID = 36 AndAlso bSkillIsAbility = False Then      'naturally talented
                                        bExclude = False
                                        Exit For
                                    End If
                                Next Z
                                If bExclude = False Then Exit For
                            Next Y
                        End If

                        If bExclude = False Then
                            lAgentListUB += 1
                            ReDim Preserve lAgentListIdx(lAgentListUB)
                            lAgentListIdx(lAgentListUB) = lTempIdx

                            mlScrollMaxValue += 1
                            lCnt += 1
                            If lCnt = mlScrollValue + 1 AndAlso lStartIdx = -1 Then
                                lStartIdx = lAgentListUB
                            End If
                        End If
                    End If
                End If
            Next X
        End If

        'mlScrollMaxValue = 0
        'For X As Int32 = 0 To goCurrentPlayer.AgentUB
        '	If goCurrentPlayer.AgentIdx(X) <> -1 AndAlso (goCurrentPlayer.Agents(X).lAgentStatus And (AgentStatus.NewRecruit Or AgentStatus.IsDead Or AgentStatus.Dismissed)) = 0 Then
        '		mlScrollMaxValue += 1
        '		lCnt += 1
        '		If lCnt = mlScrollValue + 1 AndAlso lStartIdx = -1 Then
        '			lStartIdx = X
        '		End If
        '	End If
        'Next X
        mlScrollMaxValue -= mlAgentBarCnt
        If mlScrollMaxValue < 0 Then mlScrollMaxValue = 0
        If lStartIdx = -1 Then lStartIdx = 0
        If mlScrollValue > mlScrollMaxValue Then mlScrollValue = mlScrollMaxValue

        If AgentRenderer.goAgentRenderer Is Nothing Then AgentRenderer.goAgentRenderer = New AgentRenderer()

        AgentRenderer.goAgentRenderer.PrepareForBatch()
        Dim rcDest As Rectangle
        With rcDest
            .X = 75 '50
            .Y = fraGoals.Top + fraGoals.Height + 15 '+ Me.Top
            .Width = 128
            .Height = 128
        End With

        lCnt = 0
        Dim lSelectedIndex As Int32 = -1
        If mlSelectedAgentID <> -1 Then
            lCnt = 0
            Dim rcSelectBox As Rectangle = rcDest
            For X As Int32 = lStartIdx To lAgentListUB ' goCurrentPlayer.AgentUB
                Dim oAgent As Agent = goCurrentPlayer.Agents(lAgentListIdx(X))
                If oAgent Is Nothing = False AndAlso (oAgent.lAgentStatus And (AgentStatus.NewRecruit Or AgentStatus.IsDead Or AgentStatus.Dismissed)) = 0 Then
                    If oAgent.ObjectID = mlSelectedAgentID Then
                        'Ok, found it
                        rcSelectBox.X -= 2
                        rcSelectBox.Y -= 2
                        rcSelectBox.Width += 4
                        rcSelectBox.Height += 4
                        lSelectedIndex = X
                        RenderBox(rcSelectBox, 4, Color.White)
                        Exit For
                    End If
                    rcSelectBox.X += 130
                    lCnt += 1
                    If lCnt = mlAgentBarCnt Then Exit For
                End If
            Next X
        End If

        lCnt = 0
        For X As Int32 = lStartIdx To lAgentListUB 'goCurrentPlayer.AgentUB
            Dim oAgent As Agent = goCurrentPlayer.Agents(lAgentListIdx(X))
            If oAgent Is Nothing = False AndAlso (oAgent.lAgentStatus And (AgentStatus.NewRecruit Or AgentStatus.IsDead Or AgentStatus.Dismissed)) = 0 Then
                If lSelectedIndex = X Then
                    AgentRenderer.goAgentRenderer.RenderAgent2(oAgent.ObjectID, rcDest, False, oAgent.bMale)
                Else
                    AgentRenderer.goAgentRenderer.RenderAgent2(oAgent.ObjectID, rcDest, Not mbMouseOver(lCnt), oAgent.bMale)
                End If
                rcDest.X += 130
                lCnt += 1
                If lCnt = mlAgentBarCnt Then Exit For
            End If
        Next X

        mbHasUnknowns = AgentRenderer.goAgentRenderer.bHasUnknowns
        AgentRenderer.goAgentRenderer.FinishBatch()

        rcDest.X = 75
        rcDest.Y = fraGoals.Top + fraGoals.Height + 15 + 130
        rcDest.Width = 128
        rcDest.Height = 18

        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return

        Using oFont As Font = New Font(MyBase.moUILib.oDevice, lblMission.GetFont())
            Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
                oTextSpr.Begin(SpriteFlags.AlphaBlend)

                Try
                    lCnt = 0
                    For X As Int32 = lStartIdx To lAgentListUB 'goCurrentPlayer.AgentUB
                        Dim oAgent As Agent = goCurrentPlayer.Agents(lAgentListIdx(X))
                        If oAgent Is Nothing = False AndAlso (oAgent.lAgentStatus And (AgentStatus.NewRecruit Or AgentStatus.IsDead Or AgentStatus.Dismissed)) = 0 Then
                            oFont.DrawText(oTextSpr, oAgent.sAgentName, rcDest, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, Color.White)
                            rcDest.X += 130
                            lCnt += 1
                            If lCnt = mlAgentBarCnt Then Exit For
                        End If
                    Next X
                Catch
                End Try

                oTextSpr.End()
                oTextSpr.Dispose()
            End Using
            oFont.Dispose()
        End Using

    End Sub

    Protected Overrides Sub Finalize()
        AgentRenderer.goAgentRenderer = Nothing
        MyBase.Finalize()
    End Sub

    Private Sub cboMission_ItemSelected(ByVal lItemIndex As Integer) Handles cboMission.ItemSelected
        lblMissionDesc.Caption = ""
        cboTarget2.ListIndex = -1 : cboTarget2.Enabled = False
        cboTarget3.ListIndex = -1 : cboTarget3.Enabled = False
        cboTarget1.ListIndex = -1 : cboTarget1.Clear()
        cboTarget1.Visible = True : cboTarget1.Left = 80
        cboTarget2.Left = 325 : cboTarget3.Left = 570

        cboMethod.ListIndex = -1 : cboMethod.Clear()
        lstGoals.Clear()
        lblGoalName.Caption = ""
        lblGoalTime.Caption = ""
        lblGoalRisk.Caption = ""
        txtGoalDesc.Caption = ""
        lblGoalPhase.Caption = ""

        If cboMission.ListIndex <> -1 Then
            Dim oMission As Mission = Nothing
            Dim lID As Int32 = cboMission.ItemData(cboMission.ListIndex)
            For X As Int32 = 0 To glMissionUB
                If glMissionIdx(X) = lID Then
                    oMission = goMissions(X)
                    Exit For
                End If
            Next X
            If oMission Is Nothing = False Then

                If NewTutorialManager.TutorialOn = True Then
                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionMissionSelected, lID, -1, -1, "")
                End If

                If moMission.oMission Is Nothing = True OrElse moMission.oMission.ObjectID <> oMission.ObjectID Then
                    moMission.oMission = oMission

                    ReDim moMission.oMissionGoals(oMission.GoalUB)
                    For X As Int32 = 0 To oMission.GoalUB
                        moMission.oMissionGoals(X) = New PlayerMissionGoal()
                        moMission.oMissionGoals(X).oGoal = oMission.Goals(X)
                        moMission.oMissionGoals(X).oMission = moMission
                    Next X
                End If

                Dim sMDesc As String = oMission.sMissionDesc
                Dim lVBCRLFLoc As Int32 = sMDesc.IndexOf(vbCrLf)
                If lVBCRLFLoc > 0 Then
                    lVBCRLFLoc = Math.Min(lVBCRLFLoc, 82)
                Else : lVBCRLFLoc = Math.Min(82, sMDesc.Length)
                End If
                If lVBCRLFLoc <> sMDesc.Length Then
                    sMDesc = sMDesc.Substring(0, lVBCRLFLoc) & "..."
                End If
                lblMissionDesc.Caption = sMDesc
                lblMissionDesc.ToolTipText = oMission.sMissionDesc & vbCrLf & "Required Infiltration Type: " & frmAgent.fraInfiltration.GetInfTypeText(oMission.lInfiltrationType)

                lblReqInfType.Caption = "Required Infiltration: " & frmAgent.fraInfiltration.GetInfTypeText(oMission.lInfiltrationType)

                'ok, determine what boxes to highlight (enable). for mission results 32, 39, 40 and 41 the first box is unused
                myTarget1ListType = eTargetFilterType.ePlayer   'default to player
                myTarget2ListType = eTargetFilterType.eUnusedTarget
                myTarget3ListType = eTargetFilterType.eUnusedTarget

                'Now, determine the remaining comboboxes
                Select Case CType(oMission.ProgramControlID, eMissionResult)
                    Case eMissionResult.eAcquireTechData
                        myTarget2ListType = eTargetFilterType.eTarget_Component Or eTargetFilterType.eAny
                    Case eMissionResult.eAssassinateAgent, eMissionResult.eCaptureAgent
                        myTarget2ListType = eTargetFilterType.eTarget_Agent Or eTargetFilterType.eAny
                    Case eMissionResult.eClutterCargoBay, eMissionResult.eIncreasePowerNeeds
                        myTarget2ListType = eTargetFilterType.eTarget_Colony
                    Case eMissionResult.eDoorJamHangar
                        myTarget2ListType = eTargetFilterType.eTarget_UnitOrFacility Or eTargetFilterType.eAny
                    Case eMissionResult.eFindMineral
                        'make target1 invisible
                        cboTarget1.Visible = False
                        'shift target2 and 3 to the elft
                        cboTarget2.Left = 80 : cboTarget3.Left = 325
                        myTarget1ListType = eTargetFilterType.eUnusedTarget
                        myTarget2ListType = eTargetFilterType.eMineral
                        myTarget3ListType = eTargetFilterType.eSystemAndPlanets ' eTargetFilterType.eSolarSystem
                    Case eMissionResult.eAudit
                        cboTarget1.Visible = False
                        myTarget1ListType = eTargetFilterType.eUnusedTarget
                        myTarget2ListType = eTargetFilterType.eUnusedTarget
                    Case eMissionResult.eGeologicalSurvey, eMissionResult.eTutorialFindFactory
                        'make target1 invisible
                        cboTarget1.Visible = False
                        'shift target2 and 3 to the elft
                        cboTarget2.Left = 80 : cboTarget3.Left = 325
                        myTarget1ListType = eTargetFilterType.eUnusedTarget
                        myTarget2ListType = eTargetFilterType.ePlanet
                    Case eMissionResult.eImpedeCurrentDevelopment
                        'myTarget2ListType = eTargetFilterType.eTargetSpecialTechUnresearched
                    Case eMissionResult.eDestroyCurrentSpecialProject
                        myTarget2ListType = eTargetFilterType.eTargetSpecialTechUnresearched
                    Case eMissionResult.eGetFacilityList
                        myTarget2ListType = eTargetFilterType.eProductionType
                        myTarget3ListType = eTargetFilterType.eSystemAndPlanets
                    Case eMissionResult.ePlantEvidence
                        myTarget2ListType = eTargetFilterType.ePlayer
                        myTarget3ListType = eTargetFilterType.ePlayerMissions
                    Case eMissionResult.eReconPlanetMap
                        'make target1 invisible
                        cboTarget1.Visible = False
                        'shift target2 and 3 to the elft
                        cboTarget2.Left = 80 : cboTarget3.Left = 325
                        myTarget1ListType = eTargetFilterType.eUnusedTarget
                        myTarget2ListType = eTargetFilterType.ePlanet
                        myTarget3ListType = eTargetFilterType.ePICK_A_LOCATION
                    Case eMissionResult.eSearchAndRescueAgent
                        'make target1 invisible
                        cboTarget1.Visible = False
                        'shift target2 and 3 to the elft
                        cboTarget2.Left = 80 : cboTarget3.Left = 325
                        myTarget1ListType = eTargetFilterType.eUnusedTarget
                        myTarget2ListType = eTargetFilterType.eCaptured_Agent
                        myTarget3ListType = eTargetFilterType.eUnusedTarget
                    Case eMissionResult.eShowQueueAndContents, eMissionResult.eSlowProduction, eMissionResult.eSabotageProduction, _
                      eMissionResult.eCorruptProduction, eMissionResult.eDestroyFacility
                        myTarget2ListType = eTargetFilterType.eTarget_Facility Or eTargetFilterType.eAny
                    Case eMissionResult.eStealCargo
                        myTarget2ListType = eTargetFilterType.eTarget_Colony Or eTargetFilterType.eAny
                        myTarget3ListType = eTargetFilterType.eTarget_ComponentOrKnownMineral
                    Case eMissionResult.eAssassinateGovernor, eMissionResult.eDecreaseHousingMorale, eMissionResult.eDecreaseTaxMorale, eMissionResult.eDecreaseUnemploymentMorale
                        myTarget2ListType = eTargetFilterType.eTarget_Colony Or eTargetFilterType.eAny
                    Case eMissionResult.eGetColonyBudget
                        myTarget2ListType = eTargetFilterType.ePlanet

                        'These do not use Target 2 or 3 combo boxes
                        'case eMissionResult.eAgencyAnalysis
                        'Case eMissionResult.eBadPublicity
                        'Case eMissionResult.eComponentDesignList
                        'Case eMissionResult.eComponentResearchList
                        'Case eMissionResult.eDegradePay
                        'Case eMissionResult.eDiplomaticStrength
                        'Case eMissionResult.eDropInvulField
                        'Case eMissionResult.eFindCommandCenters
                        'Case eMissionResult.eFindProductionFac
                        'Case eMissionResult.eFindResearchFac
                        'Case eMissionResult.eFindSpaceStations
                        'case eMissionResult.eGetAgentList
                        'case eMissionResult.eGetAliasList
                        'Case eMissionResult.eGetAlliesList
                        'Case eMissionResult.eGetEnemyList
                        'Case eMissionResult.eGetMineralList
                        'Case eMissionResult.eGetTradeTreaties
                        'Case eMissionResult.eIFFSabotage
                        'case eMissionResult.eLocateWormhole
                        'Case eMissionResult.eMilitaryCoup
                        'Case eMissionResult.eMilitaryStrength
                        'Case eMissionResult.ePopulationScore
                        'Case eMissionResult.eProductionScore
                        'Case eMissionResult.eRelayBattleReports
                        'Case eMissionResult.eSiphonItem
                        'Case eMissionResult.eSpecialProjectList
                        'Case eMissionResult.eTechnologyStrength
                End Select

                'Now, based on our filters, fill our combo boxes
                If myTarget1ListType <> eTargetFilterType.eUnusedTarget Then
                    FillTargetCBOFromFilter(myTarget1ListType, cboTarget1, -1, -1, -1, -1)
                ElseIf myTarget2ListType <> eTargetFilterType.eUnusedTarget Then
                    FillTargetCBOFromFilter(myTarget2ListType, cboTarget2, -1, -1, -1, -1)
                End If

                'Now, fill our methods
                For X As Int32 = 0 To oMission.GoalUB
                    Dim bFound As Boolean = False
                    For Y As Int32 = 0 To cboMethod.ListCount - 1
                        If cboMethod.ItemData(Y) = oMission.MethodIDs(X) Then
                            bFound = True
                            Exit For
                        End If
                    Next Y
                    If bFound = False Then
                        Dim sName As String = ""
                        For Y As Int32 = 0 To glMissionMethodUB
                            If glMissionMethodIdx(Y) = oMission.MethodIDs(X) Then
                                sName = gsMissionMethods(Y)
                                Exit For
                            End If
                        Next Y
                        If sName <> "" Then
                            cboMethod.AddItem(sName)
                            cboMethod.ItemData(cboMethod.NewIndex) = oMission.MethodIDs(X)
                        End If
                    End If
                Next X

            End If
        End If
    End Sub

    Private Sub FillTargetCBOFromFilter(ByVal yFilter As eTargetFilterType, ByRef cboData As UIComboBox, ByVal lParentID As Int32, ByVal iParentTypeID As Int16, ByVal lGrandParentID As Int32, ByVal iGrandParentTypeID As Int16)
        cboData.Clear()
        cboData.Visible = True
        cboData.Enabled = True

        Dim bIncludeAny As Boolean = (yFilter And eTargetFilterType.eAny) <> 0

        Dim lAnyIndex As Int32 = -1
        If bIncludeAny = True Then
            yFilter = yFilter Xor eTargetFilterType.eAny
            cboData.AddItem("ANY")
            lAnyIndex = cboData.NewIndex
            cboData.ItemData(lAnyIndex) = Int32.MinValue
        End If

        'Now, determine what to fill
        Select Case yFilter
            Case eTargetFilterType.eCaptured_Agent
                If lAnyIndex <> -1 Then cboData.ItemData2(lAnyIndex) = ObjectType.eAgent
                For X As Int32 = 0 To goCurrentPlayer.AgentUB
                    If goCurrentPlayer.AgentIdx(X) <> -1 AndAlso (goCurrentPlayer.Agents(X).lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 AndAlso (lParentID = -1 OrElse iParentTypeID <> ObjectType.ePlayer OrElse goCurrentPlayer.Agents(X).lTargetID = lParentID) Then
                        cboData.AddItem(goCurrentPlayer.Agents(X).sAgentName)
                        cboData.ItemData(cboData.NewIndex) = goCurrentPlayer.Agents(X).ObjectID
                        cboData.ItemData2(cboData.NewIndex) = ObjectType.eAgent
                    End If
                Next X
            Case eTargetFilterType.eMineral
                If lAnyIndex <> -1 Then cboData.ItemData2(lAnyIndex) = ObjectType.eMineral
                Dim lSorted() As Int32 = GetSortedMineralIdxArray(False, False)
                If lSorted Is Nothing = False Then
                    For X As Int32 = 0 To lSorted.GetUpperBound(0)
                        If glMineralIdx(lSorted(X)) <> -1 Then
                            cboData.AddItem(goMinerals(lSorted(X)).MineralName)
                            cboData.ItemData(cboData.NewIndex) = goMinerals(lSorted(X)).ObjectID
                            cboData.ItemData2(cboData.NewIndex) = ObjectType.eMineral
                        End If
                    Next X
                End If
            Case eTargetFilterType.ePICK_A_LOCATION
                If lAnyIndex <> -1 Then cboData.ItemData2(lAnyIndex) = Int32.MinValue
                cboData.AddItem("Pick A Location...")
            Case eTargetFilterType.ePlayer
                If lAnyIndex <> -1 Then cboData.ItemData2(lAnyIndex) = ObjectType.ePlayer

                'Ok, sort our rels
                Dim lIdx() As Int32 = GetSortedPlayerRelIdxArray()
                If lIdx Is Nothing = False Then
                    For X As Int32 = 0 To lIdx.GetUpperBound(0)
                        Dim oTmpRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(lIdx(X))
                        If oTmpRel Is Nothing = False Then
                            cboData.AddItem(GetCacheObjectValue(oTmpRel.lThisPlayer, ObjectType.ePlayer))
                            cboData.ItemData(cboData.NewIndex) = oTmpRel.lThisPlayer
                            cboData.ItemData2(cboData.NewIndex) = ObjectType.ePlayer
                        End If
                    Next X
                End If
                'For X As Int32 = 0 To goCurrentPlayer.PlayerRelUB
                '    Dim oTmpRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(X)
                '    If oTmpRel Is Nothing = False Then
                '        cboData.AddItem(GetCacheObjectValue(oTmpRel.lThisPlayer, ObjectType.ePlayer))
                '        cboData.ItemData(cboData.NewIndex) = oTmpRel.lThisPlayer
                '        cboData.ItemData2(cboData.NewIndex) = ObjectType.ePlayer
                '    End If
                'Next X
            Case eTargetFilterType.ePlayerMissions
                If lAnyIndex <> -1 Then cboData.ItemData2(lAnyIndex) = -1
                For X As Int32 = 0 To goCurrentPlayer.PlayerMissionUB
                    If goCurrentPlayer.PlayerMissionIdx(X) <> -1 AndAlso (lParentID = -1 OrElse iParentTypeID <> ObjectType.ePlayer OrElse (goCurrentPlayer.PlayerMissions(X).lTargetPlayerID = lParentID)) Then
                        'TODO: ???
                    End If
                Next X
            Case eTargetFilterType.eProductionType
                If lAnyIndex <> -1 Then cboData.ItemData2(lAnyIndex) = -1
                cboData.AddItem("Barracks") : cboData.ItemData(cboData.NewIndex) = ProductionType.eEnlisted
                cboData.AddItem("Command Center") : cboData.ItemData(cboData.NewIndex) = ProductionType.eCommandCenterSpecial
                cboData.AddItem("Factory") : cboData.ItemData(cboData.NewIndex) = ProductionType.eProduction
                cboData.AddItem("Land Production") : cboData.ItemData(cboData.NewIndex) = ProductionType.eProduction
                cboData.AddItem("Mining") : cboData.ItemData(cboData.NewIndex) = ProductionType.eMining
                cboData.AddItem("Naval Production") : cboData.ItemData(cboData.NewIndex) = ProductionType.eNavalProduction
                cboData.AddItem("Officers") : cboData.ItemData(cboData.NewIndex) = ProductionType.eOfficers
                cboData.AddItem("Power Center") : cboData.ItemData(cboData.NewIndex) = ProductionType.ePowerCenter
                cboData.AddItem("Refining") : cboData.ItemData(cboData.NewIndex) = ProductionType.eRefining
                cboData.AddItem("Research") : cboData.ItemData(cboData.NewIndex) = ProductionType.eResearch
                cboData.AddItem("Residence") : cboData.ItemData(cboData.NewIndex) = ProductionType.eColonists
                cboData.AddItem("Spaceport") : cboData.ItemData(cboData.NewIndex) = ProductionType.eAerialProduction
                cboData.AddItem("Space Station") : cboData.ItemData(cboData.NewIndex) = ProductionType.eSpaceStationSpecial
                cboData.AddItem("Tradepost") : cboData.ItemData(cboData.NewIndex) = ProductionType.eTradePost
                cboData.AddItem("Warehouse") : cboData.ItemData(cboData.NewIndex) = ProductionType.eWareHouse
            Case eTargetFilterType.eTarget_Agent
                If lAnyIndex <> -1 Then cboData.ItemData2(lAnyIndex) = ObjectType.eAgent
                For X As Int32 = 0 To glItemIntelUB
                    If glItemIntelIdx(X) <> -1 AndAlso goItemIntel(X).iItemTypeID = ObjectType.eAgent AndAlso (lParentID = -1 OrElse iParentTypeID <> ObjectType.ePlayer OrElse goItemIntel(X).lOtherPlayerID = lParentID) Then
                        cboData.AddItem(GetCacheObjectValue(goItemIntel(X).lItemID, goItemIntel(X).iItemTypeID))
                        cboData.ItemData(cboData.NewIndex) = goItemIntel(X).lItemID
                        cboData.ItemData2(cboData.NewIndex) = ObjectType.eAgent
                    End If
                Next X
            Case eTargetFilterType.eSolarSystem
                If goGalaxy Is Nothing = False Then
                    If goGalaxy.CurrentSystemIdx > -1 Then
                        Dim oSys As SolarSystem = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                        If oSys Is Nothing = False Then
                            cboData.AddItem(oSys.SystemName)
                            cboData.ItemData(cboData.NewIndex) = oSys.ObjectID
                            cboData.ItemData2(cboData.NewIndex) = oSys.ObjTypeID
                        End If
                    End If
                End If
            Case eTargetFilterType.ePlanet
                If NewTutorialManager.TutorialOn = True AndAlso goCurrentPlayer.yPlayerPhase = 0 Then
                    Dim oPlanet As Planet = Planet.GetTutorialPlanet()
                    If oPlanet Is Nothing = False Then
                        cboData.AddItem(oPlanet.PlanetName)
                        cboData.ItemData(cboData.NewIndex) = oPlanet.ObjectID
                        cboData.ItemData2(cboData.NewIndex) = ObjectType.ePlanet
                    End If
                ElseIf goGalaxy Is Nothing = False Then
                    If goGalaxy.CurrentSystemIdx > -1 Then
                        Dim oSys As SolarSystem = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                        If oSys Is Nothing = False Then
                            For X As Int32 = 0 To oSys.PlanetUB
                                If oSys.moPlanets(X) Is Nothing = False Then
                                    cboData.AddItem(oSys.moPlanets(X).PlanetName)
                                    cboData.ItemData(cboData.NewIndex) = oSys.moPlanets(X).ObjectID
                                    cboData.ItemData2(cboData.NewIndex) = oSys.moPlanets(X).ObjTypeID
                                End If
                            Next X
                        End If
                    End If
                End If
            Case eTargetFilterType.eSystemAndPlanets
                If goGalaxy Is Nothing = False Then
                    If goGalaxy.CurrentSystemIdx > -1 Then
                        Dim oSys As SolarSystem = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
                        If oSys Is Nothing = False Then
                            cboData.AddItem(oSys.SystemName & " (System)")
                            cboData.ItemData(cboData.NewIndex) = oSys.ObjectID
                            cboData.ItemData2(cboData.NewIndex) = oSys.ObjTypeID

                            For X As Int32 = 0 To oSys.PlanetUB
                                If oSys.moPlanets(X) Is Nothing = False Then
                                    cboData.AddItem(oSys.moPlanets(X).PlanetName)
                                    cboData.ItemData(cboData.NewIndex) = oSys.moPlanets(X).ObjectID
                                    cboData.ItemData2(cboData.NewIndex) = oSys.moPlanets(X).ObjTypeID
                                End If
                            Next X
                        End If
                    End If
                End If

            Case eTargetFilterType.eTarget_Component
                'PARENT SHOULD BE OUR PLAYER
                If iParentTypeID = ObjectType.ePlayer Then
                    FillComboFromItemIntelComponents(lParentID, cboData)
                End If
            Case eTargetFilterType.eMineralOrTargetComponent
                If iGrandParentTypeID = ObjectType.ePlayer Then 'RTP: This will never happen
                    FillComboFromItemIntelComponentsAndMinerals(lGrandParentID, cboData)
                End If
            Case eTargetFilterType.eTargetSpecialTechUnresearched
                If iParentTypeID <> ObjectType.ePlayer Then Return

                For X As Int32 = 0 To glPlayerTechKnowledgeUB
                    If glPlayerTechKnowledgeIdx(X) > -1 Then
                        Dim oPTK As PlayerTechKnowledge = goPlayerTechKnowledge(X)
                        If oPTK Is Nothing = False AndAlso oPTK.oTech.OwnerID = lParentID Then
                            If oPTK.oTech.ObjTypeID = ObjectType.eSpecialTech AndAlso oPTK.oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
                                cboData.AddItem(CType(oPTK.oTech, SpecialTech).sName)
                                cboData.ItemData(cboData.NewIndex) = oPTK.oTech.ObjectID
                                cboData.ItemData2(cboData.NewIndex) = ObjectType.eSpecialTech
                            End If
                        End If
                    End If
                Next
                cboData.SortList(False, False)
            Case eTargetFilterType.eTarget_UnitOrFacility
                'PARENT SHOULD BE OUR PLAYER
                If iParentTypeID = ObjectType.ePlayer Then
                    FillComboFromItemIntel(lParentID, ObjectType.eUnit, cboData)
                    FillComboFromItemIntel(lParentID, ObjectType.eFacility, cboData)
                End If
            Case eTargetFilterType.eTarget_Unit
                'PARENT SHOULD BE OUR PLAYER
                If iParentTypeID = ObjectType.ePlayer Then
                    FillComboFromItemIntel(lParentID, ObjectType.eUnit, cboData)
                End If
            Case eTargetFilterType.eTarget_Facility
                'PARENT SHOULD BE OUR PLAYER
                If iParentTypeID = ObjectType.ePlayer Then
                    FillComboFromItemIntel(lParentID, ObjectType.eFacility, cboData)
                End If
            Case eTargetFilterType.eTarget_Colony
                'PARENT SHOULD BE OUR PLAYER
                If iParentTypeID = ObjectType.ePlayer Then
                    FillComboFromItemIntel(lParentID, ObjectType.eColony, cboData)
                End If
            Case eTargetFilterType.eTarget_ComponentOrKnownMineral
                If iGrandParentTypeID = ObjectType.ePlayer Then
                    'FillComboFromItemIntelComponents(lGrandParentID, cboData)
                    For X As Int32 = 0 To glPlayerTechKnowledgeUB
                        If glPlayerTechKnowledgeIdx(X) > -1 Then
                            Dim oPTK As PlayerTechKnowledge = goPlayerTechKnowledge(X)
                            If oPTK Is Nothing = False AndAlso oPTK.oTech.OwnerID = lGrandParentID Then
                                Select Case oPTK.oTech.ObjTypeID
                                    Case ObjectType.eAlloyTech
                                        cboData.AddItem("Alloy-" & goPlayerTechKnowledge(X).oTech.GetComponentName)
                                        cboData.ItemData(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjectID
                                        cboData.ItemData2(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjTypeID
                                    Case ObjectType.eArmorTech
                                        cboData.AddItem("Armor-" & goPlayerTechKnowledge(X).oTech.GetComponentName)
                                        cboData.ItemData(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjectID
                                        cboData.ItemData2(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjTypeID
                                    Case ObjectType.eEngineTech
                                        cboData.AddItem("Engine-" & goPlayerTechKnowledge(X).oTech.GetComponentName)
                                        cboData.ItemData(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjectID
                                        cboData.ItemData2(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjTypeID
                                    Case ObjectType.eHullTech
                                        cboData.AddItem("Hull-" & goPlayerTechKnowledge(X).oTech.GetComponentName)
                                        cboData.ItemData(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjectID
                                        cboData.ItemData2(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjTypeID
                                    Case ObjectType.eRadarTech
                                        cboData.AddItem("Radar-" & goPlayerTechKnowledge(X).oTech.GetComponentName)
                                        cboData.ItemData(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjectID
                                        cboData.ItemData2(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjTypeID
                                    Case ObjectType.eShieldTech
                                        cboData.AddItem("Shield-" & goPlayerTechKnowledge(X).oTech.GetComponentName)
                                        cboData.ItemData(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjectID
                                        cboData.ItemData2(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjTypeID
                                    Case ObjectType.eWeaponTech
                                        cboData.AddItem("Weapon-" & goPlayerTechKnowledge(X).oTech.GetComponentName)
                                        cboData.ItemData(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjectID
                                        cboData.ItemData2(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjTypeID
                                    Case Else
                                        'next for
                                End Select
                            End If
                        End If
                    Next X

      
                    If lAnyIndex <> -1 Then cboData.ItemData2(lAnyIndex) = ObjectType.eMineral
                    Dim lSorted() As Int32 = GetSortedMineralIdxArray(False, False)
                    If lSorted Is Nothing = False Then
                        For X As Int32 = 0 To lSorted.GetUpperBound(0)
                            If glMineralIdx(lSorted(X)) <> -1 Then
                                cboData.AddItem("Mineral-" & goMinerals(lSorted(X)).MineralName)
                                cboData.ItemData(cboData.NewIndex) = goMinerals(lSorted(X)).ObjectID
                                cboData.ItemData2(cboData.NewIndex) = ObjectType.eMineral
                            End If
                        Next X
                    End If
                End If
                cboData.SortList(False, False)
        End Select

    End Sub

    Private Sub FillComboFromItemIntelComponents(ByVal lPlayerID As Int64, ByRef cboData As UIComboBox)
        For X As Int32 = 0 To glItemIntelUB
            If glItemIntelIdx(X) <> -1 Then
                If goItemIntel(X).lOtherPlayerID = lPlayerID Then
                    Select Case goItemIntel(X).iItemTypeID
                        Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, _
                         ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eSpecialTech, ObjectType.eWeaponTech
                            cboData.AddItem(GetCacheObjectValue(goItemIntel(X).lItemID, goItemIntel(X).iItemTypeID))
                            cboData.ItemData(cboData.NewIndex) = goItemIntel(X).lItemID
                            cboData.ItemData2(cboData.NewIndex) = goItemIntel(X).iItemTypeID
                    End Select
                End If
            End If
        Next X

        For X As Int32 = 0 To glPlayerTechKnowledgeUB
            If glPlayerTechKnowledgeIdx(X) > -1 Then
                Dim oPTK As PlayerTechKnowledge = goPlayerTechKnowledge(X)
                If oPTK Is Nothing = False AndAlso oPTK.oTech.OwnerID = lPlayerID Then
                    cboData.AddItem(goPlayerTechKnowledge(X).oTech.GetComponentName) 'GetCacheObjectValue(goPlayerTechKnowledge(X).oTech.ObjectID, goPlayerTechKnowledge(X).oTech.ObjTypeID))
                    cboData.ItemData(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjectID
                    cboData.ItemData2(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjTypeID
                End If
            End If
        Next X
    End Sub

    Private Sub FillComboFromItemIntelComponentsAndMinerals(ByVal lPlayerID As Int64, ByRef cboData As UIComboBox)
        For X As Int32 = 0 To glItemIntelUB
            If glItemIntelIdx(X) <> -1 Then
                If goItemIntel(X).lOtherPlayerID = lPlayerID Then
                    Select Case goItemIntel(X).iItemTypeID
                        Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, _
                         ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eSpecialTech, ObjectType.eWeaponTech, ObjectType.eMineral
                            cboData.AddItem(GetCacheObjectValue(goItemIntel(X).lItemID, goItemIntel(X).iItemTypeID))
                            cboData.ItemData(cboData.NewIndex) = goItemIntel(X).lItemID
                            cboData.ItemData2(cboData.NewIndex) = goItemIntel(X).iItemTypeID
                    End Select
                End If
            End If
        Next X
        For X As Int32 = 0 To glPlayerTechKnowledgeUB
            If glPlayerTechKnowledgeIdx(X) > -1 Then
                Dim oPTK As PlayerTechKnowledge = goPlayerTechKnowledge(X)
                If oPTK Is Nothing = False AndAlso oPTK.oTech.OwnerID = lPlayerID Then
                    cboData.AddItem(GetCacheObjectValue(goPlayerTechKnowledge(X).oTech.ObjectID, goPlayerTechKnowledge(X).oTech.ObjTypeID))
                    cboData.ItemData(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjectID
                    cboData.ItemData2(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjTypeID
                End If
            End If
        Next X
    End Sub
    Private Sub FillComboFromItemIntel(ByVal lPlayerID As Int32, ByVal iTypeID As Int16, ByRef cboData As UIComboBox)
        For X As Int32 = 0 To glItemIntelUB
            If glItemIntelIdx(X) <> -1 Then
                If goItemIntel(X).lOtherPlayerID = lPlayerID AndAlso goItemIntel(X).iItemTypeID = iTypeID Then
                    Dim sName As String = ""
                    sName = GetCacheObjectValue(goItemIntel(X).lItemID, goItemIntel(X).iItemTypeID)
                    If goItemIntel(X).EnvirTypeID > 0 AndAlso goItemIntel(X).EnvirID > 0 Then
                        sName &= " - " & GetCacheObjectValue(goItemIntel(X).EnvirID, goItemIntel(X).EnvirTypeID)
                    End If
                    cboData.AddItem(sName)
                    cboData.ItemData(cboData.NewIndex) = goItemIntel(X).lItemID
                    cboData.ItemData2(cboData.NewIndex) = goItemIntel(X).iItemTypeID
                End If
            End If
        Next X
    End Sub

    Private Sub cboTarget1_ItemSelected(ByVal lItemIndex As Integer) Handles cboTarget1.ItemSelected
        If cboTarget1.ListIndex <> -1 Then
            'ok, get our player
            If myTarget2ListType <> eTargetFilterType.eUnusedTarget Then
                'now, fill our target 2 
                Dim lPlayerID As Int32 = cboTarget1.ItemData(cboTarget1.ListIndex)
                FillTargetCBOFromFilter(myTarget2ListType, cboTarget2, lPlayerID, ObjectType.ePlayer, -1, -1)
                cboTarget2.Enabled = True
            Else : cboTarget2.Enabled = False
            End If
        End If
    End Sub

    Private Sub cboTarget2_ItemSelected(ByVal lItemIndex As Integer) Handles cboTarget2.ItemSelected
        If cboTarget2.ListIndex <> -1 Then

            Dim lID As Int32 = cboTarget2.ItemData(cboTarget2.ListIndex)
            Dim iTypeID As Int16 = CShort(cboTarget2.ItemData2(cboTarget2.ListIndex))

            Dim lID1 As Int32 = cboTarget1.ItemData(cboTarget1.ListIndex)
            Dim iTypeID1 As Int16 = CShort(cboTarget1.ItemData2(cboTarget1.ListIndex))

            If NewTutorialManager.TutorialOn = True Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionTargetSelected, lID, iTypeID, -1, "")
            End If

            If myTarget3ListType <> eTargetFilterType.eUnusedTarget Then
                FillTargetCBOFromFilter(myTarget3ListType, cboTarget3, lID, iTypeID, lID1, iTypeID1)
                cboTarget3.Enabled = True
            Else : cboTarget3.Enabled = False
            End If
        End If
    End Sub

    Private Sub cboMethod_ItemSelected(ByVal lItemIndex As Integer) Handles cboMethod.ItemSelected
        'ok, clear goals
        lstGoals.Clear()
        lblGoalName.Caption = ""
        lblGoalTime.Caption = ""
        lblGoalRisk.Caption = ""
        txtGoalDesc.Caption = ""
        lblGoalPhase.Caption = ""

        Dim oMission As Mission = GetSelectedMission()
        If oMission Is Nothing Then Return

        If cboMethod.ListIndex <> -1 Then
            Dim lMethodID As Int32 = cboMethod.ItemData(cboMethod.ListIndex)

            If NewTutorialManager.TutorialOn = True Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionMethodSelected, lMethodID, -1, -1, "")
            End If

            If chkSafehouse.Value = True Then
                Dim oSafehouse As Goal = Goal.GetSafehouseGoal()
                If oSafehouse Is Nothing = False Then
                    lstGoals.AddItem(oSafehouse.sGoalName)
                    lstGoals.ItemData(lstGoals.NewIndex) = oSafehouse.ObjectID
                End If
                moMission.ySafeHouseSetting = 1
                moMission.oSafehouseMissionGoal = New PlayerMissionGoal
            Else
                moMission.ySafeHouseSetting = 0
            End If

            moMission.lMethodID = lMethodID

            For X As Int32 = 0 To glMissionMethodUB
                If glMissionMethodIdx(X) = lMethodID Then
                    lblMethodDesc.Caption = gsMissionDescs(X)
                    Exit For
                End If
            Next X

            Dim lSorted() As Int32 = Nothing
            Dim lSortedUB As Int32 = -1
            For X As Int32 = 0 To oMission.GoalUB
                Dim lIdx As Int32 = -1
                If oMission.MethodIDs(X) = lMethodID Then
                    For Y As Int32 = 0 To lSortedUB
                        If oMission.Goals(lSorted(Y)).MissionPhase > oMission.Goals(X).MissionPhase Then
                            lIdx = Y
                            Exit For
                        End If
                    Next Y
                    lSortedUB += 1
                    ReDim Preserve lSorted(lSortedUB)
                    If lIdx = -1 Then
                        lSorted(lSortedUB) = X
                    Else
                        For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                            lSorted(Y) = lSorted(Y - 1)
                        Next Y
                        lSorted(lIdx) = X
                    End If
                End If
            Next X

            For X As Int32 = 0 To lSortedUB
                lstGoals.AddItem(oMission.Goals(lSorted(X)).sGoalName)
                lstGoals.ItemData(lstGoals.NewIndex) = oMission.Goals(lSorted(X)).ObjectID
            Next X
        End If
    End Sub

    Private Function GetSelectedMission() As Mission
        If cboMission.ListIndex <> -1 Then
            Dim lID As Int32 = cboMission.ItemData(cboMission.ListIndex)
            For X As Int32 = 0 To glMissionUB
                If glMissionIdx(X) = lID Then
                    Return goMissions(X)
                End If
            Next X
        End If
        Return Nothing
    End Function

    Private Sub lstGoals_ItemClick(ByVal lIndex As Integer) Handles lstGoals.ItemClick
        lblGoalName.Caption = ""
        lblGoalTime.Caption = ""
        lblGoalRisk.Caption = ""
        txtGoalDesc.Caption = ""
        lblGoalPhase.Caption = ""

        Dim oMission As Mission = GetSelectedMission()
        If oMission Is Nothing Then Return
        Dim lMethodID As Int32 = -1
        If cboMethod.ListIndex <> -1 Then lMethodID = cboMethod.ItemData(cboMethod.ListIndex)
        If lMethodID = -1 Then Return

        If lstGoals.ListIndex > -1 Then
            Dim lGoalID As Int32 = lstGoals.ItemData(lstGoals.ListIndex)

            Dim oGoal As Goal = Nothing
            Dim oPMG As PlayerMissionGoal = Nothing
            If lGoalID = Goal.ml_SAFEHOUSE_GOAL_ID Then
                oGoal = Goal.GetSafehouseGoal()
                If moMission.oSafehouseMissionGoal Is Nothing Then
                    moMission.oSafehouseMissionGoal = New PlayerMissionGoal()
                    With moMission.oSafehouseMissionGoal
                        .oGoal = oGoal
                        .oMission = moMission
                    End With
                End If
                oPMG = moMission.oSafehouseMissionGoal
            Else
                For X As Int32 = 0 To oMission.GoalUB
                    If oMission.MethodIDs(X) = lMethodID AndAlso oMission.Goals(X).ObjectID = lGoalID Then

                        If NewTutorialManager.TutorialOn = True Then
                            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionGoalSelected, lGoalID, -1, -1, "")
                        End If

                        oGoal = oMission.Goals(X)
                        oPMG = moMission.oMissionGoals(X)
                        Exit For
                    End If
                Next X
            End If

            If oGoal Is Nothing = False Then
                'ok, found  it
                With oGoal
                    lblGoalName.Caption = "Goal: " & .sGoalName
                    lblGoalTime.Caption = "Goal Time: about " & GetDurationFromSeconds(.BaseTime, True)
                    lblGoalRisk.Caption = "Risk Assessment: " & Goal.GetRiskAssessment(.RiskOfDetection)
                    lblGoalRisk.ForeColor = Goal.GetRiskCaptionColor(.RiskOfDetection)
                    txtGoalDesc.Caption = .sGoalDesc
                    Select Case .MissionPhase
                        Case eMissionPhase.eFlippingTheSwitch
                            lblGoalPhase.Caption = "Phase: Execution"
                        Case eMissionPhase.ePreparationTime
                            lblGoalPhase.Caption = "Phase: Preparation"
                        Case eMissionPhase.eReinfiltrationPhase
                            lblGoalPhase.Caption = "Phase: Reinfiltration"
                        Case eMissionPhase.eSettingTheStage
                            lblGoalPhase.Caption = "Phase: Setting the Stage"
                    End Select
                End With
                fraSkillSets.SetFromGoal(oGoal, moMission)

                Dim lSkillsetID As Int32 = -1
                If oPMG.lAssignmentUB <> -1 AndAlso oPMG.oSkillSet Is Nothing = False Then
                    lSkillsetID = oPMG.oSkillSet.SkillSetID
                End If
                fraSkillSets.SetSkillsetID(lSkillsetID)

                For Y As Int32 = 0 To oPMG.lAssignmentUB
                    If oPMG.oAssignments(Y) Is Nothing = False Then
                        fraSkillSets.SetAgentAssignment(oPMG.oAssignments(Y))
                    End If
                Next Y
            End If

        End If
    End Sub

    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
        '
    End Sub

    Private Sub fraSkillSets_AgentAssignment(ByRef oSkill As SkillSet_Skill, ByRef oSender As ctlAgentAssignment) Handles fraSkillSets.AgentAssignment
        'Ok, take the currently selected agent...

        If mlSelectedAgentID = -1 Then
            MyBase.moUILib.AddNotification("Select an agent to assign and then click Assign.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        Else
            Dim oAgent As Agent = Nothing
            For X As Int32 = 0 To goCurrentPlayer.AgentUB
                If goCurrentPlayer.AgentIdx(X) = mlSelectedAgentID Then
                    oAgent = goCurrentPlayer.Agents(X)
                    Exit For
                End If
            Next X

            Dim lMethodID As Int32 = -1
            If cboMethod.ListIndex > -1 Then lMethodID = cboMethod.ItemData(cboMethod.ListIndex)

            If oAgent Is Nothing = False Then

                If oAgent Is Nothing = False Then ' AndAlso oAgent.bRequestedSkillList = False Then
                    'oAgent.bRequestedSkillList = True
                    'Dim yOut(5) As Byte
                    'System.BitConverter.GetBytes(GlobalMessageCode.eGetSkillList).CopyTo(yOut, 0)
                    'System.BitConverter.GetBytes(oAgent.ObjectID).CopyTo(yOut, 2)
                    'MyBase.moUILib.SendMsgToPrimary(yOut)
                    'Return
                End If

                'check if the agent has the specified skill
                If oSkill Is Nothing = False Then

                    Dim yProf As Byte = oAgent.GetSkillProficiency(oSkill.oSkill.ObjectID)
                    If yProf < oSkill.oSkill.MinVal Then
                        yProf = oAgent.GetSkillProficiency(36)  'naturally talented
                        If yProf < oSkill.oSkill.MinVal Then
                            MyBase.moUILib.AddNotification("The agent selected does not have " & oSkill.oSkill.SkillName & " at a proficient level.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return
                        Else

                            If oSkill.oSkill.SkillType <> 0 Then
                                MyBase.moUILib.AddNotification("Unable to use Naturally Talented for " & oSkill.oSkill.SkillName & " as it is an ability.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                                Return
                            End If

                            Dim lTmpGoal As Int32 = oSkill.oSkillSet.oGoal.ObjectID
                            Dim bValid As Boolean = False
                            Dim oTmpPMG As PlayerMissionGoal = Nothing
                            If lTmpGoal = Goal.ml_SAFEHOUSE_GOAL_ID Then
                                oTmpPMG = moMission.oSafehouseMissionGoal
                            Else
                                For X As Int32 = 0 To moMission.oMissionGoals.GetUpperBound(0)
                                    If moMission.oMission.MethodIDs(X) = lMethodID AndAlso moMission.oMissionGoals(X).oGoal.ObjectID = lTmpGoal Then
                                        oTmpPMG = moMission.oMissionGoals(X)
                                        Exit For
                                    End If
                                Next X
                            End If
                            If oTmpPMG Is Nothing = False Then
                                If oTmpPMG.AgentAssignedUsingNaturalTalents(oAgent.ObjectID) = True Then
                                    MyBase.moUILib.AddNotification("The agent selected does not have " & oSkill.oSkill.SkillName & " at a proficient level.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                                    Return
                                End If
                            End If

                            MyBase.moUILib.AddNotification("Agent added using Naturally Talented.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        End If
                    End If
                End If
                If NewTutorialManager.TutorialOn = True Then
                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionSkillsetAssign, oAgent.ObjectID, oSkill.oSkillSet.SkillSetID, oSkill.oSkill.ObjectID, "")
                End If
                oSender.SetAgent(oAgent)
            End If

            'Now, make the assignment
            Dim lGoalID As Int32 = -1
            Dim yPhase As Byte = 0
            Dim oPMG As PlayerMissionGoal = Nothing
            If oSkill Is Nothing = False Then
                lGoalID = oSkill.oSkillSet.oGoal.ObjectID
                For X As Int32 = 0 To moMission.oMissionGoals.GetUpperBound(0)
                    If moMission.oMission.MethodIDs(X) = lMethodID AndAlso moMission.oMissionGoals(X).oGoal.ObjectID = lGoalID Then
                        moMission.oMissionGoals(X).AddAgentAssignment(oAgent, oSkill.oSkill)
                        oPMG = moMission.oMissionGoals(X)
                        yPhase = moMission.oMissionGoals(X).oGoal.MissionPhase
                        Exit For
                    End If
                Next X
            End If

            'Now, ensure that the agent is not assigned cover
            For X As Int32 = 0 To moMission.lPhaseUB
                If moMission.oPhases(X).yPhase = yPhase Then
                    moMission.oPhases(X).RemoveCoverAgent(oAgent.ObjectID)
                    RefreshCoverAgentList()
                    Exit For
                End If
            Next X
        End If
    End Sub

    Private Sub fraSkillSets_SkillSetSelected(ByRef oGoal As Goal, ByRef oSkillset As SkillSet) Handles fraSkillSets.SkillSetSelected
        If moMission Is Nothing = False AndAlso moMission.oMission Is Nothing = False AndAlso oGoal Is Nothing = False Then
            Dim lMethodID As Int32 = -1
            If cboMethod Is Nothing = False AndAlso cboMethod.ListIndex > -1 Then
                lMethodID = cboMethod.ItemData(cboMethod.ListIndex)
            End If
            For X As Int32 = 0 To moMission.oMission.GoalUB
                If moMission.oMission.Goals(X).ObjectID = oGoal.ObjectID AndAlso moMission.oMission.MethodIDs(X) = lMethodID Then

                    If NewTutorialManager.TutorialOn = True AndAlso oSkillset Is Nothing = False Then
                        NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionSkillsetSelected, oSkillset.SkillSetID, -1, -1, "")
                    End If

                    moMission.oMissionGoals(X).oSkillSet = oSkillset
                    moMission.oMissionGoals(X).lAssignmentUB = -1
                    moFilterSkillSet = oSkillset
                    Exit For
                End If
            Next X
        End If
    End Sub

    Private Sub fraSkillSets_SetSkillsetFilter(ByRef oSkillset As SkillSet) Handles fraSkillSets.SetSkillsetFilter
        moFilterSkillSet = oSkillset
        Me.IsDirty = True
    End Sub

    Private Sub fraSkillSets_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles fraSkillSets.OnMouseMove
        frmMission_OnMouseMove(lMouseX, lMouseY, lButton)
    End Sub

    Private Sub btnAgentLeft_Click(ByVal sName As String) Handles btnAgentLeft.Click
        mlScrollValue -= 1
        If mlScrollValue < 0 Then mlScrollValue = 0
        Me.IsDirty = True
    End Sub

    Private Sub btnAgentRight_Click(ByVal sName As String) Handles btnAgentRight.Click
        mlScrollValue += 1
        If mlScrollValue > mlScrollMaxValue Then mlScrollValue = mlScrollMaxValue
        Me.IsDirty = True
    End Sub

    Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
        'Cancels the mission
        If moMission Is Nothing = False Then
            If moMission.PM_ID > -1 Then
                'check if the mission can be cancelled
                If moMission.yCurrentPhase = eMissionPhase.eInPlanning OrElse moMission.yCurrentPhase = eMissionPhase.ePreparationTime Then
                    'send cancel msg to the primary
                    Dim yMsg() As Byte = GenerateMissionMsg(2) '2 = cancel
                    If yMsg Is Nothing = False Then MyBase.moUILib.SendMsgToPrimary(yMsg)
                    For X As Int32 = 0 To goCurrentPlayer.PlayerMissionUB
                        If goCurrentPlayer.PlayerMissionIdx(X) = moMission.PM_ID Then
                            goCurrentPlayer.PlayerMissions(X).yCurrentPhase = eMissionPhase.eCancelled
                        End If
                    Next X
                Else
                    MyBase.moUILib.AddNotification("This mission cannot be cancelled at this time.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return
                End If
            End If
        End If
        'remove the window
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnExecuteNow_Click(ByVal sName As String) Handles btnExecuteNow.Click
        '1) Send the primary a message to execute now... do a validation
        If ValidateMission(True) = False Then Return

        Dim yMsg() As Byte = GenerateMissionMsg(1)      '1 for execute now
        If yMsg Is Nothing = False Then
            MyBase.moUILib.SendMsgToPrimary(yMsg)
            If NewTutorialManager.TutorialOn = True Then
                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionExecuteNow, -1, -1, -1, "")
            End If
        End If
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Sub btnSaveForLater_Click(ByVal sName As String) Handles btnSaveForLater.Click
        '1) Send the primary a message to save for later
        Try
            If ValidateMission(False) = False Then Return
            Dim yMsg() As Byte = GenerateMissionMsg(0)              '0 = save for later
            If yMsg Is Nothing = False Then
                MyBase.moUILib.SendMsgToPrimary(yMsg)
            End If
        Catch ex As Exception
            MyBase.moUILib.AddNotification("Unable to save mission for later.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End Try
        MyBase.moUILib.RemoveWindow(Me.ControlName)
    End Sub

    Private Function ValidateMission(ByVal bExact As Boolean) As Boolean
        Try
            'is omission set?
            If moMission Is Nothing Then Return False

            If moMission.oMission Is Nothing Then
                MyBase.moUILib.AddNotification("You must select a mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If

            If moMission.oMission.ProgramControlID = eMissionResult.eGetFacilityList Then
                'ok, check our target 2...
                Dim lT2Val As Int32 = -1
                If cboTarget2.ListIndex > -1 Then lT2Val = cboTarget2.ItemData(cboTarget2.ListIndex)
                Dim bSys As Boolean = False
                If lT2Val = ProductionType.eSpaceStationSpecial Then
                    bSys = True
                End If
                Dim lT3Val As Int32 = -1
                If cboTarget3.ListIndex > -1 Then lT3Val = cboTarget3.ItemData2(cboTarget3.ListIndex)
                If lT3Val = ObjectType.ePlanet Then
                    If bSys = True Then
                        MyBase.moUILib.AddNotification("You must select the system to find space stations.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return False
                    End If
                ElseIf lT3Val = ObjectType.eSolarSystem Then
                    If bSys = False Then
                        MyBase.moUILib.AddNotification("You must select a planet when searching for anything other than space stations.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return False
                    End If
                End If

            End If
            'myTarget2ListType = eTargetFilterType.eProductionType
            'myTarget3ListType = eTargetFilterType.eSystemAndPlanets

            If bExact = True Then
                If cboMethod.ListIndex < 0 Then
                    MyBase.moUILib.AddNotification("You must select a method for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                End If
                Dim lMethodID As Int32 = cboMethod.ItemData(cboMethod.ListIndex)
                Dim bGood As Boolean = False
                For X As Int32 = 0 To moMission.oMission.GoalUB
                    If lMethodID = moMission.oMission.MethodIDs(X) Then
                        bGood = True
                        Exit For
                    End If
                Next X
                If bGood = False Then
                    MyBase.moUILib.AddNotification("You must select a method for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                End If

                bGood = False
                'check our target and target types
                Select Case CType(moMission.oMission.ProgramControlID, eMissionResult)
                    Case eMissionResult.ePlantEvidence
                        'player, otherplayerid, pm_id
                        If cboTarget1.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a target player for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                        If cboTarget2.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a target player for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                        If cboTarget3.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a mission you have ran to implicate.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                    Case eMissionResult.eImpedeCurrentDevelopment
                        If cboTarget1.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a target player for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                        'If cboTarget2.ListIndex < 0 Then
                        '    MyBase.moUILib.AddNotification("You must select a target special tech for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        '    Return False
                        'End If
                    Case eMissionResult.eDestroyCurrentSpecialProject
                        If cboTarget1.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a target player for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                        If cboTarget2.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a target special tech for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                    Case eMissionResult.eFindMineral
                        'No player, mineral
                        If cboTarget2.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a mineral to search for on this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                    Case eMissionResult.eSearchAndRescueAgent
                        'no player, agent
                        If cboTarget2.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select an agent to search for on this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                    Case eMissionResult.eReconPlanetMap
                        'no player, locx, locz
                    Case eMissionResult.eGeologicalSurvey
                        'no player, planet
                        If cboTarget2.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a planet to survey for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                    Case eMissionResult.eGetColonyBudget
                        'player, planet
                        If cboTarget1.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a target player for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                        If cboTarget2.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a planet to survey for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                    Case eMissionResult.eDecreaseHousingMorale, eMissionResult.eDecreaseTaxMorale, eMissionResult.eDecreaseUnemploymentMorale
                        If cboTarget1.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a target player for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                        If cboTarget2.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a planet to survey for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                    Case eMissionResult.eGetFacilityList
                        If cboTarget1.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a target player for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                        If cboTarget2.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a production type to find.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                        If cboTarget3.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a planet/system to search.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If
                    Case eMissionResult.eTutorialFindFactory, eMissionResult.eAudit
                        'do nothing,
                    Case Else
                        If cboTarget1.ListIndex < 0 Then
                            MyBase.moUILib.AddNotification("You must select a target player for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                            Return False
                        End If

                End Select

            'all assignments must exist...
            If chkSafehouse.Value = True Then
                If moMission.oSafehouseMissionGoal Is Nothing Then
                    MyBase.moUILib.AddNotification("Specify more detail for the safehouse goal.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                End If
                If moMission.oSafehouseMissionGoal.oSkillSet Is Nothing Then
                    MyBase.moUILib.AddNotification("Select a skillset for the safehouse goal.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    Return False
                End If
                If VerifyMissionGoal(moMission.oSafehouseMissionGoal) = False Then Return False
            End If
            For X As Int32 = 0 To moMission.oMission.GoalUB
                If moMission.oMission.MethodIDs(X) = lMethodID Then
                    If moMission.oMissionGoals(X).oSkillSet Is Nothing Then
                        MyBase.moUILib.AddNotification("All mission goals must have skillsets selected.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                        Return False
                    End If

                    If VerifyMissionGoal(moMission.oMissionGoals(X)) = False Then Return False
                    'Dim bSkillGood(moMission.oMissionGoals(X).oSkillSet.SkillUB) As Boolean
                    'For Y As Int32 = 0 To moMission.oMissionGoals(X).oSkillSet.SkillUB
                    '	bSkillGood(Y) = False
                    'Next Y

                    'Dim bAgentDetained As Boolean = False
                    'For Y As Int32 = 0 To moMission.oMissionGoals(X).lAssignmentUB
                    '	Dim lSkillID As Int32 = moMission.oMissionGoals(X).oAssignments(Y).oSkill.ObjectID
                    '	If moMission.oMissionGoals(X).oAssignments(Y).oAgent Is Nothing Then
                    '		MyBase.moUILib.AddNotification("Not all goals have fully assigned skillsets.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    '		Return False
                    '	End If
                    '	If (moMission.oMissionGoals(X).oAssignments(Y).oAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 OrElse (moMission.oMissionGoals(X).oAssignments(Y).oAgent.lAgentStatus And AgentStatus.IsDead) <> 0 Then
                    '		bAgentDetained = True
                    '	Else
                    '		For Z As Int32 = 0 To moMission.oMissionGoals(X).oSkillSet.SkillUB
                    '			If moMission.oMissionGoals(X).oSkillSet.Skills(Z).oSkill.ObjectID = lSkillID Then
                    '				bSkillGood(Z) = True
                    '				Exit For
                    '			End If
                    '		Next Z
                    '	End If
                    'Next Y

                    ''Now, verify our goods
                    'For Y As Int32 = 0 To moMission.oMissionGoals(X).oSkillSet.SkillUB
                    '	If bSkillGood(Y) = False Then
                    '		MyBase.moUILib.AddNotification("This mission is not a valid mission. ", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                    '		Return False
                    '	End If
                    'Next Y
                End If
            Next X
            End If
        Catch ex As Exception
            MyBase.moUILib.AddNotification("This mission is not a valid mission. " & ex.Message, Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return False
        End Try
        Return True
    End Function

    Private Function VerifyMissionGoal(ByVal oPMG As PlayerMissionGoal) As Boolean
        Dim bSkillGood(oPMG.oSkillSet.SkillUB) As Boolean
        For Y As Int32 = 0 To oPMG.oSkillSet.SkillUB
            bSkillGood(Y) = False
        Next Y

        Dim bAgentDetained As Boolean = False
        For Y As Int32 = 0 To oPMG.lAssignmentUB
            Dim lSkillID As Int32 = oPMG.oAssignments(Y).oSkill.ObjectID
            If oPMG.oAssignments(Y).oAgent Is Nothing Then
                MyBase.moUILib.AddNotification("Not all goals have fully assigned skillsets.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
            If (oPMG.oAssignments(Y).oAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 OrElse (oPMG.oAssignments(Y).oAgent.lAgentStatus And AgentStatus.IsDead) <> 0 Then
                bAgentDetained = True
            Else
                For Z As Int32 = 0 To oPMG.oSkillSet.SkillUB
                    If oPMG.oSkillSet.Skills(Z).oSkill.ObjectID = lSkillID Then
                        bSkillGood(Z) = True
                        Exit For
                    End If
                Next Z
            End If
        Next Y

        'Now, verify our goods
        For Y As Int32 = 0 To oPMG.oSkillSet.SkillUB
            If bSkillGood(Y) = False Then
                MyBase.moUILib.AddNotification("This mission is not a valid mission. ", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return False
            End If
        Next Y
        Return True
    End Function

    Private Function GenerateMissionMsg(ByVal yActionID As Byte) As Byte()
        'Ok, let's create our mission and submit it to the server
        Dim yMsg() As Byte
        Dim lPos As Int32 = 0
        Dim yPhaseCnt As Byte = 0
        Dim yGoalCnt As Byte = 0
        Dim lPhaseCvrCnt() As Int32 = Nothing
        Dim lAssignCnt() As Int32 = Nothing

        If moMission Is Nothing Then Return Nothing

        With moMission
            Dim lMethodID As Int32 = -1
            If cboMethod.ListIndex > -1 Then lMethodID = cboMethod.ItemData(cboMethod.ListIndex)
            'determine our counts now
            ReDim lPhaseCvrCnt(.lPhaseUB)
            For X As Int32 = 0 To .lPhaseUB
                lPhaseCvrCnt(X) = 0
                If .oPhases(X) Is Nothing = False Then
                    yPhaseCnt += CByte(1)
                    For Y As Int32 = 0 To .oPhases(X).lCoverAgentUB
                        If .oPhases(X).lCoverAgentIdx(Y) <> -1 Then lPhaseCvrCnt(X) += 1
                    Next Y
                End If
            Next X
            If .oMission Is Nothing = False Then
                ReDim lAssignCnt(.oMission.GoalUB)
                For X As Int32 = 0 To .oMission.GoalUB
                    lAssignCnt(X) = 0
                    If .oMissionGoals(X) Is Nothing = False AndAlso .oMission.MethodIDs(X) = lMethodID Then
                        yGoalCnt += CByte(1)
                        For Y As Int32 = 0 To .oMissionGoals(X).lAssignmentUB
                            If .oMissionGoals(X).oAssignments(Y) Is Nothing = False Then lAssignCnt(X) += 1
                        Next Y
                    End If
                Next X
            Else
                ReDim lAssignCnt(-1)
            End If

            'Ok, we have everything, let's determine total size
            '29 + (yPhaseCnt * 2) + (lSumOfPhaseCover * 4) + (yGoalCnt * 8) + (lSumAgentAssignment * 8)
            Dim lSumOfPhaseCover As Int32 = 0
            Dim lSumAgentAssignment As Int32 = 0
            For X As Int32 = 0 To .lPhaseUB
                lSumOfPhaseCover += lPhaseCvrCnt(X)
            Next X
            For X As Int32 = 0 To lAssignCnt.GetUpperBound(0)
                lSumAgentAssignment += lAssignCnt(X)
            Next X

            Dim lSafehouseSize As Int32 = 0
            If chkSafehouse.Value = True Then lSafehouseSize = 20

            ReDim yMsg(33 + (CInt(yPhaseCnt) * 2) + (lSumOfPhaseCover * 4) + (CInt(yGoalCnt) * 9) + (lSumAgentAssignment * 8))  '28 not 29 for ub

            System.BitConverter.GetBytes(GlobalMessageCode.eSubmitMission).CopyTo(yMsg, lPos) : lPos += 2
            yMsg(lPos) = yActionID : lPos += 1
            System.BitConverter.GetBytes(.PM_ID).CopyTo(yMsg, lPos) : lPos += 4
            If .oMission Is Nothing = False Then
                System.BitConverter.GetBytes(.oMission.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
            Else
                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
            End If
            If cboTarget1.Visible = False OrElse cboTarget1.ListIndex < 0 Then
                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
            Else : System.BitConverter.GetBytes(cboTarget1.ItemData(cboTarget1.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
            End If
            If cboTarget2.Visible = True AndAlso cboTarget2.Enabled = True AndAlso cboTarget2.ListIndex > -1 Then
                System.BitConverter.GetBytes(cboTarget2.ItemData(cboTarget2.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(CShort(cboTarget2.ItemData2(cboTarget2.ListIndex))).CopyTo(yMsg, lPos) : lPos += 2
            Else
                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(-1S).CopyTo(yMsg, lPos) : lPos += 2
            End If
            If cboTarget3.Visible = True AndAlso cboTarget3.Enabled = True AndAlso cboTarget3.ListIndex > -1 Then
                System.BitConverter.GetBytes(cboTarget3.ItemData(cboTarget3.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(CShort(cboTarget3.ItemData2(cboTarget3.ListIndex))).CopyTo(yMsg, lPos) : lPos += 2
            Else
                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                System.BitConverter.GetBytes(-1S).CopyTo(yMsg, lPos) : lPos += 2
            End If
            yMsg(lPos) = .ySafeHouseSetting : lPos += 1
            If cboMethod.ListIndex > -1 Then
                System.BitConverter.GetBytes(cboMethod.ItemData(cboMethod.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
            Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
            End If

            If chkSafehouse.Value = True Then
                If moMission.oSafehouseMissionGoal Is Nothing = False Then
                    With moMission.oSafehouseMissionGoal
                        'skillsetid
                        If .oSkillSet Is Nothing = False Then
                            System.BitConverter.GetBytes(.oSkillSet.SkillSetID).CopyTo(yMsg, lPos) : lPos += 4
                        Else : lPos += 4
                        End If
                        If .lAssignmentUB > -1 Then
                            'assignment 1 agentid
                            If .oAssignments(0).oAgent Is Nothing = False Then
                                System.BitConverter.GetBytes(.oAssignments(0).oAgent.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                            Else
                                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                            End If
                            'assignment 1 skillid
                            If .oAssignments(0).oSkill Is Nothing = False Then
                                System.BitConverter.GetBytes(.oAssignments(0).oSkill.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                            Else
                                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                            End If
                            If .lAssignmentUB > 0 Then
                                'assignment 2 agentid
                                If .oAssignments(1).oAgent Is Nothing = False Then
                                    System.BitConverter.GetBytes(.oAssignments(1).oAgent.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                                Else
                                    System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                                End If
                                'assignment 2 skillid
                                If .oAssignments(1).oSkill Is Nothing = False Then
                                    System.BitConverter.GetBytes(.oAssignments(1).oSkill.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                                Else
                                    System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                                End If
                            Else : lPos += 8
                            End If
                        Else : lPos += 16
                        End If

                    End With
                Else : lPos += 20
                End If
            End If

            'ok, get the phases
            yMsg(lPos) = yPhaseCnt : lPos += 1
            For X As Int32 = 0 To .lPhaseUB
                If .oPhases(X) Is Nothing = False Then
                    yMsg(lPos) = .oPhases(X).yPhase : lPos += 1
                    yMsg(lPos) = CByte(lPhaseCvrCnt(X)) : lPos += 1

                    For Y As Int32 = 0 To .oPhases(X).lCoverAgentUB
                        If .oPhases(X).lCoverAgentIdx(Y) <> -1 Then
                            System.BitConverter.GetBytes(.oPhases(X).lCoverAgentIdx(Y)).CopyTo(yMsg, lPos) : lPos += 4
                        End If
                    Next Y
                End If
            Next X

            'now for our goals
            yMsg(lPos) = yGoalCnt : lPos += 1
            If .oMission Is Nothing = False Then
                For X As Int32 = 0 To .oMission.GoalUB
                    If .oMissionGoals(X) Is Nothing = False AndAlso .oMission.MethodIDs(X) = lMethodID Then
                        If .oMissionGoals(X).oGoal Is Nothing = False Then
                            System.BitConverter.GetBytes(.oMissionGoals(X).oGoal.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                        Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                        End If
                        If .oMissionGoals(X).oSkillSet Is Nothing = False Then
                            System.BitConverter.GetBytes(.oMissionGoals(X).oSkillSet.SkillSetID).CopyTo(yMsg, lPos) : lPos += 4
                        Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                        End If


                        yMsg(lPos) = CByte(lAssignCnt(X)) : lPos += 1
                        For Y As Int32 = 0 To .oMissionGoals(X).lAssignmentUB
                            If .oMissionGoals(X).oAssignments(Y) Is Nothing = False Then
                                If .oMissionGoals(X).oAssignments(Y).oAgent Is Nothing = False Then
                                    System.BitConverter.GetBytes(.oMissionGoals(X).oAssignments(Y).oAgent.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                                Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                                End If
                                If .oMissionGoals(X).oAssignments(Y).oSkill Is Nothing = False Then
                                    System.BitConverter.GetBytes(.oMissionGoals(X).oAssignments(Y).oSkill.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
                                Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
                                End If
                            End If
                        Next Y
                    End If
                Next X
            End If
        End With

        MyBase.moUILib.AddNotification("Mission Plan Submitted.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

        Return yMsg
    End Function

    Public Function GetMissionPMID() As Int32
        If moMission Is Nothing = False Then Return moMission.PM_ID
        Return -1
    End Function

    Private Sub btnAssignCover_Click(ByVal sName As String) Handles btnAssignCover.Click
        If lstGoals.ListIndex = -1 Then
            MyBase.moUILib.AddNotification("Select a goal to assign cover agents to first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            Return
        End If

        If mlSelectedAgentID = -1 Then
            MyBase.moUILib.AddNotification("Select an agent to assign and then click Assign.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        Else
            Dim oAgent As Agent = Nothing
            For X As Int32 = 0 To goCurrentPlayer.AgentUB
                If goCurrentPlayer.AgentIdx(X) = mlSelectedAgentID Then
                    oAgent = goCurrentPlayer.Agents(X)
                    Exit For
                End If
            Next X

            'ok, let's see if the agent is assigned to any assignments within the same phase... so we need to get our goal
            Dim oGoal As Goal = Nothing
            Dim lGoalID As Int32 = lstGoals.ItemData(lstGoals.ListIndex)
            If lGoalID = Goal.ml_SAFEHOUSE_GOAL_ID Then
                oGoal = Goal.GetSafehouseGoal()
            Else
                For X As Int32 = 0 To moMission.oMission.GoalUB
                    If moMission.oMission.Goals(X).ObjectID = lGoalID Then
                        oGoal = moMission.oMission.Goals(X)
                        Exit For
                    End If
                Next X
            End If
            If oGoal Is Nothing Then
                MyBase.moUILib.AddNotification("Select a goal to assign cover agents to first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                Return
            End If

            'ok, now, let's see if the agent has been assigned to a skill in this phase
            Dim yPhase As Byte = oGoal.MissionPhase
            For X As Int32 = 0 To moMission.oMission.GoalUB
                If moMission.oMission.MethodIDs(X) = moMission.lMethodID Then
                    If moMission.oMission.Goals(X).MissionPhase = yPhase Then
                        For Y As Int32 = 0 To moMission.oMissionGoals(X).lAssignmentUB
                            If moMission.oMissionGoals(X).oAssignments(Y) Is Nothing = False Then
                                If moMission.oMissionGoals(X).oAssignments(Y).oAgent Is Nothing = False Then
                                    If moMission.oMissionGoals(X).oAssignments(Y).oAgent.ObjectID = oAgent.ObjectID Then
                                        'ok, not good
                                        MyBase.moUILib.AddNotification("This agent has been assigned to a task and cannot be a cover agent.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                                        Return
                                    End If
                                End If
                            End If
                        Next Y
                    End If
                End If
            Next X
            If yPhase = eMissionPhase.ePreparationTime Then
                If oGoal.ObjectID <> Goal.ml_SAFEHOUSE_GOAL_ID Then
                    Dim oPMG As PlayerMissionGoal = moMission.oSafehouseMissionGoal
                    If oPMG Is Nothing = False Then
                        For Y As Int32 = 0 To oPMG.lAssignmentUB
                            If oPMG.oAssignments(Y) Is Nothing = False Then
                                If oPMG.oAssignments(Y).oAgent Is Nothing = False Then
                                    If oPMG.oAssignments(Y).oAgent.ObjectID = oAgent.ObjectID Then
                                        'ok, not good
                                        MyBase.moUILib.AddNotification("This agent has been assigned to a task and cannot be a cover agent.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                                        Return
                                    End If
                                End If
                            End If
                        Next Y
                    End If
                End If
            End If

            'ok, assign this agent
            Dim lPhaseIdx As Int32 = -1
            If moMission.oPhases Is Nothing = False Then
                For X As Int32 = 0 To moMission.lPhaseUB
                    If moMission.oPhases(X) Is Nothing = False AndAlso moMission.oPhases(X).yPhase = yPhase Then
                        lPhaseIdx = X
                        Exit For
                    End If
                Next
            End If
            If lPhaseIdx = -1 Then
                moMission.lPhaseUB = 0
                ReDim moMission.oPhases(moMission.lPhaseUB)
                moMission.oPhases(moMission.lPhaseUB) = New PlayerMissionPhase
                lPhaseIdx = moMission.lPhaseUB
                moMission.oPhases(lPhaseIdx).yPhase = CType(yPhase, eMissionPhase)
                moMission.oPhases(lPhaseIdx).oParent = moMission
            End If
            moMission.oPhases(lPhaseIdx).AddCoverAgent(oAgent, 0)
        End If

        RefreshCoverAgentList()

    End Sub

    Private Sub btnRemoveCover_Click(ByVal sName As String) Handles btnRemoveCover.Click

        If lstCoverAgents.ListIndex > -1 Then
            Dim lAgentID As Int32 = lstCoverAgents.ItemData(lstCoverAgents.ListIndex)
            If lstGoals.ListIndex > -1 Then
                Dim lGoalID As Int32 = lstGoals.ItemData(lstGoals.ListIndex)
                Dim yPhase As Byte = 0
                If moMission Is Nothing = False AndAlso moMission.oMission Is Nothing = False Then

                    Dim oGoal As Goal = Nothing
                    If lGoalID = Goal.ml_SAFEHOUSE_GOAL_ID Then
                        yPhase = eMissionPhase.ePreparationTime
                    Else
                        For X As Int32 = 0 To moMission.oMission.GoalUB
                            If moMission.lMethodID = moMission.oMission.MethodIDs(X) AndAlso moMission.oMission.Goals(X).ObjectID = lGoalID Then
                                yPhase = moMission.oMission.Goals(X).MissionPhase
                                Exit For
                            End If
                        Next X
                    End If

                    For X As Int32 = 0 To moMission.lPhaseUB
                        If moMission.oPhases(X) Is Nothing = False AndAlso moMission.oPhases(X).yPhase = yPhase Then
                            moMission.oPhases(X).RemoveCoverAgent(lAgentID)
                            Exit For
                        End If
                    Next X
                End If
            End If
        End If

        RefreshCoverAgentList()
    End Sub

    Private Sub RefreshCoverAgentList()
        lstCoverAgents.Clear()
        If lstGoals.ListIndex > -1 Then
            Dim lGoalID As Int32 = lstGoals.ItemData(lstGoals.ListIndex)
            Dim yPhase As Byte = 0
            If moMission Is Nothing = False AndAlso moMission.oMission Is Nothing = False Then

                If lGoalID = Goal.ml_SAFEHOUSE_GOAL_ID Then
                    yPhase = eMissionPhase.ePreparationTime
                Else
                    For X As Int32 = 0 To moMission.oMission.GoalUB
                        If moMission.lMethodID = moMission.oMission.MethodIDs(X) AndAlso moMission.oMission.Goals(X).ObjectID = lGoalID Then
                            yPhase = moMission.oMission.Goals(X).MissionPhase
                            Exit For
                        End If
                    Next X
                End If

                Dim oList() As Agent = Nothing
                Dim lListIdx() As Int32 = Nothing
                Dim lListUB As Int32 = -1
                For X As Int32 = 0 To moMission.lPhaseUB
                    If moMission.oPhases(X) Is Nothing = False AndAlso moMission.oPhases(X).yPhase = yPhase Then
                        For Y As Int32 = 0 To moMission.oPhases(X).lCoverAgentUB
                            If moMission.oPhases(X).lCoverAgentIdx(Y) <> -1 Then
                                lListUB += 1
                                ReDim Preserve oList(lListUB)
                                ReDim Preserve lListIdx(lListUB)
                                oList(lListUB) = moMission.oPhases(X).oCoverAgents(Y)
                                lListIdx(lListUB) = moMission.oPhases(X).lCoverAgentIdx(Y)
                            End If
                        Next Y
                        Exit For
                    End If
                Next X

                If lListUB <> -1 Then
                    Dim lSorted() As Int32 = GetSortedIndexArray(oList, lListIdx, GetSortedIndexArrayType.eAgentName)
                    If lSorted Is Nothing = False Then
                        For X As Int32 = 0 To lListUB
                            lstCoverAgents.AddItem(oList(lSorted(X)).sAgentName, False)
                            lstCoverAgents.ItemData(lstCoverAgents.NewIndex) = oList(lSorted(X)).ObjectID
                            lstCoverAgents.ItemData2(lstCoverAgents.NewIndex) = ObjectType.eAgent
                        Next X
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub chkSafehouse_Click() Handles chkSafehouse.Click
        If moMission Is Nothing = False AndAlso moMission.oMission Is Nothing = False Then
            cboMethod_ItemSelected(cboMethod.ListIndex)
        Else
            MyBase.moUILib.AddNotification("Select a mission first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
            chkSafehouse.Value = False
        End If
    End Sub


    Private Sub frmMission_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.AgentMissionCreateX = Me.Left
            muSettings.AgentMissionCreateY = Me.Top
        End If
    End Sub
End Class
#End Region
#Region "new code"
'Public Class frmMission
'    Inherits UIWindow

'    Private Enum eTargetFilterType As Byte
'        eUnusedTarget = 0       'the target is unused
'        ''' <summary>
'        ''' Component belonging to the targeted player
'        ''' </summary>
'        ''' <remarks></remarks>
'        eTarget_Component = 1
'        ''' <summary>
'        ''' Agent belonging to the targeted player
'        ''' </summary>
'        ''' <remarks></remarks>
'        eTarget_Agent
'        ''' <summary>
'        ''' An agent belonging to Current Player that is Captured by Target player
'        ''' </summary>
'        ''' <remarks></remarks>
'        eCaptured_Agent
'        ''' <summary>
'        ''' A known Mineral
'        ''' </summary>
'        ''' <remarks></remarks>
'        eMineral
'        ''' <summary>
'        ''' Production Type list
'        ''' </summary>
'        ''' <remarks></remarks>
'        eProductionType
'        ''' <summary>
'        ''' Known Players
'        ''' </summary>
'        ''' <remarks></remarks>
'        ePlayer
'        ''' <summary>
'        ''' Player Missions (non-state-specific)
'        ''' </summary>
'        ''' <remarks></remarks>
'        ePlayerMissions
'        ''' <summary>
'        ''' Will place an item in the list to select a location... which will change UI state to selecting point
'        ''' </summary>
'        ''' <remarks></remarks>
'        ePICK_A_LOCATION
'        ''' <summary>
'        ''' Known Units or facilities belonging to the target player
'        ''' </summary>
'        ''' <remarks></remarks>
'        eTarget_UnitOrFacility
'        ''' <summary>
'        ''' Known units belonging to the target player
'        ''' </summary>
'        ''' <remarks></remarks>
'        eTarget_Unit
'        ''' <summary>
'        ''' Known facilities belonging to the target player
'        ''' </summary>
'        ''' <remarks></remarks>
'        eTarget_Facility
'        ''' <summary>
'        ''' A solar system
'        ''' </summary>
'        ''' <remarks></remarks>
'        eSolarSystem
'        ''' <summary>
'        ''' A Planet
'        ''' </summary>
'        ''' <remarks></remarks>
'        ePlanet
'        ''' <summary>
'        ''' The current system and all planets within that system will be listed
'        ''' </summary>
'        ''' <remarks></remarks>
'        eSystemAndPlanets
'        ''' <summary>
'        ''' Colony belonging to the target player
'        ''' </summary>
'        ''' <remarks></remarks>
'        eTarget_Colony

'        ''' <summary>
'        ''' Special tech in the target player's arsenal that has not been fully researched
'        ''' </summary>
'        ''' <remarks></remarks>
'        eTargetSpecialTechUnresearched

'        eMineralOrTargetComponent

'        ''' <summary>
'        ''' bit-wise (use AND/OR on value to determine) - indicates that the ANY work will be included in the list
'        ''' </summary>
'        ''' <remarks></remarks>
'        eAny = 128
'    End Enum

'    Private lblTitle As UILabel
'    Private lblMission As UILabel
'    Private lblMissionDesc As UILabel
'    Private lblTarget As UILabel
'    Private lblMethod As UILabel
'    Private lblMethodDesc As UILabel
'    'Private fraGoals As UIWindow
'    Private lblGoalTime As UILabel
'    Private lblGoalRisk As UILabel
'    Private lblGoalName As UILabel
'    Private lblGoalPhase As UILabel
'    'Private txtGoalDesc As UITextBox
'    Private lnDiv1 As UILine
'    Private lblReqInfType As UILabel
'    Private WithEvents fraSkillSets As fraSkillSet
'    Private WithEvents cboMission As UIComboBox
'    Private WithEvents cboTarget1 As UIComboBox
'    Private WithEvents cboTarget2 As UIComboBox
'    Private WithEvents cboTarget3 As UIComboBox
'    Private WithEvents cboMethod As UIComboBox
'    Private WithEvents lstGoals As UIListBox
'    Private WithEvents btnClose As UIButton
'    Private WithEvents btnHelp As UIButton
'    Private WithEvents tvwGoals As UITreeView

'    'Private WithEvents chkSafehouse As UICheckBox

'    Private lstCoverAgents As UIListBox
'    Private WithEvents btnAssignCover As UIButton
'    Private WithEvents btnRemoveCover As UIButton

'    Private WithEvents btnExecuteNow As UIButton
'    Private WithEvents btnSaveForLater As UIButton
'    Private WithEvents btnCancel As UIButton

'    Private WithEvents btnAgentLeft As UIButton
'    Private WithEvents btnAgentRight As UIButton

'    Private mbMouseOver() As Boolean
'    Private moMission As PlayerMission

'    Private Const ml_MOUSE_OVER_DELAY As Int32 = 45     '1.5 seconds
'    Private moAgentPop As frmAgent = Nothing
'    Private mlMouseOverEvent As Int32 = -1

'    Private mlAgentBarCnt As Int32
'    Private mlScrollValue As Int32 = 0
'    Private mlScrollMaxValue As Int32 = 0
'    Private mlSelectedAgentID As Int32 = -1

'    Private mbHasUnknowns As Boolean = False
'    Private mbUseSafehouse As Boolean = False

'    Private myTarget1ListType As eTargetFilterType = eTargetFilterType.ePlayer   'default to player
'    Private myTarget2ListType As eTargetFilterType = eTargetFilterType.eUnusedTarget
'    Private myTarget3ListType As eTargetFilterType = eTargetFilterType.eUnusedTarget

'    Private moFilterSkillSet As SkillSet = Nothing
'    Private mbLoading As Boolean = True

'    Public Sub New(ByRef oUILib As UILib)
'        MyBase.New(oUILib)

'        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenCreateMissionWindow)

'        'frmMission initial props
'        With Me
'            .lWindowMetricID = BPMetrics.eWindow.eCreateMission
'            .ControlName = "frmMission"

'            .Width = 800
'            .Height = 670
'            Dim lLeft As Int32 = -1
'            Dim lTop As Int32 = -1
'            If NewTutorialManager.TutorialOn = False Then
'                lLeft = muSettings.AgentMissionCreateX
'                lTop = muSettings.AgentMissionCreateY
'            End If
'            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 400
'            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 310
'            If lLeft + .Width > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
'            If lTop + .Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
'            .Left = lLeft
'            .Top = lTop

'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceFillColor
'            .FullScreen = True
'            .BorderLineWidth = 2
'        End With

'        'lblTitle initial props
'        lblTitle = New UILabel(oUILib)
'        With lblTitle
'            .ControlName = "lblTitle"
'            .Left = 5
'            .Top = 3
'            .Width = 175
'            .Height = 24
'            .Enabled = True
'            .Visible = True
'            .Caption = "Mission Management"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 12.0F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblTitle, UIControl))

'        'lblMission initial props
'        lblMission = New UILabel(oUILib)
'        With lblMission
'            .ControlName = "lblMission"
'            .Left = 10
'            .Top = 35
'            .Width = 70
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Mission:"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblMission, UIControl))

'        'lblMissionDesc initial props
'        lblMissionDesc = New UILabel(oUILib)
'        With lblMissionDesc
'            .ControlName = "lblMissionDesc"
'            .Left = 310
'            .Top = 35
'            .Width = 490
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = ""
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblMissionDesc, UIControl))

'        'lblTarget initial props
'        lblTarget = New UILabel(oUILib)
'        With lblTarget
'            .ControlName = "lblTarget"
'            .Left = 10
'            .Top = 85
'            .Width = 70
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Target:"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblTarget, UIControl))

'        'lblMethod initial props
'        lblMethod = New UILabel(oUILib)
'        With lblMethod
'            .ControlName = "lblMethod"
'            .Left = 10
'            .Top = 60
'            .Width = 70
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Method:"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblMethod, UIControl))

'        'lblMethodDesc initial props
'        lblMethodDesc = New UILabel(oUILib)
'        With lblMethodDesc
'            .ControlName = "lblMethodDesc"
'            .Left = 310
'            .Top = 60
'            .Width = 490
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = ""
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblMethodDesc, UIControl))

'        'fraGoals initial props
'        'fraGoals = New UIWindow(oUILib)
'        'With fraGoals
'        '    .ControlName = "fraGoals"
'        '    .Left = 10
'        '    .Top = 120
'        '    .Width = 780
'        '    .Height = 125
'        '    .Enabled = False
'        '    .Visible = False
'        '    .BorderColor = muSettings.InterfaceBorderColor
'        '    .FillColor = muSettings.InterfaceFillColor
'        '    .FullScreen = False
'        '    .Moveable = False
'        '    .Caption = "Mission Goals"
'        '    .BorderLineWidth = 2
'        'End With
'        'Me.AddChild(CType(fraGoals, UIControl))

'        'chkSafehouse initial props
'        'chkSafehouse = New UICheckBox(oUILib)
'        'With chkSafehouse
'        '    .ControlName = "chkSafehouse"
'        '    .Left = 80
'        '    .Top = 160
'        '    .Width = 140
'        '    .Height = 18
'        '    .Enabled = True
'        '    .Visible = True
'        '    .Caption = "Create Safehouse"
'        '    .ForeColor = muSettings.InterfaceBorderColor
'        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'        '    .DrawBackImage = False
'        '    .FontFormat = CType(6, DrawTextFormat)
'        '    .Value = False
'        '    .DisplayType = UICheckBox.eCheckTypes.eSmallCheck
'        '    .ToolTipText = "Indicates whether or not to include the create safehouse goal." & vbCrLf & "Safehouses provide bonuses to preparation and re-infiltration."
'        'End With
'        'Me.AddChild(CType(chkSafehouse, UIControl))

'        'lstGoals initial props
'        lstGoals = New UIListBox(oUILib)
'        With lstGoals
'            .ControlName = "lstGoals"
'            .Left = 20
'            .Top = 145
'            .Width = 250
'            .Height = 90
'            .Enabled = True
'            .Visible = False
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'        End With
'        Me.AddChild(CType(lstGoals, UIControl))

'        'lblGoalTime initial props
'        lblGoalTime = New UILabel(oUILib)
'        With lblGoalTime
'            .ControlName = "lblGoalTime"
'            .Left = 310
'            .Top = 110
'            .Width = 250
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Goal Time: Approx. 3 days"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblGoalTime, UIControl))

'        'lblGoalRisk initial props
'        lblGoalRisk = New UILabel(oUILib)
'        With lblGoalRisk
'            .ControlName = "lblGoalRisk"
'            .Left = 310
'            .Top = 135
'            .Width = 250
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Risk Assessment: HIGH"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblGoalRisk, UIControl))

'        'lblGoalPhase initial props
'        lblGoalPhase = New UILabel(oUILib)
'        With lblGoalPhase
'            .ControlName = "lblGoalPhase"
'            .Left = 310
'            .Top = 160
'            .Width = 250
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = ""
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblGoalPhase, UIControl))

'        'lblGoalName initial props
'        lblGoalName = New UILabel(oUILib)
'        With lblGoalName
'            .ControlName = "lblGoalName"
'            .Left = 310
'            .Top = 85
'            .Width = 250
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .Caption = "Goal:"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = CType(4, DrawTextFormat)
'        End With
'        Me.AddChild(CType(lblGoalName, UIControl))

'        'txtGoalDesc initial props
'        'txtGoalDesc = New UITextBox(oUILib)
'        'With txtGoalDesc
'        '    .ControlName = "txtGoalDesc"
'        '    .Left = 540
'        '    .Top = 135
'        '    .Width = 240
'        '    .Height = 100
'        '    .Enabled = True
'        '    .Visible = True
'        '    .Caption = "Goal Description...."
'        '    .ForeColor = muSettings.InterfaceBorderColor
'        '    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'        '    .DrawBackImage = False
'        '    .FontFormat = CType(0, DrawTextFormat)
'        '    .BackColorEnabled = muSettings.InterfaceFillColor
'        '    .MaxLength = 0
'        '    .BorderLineWidth = 0
'        '    .MultiLine = True
'        '    .Locked = True
'        'End With
'        'Me.AddChild(CType(txtGoalDesc, UIControl))

'        'lnDiv1 initial props
'        lnDiv1 = New UILine(oUILib)
'        With lnDiv1
'            .ControlName = "lnDiv1"
'            .Left = 1
'            .Top = 27
'            .Width = 799
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
'            .Left = 774
'            .Top = 2
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

'        'btnHelp initial props
'        btnHelp = New UIButton(oUILib)
'        With btnHelp
'            .ControlName = "btnHelp"
'            .Left = 750
'            .Top = 2
'            .Width = 24
'            .Height = 24
'            .Enabled = True
'            .Visible = True
'            .Caption = "?"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'        End With
'        Me.AddChild(CType(btnHelp, UIControl))

'        'btnAgentLeft initial props
'        btnAgentLeft = New UIButton(oUILib)
'        With btnAgentLeft
'            .ControlName = "btnAgentLeft"
'            .Left = 310
'            .Top = 300 'fraGoals.Top + fraGoals.Height + 54      '(btnHeight \ 2) + (ImageHeight \ 2) + 15
'            .Width = 21
'            .Height = 50
'            .Enabled = True
'            .Visible = True
'            .Caption = ""
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = grc_UI(elInterfaceRectangle.eAgentScrollLeftButton)
'            .ControlImageRect_Disabled = .ControlImageRect
'            .ControlImageRect_Normal = .ControlImageRect
'            .ControlImageRect_Pressed = .ControlImageRect
'        End With
'        Me.AddChild(CType(btnAgentLeft, UIControl))

'        'btnAgentRight initial props
'        btnAgentRight = New UIButton(oUILib)
'        With btnAgentRight
'            .ControlName = "btnAgentRight"
'            .Left = Me.Width - 21 - btnAgentLeft.Left + 260
'            .Top = btnAgentLeft.Top
'            .Width = 21
'            .Height = 50
'            .Enabled = True
'            .Visible = True
'            .Caption = ""
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = grc_UI(elInterfaceRectangle.eAgentScrollRightButton)
'            .ControlImageRect_Disabled = .ControlImageRect
'            .ControlImageRect_Normal = .ControlImageRect
'            .ControlImageRect_Pressed = .ControlImageRect
'        End With
'        Me.AddChild(CType(btnAgentRight, UIControl))

'        'fraSkillSets initial props
'        fraSkillSets = New fraSkillSet(oUILib)
'        With fraSkillSets
'            .ControlName = "fraSkillSets"
'            .Left = 15 'fraGoals.Left
'            .Top = 420
'            '.Width = fraGoals.Width
'            .Visible = True
'            .Enabled = True
'            .Caption = "Skillset Assignments"
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceFillColor
'            .FullScreen = False
'            .Moveable = False
'            .BorderLineWidth = 2
'            .bRoundedBorder = True
'        End With
'        Me.AddChild(CType(fraSkillSets, UIControl))

'        'btnAssignCover initial props
'        btnAssignCover = New UIButton(oUILib)
'        With btnAssignCover
'            .ControlName = "btnAssignCover"
'            .Left = fraSkillSets.Left + fraSkillSets.Width + 5
'            .Top = fraSkillSets.Top
'            .Width = Me.Width - .Left - 5
'            .Height = 28
'            .Enabled = True
'            .Visible = True
'            .Caption = "Assign to Cover"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'        End With
'        Me.AddChild(CType(btnAssignCover, UIControl))

'        'lstCoverAgents initial props
'        lstCoverAgents = New UIListBox(oUILib)
'        With lstCoverAgents
'            .ControlName = "lstCoverAgents"
'            .Left = btnAssignCover.Left
'            .Top = btnAssignCover.Top + btnAssignCover.Height + 5
'            .Width = btnAssignCover.Width
'            .Height = fraSkillSets.Height - (btnAssignCover.Height * 2) - 10
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'        End With
'        Me.AddChild(CType(lstCoverAgents, UIControl))

'        'btnRemoveCover initial props
'        btnRemoveCover = New UIButton(oUILib)
'        With btnRemoveCover
'            .ControlName = "btnRemoveCover"
'            .Left = btnAssignCover.Left
'            .Top = lstCoverAgents.Top + lstCoverAgents.Height + 5
'            .Width = btnAssignCover.Width
'            .Height = btnAssignCover.Height
'            .Enabled = True
'            .Visible = True
'            .Caption = "Remove Assignment"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'        End With
'        Me.AddChild(CType(btnRemoveCover, UIControl))

'        'btnSaveForLater initial props
'        btnSaveForLater = New UIButton(oUILib)
'        With btnSaveForLater
'            .ControlName = "btnSaveForLater"
'            .Left = fraSkillSets.Left + fraSkillSets.Width - 100
'            .Top = fraSkillSets.Top + fraSkillSets.Height + 5
'            .Width = 100
'            .Height = 24
'            .Enabled = True
'            .Visible = True
'            .Caption = "Execute Later"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'            .ToolTipText = "The mission is planned but will not execute until Execute Now is pressed."
'        End With
'        Me.AddChild(CType(btnSaveForLater, UIControl))

'        'btnExecuteNow initial props
'        btnExecuteNow = New UIButton(oUILib)
'        With btnExecuteNow
'            .ControlName = "btnExecuteNow"
'            .Left = btnSaveForLater.Left - 105
'            .Top = btnSaveForLater.Top
'            .Width = 100
'            .Height = 24
'            .Enabled = True
'            .Visible = True
'            .Caption = "Execute Now"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'            .ToolTipText = "The mission will execute as soon as all agents are infiltrated and in place."
'        End With
'        Me.AddChild(CType(btnExecuteNow, UIControl))

'        'btnCancel initial props
'        btnCancel = New UIButton(oUILib)
'        With btnCancel
'            .ControlName = "btnCancel"
'            .Left = fraSkillSets.Left
'            .Top = btnSaveForLater.Top
'            .Width = 130
'            .Height = 24
'            .Enabled = True
'            .Visible = True
'            .Caption = "Cancel Mission"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = True
'            .FontFormat = CType(5, DrawTextFormat)
'            .ControlImageRect = New Rectangle(0, 0, 120, 32)
'            .ToolTipText = "Cancels the mission if the mission can be cancelled." & vbCrLf & _
'                           "Once missions reach a certain phase of execution," & vbCrLf & _
'                           "they can no longer be cancelled."
'        End With
'        Me.AddChild(CType(btnCancel, UIControl))

'        'lblReqInfType initial props
'        lblReqInfType = New UILabel(oUILib)
'        With lblReqInfType
'            .ControlName = "lblReqInfType"
'            .Left = btnCancel.Left + btnCancel.Width + 10
'            .Top = btnCancel.Top
'            .Width = btnExecuteNow.Left - .Left
'            .Height = btnCancel.Height
'            .Enabled = True
'            .Visible = True
'            .Caption = "Required Infiltration: Select Mission"
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
'            .DrawBackImage = False
'            .FontFormat = DrawTextFormat.Left Or DrawTextFormat.VerticalCenter
'        End With
'        Me.AddChild(CType(lblReqInfType, UIControl))

'        'cboMethod initial props
'        cboMethod = New UIComboBox(oUILib)
'        With cboMethod
'            .ControlName = "cboMethod"
'            .Left = 80
'            .Top = 60
'            .Width = 220
'            .Height = 18
'            .Enabled = False
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceTextBoxFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'            .DropDownListBorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(cboMethod, UIControl))

'        'cboTarget3 initial props
'        cboTarget3 = New UIComboBox(oUILib)
'        With cboTarget3
'            .ControlName = "cboTarget3"
'            .Left = 80
'            .Top = 135
'            .Width = 220
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceTextBoxFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'            .DropDownListBorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(cboTarget3, UIControl))

'        'cboTarget2 initial props
'        cboTarget2 = New UIComboBox(oUILib)
'        With cboTarget2
'            .ControlName = "cboTarget2"
'            .Left = 80
'            .Top = 110
'            .Width = 220
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceTextBoxFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'            .DropDownListBorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(cboTarget2, UIControl))

'        'cboTarget1 initial props
'        cboTarget1 = New UIComboBox(oUILib)
'        With cboTarget1
'            .ControlName = "cboTarget1"
'            .Left = 80
'            .Top = 85
'            .Width = 220
'            .Height = 18
'            .Enabled = False
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceTextBoxFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'            .DropDownListBorderColor = muSettings.InterfaceBorderColor
'        End With
'        Me.AddChild(CType(cboTarget1, UIControl))

'        'cboMission initial props
'        cboMission = New UIComboBox(oUILib)
'        With cboMission
'            .ControlName = "cboMission"
'            .Left = 80
'            .Top = 35
'            .Width = 220
'            .Height = 18
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceTextBoxFillColor
'            .ForeColor = muSettings.InterfaceBorderColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'            .DropDownListBorderColor = muSettings.InterfaceBorderColor
'            .l_ListBoxHeight = 200
'        End With
'        Me.AddChild(CType(cboMission, UIControl))

'        'tvwGoals initial props
'        tvwGoals = New UITreeView(oUILib)
'        With tvwGoals
'            .ControlName = "tvwGoals"
'            .Left = 80
'            .Top = 180
'            .Width = 220
'            .Height = 220
'            .Enabled = True
'            .Visible = True
'            .BorderColor = muSettings.InterfaceBorderColor
'            .FillColor = muSettings.InterfaceTextBoxFillColor
'            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
'            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
'            .ForeColor = muSettings.InterfaceBorderColor
'            .mbAcceptReprocessEvents = True
'        End With
'        Me.AddChild(CType(tvwGoals, UIControl))

'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'        MyBase.moUILib.AddWindow(Me)

'        FillMissionList()
'        mbLoading = False
'    End Sub

'    Public Sub SetFromMission(ByRef oMission As PlayerMission, ByVal bCreateNew As Boolean)
'        If oMission Is Nothing Then
'            moMission = New PlayerMission()
'            moMission.PM_ID = -1
'        Else
'            moMission = oMission.CreateClone()           'now, moMission is a workable copy
'            moMission.PM_ID = -1
'        End If
'        If moMission.oMission Is Nothing = False Then cboMission.FindComboItemData(moMission.oMission.ObjectID)
'        cboMission_ItemSelected(cboMission.ListIndex)
'        cboTarget1.FindComboItemData(moMission.lTargetPlayerID)
'        cboTarget1_ItemSelected(cboTarget1.ListIndex)

'        If cboTarget2.Visible = True Then
'            cboTarget2.ListIndex = -1
'            For X As Int32 = 0 To cboTarget2.ListCount - 1
'                If cboTarget2.ItemData(X) = moMission.lTargetID AndAlso cboTarget2.ItemData2(X) = moMission.iTargetTypeID Then
'                    cboTarget2.ListIndex = X
'                    Exit For
'                End If
'            Next X
'        End If
'        If cboTarget3.Visible = True Then
'            cboTarget3.ListIndex = -1
'            For X As Int32 = 0 To cboTarget3.ListCount - 1
'                If cboTarget3.ItemData(X) = moMission.lTargetID2 AndAlso cboTarget3.ItemData2(X) = moMission.iTargetTypeID2 Then
'                    cboTarget3.ListIndex = X
'                    Exit For
'                End If
'            Next X
'        End If

'        cboMethod.FindComboItemData(moMission.lMethodID)
'        cboMethod_ItemSelected(cboMethod.ListIndex)

'        cboMission.Enabled = moMission.PM_ID < 1

'    End Sub

'    Private Sub FillMissionList()
'        Dim lSorted() As Int32 = GetSortedIndexArray(goMissions, glMissionIdx, GetSortedIndexArrayType.eMissionName)

'        'Fill our mission list
'        cboMission.Clear()
'        If lSorted Is Nothing = False Then
'            For X As Int32 = 0 To lSorted.GetUpperBound(0)

'                If NewTutorialManager.TutorialOn = True AndAlso goCurrentPlayer.yPlayerPhase = 0 Then
'                    If glMissionIdx(lSorted(X)) = 999 Then
'                        cboMission.AddItem(goMissions(lSorted(X)).sMissionName)
'                        cboMission.ItemData(cboMission.NewIndex) = goMissions(lSorted(X)).ObjectID
'                    End If
'                Else
'                    If glMissionIdx(lSorted(X)) <> -1 AndAlso glMissionIdx(lSorted(X)) <> 999 Then
'                        If goMissions(lSorted(X)).ProgramControlID = eMissionResult.eLocateWormhole Then
'                            If goCurrentPlayer Is Nothing OrElse (goCurrentPlayer.GetSpecialTechPropertyValue(PlayerSpecialAttributeID.eSuperSpecials) And 4) = 0 Then
'                                Continue For
'                            End If
'                        End If
'                        cboMission.AddItem(goMissions(lSorted(X)).sMissionName)
'                        cboMission.ItemData(cboMission.NewIndex) = goMissions(lSorted(X)).ObjectID
'                    End If
'                End If
'            Next X
'        End If
'    End Sub

'    Private Sub frmMission_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
'        If mbMouseOver Is Nothing = False Then
'            For X As Int32 = 0 To mbMouseOver.GetUpperBound(0)
'                If mbMouseOver(X) = True Then
'                    'ok, select that item
'                    'If mlScrollValue + X > goCurrentPlayer.AgentUB Then Return
'                    Dim lIdx As Int32 = GetScrolledSelectionIndex(X)
'                    If lIdx > -1 AndAlso lIdx <= goCurrentPlayer.AgentUB Then AgentSelected(goCurrentPlayer.Agents(lIdx))
'                    Exit For
'                End If
'            Next X
'        End If
'    End Sub

'    Private Function GetScrolledSelectionIndex(ByVal lSelectionIndex As Int32) As Int32
'        Dim lTempIdx As Int32 = -1

'        Dim lAgentListUB As Int32 = -1
'        Dim lAgentListIdx() As Int32 = Nothing
'        Dim lCnt As Int32 = 0
'        Dim lStartIdx As Int32 = -1
'        Dim lSorted() As Int32 = GetSortedIndexArray(goCurrentPlayer.Agents, goCurrentPlayer.AgentIdx, GetSortedIndexArrayType.eAgentName)
'        If lSorted Is Nothing = False Then
'            For X As Int32 = 0 To lSorted.GetUpperBound(0)
'                lTempIdx = lSorted(X)
'                If goCurrentPlayer.AgentIdx(lTempIdx) <> -1 AndAlso (goCurrentPlayer.Agents(lTempIdx).lAgentStatus And (AgentStatus.NewRecruit Or AgentStatus.IsDead Or AgentStatus.Dismissed Or AgentStatus.HasBeenCaptured)) = 0 Then
'                    Dim oAgent As Agent = goCurrentPlayer.Agents(lTempIdx)
'                    If oAgent Is Nothing = False Then
'                        Dim bExclude As Boolean = False
'                        If moFilterSkillSet Is Nothing = False Then
'                            bExclude = True
'                            For Y As Int32 = 0 To moFilterSkillSet.SkillUB
'                                Dim bSkillIsAbility As Boolean = moFilterSkillSet.Skills(Y).oSkill.SkillType <> 0
'                                For Z As Int32 = 0 To oAgent.SkillUB
'                                    If oAgent.Skills(Z).ObjectID = moFilterSkillSet.Skills(Y).oSkill.ObjectID Then
'                                        bExclude = False
'                                        Exit For
'                                    ElseIf oAgent.Skills(Z).ObjectID = 36 AndAlso bSkillIsAbility = False Then          '36 is Naturally Talented ID
'                                        bExclude = False
'                                        Exit For
'                                    End If
'                                Next Z
'                                If bExclude = False Then Exit For
'                            Next Y
'                        End If

'                        If bExclude = False Then
'                            lAgentListUB += 1
'                            ReDim Preserve lAgentListIdx(lAgentListUB)
'                            lAgentListIdx(lAgentListUB) = lTempIdx

'                            mlScrollMaxValue += 1
'                            lCnt += 1
'                            If lCnt = mlScrollValue + 1 AndAlso lStartIdx = -1 Then
'                                lStartIdx = lAgentListUB
'                            End If
'                        End If
'                    End If
'                End If
'            Next X
'        End If
'        lTempIdx = -1

'        For X As Int32 = 0 To lAgentListUB
'            Dim oAgent As Agent = goCurrentPlayer.Agents(lAgentListIdx(X))
'            If oAgent Is Nothing = False Then
'                lTempIdx += 1
'                If lTempIdx = mlScrollValue + lSelectionIndex Then Return lAgentListIdx(X)
'            End If
'        Next X
'        Return -1
'    End Function

'    Private Sub frmMission_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
'        Dim oLoc As Point = Me.GetAbsolutePosition()
'        lMouseY -= oLoc.Y
'        lMouseX -= oLoc.X
'        Dim bMouseOverAgent As Boolean = False
'        If lMouseX > 335 Then ' 50 Then
'            If lMouseY > 260 AndAlso lMouseY < 404 Then

'                'rcDest.X += 130
'                lMouseX -= 335 '50
'                Dim lItem As Int32 = lMouseX \ 130
'                If lItem < mlAgentBarCnt Then
'                    For X As Int32 = 0 To mlAgentBarCnt - 1
'                        mbMouseOver(X) = False
'                    Next X

'                    bMouseOverAgent = True

'                    If mlMouseOverEvent = -1 Then
'                        mlMouseOverEvent = glCurrentCycle
'                    ElseIf moAgentPop Is Nothing = False Then
'                        Dim lIdx As Int32 = GetScrolledSelectionIndex(lItem)
'                        If lIdx > -1 AndAlso lIdx <= goCurrentPlayer.AgentUB Then moAgentPop.SetFromAgent(goCurrentPlayer.Agents(lIdx))
'                        'If mlScrollValue + lItem < goCurrentPlayer.AgentUB Then moAgentPop.SetFromAgent(goCurrentPlayer.Agents(mlScrollValue + lItem))
'                    End If
'                    mbMouseOver(lItem) = True
'                    Me.IsDirty = True
'                End If
'            End If
'        End If

'        If bMouseOverAgent = False Then
'            For X As Int32 = 0 To mlAgentBarCnt - 1
'                If mbMouseOver(X) = True Then
'                    mbMouseOver(X) = False
'                    Me.IsDirty = True
'                    mlMouseOverEvent = -1
'                End If
'            Next X
'        End If
'    End Sub

'    Private Sub AgentSelected(ByRef oAgent As Agent)
'        mlSelectedAgentID = -1
'        If oAgent Is Nothing = False Then
'            mlSelectedAgentID = oAgent.ObjectID
'            If NewTutorialManager.TutorialOn = True Then
'                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionSelectAgent, oAgent.ObjectID, -1, -1, "")
'            End If
'            Me.IsDirty = True
'        End If
'    End Sub

'    Private Sub frmMission_OnNewFrame() Handles Me.OnNewFrame

'        If cboTarget1 Is Nothing = False Then
'            With cboTarget1
'                For X As Int32 = 0 To .ListCount - 1
'                    If .ItemData2(X) = ObjectType.ePlayer Then
'                        Dim sText As String = GetCacheObjectValue(.ItemData(X), ObjectType.ePlayer)
'                        If .List(X) <> sText Then .List(X) = sText
'                    ElseIf .ItemData2(X) = ObjectType.eAgent Then
'                        If .List(X).ToUpper = "UNKNOWN" Then
'                            Dim sText As String = GetCacheObjectValue(.ItemData(X), ObjectType.eAgent)
'                            If .List(X) <> sText Then .List(X) = sText
'                        End If
'                    End If
'                Next X
'            End With
'        End If
'        If cboTarget2 Is Nothing = False Then
'            With cboTarget2
'                For X As Int32 = 0 To .ListCount - 1
'                    If .ItemData2(X) = ObjectType.ePlayer Then
'                        Dim sText As String = GetCacheObjectValue(.ItemData(X), ObjectType.ePlayer)
'                        If .List(X) <> sText Then .List(X) = sText
'                    ElseIf .ItemData2(X) = ObjectType.eAgent Then
'                        If .List(X).ToUpper = "UNKNOWN" Then
'                            Dim sText As String = GetCacheObjectValue(.ItemData(X), ObjectType.eAgent)
'                            If .List(X) <> sText Then .List(X) = sText
'                        End If
'                    End If
'                Next X
'            End With
'        End If
'        If cboTarget3 Is Nothing = False Then
'            With cboTarget3
'                For X As Int32 = 0 To .ListCount - 1
'                    If .ItemData2(X) = ObjectType.ePlayer Then
'                        Dim sText As String = GetCacheObjectValue(.ItemData(X), ObjectType.ePlayer)
'                        If .List(X) <> sText Then .List(X) = sText
'                    ElseIf .ItemData2(X) = ObjectType.eAgent Then
'                        If .List(X).ToUpper = "UNKNOWN" Then
'                            Dim sText As String = GetCacheObjectValue(.ItemData(X), ObjectType.eAgent)
'                            If .List(X) <> sText Then .List(X) = sText
'                        End If
'                    End If
'                Next X
'            End With
'        End If

'        If mlMouseOverEvent = -1 Then
'            If moAgentPop Is Nothing = False Then
'                RemoveHandler moAgentPop.AgentSelected, AddressOf AgentSelected
'                MyBase.moUILib.RemoveWindow(moAgentPop.ControlName)
'                moAgentPop = Nothing
'            End If
'        ElseIf glCurrentCycle - mlMouseOverEvent > ml_MOUSE_OVER_DELAY Then
'            'figure out which item we are hovered over
'            Dim lIdx As Int32 = -1
'            Dim lAgentIdx As Int32 = -1
'            For X As Int32 = 0 To mlAgentBarCnt - 1
'                If mbMouseOver(X) = True Then
'                    lIdx = X
'                    Exit For
'                End If
'            Next X
'            If lIdx = -1 Then
'                mlMouseOverEvent = -1
'                Return
'            Else
'                lAgentIdx = GetScrolledSelectionIndex(lIdx)
'                If lAgentIdx < 0 OrElse lAgentIdx > goCurrentPlayer.AgentUB Then
'                    mlMouseOverEvent = -1
'                    Return
'                End If
'            End If

'            'ok, is moAgentPop nothing?
'            If moAgentPop Is Nothing = True Then
'                moAgentPop = New frmAgent(goUILib, True, False)
'                AddHandler moAgentPop.AgentSelected, AddressOf AgentSelected
'            End If
'            Dim lTemp As Int32 = 289 + ((lIdx) * 130) + Me.Left  '4
'            If moAgentPop.Left <> lTemp Then moAgentPop.Left = lTemp
'            lTemp = 245 + Me.Top + 15 - 249
'            If moAgentPop.Top <> lTemp Then moAgentPop.Top = lTemp
'            If moAgentPop.Visible = False Then moAgentPop.Visible = True
'            moAgentPop.SetAgentIfNeeded(goCurrentPlayer.Agents(lAgentIdx))
'        End If

'        If mbHasUnknowns = True Then Me.IsDirty = True
'    End Sub

'    Private Sub frmMission_OnRender() Handles Me.OnRender
'        mlAgentBarCnt = 3
'        If mbMouseOver Is Nothing OrElse mbMouseOver.GetUpperBound(0) + 1 <> mlAgentBarCnt Then ReDim mbMouseOver(mlAgentBarCnt - 1)

'        If goCurrentPlayer Is Nothing = True Then Return
'        If goCurrentPlayer.AgentUB = -1 Then Return
'        'If fraGoals Is Nothing Then Return

'        'get our total scroll count
'        Dim lStartIdx As Int32 = -1
'        Dim lCnt As Int32 = 0

'        Dim lAgentListIdx() As Int32 = Nothing
'        Dim lAgentListUB As Int32 = -1
'        mlScrollMaxValue = 0

'        Dim lSorted() As Int32 = GetSortedIndexArray(goCurrentPlayer.Agents, goCurrentPlayer.AgentIdx, GetSortedIndexArrayType.eAgentName)
'        If lSorted Is Nothing = False Then
'            For X As Int32 = 0 To lSorted.GetUpperBound(0)
'                Dim lTempIdx As Int32 = lSorted(X)
'                If goCurrentPlayer.AgentIdx(lTempIdx) <> -1 Then
'                    Dim oAgent As Agent = goCurrentPlayer.Agents(lTempIdx)
'                    If oAgent Is Nothing = False Then
'                        If (oAgent.lAgentStatus And (AgentStatus.NewRecruit Or AgentStatus.IsDead Or AgentStatus.Dismissed Or AgentStatus.HasBeenCaptured)) <> 0 Then Continue For

'                        Dim bExclude As Boolean = False
'                        If moFilterSkillSet Is Nothing = False Then
'                            bExclude = True
'                            For Y As Int32 = 0 To moFilterSkillSet.SkillUB
'                                Dim bSkillIsAbility As Boolean = moFilterSkillSet.Skills(Y).oSkill.SkillType <> 0
'                                For Z As Int32 = 0 To oAgent.SkillUB
'                                    If oAgent.Skills(Z).ObjectID = moFilterSkillSet.Skills(Y).oSkill.ObjectID Then
'                                        bExclude = False
'                                        Exit For
'                                    ElseIf oAgent.Skills(Z).ObjectID = 36 AndAlso bSkillIsAbility = False Then      'naturally talented
'                                        bExclude = False
'                                        Exit For
'                                    End If
'                                Next Z
'                                If bExclude = False Then Exit For
'                            Next Y
'                        End If

'                        If bExclude = False Then
'                            lAgentListUB += 1
'                            ReDim Preserve lAgentListIdx(lAgentListUB)
'                            lAgentListIdx(lAgentListUB) = lTempIdx

'                            mlScrollMaxValue += 1
'                            lCnt += 1
'                            If lCnt = mlScrollValue + 1 AndAlso lStartIdx = -1 Then
'                                lStartIdx = lAgentListUB
'                            End If
'                        End If
'                    End If
'                End If
'            Next X
'        End If

'        'mlScrollMaxValue = 0
'        'For X As Int32 = 0 To goCurrentPlayer.AgentUB
'        '	If goCurrentPlayer.AgentIdx(X) <> -1 AndAlso (goCurrentPlayer.Agents(X).lAgentStatus And (AgentStatus.NewRecruit Or AgentStatus.IsDead Or AgentStatus.Dismissed)) = 0 Then
'        '		mlScrollMaxValue += 1
'        '		lCnt += 1
'        '		If lCnt = mlScrollValue + 1 AndAlso lStartIdx = -1 Then
'        '			lStartIdx = X
'        '		End If
'        '	End If
'        'Next X
'        mlScrollMaxValue -= mlAgentBarCnt
'        If mlScrollMaxValue < 0 Then mlScrollMaxValue = 0
'        If lStartIdx = -1 Then lStartIdx = 0
'        If mlScrollValue > mlScrollMaxValue Then mlScrollValue = mlScrollMaxValue

'        If AgentRenderer.goAgentRenderer Is Nothing Then AgentRenderer.goAgentRenderer = New AgentRenderer()

'        AgentRenderer.goAgentRenderer.PrepareForBatch()
'        Dim rcDest As Rectangle
'        With rcDest
'            .X = 335 '50
'            .Y = 260 '+ Me.Top
'            .Width = 128
'            .Height = 128
'        End With

'        lCnt = 0
'        Dim lSelectedIndex As Int32 = -1
'        If mlSelectedAgentID <> -1 Then
'            lCnt = 0
'            Dim rcSelectBox As Rectangle = rcDest
'            For X As Int32 = lStartIdx To lAgentListUB ' goCurrentPlayer.AgentUB
'                Dim oAgent As Agent = goCurrentPlayer.Agents(lAgentListIdx(X))
'                If oAgent Is Nothing = False AndAlso (oAgent.lAgentStatus And (AgentStatus.NewRecruit Or AgentStatus.IsDead Or AgentStatus.Dismissed)) = 0 Then
'                    If oAgent.ObjectID = mlSelectedAgentID Then
'                        'Ok, found it
'                        rcSelectBox.X -= 2
'                        rcSelectBox.Y -= 2
'                        rcSelectBox.Width += 4
'                        rcSelectBox.Height += 4
'                        lSelectedIndex = X
'                        RenderBox(rcSelectBox, 4, Color.White)
'                        Exit For
'                    End If
'                    rcSelectBox.X += 130
'                    lCnt += 1
'                    If lCnt = mlAgentBarCnt Then Exit For
'                End If
'            Next X
'        End If

'        lCnt = 0
'        For X As Int32 = lStartIdx To lAgentListUB 'goCurrentPlayer.AgentUB
'            Dim oAgent As Agent = goCurrentPlayer.Agents(lAgentListIdx(X))
'            If oAgent Is Nothing = False AndAlso (oAgent.lAgentStatus And (AgentStatus.NewRecruit Or AgentStatus.IsDead Or AgentStatus.Dismissed)) = 0 Then
'                If lSelectedIndex = X Then
'                    AgentRenderer.goAgentRenderer.RenderAgent2(oAgent.ObjectID, rcDest, False, oAgent.bMale)
'                Else
'                    AgentRenderer.goAgentRenderer.RenderAgent2(oAgent.ObjectID, rcDest, Not mbMouseOver(lCnt), oAgent.bMale)
'                End If
'                rcDest.X += 130
'                lCnt += 1
'                If lCnt = mlAgentBarCnt Then Exit For
'            End If
'        Next X

'        mbHasUnknowns = AgentRenderer.goAgentRenderer.bHasUnknowns
'        AgentRenderer.goAgentRenderer.FinishBatch()

'        rcDest.X = 335
'        rcDest.Y = 245 + 15 + 130
'        rcDest.Width = 128
'        rcDest.Height = 18

'        If GFXEngine.gbPaused = True OrElse GFXEngine.gbDeviceLost = True Then Return

'        Using oFont As Font = New Font(MyBase.moUILib.oDevice, lblMission.GetFont())
'            Using oTextSpr As Sprite = New Sprite(MyBase.moUILib.oDevice)
'                oTextSpr.Begin(SpriteFlags.AlphaBlend)

'                Try
'                    lCnt = 0
'                    For X As Int32 = lStartIdx To lAgentListUB 'goCurrentPlayer.AgentUB
'                        Dim oAgent As Agent = goCurrentPlayer.Agents(lAgentListIdx(X))
'                        If oAgent Is Nothing = False AndAlso (oAgent.lAgentStatus And (AgentStatus.NewRecruit Or AgentStatus.IsDead Or AgentStatus.Dismissed)) = 0 Then
'                            oFont.DrawText(oTextSpr, oAgent.sAgentName, rcDest, DrawTextFormat.Center Or DrawTextFormat.VerticalCenter, Color.White)
'                            rcDest.X += 130
'                            lCnt += 1
'                            If lCnt = mlAgentBarCnt Then Exit For
'                        End If
'                    Next X
'                Catch
'                End Try

'                oTextSpr.End()
'                oTextSpr.Dispose()
'            End Using
'            oFont.Dispose()
'        End Using

'    End Sub

'    Protected Overrides Sub Finalize()
'        AgentRenderer.goAgentRenderer = Nothing
'        MyBase.Finalize()
'    End Sub

'    Private Sub cboMission_ItemSelected(ByVal lItemIndex As Integer) Handles cboMission.ItemSelected
'        lblMissionDesc.Caption = "" : lblMethodDesc.Caption = ""
'        cboTarget1.ListIndex = -1 : cboTarget1.Clear()
'        cboTarget2.ListIndex = -1 : cboTarget2.Enabled = False
'        cboTarget3.ListIndex = -1 : cboTarget3.Enabled = False
'        cboTarget2.Top = 110 : cboTarget3.Top = 135

'        Dim lMethodID As Int32 = -1
'        If cboMethod.ListIndex <> -1 Then lMethodID = cboMethod.ItemData(cboMethod.ListIndex)
'        cboMethod.ListIndex = -1 : cboMethod.Clear()

'        tvwGoals.Clear()
'        lblGoalName.Caption = ""
'        lblGoalTime.Caption = ""
'        lblGoalRisk.Caption = ""
'        'txtGoalDesc.Caption = "" ' OLD
'        lblGoalPhase.Caption = ""

'        If cboMission.ListIndex <> -1 Then
'            Dim oMission As Mission = Nothing
'            Dim lID As Int32 = cboMission.ItemData(cboMission.ListIndex)
'            For X As Int32 = 0 To glMissionUB
'                If glMissionIdx(X) = lID Then
'                    oMission = goMissions(X)
'                    Exit For
'                End If
'            Next X
'            If oMission Is Nothing = False Then

'                If NewTutorialManager.TutorialOn = True Then
'                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionMissionSelected, lID, -1, -1, "")
'                End If

'                If moMission.oMission Is Nothing = True OrElse moMission.oMission.ObjectID <> oMission.ObjectID Then
'                    moMission.oMission = oMission

'                    ReDim moMission.oMissionGoals(oMission.GoalUB)
'                    For X As Int32 = 0 To oMission.GoalUB
'                        moMission.oMissionGoals(X) = New PlayerMissionGoal()
'                        moMission.oMissionGoals(X).oGoal = oMission.Goals(X)
'                        moMission.oMissionGoals(X).oMission = moMission
'                    Next X
'                End If

'                Dim sMDesc As String = oMission.sMissionDesc
'                Dim lVBCRLFLoc As Int32 = sMDesc.IndexOf(vbCrLf)
'                If lVBCRLFLoc > 0 Then
'                    lVBCRLFLoc = Math.Min(lVBCRLFLoc, 82)
'                Else : lVBCRLFLoc = Math.Min(82, sMDesc.Length)
'                End If
'                If lVBCRLFLoc <> sMDesc.Length Then
'                    sMDesc = sMDesc.Substring(0, lVBCRLFLoc) & "..."
'                End If
'                lblMissionDesc.Caption = sMDesc
'                lblMissionDesc.ToolTipText = oMission.sMissionDesc & vbCrLf & "Required Infiltration Type: " & frmAgent.fraInfiltration.GetInfTypeText(oMission.lInfiltrationType)

'                lblReqInfType.Caption = "Required Infiltration: " & frmAgent.fraInfiltration.GetInfTypeText(oMission.lInfiltrationType)

'                'ok, determine what boxes to highlight (enable). for mission results 32, 39, 40 and 41 the first box is unused
'                myTarget1ListType = eTargetFilterType.ePlayer   'default to player
'                myTarget2ListType = eTargetFilterType.eUnusedTarget
'                myTarget3ListType = eTargetFilterType.eUnusedTarget

'                'Now, determine the remaining comboboxes
'                Select Case CType(oMission.ProgramControlID, eMissionResult)
'                    Case eMissionResult.eAcquireTechData
'                        myTarget2ListType = eTargetFilterType.eTarget_Component Or eTargetFilterType.eAny
'                    Case eMissionResult.eAssassinateAgent, eMissionResult.eCaptureAgent
'                        myTarget2ListType = eTargetFilterType.eTarget_Agent Or eTargetFilterType.eAny
'                    Case eMissionResult.eClutterCargoBay, eMissionResult.eIncreasePowerNeeds
'                        myTarget2ListType = eTargetFilterType.eTarget_Colony
'                    Case eMissionResult.eDoorJamHangar
'                        myTarget2ListType = eTargetFilterType.eTarget_UnitOrFacility Or eTargetFilterType.eAny
'                    Case eMissionResult.eFindMineral
'                        'make target1 invisible
'                        cboTarget1.Visible = False
'                        'shift target2 and 3 up one spot
'                        cboTarget2.Top = 85 : cboTarget3.Top = 110
'                        myTarget1ListType = eTargetFilterType.eUnusedTarget
'                        myTarget2ListType = eTargetFilterType.eMineral
'                        myTarget3ListType = eTargetFilterType.eSystemAndPlanets ' eTargetFilterType.eSolarSystem
'                    Case eMissionResult.eAudit
'                        cboTarget1.Visible = False
'                        cboTarget2.Visible = False
'                        cboTarget3.Visible = False
'                        myTarget1ListType = eTargetFilterType.eUnusedTarget
'                        myTarget2ListType = eTargetFilterType.eUnusedTarget
'                    Case eMissionResult.eGeologicalSurvey, eMissionResult.eTutorialFindFactory
'                        'make target1 invisible
'                        cboTarget1.Visible = False
'                        'shift target2 and 3 up one spot
'                        cboTarget2.Top = 85 : cboTarget3.Top = 110
'                        myTarget1ListType = eTargetFilterType.eUnusedTarget
'                        myTarget2ListType = eTargetFilterType.ePlanet
'                    Case eMissionResult.eImpedeCurrentDevelopment, eMissionResult.eDestroyCurrentSpecialProject
'                        myTarget2ListType = eTargetFilterType.eTargetSpecialTechUnresearched
'                    Case eMissionResult.eGetFacilityList
'                        myTarget2ListType = eTargetFilterType.eProductionType
'                        myTarget3ListType = eTargetFilterType.eSystemAndPlanets
'                    Case eMissionResult.ePlantEvidence
'                        myTarget2ListType = eTargetFilterType.ePlayer
'                        myTarget3ListType = eTargetFilterType.ePlayerMissions
'                    Case eMissionResult.eReconPlanetMap
'                        'make target1 invisible
'                        cboTarget1.Visible = False
'                        'shift target2 and 3 up one spot
'                        cboTarget2.Top = 85 : cboTarget3.Top = 110
'                        myTarget1ListType = eTargetFilterType.eUnusedTarget
'                        myTarget2ListType = eTargetFilterType.ePlanet
'                        myTarget3ListType = eTargetFilterType.ePICK_A_LOCATION
'                    Case eMissionResult.eSearchAndRescueAgent
'                        'make target1 invisible
'                        cboTarget1.Visible = False
'                        'shift target2 and 3 up one spot
'                        cboTarget2.Top = 85 : cboTarget3.Top = 110
'                        myTarget1ListType = eTargetFilterType.eUnusedTarget
'                        myTarget2ListType = eTargetFilterType.eCaptured_Agent
'                        myTarget3ListType = eTargetFilterType.eUnusedTarget
'                    Case eMissionResult.eShowQueueAndContents, eMissionResult.eSlowProduction, eMissionResult.eSabotageProduction, _
'                      eMissionResult.eCorruptProduction, eMissionResult.eDestroyFacility
'                        myTarget2ListType = eTargetFilterType.eTarget_Facility Or eTargetFilterType.eAny
'                    Case eMissionResult.eStealCargo
'                        myTarget2ListType = eTargetFilterType.eTarget_UnitOrFacility Or eTargetFilterType.eAny
'                        myTarget3ListType = eTargetFilterType.eMineralOrTargetComponent
'                        'TODO: Shouldn't this be any Foreign component intel or Minerals I know about?
'                        ' Tested, server likes this.. so it can change easily.
'                        'myTarget3ListType = eTargetFilterType.eMineral

'                    Case eMissionResult.eAssassinateGovernor, eMissionResult.eDecreaseHousingMorale, eMissionResult.eDecreaseTaxMorale, eMissionResult.eDecreaseUnemploymentMorale
'                        myTarget2ListType = eTargetFilterType.eTarget_Colony Or eTargetFilterType.eAny
'                    Case eMissionResult.eGetColonyBudget
'                        myTarget2ListType = eTargetFilterType.ePlanet

'                        'These do not use Target 2 or 3 combo boxes
'                        'case eMissionResult.eAgencyAnalysis
'                        'Case eMissionResult.eBadPublicity
'                        'Case eMissionResult.eComponentDesignList
'                        'Case eMissionResult.eComponentResearchList
'                        'Case eMissionResult.eDegradePay
'                        'Case eMissionResult.eDiplomaticStrength
'                        'Case eMissionResult.eDropInvulField
'                        'Case eMissionResult.eFindCommandCenters
'                        'Case eMissionResult.eFindProductionFac
'                        'Case eMissionResult.eFindResearchFac
'                        'Case eMissionResult.eFindSpaceStations
'                        'case eMissionResult.eGetAgentList
'                        'case eMissionResult.eGetAliasList
'                        'Case eMissionResult.eGetAlliesList
'                        'Case eMissionResult.eGetEnemyList
'                        'Case eMissionResult.eGetMineralList
'                        'Case eMissionResult.eGetTradeTreaties
'                        'Case eMissionResult.eIFFSabotage
'                        'case eMissionResult.eLocateWormhole
'                        'Case eMissionResult.eMilitaryCoup
'                        'Case eMissionResult.eMilitaryStrength
'                        'Case eMissionResult.ePopulationScore
'                        'Case eMissionResult.eProductionScore
'                        'Case eMissionResult.eRelayBattleReports
'                        'Case eMissionResult.eSiphonItem
'                        'Case eMissionResult.eSpecialProjectList
'                        'Case eMissionResult.eTechnologyStrength
'                End Select

'                'Now, based on our filters, fill our combo boxes
'                If myTarget1ListType <> eTargetFilterType.eUnusedTarget Then
'                    FillTargetCBOFromFilter(myTarget1ListType, cboTarget1, -1, -1)
'                ElseIf myTarget2ListType <> eTargetFilterType.eUnusedTarget Then
'                    FillTargetCBOFromFilter(myTarget2ListType, cboTarget2, -1, -1)
'                End If

'                'Now, fill our methods
'                For X As Int32 = 0 To oMission.GoalUB
'                    Dim bFound As Boolean = False
'                    For Y As Int32 = 0 To cboMethod.ListCount - 1
'                        If cboMethod.ItemData(Y) = oMission.MethodIDs(X) Then
'                            bFound = True
'                            Exit For
'                        End If
'                    Next Y
'                    If bFound = False Then
'                        Dim sName As String = ""
'                        For Y As Int32 = 0 To glMissionMethodUB
'                            If glMissionMethodIdx(Y) = oMission.MethodIDs(X) Then
'                                sName = gsMissionMethods(Y)
'                                Exit For
'                            End If
'                        Next Y
'                        If sName <> "" Then
'                            cboMethod.AddItem(sName)
'                            cboMethod.ItemData(cboMethod.NewIndex) = oMission.MethodIDs(X)
'                        End If
'                    End If
'                Next X
'                If cboMethod.ListCount > 1 Then
'                    cboMethod.Enabled = True
'                    For X As Int32 = 0 To cboMethod.ListCount - 1
'                        If cboMethod.ItemData(X) = lMethodID Then
'                            cboMethod.ListIndex = X
'                            Exit For
'                        End If
'                    Next
'                Else
'                    cboMethod.Enabled = False
'                    cboMethod.ListIndex = 0
'                End If
'            End If
'        End If
'    End Sub

'    Private Sub FillTargetCBOFromFilter(ByVal yFilter As eTargetFilterType, ByRef cboData As UIComboBox, ByVal lParentID As Int32, ByVal iParentTypeID As Int16)
'        cboData.Clear()
'        cboData.Visible = True
'        cboData.Enabled = True

'        Dim bIncludeAny As Boolean = (yFilter And eTargetFilterType.eAny) <> 0

'        Dim lAnyIndex As Int32 = -1
'        If bIncludeAny = True Then
'            yFilter = yFilter Xor eTargetFilterType.eAny
'            cboData.AddItem("ANY")
'            lAnyIndex = cboData.NewIndex
'            cboData.ItemData(lAnyIndex) = Int32.MinValue
'        End If

'        'Now, determine what to fill
'        Select Case yFilter
'            Case eTargetFilterType.eCaptured_Agent
'                If lAnyIndex <> -1 Then cboData.ItemData2(lAnyIndex) = ObjectType.eAgent
'                For X As Int32 = 0 To goCurrentPlayer.AgentUB
'                    If goCurrentPlayer.AgentIdx(X) <> -1 AndAlso (goCurrentPlayer.Agents(X).lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 AndAlso (lParentID = -1 OrElse iParentTypeID <> ObjectType.ePlayer OrElse goCurrentPlayer.Agents(X).lTargetID = lParentID) Then
'                        cboData.AddItem(goCurrentPlayer.Agents(X).sAgentName)
'                        cboData.ItemData(cboData.NewIndex) = goCurrentPlayer.Agents(X).ObjectID
'                        cboData.ItemData2(cboData.NewIndex) = ObjectType.eAgent
'                    End If
'                Next X
'            Case eTargetFilterType.eMineral
'                If lAnyIndex <> -1 Then cboData.ItemData2(lAnyIndex) = ObjectType.eMineral
'                Dim lSorted() As Int32 = GetSortedMineralIdxArray(False, False)
'                If lSorted Is Nothing = False Then
'                    For X As Int32 = 0 To lSorted.GetUpperBound(0)
'                        If glMineralIdx(lSorted(X)) <> -1 Then
'                            cboData.AddItem(goMinerals(lSorted(X)).MineralName)
'                            cboData.ItemData(cboData.NewIndex) = goMinerals(lSorted(X)).ObjectID
'                            cboData.ItemData2(cboData.NewIndex) = ObjectType.eMineral
'                        End If
'                    Next X
'                End If
'            Case eTargetFilterType.ePICK_A_LOCATION
'                If lAnyIndex <> -1 Then cboData.ItemData2(lAnyIndex) = Int32.MinValue
'                cboData.AddItem("Pick A Location...")
'            Case eTargetFilterType.ePlayer
'                If lAnyIndex <> -1 Then cboData.ItemData2(lAnyIndex) = ObjectType.ePlayer

'                'Ok, sort our rels
'                Dim lIdx() As Int32 = GetSortedPlayerRelIdxArray()
'                If lIdx Is Nothing = False Then
'                    For X As Int32 = 0 To lIdx.GetUpperBound(0)
'                        Dim oTmpRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(lIdx(X))
'                        If oTmpRel Is Nothing = False Then
'                            cboData.AddItem(GetCacheObjectValue(oTmpRel.lThisPlayer, ObjectType.ePlayer))
'                            cboData.ItemData(cboData.NewIndex) = oTmpRel.lThisPlayer
'                            cboData.ItemData2(cboData.NewIndex) = ObjectType.ePlayer
'                        End If
'                    Next X
'                End If
'                'For X As Int32 = 0 To goCurrentPlayer.PlayerRelUB
'                '    Dim oTmpRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(X)
'                '    If oTmpRel Is Nothing = False Then
'                '        cboData.AddItem(GetCacheObjectValue(oTmpRel.lThisPlayer, ObjectType.ePlayer))
'                '        cboData.ItemData(cboData.NewIndex) = oTmpRel.lThisPlayer
'                '        cboData.ItemData2(cboData.NewIndex) = ObjectType.ePlayer
'                '    End If
'                'Next X
'            Case eTargetFilterType.ePlayerMissions
'                If lAnyIndex <> -1 Then cboData.ItemData2(lAnyIndex) = -1
'                For X As Int32 = 0 To goCurrentPlayer.PlayerMissionUB
'                    If goCurrentPlayer.PlayerMissionIdx(X) <> -1 AndAlso (lParentID = -1 OrElse iParentTypeID <> ObjectType.ePlayer OrElse (goCurrentPlayer.PlayerMissions(X).lTargetPlayerID = lParentID)) Then
'                        'TODO: ???
'                    End If
'                Next X
'            Case eTargetFilterType.eProductionType
'                If lAnyIndex <> -1 Then cboData.ItemData2(lAnyIndex) = -1
'                cboData.AddItem("Barracks") : cboData.ItemData(cboData.NewIndex) = ProductionType.eEnlisted
'                cboData.AddItem("Command Center") : cboData.ItemData(cboData.NewIndex) = ProductionType.eCommandCenterSpecial
'                cboData.AddItem("Factory") : cboData.ItemData(cboData.NewIndex) = ProductionType.eProduction
'                cboData.AddItem("Land Production") : cboData.ItemData(cboData.NewIndex) = ProductionType.eProduction
'                cboData.AddItem("Mining") : cboData.ItemData(cboData.NewIndex) = ProductionType.eMining
'                cboData.AddItem("Naval Production") : cboData.ItemData(cboData.NewIndex) = ProductionType.eNavalProduction
'                cboData.AddItem("Officers") : cboData.ItemData(cboData.NewIndex) = ProductionType.eOfficers
'                cboData.AddItem("Power Center") : cboData.ItemData(cboData.NewIndex) = ProductionType.ePowerCenter
'                cboData.AddItem("Refining") : cboData.ItemData(cboData.NewIndex) = ProductionType.eRefining
'                cboData.AddItem("Research") : cboData.ItemData(cboData.NewIndex) = ProductionType.eResearch
'                cboData.AddItem("Residence") : cboData.ItemData(cboData.NewIndex) = ProductionType.eColonists
'                cboData.AddItem("Spaceport") : cboData.ItemData(cboData.NewIndex) = ProductionType.eAerialProduction
'                cboData.AddItem("Space Station") : cboData.ItemData(cboData.NewIndex) = ProductionType.eSpaceStationSpecial
'                cboData.AddItem("Tradepost") : cboData.ItemData(cboData.NewIndex) = ProductionType.eTradePost
'                cboData.AddItem("Warehouse") : cboData.ItemData(cboData.NewIndex) = ProductionType.eWareHouse
'            Case eTargetFilterType.eTarget_Agent
'                If lAnyIndex <> -1 Then cboData.ItemData2(lAnyIndex) = ObjectType.eAgent
'                For X As Int32 = 0 To glItemIntelUB
'                    If glItemIntelIdx(X) <> -1 AndAlso goItemIntel(X).iItemTypeID = ObjectType.eAgent AndAlso (lParentID = -1 OrElse iParentTypeID <> ObjectType.ePlayer OrElse goItemIntel(X).lOtherPlayerID = lParentID) Then
'                        cboData.AddItem(GetCacheObjectValue(goItemIntel(X).lItemID, goItemIntel(X).iItemTypeID))
'                        cboData.ItemData(cboData.NewIndex) = goItemIntel(X).lItemID
'                        cboData.ItemData2(cboData.NewIndex) = ObjectType.eAgent
'                    End If
'                Next X
'            Case eTargetFilterType.eSolarSystem
'                If goGalaxy Is Nothing = False Then
'                    If goGalaxy.CurrentSystemIdx > -1 Then
'                        Dim oSys As SolarSystem = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
'                        If oSys Is Nothing = False Then
'                            cboData.AddItem(oSys.SystemName)
'                            cboData.ItemData(cboData.NewIndex) = oSys.ObjectID
'                            cboData.ItemData2(cboData.NewIndex) = oSys.ObjTypeID
'                        End If
'                    End If
'                End If
'            Case eTargetFilterType.ePlanet
'                If NewTutorialManager.TutorialOn = True AndAlso goCurrentPlayer.yPlayerPhase = 0 Then
'                    Dim oPlanet As Planet = Planet.GetTutorialPlanet()
'                    If oPlanet Is Nothing = False Then
'                        cboData.AddItem(oPlanet.PlanetName)
'                        cboData.ItemData(cboData.NewIndex) = oPlanet.ObjectID
'                        cboData.ItemData2(cboData.NewIndex) = ObjectType.ePlanet
'                    End If
'                ElseIf goGalaxy Is Nothing = False Then
'                    If goGalaxy.CurrentSystemIdx > -1 Then
'                        Dim oSys As SolarSystem = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
'                        If oSys Is Nothing = False Then
'                            For X As Int32 = 0 To oSys.PlanetUB
'                                If oSys.moPlanets(X) Is Nothing = False Then
'                                    cboData.AddItem(oSys.moPlanets(X).PlanetName)
'                                    cboData.ItemData(cboData.NewIndex) = oSys.moPlanets(X).ObjectID
'                                    cboData.ItemData2(cboData.NewIndex) = oSys.moPlanets(X).ObjTypeID
'                                End If
'                            Next X
'                        End If
'                    End If
'                End If
'            Case eTargetFilterType.eSystemAndPlanets
'                If goGalaxy Is Nothing = False Then
'                    If goGalaxy.CurrentSystemIdx > -1 Then
'                        Dim oSys As SolarSystem = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
'                        If oSys Is Nothing = False Then
'                            cboData.AddItem(oSys.SystemName & " (System)")
'                            cboData.ItemData(cboData.NewIndex) = oSys.ObjectID
'                            cboData.ItemData2(cboData.NewIndex) = oSys.ObjTypeID

'                            For X As Int32 = 0 To oSys.PlanetUB
'                                If oSys.moPlanets(X) Is Nothing = False Then
'                                    cboData.AddItem(oSys.moPlanets(X).PlanetName)
'                                    cboData.ItemData(cboData.NewIndex) = oSys.moPlanets(X).ObjectID
'                                    cboData.ItemData2(cboData.NewIndex) = oSys.moPlanets(X).ObjTypeID
'                                End If
'                            Next X
'                        End If
'                    End If
'                End If

'            Case eTargetFilterType.eTarget_Component
'                'PARENT SHOULD BE OUR PLAYER
'                If iParentTypeID = ObjectType.ePlayer Then
'                    FillComboFromItemIntelComponents(lParentID, cboData)
'                End If
'            Case eTargetFilterType.eMineralOrTargetComponent
'                If iParentTypeID = ObjectType.ePlayer Then 'RTP: This will never happen
'                    FillComboFromItemIntelComponentsAndMinerals(lParentID, cboData)
'                    'TODO: Test the following addition more
'                    'Else
'                    '    For X As Int32 = 0 To glItemIntelUB
'                    '        If glItemIntelIdx(X) <> -1 AndAlso goItemIntel(X).iItemTypeID = iParentTypeID AndAlso goItemIntel(X).lItemID = lParentID Then
'                    '            FillComboFromItemIntelComponentsAndMinerals(goItemIntel(X).lOtherPlayerID, cboData)
'                    '            Exit For
'                    '        End If
'                    '    Next X
'                End If
'            Case eTargetFilterType.eTargetSpecialTechUnresearched
'                If iParentTypeID <> ObjectType.ePlayer Then Return

'                For X As Int32 = 0 To glPlayerTechKnowledgeUB
'                    If glPlayerTechKnowledgeIdx(X) > -1 Then
'                        Dim oPTK As PlayerTechKnowledge = goPlayerTechKnowledge(X)
'                        If oPTK Is Nothing = False AndAlso oPTK.oTech.OwnerID = lParentID Then
'                            If oPTK.oTech.ObjTypeID = ObjectType.eSpecialTech AndAlso oPTK.oTech.ComponentDevelopmentPhase <> Base_Tech.eComponentDevelopmentPhase.eResearched Then
'                                cboData.AddItem(CType(oPTK.oTech, SpecialTech).sName)
'                                cboData.ItemData(cboData.NewIndex) = oPTK.oTech.ObjectID
'                                cboData.ItemData2(cboData.NewIndex) = ObjectType.eSpecialTech
'                            End If
'                        End If
'                    End If
'                Next
'            Case eTargetFilterType.eTarget_UnitOrFacility
'                'PARENT SHOULD BE OUR PLAYER
'                If iParentTypeID = ObjectType.ePlayer Then
'                    FillComboFromItemIntel(lParentID, ObjectType.eUnit, cboData)
'                    FillComboFromItemIntel(lParentID, ObjectType.eFacility, cboData)
'                End If
'            Case eTargetFilterType.eTarget_Unit
'                'PARENT SHOULD BE OUR PLAYER
'                If iParentTypeID = ObjectType.ePlayer Then
'                    FillComboFromItemIntel(lParentID, ObjectType.eUnit, cboData)
'                End If
'            Case eTargetFilterType.eTarget_Facility
'                'PARENT SHOULD BE OUR PLAYER
'                If iParentTypeID = ObjectType.ePlayer Then
'                    FillComboFromItemIntel(lParentID, ObjectType.eFacility, cboData)
'                End If
'            Case eTargetFilterType.eTarget_Colony
'                'PARENT SHOULD BE OUR PLAYER
'                If iParentTypeID = ObjectType.ePlayer Then
'                    FillComboFromItemIntel(lParentID, ObjectType.eColony, cboData)
'                End If
'        End Select

'    End Sub

'    Private Sub FillComboFromItemIntelComponents(ByVal lPlayerID As Int64, ByRef cboData As UIComboBox)
'        For X As Int32 = 0 To glItemIntelUB
'            If glItemIntelIdx(X) <> -1 Then
'                If goItemIntel(X).lOtherPlayerID = lPlayerID Then
'                    Select Case goItemIntel(X).iItemTypeID
'                        Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, _
'                         ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eSpecialTech, ObjectType.eWeaponTech
'                            cboData.AddItem(GetCacheObjectValue(goItemIntel(X).lItemID, goItemIntel(X).iItemTypeID))
'                            cboData.ItemData(cboData.NewIndex) = goItemIntel(X).lItemID
'                            cboData.ItemData2(cboData.NewIndex) = goItemIntel(X).iItemTypeID
'                    End Select
'                End If
'            End If
'        Next X
'        For X As Int32 = 0 To glPlayerTechKnowledgeUB
'            If glPlayerTechKnowledgeIdx(X) > -1 Then
'                Dim oPTK As PlayerTechKnowledge = goPlayerTechKnowledge(X)
'                If oPTK Is Nothing = False AndAlso oPTK.oTech.OwnerID = lPlayerID Then
'                    cboData.AddItem(goPlayerTechKnowledge(X).oTech.GetComponentName) 'GetCacheObjectValue(goPlayerTechKnowledge(X).oTech.ObjectID, goPlayerTechKnowledge(X).oTech.ObjTypeID))
'                    cboData.ItemData(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjectID
'                    cboData.ItemData2(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjTypeID
'                End If
'            End If
'        Next X
'    End Sub
'    Private Sub FillComboFromItemIntelComponentsAndMinerals(ByVal lPlayerID As Int64, ByRef cboData As UIComboBox)
'        For X As Int32 = 0 To glItemIntelUB
'            If glItemIntelIdx(X) <> -1 Then
'                If goItemIntel(X).lOtherPlayerID = lPlayerID Then
'                    Select Case goItemIntel(X).iItemTypeID
'                        Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.eHullTech, ObjectType.ePrototype, _
'                         ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eSpecialTech, ObjectType.eWeaponTech, ObjectType.eMineral
'                            cboData.AddItem(GetCacheObjectValue(goItemIntel(X).lItemID, goItemIntel(X).iItemTypeID))
'                            cboData.ItemData(cboData.NewIndex) = goItemIntel(X).lItemID
'                            cboData.ItemData2(cboData.NewIndex) = goItemIntel(X).iItemTypeID
'                    End Select
'                End If
'            End If
'        Next X
'        For X As Int32 = 0 To glPlayerTechKnowledgeUB
'            If glPlayerTechKnowledgeIdx(X) > -1 Then
'                Dim oPTK As PlayerTechKnowledge = goPlayerTechKnowledge(X)
'                If oPTK Is Nothing = False AndAlso oPTK.oTech.OwnerID = lPlayerID Then
'                    cboData.AddItem(GetCacheObjectValue(goPlayerTechKnowledge(X).oTech.ObjectID, goPlayerTechKnowledge(X).oTech.ObjTypeID))
'                    cboData.ItemData(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjectID
'                    cboData.ItemData2(cboData.NewIndex) = goPlayerTechKnowledge(X).oTech.ObjTypeID
'                End If
'            End If
'        Next X
'    End Sub
'    Private Sub FillComboFromItemIntel(ByVal lPlayerID As Int32, ByVal iTypeID As Int16, ByRef cboData As UIComboBox)
'        For X As Int32 = 0 To glItemIntelUB
'            If glItemIntelIdx(X) <> -1 Then
'                If goItemIntel(X).lOtherPlayerID = lPlayerID AndAlso goItemIntel(X).iItemTypeID = iTypeID Then
'                    Dim sName As String = ""
'                    sName = GetCacheObjectValue(goItemIntel(X).lItemID, goItemIntel(X).iItemTypeID)
'                    If goItemIntel(X).EnvirTypeID > 0 AndAlso goItemIntel(X).EnvirID > 0 Then
'                        sName &= " - " & GetCacheObjectValue(goItemIntel(X).EnvirID, goItemIntel(X).EnvirTypeID)
'                    End If
'                    cboData.AddItem(sName)
'                    cboData.ItemData(cboData.NewIndex) = goItemIntel(X).lItemID
'                    cboData.ItemData2(cboData.NewIndex) = goItemIntel(X).iItemTypeID
'                End If
'            End If
'        Next X
'    End Sub

'    Private Sub cboTarget1_ItemSelected(ByVal lItemIndex As Integer) Handles cboTarget1.ItemSelected
'        If cboTarget1.ListIndex <> -1 Then
'            'ok, get our player
'            If myTarget2ListType <> eTargetFilterType.eUnusedTarget Then
'                'now, fill our target 2 
'                Dim lPlayerID As Int32 = cboTarget1.ItemData(cboTarget1.ListIndex)
'                FillTargetCBOFromFilter(myTarget2ListType, cboTarget2, lPlayerID, ObjectType.ePlayer)
'                cboTarget2.Enabled = True
'            Else : cboTarget2.Enabled = False
'            End If
'        End If
'    End Sub

'    Private Sub cboTarget2_ItemSelected(ByVal lItemIndex As Integer) Handles cboTarget2.ItemSelected
'        If cboTarget2.ListIndex <> -1 Then

'            Dim lID As Int32 = cboTarget2.ItemData(cboTarget2.ListIndex)
'            Dim iTypeID As Int16 = CShort(cboTarget2.ItemData2(cboTarget2.ListIndex))

'            If NewTutorialManager.TutorialOn = True Then
'                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionTargetSelected, lID, iTypeID, -1, "")
'            End If

'            If myTarget3ListType <> eTargetFilterType.eUnusedTarget Then
'                FillTargetCBOFromFilter(myTarget3ListType, cboTarget3, lID, iTypeID)
'                cboTarget3.Enabled = True
'            Else : cboTarget3.Enabled = False
'            End If
'        End If
'    End Sub

'    Private Sub cboMethod_ItemSelected(ByVal lItemIndex As Integer) Handles cboMethod.ItemSelected
'        'ok, clear goals
'        lstGoals.Clear()
'        tvwGoals.Clear()
'        lblGoalName.Caption = ""
'        lblGoalTime.Caption = ""
'        lblGoalRisk.Caption = ""
'        'txtGoalDesc.Caption = ""
'        lblGoalPhase.Caption = ""

'        If tvwGoals.oIconTexture Is Nothing OrElse tvwGoals.oIconTexture.Disposed = True Then
'            tvwGoals.oIconTexture = goResMgr.GetTexture("Interface.dds", GFXResourceManager.eGetTextureType.UserInterface, "Interface.pak")
'        End If

'        Dim oMission As Mission = GetSelectedMission()
'        If oMission Is Nothing Then Return

'        If cboMethod.ListIndex <> -1 Then
'            Dim lMethodID As Int32 = cboMethod.ItemData(cboMethod.ListIndex)

'            If NewTutorialManager.TutorialOn = True Then
'                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionMethodSelected, lMethodID, -1, -1, "")
'            End If

'            Dim oSafehouse As Goal = Goal.GetSafehouseGoal()
'            If oSafehouse Is Nothing = False Then
'                Dim oSafehouseNode As UITreeView.UITreeViewItem = tvwGoals.AddNode(oSafehouse.sGoalName, oSafehouse.ObjectID, 1, -1, Nothing, Nothing)
'                oSafehouseNode.bRenderIcon = True
'                oSafehouseNode.IconColor = muSettings.InterfaceBorderColor
'                oSafehouseNode.bClickIcon = True : oSafehouseNode.bRenderIcon = True
'                If mbUseSafehouse = True Then
'                    moMission.ySafeHouseSetting = 1
'                    moMission.oSafehouseMissionGoal = New PlayerMissionGoal
'                    oSafehouseNode.IconRect = grc_UI(elInterfaceRectangle.eCheck_Checked)
'                Else
'                    oSafehouseNode.IconRect = grc_UI(elInterfaceRectangle.eCheck_Unchecked)
'                    moMission.ySafeHouseSetting = 0
'                    osafehousenode.lItemData2 =0
'                End If
'            End If

'            moMission.lMethodID = lMethodID

'            For X As Int32 = 0 To glMissionMethodUB
'                If glMissionMethodIdx(X) = lMethodID Then
'                    lblMethodDesc.Caption = gsMissionDescs(X)
'                    Exit For
'                End If
'            Next X

'            Dim lSorted() As Int32 = Nothing
'            Dim lSortedUB As Int32 = -1
'            For X As Int32 = 0 To oMission.GoalUB
'                Dim lIdx As Int32 = -1
'                If oMission.MethodIDs(X) = lMethodID Then
'                    For Y As Int32 = 0 To lSortedUB
'                        If oMission.Goals(lSorted(Y)).MissionPhase > oMission.Goals(X).MissionPhase Then
'                            lIdx = Y
'                            Exit For
'                        End If
'                    Next Y
'                    lSortedUB += 1
'                    ReDim Preserve lSorted(lSortedUB)
'                    If lIdx = -1 Then
'                        lSorted(lSortedUB) = X
'                    Else
'                        For Y As Int32 = lSortedUB To lIdx + 1 Step -1
'                            lSorted(Y) = lSorted(Y - 1)
'                        Next Y
'                        lSorted(lIdx) = X
'                    End If
'                End If
'            Next X

'            For X As Int32 = 0 To lSortedUB
'                Dim lGoalID As Int32 = oMission.Goals(lSorted(X)).ObjectID
'                Dim oGoalNode As UITreeView.UITreeViewItem = tvwGoals.AddNode(oMission.Goals(lSorted(X)).sGoalName, lGoalID, 1, 0, Nothing, Nothing)
'                oGoalNode.bRenderIcon = True
'                oGoalNode.IconColor = muSettings.InterfaceTextBoxForeColor
'                oGoalNode.IconRect = grc_UI(elInterfaceRectangle.eCheck_Unchecked)
'                Dim oGoal As Goal = Nothing
'                Dim oPMG As PlayerMissionGoal = Nothing

'                For Y As Int32 = 0 To oMission.GoalUB
'                    If oMission.MethodIDs(Y) = lMethodID AndAlso oMission.Goals(Y).ObjectID = lGoalID Then

'                        If NewTutorialManager.TutorialOn = True Then
'                            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionGoalSelected, lGoalID, -1, -1, "")
'                        End If

'                        oGoal = oMission.Goals(Y)
'                        oPMG = moMission.oMissionGoals(Y)
'                        Exit For
'                    End If
'                Next Y
'                If oGoal Is Nothing = False Then
'                    For lSkillSetID As Int32 = 0 To oGoal.SkillSetUB
'                        Dim oSkillSet As SkillSet = oGoal.SkillSets(lSkillSetID)
'                        If oSkillSet Is Nothing = False Then
'                            Dim oSkillSetNode As UITreeView.UITreeViewItem = tvwGoals.AddNode(oSkillSet.sSkillSetName, oSkillSet.SkillSetID, 2, 0, oGoalNode, Nothing)

'                        End If
'                    Next lSkillSetID
'                End If

'            Next X

'        End If
'    End Sub

'    Private Function GetSelectedMission() As Mission
'        If cboMission.ListIndex <> -1 Then
'            Dim lID As Int32 = cboMission.ItemData(cboMission.ListIndex)
'            For X As Int32 = 0 To glMissionUB
'                If glMissionIdx(X) = lID Then
'                    Return goMissions(X)
'                End If
'            Next X
'        End If
'        Return Nothing
'    End Function

'    Private Sub lstGoals_ItemClick(ByVal lIndex As Integer) Handles lstGoals.ItemClick
'        lblGoalName.Caption = ""
'        lblGoalTime.Caption = ""
'        lblGoalRisk.Caption = ""
'        'txtGoalDesc.Caption = ""
'        lblGoalPhase.Caption = ""

'        Dim oMission As Mission = GetSelectedMission()
'        If oMission Is Nothing Then Return
'        Dim lMethodID As Int32 = -1
'        If cboMethod.ListIndex <> -1 Then lMethodID = cboMethod.ItemData(cboMethod.ListIndex)
'        If lMethodID = -1 Then Return

'        If lstGoals.ListIndex > -1 Then
'            Dim lGoalID As Int32 = lstGoals.ItemData(lstGoals.ListIndex)

'            Dim oGoal As Goal = Nothing
'            Dim oPMG As PlayerMissionGoal = Nothing
'            If lGoalID = Goal.ml_SAFEHOUSE_GOAL_ID Then
'                oGoal = Goal.GetSafehouseGoal()
'                If moMission.oSafehouseMissionGoal Is Nothing Then
'                    moMission.oSafehouseMissionGoal = New PlayerMissionGoal()
'                    With moMission.oSafehouseMissionGoal
'                        .oGoal = oGoal
'                        .oMission = moMission
'                    End With
'                End If
'                oPMG = moMission.oSafehouseMissionGoal
'            Else
'                For X As Int32 = 0 To oMission.GoalUB
'                    If oMission.MethodIDs(X) = lMethodID AndAlso oMission.Goals(X).ObjectID = lGoalID Then

'                        If NewTutorialManager.TutorialOn = True Then
'                            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionGoalSelected, lGoalID, -1, -1, "")
'                        End If

'                        oGoal = oMission.Goals(X)
'                        oPMG = moMission.oMissionGoals(X)
'                        Exit For
'                    End If
'                Next X
'            End If

'            If oGoal Is Nothing = False Then
'                'ok, found  it
'                With oGoal
'                    lblGoalName.Caption = "Goal: " & .sGoalName
'                    lblGoalTime.Caption = "Goal Time: about " & GetDurationFromSeconds(.BaseTime, True)
'                    lblGoalRisk.Caption = "Risk Assessment: " & Goal.GetRiskAssessment(.RiskOfDetection)
'                    lblGoalRisk.ForeColor = Goal.GetRiskCaptionColor(.RiskOfDetection)
'                    'txtGoalDesc.Caption = .sGoalDesc
'                    Select Case .MissionPhase
'                        Case eMissionPhase.eFlippingTheSwitch
'                            lblGoalPhase.Caption = "Phase: Execution"
'                        Case eMissionPhase.ePreparationTime
'                            lblGoalPhase.Caption = "Phase: Preparation"
'                        Case eMissionPhase.eReinfiltrationPhase
'                            lblGoalPhase.Caption = "Phase: Reinfiltration"
'                        Case eMissionPhase.eSettingTheStage
'                            lblGoalPhase.Caption = "Phase: Setting the Stage"
'                    End Select
'                End With
'                fraSkillSets.SetFromGoal(oGoal, moMission)

'                Dim lSkillsetID As Int32 = -1
'                If oPMG.lAssignmentUB <> -1 AndAlso oPMG.oSkillSet Is Nothing = False Then
'                    lSkillsetID = oPMG.oSkillSet.SkillSetID
'                End If
'                fraSkillSets.SetSkillsetID(lSkillsetID)

'                For Y As Int32 = 0 To oPMG.lAssignmentUB
'                    If oPMG.oAssignments(Y) Is Nothing = False Then
'                        fraSkillSets.SetAgentAssignment(oPMG.oAssignments(Y))
'                    End If
'                Next Y
'            End If

'        End If
'    End Sub

'    Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'    End Sub

'    Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
'        '
'    End Sub

'    Private Sub fraSkillSets_AgentAssignment(ByRef oSkill As SkillSet_Skill, ByRef oSender As ctlAgentAssignment) Handles fraSkillSets.AgentAssignment
'        'Ok, take the currently selected agent...

'        If mlSelectedAgentID = -1 Then
'            MyBase.moUILib.AddNotification("Select an agent to assign and then click Assign.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'        Else
'            Dim oAgent As Agent = Nothing
'            For X As Int32 = 0 To goCurrentPlayer.AgentUB
'                If goCurrentPlayer.AgentIdx(X) = mlSelectedAgentID Then
'                    oAgent = goCurrentPlayer.Agents(X)
'                    Exit For
'                End If
'            Next X

'            Dim lMethodID As Int32 = -1
'            If cboMethod.ListIndex > -1 Then lMethodID = cboMethod.ItemData(cboMethod.ListIndex)

'            If oAgent Is Nothing = False Then

'                If oAgent Is Nothing = False Then ' AndAlso oAgent.bRequestedSkillList = False Then
'                    'oAgent.bRequestedSkillList = True
'                    'Dim yOut(5) As Byte
'                    'System.BitConverter.GetBytes(GlobalMessageCode.eGetSkillList).CopyTo(yOut, 0)
'                    'System.BitConverter.GetBytes(oAgent.ObjectID).CopyTo(yOut, 2)
'                    'MyBase.moUILib.SendMsgToPrimary(yOut)
'                    'Return
'                End If

'                'check if the agent has the specified skill
'                If oSkill Is Nothing = False Then

'                    Dim yProf As Byte = oAgent.GetSkillProficiency(oSkill.oSkill.ObjectID)
'                    If yProf < oSkill.oSkill.MinVal Then
'                        yProf = oAgent.GetSkillProficiency(36)  'naturally talented
'                        If yProf < oSkill.oSkill.MinVal Then
'                            MyBase.moUILib.AddNotification("The agent selected does not have " & oSkill.oSkill.SkillName & " at a proficient level.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return
'                        Else

'                            If oSkill.oSkill.SkillType <> 0 Then
'                                MyBase.moUILib.AddNotification("Unable to use Naturally Talented for " & oSkill.oSkill.SkillName & " as it is an ability.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                                Return
'                            End If

'                            Dim lTmpGoal As Int32 = oSkill.oSkillSet.oGoal.ObjectID
'                            Dim bValid As Boolean = False
'                            Dim oTmpPMG As PlayerMissionGoal = Nothing
'                            If lTmpGoal = Goal.ml_SAFEHOUSE_GOAL_ID Then
'                                oTmpPMG = moMission.oSafehouseMissionGoal
'                            Else
'                                For X As Int32 = 0 To moMission.oMissionGoals.GetUpperBound(0)
'                                    If moMission.oMission.MethodIDs(X) = lMethodID AndAlso moMission.oMissionGoals(X).oGoal.ObjectID = lTmpGoal Then
'                                        oTmpPMG = moMission.oMissionGoals(X)
'                                        Exit For
'                                    End If
'                                Next X
'                            End If
'                            If oTmpPMG Is Nothing = False Then
'                                If oTmpPMG.AgentAssignedUsingNaturalTalents(oAgent.ObjectID) = True Then
'                                    MyBase.moUILib.AddNotification("The agent selected does not have " & oSkill.oSkill.SkillName & " at a proficient level.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                                    Return
'                                End If
'                            End If

'                            MyBase.moUILib.AddNotification("Agent added using Naturally Talented.", Color.Yellow, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                        End If
'                    End If
'                End If
'                If NewTutorialManager.TutorialOn = True Then
'                    NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionSkillsetAssign, oAgent.ObjectID, oSkill.oSkillSet.SkillSetID, oSkill.oSkill.ObjectID, "")
'                End If
'                oSender.SetAgent(oAgent)
'            End If

'            'Now, make the assignment
'            Dim lGoalID As Int32 = -1
'            Dim yPhase As Byte = 0
'            Dim oPMG As PlayerMissionGoal = Nothing
'            If oSkill Is Nothing = False Then
'                lGoalID = oSkill.oSkillSet.oGoal.ObjectID
'                For X As Int32 = 0 To moMission.oMissionGoals.GetUpperBound(0)
'                    If moMission.oMission.MethodIDs(X) = lMethodID AndAlso moMission.oMissionGoals(X).oGoal.ObjectID = lGoalID Then
'                        moMission.oMissionGoals(X).AddAgentAssignment(oAgent, oSkill.oSkill)
'                        oPMG = moMission.oMissionGoals(X)
'                        yPhase = moMission.oMissionGoals(X).oGoal.MissionPhase
'                        Exit For
'                    End If
'                Next X
'            End If

'            'Now, ensure that the agent is not assigned cover
'            For X As Int32 = 0 To moMission.lPhaseUB
'                If moMission.oPhases(X).yPhase = yPhase Then
'                    moMission.oPhases(X).RemoveCoverAgent(oAgent.ObjectID)
'                    RefreshCoverAgentList()
'                    Exit For
'                End If
'            Next X
'        End If
'    End Sub

'    Private Sub fraSkillSets_SkillSetSelected(ByRef oGoal As Goal, ByRef oSkillset As SkillSet) Handles fraSkillSets.SkillSetSelected
'        If moMission Is Nothing = False AndAlso moMission.oMission Is Nothing = False AndAlso oGoal Is Nothing = False Then
'            Dim lMethodID As Int32 = -1
'            If cboMethod Is Nothing = False AndAlso cboMethod.ListIndex > -1 Then
'                lMethodID = cboMethod.ItemData(cboMethod.ListIndex)
'            End If
'            For X As Int32 = 0 To moMission.oMission.GoalUB
'                If moMission.oMission.Goals(X).ObjectID = oGoal.ObjectID AndAlso moMission.oMission.MethodIDs(X) = lMethodID Then

'                    If NewTutorialManager.TutorialOn = True AndAlso oSkillset Is Nothing = False Then
'                        NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionSkillsetSelected, oSkillset.SkillSetID, -1, -1, "")
'                    End If

'                    moMission.oMissionGoals(X).oSkillSet = oSkillset
'                    moMission.oMissionGoals(X).lAssignmentUB = -1
'                    moFilterSkillSet = oSkillset
'                    Exit For
'                End If
'            Next X
'        End If
'    End Sub

'    Private Sub fraSkillSets_SetSkillsetFilter(ByRef oSkillset As SkillSet) Handles fraSkillSets.SetSkillsetFilter
'        moFilterSkillSet = oSkillset
'        Me.IsDirty = True
'    End Sub

'    Private Sub fraSkillSets_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles fraSkillSets.OnMouseMove
'        frmMission_OnMouseMove(lMouseX, lMouseY, lButton)
'    End Sub

'    Private Sub btnAgentLeft_Click(ByVal sName As String) Handles btnAgentLeft.Click
'        mlScrollValue -= 1
'        If mlScrollValue < 0 Then mlScrollValue = 0
'        Me.IsDirty = True
'    End Sub

'    Private Sub btnAgentRight_Click(ByVal sName As String) Handles btnAgentRight.Click
'        mlScrollValue += 1
'        If mlScrollValue > mlScrollMaxValue Then mlScrollValue = mlScrollMaxValue
'        Me.IsDirty = True
'    End Sub

'    Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
'        'Cancels the mission
'        If moMission Is Nothing = False Then
'            If moMission.PM_ID > -1 Then
'                'check if the mission can be cancelled
'                If moMission.yCurrentPhase = eMissionPhase.eInPlanning OrElse moMission.yCurrentPhase = eMissionPhase.ePreparationTime Then
'                    'send cancel msg to the primary
'                    Dim yMsg() As Byte = GenerateMissionMsg(2) '2 = cancel
'                    If yMsg Is Nothing = False Then MyBase.moUILib.SendMsgToPrimary(yMsg)
'                    For X As Int32 = 0 To goCurrentPlayer.PlayerMissionUB
'                        If goCurrentPlayer.PlayerMissionIdx(X) = moMission.PM_ID Then
'                            goCurrentPlayer.PlayerMissions(X).yCurrentPhase = eMissionPhase.eCancelled
'                        End If
'                    Next X
'                Else
'                    MyBase.moUILib.AddNotification("This mission cannot be cancelled at this time.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                    Return
'                End If
'            End If
'        End If
'        'remove the window
'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'    End Sub

'    Private Sub btnExecuteNow_Click(ByVal sName As String) Handles btnExecuteNow.Click
'        '1) Send the primary a message to execute now... do a validation
'        If ValidateMission(True) = False Then Return

'        Dim yMsg() As Byte = GenerateMissionMsg(1)      '1 for execute now
'        If yMsg Is Nothing = False Then
'            MyBase.moUILib.SendMsgToPrimary(yMsg)
'            If NewTutorialManager.TutorialOn = True Then
'                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionExecuteNow, -1, -1, -1, "")
'            End If
'        End If
'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'    End Sub

'    Private Sub btnSaveForLater_Click(ByVal sName As String) Handles btnSaveForLater.Click
'        '1) Send the primary a message to save for later
'        Try
'            If ValidateMission(False) = False Then Return
'            Dim yMsg() As Byte = GenerateMissionMsg(0)              '0 = save for later
'            If yMsg Is Nothing = False Then
'                MyBase.moUILib.SendMsgToPrimary(yMsg)
'            End If
'        Catch ex As Exception
'            MyBase.moUILib.AddNotification("Unable to save mission for later.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'        End Try
'        MyBase.moUILib.RemoveWindow(Me.ControlName)
'    End Sub

'    Private Function ValidateMission(ByVal bExact As Boolean) As Boolean
'        Try
'            'is omission set?
'            If moMission Is Nothing Then Return False

'            If moMission.oMission Is Nothing Then
'                MyBase.moUILib.AddNotification("You must select a mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                Return False
'            End If

'            If moMission.oMission.ProgramControlID = eMissionResult.eGetFacilityList Then
'                'ok, check our target 2...
'                Dim lT2Val As Int32 = -1
'                If cboTarget2.ListIndex > -1 Then lT2Val = cboTarget2.ItemData(cboTarget2.ListIndex)
'                Dim bSys As Boolean = False
'                If lT2Val = ProductionType.eSpaceStationSpecial Then
'                    bSys = True
'                End If
'                Dim lT3Val As Int32 = -1
'                If cboTarget3.ListIndex > -1 Then lT3Val = cboTarget3.ItemData2(cboTarget3.ListIndex)
'                If lT3Val = ObjectType.ePlanet Then
'                    If bSys = True Then
'                        MyBase.moUILib.AddNotification("You must select the system to find space stations.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                        Return False
'                    End If
'                ElseIf lT3Val = ObjectType.eSolarSystem Then
'                    If bSys = False Then
'                        MyBase.moUILib.AddNotification("You must select a planet when searching for anything other than space stations.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                        Return False
'                    End If
'                End If

'            End If
'            'myTarget2ListType = eTargetFilterType.eProductionType
'            'myTarget3ListType = eTargetFilterType.eSystemAndPlanets

'            If bExact = True Then
'                If cboMethod.ListIndex < 0 Then
'                    MyBase.moUILib.AddNotification("You must select a method for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                    Return False
'                End If
'                Dim lMethodID As Int32 = cboMethod.ItemData(cboMethod.ListIndex)
'                Dim bGood As Boolean = False
'                For X As Int32 = 0 To moMission.oMission.GoalUB
'                    If lMethodID = moMission.oMission.MethodIDs(X) Then
'                        bGood = True
'                        Exit For
'                    End If
'                Next X
'                If bGood = False Then
'                    MyBase.moUILib.AddNotification("You must select a method for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                    Return False
'                End If

'                bGood = False
'                'check our target and target types
'                Select Case CType(moMission.oMission.ProgramControlID, eMissionResult)
'                    Case eMissionResult.ePlantEvidence
'                        'player, otherplayerid, pm_id
'                        If cboTarget1.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select a target player for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If
'                        If cboTarget2.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select a target player for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If
'                        If cboTarget3.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select a mission you have ran to implicate.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If
'                    Case eMissionResult.eImpedeCurrentDevelopment, eMissionResult.eDestroyCurrentSpecialProject
'                        If cboTarget1.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select a target player for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If
'                        If cboTarget2.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select a target special tech for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If
'                    Case eMissionResult.eFindMineral
'                        'No player, mineral
'                        If cboTarget2.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select a mineral to search for on this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If
'                    Case eMissionResult.eSearchAndRescueAgent
'                        'no player, agent
'                        If cboTarget2.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select an agent to search for on this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If
'                    Case eMissionResult.eReconPlanetMap
'                        'no player, locx, locz
'                    Case eMissionResult.eGeologicalSurvey
'                        'no player, planet
'                        If cboTarget2.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select a planet to survey for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If
'                    Case eMissionResult.eGetColonyBudget
'                        'player, planet
'                        If cboTarget1.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select a target player for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If
'                        If cboTarget2.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select a planet to survey for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If
'                    Case eMissionResult.eDecreaseHousingMorale, eMissionResult.eDecreaseTaxMorale, eMissionResult.eDecreaseUnemploymentMorale
'                        If cboTarget1.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select a target player for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If
'                        If cboTarget2.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select a planet to survey for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If
'                    Case eMissionResult.eGetFacilityList
'                        If cboTarget1.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select a target player for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If
'                        If cboTarget2.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select a production type to find.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If
'                        If cboTarget3.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select a planet/system to search.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If
'                    Case eMissionResult.eTutorialFindFactory, eMissionResult.eAudit
'                        'do nothing,
'                    Case Else
'                        If cboTarget1.ListIndex < 0 Then
'                            MyBase.moUILib.AddNotification("You must select a target player for this mission.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If

'                End Select


'                'all assignments must exist...
'                If mbUseSafehouse = True Then
'                    If moMission.oSafehouseMissionGoal Is Nothing Then
'                        MyBase.moUILib.AddNotification("Specify more detail for the safehouse goal.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                        Return False
'                    End If
'                    If moMission.oSafehouseMissionGoal.oSkillSet Is Nothing Then
'                        MyBase.moUILib.AddNotification("Select a skillset for the safehouse goal.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                        Return False
'                    End If
'                    If VerifyMissionGoal(moMission.oSafehouseMissionGoal) = False Then Return False
'                End If
'                For X As Int32 = 0 To moMission.oMission.GoalUB
'                    If moMission.oMission.MethodIDs(X) = lMethodID Then
'                        If moMission.oMissionGoals(X).oSkillSet Is Nothing Then
'                            MyBase.moUILib.AddNotification("All mission goals must have skillsets selected.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                            Return False
'                        End If

'                        If VerifyMissionGoal(moMission.oMissionGoals(X)) = False Then Return False
'                        'Dim bSkillGood(moMission.oMissionGoals(X).oSkillSet.SkillUB) As Boolean
'                        'For Y As Int32 = 0 To moMission.oMissionGoals(X).oSkillSet.SkillUB
'                        '	bSkillGood(Y) = False
'                        'Next Y

'                        'Dim bAgentDetained As Boolean = False
'                        'For Y As Int32 = 0 To moMission.oMissionGoals(X).lAssignmentUB
'                        '	Dim lSkillID As Int32 = moMission.oMissionGoals(X).oAssignments(Y).oSkill.ObjectID
'                        '	If moMission.oMissionGoals(X).oAssignments(Y).oAgent Is Nothing Then
'                        '		MyBase.moUILib.AddNotification("Not all goals have fully assigned skillsets.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                        '		Return False
'                        '	End If
'                        '	If (moMission.oMissionGoals(X).oAssignments(Y).oAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 OrElse (moMission.oMissionGoals(X).oAssignments(Y).oAgent.lAgentStatus And AgentStatus.IsDead) <> 0 Then
'                        '		bAgentDetained = True
'                        '	Else
'                        '		For Z As Int32 = 0 To moMission.oMissionGoals(X).oSkillSet.SkillUB
'                        '			If moMission.oMissionGoals(X).oSkillSet.Skills(Z).oSkill.ObjectID = lSkillID Then
'                        '				bSkillGood(Z) = True
'                        '				Exit For
'                        '			End If
'                        '		Next Z
'                        '	End If
'                        'Next Y

'                        ''Now, verify our goods
'                        'For Y As Int32 = 0 To moMission.oMissionGoals(X).oSkillSet.SkillUB
'                        '	If bSkillGood(Y) = False Then
'                        '		MyBase.moUILib.AddNotification("This mission is not a valid mission. ", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                        '		Return False
'                        '	End If
'                        'Next Y
'                    End If
'                Next X
'            End If
'        Catch ex As Exception
'            MyBase.moUILib.AddNotification("This mission is not a valid mission. " & ex.Message, Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return False
'        End Try
'        Return True
'    End Function

'    Private Function VerifyMissionGoal(ByVal oPMG As PlayerMissionGoal) As Boolean
'        Dim bSkillGood(oPMG.oSkillSet.SkillUB) As Boolean
'        For Y As Int32 = 0 To oPMG.oSkillSet.SkillUB
'            bSkillGood(Y) = False
'        Next Y

'        Dim bAgentDetained As Boolean = False
'        For Y As Int32 = 0 To oPMG.lAssignmentUB
'            Dim lSkillID As Int32 = oPMG.oAssignments(Y).oSkill.ObjectID
'            If oPMG.oAssignments(Y).oAgent Is Nothing Then
'                MyBase.moUILib.AddNotification("Not all goals have fully assigned skillsets.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                Return False
'            End If
'            If (oPMG.oAssignments(Y).oAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 OrElse (oPMG.oAssignments(Y).oAgent.lAgentStatus And AgentStatus.IsDead) <> 0 Then
'                bAgentDetained = True
'            Else
'                For Z As Int32 = 0 To oPMG.oSkillSet.SkillUB
'                    If oPMG.oSkillSet.Skills(Z).oSkill.ObjectID = lSkillID Then
'                        bSkillGood(Z) = True
'                        Exit For
'                    End If
'                Next Z
'            End If
'        Next Y

'        'Now, verify our goods
'        For Y As Int32 = 0 To oPMG.oSkillSet.SkillUB
'            If bSkillGood(Y) = False Then
'                MyBase.moUILib.AddNotification("This mission is not a valid mission. ", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                Return False
'            End If
'        Next Y
'        Return True
'    End Function

'    Private Function GenerateMissionMsg(ByVal yActionID As Byte) As Byte()
'        'Ok, let's create our mission and submit it to the server
'        Dim yMsg() As Byte
'        Dim lPos As Int32 = 0
'        Dim yPhaseCnt As Byte = 0
'        Dim yGoalCnt As Byte = 0
'        Dim lPhaseCvrCnt() As Int32 = Nothing
'        Dim lAssignCnt() As Int32 = Nothing

'        If moMission Is Nothing Then Return Nothing

'        With moMission
'            Dim lMethodID As Int32 = -1
'            If cboMethod.ListIndex > -1 Then lMethodID = cboMethod.ItemData(cboMethod.ListIndex)
'            'determine our counts now
'            ReDim lPhaseCvrCnt(.lPhaseUB)
'            For X As Int32 = 0 To .lPhaseUB
'                lPhaseCvrCnt(X) = 0
'                If .oPhases(X) Is Nothing = False Then
'                    yPhaseCnt += CByte(1)
'                    For Y As Int32 = 0 To .oPhases(X).lCoverAgentUB
'                        If .oPhases(X).lCoverAgentIdx(Y) <> -1 Then lPhaseCvrCnt(X) += 1
'                    Next Y
'                End If
'            Next X
'            If .oMission Is Nothing = False Then
'                ReDim lAssignCnt(.oMission.GoalUB)
'                For X As Int32 = 0 To .oMission.GoalUB
'                    lAssignCnt(X) = 0
'                    If .oMissionGoals(X) Is Nothing = False AndAlso .oMission.MethodIDs(X) = lMethodID Then
'                        yGoalCnt += CByte(1)
'                        For Y As Int32 = 0 To .oMissionGoals(X).lAssignmentUB
'                            If .oMissionGoals(X).oAssignments(Y) Is Nothing = False Then lAssignCnt(X) += 1
'                        Next Y
'                    End If
'                Next X
'            Else
'                ReDim lAssignCnt(-1)
'            End If

'            'Ok, we have everything, let's determine total size
'            '29 + (yPhaseCnt * 2) + (lSumOfPhaseCover * 4) + (yGoalCnt * 8) + (lSumAgentAssignment * 8)
'            Dim lSumOfPhaseCover As Int32 = 0
'            Dim lSumAgentAssignment As Int32 = 0
'            For X As Int32 = 0 To .lPhaseUB
'                lSumOfPhaseCover += lPhaseCvrCnt(X)
'            Next X
'            For X As Int32 = 0 To lAssignCnt.GetUpperBound(0)
'                lSumAgentAssignment += lAssignCnt(X)
'            Next X

'            Dim lSafehouseSize As Int32 = 0
'            If mbUseSafehouse = True Then lSafehouseSize = 20

'            ReDim yMsg(33 + (CInt(yPhaseCnt) * 2) + (lSumOfPhaseCover * 4) + (CInt(yGoalCnt) * 9) + (lSumAgentAssignment * 8))  '28 not 29 for ub

'            System.BitConverter.GetBytes(GlobalMessageCode.eSubmitMission).CopyTo(yMsg, lPos) : lPos += 2
'            yMsg(lPos) = yActionID : lPos += 1
'            System.BitConverter.GetBytes(.PM_ID).CopyTo(yMsg, lPos) : lPos += 4
'            If .oMission Is Nothing = False Then
'                System.BitConverter.GetBytes(.oMission.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
'            Else
'                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'            End If
'            If cboTarget1.Visible = False OrElse cboTarget1.ListIndex < 0 Then
'                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'            Else : System.BitConverter.GetBytes(cboTarget1.ItemData(cboTarget1.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
'            End If
'            If cboTarget2.Visible = True AndAlso cboTarget2.Enabled = True AndAlso cboTarget2.ListIndex > -1 Then
'                System.BitConverter.GetBytes(cboTarget2.ItemData(cboTarget2.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
'                System.BitConverter.GetBytes(CShort(cboTarget2.ItemData2(cboTarget2.ListIndex))).CopyTo(yMsg, lPos) : lPos += 2
'            Else
'                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'                System.BitConverter.GetBytes(-1S).CopyTo(yMsg, lPos) : lPos += 2
'            End If
'            If cboTarget3.Visible = True AndAlso cboTarget3.Enabled = True AndAlso cboTarget3.ListIndex > -1 Then
'                System.BitConverter.GetBytes(cboTarget3.ItemData(cboTarget3.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
'                System.BitConverter.GetBytes(CShort(cboTarget3.ItemData2(cboTarget3.ListIndex))).CopyTo(yMsg, lPos) : lPos += 2
'            Else
'                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'                System.BitConverter.GetBytes(-1S).CopyTo(yMsg, lPos) : lPos += 2
'            End If
'            yMsg(lPos) = .ySafeHouseSetting : lPos += 1
'            If cboMethod.ListIndex > -1 Then
'                System.BitConverter.GetBytes(cboMethod.ItemData(cboMethod.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
'            Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'            End If

'            If mbUseSafehouse = True Then
'                If moMission.oSafehouseMissionGoal Is Nothing = False Then
'                    With moMission.oSafehouseMissionGoal
'                        'skillsetid
'                        If .oSkillSet Is Nothing = False Then
'                            System.BitConverter.GetBytes(.oSkillSet.SkillSetID).CopyTo(yMsg, lPos) : lPos += 4
'                        Else : lPos += 4
'                        End If
'                        If .lAssignmentUB > -1 Then
'                            'assignment 1 agentid
'                            If .oAssignments(0).oAgent Is Nothing = False Then
'                                System.BitConverter.GetBytes(.oAssignments(0).oAgent.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
'                            Else
'                                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'                            End If
'                            'assignment 1 skillid
'                            If .oAssignments(0).oSkill Is Nothing = False Then
'                                System.BitConverter.GetBytes(.oAssignments(0).oSkill.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
'                            Else
'                                System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'                            End If
'                            If .lAssignmentUB > 0 Then
'                                'assignment 2 agentid
'                                If .oAssignments(1).oAgent Is Nothing = False Then
'                                    System.BitConverter.GetBytes(.oAssignments(1).oAgent.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
'                                Else
'                                    System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'                                End If
'                                'assignment 2 skillid
'                                If .oAssignments(1).oSkill Is Nothing = False Then
'                                    System.BitConverter.GetBytes(.oAssignments(1).oSkill.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
'                                Else
'                                    System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'                                End If
'                            Else : lPos += 8
'                            End If
'                        Else : lPos += 16
'                        End If

'                    End With
'                Else : lPos += 20
'                End If
'            End If

'            'ok, get the phases
'            yMsg(lPos) = yPhaseCnt : lPos += 1
'            For X As Int32 = 0 To .lPhaseUB
'                If .oPhases(X) Is Nothing = False Then
'                    yMsg(lPos) = .oPhases(X).yPhase : lPos += 1
'                    yMsg(lPos) = CByte(lPhaseCvrCnt(X)) : lPos += 1

'                    For Y As Int32 = 0 To .oPhases(X).lCoverAgentUB
'                        If .oPhases(X).lCoverAgentIdx(Y) <> -1 Then
'                            System.BitConverter.GetBytes(.oPhases(X).lCoverAgentIdx(Y)).CopyTo(yMsg, lPos) : lPos += 4
'                        End If
'                    Next Y
'                End If
'            Next X

'            'now for our goals
'            yMsg(lPos) = yGoalCnt : lPos += 1
'            If .oMission Is Nothing = False Then
'                For X As Int32 = 0 To .oMission.GoalUB
'                    If .oMissionGoals(X) Is Nothing = False AndAlso .oMission.MethodIDs(X) = lMethodID Then
'                        If .oMissionGoals(X).oGoal Is Nothing = False Then
'                            System.BitConverter.GetBytes(.oMissionGoals(X).oGoal.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
'                        Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'                        End If
'                        If .oMissionGoals(X).oSkillSet Is Nothing = False Then
'                            System.BitConverter.GetBytes(.oMissionGoals(X).oSkillSet.SkillSetID).CopyTo(yMsg, lPos) : lPos += 4
'                        Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'                        End If


'                        yMsg(lPos) = CByte(lAssignCnt(X)) : lPos += 1
'                        For Y As Int32 = 0 To .oMissionGoals(X).lAssignmentUB
'                            If .oMissionGoals(X).oAssignments(Y) Is Nothing = False Then
'                                If .oMissionGoals(X).oAssignments(Y).oAgent Is Nothing = False Then
'                                    System.BitConverter.GetBytes(.oMissionGoals(X).oAssignments(Y).oAgent.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
'                                Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'                                End If
'                                If .oMissionGoals(X).oAssignments(Y).oSkill Is Nothing = False Then
'                                    System.BitConverter.GetBytes(.oMissionGoals(X).oAssignments(Y).oSkill.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
'                                Else : System.BitConverter.GetBytes(-1I).CopyTo(yMsg, lPos) : lPos += 4
'                                End If
'                            End If
'                        Next Y
'                    End If
'                Next X
'            End If
'        End With

'        MyBase.moUILib.AddNotification("Mission Plan Submitted.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)

'        Return yMsg
'    End Function

'    Public Function GetMissionPMID() As Int32
'        If moMission Is Nothing = False Then Return moMission.PM_ID
'        Return -1
'    End Function

'    Private Sub btnAssignCover_Click(ByVal sName As String) Handles btnAssignCover.Click
'        If lstGoals.ListIndex = -1 Then
'            MyBase.moUILib.AddNotification("Select a goal to assign cover agents to first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'            Return
'        End If

'        If mlSelectedAgentID = -1 Then
'            MyBase.moUILib.AddNotification("Select an agent to assign and then click Assign.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'        Else
'            Dim oAgent As Agent = Nothing
'            For X As Int32 = 0 To goCurrentPlayer.AgentUB
'                If goCurrentPlayer.AgentIdx(X) = mlSelectedAgentID Then
'                    oAgent = goCurrentPlayer.Agents(X)
'                    Exit For
'                End If
'            Next X

'            'ok, let's see if the agent is assigned to any assignments within the same phase... so we need to get our goal
'            Dim oGoal As Goal = Nothing
'            Dim lGoalID As Int32 = lstGoals.ItemData(lstGoals.ListIndex)
'            If lGoalID = Goal.ml_SAFEHOUSE_GOAL_ID Then
'                oGoal = Goal.GetSafehouseGoal()
'            Else
'                For X As Int32 = 0 To moMission.oMission.GoalUB
'                    If moMission.oMission.Goals(X).ObjectID = lGoalID Then
'                        oGoal = moMission.oMission.Goals(X)
'                        Exit For
'                    End If
'                Next X
'            End If
'            If oGoal Is Nothing Then
'                MyBase.moUILib.AddNotification("Select a goal to assign cover agents to first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                Return
'            End If

'            'ok, now, let's see if the agent has been assigned to a skill in this phase
'            Dim yPhase As Byte = oGoal.MissionPhase
'            For X As Int32 = 0 To moMission.oMission.GoalUB
'                If moMission.oMission.MethodIDs(X) = moMission.lMethodID Then
'                    If moMission.oMission.Goals(X).MissionPhase = yPhase Then
'                        For Y As Int32 = 0 To moMission.oMissionGoals(X).lAssignmentUB
'                            If moMission.oMissionGoals(X).oAssignments(Y) Is Nothing = False Then
'                                If moMission.oMissionGoals(X).oAssignments(Y).oAgent Is Nothing = False Then
'                                    If moMission.oMissionGoals(X).oAssignments(Y).oAgent.ObjectID = oAgent.ObjectID Then
'                                        'ok, not good
'                                        MyBase.moUILib.AddNotification("This agent has been assigned to a task and cannot be a cover agent.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                                        Return
'                                    End If
'                                End If
'                            End If
'                        Next Y
'                    End If
'                End If
'            Next X
'            If yPhase = eMissionPhase.ePreparationTime Then
'                If oGoal.ObjectID <> Goal.ml_SAFEHOUSE_GOAL_ID Then
'                    Dim oPMG As PlayerMissionGoal = moMission.oSafehouseMissionGoal
'                    If oPMG Is Nothing = False Then
'                        For Y As Int32 = 0 To oPMG.lAssignmentUB
'                            If oPMG.oAssignments(Y) Is Nothing = False Then
'                                If oPMG.oAssignments(Y).oAgent Is Nothing = False Then
'                                    If oPMG.oAssignments(Y).oAgent.ObjectID = oAgent.ObjectID Then
'                                        'ok, not good
'                                        MyBase.moUILib.AddNotification("This agent has been assigned to a task and cannot be a cover agent.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'                                        Return
'                                    End If
'                                End If
'                            End If
'                        Next Y
'                    End If
'                End If
'            End If

'            'ok, assign this agent
'            Dim lPhaseIdx As Int32 = -1
'            If moMission.oPhases Is Nothing = False Then
'                For X As Int32 = 0 To moMission.lPhaseUB
'                    If moMission.oPhases(X) Is Nothing = False AndAlso moMission.oPhases(X).yPhase = yPhase Then
'                        lPhaseIdx = X
'                        Exit For
'                    End If
'                Next
'            End If
'            If lPhaseIdx = -1 Then
'                moMission.lPhaseUB = 0
'                ReDim moMission.oPhases(moMission.lPhaseUB)
'                moMission.oPhases(moMission.lPhaseUB) = New PlayerMissionPhase
'                lPhaseIdx = moMission.lPhaseUB
'                moMission.oPhases(lPhaseIdx).yPhase = CType(yPhase, eMissionPhase)
'                moMission.oPhases(lPhaseIdx).oParent = moMission
'            End If
'            moMission.oPhases(lPhaseIdx).AddCoverAgent(oAgent, 0)
'        End If

'        RefreshCoverAgentList()

'    End Sub

'    Private Sub btnRemoveCover_Click(ByVal sName As String) Handles btnRemoveCover.Click

'        If lstCoverAgents.ListIndex > -1 Then
'            Dim lAgentID As Int32 = lstCoverAgents.ItemData(lstCoverAgents.ListIndex)
'            If lstGoals.ListIndex > -1 Then
'                Dim lGoalID As Int32 = lstGoals.ItemData(lstGoals.ListIndex)
'                Dim yPhase As Byte = 0
'                If moMission Is Nothing = False AndAlso moMission.oMission Is Nothing = False Then

'                    Dim oGoal As Goal = Nothing
'                    If lGoalID = Goal.ml_SAFEHOUSE_GOAL_ID Then
'                        yPhase = eMissionPhase.ePreparationTime
'                    Else
'                        For X As Int32 = 0 To moMission.oMission.GoalUB
'                            If moMission.lMethodID = moMission.oMission.MethodIDs(X) AndAlso moMission.oMission.Goals(X).ObjectID = lGoalID Then
'                                yPhase = moMission.oMission.Goals(X).MissionPhase
'                                Exit For
'                            End If
'                        Next X
'                    End If

'                    For X As Int32 = 0 To moMission.lPhaseUB
'                        If moMission.oPhases(X) Is Nothing = False AndAlso moMission.oPhases(X).yPhase = yPhase Then
'                            moMission.oPhases(X).RemoveCoverAgent(lAgentID)
'                            Exit For
'                        End If
'                    Next X
'                End If
'            End If
'        End If

'        RefreshCoverAgentList()
'    End Sub

'    Private Sub RefreshCoverAgentList()
'        lstCoverAgents.Clear()
'        If lstGoals.ListIndex > -1 Then
'            Dim lGoalID As Int32 = lstGoals.ItemData(lstGoals.ListIndex)
'            Dim yPhase As Byte = 0
'            If moMission Is Nothing = False AndAlso moMission.oMission Is Nothing = False Then

'                If lGoalID = Goal.ml_SAFEHOUSE_GOAL_ID Then
'                    yPhase = eMissionPhase.ePreparationTime
'                Else
'                    For X As Int32 = 0 To moMission.oMission.GoalUB
'                        If moMission.lMethodID = moMission.oMission.MethodIDs(X) AndAlso moMission.oMission.Goals(X).ObjectID = lGoalID Then
'                            yPhase = moMission.oMission.Goals(X).MissionPhase
'                            Exit For
'                        End If
'                    Next X
'                End If

'                Dim oList() As Agent = Nothing
'                Dim lListIdx() As Int32 = Nothing
'                Dim lListUB As Int32 = -1
'                For X As Int32 = 0 To moMission.lPhaseUB
'                    If moMission.oPhases(X) Is Nothing = False AndAlso moMission.oPhases(X).yPhase = yPhase Then
'                        For Y As Int32 = 0 To moMission.oPhases(X).lCoverAgentUB
'                            If moMission.oPhases(X).lCoverAgentIdx(Y) <> -1 Then
'                                lListUB += 1
'                                ReDim Preserve oList(lListUB)
'                                ReDim Preserve lListIdx(lListUB)
'                                oList(lListUB) = moMission.oPhases(X).oCoverAgents(Y)
'                                lListIdx(lListUB) = moMission.oPhases(X).lCoverAgentIdx(Y)
'                            End If
'                        Next Y
'                        Exit For
'                    End If
'                Next X

'                If lListUB <> -1 Then
'                    Dim lSorted() As Int32 = GetSortedIndexArray(oList, lListIdx, GetSortedIndexArrayType.eAgentName)
'                    If lSorted Is Nothing = False Then
'                        For X As Int32 = 0 To lListUB
'                            lstCoverAgents.AddItem(oList(lSorted(X)).sAgentName, False)
'                            lstCoverAgents.ItemData(lstCoverAgents.NewIndex) = oList(lSorted(X)).ObjectID
'                            lstCoverAgents.ItemData2(lstCoverAgents.NewIndex) = ObjectType.eAgent
'                        Next X
'                    End If
'                End If
'            End If
'        End If
'    End Sub

'    'Private Sub chkSafehouse_Click() Handles chkSafehouse.Click
'    '    If moMission Is Nothing = False AndAlso moMission.oMission Is Nothing = False Then
'    '        cboMethod_ItemSelected(cboMethod.ListIndex)
'    '    Else
'    '        MyBase.moUILib.AddNotification("Select a mission first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
'    '        chkSafehouse.Value = False
'    '    End If
'    'End Sub


'    Private Sub frmMission_WindowMoved() Handles Me.WindowMoved
'        If mbLoading = False Then
'            muSettings.AgentMissionCreateX = Me.Left
'            muSettings.AgentMissionCreateY = Me.Top
'        End If
'    End Sub

'    Private Sub tvwGoals_NodeIconClicked(ByRef oNode As UITreeView.UITreeViewItem) Handles tvwGoals.NodeIconClicked
'        If oNode.bClickIcon = False Then Return
'        'oNode.IconRect = grc_UI(elInterfaceRectangle.eCheck_Unchecked)

'    End Sub


'    Private Sub tvwGoals_NodeSelected(ByRef oNode As UITreeView.UITreeViewItem) Handles tvwGoals.NodeSelected
'        If oNode Is Nothing = True Then Return
'        Dim lTeir As Int32 = oNode.lItemData3

'        Dim oMission As Mission = GetSelectedMission()
'        If oMission Is Nothing Then Return
'        Dim lMethodID As Int32 = -1
'        If cboMethod.ListIndex <> -1 Then lMethodID = cboMethod.ItemData(cboMethod.ListIndex)
'        If lMethodID = -1 Then Return

'        If lTeir = 2 Then
'            'skillset was clicked

'        End If

'        If lTeir = 1 Then
'            lblGoalName.Caption = ""
'            lblGoalTime.Caption = ""
'            lblGoalRisk.Caption = ""
'            'txtGoalDesc.Caption = ""
'            lblGoalPhase.Caption = ""

'            Dim lGoalID As Int32 = oNode.lItemData
'            Dim oGoal As Goal = Nothing

'            If lGoalID = Goal.ml_SAFEHOUSE_GOAL_ID Then
'                oGoal = Goal.GetSafehouseGoal()
'                If moMission.oSafehouseMissionGoal Is Nothing Then
'                    moMission.oSafehouseMissionGoal = New PlayerMissionGoal()
'                    With moMission.oSafehouseMissionGoal
'                        .oGoal = oGoal
'                        .oMission = moMission
'                    End With
'                End If
'            Else
'                For X As Int32 = 0 To oMission.GoalUB
'                    If oMission.MethodIDs(X) = lMethodID AndAlso oMission.Goals(X).ObjectID = lGoalID Then

'                        If NewTutorialManager.TutorialOn = True Then
'                            NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eCreateMissionGoalSelected, lGoalID, -1, -1, "")
'                        End If

'                        oGoal = oMission.Goals(X)
'                        Exit For
'                    End If
'                Next X
'            End If

'            If oGoal Is Nothing = False Then
'                'ok, found  it
'                With oGoal
'                    lblGoalName.Caption = "Goal: " & .sGoalName
'                    lblGoalTime.Caption = "Goal Time: about " & GetDurationFromSeconds(.BaseTime, True)
'                    lblGoalRisk.Caption = "Risk Assessment: " & Goal.GetRiskAssessment(.RiskOfDetection)
'                    lblGoalRisk.ForeColor = Goal.GetRiskCaptionColor(.RiskOfDetection)
'                    'txtGoalDesc.Caption = .sGoalDesc
'                    Select Case .MissionPhase
'                        Case eMissionPhase.eFlippingTheSwitch
'                            lblGoalPhase.Caption = "Phase: Execution"
'                        Case eMissionPhase.ePreparationTime
'                            lblGoalPhase.Caption = "Phase: Preparation"
'                        Case eMissionPhase.eReinfiltrationPhase
'                            lblGoalPhase.Caption = "Phase: Reinfiltration"
'                        Case eMissionPhase.eSettingTheStage
'                            lblGoalPhase.Caption = "Phase: Setting the Stage"
'                    End Select
'                End With
'            End If
'        End If
'    End Sub
'End Class
#End Region