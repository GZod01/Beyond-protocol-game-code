Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmAgentMain
    Inherits UIWindow


    Private Enum AgentEffectType As Byte
        eMorale = 0
        eProduction = 1
        eCargoBay = 2
        eHangarBay = 3
        eTradeIncome = 4
        'Per colony, per morale modifier type
        eColonyTaxMorale = 5
        eColonyHousingMorale = 6
        eColonyUnemploymentMorale = 7

        eSiphonTradepost = 8
        eSiphonMining = 9
        eSiphonCorporation = 10
        eGovernorMorale = 11
        eIncreasePowerNeed = 12
    End Enum

    Private Structure ImposedAgentEffects
        Public lTargetID As Int32
        Public lCyclesRemaining As Int32
        Public lSetOnCycle As Int32
        Public lAmount As Int32
        Public yType As AgentEffectType
        Public bAmountAsPerc As Boolean
        Public sSpecificName As String

        Public Function GetListBoxText() As String
            Dim sPlayerName As String = "Enoch Dagor" ' GetCacheObjectValue(lTargetID, ObjectType.ePlayer)
            Dim lActualCyclesRemain As Int32 = (lCyclesRemaining - (glCurrentCycle - lSetOnCycle)) \ 30

            Dim sAmt As String = lAmount.ToString()
            If bAmountAsPerc = True Then sAmt &= "%"

            Dim sFinal As String = ""
            If lActualCyclesRemain < 30 Then
                sFinal = "Expiring...".PadRight(12, " "c)
            Else
                Dim sTemp As String = GetDurationFromSeconds(lActualCyclesRemain, False).PadRight(12, " "c)
                
                'Dim lTmpVal As Int32 = sTemp.IndexOf(" ")
                'If lTmpVal > -1 Then
                '    lTmpVal = sTemp.IndexOf(" ", lTmpVal + 1)
                '    If lTmpVal > -1 Then
                '        sTemp = sTemp.Substring(0, lTmpVal)
                '    End If
                'End If
                ''sTemp = sTemp.Substring(0, sTemp.Length - 3)
                'sFinal = sTemp.PadLeft(7, " "c).PadRight(8, " "c)
            End If
            sFinal &= (GetEffectTypeName(yType) & " (" & sAmt & ")").PadRight(34, " "c)
            sFinal &= sSpecificName & " (" & sPlayerName & ")"

            Return sFinal
        End Function

        Private Shared Function GetEffectTypeName(ByVal yType As AgentEffectType) As String
            Select Case yType
                Case AgentEffectType.eCargoBay
                    Return "Clutter Cargo"
                Case AgentEffectType.eColonyHousingMorale
                    Return "Incite Unrest Housing"
                Case AgentEffectType.eColonyTaxMorale
                    Return "Incite Unrest Taxes"
                Case AgentEffectType.eColonyUnemploymentMorale
                    Return "Incite Unrest Unemployment"
                Case AgentEffectType.eGovernorMorale
                    Return "Assassinate Governor"
                Case AgentEffectType.eHangarBay
                    Return "Slow Docking"
                Case AgentEffectType.eIncreasePowerNeed
                    Return "Increase Power Need"
                Case AgentEffectType.eMorale
                    Return "Colony Morale"
                Case AgentEffectType.eProduction
                    Return "Production Rate"
                Case AgentEffectType.eSiphonCorporation
                    Return "Siphon Corporation"
                Case AgentEffectType.eSiphonMining
                    Return "Siphon Minerals"
                Case AgentEffectType.eSiphonTradepost
                    Return "Siphon Trade"
                Case AgentEffectType.eTradeIncome
                    Return "Trade Treaty Income"
            End Select
            Return "Unknown"
        End Function
    End Structure
    Private muEffects(-1) As ImposedAgentEffects

    Private lblAgents As UILabel
    Private lblIntel As UILabel
	Private lblTitle As UILabel
	Private lnDiv1 As UILine

    Private WithEvents optCurrentMissions As UIOption
    Private WithEvents optCurrentEffects As UIOption

    Private WithEvents lstMissions As UIListBox
    Private WithEvents btnViewMission As UIButton
    Private WithEvents btnCreateMission As UIButton
    Private WithEvents lstAgents As UIListBox
    Private WithEvents lstIntel As UIListBox
    Private WithEvents btnViewAgent As UIButton
    Private WithEvents btnDismissAgent As UIButton
    Private WithEvents btnViewIntel As UIButton
    Private WithEvents btnClose As UIButton
	Private WithEvents btnHelp As UIButton

	Private WithEvents chkFilterArchivedMissions As UICheckBox
	Private WithEvents chkSetFilterMission As UICheckBox
	Private WithEvents chkFilterArchivedIntel As UICheckBox
	Private mbIgnoreSetFilter As Boolean = False

	Private WithEvents cboIntelFilter As UIComboBox
    Private WithEvents optMyAgents As UIOption
    Private WithEvents optEnemyAgents As UIOption

    Private WithEvents btnExport As UIButton
    Private myFilterValue As Byte = 0       '0 for mine, 1 for enemy

    Public Shared moAgentIcons As Texture = Nothing

    Private myViewState As Byte = 0     '0 for mission view, 1 for effect view
    Private mbLoading As Boolean = True
    Private mlLastAgentRequestTime As Int32
    Private myAgentSort As Int32 = 0

    Private mlAgentCount As Int32 = 0
    Private mbHasUnknownsIntel As Boolean
    Private mlTimeOpened As Int32 = -1

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenAgentMainWindow)

        'frmAgents initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eAgentMain
            .ControlName = "frmAgentMain"


            .Width = 800
            .Height = 600
            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1
            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.AgentMainX
                lTop = muSettings.AgentMainY
            End If
            If lLeft < 0 Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 400
            If lTop < 0 Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 300
            If lLeft + .Width > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If lTop + .Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Left = lLeft
            .Top = lTop

            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
        End With

        'optCurrentMissions initial props
        optCurrentMissions = New UIOption(oUILib)
        With optCurrentMissions
            .ControlName = "optCurrentMissions"
            .Left = 10
            .Top = 35
            .Width = 130
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Current Missions"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
        End With
        Me.AddChild(CType(optCurrentMissions, UIControl))

        'optCurrentEffects initial props
        optCurrentEffects = New UIOption(oUILib)
        With optCurrentEffects
            .ControlName = "optCurrentEffects"
            .Left = 235
            .Top = 35
            .Width = 311
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Effects Imposed On Others By Your Agents"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optCurrentEffects, UIControl))

        'lstMissions initial props
        lstMissions = New UIListBox(oUILib)
        With lstMissions
            .ControlName = "lstMissions"
            .Left = 10
            .Top = 55
            .Width = 775
            .Height = 210
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstMissions, UIControl))

        'btnViewMission initial props
        btnViewMission = New UIButton(oUILib)
        With btnViewMission
            .ControlName = "btnViewMission"
            .Left = 665
            .Top = 270
            .Width = 120
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "View Details"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnViewMission, UIControl))

        'btnCreateMission initial props
        btnCreateMission = New UIButton(oUILib)
        With btnCreateMission
            .ControlName = "btnCreateMission"
            .Width = 120
            .Height = 24
            .Left = btnViewMission.Left - .Width - 1
            .Top = 270
            .Enabled = True
            .Visible = True
            .Caption = "Create Mission"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnCreateMission, UIControl))

        'lblAgents initial props
        lblAgents = New UILabel(oUILib)
        With lblAgents
            .ControlName = "lblAgents"
            .Left = 10
            .Top = 300
            .Width = 115
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Agents"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblAgents, UIControl))

        'lblIntel initial props
        lblIntel = New UILabel(oUILib)
        With lblIntel
            .ControlName = "lblIntel"
            .Left = 455
            .Top = 300
            .Width = 111
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Intelligence for"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
        End With
        Me.AddChild(CType(lblIntel, UIControl))

        'lstAgents initial props
        lstAgents = New UIListBox(oUILib)
        With lstAgents
            .ControlName = "lstAgents"
            .Left = 10
            .Top = 320
            .Width = 435
            .Height = 240
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
		Me.AddChild(CType(lstAgents, UIControl))

		'lstIntel initial props
        lstIntel = New UIListBox(oUILib)
        With lstIntel
            .ControlName = "lstIntel"
            .Left = 455
            .Top = 320
            .Width = 330
            .Height = 240
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
            .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
        End With
        Me.AddChild(CType(lstIntel, UIControl))

        'btnViewAgent initial props
        btnViewAgent = New UIButton(oUILib)
        With btnViewAgent
            .ControlName = "btnViewAgent"
            .Left = lstAgents.Left + lstAgents.Width - 100
            .Top = lstAgents.Top + lstAgents.Height + 5
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "View Agent"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnViewAgent, UIControl))

        'btnViewIntel initial props
        btnViewIntel = New UIButton(oUILib)
        With btnViewIntel
            .ControlName = "btnViewIntel"
            .Width = 120
            .Height = 24
            .Left = lstIntel.Left + lstIntel.Width - .Width
            .Top = lstIntel.Top + lstIntel.Height + 5
            .Enabled = True
            .Visible = True
            .Caption = "View Report"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnViewIntel, UIControl))

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 10
            .Top = 3
            .Width = 260
            .Height = 24
            .Enabled = True
            .Visible = True
            .Caption = "Agent and Mission Management"
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


        'lnDiv1 initial props
        lnDiv1 = New UILine(oUILib)
        With lnDiv1
            .ControlName = "lnDiv1"
            .Left = 2 '1
            .Top = 27
            .Width = Me.Width - 4 ' 799
            .Height = 0
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
        End With
		Me.AddChild(CType(lnDiv1, UIControl))

        optMyAgents = New UIOption(oUILib)
        With optMyAgents
            .ControlName = "optMyAgents"
            .Left = 10
            .Top = lstAgents.Top + lstAgents.Height + 6
            .Width = 85
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "My Agents"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = True
        End With
        Me.AddChild(CType(optMyAgents, UIControl))

        optEnemyAgents = New UIOption(oUILib)
        With optEnemyAgents
            .ControlName = "optEnemyAgents"
            '.Left = 10
            '.Top = optMyAgents.Top + optMyAgents.Height - 2
            .Left = optMyAgents.Left + optMyAgents.Width + 10
            .Top = optMyAgents.Top

            .Width = 112
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Enemy Agents"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(6, DrawTextFormat)
            .Value = False
        End With
        Me.AddChild(CType(optEnemyAgents, UIControl))

        'btnDismissAgent initial props
        btnDismissAgent = New UIButton(oUILib)
        With btnDismissAgent
            .ControlName = "btnDismissAgent"
            .Width = 100
            .Height = 24
            .Left = btnViewAgent.Left - .Width - 1 '180 'btnShow.Left + btnShow.Width + 5
            .Top = lstAgents.Top + lstAgents.Height + 5
            .Enabled = True
            .Visible = True
            .Caption = "Dismiss"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnDismissAgent, UIControl))

        'cboIntelFilter initial props
		cboIntelFilter = New UIComboBox(oUILib)
		With cboIntelFilter
			.ControlName = "cboIntelFilter"
			.Left = lblIntel.Left + lblIntel.Width + 5
			.Top = lblIntel.Top - 2
			.Width = lstIntel.Width - (.Left - lstIntel.Left)
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
		Me.AddChild(CType(cboIntelFilter, UIControl))

		'chkFilterArchivedMissions initial props
		chkFilterArchivedMissions = New UICheckBox(oUILib)
		With chkFilterArchivedMissions
			.ControlName = "chkFilterArchivedMissions"
			.Left = lstMissions.Left
			.Top = lstMissions.Top + lstMissions.Height + 10
			.Width = 102
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Filter Archived"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = True
			.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
		End With
		Me.AddChild(CType(chkFilterArchivedMissions, UIControl))

		'chkSetFilterMission initial props
		chkSetFilterMission = New UICheckBox(oUILib)
		With chkSetFilterMission
			.ControlName = "chkSetFilterMission"
			.Left = chkFilterArchivedMissions.Left + chkFilterArchivedMissions.Width + 15
			.Top = lstMissions.Top + lstMissions.Height + 10
			.Width = 142
			.Height = 18
            .Enabled = Not gbAliased
			.Visible = True
			.Caption = "Set Mission Archived"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = False
			.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
		End With
		Me.AddChild(CType(chkSetFilterMission, UIControl))

		'chkFilterArchivedIntel initial props
		chkFilterArchivedIntel = New UICheckBox(oUILib)
		With chkFilterArchivedIntel
			.ControlName = "chkFilterArchivedIntel"
			.Left = lstIntel.Left
			.Top = lstIntel.Top + lstIntel.Height + 5
			.Width = 102
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Filter Archived"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(6, DrawTextFormat)
			.Value = True
			.DisplayType = UICheckBox.eCheckTypes.eSmallCheck
		End With
		Me.AddChild(CType(chkFilterArchivedIntel, UIControl))


        MyBase.moUILib.RemoveWindow(Me.ControlName)
        If HasAliasedRights(AliasingRights.eViewAgents) = True Then
            MyBase.moUILib.AddWindow(Me)
		Else
			Me.Visible = False
			goUILib.AddNotification("You lack the rights to view the agents interface.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
        End If

        'Call this method because it will return a scratch surface
        If moAgentIcons Is Nothing OrElse moAgentIcons.Disposed = True Then moAgentIcons = goResMgr.LoadScratchTexture("AgentIcons.dds", "Interface.pak")
		Me.lstAgents.oIconTexture = moAgentIcons
        Me.lstMissions.oIconTexture = moAgentIcons

        'remove this when ready
		'		CreateTestData()

        RefreshMissionList()
		RefreshAgentList()
		FullRefreshIntelFilterList()
        cboIntelFilter.FindComboItemData(-1)        'set our filter to all
        mlTimeOpened = glCurrentCycle
        mbLoading = False
    End Sub

    Private Sub RefreshMissionList()
        lstMissions.Clear()
        lstMissions.oIconTexture = moAgentIcons
        'lstMissions.sHeaderRow = "  Mission".PadRight(52, " "c) & "Target".PadRight(25, " "c) & "Status"
        lstMissions.sHeaderRow = "  Mission".PadRight(37, " "c) & "Target".PadRight(40, " "c) & "Status"

		If goCurrentPlayer Is Nothing = False Then
			Dim bFilter As Boolean = chkFilterArchivedMissions.Value

			Dim lSorted() As Int32 = GetSortedIndexArray(goCurrentPlayer.PlayerMissions, goCurrentPlayer.PlayerMissionIdx, GetSortedIndexArrayType.ePlayerMissionGetListboxText)
			For X As Int32 = 0 To goCurrentPlayer.PlayerMissionUB
				If goCurrentPlayer.PlayerMissionIdx(lSorted(X)) <> -1 Then
					With goCurrentPlayer.PlayerMissions(lSorted(X))
						If .bArchived = False OrElse bFilter = False Then
							Dim sText As String = .GetListBoxText()
							lstMissions.AddItem(sText, False)
							lstMissions.ItemData(lstMissions.NewIndex) = .PM_ID
							lstMissions.ItemData2(lstMissions.NewIndex) = .yCurrentPhase
							lstMissions.ApplyIconOffset(lstMissions.NewIndex) = True
							lstMissions.rcIconRectangle(lstMissions.NewIndex) = GetMissionStatusIcon(.yCurrentPhase)
							lstMissions.IconForeColor(lstMissions.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 255, 255)
						End If
					End With
				End If
			Next X
		End If
        lstMissions.RenderIcons = True
    End Sub

    Private Function GetMissionStatusIcon(ByVal lStatus As eMissionPhase) As Rectangle

        'Mission Status:
        '===============
        'In Planning - lightbulb - 0,0,16,16
        'Cancelled - red X - 16, 0, 16, 16
        'Preparation - wrench - 32, 0, 16, 16
        'In Progress - lightning - 48, 0, 16, 16
        'Executing - double lightning - 0, 16, 16, 16
        'failed - negative sign - 16, 16, 16, 16
		'success - positive sign - 32, 16, 16, 16

		If (lStatus And eMissionPhase.eMissionPaused) <> 0 Then
			Return New Rectangle(64, 16, 16, 16)
		End If

        Select Case lStatus
            Case eMissionPhase.eCancelled
                Return New Rectangle(16, 0, 16, 16)
            Case eMissionPhase.ePreparationTime
                Return New Rectangle(32, 0, 16, 16)
            Case eMissionPhase.eSettingTheStage
                Return New Rectangle(48, 0, 16, 16)
            Case eMissionPhase.eFlippingTheSwitch
                Return New Rectangle(0, 16, 16, 16)
            Case eMissionPhase.eMissionOverFailure
                Return New Rectangle(16, 16, 16, 16)
            Case eMissionPhase.eMissionOverSuccess
				Return New Rectangle(32, 16, 16, 16)
			Case eMissionPhase.eWaitingToExecute
				Return New Rectangle(64, 16, 16, 16)
			Case Else 'eMissionPhase.eInPlanning
				Return New Rectangle(0, 0, 16, 16)
		End Select

    End Function

    Private Sub RefreshAgentList()
        lstAgents.Clear()
        lstAgents.oIconTexture = moAgentIcons
        mlAgentCount = 0

		If goCurrentPlayer Is Nothing = False Then
			If myFilterValue = 0 Then
                lstAgents.sHeaderRow = "  Agent Name".PadRight(23, " "c) & "Status"
                Dim lSorted() As Int32
                If myAgentSort = 0 OrElse myAgentSort = 1 OrElse myAgentSort = -1 Then
                    lSorted = GetSortedIndexArray(goCurrentPlayer.Agents, goCurrentPlayer.AgentIdx, GetSortedIndexArrayType.eAgentName, myAgentSort = -1)
                ElseIf myAgentSort = 2 OrElse myAgentSort = -2 Then
                    lSorted = GetSortedIndexArray(goCurrentPlayer.Agents, goCurrentPlayer.AgentIdx, GetSortedIndexArrayType.eAgentStatus, myAgentSort = -2)
                Else
                    lSorted = GetSortedIndexArray(goCurrentPlayer.Agents, goCurrentPlayer.AgentIdx, GetSortedIndexArrayType.eAgentName)
                End If
                If lSorted Is Nothing = False Then
                    For X As Int32 = 0 To lSorted.GetUpperBound(0)
                        If goCurrentPlayer.AgentIdx(lSorted(X)) <> -1 Then
                            With goCurrentPlayer.Agents(lSorted(X))
                                Dim sText As String = .sAgentName.PadRight(21, " "c)
                                sText &= Agent.GetStatusText(.lAgentStatus, .lTargetID, .iTargetTypeID, .InfiltrationType)
                                lstAgents.AddItem(sText, False)
                                lstAgents.ItemData(lstAgents.NewIndex) = .ObjectID
                                lstAgents.ApplyIconOffset(lstAgents.NewIndex) = True
                                lstAgents.rcIconRectangle(lstAgents.NewIndex) = .GetStatusIconRect
                                lstAgents.ItemData2(lstAgents.NewIndex) = .lAgentStatus
                                lstAgents.IconForeColor(lstAgents.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 255, 255)
                                mlAgentCount += 1
                            End With
                        End If
                    Next X
                End If
            Else
                lstAgents.sHeaderRow = "  Agent Name".PadRight(23, " "c) & "Owner"

                Dim lSorted() As Int32 = GetSortedIndexArray(goCurrentPlayer.CapturedAgents, goCurrentPlayer.CapturedAgentIdx, GetSortedIndexArrayType.eCapturedAgentName)
                If lSorted Is Nothing = False Then
                    For X As Int32 = 0 To lSorted.GetUpperBound(0)
                        If goCurrentPlayer.CapturedAgentIdx(lSorted(X)) <> -1 Then
                            With goCurrentPlayer.CapturedAgents(lSorted(X))
                                Dim sText As String = .sAgentName.PadRight(21, " "c)
                                sText &= GetCacheObjectValue(.lOwnerID, ObjectType.ePlayer)
                                lstAgents.AddItem(sText, False)
                                lstAgents.ItemData(lstAgents.NewIndex) = .ObjectID
                                lstAgents.ApplyIconOffset(lstAgents.NewIndex) = True
                                lstAgents.rcIconRectangle(lstAgents.NewIndex) = New Rectangle(0, 48, 16, 16)
                                lstAgents.ItemData2(lstAgents.NewIndex) = AgentStatus.HasBeenCaptured
                                lstAgents.IconForeColor(lstAgents.NewIndex) = System.Drawing.Color.FromArgb(255, 255, 255, 255)
                                mlAgentCount += 1
                            End With
                        End If
                    Next X
                End If
            End If
            End If

        lblAgents.Caption = "Agents (" & mlAgentCount.ToString & ")"
		lstAgents.RenderIcons = True
    End Sub

    Public Sub RefreshIntelList()
        'ok, we fill the list based on... PlayerItemIntel, PlayerIntel, PlayerTechKnowledge
        lstIntel.Clear()
        mbHasUnknownsIntel = False

        Dim lFilterID As Int32 = -1
        If cboIntelFilter.ListIndex > -1 Then lFilterID = cboIntelFilter.ItemData(cboIntelFilter.ListIndex)

        Dim bFilter As Boolean = chkFilterArchivedIntel.Value

        'Ok, so... we have the following intel... we'll sort it by Group Name then by Name
        Dim lSorted() As Int32 = Nothing
        Dim sSortedVal() As String = Nothing
        Dim lGroupVal() As Int32 = Nothing
        Dim lSortedUB As Int32 = -1

        'Item intel
        For X As Int32 = 0 To glItemIntelUB
            If glItemIntelIdx(X) <> -1 Then
                Dim lIdx As Int32 = -1
                Dim oItem As PlayerItemIntel = goItemIntel(X)
                If oItem Is Nothing = False AndAlso oItem.lPlayerID = glPlayerID AndAlso (lFilterID = -1 OrElse oItem.lOtherPlayerID = lFilterID) AndAlso (bFilter = False OrElse oItem.bArchived = False) Then
                    Dim lNewGroup As Int32 = PlayerItemIntel.GetGroupValue(oItem.iItemTypeID)

                    'Now, check if the group val has been added already
                    lIdx = -1
                    For Y As Int32 = 0 To lSortedUB
                        If lGroupVal(Y) = lNewGroup AndAlso lSorted(Y) = -1 Then
                            lIdx = Y
                            Exit For
                        End If
                    Next Y
                    If lIdx = -1 Then
                        For Y As Int32 = 0 To lSortedUB
                            If lGroupVal(Y) > lNewGroup Then
                                lIdx = Y
                                Exit For
                            End If
                        Next Y
                        'make room
                        lSortedUB += 1
                        ReDim Preserve lSorted(lSortedUB)
                        ReDim Preserve sSortedVal(lSortedUB)
                        ReDim Preserve lGroupVal(lSortedUB)
                        If lIdx = -1 Then
                            lSorted(lSortedUB) = -1
                            lGroupVal(lSortedUB) = lNewGroup
                            sSortedVal(lSortedUB) = PlayerItemIntel.GetGroupValueName(oItem.iItemTypeID)
                        Else
                            For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                                lSorted(Y) = lSorted(Y - 1)
                                lGroupVal(Y) = lGroupVal(Y - 1)
                                sSortedVal(Y) = sSortedVal(Y - 1)
                            Next Y
                            lSorted(lIdx) = -1
                            lGroupVal(lIdx) = lNewGroup
                            sSortedVal(lIdx) = PlayerItemIntel.GetGroupValueName(oItem.iItemTypeID)
                        End If
                    End If

                    'Now, find our place to place this item...
                    lIdx = -1
                    Dim sName As String = GetCacheObjectValueCheckUnknowns(oItem.lItemID, oItem.iItemTypeID, mbHasUnknownsIntel).ToUpper
                    For Y As Int32 = 0 To lSortedUB
                        If (lGroupVal(Y) = lNewGroup AndAlso lSorted(Y) <> -1 AndAlso sSortedVal(Y) > sName) OrElse lGroupVal(Y) > lNewGroup Then
                            lIdx = Y
                            Exit For
                        End If
                    Next Y
                    lSortedUB += 1
                    ReDim Preserve lSorted(lSortedUB)
                    ReDim Preserve sSortedVal(lSortedUB)
                    ReDim Preserve lGroupVal(lSortedUB)
                    If lIdx = -1 Then
                        lSorted(lSortedUB) = X
                        lGroupVal(lSortedUB) = lNewGroup
                        sSortedVal(lSortedUB) = sName
                    Else
                        For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                            lSorted(Y) = lSorted(Y - 1)
                            lGroupVal(Y) = lGroupVal(Y - 1)
                            sSortedVal(Y) = sSortedVal(Y - 1)
                        Next Y
                        lSorted(lIdx) = X
                        lGroupVal(lIdx) = lNewGroup
                        sSortedVal(lIdx) = sName
                    End If
                End If
            End If
        Next X

        '  TECHNOLOGY DESIGNS  (PlayerTechKnowledge)
        '  Alloy Designs
        '  Armor Designs
        '  Engine Designs
        '  Prototype Designs?
        '  Radar Designs
        '  Shield Designs
        '  Weapon Designs
        For X As Int32 = 0 To glPlayerTechKnowledgeUB
            If glPlayerTechKnowledgeIdx(X) <> -1 Then
                Dim lIdx As Int32 = -1
                Dim oItem As PlayerTechKnowledge = goPlayerTechKnowledge(X)
                If oItem Is Nothing = False AndAlso oItem.oPlayer Is Nothing = False AndAlso oItem.oPlayer.ObjectID = glPlayerID AndAlso oItem.oTech Is Nothing = False AndAlso (oItem.oTech.OwnerID = lFilterID OrElse lFilterID = -1) AndAlso (oItem.bArchived = False OrElse bFilter = False) Then
                    Dim lNewGroup As Int32 = PlayerTechKnowledge.GetGroupValue(oItem.oTech.ObjTypeID)

                    'Now, check if the group val has been added already
                    lIdx = -1
                    For Y As Int32 = 0 To lSortedUB
                        If lGroupVal(Y) = lNewGroup AndAlso lSorted(Y) = -1 Then
                            lIdx = Y
                            Exit For
                        End If
                    Next Y
                    If lIdx = -1 Then
                        For Y As Int32 = 0 To lSortedUB
                            If lGroupVal(Y) > lNewGroup Then
                                lIdx = Y
                                Exit For
                            End If
                        Next Y
                        'make room
                        lSortedUB += 1
                        ReDim Preserve lSorted(lSortedUB)
                        ReDim Preserve sSortedVal(lSortedUB)
                        ReDim Preserve lGroupVal(lSortedUB)
                        If lIdx = -1 Then
                            lSorted(lSortedUB) = -1
                            lGroupVal(lSortedUB) = lNewGroup
                            sSortedVal(lSortedUB) = PlayerTechKnowledge.GetGroupName(oItem.oTech.ObjTypeID)
                        Else
                            For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                                lSorted(Y) = lSorted(Y - 1)
                                lGroupVal(Y) = lGroupVal(Y - 1)
                                sSortedVal(Y) = sSortedVal(Y - 1)
                            Next Y
                            lSorted(lIdx) = -1
                            lGroupVal(lIdx) = lNewGroup
                            sSortedVal(lIdx) = PlayerTechKnowledge.GetGroupName(oItem.oTech.ObjTypeID)
                        End If
                    End If

                    'Now, find our place to place this item...
                    lIdx = -1
                    Dim sName As String = oItem.oTech.GetComponentName().ToUpper
                    For Y As Int32 = 0 To lSortedUB
                        'If lGroupVal(Y) >= lNewGroup AndAlso lSorted(Y) <> -1 AndAlso sSortedVal(Y) > sName Then
                        If (lGroupVal(Y) = lNewGroup AndAlso lSorted(Y) <> -1 AndAlso sSortedVal(Y) > sName) OrElse lGroupVal(Y) > lNewGroup Then
                            lIdx = Y
                            Exit For
                        End If
                    Next Y
                    lSortedUB += 1
                    ReDim Preserve lSorted(lSortedUB)
                    ReDim Preserve sSortedVal(lSortedUB)
                    ReDim Preserve lGroupVal(lSortedUB)
                    If lIdx = -1 Then
                        lSorted(lSortedUB) = X
                        lGroupVal(lSortedUB) = lNewGroup
                        sSortedVal(lSortedUB) = sName
                    Else
                        For Y As Int32 = lSortedUB To lIdx + 1 Step -1
                            lSorted(Y) = lSorted(Y - 1)
                            lGroupVal(Y) = lGroupVal(Y - 1)
                            sSortedVal(Y) = sSortedVal(Y - 1)
                        Next Y
                        lSorted(lIdx) = X
                        lGroupVal(lIdx) = lNewGroup
                        sSortedVal(lIdx) = sName
                    End If
                End If
            End If
        Next X

        'Now, fill our list...
        If lSorted Is Nothing = False Then
            For X As Int32 = 0 To lSorted.GetUpperBound(0)
                'if lSorted is -1
                If lSorted(X) = -1 Then
                    'ok, it is a title
                    lstIntel.AddItem(sSortedVal(X), True)
                    lstIntel.ItemData(lstIntel.NewIndex) = -1
                    lstIntel.ItemLocked(lstIntel.NewIndex) = True
                Else
                    If lGroupVal(X) < 100 OrElse lGroupVal(X) = 200 Then
                        'item intel
                        Dim oItem As PlayerItemIntel = goItemIntel(lSorted(X))
                        If oItem Is Nothing = False Then
                            Dim sText As String = GetCacheObjectValueCheckUnknowns(oItem.lItemID, oItem.iItemTypeID, mbHasUnknownsIntel).PadRight(21, " "c)
                            sText &= GetCacheObjectValueCheckUnknowns(oItem.lOtherPlayerID, ObjectType.ePlayer, mbHasUnknownsIntel)
                            lstIntel.AddItem(sText, False)
                            lstIntel.ItemData(lstIntel.NewIndex) = oItem.lItemID
                            lstIntel.ItemData2(lstIntel.NewIndex) = oItem.iItemTypeID
                        End If
                    Else
                        'Player tech knowledge
                        Dim oItem As PlayerTechKnowledge = goPlayerTechKnowledge(lSorted(X))
                        If oItem Is Nothing = False AndAlso oItem.oTech Is Nothing = False Then
                            Dim sText As String = oItem.oTech.GetComponentName().PadRight(21, " "c)
                            sText &= GetCacheObjectValueCheckUnknowns(oItem.oTech.OwnerID, ObjectType.ePlayer, mbHasUnknownsIntel)
                            lstIntel.AddItem(sText, False)
                            lstIntel.ItemData(lstIntel.NewIndex) = oItem.oTech.ObjectID
                            lstIntel.ItemData2(lstIntel.NewIndex) = oItem.oTech.ObjTypeID
                        End If
                    End If
                End If
            Next X
        End If
    End Sub

    Protected Overrides Sub Finalize()
        If AgentRenderer.goAgentRenderer Is Nothing = False Then
            AgentRenderer.goAgentRenderer.ClearTextures()
        End If

        Try
            Dim oWin As UIWindow = MyBase.moUILib.GetWindow("frmAgent")
            If oWin Is Nothing = True OrElse oWin.Visible = False Then
                If moAgentIcons Is Nothing = False Then
                    moAgentIcons.Dispose()
                End If
                moAgentIcons = Nothing
            End If
        Catch ex As Exception
        End Try
        MyBase.Finalize()
    End Sub

	Private Sub btnViewAgent_Click(ByVal sName As String) Handles btnViewAgent.Click
		If lstAgents.ListIndex > -1 Then

			If myFilterValue = 0 Then
				Dim oAgent As Agent = Nothing
				Dim lID As Int32 = lstAgents.ItemData(lstAgents.ListIndex)
				For X As Int32 = 0 To goCurrentPlayer.AgentUB
					If goCurrentPlayer.AgentIdx(X) = lID Then
						oAgent = goCurrentPlayer.Agents(X)
						Exit For
					End If
				Next X
				If oAgent Is Nothing Then Return

				Dim ofrm As frmAgent = CType(goUILib.GetWindow("frmAgent"), frmAgent)
				If ofrm Is Nothing Then ofrm = New frmAgent(goUILib, False, False)
				ofrm.Visible = True
                ofrm.SetFromAgent(oAgent)
            Else
                Dim oAgent As CapturedAgent = Nothing
                Dim lID As Int32 = lstAgents.ItemData(lstAgents.ListIndex)
                For X As Int32 = 0 To goCurrentPlayer.CapturedAgentUB
                    If goCurrentPlayer.CapturedAgentIdx(X) = lID Then
                        oAgent = goCurrentPlayer.CapturedAgents(X)
                        Exit For
                    End If
                Next X
                If oAgent Is Nothing Then Return

                Dim ofrm As frmCapturedAgent = CType(goUILib.GetWindow("frmCapturedAgent"), frmCapturedAgent)
                If ofrm Is Nothing Then ofrm = New frmCapturedAgent(goUILib)
                ofrm.Visible = True
                ofrm.SetFromCapturedAgent(oAgent)
            End If

		Else
			MyBase.moUILib.AddNotification("Select an agent in the list to view.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		End If
    End Sub

    Private Sub lstAgents_ItemDblClick(ByVal lIndex As Integer) Handles lstAgents.ItemDblClick
        If btnViewAgent.Enabled = False Then Exit Sub
        If lstAgents.ListIndex < 0 OrElse lstAgents.ItemData(lstAgents.ListIndex) < 0 Then Exit Sub
        frmMain.mbIgnoreNextMouseUp = True
        btnViewAgent_Click(Nothing)
        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
    End Sub

	Private Sub btnCreateMission_Click(ByVal sName As String) Handles btnCreateMission.Click
		Dim ofrm As frmMission = New frmMission(goUILib)
		ofrm.Visible = True
        ofrm.SetFromMission(Nothing, True)
	End Sub

	Private Sub btnViewMission_Click(ByVal sName As String) Handles btnViewMission.Click
		If lstMissions.ListIndex > -1 Then
			Dim oMission As PlayerMission = Nothing
			Dim lPM_ID As Int32 = lstMissions.ItemData(lstMissions.ListIndex)
			For X As Int32 = 0 To goCurrentPlayer.PlayerMissionUB
				If goCurrentPlayer.PlayerMissionIdx(X) = lPM_ID Then
					oMission = goCurrentPlayer.PlayerMissions(X)
					Exit For
				End If
			Next X
			If oMission Is Nothing Then Return
			If oMission.yCurrentPhase = eMissionPhase.eInPlanning Then
				Dim ofrm As frmMission = New frmMission(goUILib)
				ofrm.Visible = True
                ofrm.SetFromMission(oMission, False)
			Else
				Dim ofrm As frmViewMission = New frmViewMission(goUILib)
				ofrm.Visible = True
				ofrm.SetFromMission(oMission)
			End If

		End If
	End Sub

    Private Sub lstMissions_ItemDblClick(ByVal lIndex As Integer) Handles lstMissions.ItemDblClick
        If btnViewMission.Enabled = False Then Exit Sub
        If lstMissions.ListIndex < 0 OrElse lstMissions.ItemData(lstMissions.ListIndex) < 0 Then Exit Sub
        frmMain.mbIgnoreNextMouseUp = True
        btnViewMission_Click(Nothing)
        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
    End Sub

	Private Sub RefreshIntelFilterList()
		If goCurrentPlayer Is Nothing = False Then
			Dim lSorted() As Int32 = GetSortedPlayerRelIdxArray()
			If lSorted Is Nothing = False Then
				For X As Int32 = 0 To lSorted.GetUpperBound(0)
					Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(lSorted(X))
					If oRel Is Nothing = False Then
						Dim lID As Int32 = -1
						If oRel.lPlayerRegards = glPlayerID Then
							lID = oRel.lThisPlayer
						Else : lID = oRel.lPlayerRegards
						End If
						If lID <> -1 Then
							If cboIntelFilter.ItemData(X + 1) <> lID Then
								FullRefreshIntelFilterList()
								Return
							End If
							Dim sText As String = GetCacheObjectValue(lID, ObjectType.ePlayer)
							If cboIntelFilter.List(X + 1) <> sText Then cboIntelFilter.List(X) = sText
						End If
					End If
				Next X
			End If
		End If
	End Sub
	Private Sub FullRefreshIntelFilterList()
		Dim lIdx As Int32 = -1
		If cboIntelFilter.ListIndex > -1 Then lIdx = cboIntelFilter.ItemData(cboIntelFilter.ListIndex)

		cboIntelFilter.Clear()

		If goCurrentPlayer Is Nothing = False Then
			cboIntelFilter.AddItem("All Player Intel")
			cboIntelFilter.ItemData(cboIntelFilter.NewIndex) = -1
			Dim lSorted() As Int32 = GetSortedPlayerRelIdxArray()
			If lSorted Is Nothing = False Then
				For X As Int32 = 0 To lSorted.GetUpperBound(0)
					Dim oRel As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(lSorted(X))
					If oRel Is Nothing = False Then
						Dim lID As Int32 = -1
						If oRel.lPlayerRegards = glPlayerID Then
							lID = oRel.lThisPlayer
						Else : lID = oRel.lPlayerRegards
						End If
						If lID <> -1 Then
							Dim sText As String = GetCacheObjectValue(lID, ObjectType.ePlayer)
							cboIntelFilter.AddItem(sText)
							cboIntelFilter.ItemData(cboIntelFilter.NewIndex) = lID
						End If
					End If
				Next X
			End If
			'now, see if we can find it
			For X As Int32 = 0 To cboIntelFilter.ListCount - 1
				If cboIntelFilter.ItemData(X) = lIdx Then
					If cboIntelFilter.ListIndex <> X Then cboIntelFilter.ListIndex = X
					Exit For
				End If
			Next X
		End If


	End Sub

    Private mlPreviousCheck As Int32
	Private Sub frmAgentMain_OnNewFrame() Handles Me.OnNewFrame
		If goCurrentPlayer Is Nothing Then Return

		If glCurrentCycle - mlLastAgentRequestTime > 15 Then
			mlLastAgentRequestTime = glCurrentCycle
			For X As Int32 = 0 To goCurrentPlayer.AgentUB
                If goCurrentPlayer.AgentIdx(X) <> -1 AndAlso goCurrentPlayer.Agents(X).bRequestedSkillList = False Then
                    'goCurrentPlayer.Agents(X).bRequestedSkillList = True
                    'Dim yOut(5) As Byte
                    'System.BitConverter.GetBytes(GlobalMessageCode.eGetSkillList).CopyTo(yOut, 0)
                    'System.BitConverter.GetBytes(goCurrentPlayer.Agents(X).ObjectID).CopyTo(yOut, 2)
                    'MyBase.moUILib.SendMsgToPrimary(yOut)
                    goCurrentPlayer.Agents(X).CheckRequestSkillList()
                    Exit For
                End If
			Next X
		End If

        If glCurrentCycle - mlPreviousCheck > 15 Then
            'Only if there were unknowns from the initial load, lets attempt to refresh a few times limited to 3 seconds.  Assume some unknowns will never be known
            If mbHasUnknownsIntel AndAlso mlTimeOpened + 30 > glCurrentCycle Then RefreshIntelList()

            mlPreviousCheck = glCurrentCycle

            If cboIntelFilter Is Nothing = False Then
                RefreshIntelFilterList()
            End If

            If lstMissions Is Nothing = False Then

                If myViewState = 0 Then
                    Try
                        Dim lVisibleMissionCnt As Int32 = 0
                        Dim bMissionFilter As Boolean = chkFilterArchivedMissions.Value
                        For X As Int32 = 0 To goCurrentPlayer.PlayerMissionUB
                            If goCurrentPlayer.PlayerMissionIdx(X) <> -1 AndAlso (bMissionFilter = False OrElse goCurrentPlayer.PlayerMissions(X).bArchived = False) Then
                                lVisibleMissionCnt += 1
                            End If
                        Next X

                        If lstMissions.ListCount <> lVisibleMissionCnt Then
                            RefreshMissionList()
                        Else
                            For X As Int32 = 0 To lstMissions.ListCount - 1
                                Dim lID As Int32 = lstMissions.ItemData(X)
                                For Y As Int32 = 0 To goCurrentPlayer.PlayerMissionUB
                                    If goCurrentPlayer.PlayerMissionIdx(Y) = lID Then

                                        If NewTutorialManager.TutorialOn = True Then
                                            If lstMissions.ItemData2(X) <> goCurrentPlayer.PlayerMissions(Y).yCurrentPhase Then
                                                NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eMissionPhaseChange, goCurrentPlayer.PlayerMissions(Y).oMission.ObjectID, goCurrentPlayer.PlayerMissions(Y).yCurrentPhase, -1, "")
                                                lstMissions.ItemData2(X) = goCurrentPlayer.PlayerMissions(Y).yCurrentPhase
                                            End If
                                        End If

                                        Dim sText As String = goCurrentPlayer.PlayerMissions(Y).GetListBoxText()
                                        If lstMissions.List(X) <> sText Then
                                            lstMissions.List(X) = sText
                                            lstMissions.rcIconRectangle(X) = GetMissionStatusIcon(goCurrentPlayer.PlayerMissions(Y).yCurrentPhase)
                                        End If
                                        Exit For
                                    End If
                                Next Y
                            Next X
                        End If
                    Catch
                        lstMissions.Clear()
                    End Try
                Else
                    If muEffects Is Nothing = False Then
                        If muEffects.Length <> lstMissions.ListCount Then
                            lstMissions.Clear()
                            lstMissions.sHeaderRow = "Expires".PadRight(8, " "c) & "Effect".PadRight(34, " "c) & "Target"
                            Try
                                For X As Int32 = 0 To muEffects.GetUpperBound(0)
                                    lstMissions.AddItem(muEffects(X).GetListBoxText, False)
                                    lstMissions.ItemData(lstMissions.NewIndex) = X
                                Next X
                            Catch
                                lstMissions.Clear()
                            End Try
                        Else
                            Dim lUB As Int32 = muEffects.GetUpperBound(0)
                            For X As Int32 = 0 To lstMissions.ListCount - 1
                                Dim lIdx As Int32 = lstMissions.ItemData(X)
                                If lIdx > -1 AndAlso lIdx <= lUB Then
                                    Dim sTemp As String = muEffects(lIdx).GetListBoxText
                                    If lstMissions.List(X) <> sTemp Then lstMissions.List(X) = sTemp
                                End If
                            Next X
                        End If
                    Else
                        lstMissions.Clear()
                    End If
                End If

            End If
            If lstAgents Is Nothing = False Then
                Dim lActualAgentCnt As Int32 = 0

                If myFilterValue = 0 Then
                    For X As Int32 = 0 To goCurrentPlayer.AgentUB
                        If goCurrentPlayer.AgentIdx(X) <> -1 Then lActualAgentCnt += 1
                    Next X
                Else
                    For X As Int32 = 0 To goCurrentPlayer.CapturedAgentUB
                        If goCurrentPlayer.CapturedAgentIdx(X) <> -1 Then lActualAgentCnt += 1
                    Next X
                End If

                If lstAgents.ListCount <> lActualAgentCnt Then
                    RefreshAgentList()
                Else
                    For X As Int32 = 0 To lstAgents.ListCount - 1
                        Dim lID As Int32 = lstAgents.ItemData(X)

                        If myFilterValue = 0 Then
                            For Y As Int32 = 0 To goCurrentPlayer.AgentUB
                                If goCurrentPlayer.AgentIdx(Y) = lID Then
                                    With goCurrentPlayer.Agents(Y)
                                        Dim sText As String = .sAgentName.PadRight(21, " "c)
                                        sText &= Agent.GetStatusText(.lAgentStatus, .lTargetID, .iTargetTypeID, .InfiltrationType)
                                        If lstAgents.List(X) <> sText Then lstAgents.List(X) = sText
                                        If lstAgents.ItemData2(X) <> .lAgentStatus Then
                                            lstAgents.rcIconRectangle(X) = .GetStatusIconRect
                                            lstAgents.ItemData2(X) = .lAgentStatus
                                        End If
                                        Exit For
                                    End With
                                End If
                            Next Y
                        Else
                            For Y As Int32 = 0 To goCurrentPlayer.CapturedAgentUB
                                If goCurrentPlayer.CapturedAgentIdx(Y) = lID Then
                                    Dim sText As String = goCurrentPlayer.CapturedAgents(Y).sAgentName.PadRight(21, " "c)
                                    sText &= GetCacheObjectValue(goCurrentPlayer.CapturedAgents(Y).lOwnerID, ObjectType.ePlayer)
                                    If lstAgents.List(X) <> sText Then lstAgents.List(X) = sText
                                    Exit For
                                End If
                            Next Y
                        End If


                    Next X
                End If
            End If
        End If
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		CloseForm()
	End Sub

	Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
		goUILib.AddNotification("Agent Tutorial is not yet implemented", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
	End Sub

	Private Sub frmAgentMain_OnRender() Handles Me.OnRender
		If moAgentIcons Is Nothing OrElse moAgentIcons.Disposed = True Then
			moAgentIcons = goResMgr.LoadScratchTexture("AgentIcons.dds", "Interface.pak")
			If lstAgents Is Nothing = False Then Me.lstAgents.oIconTexture = moAgentIcons
			If lstMissions Is Nothing = False Then Me.lstMissions.oIconTexture = moAgentIcons
		End If
	End Sub

	Public Sub CloseForm()
		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.RemoveWindow("frmAgent")
		MyBase.moUILib.RemoveWindow("frmViewMission")
		MyBase.moUILib.RemoveWindow("frmMission")
		MyBase.moUILib.RemoveWindow("frmCapturedAgent")
		MyBase.moUILib.RemoveWindow("frmAgentPopup")
	End Sub

    'Private Sub btnShow_Click(ByVal sName As String) Handles btnShow.Click
    '       If myFilterValue = 0 Then
    '           myFilterValue = 1
    '           btnDismissAgent.Enabled = False
    '           btnShow.Caption = "My Agents"
    '       Else
    '           myFilterValue = 0
    '           btnDismissAgent.Enabled = True
    '           btnShow.Caption = "Enemy Agents"
    '       End If
    'End Sub

    Private Sub optMyAgents_Click() Handles optMyAgents.Click
        optMyAgents.Value = True
        If myFilterValue = 0 Then Return
        optEnemyAgents.Value = False

        myFilterValue = 0
        btnDismissAgent.Enabled = True
        RefreshAgentList()
    End Sub

    Private Sub optEnemyAgents_Click() Handles optEnemyAgents.Click
        optEnemyAgents.Value = True
        If myFilterValue = 1 Then Return
        optMyAgents.Value = False
        myFilterValue = 1
        btnDismissAgent.Enabled = False
        RefreshAgentList()
    End Sub

	Private Sub cboIntelFilter_ItemSelected(ByVal lItemIndex As Integer) Handles cboIntelFilter.ItemSelected
		RefreshIntelList()
	End Sub

	Private Sub btnViewIntel_Click(ByVal sName As String) Handles btnViewIntel.Click

		If lstIntel.ListIndex < 0 Then Return

		Dim lID As Int32 = lstIntel.ItemData(lstIntel.ListIndex)
		Dim iTypeID As Int16 = CShort(lstIntel.ItemData2(lstIntel.ListIndex))

		Dim oTech As PlayerTechKnowledge = Nothing
		Dim oItem As PlayerItemIntel = Nothing

		Select Case iTypeID
			Case ObjectType.eAlloyTech, ObjectType.eArmorTech, ObjectType.eEngineTech, ObjectType.ePrototype, ObjectType.eRadarTech, ObjectType.eShieldTech, ObjectType.eWeaponTech, ObjectType.eSpecialTech, ObjectType.eHullTech
				For X As Int32 = 0 To glPlayerTechKnowledgeUB
					If glPlayerTechKnowledgeIdx(X) = lID Then
						If goPlayerTechKnowledge(X).oTech Is Nothing = False Then
							If goPlayerTechKnowledge(X).oTech.ObjTypeID = iTypeID Then
								oTech = goPlayerTechKnowledge(X)
								Exit For
							End If
						End If
					End If
				Next X
			Case Else
				For X As Int32 = 0 To glItemIntelUB
					If glItemIntelIdx(X) <> -1 Then
						If goItemIntel(X).lItemID = lID AndAlso goItemIntel(X).iItemTypeID = iTypeID Then
							oItem = goItemIntel(X)
							Exit For
						End If
					End If
				Next X
		End Select

		If oTech Is Nothing = False OrElse oItem Is Nothing = False Then
			Dim ofrm As frmIntelReport = CType(goUILib.GetWindow("frmIntelReport"), frmIntelReport)
			If ofrm Is Nothing Then ofrm = New frmIntelReport(goUILib)
			ofrm.Visible = True
            If oTech Is Nothing = False Then ofrm.SetFromTechKnowledge(oTech) Else ofrm.SetFromItemIntel(oItem)
            moUILib.BringWindowToForeground(ofrm.ControlName)
		Else
			MyBase.moUILib.AddNotification("That intelligence report has been misplaced.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		End If
    End Sub

    Private Sub lstIntel_ItemDblClick(ByVal lIndex As Integer) Handles lstIntel.ItemDblClick
        If btnViewIntel.Enabled = False Then Exit Sub
        If lstIntel.ListIndex < 0 OrElse lstIntel.ItemData(lstIntel.ListIndex) < 0 Then Exit Sub
        frmMain.mbIgnoreNextMouseUp = True
        btnViewIntel_Click(Nothing)
        If goSound Is Nothing = False Then goSound.StartSound("UserInterface\ButtonClick.wav", False, SoundMgr.SoundUsage.eUserInterface, Vector3.Empty, Vector3.Empty)
    End Sub

	Private Sub chkFilterArchivedIntel_Click() Handles chkFilterArchivedIntel.Click
		MyBase.moUILib.GetMsgSys.LoadArchived()
		RefreshIntelList()
	End Sub

	Private Sub chkSetFilterMission_Click() Handles chkSetFilterMission.Click
		If mbIgnoreSetFilter = True Then Return
		If lstMissions Is Nothing = False AndAlso lstMissions.ListIndex > -1 Then
			Dim lPM_ID As Int32 = lstMissions.ItemData(lstMissions.ListIndex)

			If lPM_ID < 1 Then Return

			For X As Int32 = 0 To goCurrentPlayer.PlayerMissionUB
				If goCurrentPlayer.PlayerMissionIdx(X) = lPM_ID Then
					goCurrentPlayer.PlayerMissions(X).bArchived = chkSetFilterMission.Value
					Exit For
				End If
			Next X

			Dim yMsg(8) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eSetArchiveState).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(lPM_ID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(ObjectType.eMission).CopyTo(yMsg, lPos) : lPos += 2
			If chkSetFilterMission.Value = True Then yMsg(lPos) = 1 Else yMsg(lPos) = 0
			lPos += 1

			MyBase.moUILib.SendMsgToPrimary(yMsg)
			lstMissions.ListIndex = -1
		End If
	End Sub

	Private Sub lstMissions_ItemClick(ByVal lIndex As Integer) Handles lstMissions.ItemClick
		Dim lMissionID As Int32 = -1
		If lstMissions.ListIndex > -1 Then
			lMissionID = lstMissions.ItemData(lstMissions.ListIndex)
			If lMissionID > 0 Then
				For X As Int32 = 0 To goCurrentPlayer.PlayerMissionUB
					If goCurrentPlayer.PlayerMissionIdx(X) = lMissionID Then
						mbIgnoreSetFilter = True
						chkSetFilterMission.Value = goCurrentPlayer.PlayerMissions(X).bArchived
						mbIgnoreSetFilter = False
						Exit For
					End If
				Next X
			End If
		End If
	End Sub

	Private Sub chkFilterArchivedMissions_Click() Handles chkFilterArchivedMissions.Click
		MyBase.moUILib.GetMsgSys.LoadArchived()
		RefreshMissionList()
	End Sub

    Private Sub btnDismissAgent_Click(ByVal sName As String) Handles btnDismissAgent.Click
        If lstAgents.ListIndex > -1 Then
            If btnDismissAgent.Caption.ToUpper = "CONFIRM" Then
                Dim lID As Int32 = lstAgents.ItemData(lstAgents.ListIndex)

                Dim yMsg(9) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eSetAgentStatus).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(lID).CopyTo(yMsg, 2)
                System.BitConverter.GetBytes(AgentStatus.Dismissed).CopyTo(yMsg, 6)
                MyBase.moUILib.SendMsgToPrimary(yMsg)
                lstAgents.ListIndex = -1
                btnDismissAgent.Enabled = False
            Else : btnDismissAgent.Caption = "Confirm"
            End If
        Else
            MyBase.moUILib.AddNotification("Select an agent in the list first.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        End If
    End Sub

    Private Sub lstAgents_HeaderRowClick(ByVal lX As Integer) Handles lstAgents.HeaderRowClick
        If myViewState <> 0 Then Return
        If lX < 153 Then
            If Math.Abs(myAgentSort) = 1 Then myAgentSort = -myAgentSort Else myAgentSort = 1
        Else
            If Math.Abs(myAgentSort) = 2 Then myAgentSort = -myAgentSort Else myAgentSort = 2
        End If
        RefreshAgentList()
    End Sub

    Private Sub lstAgents_ItemClick(ByVal lIndex As Integer) Handles lstAgents.ItemClick
        If myFilterValue <> 0 Then Return
        btnDismissAgent.Caption = "Dismiss"
        btnDismissAgent.Enabled = True
    End Sub

    Private Sub optCurrentEffects_Click() Handles optCurrentEffects.Click
        optCurrentMissions.Value = False
        optCurrentEffects.Value = True
        lstMissions.Clear()
        myViewState = 1

        chkFilterArchivedMissions.Enabled = False
        chkSetFilterMission.Enabled = False
        btnCreateMission.Enabled = False
        btnViewMission.Enabled = False

        Dim yMsg(1) As Byte
        System.BitConverter.GetBytes(GlobalMessageCode.eGetImposedAgentEffects).CopyTo(yMsg, 0)
        MyBase.moUILib.SendMsgToPrimary(yMsg)
    End Sub

    Private Sub optCurrentMissions_Click() Handles optCurrentMissions.Click
        optCurrentEffects.Value = False
        optCurrentMissions.Value = True
        lstMissions.Clear()
        myViewState = 0

        chkFilterArchivedMissions.Enabled = True
        chkSetFilterMission.Enabled = True
        btnCreateMission.Enabled = True
        btnViewMission.Enabled = True
    End Sub

    Public Sub HandleGetImposedAgentEffects(ByRef yData() As Byte)
        Dim lPos As Int32 = 2 'for msgcode
        Dim lCnt As Int32 = System.BitConverter.ToInt32(yData, lPos) : lPos += 4

        Dim uEffects(lCnt - 1) As ImposedAgentEffects
        For X As Int32 = 0 To lCnt - 1
            With uEffects(X)
                .lTargetID = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lSetOnCycle = glCurrentCycle
                .lCyclesRemaining = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .lAmount = System.BitConverter.ToInt32(yData, lPos) : lPos += 4
                .yType = CType(yData(lPos), AgentEffectType) : lPos += 1
                .bAmountAsPerc = yData(lPos) <> 0 : lPos += 1
                .sSpecificName = GetStringFromBytes(yData, lPos, 20) : lPos += 20
            End With
        Next X

        lstMissions.Clear()
        muEffects = uEffects
    End Sub

    Private Sub frmAgentMain_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.AgentMainX = Me.Left
            muSettings.AgentMainY = Me.Top
        End If
    End Sub

    Private Sub btnExport_Click(ByVal sName As String) Handles btnExport.Click
        MyBase.moUILib.AddNotification("Beginning agent export...", muSettings.InterfaceBorderColor, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
        Agent.Export_AgentInfo()
    End Sub
End Class