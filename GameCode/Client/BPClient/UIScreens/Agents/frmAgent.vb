Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmAgent
    Inherits UIWindow

    Private WithEvents fraPhoto As UIWindow
    Private lblName As UILabel
    Private lblLocation As UILabel
    Private lblStatus As UILabel
	Private lblExperience As UILabel
	Private lblETA As UILabel 
    Private mfraProficiencies As fraProficiencies
    Private mfraSkills As fraSkills
    Private mfraInfiltration As fraInfiltration
    Private mfraHistory As fraHistory
    Private WithEvents btnClose As UIButton
	Private WithEvents btnHelp As UIButton
	Private WithEvents btnRecruit As UIButton
	Private WithEvents btnDismiss As UIButton

	Private moAgent As Agent
	Private mlPrevStatus As Int32

	Private mbAsPopup As Boolean = False
	Private mbAsChild As Boolean = False

    Private mlLastRequest As Int32 = 0

    Private mbHasUnknowns As Boolean = False

    Public Event AgentSelected(ByRef oAgent As Agent)

    Private mlFirstRender As Int32 = -1
    Private mbLoading As Boolean = True

	Public Sub New(ByRef oUILib As UILib, ByVal bAsPopup As Boolean, ByVal bAsChild As Boolean)
		MyBase.New(oUILib)

		mbAsPopup = bAsPopup
        mbAsChild = bAsChild

        'TutorialAlert.CheckTrigger(TutorialAlert.eTriggerType.OpenAgentWindow)

		'frmAgent initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eAgentView

            If bAsPopup = True Then
                .ControlName = "frmAgentPopup"
                .Width = 221 '186
                .Height = 382
            Else
                .ControlName = "frmAgent"
                .Width = 455 '420
                .Height = 512
            End If

            Dim lLeft As Int32 = -1
            Dim lTop As Int32 = -1
            muSettings.AgentMissionDetailsX = -100
            muSettings.AgentMissionDetailsY = -100

            If NewTutorialManager.TutorialOn = False Then
                lLeft = muSettings.AgentDetailsX
                lTop = muSettings.AgentDetailsY
            End If
            If lLeft < 0 Then lLeft = 241
            If lTop < 0 Then lTop = 106
            If lLeft + .Width > MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth Then lLeft = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth - .Width
            If lTop + .Height > MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight Then lTop = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight - .Height
            .Left = lLeft
            .Top = lTop

            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True

            If bAsPopup = True Then
                .FillColor = System.Drawing.Color.FromArgb(255, CInt(.FillColor.R * 0.375F), CInt(.FillColor.G * 0.375F), CInt(.FillColor.B * 0.375F))
                .BorderLineWidth = 2
            End If
            .mbAcceptReprocessEvents = True
        End With

		'btnClose initial props
		btnClose = New UIButton(oUILib)
		With btnClose
			.ControlName = "btnClose"
            .Left = Me.Width - 24 - Me.BorderLineWidth
			.Top = Me.BorderLineWidth
			.Width = 24
			.Height = 24
			.Enabled = True
			.Visible = Not bAsChild
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
			.Visible = Not bAsPopup
			.Caption = "?"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnHelp, UIControl))

        'btnDismiss initial props
		btnDismiss = New UIButton(oUILib)
		With btnDismiss
			.ControlName = "btnDismiss"
            '.Left = btnRecruit.Left + btnRecruit.Width + 10
            '.Top = btnRecruit.Top
            .Left = 5
            .Top = 482
			.Width = 100
			.Height = 24
			.Enabled = True
            .Visible = Not bAsPopup
			.Caption = "Dismiss"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnDismiss, UIControl))

        'btnRecruit initial props
        btnRecruit = New UIButton(oUILib)
        With btnRecruit
            .ControlName = "btnRecruit"
            '.Left = 5
            '.Top = Me.Height - 30
            .Left = btnDismiss.Left + btnDismiss.Width + 10
            .Top = btnDismiss.Top
            .Width = 100
            .Height = 24
            .Enabled = True
            .Visible = Not bAsPopup
            .Caption = "Recruit"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = True
            .FontFormat = CType(5, DrawTextFormat)
            .ControlImageRect = New Rectangle(0, 0, 120, 32)
        End With
        Me.AddChild(CType(btnRecruit, UIControl))

		'fraPhoto initial props
		fraPhoto = New UIWindow(oUILib)
		With fraPhoto
			.ControlName = "fraPhoto"
			If bAsPopup = True Then
				.Left = (Me.Width - 128) \ 2
				.Top = 249	'379
			Else
				.Left = 5
				.Top = 5
			End If
			.Width = 128
			.Height = 128
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = System.Drawing.Color.FromArgb(0, 0, 0, 0)
			.FillWindow = False
			.FullScreen = False
			.BorderLineWidth = 1
			.bRoundedBorder = False
			.Moveable = False
		End With
		Me.AddChild(CType(fraPhoto, UIControl))

		'lblName initial props
		lblName = New UILabel(oUILib)
		With lblName
			.ControlName = "lblName"
			If bAsPopup = True Then
				.Left = 5
				.Top = 10
			Else
				.Left = 145
				.Top = 10
			End If

			.Width = 250
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Name: "
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
            .FontFormat = CType(4, DrawTextFormat)
            .bFilterBadWords = False
		End With
		Me.AddChild(CType(lblName, UIControl))

		'lblLocation initial props
		lblLocation = New UILabel(oUILib)
		With lblLocation
			.ControlName = "lblLocation"
			.Left = 145
			.Top = 30
			.Width = 250
			.Height = 18
			.Enabled = True
			.Visible = Not bAsPopup
			.Caption = "Location: Classified"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblLocation, UIControl))

		'lblStatus initial props
		lblStatus = New UILabel(oUILib)
		With lblStatus
			.ControlName = "lblStatus"
			.Left = 145
			.Top = 50
			.Width = 250
			.Height = 18
			.Enabled = True
			.Visible = Not bAsPopup
			.Caption = "Status: "
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblStatus, UIControl))

		'lblExperience initial props
		lblExperience = New UILabel(oUILib)
		With lblExperience
			.ControlName = "lblExperience"
			.Left = 145
			.Top = 70
			.Width = 250
			.Height = 18
			.Enabled = True
			.Visible = Not bAsPopup
			.Caption = "Time With Agency: "
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblExperience, UIControl))

		'lblEta initial props
		lblETA = New UILabel(oUILib)
		With lblETA
			.ControlName = "lblETA"
			.Left = 145
			.Top = 90
			.Width = 250
			.Height = 18
			.Enabled = True
			.Visible = Not bAsPopup
			.Caption = "Current ETA: "
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
			.ToolTipText = "Indicates any travel time associated with the agent's movements" & vbCrLf & "as well as any delays between actions caused by various activities."
		End With
		Me.AddChild(CType(lblETA, UIControl))

		'mfraProficiencies initial props
		mfraProficiencies = New fraProficiencies(oUILib, bAsChild)
		With mfraProficiencies
			.ControlName = "mfraProficiencies"
			If bAsPopup = True Then
				.Left = 5
				.Top = 40
			Else
				.Left = 5
				.Top = 145
			End If
			.Width = 210 '175
			.Height = 95
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			.BorderLineWidth = 2
			If bAsPopup = True Then
				.FillColor = Me.FillColor
			End If
		End With
		Me.AddChild(CType(mfraProficiencies, UIControl))

		'mfraSkills initial props
		mfraSkills = New fraSkills(oUILib)
		With mfraSkills
			.ControlName = "mfraSkills"
			If bAsPopup = True Then
				.Left = 5
				.Top = 146
			Else
				.Left = 5
				.Top = 250
			End If
			.Width = 210 ' 175
			.Height = 225
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			If bAsPopup = True Then
				.FillColor = Me.FillColor
				.Height = Me.mfraProficiencies.Height
			End If
		End With
		Me.AddChild(CType(mfraSkills, UIControl))

		'mfraHistory initial props
		mfraHistory = New fraHistory(oUILib)
		With mfraHistory
			.ControlName = "mfraHistory"
			.Left = 225	'190
			.Top = 375
			.Width = 225
			.Height = 100
			.Enabled = True
			.Visible = Not bAsPopup
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			If bAsPopup = True Then
				.FillColor = Me.FillColor
			End If
		End With
		Me.AddChild(CType(mfraHistory, UIControl))

		'mfraInfiltration initial props
		mfraInfiltration = New fraInfiltration(oUILib)
		With mfraInfiltration
			.ControlName = "mfraInfiltration"
			.Left = 225	'190
			.Top = 145
			.Width = 225
			.Height = 220
			.Enabled = True
			.Visible = Not bAsPopup
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			If bAsPopup = True Then
				.FillColor = Me.FillColor
			End If
		End With
		Me.AddChild(CType(mfraInfiltration, UIControl))

		If bAsChild = False Then
			MyBase.moUILib.RemoveWindow(Me.ControlName)
			MyBase.moUILib.AddWindow(Me)
        End If
        mbLoading = False
	End Sub

    Public Sub SetFromAgent(ByRef oAgent As Agent)
        moAgent = oAgent

        lblName.Caption = "Name: " & moAgent.sAgentName
        lblStatus.Caption = "Status: " & Agent.GetStatusText(moAgent.lAgentStatus, moAgent.lTargetID, moAgent.iTargetTypeID, moAgent.InfiltrationType)
        If btnDismiss.Caption <> "Dismiss" Then btnDismiss.Caption = "Dismiss"
        If (moAgent.lAgentStatus And AgentStatus.NewRecruit) <> 0 Then
			lblExperience.Caption = "Agent Not Recruited"
            'btnRecruit.Caption = "Recruit"
			btnRecruit.ToolTipText = "Click to hire this agent into the agency." & vbCrLf & _
			  "Doing so will cost " & moAgent.lUpfrontCost.ToString("#,##0") & " credits upfront" & vbCrLf & _
			  "and will require " & moAgent.lMaintCost.ToString("#,##0") & " credits per " & vbCrLf & _
			  "payment cycle to maintain the agent's payroll."
            'btnRecruit.Enabled = True
            'btnDismiss.Enabled = True
            btnRecruit.Visible = True
            btnRecruit.Caption = "Recruit"
			mfraInfiltration.Enabled = False
		Else
            'btnDismiss.Visible = True
            btnRecruit.Visible = False
			Dim lTemp As Int32 = CInt(Now.Subtract(moAgent.dtRecruited).TotalSeconds)
			lblExperience.Caption = "Time With Agency: " & GetDurationFromSeconds(lTemp, True)
            'btnRecruit.Caption = "Dismiss"
			btnRecruit.ToolTipText = "This agent requires " & moAgent.lMaintCost.ToString("#,##0") & " credits per " & vbCrLf & _
			 "payment cycle to maintain. Click this button to dismiss this agent from the agency."
			If moAgent.lAgentStatus <> AgentStatus.NormalStatus AndAlso (moAgent.lAgentStatus Or AgentStatus.IsDead) = 0 Then
				btnRecruit.Enabled = False
				btnRecruit.ToolTipText &= vbCrLf & "Agents can only be dismissed when they are at home, unassigned and uninfiltrated."
			End If
			mfraInfiltration.Enabled = (moAgent.lAgentStatus And (AgentStatus.Dismissed Or AgentStatus.HasBeenCaptured Or AgentStatus.IsDead)) = 0
		End If

		mlPrevStatus = moAgent.lAgentStatus

        mfraProficiencies.SetFromAgent(oAgent)
        mfraSkills.SetFromAgent(oAgent)
        mfraInfiltration.SetFromAgent(oAgent)
        mfraHistory.SetFromAgent(oAgent)
        MyBase.moUILib.BringWindowToForeground(Me.ControlName)
    End Sub

    Public Sub SetAgentIfNeeded(ByRef oAgent As Agent)
        If moAgent Is Nothing = True OrElse moAgent.ObjectID <> oAgent.ObjectID Then SetFromAgent(oAgent)
    End Sub

    Private Sub frmAgent_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
        mfraSkills.ClearSkillPopup()
    End Sub

	Private Sub frmAgent_OnNewFrame() Handles Me.OnNewFrame
        If glCurrentCycle - mlLastRequest > 150 Then
            mlLastRequest = glCurrentCycle
            If moAgent Is Nothing = False Then
                Dim yMsg(5) As Byte
                System.BitConverter.GetBytes(GlobalMessageCode.eGetAgentStatus).CopyTo(yMsg, 0)
                System.BitConverter.GetBytes(moAgent.ObjectID).CopyTo(yMsg, 2)
                MyBase.moUILib.SendMsgToPrimary(yMsg)
            End If
        End If

        Dim sText As String = "Status: " & Agent.GetStatusText(moAgent.lAgentStatus, moAgent.lTargetID, moAgent.iTargetTypeID, moAgent.InfiltrationType)
		If lblStatus.Caption <> sText Then lblStatus.Caption = sText

		Dim lSecs As Int32 = moAgent.GetSecondsTillArrival()
		If lSecs <> 0 Then
			sText = "Current ETA: " & GetDurationFromSeconds(lSecs, False)
		Else : sText = ""
		End If
		If lblETA.Caption <> sText Then lblETA.Caption = sText

		'"Location: " & moagent.      'TODO: Fix this

		If (moAgent.lAgentStatus And AgentStatus.NewRecruit) <> 0 Then
			sText = "Agent Not Recruited"
			If lblExperience.Caption <> sText Then lblExperience.Caption = sText
		Else
			Dim lTemp As Int32 = CInt(Now.Subtract(moAgent.dtRecruited).TotalSeconds)
			sText = "Time With Agency: " & GetDurationFromSeconds(lTemp, True)
			If lblExperience.Caption <> sText Then lblExperience.Caption = sText
		End If

		If moAgent.lAgentStatus <> mlPrevStatus Then
			SetFromAgent(moAgent)
		End If

		mfraInfiltration.fraInfiltration_OnNewFrame()
		mfraHistory.fraHistory_OnNewFrame()
        mfraSkills.fraSkills_OnNewFrame()

        If mbHasUnknowns = False Then
            If mlFirstRender <> Int32.MaxValue Then
                If mlFirstRender = -1 Then
                    mlFirstRender = glCurrentCycle
                ElseIf glCurrentCycle - mlFirstRender > 30 Then
                    mlFirstRender = Int32.MaxValue
                    Me.IsDirty = True
                End If
            End If
        Else
            Me.IsDirty = True
        End If

    End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		If mfraSkills Is Nothing = False Then mfraSkills.ClearSkillPopup()
		If mbAsChild = False Then
			MyBase.moUILib.RemoveWindow(Me.ControlName)
		Else : Me.Visible = False
		End If
	End Sub

	Private Sub btnHelp_Click(ByVal sName As String) Handles btnHelp.Click
		'
    End Sub
 
    Private Sub frmAgent_OnRender() Handles Me.OnRender
        If AgentRenderer.goAgentRenderer Is Nothing Then AgentRenderer.goAgentRenderer = New AgentRenderer()
		Dim rcDest As Rectangle = New Rectangle(fraPhoto.Left, fraPhoto.Top, fraPhoto.Width, fraPhoto.Height)
		If mbAsChild = True Then
			rcDest.X += Me.Left + Me.ParentControl.Left
			rcDest.Y += Me.Top + Me.ParentControl.Top
		End If
        AgentRenderer.goAgentRenderer.RenderAgent2(moAgent.ObjectID, rcDest, False, moAgent.bMale)
        mbHasUnknowns = AgentRenderer.goAgentRenderer.bHasUnknowns
    End Sub

    Private Sub fraPhoto_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles fraPhoto.OnMouseDown
        If mbAsPopup = True Then RaiseEvent AgentSelected(moAgent)
    End Sub

	Private Sub btnRecruit_Click(ByVal sName As String) Handles btnRecruit.Click
		If btnRecruit.Caption.ToUpper = "CONFIRM" Then
			'do it
			Dim yMsg(9) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eSetAgentStatus).CopyTo(yMsg, 0)
			System.BitConverter.GetBytes(moAgent.ObjectID).CopyTo(yMsg, 2)

			If (moAgent.lAgentStatus And AgentStatus.NewRecruit) <> 0 Then
				System.BitConverter.GetBytes(AgentStatus.NewRecruit).CopyTo(yMsg, 6)
			Else
				System.BitConverter.GetBytes(AgentStatus.Dismissed).CopyTo(yMsg, 6)
			End If
			MyBase.moUILib.SendMsgToPrimary(yMsg)

			If NewTutorialManager.TutorialOn = True Then
				NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eRecruitButtonClicked, -1, -1, -1, "")
			End If
		Else
			btnRecruit.Caption = "Confirm"
		End If
	End Sub

	Protected Overrides Sub Finalize()
		If mfraSkills Is Nothing = False Then mfraSkills.ClearSkillPopup()
		MyBase.Finalize()
	End Sub

	Private Sub btnDismiss_Click(ByVal sName As String) Handles btnDismiss.Click
		If btnDismiss.Caption.ToUpper = "CONFIRM" Then
			Dim yMsg(9) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eSetAgentStatus).CopyTo(yMsg, 0)
			System.BitConverter.GetBytes(moAgent.ObjectID).CopyTo(yMsg, 2)
			System.BitConverter.GetBytes(AgentStatus.Dismissed).CopyTo(yMsg, 6)
			MyBase.moUILib.SendMsgToPrimary(yMsg)
			btnClose_Click(btnClose.ControlName)
		Else : btnDismiss.Caption = "Confirm"
		End If
	End Sub

    Private Sub frmAgent_WindowMoved() Handles Me.WindowMoved
        If mbLoading = False Then
            muSettings.AgentDetailsX = Me.Left
            muSettings.AgentDetailsY = Me.Top
        End If
    End Sub
End Class