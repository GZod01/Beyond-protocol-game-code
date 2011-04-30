Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class ctlGoalSS
	Inherits UIWindow

	Private hscrScroll As UIScrollBar
	Private WithEvents muAA1 As ctlGoalAA
	Private WithEvents muAA2 As ctlGoalAA
	Private WithEvents muAA3 As ctlGoalAA

	Private moPMG As PlayerMissionGoal = Nothing

	Public Event AAClicked(ByRef oAA As AgentAssignment, ByRef oPMG As PlayerMissionGoal)

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'ctlGoalSS initial props
		With Me
			.ControlName = "ctlGoalSS"
			.Left = 105
			.Top = 173
			.Width = 520
			.Height = 100
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.BorderLineWidth = 2
			.Moveable = False
		End With

		'muAA1 initial props
		muAA1 = New ctlGoalAA(oUILib)
		With muAA1
			.ControlName = "muAA1"
			.Left = 5
			.Top = 10
			.Width = 170
			.Height = 60
			.Enabled = True
			.Visible = False
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			.BorderLineWidth = 2
		End With
		Me.AddChild(CType(muAA1, UIControl))

		'hscrScroll initial props
		hscrScroll = New UIScrollBar(oUILib, False)
		With hscrScroll
			.ControlName = "hscrScroll"
			.Left = 5
			.Top = 73
			.Width = 510
			.Height = 24
			.Enabled = False
			.Visible = True
			.Value = 0
			.MaxValue = 0
			.MinValue = 0
			.SmallChange = 1
			.LargeChange = 1
			.ReverseDirection = False
		End With
		Me.AddChild(CType(hscrScroll, UIControl))

		'muAA2 initial props
		muAA2 = New ctlGoalAA(oUILib)
		With muAA2
			.ControlName = "muAA2"
			.Left = 175
			.Top = 10
			.Width = 170
			.Height = 60
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			.BorderLineWidth = 2
		End With
		Me.AddChild(CType(muAA2, UIControl))

		'muAA3 initial props
		muAA3 = New ctlGoalAA(oUILib)
		With muAA3
			.ControlName = "muAA3"
			.Left = 345
			.Top = 10
			.Width = 170
			.Height = 60
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
			.Moveable = False
			.BorderLineWidth = 2
		End With
		Me.AddChild(CType(muAA3, UIControl))
	End Sub

	Public Sub SetFromPlayerMissionGoal(ByRef oPMG As PlayerMissionGoal)
		moPMG = oPMG
		If oPMG Is Nothing Then Return
		If oPMG.oGoal Is Nothing Then Return

		If oPMG.oGoal.MissionPhase = oPMG.oMission.yCurrentPhase Then
			Me.BorderColor = System.Drawing.Color.FromArgb(255, 64, 255, 64)
		Else : Me.BorderColor = muSettings.InterfaceBorderColor
		End If

		If oPMG.oSkillSet Is Nothing = False Then Me.Caption = oPMG.oGoal.sGoalName & " - " & oPMG.oSkillSet.sSkillSetName
		Dim lCnt As Int32 = 0

		For X As Int32 = 0 To oPMG.lAssignmentUB
			If oPMG.oAssignments(X) Is Nothing = False Then
				lCnt += 1
				Select Case lCnt
					Case 1
						muAA1.SetFromAgentAssignment(oPMG.oAssignments(X))
					Case 2
						muAA2.SetFromAgentAssignment(oPMG.oAssignments(X))
					Case 3
						muAA3.SetFromAgentAssignment(oPMG.oAssignments(X))
					Case Else
						'do nothing
				End Select
			End If
		Next X

		If lCnt <= 3 Then
			hscrScroll.MaxValue = 1
			hscrScroll.Value = 0
			hscrScroll.Enabled = False
		Else
			hscrScroll.MaxValue = lCnt - 3
			'hscrScroll.Value = 0
			If hscrScroll.Value > hscrScroll.MaxValue Then hscrScroll.Value = hscrScroll.MaxValue
			hscrScroll.Enabled = True
		End If

	End Sub

	Public Sub ctlGoalSS_OnNewFrame() Handles Me.OnNewFrame
		Try
			If muAA1 Is Nothing = False Then muAA1.ctlGoalAA_OnNewFrame()
			If muAA2 Is Nothing = False Then muAA2.ctlGoalAA_OnNewFrame()
			If muAA3 Is Nothing = False Then muAA3.ctlGoalAA_OnNewFrame()

			If moPMG Is Nothing = False AndAlso moPMG.oGoal Is Nothing = False AndAlso moPMG.oGoal Is Nothing = False Then
				If moPMG.oGoal.MissionPhase = moPMG.oMission.yCurrentPhase Then
					Dim clrVal As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 64, 255, 64)
					If Me.BorderColor <> clrVal Then Me.BorderColor = clrVal
				ElseIf Me.BorderColor <> muSettings.InterfaceBorderColor Then
					Me.BorderColor = muSettings.InterfaceBorderColor
				End If
			End If
		Catch
		End Try
	End Sub

	Private Sub ctlGoalSS_OnRender() Handles Me.OnRender
		If moPMG Is Nothing = False AndAlso hscrScroll Is Nothing = False Then
			Dim lCnt As Int32 = hscrScroll.Value + 1
			Dim lIdx As Int32 = 0
			For X As Int32 = 0 To moPMG.lAssignmentUB
				If moPMG.oAssignments(X) Is Nothing = False Then
					lCnt -= 1
					If lCnt = 0 Then
						lIdx = X
						Exit For
					End If
				End If
			Next X

			muAA1.Visible = False
			muAA2.Visible = False
			muAA3.Visible = False

			lCnt = 0
			For X As Int32 = lIdx To moPMG.lAssignmentUB
				If moPMG.oAssignments(X) Is Nothing = False Then
					lCnt += 1
					'If lCnt >= hscrScroll.Value Then
					Select Case lCnt '- hscrScroll.Value
						Case 1
							muAA1.SetFromAgentAssignment(moPMG.oAssignments(X))
							muAA1.Visible = True
						Case 2
							muAA2.SetFromAgentAssignment(moPMG.oAssignments(X))
							muAA2.Visible = True
						Case 3
							muAA3.SetFromAgentAssignment(moPMG.oAssignments(X))
							muAA3.Visible = True
							Exit For
					End Select
					'End If
				End If
			Next X
		End If
	End Sub

	Private Sub muAA1_AAClick(ByRef oAA As AgentAssignment) Handles muAA1.AAClick, muAA2.AAClick, muAA3.AAClick
		If oAA Is Nothing = False AndAlso moPMG Is Nothing = False Then RaiseEvent AAClicked(oAA, moPMG)
	End Sub

#Region "  ctlGoalAA  "
	'Interface created from Interface Builder
	Public Class ctlGoalAA
		Inherits UIWindow

		Private moAA As AgentAssignment = Nothing

		Private lblAgent As UILabel
		Private lblSkill As UILabel
		Private lblScore As UILabel

		Public Event AAClick(ByRef oAA As AgentAssignment)

		Public Sub New(ByRef oUILib As UILib)
			MyBase.New(oUILib)

			'ctlGoalAA initial props
			With Me
				.ControlName = "ctlGoalAA"
				.Left = 197
				.Top = 240
				.Width = 170
				.Height = 60
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.BorderLineWidth = 2
				.bRoundedBorder = True
				.Moveable = False
			End With

			'lblAgent initial props
			lblAgent = New UILabel(oUILib)
			With lblAgent
				.ControlName = "lblAgent"
				.Left = 5
				.Top = 2
				.Width = 165
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = ""
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
				.bAcceptEvents = False
			End With
			Me.AddChild(CType(lblAgent, UIControl))

			'lblSkill initial props
			lblSkill = New UILabel(oUILib)
			With lblSkill
				.ControlName = "lblSkill"
				.Left = 5
				.Top = 20
				.Width = 165
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = ""
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
				.bAcceptEvents = False
			End With
			Me.AddChild(CType(lblSkill, UIControl))

			'lblScore initial props
			lblScore = New UILabel(oUILib)
			With lblScore
				.ControlName = "lblScore"
				.Left = 5
				.Top = 40
				.Width = 160
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = ""
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(5, DrawTextFormat)
				.bAcceptEvents = False
			End With
			Me.AddChild(CType(lblScore, UIControl))
		End Sub

		Public Sub SetFromAgentAssignment(ByRef oAA As AgentAssignment)
            moAA = oAA
            Dim sText As String = ""
            sText = oAA.oAgent.sAgentName
            If lblAgent.Caption <> sText Then lblAgent.Caption = sText

			Dim lInfTarget As Int32 = -1
			Dim iInfTargetTypeID As Int16 = -1
			Select Case CType(oAA.oParent.oMission.oMission.ProgramControlID, eMissionResult)
				Case eMissionResult.eFindMineral, eMissionResult.eSearchAndRescueAgent, eMissionResult.eReconPlanetMap
					lInfTarget = oAA.oParent.oMission.lTargetID2 : iInfTargetTypeID = oAA.oParent.oMission.iTargetTypeID2
				Case eMissionResult.eGeologicalSurvey
					lInfTarget = oAA.oParent.oMission.lTargetID : iInfTargetTypeID = oAA.oParent.oMission.iTargetTypeID
				Case Else
					'Player....
					lInfTarget = oAA.oParent.oMission.lTargetPlayerID : iInfTargetTypeID = ObjectType.ePlayer
			End Select

            If oAA.oAgent Is Nothing = False Then 'AndAlso oAA.oAgent.bRequestedSkillList = False Then
                oAA.oAgent.CheckRequestSkillList()
                'oAA.oAgent.bRequestedSkillList = True
                'Dim yOut(5) As Byte
                'System.BitConverter.GetBytes(GlobalMessageCode.eGetSkillList).CopyTo(yOut, 0)
                'System.BitConverter.GetBytes(oAA.oAgent.ObjectID).CopyTo(yOut, 2)
                'MyBase.moUILib.SendMsgToPrimary(yOut)
            End If


            sText = oAA.oSkill.SkillName
			For X As Int32 = 0 To oAA.oAgent.SkillUB
				If oAA.oAgent.Skills(X).ObjectID = oAA.oSkill.ObjectID Then
					sText &= " (" & oAA.oAgent.SkillProf(X) & ")"
					Exit For
				End If
			Next X
            If lblSkill.Caption <> sText Then lblSkill.Caption = sText
            Dim clrCaption As System.Drawing.Color = System.Drawing.Color.FromArgb(255, 128, 128, 128)
			If oAA.oParent.oGoal.MissionPhase = eMissionPhase.ePreparationTime Then
				If (oAA.yStatus And AgentAssignmentStatus.Skipped) <> 0 Then
                    sText = "Skipped"
                    clrCaption = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                ElseIf (oAA.oAgent.lAgentStatus And AgentStatus.IsInfiltrated) = 0 OrElse oAA.oAgent.lTargetID <> lInfTarget OrElse oAA.oAgent.iTargetTypeID <> iInfTargetTypeID Then
                    sText = "Agent Not Present"
                    clrCaption = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                Else
                    sText = oAA.PointsProduced.ToString & " / " & oAA.PointsRequired.ToString
                    If oAA.PointsProduced < 0 Then
                        clrCaption = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                    ElseIf oAA.PointsProduced >= oAA.PointsRequired Then
                        clrCaption = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                    ElseIf oAA.PointsProduced <> 0 Then
                        clrCaption = Color.Cyan
                    Else
                        clrCaption = muSettings.InterfaceBorderColor
                    End If
                End If
            Else
                If (oAA.yStatus And AgentAssignmentStatus.Finished) <> 0 Then
                    If (oAA.yStatus And AgentAssignmentStatus.Success) <> 0 Then
                        sText = "Succeeded"
                        clrCaption = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                    Else
                        sText = "Failure"
                        clrCaption = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                    End If
                    If lblScore.Caption <> sText Then lblScore.Caption = sText
                    If lblScore.ForeColor <> clrCaption Then lblScore.ForeColor = clrCaption
                ElseIf (oAA.oAgent.lAgentStatus And AgentStatus.IsInfiltrated) = 0 OrElse oAA.oAgent.lTargetID <> lInfTarget OrElse oAA.oAgent.iTargetTypeID <> iInfTargetTypeID Then
                    sText = "Agent Not Present"
                    clrCaption = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                Else
                    sText = "Working"
                    clrCaption = muSettings.InterfaceBorderColor
                End If
            End If
            If lblScore.Caption <> sText Then lblScore.Caption = sText
            If lblScore.ForeColor <> clrCaption Then lblScore.ForeColor = clrCaption
		End Sub

		Private Sub ctlGoalAA_OnMouseDown(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseDown
			If moAA Is Nothing = False Then RaiseEvent AAClick(moAA)
		End Sub

		Public Sub ctlGoalAA_OnNewFrame() Handles Me.OnNewFrame
            If moAA Is Nothing = False Then
                Dim lInfTarget As Int32 = -1
                Dim iInfTargetTypeID As Int16 = -1
                Select Case CType(moAA.oParent.oMission.oMission.ProgramControlID, eMissionResult)
                    Case eMissionResult.eFindMineral, eMissionResult.eReconPlanetMap
                        lInfTarget = moAA.oParent.oMission.lTargetID2 : iInfTargetTypeID = moAA.oParent.oMission.iTargetTypeID2
                    Case eMissionResult.eGeologicalSurvey
                        lInfTarget = moAA.oParent.oMission.lTargetID : iInfTargetTypeID = moAA.oParent.oMission.iTargetTypeID
                    Case eMissionResult.eSearchAndRescueAgent
                        Dim lTmp As Int32 = -1
                        lTmp = moAA.oParent.oMission.lTargetID
                        For X As Int32 = 0 To goCurrentPlayer.AgentUB
                            If goCurrentPlayer.AgentIdx(X) = lTmp Then
                                'GetCacheObjectValue(lTargetID, iTargetTypeID)
                                lInfTarget = goCurrentPlayer.Agents(X).lTargetID
                                iInfTargetTypeID = goCurrentPlayer.Agents(X).iTargetTypeID
                                Exit For
                            End If
                        Next X

                        'Dim oTmpAgent As Agent = GetEpicaAgent(Me.lTargetID)
                        'If oTmpAgent Is Nothing = False AndAlso (oTmpAgent.lAgentStatus And AgentStatus.HasBeenCaptured) <> 0 Then
                        ' lInfTarget = oTmpAgent.lCapturedBy
                        ' iInfTargetTypeID = ObjectType.ePlayer
                        ' End If
                    Case Else
                        'Player....
                        lInfTarget = moAA.oParent.oMission.lTargetPlayerID : iInfTargetTypeID = ObjectType.ePlayer
                End Select

                If lblAgent.Caption <> moAA.oAgent.sAgentName Then lblAgent.Caption = moAA.oAgent.sAgentName

                If moAA.oAgent Is Nothing = False Then 'AndAlso moAA.oAgent.bRequestedSkillList = False Then
                    'moAA.oAgent.bRequestedSkillList = True
                    'Dim yOut(5) As Byte
                    'System.BitConverter.GetBytes(GlobalMessageCode.eGetSkillList).CopyTo(yOut, 0)
                    'System.BitConverter.GetBytes(moAA.oAgent.ObjectID).CopyTo(yOut, 2)
                    'MyBase.moUILib.SendMsgToPrimary(yOut)
                End If

                Dim sText As String = moAA.oSkill.SkillName
                Dim clrCaption As System.Drawing.Color

                For X As Int32 = 0 To moAA.oAgent.SkillUB
                    If moAA.oAgent.Skills(X).ObjectID = moAA.oSkill.ObjectID Then
                        sText &= " (" & moAA.oAgent.SkillProf(X) & ")"
                        Exit For
                    End If
                Next X
                If lblSkill.Caption <> sText Then lblSkill.Caption = sText
                sText = ""
                If moAA.oParent.oGoal.MissionPhase = eMissionPhase.ePreparationTime Then
                    If (moAA.yStatus And AgentAssignmentStatus.Skipped) <> 0 Then
                        sText = "Skipped"
                        clrCaption = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                    ElseIf (moAA.oAgent.lAgentStatus And AgentStatus.IsInfiltrated) = 0 OrElse moAA.oAgent.lTargetID <> lInfTarget OrElse moAA.oAgent.iTargetTypeID <> iInfTargetTypeID Then
                        sText = "Agent Not Present"
                        clrCaption = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                    Else
                        sText = moAA.PointsProduced.ToString & " / " & moAA.PointsRequired.ToString
                        If moAA.PointsProduced < 0 Then
                            clrCaption = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                        ElseIf moAA.PointsProduced >= moAA.PointsRequired Then
                            clrCaption = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                        ElseIf moAA.PointsProduced <> 0 Then
                            clrCaption = Color.Cyan
                        Else
                            clrCaption = muSettings.InterfaceBorderColor
                        End If
                    End If
                Else
                    If (moAA.yStatus And AgentAssignmentStatus.Finished) <> 0 Then
                        If (moAA.yStatus And AgentAssignmentStatus.Success) <> 0 Then
                            sText = "Succeeded"
                            clrCaption = System.Drawing.Color.FromArgb(255, 0, 255, 0)
                        Else
                            sText = "Failure"
                            clrCaption = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                        End If
                    ElseIf (moAA.oAgent.lAgentStatus And AgentStatus.IsInfiltrated) = 0 OrElse moAA.oAgent.lTargetID <> lInfTarget OrElse moAA.oAgent.iTargetTypeID <> iInfTargetTypeID Then
                        sText = "Agent Not Present"
                        clrCaption = System.Drawing.Color.FromArgb(255, 255, 0, 0)
                    Else
                        sText = "Working"
                        clrCaption = muSettings.InterfaceBorderColor
                    End If
                End If
                If lblScore.Caption <> sText Then lblScore.Caption = sText
                If lblScore.ForeColor <> clrCaption Then lblScore.ForeColor = clrCaption
            End If
		End Sub
	End Class
#End Region

	
End Class