Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmViewMission
	'Interface created from Interface Builder
	Private Class fraGoalList
		Inherits UIWindow 

		Private moPM As PlayerMission = Nothing

		Private WithEvents muGSS1 As ctlGoalSS
		Private WithEvents muGSS2 As ctlGoalSS
		Private WithEvents muGSS3 As ctlGoalSS
		Private WithEvents muGSS4 As ctlGoalSS
		Private vscrScroller As UIScrollBar
		Private WithEvents btnSkipAssignment As UIButton
		Private mfraPopup As frmAgent = Nothing

		Private moSelectedAA As AgentAssignment = Nothing

		Public Sub New(ByRef oUILib As UILib)
			MyBase.New(oUILib)

			'fraGoalList initial props
			With Me
				.ControlName = "fraGoalList"
				.Left = 73
				.Top = 106
				.Width = 780
				.Height = 480
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
			End With

			'muGSS1 initial props
			muGSS1 = New ctlGoalSS(oUILib)
			With muGSS1
				.ControlName = "muGSS1"
				.Left = 10
				.Top = 15
				.Width = 520
				.Height = 100
				.Enabled = True
				.Visible = False
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.Moveable = False
			End With
			Me.AddChild(CType(muGSS1, UIControl))

			'muGSS2 initial props
			muGSS2 = New ctlGoalSS(oUILib)
			With muGSS2
				.ControlName = "muGSS2"
				.Left = 10
				.Top = 130
				.Width = 520
				.Height = 100
				.Enabled = True
				.Visible = False
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.Moveable = False
			End With
			Me.AddChild(CType(muGSS2, UIControl))

			'muGSS3 initial props
			muGSS3 = New ctlGoalSS(oUILib)
			With muGSS3
				.ControlName = "muGSS3"
				.Left = 10
				.Top = 245
				.Width = 520
				.Height = 100
				.Enabled = True
				.Visible = False
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.Moveable = False
			End With
			Me.AddChild(CType(muGSS3, UIControl))

			'muGSS4 initial props
			muGSS4 = New ctlGoalSS(oUILib)
			With muGSS4
				.ControlName = "muGSS4"
				.Left = 10
				.Top = 360
				.Width = 520
				.Height = 100
				.Enabled = True
				.Visible = False
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.Moveable = False
			End With
			Me.AddChild(CType(muGSS4, UIControl))

			'vscrScroller initial props
			vscrScroller = New UIScrollBar(oUILib, True)
			With vscrScroller
				.ControlName = "vscrScroller"
				.Left = 535
				.Top = 15
				.Width = 24
				.Height = 450
				.Enabled = False
				.Visible = True
				.Value = 0
				.MaxValue = 0
				.MinValue = 0
				.SmallChange = 1
				.LargeChange = 1
				.ReverseDirection = True
			End With
			Me.AddChild(CType(vscrScroller, UIControl))

			'btnSkipAssignment initial props
			btnSkipAssignment = New UIButton(oUILib)
			With btnSkipAssignment
				.ControlName = "btnSkipAssignment"
				.Left = 620
				.Top = 405
				.Width = 100
				.Height = 24
				.Enabled = True
				.Visible = False
				.Caption = "Skip"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
				.ToolTipText = "Click to skip this preparation step." & vbCrLf & "Each preparation step skipped incurs a" & vbCrLf & "penalty to all subsequent actions on this mission."
			End With
			Me.AddChild(CType(btnSkipAssignment, UIControl))
		End Sub

		Public Sub SetPlayerMission(ByRef oPM As PlayerMission)
			moPM = oPM

			Dim lCnt As Int32 = 0
			For X As Int32 = 0 To moPM.oMission.GoalUB
				If moPM.oMission.MethodIDs(X) = moPM.lMethodID Then
					lCnt += 1
					Select Case lCnt
						Case 1
							muGSS1.SetFromPlayerMissionGoal(moPM.oMissionGoals(X))
						Case 2
							muGSS2.SetFromPlayerMissionGoal(moPM.oMissionGoals(X))
						Case 3
							muGSS3.SetFromPlayerMissionGoal(moPM.oMissionGoals(X))
						Case 4
							muGSS4.SetFromPlayerMissionGoal(moPM.oMissionGoals(X))
					End Select
				End If
			Next X

			vscrScroller.Value = 0
			If lCnt > 4 Then
				vscrScroller.Enabled = True
				vscrScroller.MaxValue = lCnt - 4
			Else
				vscrScroller.Enabled = False
				vscrScroller.MaxValue = 1
			End If
		End Sub

		Public Sub fraGoalList_OnNewFrame() Handles Me.OnNewFrame
			If muGSS1 Is Nothing = False Then muGSS1.ctlGoalSS_OnNewFrame()
			If muGSS2 Is Nothing = False Then muGSS2.ctlGoalSS_OnNewFrame()
			If muGSS3 Is Nothing = False Then muGSS3.ctlGoalSS_OnNewFrame()
			If muGSS4 Is Nothing = False Then muGSS4.ctlGoalSS_OnNewFrame()
		End Sub

		Private Sub fraGoalList_OnRender() Handles Me.OnRender
			If moPM Is Nothing = False Then

				Dim lCnt As Int32 = 1
				If vscrScroller Is Nothing = False AndAlso vscrScroller.Enabled = True Then
					'	lCnt = vscrScroller.MaxValue - vscrScroller.Value + 1
					lCnt = vscrScroller.Value + 1
				End If

				Dim lIdx As Int32 = 0
				For X As Int32 = 0 To moPM.oMission.GoalUB
					If moPM.oMission.MethodIDs(X) = moPM.lMethodID Then
						lCnt -= 1
						If lCnt = 0 Then
							lIdx = X
							Exit For
						End If
					End If
				Next X

				muGSS1.Visible = False
				muGSS2.Visible = False
				muGSS3.Visible = False
				muGSS4.Visible = False

				lCnt = 0
				For X As Int32 = lIdx To moPM.oMission.GoalUB
					If moPM.oMission.MethodIDs(X) = moPM.lMethodID Then
						lCnt += 1
						Select Case lCnt
							Case 1
								muGSS1.SetFromPlayerMissionGoal(moPM.oMissionGoals(X))
								muGSS1.Visible = True
							Case 2
								muGSS2.SetFromPlayerMissionGoal(moPM.oMissionGoals(X))
								muGSS2.Visible = True
							Case 3
								muGSS3.SetFromPlayerMissionGoal(moPM.oMissionGoals(X))
								muGSS3.Visible = True
							Case 4
								muGSS4.SetFromPlayerMissionGoal(moPM.oMissionGoals(X))
								muGSS4.Visible = True
								Exit For
						End Select
					End If
				Next X
			End If

		End Sub

		Private Sub muGSS1_AAClicked(ByRef oAA As AgentAssignment, ByRef oPMG As PlayerMissionGoal) Handles muGSS1.AAClicked, muGSS2.AAClicked, muGSS3.AAClicked, muGSS4.AAClicked
			If oAA Is Nothing = False AndAlso oPMG Is Nothing = False Then
				moSelectedAA = oAA
				DisplayPopup(oAA.oAgent)
				If oPMG.oGoal.MissionPhase = eMissionPhase.ePreparationTime AndAlso oPMG.oMission.yCurrentPhase <= eMissionPhase.ePreparationTime Then
					btnSkipAssignment.Visible = True
					btnSkipAssignment.Enabled = oAA.PointsProduced <> oAA.PointsRequired
					If (oAA.yStatus And AgentAssignmentStatus.Skipped) <> 0 Then
						btnSkipAssignment.Caption = "Resume"
					Else : btnSkipAssignment.Caption = "Skip"
					End If
				Else : btnSkipAssignment.Visible = False
				End If
			End If
		End Sub

		Private Sub DisplayPopup(ByRef oAgent As Agent)
			If mfraPopup Is Nothing Then
				mfraPopup = New frmAgent(MyBase.moUILib, True, True)
				Me.AddChild(CType(mfraPopup, UIControl))
			End If
			With mfraPopup
				.Left = 559
				.Top = 15
				.BorderLineWidth = 2
				.Moveable = False
				.Visible = True
				.Enabled = True
			End With

			mfraPopup.SetAgentIfNeeded(oAgent)
		End Sub

		Protected Overrides Sub Finalize()
			Try
				If mfraPopup Is Nothing = False Then
					MyBase.moUILib.RemoveWindow(mfraPopup.ControlName)
				End If
				mfraPopup = Nothing
			Catch
			End Try
			MyBase.Finalize()
		End Sub

		Private Sub btnSkipAssignment_Click(ByVal sName As String) Handles btnSkipAssignment.Click
			If moSelectedAA Is Nothing = False Then
				If btnSkipAssignment.Caption.ToUpper = "CONFIRM" Then
					btnSkipAssignment.Visible = False
					If mfraPopup Is Nothing = False Then mfraPopup.Visible = False
					Dim yMsg(18) As Byte
					System.BitConverter.GetBytes(GlobalMessageCode.eSetSkipStatus).CopyTo(yMsg, 0)
					System.BitConverter.GetBytes(moPM.PM_ID).CopyTo(yMsg, 2)
					System.BitConverter.GetBytes(moSelectedAA.oParent.oGoal.ObjectID).CopyTo(yMsg, 6)
					System.BitConverter.GetBytes(moSelectedAA.oAgent.ObjectID).CopyTo(yMsg, 10)
					System.BitConverter.GetBytes(moSelectedAA.oSkill.ObjectID).CopyTo(yMsg, 14)
					If (moSelectedAA.yStatus And AgentAssignmentStatus.Skipped) <> 0 Then
						yMsg(18) = 0
					Else : yMsg(18) = 1
					End If
					MyBase.moUILib.SendMsgToPrimary(yMsg)
				Else : btnSkipAssignment.Caption = "Confirm"
				End If
			End If
		End Sub
	End Class
End Class
