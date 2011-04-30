Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmMission
    'Interface created from Interface Builder
    Private Class fraSkillSet
        Inherits UIWindow

        Private WithEvents lstSkillset As UIListBox
        Private lblSelect As UILabel
		Private WithEvents hscrAssignments As UIScrollBar

		Private moAgentAssign() As ctlAgentAssignment

        Private moGoal As Goal
		Private moPM As PlayerMission

        Public Event AgentAssignment(ByRef oSkillSetSkill As SkillSet_Skill, ByRef oSender As ctlAgentAssignment)
        Public Event SkillSetSelected(ByRef oGoal As Goal, ByRef oSkillset As SkillSet)
        Public Event SetSkillsetFilter(ByRef oSkillset As SkillSet)

        Private mlCurrentSkillset As Int32
		Private mbIgnoreEvents As Boolean = False

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraSkillSet initial props
            With Me
                .ControlName = "fraSkillSet"
                .Left = 64
                .Top = 405
				.Width = 630 '780
                .Height = 220
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
            End With

            'lstSkillset initial props
            lstSkillset = New UIListBox(oUILib)
            With lstSkillset
                .ControlName = "lstSkillset"
                .Left = 5
                .Top = 25
                .Width = 165
                .Height = 190
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
            End With
            Me.AddChild(CType(lstSkillset, UIControl))

            'lblSelect initial props
            lblSelect = New UILabel(oUILib)
            With lblSelect
                .ControlName = "lblSelect"
                .Left = 5
                .Top = 8
                .Width = 129
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Select Skillset to Use"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblSelect, UIControl))

			ReDim moAgentAssign(2)

            'moAgentAssign(0) initial props
            moAgentAssign(0) = New ctlAgentAssignment(oUILib)
            With moAgentAssign(0)
                .ControlName = "moAgentAssign(0)"
                .Left = 175
                .Top = 5
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
            End With
            Me.AddChild(CType(moAgentAssign(0), UIControl))

            'moAgentAssign(1) initial props
            moAgentAssign(1) = New ctlAgentAssignment(oUILib)
            With moAgentAssign(1)
                .ControlName = "moAgentAssign(1)"
                .Left = 325
                .Top = 5
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
            End With
            Me.AddChild(CType(moAgentAssign(1), UIControl))

            'moAgentAssign(2) initial props
            moAgentAssign(2) = New ctlAgentAssignment(oUILib)
            With moAgentAssign(2)
                .ControlName = "moAgentAssign(2)"
                .Left = 475
                .Top = 5
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
            End With
            Me.AddChild(CType(moAgentAssign(2), UIControl))

			''moAgentAssign(3) initial props
			'moAgentAssign(3) = New ctlAgentAssignment(oUILib)
			'With moAgentAssign(3)
			'	.ControlName = "moAgentAssign(3)"
			'	.Left = 625
			'	.Top = 5
			'	.Enabled = True
			'	.Visible = True
			'	.BorderColor = muSettings.InterfaceBorderColor
			'	.FillColor = muSettings.InterfaceFillColor
			'	.FullScreen = False
			'End With
			'Me.AddChild(CType(moAgentAssign(3), UIControl))

			'hscrAssignments initial props
            hscrAssignments = New UIScrollBar(oUILib, False)
            With hscrAssignments
                .ControlName = "hscrAssignments"
                .Left = 175
                .Top = 190
				.Width = 450	'600
                .Height = 24
                .Enabled = True
                .Visible = True
                .Value = 0
                .MaxValue = 100
                .MinValue = 0
                .SmallChange = 1
                .LargeChange = 1
                .ReverseDirection = False
            End With
            Me.AddChild(CType(hscrAssignments, UIControl))

			For X As Int32 = 0 To moAgentAssign.GetUpperBound(0)
				AddHandler moAgentAssign(X).AssignButtonClicked, AddressOf AgentAssignmentButtonClick
			Next
            
            lstSkillset.ListIndex = -1
            lstSkillset_ItemClick(-1)
        End Sub

        Private Sub AgentAssignmentButtonClick(ByRef oSkillSet_Skill As SkillSet_Skill, ByRef oSender As ctlAgentAssignment)
            RaiseEvent AgentAssignment(oSkillSet_Skill, oSender)
        End Sub

		Public Sub SetFromGoal(ByRef oGoal As Goal, ByRef oPM As PlayerMission)
			moGoal = oGoal
			moPM = oPM

			'Ok, create our skillsets...
			mbIgnoreEvents = True
			lstSkillset.Clear()
			If moGoal Is Nothing = False Then
				For X As Int32 = 0 To moGoal.SkillSetUB
					If moGoal.SkillSets(X) Is Nothing = False Then
						lstSkillset.AddItem(moGoal.SkillSets(X).sSkillSetName)
						lstSkillset.ItemData(lstSkillset.NewIndex) = moGoal.SkillSets(X).SkillSetID
					End If
				Next X
			End If
			'If lstSkillset.ListCount > 0 Then
			'	lstSkillset.ListIndex = 0
			'Else : 
			'End If
			lstSkillset.ListIndex = -1
			mbIgnoreEvents = False
		End Sub

		Public Sub SetSkillsetID(ByVal lSkillsetID As Int32)
			mbIgnoreEvents = True
			For X As Int32 = 0 To lstSkillset.ListCount - 1
				If lstSkillset.ItemData(X) = lSkillsetID Then
					lstSkillset.ListIndex = X
					Exit For
				End If
			Next X
			mbIgnoreEvents = False
		End Sub

		Public Sub SetAgentAssignment(ByRef oAA As AgentAssignment)
			If moAgentAssign Is Nothing = False Then
				Dim lSSID As Int32 = oAA.oParent.oSkillSet.SkillSetID

				For X As Int32 = 0 To moAgentAssign.GetUpperBound(0)
					If moAgentAssign(X).SkillsetID = lSSID AndAlso moAgentAssign(X).SkillID = oAA.oSkill.ObjectID Then
						moAgentAssign(X).SetAgent(oAA.oAgent)
					End If
				Next X
			End If
		End Sub

        Private Sub lstSkillset_ItemClick(ByVal lIndex As Integer) Handles lstSkillset.ItemClick
            Dim lID As Int32 = -1
            If lstSkillset.ListIndex > -1 Then
                lID = lstSkillset.ItemData(lstSkillset.ListIndex)
            End If

            Dim oSkillSet As SkillSet = Nothing

            If lID = mlCurrentSkillset Then
                If lID > 0 Then
                    For X As Int32 = 0 To moGoal.SkillSetUB
                        If moGoal.SkillSets(X) Is Nothing = False AndAlso moGoal.SkillSets(X).SkillSetID = lID Then
                            oSkillSet = moGoal.SkillSets(X)
                            Exit For
                        End If
                    Next X
                End If

                RaiseEvent SetSkillsetFilter(oSkillSet)
                Return
            End If
            mlCurrentSkillset = lID

            'ok, set all moAgentAssigns to invisible and clear
            If moAgentAssign Is Nothing = False Then
                For X As Int32 = 0 To moAgentAssign.GetUpperBound(0)
                    moAgentAssign(X).SetAgent(Nothing)
                    moAgentAssign(X).SetSkillSetSkill(Nothing)
                Next X
            End If

            hscrAssignments.Enabled = False
            hscrAssignments.Value = 0
            hscrAssignments.MaxValue = 0

            If lID > 0 Then
                For X As Int32 = 0 To moGoal.SkillSetUB
                    If moGoal.SkillSets(X) Is Nothing = False AndAlso moGoal.SkillSets(X).SkillSetID = lID Then
                        oSkillSet = moGoal.SkillSets(X)
                        Exit For
                    End If
                Next X
                If oSkillSet Is Nothing = False Then
                    Dim lCnt As Int32 = RefreshSkillList()
                    If lCnt - 1 > moAgentAssign.GetUpperBound(0) Then
                        'ok...
                        hscrAssignments.MaxValue = lCnt - moAgentAssign.GetUpperBound(0) - 1
                        hscrAssignments.Enabled = True
                        hscrAssignments.Value = 0
                    Else
                        hscrAssignments.Value = 0
                        hscrAssignments.Enabled = False
                    End If

                    If mbIgnoreEvents = False Then RaiseEvent SkillSetSelected(moGoal, oSkillSet)
                End If
            End If

        End Sub

		Private Sub hscrAssignments_ValueChange() Handles hscrAssignments.ValueChange
			Dim lCnt As Int32 = RefreshSkillList()
		End Sub

		Private Function RefreshSkillList() As Int32
			'ok, get our selected skillset
			Dim lCnt As Int32 = 0
			Dim oSkillSet As SkillSet = Nothing
			If lstSkillset.ListIndex > -1 Then
				Dim lID As Int32 = lstSkillset.ItemData(lstSkillset.ListIndex)
				For X As Int32 = 0 To moGoal.SkillSetUB
					If moGoal.SkillSets(X) Is Nothing = False AndAlso moGoal.SkillSets(X).SkillSetID = lID Then
						oSkillSet = moGoal.SkillSets(X)
						Exit For
					End If
				Next X
			End If

			'do we have skillset?
			If oSkillSet Is Nothing = False Then
				'ok, now... get our mission goal (if there is one)
				Dim oMG As PlayerMissionGoal = Nothing
				If moPM Is Nothing = False Then
					For X As Int32 = 0 To moPM.oMission.GoalUB
						If moPM.oMission.Goals(X).ObjectID = moGoal.ObjectID AndAlso moPM.lMethodID = moPM.oMission.MethodIDs(X) Then
							oMG = moPM.oMissionGoals(X)
							Exit For
						End If
					Next X
				End If

				For X As Int32 = 0 To oSkillSet.SkillUB
					If oSkillSet.Skills(X) Is Nothing = False Then
						lCnt += 1
						If lCnt >= hscrAssignments.Value + 1 Then
							If lCnt - hscrAssignments.Value - 1 < moAgentAssign.GetUpperBound(0) + 1 Then
								moAgentAssign(lCnt - hscrAssignments.Value - 1).SetSkillSetSkill(oSkillSet.Skills(X))
								moAgentAssign(lCnt - hscrAssignments.Value - 1).SetAgent(Nothing)
								If oMG Is Nothing = False Then
									For Y As Int32 = 0 To oMG.lAssignmentUB
										If oMG.oAssignments(Y) Is Nothing = False AndAlso oMG.oAssignments(Y).oSkill Is Nothing = False AndAlso oMG.oAssignments(Y).oAgent Is Nothing = False Then
											If oMG.oAssignments(Y).oSkill.ObjectID = oSkillSet.Skills(X).oSkill.ObjectID Then
												If oMG.oAssignments(Y).oAgent Is Nothing = False Then
													moAgentAssign(lCnt - hscrAssignments.Value - 1).SetAgent(oMG.oAssignments(Y).oAgent)
												End If
												Exit For
											End If
										End If
									Next Y
								End If

								moAgentAssign(lCnt - hscrAssignments.Value - 1).Visible = True
							End If
						End If
					End If
				Next X
			End If

			Return lCnt
        End Function

	End Class
End Class
