Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmAgent

    Public Class fraProficiencies
        Inherits UIWindow

        Private lblDagger As UILabel
        Private lblInfiltration As UILabel
        Private lblResourcefulness As UILabel
        Private lblSuspicion As UILabel

        Private mrcDagger As Rectangle = New Rectangle(32, 48, 16, 16)
        Private mrcInfiltration As Rectangle = New Rectangle(32, 32, 16, 16)
        Private mrcResource As Rectangle = New Rectangle(0, 0, 16, 16)
        Private mrcSuspicion As Rectangle = New Rectangle(48, 48, 16, 16)

		Private moSprite As Sprite

		Private mbParentIsChild As Boolean = False

		Private mbKillAgentMainTex As Boolean = False

		Public Sub New(ByRef oUILib As UILib, ByVal bParentIsChild As Boolean)
			MyBase.New(oUILib)

			mbParentIsChild = bParentIsChild

			Device.IsUsingEventHandlers = False
			moSprite = New Sprite(oUILib.oDevice)
			Device.IsUsingEventHandlers = True

			'fraProficiencies initial props
			With Me
				.ControlName = "fraProficiencies"
				.Left = 127
				.Top = 132
				.Width = 190 '175
				.Height = 95
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.BorderLineWidth = 2
				.Caption = "Proficiencies"
			End With

			'lblDagger initial props
			lblDagger = New UILabel(oUILib)
			With lblDagger
				.ControlName = "lblDagger"
				.Left = 27
				.Top = 10
				.Width = 135
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Dagger: XXX"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
				.ToolTipText = "Dagger represents this agent's black-ops ability." & vbCrLf & _
							   "This score is important for counter agents and for" & vbCrLf & _
							   "black ops work such as assassination and sabotage." & vbCrLf & _
							   "Higher values are better while lower values would" & vbCrLf & _
							   "represent an agent more suitable for espionage."
			End With
			Me.AddChild(CType(lblDagger, UIControl))

			'lblInfiltration initial props
			lblInfiltration = New UILabel(oUILib)
			With lblInfiltration
				.ControlName = "lblInfiltration"
				.Left = 27
				.Top = 30
				.Width = 135
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Infiltration: XXX"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
				.ToolTipText = "This score represents the agent's ability to infiltrate empires." & vbCrLf & _
							   "Higher values represent a better success rate at gaining access" & vbCrLf & _
							   "to other empires including those that enforce the strictest" & vbCrLf & _
							   "security programs. Lower values are better suited for counter" & vbCrLf & _
							   "agent assignments within the homeland."
			End With
			Me.AddChild(CType(lblInfiltration, UIControl))

			'lblResourcefulness initial props
			lblResourcefulness = New UILabel(oUILib)
			With lblResourcefulness
				.ControlName = "lblResourcefulness"
				.Left = 27
				.Top = 50
				.Width = 135
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Resourcefulness: XXX"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
				.ToolTipText = "Resourcefulness is an ability that indicates the agent's" & vbCrLf & _
							   "ability to think on their feet, improvise, and obtain" & vbCrLf & _
							   "necessary supplies when those supplies are not readily" & vbCrLf & _
							   "available. Resourcefulness impacts a large amount of" & vbCrLf & _
							   "what an agent does while on missions. Low Resourcefulness" & vbCrLf & _
							   "can still be utilized in counter agent assignments."
			End With
			Me.AddChild(CType(lblResourcefulness, UIControl))

			'lblSuspicion initial props
			lblSuspicion = New UILabel(oUILib)
			With lblSuspicion
				.ControlName = "lblSuspicion"
				.Left = 27
				.Top = 70
				.Width = 135
				.Height = 18
				.Enabled = True
				.Visible = True
				.Caption = "Suspicion: XXX"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(4, DrawTextFormat)
				.ToolTipText = "This value represents the amount of suspicion the agent" & vbCrLf & _
							   "has from the currently infiltrated empire's authorities." & vbCrLf & _
							   "This value is an approximation as reported by the agent" & vbCrLf & _
							   "and may not be accurate. Higher values represent an" & vbCrLf & _
							   "increase in pressure from the authorities on this agent" & vbCrLf & _
							   "which can make it easier for this agent to get caught." & vbCrLf & _
							   "Certain activities can raise or lower this score over the" & vbCrLf & _
							   "course of a mission. This score is also very important" & vbCrLf & _
							   "when the agent attempts to deinfiltrate a target empire."
			End With
			Me.AddChild(CType(lblSuspicion, UIControl))
		End Sub

        Public Sub SetFromAgent(ByRef oAgent As Agent)
            lblDagger.Caption = "Dagger: " & oAgent.Dagger
            lblInfiltration.Caption = "Infiltration: " & oAgent.Infiltration
            lblResourcefulness.Caption = "Resourcefulness: " & oAgent.Resourcefulness
            lblSuspicion.Caption = "Suspicion: " & oAgent.Suspicion
        End Sub

        Private Sub fraProficiencies_OnRenderEnd() Handles Me.OnRenderEnd
			'Now, draw our icons
			If frmAgentMain.moAgentIcons Is Nothing OrElse frmAgentMain.moAgentIcons.Disposed = True Then frmAgentMain.moAgentIcons = goResMgr.LoadScratchTexture("AgentIcons.dds", "Interface.pak")

            With moSprite
                .Begin(SpriteFlags.AlphaBlend)
				Dim rcDest As Rectangle = New Rectangle(Me.Left + 10, Me.Top + 11, 16, 16)
				Dim lOffsetY As Int32 = 0
				If mbParentIsChild = True Then
					Dim ptTmp As Point = Me.ParentControl.GetAbsolutePosition()
					rcDest.X += ptTmp.X
					rcDest.Y += ptTmp.Y
					lOffsetY = ptTmp.Y
				End If
                .Draw2D(frmAgentMain.moAgentIcons, mrcDagger, rcDest, Point.Empty, 0, rcDest.Location, Color.White)
				rcDest.Y = Me.Top + 31 + lOffsetY
				.Draw2D(frmAgentMain.moAgentIcons, mrcInfiltration, rcDest, Point.Empty, 0, rcDest.Location, Color.White)
				rcDest.Y = Me.Top + 51 + lOffsetY
				.Draw2D(frmAgentMain.moAgentIcons, mrcResource, rcDest, Point.Empty, 0, rcDest.Location, Color.White)
				rcDest.Y = Me.Top + 71 + lOffsetY
				.Draw2D(frmAgentMain.moAgentIcons, mrcSuspicion, rcDest, Point.Empty, 0, rcDest.Location, Color.White)
                .End()
            End With
        End Sub

        Protected Overrides Sub Finalize()
			Try
				Dim oWin As UIWindow = MyBase.moUILib.GetWindow("frmAgentMain")
				If oWin Is Nothing = True OrElse oWin.Visible = False Then
					If frmAgentMain.moAgentIcons Is Nothing = False Then frmAgentMain.moAgentIcons.Dispose()
					frmAgentMain.moAgentIcons = Nothing
				End If
				oWin = Nothing
			Catch
			End Try
			Try
				If moSprite Is Nothing = False Then
					moSprite.Dispose()
				End If
				moSprite = Nothing
			Catch
			End Try
			MyBase.Finalize()
        End Sub
    End Class

    Public Class fraSkills
        Inherits UIWindow

        Private WithEvents lstSkills As UIListBox
        Private moAgent As Agent

        Private mbLoading As Boolean = True

        Private moSkillPopup As frmSkillDetail = Nothing

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraSkills initial props
            With Me
                .ControlName = "fraSkills"
                .Left = 154
                .Top = 137
				.Width = 210 '175
                .Height = 225
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 2
                .Caption = "Skills"
                .mbAcceptReprocessEvents = True
            End With

            'lstSkills initial props
            lstSkills = New UIListBox(oUILib)
            With lstSkills
                .ControlName = "lstSkills"
                .Left = 5
                .Top = 10
				.Width = 200 '165
                .Height = 210
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = Me.FillColor
                .ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(lstSkills, UIControl))

            mbLoading = False
        End Sub

        Public Sub SetFromAgent(ByRef oAgent As Agent)
            moAgent = oAgent
            Me.lstSkills.FillColor = Me.FillColor
			lstSkills.Clear()

            If moAgent Is Nothing = False Then ' AndAlso moAgent.bRequestedSkillList = False Then
                'moAgent.bRequestedSkillList = True
                'Dim yOut(5) As Byte
                'System.BitConverter.GetBytes(GlobalMessageCode.eGetSkillList).CopyTo(yOut, 0)
                'System.BitConverter.GetBytes(moAgent.ObjectID).CopyTo(yOut, 2)
                'MyBase.moUILib.SendMsgToPrimary(yOut)
                moAgent.CheckRequestSkillList()
            End If

            'Sort by proficiency
            Dim lSorted() As Int32 = Nothing
            Dim lSortedUB As Int32 = -1
            For X As Int32 = 0 To oAgent.SkillUB
                Dim lIdx As Int32 = -1

                For Y As Int32 = 0 To lSortedUB
                    If oAgent.SkillProf(lSorted(Y)) < oAgent.SkillProf(X) Then
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
            Next X

			For X As Int32 = 0 To lSortedUB 'oAgent.SkillUB
				Dim sText As String = oAgent.Skills(lSorted(X)).SkillName.PadRight(21, " "c)
				sText &= oAgent.SkillProf(lSorted(X))
				lstSkills.AddItem(sText, False)
				lstSkills.ItemData(lstSkills.NewIndex) = oAgent.Skills(lSorted(X)).ObjectID
			Next X
        End Sub

        Public Sub ClearSkillPopup()
            If moSkillPopup Is Nothing = False Then
                If MyBase.moUILib Is Nothing = False Then MyBase.moUILib.RemoveWindow(moSkillPopup.ControlName)
            End If
            moSkillPopup = Nothing
        End Sub

        Private Sub fraSkills_OnMouseMove(ByVal lMouseX As Integer, ByVal lMouseY As Integer, ByVal lButton As System.Windows.Forms.MouseButtons) Handles Me.OnMouseMove
            ClearSkillPopup()
		End Sub

		Public Sub fraSkills_OnNewFrame() Handles Me.OnNewFrame
			If moAgent Is Nothing = False AndAlso lstSkills.ListCount - 1 <> moAgent.SkillUB Then
				Me.SetFromAgent(moAgent)
			End If
		End Sub

        Private Sub fraSkills_OnResize() Handles Me.OnResize
            If mbLoading = False AndAlso lstSkills Is Nothing = False Then
                lstSkills.Height = Me.Height - 15
            End If
        End Sub

        Private Sub lstSkills_ItemMouseOver(ByVal lIndex As Integer, ByVal lX As Integer, ByVal lY As Integer) Handles lstSkills.ItemMouseOver
            If lIndex > -1 AndAlso moAgent Is Nothing = False Then
                Dim lID As Int32 = lstSkills.ItemData(lIndex)
                Dim oSkill As Skill = Nothing
                Dim sText As String = ""
                For X As Int32 = 0 To moAgent.SkillUB
                    If moAgent.Skills(X).ObjectID = lID Then
                        If moSkillPopup Is Nothing = True Then moSkillPopup = New frmSkillDetail(goUILib)
                        moSkillPopup.SetFromSkill(moAgent.Skills(X), moAgent.SkillProf(X))
                        moSkillPopup.Left = lX
                        moSkillPopup.Top = lY
                        MyBase.moUILib.BringWindowToForeground(moSkillPopup.ControlName)
                        Return
                    End If
                Next X

                ClearSkillPopup()
            End If
        End Sub

        Protected Overrides Sub Finalize()
			ClearSkillPopup()
			MyBase.Finalize()
        End Sub
    End Class
 
    Public Class fraInfiltration
        Inherits UIWindow

        Private lblCurrent As UILabel
        Private lblCurrentInfType As UILabel
        Private lblCurrentTarget As UILabel
        Private lblCurrentFreq As UILabel
        Private lblNew As UILabel
        Private lblNewType As UILabel
        Private lblNewFreq As UILabel
        Private lblNewTarget As UILabel
        Private WithEvents cboNewInfType As UIComboBox
        Private WithEvents cboNewTarget As UIComboBox
        Private WithEvents cboReportFreq As UIComboBox
        Private WithEvents btnDeinfiltrate As UIButton
        Private WithEvents btnSet As UIButton

        Private moAgent As Agent

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraInfiltration initial props
            With Me
                .ControlName = "fraInfiltration"
                .Left = 264
                .Top = 165
                .Width = 225
                .Height = 220
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 2
                .Caption = "Infiltration Status"
                .mbAcceptReprocessEvents = True
            End With

            'lblCurrent initial props
            lblCurrent = New UILabel(oUILib)
            With lblCurrent
                .ControlName = "lblCurrent"
                .Left = 10
                .Top = 10
                .Width = 181
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Current Infiltration Settings"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblCurrent, UIControl))

            'lblCurrentInfType initial props
            lblCurrentInfType = New UILabel(oUILib)
            With lblCurrentInfType
                .ControlName = "lblCurrentInfType"
                .Left = 20
                .Top = 27
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Type: General"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblCurrentInfType, UIControl))

            'lblCurrentTarget initial props
            lblCurrentTarget = New UILabel(oUILib)
            With lblCurrentTarget
                .ControlName = "lblCurrentTarget"
                .Left = 20
                .Top = 45
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Target: Csaj"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblCurrentTarget, UIControl))

            'lblCurrentFreq initial props
            lblCurrentFreq = New UILabel(oUILib)
            With lblCurrentFreq
                .ControlName = "lblCurrentFreq"
                .Left = 20
                .Top = 62
                .Width = 200
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Report Frequency: once per day"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblCurrentFreq, UIControl))

            'lblNew initial props
            lblNew = New UILabel(oUILib)
            With lblNew
                .ControlName = "lblNew"
                .Left = 10
                .Top = 85
                .Width = 163
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "New Infiltration Settings"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblNew, UIControl))

            'lblNewType initial props
            lblNewType = New UILabel(oUILib)
            With lblNewType
                .ControlName = "lblNewType"
                .Left = 20
                .Top = 105
                .Width = 36
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Type:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblNewType, UIControl))

            'lblNewTarget initial props
            lblNewTarget = New UILabel(oUILib)
            With lblNewTarget
                .ControlName = "lblNewTarget"
                .Left = 20
                .Top = 130
                .Width = 44
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Target:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblNewTarget, UIControl))

            'lblNewFreq initial props
            lblNewFreq = New UILabel(oUILib)
            With lblNewFreq
                .ControlName = "lblNewFreq"
                .Left = 20
                .Top = 153
                .Width = 44
                .Height = 18
                .Enabled = True
                .Visible = True
                .Caption = "Report:"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .DrawBackImage = False
                .FontFormat = CType(4, DrawTextFormat)
            End With
            Me.AddChild(CType(lblNewFreq, UIControl))

            'btnDeinfiltrate initial props
            btnDeinfiltrate = New UIButton(oUILib)
            With btnDeinfiltrate
                .ControlName = "btnDeinfiltrate"
                .Left = 10
                .Top = 190
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Deinfiltrate"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnDeinfiltrate, UIControl))

            'btnSet initial props
            btnSet = New UIButton(oUILib)
            With btnSet
                .ControlName = "btnSet"
                .Left = 115
                .Top = 190
                .Width = 100
                .Height = 24
                .Enabled = True
                .Visible = True
                .Caption = "Set Settings"
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                .DrawBackImage = True
                .FontFormat = CType(5, DrawTextFormat)
                .ControlImageRect = New Rectangle(0, 0, 120, 32)
            End With
            Me.AddChild(CType(btnSet, UIControl))

            'cboReportFreq initial props
            cboReportFreq = New UIComboBox(oUILib)
            With cboReportFreq
                .ControlName = "cboReportFreq"
                .Left = 80
                .Top = 155
                .Width = 140
                .Height = 18
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(cboReportFreq, UIControl))

            'cboNewTarget initial props
            cboNewTarget = New UIComboBox(oUILib)
            With cboNewTarget
                .ControlName = "cboNewTarget"
                .Left = 80
                .Top = 130
                .Width = 140
                .Height = 18
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(cboNewTarget, UIControl))

            'cboNewInfType initial props
            cboNewInfType = New UIComboBox(oUILib)
            With cboNewInfType
                .ControlName = "cboNewInfType"
                .Left = 80
                .Top = 105
                .Width = 140
                .Height = 18
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceTextBoxFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .DropDownListBorderColor = muSettings.InterfaceBorderColor
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(cboNewInfType, UIControl))

            FillData()
        End Sub

        Public Sub SetFromAgent(ByRef oAgent As Agent)
            moAgent = oAgent
            lblCurrentInfType.Caption = "Type: " & GetInfTypeText(oAgent.InfiltrationType)
			lblCurrentFreq.Caption = "Report Frequency: " & GetFreqText(oAgent.yReportFreq)

			Me.cboNewInfType.FindComboItemData(oAgent.InfiltrationType)
			For X As Int32 = 0 To Me.cboNewTarget.ListCount - 1
				If Me.cboNewTarget.ItemData(X) = oAgent.lTargetID AndAlso Me.cboNewTarget.ItemData2(X) = oAgent.iTargetTypeID Then
					Me.cboNewTarget.ListIndex = X
					Exit For
				End If
			Next X

			Me.cboReportFreq.FindComboItemData(oAgent.yReportFreq)
        End Sub

        Public Sub fraInfiltration_OnNewFrame() Handles Me.OnNewFrame
			Dim sText As String = "Target: "
			If moAgent.lTargetID > -1 AndAlso moAgent.lTargetID <> glPlayerID Then
				sText &= GetCacheObjectValue(moAgent.lTargetID, moAgent.iTargetTypeID)
			Else
				If (moAgent.lAgentStatus And AgentStatus.CounterAgent) <> 0 Then
					sText &= "Counter-Agent"
				Else : sText &= "Not Infiltrated"
				End If
			End If
			If lblCurrentTarget.Caption <> sText Then lblCurrentTarget.Caption = sText

            If (moAgent.lAgentStatus And AgentStatus.IsInfiltrated) <> 0 Then
                If btnDeinfiltrate.Enabled = False Then btnDeinfiltrate.Enabled = True
            ElseIf btnDeinfiltrate.Enabled = True Then
                btnDeinfiltrate.Enabled = False
            End If


            sText = "Type: " & GetInfTypeText(moAgent.InfiltrationType)
            If moAgent.InfiltrationLevel <> 0 Then sText &= " @ " & moAgent.InfiltrationLevel.ToString
            If sText <> lblCurrentInfType.Caption Then lblCurrentInfType.Caption = sText
            sText = "Report Frequency: " & GetFreqText(moAgent.yReportFreq)
            If sText <> lblCurrentFreq.Caption Then lblCurrentFreq.Caption = sText

            If cboNewTarget Is Nothing = False Then
                For X As Int32 = 0 To cboNewTarget.ListCount - 1
                    If cboNewTarget.ItemData2(X) = ObjectType.ePlayer AndAlso cboNewTarget.ItemData(X) = glPlayerID Then
                        sText = "ME (Counter Spy)"
                    Else : sText = GetCacheObjectValue(cboNewTarget.ItemData(X), CShort(cboNewTarget.ItemData2(X)))
                    End If
                    If cboNewTarget.List(X) <> sText Then cboNewTarget.List(X) = sText
                Next X
            End If
        End Sub

		Public Shared Function GetInfTypeText(ByVal lInfType As Byte) As String
			Select Case lInfType
				Case eInfiltrationType.eAgencyInfiltration
					Return "Agency"
				Case eInfiltrationType.eCapitalShipInfiltration
                    Return "Fleet"
				Case eInfiltrationType.eColonialInfiltration
                    Return "Colonial"
				Case eInfiltrationType.eCombatUnitInfiltration
                    Return "Unit"
				Case eInfiltrationType.eCorporationInfiltration
                    Return "Corporate"
				Case eInfiltrationType.eFederalInfiltration
                    Return "Government"
				Case eInfiltrationType.eGeneralInfiltration
					Return "General"
				Case eInfiltrationType.eMilitaryInfiltration
                    Return "Military"
				Case eInfiltrationType.eMiningInfiltration
					Return "Mining"
				Case eInfiltrationType.ePlanetInfiltration
					Return "Planet"
				Case eInfiltrationType.ePowerCenterInfiltration
                    Return "Power Center"
				Case eInfiltrationType.eProductionInfiltration
					Return "Production"
				Case eInfiltrationType.eResearchInfiltration
					Return "Research"
				Case eInfiltrationType.eSolarSystemInfiltration
					Return "Solar System"
				Case eInfiltrationType.eTradeInfiltration
					Return "Trade"
			End Select
			Return "Unknown"
		End Function
        Private Function GetFreqText(ByVal yFreq As Byte) As String
            Select Case yFreq
                Case eReportFreq.OncePerDay
                    Return "per day"
                Case eReportFreq.OncePerHalfHour
                    Return "per half hour"
                Case eReportFreq.OncePerHour
                    Return "per hour"
                Case eReportFreq.OncePerSixHours
                    Return "per 6 hours"
                Case eReportFreq.OncePerTwelveHours
                    Return "per 12 hours"
                Case eReportFreq.OncePerTwoDays
                    Return "per two days"
                Case eReportFreq.OncePerTwoHours
                    Return "per two hours"
                Case eReportFreq.OncePerWeek
                    Return "per week"
            End Select
            Return "per half hour"
        End Function
        Private Sub FillData()
            cboNewInfType.Clear() : cboNewTarget.Clear() : cboReportFreq.Clear()

            'general always at top
            cboNewInfType.AddItem("General") : cboNewInfType.ItemData(cboNewInfType.NewIndex) = eInfiltrationType.eGeneralInfiltration
            'alphabetized
            cboNewInfType.AddItem("Agency") : cboNewInfType.ItemData(cboNewInfType.NewIndex) = eInfiltrationType.eAgencyInfiltration
            'cboNewInfType.AddItem("Capital Ship") : cboNewInfType.ItemData(cboNewInfType.NewIndex) = eInfiltrationType.eCapitalShipInfiltration
            cboNewInfType.AddItem("Colony") : cboNewInfType.ItemData(cboNewInfType.NewIndex) = eInfiltrationType.eColonialInfiltration
            'cboNewInfType.AddItem("Combat Unit") : cboNewInfType.ItemData(cboNewInfType.NewIndex) = eInfiltrationType.eCombatUnitInfiltration
            'cboNewInfType.AddItem("Corporation") : cboNewInfType.ItemData(cboNewInfType.NewIndex) = eInfiltrationType.eCorporationInfiltration
            cboNewInfType.AddItem("Government") : cboNewInfType.ItemData(cboNewInfType.NewIndex) = eInfiltrationType.eFederalInfiltration
            cboNewInfType.AddItem("Military") : cboNewInfType.ItemData(cboNewInfType.NewIndex) = eInfiltrationType.eMilitaryInfiltration
            cboNewInfType.AddItem("Mining") : cboNewInfType.ItemData(cboNewInfType.NewIndex) = eInfiltrationType.eMiningInfiltration
            cboNewInfType.AddItem("Planet") : cboNewInfType.ItemData(cboNewInfType.NewIndex) = eInfiltrationType.ePlanetInfiltration
            cboNewInfType.AddItem("Power Center") : cboNewInfType.ItemData(cboNewInfType.NewIndex) = eInfiltrationType.ePowerCenterInfiltration
            cboNewInfType.AddItem("Production") : cboNewInfType.ItemData(cboNewInfType.NewIndex) = eInfiltrationType.eProductionInfiltration
            cboNewInfType.AddItem("Research") : cboNewInfType.ItemData(cboNewInfType.NewIndex) = eInfiltrationType.eResearchInfiltration
            cboNewInfType.AddItem("Solar System") : cboNewInfType.ItemData(cboNewInfType.NewIndex) = eInfiltrationType.eSolarSystemInfiltration
            cboNewInfType.AddItem("Trade") : cboNewInfType.ItemData(cboNewInfType.NewIndex) = eInfiltrationType.eTradeInfiltration

            'Now, report freqs
            cboReportFreq.AddItem("Per Half Hour") : cboReportFreq.ItemData(cboReportFreq.NewIndex) = eReportFreq.OncePerHalfHour
            cboReportFreq.AddItem("Per Hour") : cboReportFreq.ItemData(cboReportFreq.NewIndex) = eReportFreq.OncePerHour
            cboReportFreq.AddItem("Per Two Hours") : cboReportFreq.ItemData(cboReportFreq.NewIndex) = eReportFreq.OncePerTwoHours
			cboReportFreq.AddItem("Per Six Hours") : cboReportFreq.ItemData(cboReportFreq.NewIndex) = eReportFreq.OncePerSixHours
			cboReportFreq.AddItem("Per 12 Hours") : cboReportFreq.ItemData(cboReportFreq.NewIndex) = eReportFreq.OncePerTwelveHours
			cboReportFreq.AddItem("Daily") : cboReportFreq.ItemData(cboReportFreq.NewIndex) = eReportFreq.OncePerDay
			cboReportFreq.AddItem("Per Two Days") : cboReportFreq.ItemData(cboReportFreq.NewIndex) = eReportFreq.OncePerTwoDays
			cboReportFreq.AddItem("Weekly") : cboReportFreq.ItemData(cboReportFreq.NewIndex) = eReportFreq.OncePerWeek
        End Sub

		Private Sub btnDeinfiltrate_Click(ByVal sName As String) Handles btnDeinfiltrate.Click
			'check if this agent is on a mission
			If btnDeinfiltrate.Caption.ToUpper <> "CONFIRM" Then
				If moAgent.lMissionID <> -1 Then
					MyBase.moUILib.AddNotification("This will remove the agent from the selected empire and any missions that the agent may be participating in!", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				End If
				MyBase.moUILib.AddNotification("Click Confirm to confirm De-Infiltrate order.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				btnDeinfiltrate.Caption = "Confirm"
			Else
				Dim yMsg(5) As Byte
				System.BitConverter.GetBytes(GlobalMessageCode.eDeinfiltrateAgent).CopyTo(yMsg, 0)
				System.BitConverter.GetBytes(moAgent.ObjectID).CopyTo(yMsg, 2)
				MyBase.moUILib.SendMsgToPrimary(yMsg)

                MyBase.moUILib.AddNotification("De-Infiltrate Order Sent to Agent.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
                btnDeinfiltrate.Caption = "Deinfiltrate"
                btnDeinfiltrate.Enabled = False
			End If
		End Sub

		Private Sub btnSet_Click(ByVal sName As String) Handles btnSet.Click

			If cboNewInfType.ListIndex = -1 Then
				MyBase.moUILib.AddNotification("Select an Infiltration Type for the new setting.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If
			If cboNewTarget.ListIndex = -1 Then
				MyBase.moUILib.AddNotification("Select an Infiltration target for the new setting.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If
			If cboReportFreq.ListIndex = -1 Then
				MyBase.moUILib.AddNotification("Select a report frequency for the new setting.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
				Return
			End If

			Dim yMsg(19) As Byte
			Dim lPos As Int32 = 0
			System.BitConverter.GetBytes(GlobalMessageCode.eSetInfiltrateSettings).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(moAgent.ObjectID).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(cboNewInfType.ItemData(cboNewInfType.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(cboNewTarget.ItemData(cboNewTarget.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
			System.BitConverter.GetBytes(CShort(cboNewTarget.ItemData2(cboNewTarget.ListIndex))).CopyTo(yMsg, lPos) : lPos += 2
			System.BitConverter.GetBytes(cboReportFreq.ItemData(cboReportFreq.ListIndex)).CopyTo(yMsg, lPos) : lPos += 4
			MyBase.moUILib.SendMsgToPrimary(yMsg)

			If NewTutorialManager.TutorialOn = True Then
				NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eAgentInfiltrationSetSettings, -1, -1, -1, "")
			End If

			MyBase.moUILib.AddNotification("Infiltration Orders Submitted.", Color.Blue, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		End Sub

		Private Sub cboNewInfType_ItemSelected(ByVal lItemIndex As Integer) Handles cboNewInfType.ItemSelected
			'fill the target list...
			cboNewTarget.Clear()
			If cboNewInfType.ListIndex > -1 Then
				Dim lID As Int32 = cboNewInfType.ItemData(cboNewInfType.ListIndex)

				If NewTutorialManager.TutorialOn = True Then
					NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eAgentInfiltrationTypeSelected, lID, -1, -1, "")
				End If

				If lID = eInfiltrationType.ePlanetInfiltration Then
					'fill with planets
					If goCurrentPlayer.yPlayerPhase <> 255 Then
                        Dim oPlanet As Planet = Planet.GetTutorialPlanet()
						If oPlanet Is Nothing = False Then
							cboNewTarget.AddItem(oPlanet.PlanetName)
							cboNewTarget.ItemData(cboNewTarget.NewIndex) = oPlanet.ObjectID
							cboNewTarget.ItemData2(cboNewTarget.NewIndex) = oPlanet.ObjTypeID
						End If
					ElseIf goGalaxy Is Nothing = False Then
						If goGalaxy.CurrentSystemIdx > -1 Then
							Dim oSys As SolarSystem = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
							If oSys Is Nothing = False Then
								For X As Int32 = 0 To oSys.PlanetUB
									If oSys.moPlanets(X) Is Nothing = False Then
										cboNewTarget.AddItem(oSys.moPlanets(X).PlanetName)
										cboNewTarget.ItemData(cboNewTarget.NewIndex) = oSys.moPlanets(X).ObjectID
										cboNewTarget.ItemData2(cboNewTarget.NewIndex) = oSys.moPlanets(X).ObjTypeID
									End If
								Next X
							End If
						End If
					End If
				ElseIf lID = eInfiltrationType.eSolarSystemInfiltration Then
					'fill with systems
					If goGalaxy Is Nothing = False Then
						If goGalaxy.CurrentSystemIdx > -1 Then
							Dim oSys As SolarSystem = goGalaxy.moSystems(goGalaxy.CurrentSystemIdx)
							If oSys Is Nothing = False Then
								cboNewTarget.AddItem(oSys.SystemName)
								cboNewTarget.ItemData(cboNewTarget.NewIndex) = oSys.ObjectID
								cboNewTarget.ItemData2(cboNewTarget.NewIndex) = oSys.ObjTypeID
							End If
						End If
					End If
				Else
					'fill with players
					cboNewTarget.AddItem("ME (Counter)")
					cboNewTarget.ItemData(cboNewTarget.NewIndex) = glPlayerID
					cboNewTarget.ItemData2(cboNewTarget.NewIndex) = ObjectType.ePlayer

                    'Ok, let's sort our player list
                    Dim lIdx() As Int32 = GetSortedPlayerRelIdxArray()
                    If lIdx Is Nothing = False Then
                        For X As Int32 = 0 To lIdx.GetUpperBound(0)
                            Dim oPR As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(lIdx(X))
                            If oPR Is Nothing = False Then
                                cboNewTarget.AddItem(GetCacheObjectValue(oPR.lThisPlayer, ObjectType.ePlayer))
                                cboNewTarget.ItemData(cboNewTarget.NewIndex) = oPR.lThisPlayer
                                cboNewTarget.ItemData2(cboNewTarget.NewIndex) = ObjectType.ePlayer
                            End If
                        Next X
                    End If

                    'For X As Int32 = 0 To goCurrentPlayer.PlayerRelUB
                    '	Dim oPR As PlayerRel = goCurrentPlayer.GetPlayerRelByIndex(X)
                    '	If oPR Is Nothing = False Then
                    '		cboNewTarget.AddItem(GetCacheObjectValue(oPR.lThisPlayer, ObjectType.ePlayer))
                    '		cboNewTarget.ItemData(cboNewTarget.NewIndex) = oPR.lThisPlayer
                    '		cboNewTarget.ItemData2(cboNewTarget.NewIndex) = ObjectType.ePlayer
                    '	End If
                    'Next X
				End If
			End If
		End Sub

		Private Sub cboNewTarget_ItemSelected(ByVal lItemIndex As Integer) Handles cboNewTarget.ItemSelected
			If NewTutorialManager.TutorialOn = True Then
				If lItemIndex > -1 Then
					Dim lID As Int32 = cboNewTarget.ItemData(lItemIndex)
					Dim lID2 As Int32 = cboNewTarget.ItemData2(lItemIndex)
					NewTutorialManager.TriggerFired(NewTutorialManager.StartupTriggerType.eAgentInfiltrationTargetSelected, lID, lID2, -1, "")
				End If
			End If
		End Sub
	End Class

    Public Class fraHistory
        Inherits UIWindow

        Private lstHistory As UIListBox
        Private moAgent As Agent

        Public Sub New(ByRef oUILib As UILib)
            MyBase.New(oUILib)

            'fraHistory initial props
            With Me
                .ControlName = "fraHistory"
                .Left = 260
                .Top = 203
                .Width = 225
                .Height = 100
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .FullScreen = False
                .BorderLineWidth = 2
                .Caption = "History"
                .mbAcceptReprocessEvents = True
            End With

            'lstHistory initial props
            lstHistory = New UIListBox(oUILib)
            With lstHistory
                .ControlName = "lstHistory"
                .Left = 5
                .Top = 10
                .Width = 215
                .Height = 85
                .Enabled = True
                .Visible = True
                .BorderColor = muSettings.InterfaceBorderColor
                .FillColor = muSettings.InterfaceFillColor
                .ForeColor = muSettings.InterfaceBorderColor
                .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
                .HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
                .mbAcceptReprocessEvents = True
            End With
            Me.AddChild(CType(lstHistory, UIControl))
        End Sub

        Public Sub SetFromAgent(ByRef oAgent As Agent)
            moAgent = oAgent
        End Sub

        Public Sub fraHistory_OnNewFrame() Handles Me.OnNewFrame
            moAgent.SmartFillHistoryList(lstHistory)
        End Sub
    End Class
End Class
