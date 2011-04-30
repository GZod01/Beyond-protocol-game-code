Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

Partial Class frmBudget
	'Interface created from Interface Builder
	Private Class fraAgentBudget
		Inherits UIWindow

		Private txtAgentBudget As UITextBox
		Public Sub New(ByRef oUILib As UILib)
			MyBase.New(oUILib)

			'fraAgentBudget initial props
			With Me
				.ControlName = "fraAgentBudget"
				.Left = 138
				.Top = 176
				.Width = 250
				.Height = 86
				.Enabled = True
				.Visible = True
				.BorderColor = muSettings.InterfaceBorderColor
				.FillColor = muSettings.InterfaceFillColor
				.FullScreen = False
				.Moveable = False
				.BorderLineWidth = 2
                .Caption = "Agent Budget"
            End With

			'txtAgentBudget initial props
			txtAgentBudget = New UITextBox(oUILib)
			With txtAgentBudget
				.ControlName = "txtAgentBudget"
				.Left = 5
				.Top = 10
				.Width = 240
				.Height = 70
				.Enabled = True
				.Visible = True
				.Caption = ""
				.ForeColor = muSettings.InterfaceTextBoxForeColor
				.SetFont(New System.Drawing.Font("Courier New", 9.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
				.DrawBackImage = False
				.FontFormat = CType(0, DrawTextFormat)
				.BackColorEnabled = muSettings.InterfaceTextBoxFillColor
				.BackColorDisabled = System.Drawing.Color.FromArgb(255, 105, 105, 105)
				.MaxLength = 0
				.BorderColor = muSettings.InterfaceBorderColor
				.MultiLine = True
                .Locked = True
                .AcceptMouseWheelEvents = False
            End With
			Me.AddChild(CType(txtAgentBudget, UIControl))
		End Sub

		Public Sub SetText()
			Dim sValue As String = "Total: ".PadRight(20, " "c) & goCurrentPlayer.oBudget.lAgentMaintCost.ToString("#,##0").PadLeft(9, " "c)
			Dim lFilterStatus As Int32 = AgentStatus.NewRecruit Or AgentStatus.Dismissed Or AgentStatus.IsDead Or AgentStatus.HasBeenCaptured
			For X As Int32 = 0 To goCurrentPlayer.AgentUB
				If goCurrentPlayer.AgentIdx(X) <> -1 AndAlso (goCurrentPlayer.Agents(X).lAgentStatus And lFilterStatus) = 0 Then
					sValue &= vbCrLf & (goCurrentPlayer.Agents(X).sAgentName & ": ").PadRight(20, " "c) & goCurrentPlayer.Agents(X).lMaintCost.ToString("#,##0").PadLeft(9, " "c)
				End If
			Next X
            If txtAgentBudget.Caption <> sValue Then txtAgentBudget.Caption = sValue

            For X As Int32 = 0 To txtAgentBudget.ChildrenUB
                If txtAgentBudget.moChildren(X) Is Nothing = False Then
                    txtAgentBudget.moChildren(X).AcceptMouseWheelEvents = False
                End If
            Next
		End Sub
	End Class
End Class