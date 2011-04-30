Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmCapturedAgent
	Inherits UIWindow

	Private lblTitle As UILabel
	Private lnDiv1 As UILine
	Private picPhoto As UIWindow
	Private lblName As UILabel
	Private lblInfType As UILabel
	Private lblInfLevel As UILabel
	Private lblOwner As UILabel
	Private lblMission As UILabel
	Private lblTarget As UILabel
	Private lblHealth As UILabel
	Private lblHealthValue As UILabel
	Private lnDiv2 As UILine
	Private lblInterrogator As UILabel
	Private lblIntProg As UILabel
	Private WithEvents lstInterrogator As UIListBox
	Private WithEvents btnInterrogate As UIButton
	Private WithEvents btnExecute As UIButton
	Private WithEvents btnRelease As UIButton
	Private WithEvents btnClose As UIButton

	Private moCA As CapturedAgent = Nothing

	Private mfrmPopup As frmAgent = Nothing
	Private mlLastForcedRefresh As Int32

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'frmCapturedAgent initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eCapturedAgent
            .ControlName = "frmCapturedAgent"
            .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 256
            .Width = 256
            .Height = 512
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .BorderLineWidth = 2
            .Moveable = True
        End With

		'lblTitle initial props
		lblTitle = New UILabel(oUILib)
		With lblTitle
			.ControlName = "lblTitle"
			.Left = 5
			.Top = 0
			.Width = 200
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Captured Agent Data"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTitle, UIControl))

		'btnClose initial props
		btnClose = New UIButton(oUILib)
		With btnClose
			.ControlName = "btnClose"
			.Left = 231
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

		'lnDiv1 initial props
		lnDiv1 = New UILine(oUILib)
		With lnDiv1
			.ControlName = "lnDiv1"
			.Left = 1
			.Top = 24
			.Width = 255
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv1, UIControl))

		'picPhoto initial props
		picPhoto = New UIWindow(oUILib)
		With picPhoto
			.ControlName = "picPhoto"
			.Left = 64
			.Top = 35
			.Width = 128
			.Height = 128
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = False
		End With
		Me.AddChild(CType(picPhoto, UIControl))

		'lblName initial props
		lblName = New UILabel(oUILib)
		With lblName
			.ControlName = "lblName"
			.Left = 10
			.Top = 170
			.Width = 230
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Name: Unknown"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblName, UIControl))

		'lblInfType initial props
		lblInfType = New UILabel(oUILib)
		With lblInfType
			.ControlName = "lblInfType"
			.Left = 10
			.Top = 190
			.Width = 230
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Infiltration Mission: Unknown"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblInfType, UIControl))

		'lblInfLevel initial props
		lblInfLevel = New UILabel(oUILib)
		With lblInfLevel
			.ControlName = "lblInfLevel"
			.Left = 10
			.Top = 210
			.Width = 230
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Estimated Infiltration: Unknown"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblInfLevel, UIControl))

		'lblOwner initial props
		lblOwner = New UILabel(oUILib)
		With lblOwner
			.ControlName = "lblOwner"
			.Left = 10
			.Top = 230
			.Width = 230
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Owner: Unknown"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblOwner, UIControl))

		'lblMission initial props
		lblMission = New UILabel(oUILib)
		With lblMission
			.ControlName = "lblMission"
			.Left = 10
			.Top = 250
			.Width = 230
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Mission: Unknown"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblMission, UIControl))

		'lblTarget initial props
		lblTarget = New UILabel(oUILib)
		With lblTarget
			.ControlName = "lblTarget"
			.Left = 10
			.Top = 270
			.Width = 230
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Target: Unknown"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTarget, UIControl))

		'lblHealth initial props
		lblHealth = New UILabel(oUILib)
		With lblHealth
			.ControlName = "lblHealth"
			.Left = 205
			.Top = 45
			.Width = 41
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Health"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblHealth, UIControl))

		'lblHealthValue initial props
		lblHealthValue = New UILabel(oUILib)
		With lblHealthValue
			.ControlName = "lblHealthValue"
			.Left = 205
			.Top = 65
			.Width = 41
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "100%"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = DrawTextFormat.Center Or DrawTextFormat.VerticalCenter
		End With
		Me.AddChild(CType(lblHealthValue, UIControl))

		'lnDiv2 initial props
		lnDiv2 = New UILine(oUILib)
		With lnDiv2
			.ControlName = "lnDiv2"
			.Left = 1
			.Top = 295
			.Width = 255
			.Height = 0
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
		End With
		Me.AddChild(CType(lnDiv2, UIControl))

		'lblInterrogator initial props
		lblInterrogator = New UILabel(oUILib)
		With lblInterrogator
			.ControlName = "lblInterrogator"
			.Left = 10
			.Top = 300
			.Width = 230
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Interrogator:"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblInterrogator, UIControl))

		'lstInterrogator initial props
		lstInterrogator = New UIListBox(oUILib)
		With lstInterrogator
			.ControlName = "lstInterrogator"
			.Left = 10
			.Top = 320
			.Width = 236
			.Height = 100
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.HighlightColor = System.Drawing.Color.FromArgb(255, 32, 64, 128)
		End With
		Me.AddChild(CType(lstInterrogator, UIControl))

		'btnInterrogate initial props
		btnInterrogate = New UIButton(oUILib)
		With btnInterrogate
			.ControlName = "btnInterrogate"
			.Left = 80
			.Top = 450
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Interrogate"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
			.ToolTipText = "Begin interrogating the agent to determine more information." & vbCrLf & _
				  "Careful when selecting an interrogator as brutality may kill the agent."
		End With
		Me.AddChild(CType(btnInterrogate, UIControl))

		'lblIntProg initial props
		lblIntProg = New UILabel(oUILib)
		With lblIntProg
			.ControlName = "lblIntProg"
			.Left = 10
			.Top = 425
			.Width = 230
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Progress: Interrogation not started"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 10.0F, FontStyle.Regular, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(4, DrawTextFormat)
		End With
		Me.AddChild(CType(lblIntProg, UIControl))

		'btnExecute initial props
		btnExecute = New UIButton(oUILib)
		With btnExecute
			.ControlName = "btnExecute"
			.Left = 10
			.Top = 480
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Execute"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
			.ToolTipText = "Order this agent to their death." & vbCrLf & _
				  "This is a public act and the owner" & vbCrLf & _
				  "will be notified that you ordered it."
		End With
		Me.AddChild(CType(btnExecute, UIControl))

		'btnRelease initial props
		btnRelease = New UIButton(oUILib)
		With btnRelease
			.ControlName = "btnRelease"
			.Left = 146
			.Top = 480
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Release"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
			.ToolTipText = "Release this agent to their rightful owner."
		End With
		Me.AddChild(CType(btnRelease, UIControl))

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)

		FillInterrogators()
	End Sub

	Public Sub SetFromCapturedAgent(ByRef oCA As CapturedAgent)
		moCA = oCA
	End Sub

	Private Sub FillInterrogators()
		lstInterrogator.Clear()
		If goCurrentPlayer Is Nothing = False Then

			Dim lSorted() As Int32 = GetSortedIndexArray(goCurrentPlayer.Agents, goCurrentPlayer.AgentIdx, GetSortedIndexArrayType.eAgentName)
			If lSorted Is Nothing = False Then
				For X As Int32 = 0 To lSorted.GetUpperBound(0)
					If goCurrentPlayer.AgentIdx(lSorted(X)) <> -1 Then
						Dim oAgent As Agent = goCurrentPlayer.Agents(lSorted(X))
						If oAgent Is Nothing = False AndAlso ((oAgent.lAgentStatus And AgentStatus.CounterAgent) <> 0 OrElse oAgent.lAgentStatus = AgentStatus.NormalStatus) Then
							lstInterrogator.AddItem(oAgent.sAgentName)
							lstInterrogator.ItemData(lstInterrogator.NewIndex) = oAgent.ObjectID
						End If
					End If
				Next X
			End If
		End If
	End Sub

	Private Sub frmCapturedAgent_OnNewFrame() Handles Me.OnNewFrame
		If moCA Is Nothing = False Then
			Dim sText As String = "Name: " & moCA.sAgentName
			If lblName.Caption <> sText Then lblName.Caption = sText
			sText = "Infiltration Mission: "
			If moCA.yInfType = 255 Then sText &= "Unknown" Else sText &= frmAgent.fraInfiltration.GetInfTypeText(moCA.yInfType)
			If lblInfType.Caption <> sText Then lblInfType.Caption = sText
			sText = "Estimated Infiltration: "
			If moCA.yInfLevel = 0 Then sText &= "Unknown" Else sText &= moCA.yInfLevel
			If lblInfLevel.Caption <> sText Then lblInfLevel.Caption = sText
			sText = "Owner: "
			If moCA.lOwnerID = -1 Then sText &= "Unknown" Else sText &= GetCacheObjectValue(moCA.lOwnerID, ObjectType.ePlayer)
			If lblOwner.Caption <> sText Then lblOwner.Caption = sText
			sText = "Mission: "
			If moCA.lMissionID = -1 Then
				sText &= "Unknown"
			Else
				For X As Int32 = 0 To glMissionUB
					If glMissionIdx(X) = moCA.lMissionID Then
						sText &= goMissions(X).sMissionName
						Exit For
					End If
				Next X
			End If
			If lblMission.Caption <> sText Then lblMission.Caption = sText

			sText = "Target: "
			If moCA.lMissionTargetID = -1 OrElse moCA.iMissionTargetTypeID = -1 Then sText &= "Unknown" Else sText &= GetCacheObjectValue(moCA.lMissionTargetID, moCA.iMissionTargetTypeID)
			If lblTarget.Caption <> sText Then lblTarget.Caption = sText

			sText = "Progress: " & GetInterrogationProgressText(moCA.yInterrogationProgress)
			If lblIntProg.Caption <> sText Then lblIntProg.Caption = sText

			If moCA.yInterrogationProgress <> 1 Then
				If btnInterrogate.Enabled = False Then btnInterrogate.Enabled = True
			ElseIf btnInterrogate.Enabled = True Then
				btnInterrogate.Enabled = False
			End If

			sText = moCA.yHealth.ToString & "%"
			If lblHealthValue.Caption <> sText Then
				lblHealthValue.Caption = sText
				If moCA.yHealth < 25 Then
					lblHealthValue.ForeColor = System.Drawing.Color.FromArgb(255, 255, 0, 0)
				ElseIf moCA.yHealth < 50 Then
					lblHealthValue.ForeColor = System.Drawing.Color.FromArgb(255, 255, 255, 0)
				Else : lblHealthValue.ForeColor = System.Drawing.Color.FromArgb(255, 0, 255, 0)
				End If
			End If
		End If

		If glCurrentCycle - mlLastForcedRefresh > 30 Then
			mlLastForcedRefresh = glCurrentCycle
			Me.IsDirty = True
		End If
	End Sub

	Private Function GetInterrogationProgressText(ByVal yVal As Byte) As String
		Select Case yVal
			Case 2
				Return "Interrogation Paused"
			Case 1
				Return "Interrogation In Progress"
			Case Else
				Return "Interrogation Not Started"
		End Select
	End Function

	Private Sub lstInterrogator_ItemMouseOver(ByVal lIndex As Integer, ByVal lX As Integer, ByVal lY As Integer) Handles lstInterrogator.ItemMouseOver
		If lIndex > -1 AndAlso lIndex < lstInterrogator.ListCount Then
			Dim lID As Int32 = lstInterrogator.ItemData(lIndex)
			If goCurrentPlayer Is Nothing = False Then
				For X As Int32 = 0 To goCurrentPlayer.AgentUB
					If goCurrentPlayer.AgentIdx(X) = lID Then
						'ok, is mfrmPopup nothing?
						If mfrmPopup Is Nothing = True Then mfrmPopup = New frmAgent(goUILib, True, False)

						Dim lTemp As Int32 = Me.Left + Me.Width + 5
						If mfrmPopup.Left <> lTemp Then mfrmPopup.Left = lTemp
						lTemp = Me.Top
						If mfrmPopup.Top <> lTemp Then mfrmPopup.Top = lTemp
						If mfrmPopup.Visible = False Then mfrmPopup.Visible = True
						goUILib.RemoveWindow(mfrmPopup.ControlName)
						goUILib.AddWindow(mfrmPopup)
						mfrmPopup.SetAgentIfNeeded(goCurrentPlayer.Agents(X))
						Exit For
					End If
				Next X
			End If
		End If
	End Sub

	Private Sub btnClose_Click(ByVal sName As String) Handles btnClose.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnExecute_Click(ByVal sName As String) Handles btnExecute.Click
		If btnExecute.Caption.ToUpper = "CONFIRM" Then
			'execute
			Dim yMsg(10) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eCaptureKillAgent).CopyTo(yMsg, 0)
			System.BitConverter.GetBytes(moCA.ObjectID).CopyTo(yMsg, 2)
			yMsg(6) = 1	'1 indicates execute
			System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 7)
			MyBase.moUILib.SendMsgToPrimary(yMsg)
			For X As Int32 = 0 To goCurrentPlayer.CapturedAgentUB
				If goCurrentPlayer.CapturedAgentIdx(X) = moCA.ObjectID Then
					goCurrentPlayer.CapturedAgentIdx(X) = -1
					Exit For
				End If
			Next X
			btnClose_Click(btnClose.ControlName)
		Else : btnExecute.Caption = "Confirm"
		End If
	End Sub

	Private Sub btnInterrogate_Click(ByVal sName As String) Handles btnInterrogate.Click
		'interrogate
		If lstInterrogator.ListIndex > -1 Then
			btnInterrogate.Enabled = False
			Dim lID As Int32 = lstInterrogator.ItemData(lstInterrogator.ListIndex)
			Dim yMsg(10) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eCaptureKillAgent).CopyTo(yMsg, 0)
			System.BitConverter.GetBytes(moCA.ObjectID).CopyTo(yMsg, 2)
			yMsg(6) = 2	'2 indicates interrogate
			System.BitConverter.GetBytes(lID).CopyTo(yMsg, 7)
			MyBase.moUILib.SendMsgToPrimary(yMsg)
		Else : MyBase.moUILib.AddNotification("Select an agent to do the interrogation.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
		End If
	End Sub

	Private Sub btnRelease_Click(ByVal sName As String) Handles btnRelease.Click
		If btnRelease.Caption.ToUpper = "CONFIRM" Then
			'release
			btnRelease.Enabled = False
			Dim yMsg(10) As Byte
			System.BitConverter.GetBytes(GlobalMessageCode.eCaptureKillAgent).CopyTo(yMsg, 0)
			System.BitConverter.GetBytes(moCA.ObjectID).CopyTo(yMsg, 2)
			yMsg(6) = 3	'3 indicates release
			System.BitConverter.GetBytes(-1I).CopyTo(yMsg, 7)
			MyBase.moUILib.SendMsgToPrimary(yMsg)
			For X As Int32 = 0 To goCurrentPlayer.CapturedAgentUB
				If goCurrentPlayer.CapturedAgentIdx(X) = moCA.ObjectID Then
					goCurrentPlayer.CapturedAgentIdx(X) = -1
					Exit For
				End If
			Next X
			btnClose_Click(btnClose.ControlName)
		Else : btnRelease.Caption = "Confirm"
		End If
	End Sub

	Protected Overrides Sub Finalize()
		MyBase.moUILib.RemoveWindow("frmAgentPopup")
		MyBase.Finalize()
	End Sub

    Private Sub frmCapturedAgent_WindowClosed() Handles Me.WindowClosed
        MyBase.moUILib.RemoveWindow("frmAgentPopup")
    End Sub
End Class