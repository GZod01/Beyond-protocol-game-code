Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmInviteFormGuild
	Inherits UIWindow

	Private lblTitle As UILabel
	Private txtDesc As UITextBox
	Private WithEvents btnAccept As UIButton
	Private WithEvents btnReject As UIButton

	Private mlPlayerID As Int32

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'frmInviteFormGuild initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eInviteFromGuild
            .ControlName = "frmInviteFormGuild"
            .Left = 120
            .Top = 167
            .Width = 256
            .Height = 128
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = False
            .Moveable = True
        End With

		'lblTitle initial props
		lblTitle = New UILabel(oUILib)
		With lblTitle
			.ControlName = "lblTitle"
			.Left = 0
			.Top = 0
			.Width = 255
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Form Guild Invitation"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTitle, UIControl))

		'txtDesc initial props
		txtDesc = New UITextBox(oUILib)
		With txtDesc
			.ControlName = "txtDesc"
			.Left = 5
			.Top = 25
			.Width = 245
			.Height = 70
			.Enabled = True
			.Visible = True
			.Caption = ""
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
		Me.AddChild(CType(txtDesc, UIControl))

		'btnAccept initial props
		btnAccept = New UIButton(oUILib)
		With btnAccept
			.ControlName = "btnAccept"
			.Left = 5
			.Top = 100
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Accept"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnAccept, UIControl))

		'btnReject initial props
		btnReject = New UIButton(oUILib)
		With btnReject
			.ControlName = "btnReject"
			.Left = 150
			.Top = 100
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Reject"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnReject, UIControl))

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)
	End Sub

	Private Sub btnAccept_Click(ByVal sName As String) Handles btnAccept.Click

		Dim ofrm As New frmGuildSetup(MyBase.moUILib)
		ofrm.Visible = True

		Dim yMsg(6) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eInviteFormGuild).CopyTo(yMsg, lPos) : lPos += 2
		yMsg(lPos) = 2 : lPos += 1		'player response - ACCEPT
		System.BitConverter.GetBytes(mlPlayerID).CopyTo(yMsg, lPos) : lPos += 4
		MyBase.moUILib.SendMsgToPrimary(yMsg)
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnReject_Click(ByVal sName As String) Handles btnReject.Click
		Dim yMsg(6) As Byte
		Dim lPos As Int32 = 0
		System.BitConverter.GetBytes(GlobalMessageCode.eInviteFormGuild).CopyTo(yMsg, lPos) : lPos += 2
		yMsg(lPos) = 255 : lPos += 1		'player response - REJECT
		System.BitConverter.GetBytes(mlPlayerID).CopyTo(yMsg, lPos) : lPos += 4
		MyBase.moUILib.SendMsgToPrimary(yMsg)
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Public Sub SetFromRequest(ByVal sPlayerName As String, ByVal lPlayerID As Int32)
		txtDesc.Caption = sPlayerName & " has invited you to join in forming a new guild of players. Do you wish to accept?"
		mlPlayerID = lPlayerID
	End Sub
End Class