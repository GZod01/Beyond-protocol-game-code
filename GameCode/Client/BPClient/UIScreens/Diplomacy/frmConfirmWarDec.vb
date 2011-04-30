Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmConfirmWar
	Inherits UIWindow

	Private lblTitle As UILabel
	Private txtDesc As UITextBox
	Private WithEvents btnConfirm As UIButton
	Private WithEvents btnCancel As UIButton

	Private moPlayerRel As PlayerRel = Nothing
	Private myScore As Byte

	Public Sub New(ByRef oUILib As UILib)
		MyBase.New(oUILib)

		'frmConfirmWar initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eConfirmWarDec
            .ControlName = "frmConfirmWar"
            .Left = MyBase.moUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            .Top = MyBase.moUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 64
            .Width = 256
            .Height = 128
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
			.Left = 1
			.Top = 1
			.Width = 255
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Confirm War Declaration"
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
			.Height = 74
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

		'btnConfirm initial props
		btnConfirm = New UIButton(oUILib)
		With btnConfirm
			.ControlName = "btnConfirm"
			.Left = 10
			.Top = 102
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Confirm"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnConfirm, UIControl))

		'btnCancel initial props
		btnCancel = New UIButton(oUILib)
		With btnCancel
			.ControlName = "btnCancel"
			.Left = 146
			.Top = 102
			.Width = 100
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Cancel"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnCancel, UIControl))

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)
	End Sub

    Public Sub SetPlayerRel(ByRef oRel As PlayerRel, ByVal yNewValue As Byte, ByVal lRankDiff As Int32, ByVal bGuildRank As Boolean)
        moPlayerRel = oRel
        myScore = yNewValue
        Dim sFinal As String
        If bGuildRank = False Then
            sFinal = "Because your rank is higher than the target, declaring war on them will incur a CP penalty across your empire. If they declare war on you first, there will be no penalty."
        Else
            sFinal = "Because your guild has a member with a rank higher than the target, declaring war on them will in a CP penalty across your empire. If they declare war on your first, there will be no penalty."
        End If
        sFinal &= vbCrLf & vbCrLf & "Declaring war would incur a " & lRankDiff & " CP Penalty. Do you wish to proceed?"
        txtDesc.Caption = sFinal
    End Sub

	Private Sub btnConfirm_Click(ByVal sName As String) Handles btnConfirm.Click
		Dim yData(10) As Byte
		Dim X As Int32

		If HasAliasedRights(AliasingRights.eAlterDiplomacy) = False Then
			MyBase.moUILib.AddNotification("You lack rights to alter diplomatic relations.", Color.Red, -1, -1, -1, -1, Int32.MinValue, Int32.MinValue)
			If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Deny.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
			Return
		End If

		If moPlayerRel Is Nothing Then Return
		If moPlayerRel.lPlayerRegards <> glPlayerID Then Return
        'moPlayerRel.WithThisScore = myScore
        moPlayerRel.TargetScore = myScore
        If moPlayerRel.WithThisScore < moPlayerRel.TargetScore Then moPlayerRel.WithThisScore = moPlayerRel.TargetScore
        'moRel.WithThisScore = CByte(hscrRel.Value)
        goCurrentPlayer.SetPlayerRel(moPlayerRel.lThisPlayer, moPlayerRel.WithThisScore)

		If goCurrentEnvir Is Nothing = False Then
			For X = 0 To goCurrentEnvir.lEntityUB
				If goCurrentEnvir.lEntityIdx(X) <> -1 Then
					If goCurrentEnvir.oEntity(X).OwnerID = moPlayerRel.lThisPlayer Then
                        goCurrentEnvir.oEntity(X).yRelID = moPlayerRel.WithThisScore
					End If
				End If
			Next X
		End If

		System.BitConverter.GetBytes(GlobalMessageCode.eSetPlayerRel).CopyTo(yData, 0)
		System.BitConverter.GetBytes(goCurrentPlayer.ObjectID).CopyTo(yData, 2)
		System.BitConverter.GetBytes(moPlayerRel.lThisPlayer).CopyTo(yData, 6)
		yData(10) = myScore
		'Now, send to primary
		MyBase.moUILib.SendMsgToPrimary(yData)

		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub

	Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub
End Class