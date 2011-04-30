Option Strict On

Imports Microsoft.DirectX
Imports Microsoft.DirectX.Direct3D

'Interface created from Interface Builder
Public Class frmDeath
    Inherits UIWindow

    Private lblTitle As UILabel
    Private WithEvents btnDeath As UIButton
	Private txtDetails As UITextBox
	Private WithEvents btnSpawnLive As UIButton

    Public Shared mbInDeathSequence As Boolean = False

    Public Shared mbSelfDestruct As Boolean = False

    Private mbInGuild As Boolean = False

    Public Sub New(ByRef oUILib As UILib)
        MyBase.New(oUILib)

        'frmDeath initial props
        With Me
            .lWindowMetricID = BPMetrics.eWindow.eDeath
            .ControlName = "frmDeath"
            .Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
            .Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 64
            .Width = 256
            .Height = 256
            .Enabled = True
            .Visible = True
            .BorderColor = muSettings.InterfaceBorderColor
            .FillColor = muSettings.InterfaceFillColor
            .FullScreen = True
            .Moveable = True
        End With

        'lblTitle initial props
        lblTitle = New UILabel(oUILib)
        With lblTitle
            .ControlName = "lblTitle"
            .Left = 0
            .Top = 3
            .Width = 256
            .Height = 18
            .Enabled = True
            .Visible = True
            .Caption = "Contigency Plan"
            .ForeColor = muSettings.InterfaceBorderColor
            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            .DrawBackImage = False
            .FontFormat = CType(5, DrawTextFormat)
        End With
		Me.AddChild(CType(lblTitle, UIControl))

		'txtDetails initial props
		txtDetails = New UITextBox(oUILib)
		With txtDetails
			.ControlName = "txtDetails"
			.Left = 5
			.Top = 25
			.Width = 246
            .Height = 203
			.Enabled = True
			.Visible = True
			.Caption = " "
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
		Me.AddChild(CType(txtDetails, UIControl))

		Dim sCaption As String = ""
		If goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.yPlayerPhase <> 255 Then
			'btnDeath initial props
			btnDeath = New UIButton(oUILib)
			With btnDeath
				.ControlName = "btnDeath"
				.Width = 120
				.Left = txtDetails.Left
                .Top = 228
				.Height = 24
                .Enabled = True
                'If goCurrentPlayer Is Nothing = False Then
                '    If goCurrentPlayer.yPlayerPhase = 1 Then .Enabled = False
                '    If goCurrentPlayer.lAccountStatus <> AccountStatusType.eActiveAccount AndAlso goCurrentPlayer.lAccountStatus <> AccountStatusType.eMondelisActive Then .Enabled = True
                'End If
				.Visible = True
				.Caption = "Initiate"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			Me.AddChild(CType(btnDeath, UIControl))

			'btnSpawnLive initial props
			btnSpawnLive = New UIButton(oUILib)
			With btnSpawnLive
				.ControlName = "btnSpawnLive"
				.Width = 120
				.Left = txtDetails.Width + txtDetails.Left - 120
                .Top = 228
				.Height = 24
				.Enabled = True
                .Visible = goCurrentPlayer.yPlayerPhase <> 0 'True
                .Enabled = goCurrentPlayer.lAccountStatus = AccountStatusType.eActiveAccount OrElse goCurrentPlayer.lAccountStatus = AccountStatusType.eMondelisActive OrElse goCurrentPlayer.lAccountStatus = AccountStatusType.eTrialAccount
				.Caption = "Goto Live"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
			Me.AddChild(CType(btnSpawnLive, UIControl))
            If mbSelfDestruct = True Then
                sCaption = "Pick 'Initiate' which will cause you to forfeit all units and facilities left under your control. Funds allocated to the Death Budget will be lost. Once initiated, it cannot be undone." & vbCrLf & vbCrLf
                If goCurrentPlayer.lAccountStatus <> AccountStatusType.eActiveAccount AndAlso goCurrentPlayer.lAccountStatus <> AccountStatusType.eMondelisActive AndAlso goCurrentPlayer.lAccountStatus <> AccountStatusType.eTrialAccount Then
                    sCaption &= "You are unable to Goto Live because you have not activated your subscription for this account!"
                Else
                    sCaption &= "Click 'Goto Live' to bypass the remainder of the tutorial and go straight to the live servers. Your Contingency Plan will be lost and you will be starting fresh without any possessions. Do this only if you are absolutely sure you know what you are doing."
                End If
            Else
                sCaption = "Pick 'Initiate' to begin the contingency plan which will cause you to forfeit all units and facilities left under your control. Funds allocated to the Death Budget will be freed and any assets created using these funds are immediately built. Once initiated, the contingency plan is active for 24 hours." & vbCrLf & vbCrLf
                If goCurrentPlayer.lAccountStatus <> AccountStatusType.eActiveAccount AndAlso goCurrentPlayer.lAccountStatus <> AccountStatusType.eMondelisActive AndAlso goCurrentPlayer.lAccountStatus <> AccountStatusType.eTrialAccount Then
                    sCaption &= "You are unable to Goto Live because you have not activated your subscription for this account!"
                Else
                    sCaption &= "Click 'Goto Live' to bypass the remainder of the tutorial and go straight to the live servers. Your Contingency Plan will be lost and your will be starting fresh without any possessions. Do this only if you are absolutely sure you know what you are doing."
                End If
            End If

		Else
			'btnDeath initial props
			btnDeath = New UIButton(oUILib)
			With btnDeath
				.ControlName = "btnDeath"
				.Width = 120
				.Left = Me.Width \ 2 - (.Width \ 2)
                .Top = 228
				.Height = 24
				.Enabled = True
				.Visible = True
				.Caption = "Initiate"
				.ForeColor = muSettings.InterfaceBorderColor
				.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
				.DrawBackImage = True
				.FontFormat = CType(5, DrawTextFormat)
				.ControlImageRect = New Rectangle(0, 0, 120, 32)
			End With
            Me.AddChild(CType(btnDeath, UIControl))
            If mbSelfDestruct = True Then
                sCaption = "By clicking 'Initiate' you forfeit all units and facilities left under your control and your funds will be lost. Funds allocated to the Death Budget will NOT be available to use. Effectively, this is restarting the game." & vbCrLf & vbCrLf & "Because you chose to self-destruct, your contingency plan is forfeit."
            Else
                sCaption = "By clicking 'Initiate' you forfeit all units and facilities left under your control and your funds will be lost. Funds allocated to the Death Budget will be freed and any assets created using these funds are immediately built. Effectively, this is restarting the game." & vbCrLf & vbCrLf & "Any funds remaining after 24 hours of initiating the contingency plan will be forfeit and the base 100k will be allocated. Therefore, if you do not have 24 hours to commit to restarting, postponing this process may be necessary."
            End If

            'If goCurrentPlayer Is Nothing = False Then
            '    If goCurrentPlayer.oGuild Is Nothing = False Then
            '        ' btnDeath.Left = txtDetails.Left
            '        'btnSpawnLive initial props
            '        btnSpawnLive = New UIButton(oUILib)
            '        With btnSpawnLive
            '            .ControlName = "btnSpawnLive"
            '            .Width = 120
            '            .Left = txtDetails.Width + txtDetails.Left - 120
            '            .Top = 100
            '            .Height = 24
            '            .Enabled = True
            '            .Visible = False
            '            .Caption = "With Guild"
            '            .ForeColor = muSettings.InterfaceBorderColor
            '            .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
            '            .DrawBackImage = True
            '            .FontFormat = CType(5, DrawTextFormat)
            '            .ControlImageRect = New Rectangle(0, 0, 120, 32)
            '        End With
            '        Me.AddChild(CType(btnSpawnLive, UIControl))
            '        If mbSelfDestruct = True Then
            '            sCaption = "Pick 'Initiate' which will cause you to forfeit all units and facilities left under your control. Funds allocated to the Death Budget will be lost. Once initiated, it cannot be undone." & vbCrLf & vbCrLf & "Click 'With Guild' to be placed near your guild mates."
            '        Else
            '            sCaption = "Pick 'Initiate' to begin the contingency plan which will cause you to forfeit all units and facilities left under your control. Funds allocated to the Death Budget will be freed and any assets created using these funds are immediately built. Once initiated, the contingency plan is active for 24 hours." & vbCrLf & vbCrLf & "Click 'With Guild' to spawn near your guild mates. Your Contingency Plan will be lost and your will be starting fresh without any possessions. Do this only if you are absolutely sure you know what you are doing."
            '        End If
            '    End If
            'End If
		End If

		txtDetails.Caption = sCaption

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)
    End Sub

	Private Sub btnDeath_Click(ByVal sName As String) Handles btnDeath.Click
        If btnDeath.Caption.ToLower = "confirm" Then

            If CheckSpawnOptions() = True Then Return

            mbInDeathSequence = True
            MyBase.moUILib.RemoveWindow(Me.ControlName)
            Dim yMsg(2) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerIsDead).CopyTo(yMsg, 0)
            yMsg(2) = 1
            goUILib.SendMsgToPrimary(yMsg)
        Else
            btnDeath.Caption = "Confirm"
            txtDetails.Caption = "This will quit you out of the game and restart the client." & vbCrLf & vbCrLf & "Are you sure?"
        End If
	End Sub

	Private Sub btnSpawnLive_Click(ByVal sName As String) Handles btnSpawnLive.Click
        If btnSpawnLive.Caption.ToLower = "confirm" Then

            If CheckSpawnOptions() = True Then Return

            mbInDeathSequence = True
            MyBase.moUILib.RemoveWindow(Me.ControlName)
            Dim yMsg(2) As Byte
            System.BitConverter.GetBytes(GlobalMessageCode.ePlayerIsDead).CopyTo(yMsg, 0)
            If mbInGuild = True Then yMsg(2) = 254 Else yMsg(2) = 255
            goUILib.SendMsgToPrimary(yMsg)
        Else
            btnSpawnLive.Caption = "Confirm"
            txtDetails.Caption = "This will quit you out of the game and restart the client." & vbCrLf & vbCrLf & "Are you sure?"
        End If
    End Sub

    Private Function CheckSpawnOptions() As Boolean
        If mbInGuild = False AndAlso goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False Then
            btnDeath.Caption = "Normal"
            If btnSpawnLive Is Nothing Then

                btnDeath.Left = txtDetails.Left

                'btnSpawnLive initial props
                btnSpawnLive = New UIButton(MyBase.moUILib)
                With btnSpawnLive
                    .ControlName = "btnSpawnLive"
                    .Width = 120
                    .Left = txtDetails.Width + txtDetails.Left - 120
                    .Top = 228
                    .Height = 24
                    .Enabled = True
                    .Visible = True
                    .Caption = "Goto Live"
                    .ForeColor = muSettings.InterfaceBorderColor
                    .SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
                    .DrawBackImage = True
                    .FontFormat = CType(5, DrawTextFormat)
                    .ControlImageRect = New Rectangle(0, 0, 120, 32)
                End With
                Me.AddChild(CType(btnSpawnLive, UIControl))

            End If
            btnSpawnLive.Caption = "With Guild"
            txtDetails.Caption = "Select whether you want to spawn near your guild bank facility or if you would rather spawn normally. Spawning normally may result in placement far away from your guild."
            btnSpawnLive.Enabled = goCurrentPlayer Is Nothing = False AndAlso goCurrentPlayer.oGuild Is Nothing = False 'AndAlso goCurrentPlayer.oGuild.lGuildHallID > -1 AndAlso goCurrentPlayer.oGuild.iGuildHallLocTypeID > -1
            mbInGuild = True
            Return True
        End If
        Return False
    End Function
End Class

Public Class frmConfirmQuit
	Inherits UIWindow

	Private lblTitle As UILabel
	Private WithEvents btnQuit As UIButton
	Private WithEvents btnCancel As UIButton
	Private txtDetails As UITextBox

	Public Sub New(ByRef oUILib As UILib, ByVal sText As String)
		MyBase.New(oUILib)

		'frmDeath initial props
		With Me
			.ControlName = "frmConfirmQuit"
			.Left = oUILib.oDevice.PresentationParameters.BackBufferWidth \ 2 - 128
			.Top = oUILib.oDevice.PresentationParameters.BackBufferHeight \ 2 - 64
			.Width = 256
			.Height = 256
			.Enabled = True
			.Visible = True
			.BorderColor = muSettings.InterfaceBorderColor
			.FillColor = muSettings.InterfaceFillColor
			.FullScreen = True
			.Moveable = True
		End With

		'lblTitle initial props
		lblTitle = New UILabel(oUILib)
		With lblTitle
			.ControlName = "lblTitle"
			.Left = 0
			.Top = 3
			.Width = 256
			.Height = 18
			.Enabled = True
			.Visible = True
			.Caption = "Confirm Quit"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = False
			.FontFormat = CType(5, DrawTextFormat)
		End With
		Me.AddChild(CType(lblTitle, UIControl))

		'btnDeath initial props
		btnQuit = New UIButton(oUILib)
		With btnQuit
			.ControlName = "btnQuit"
			.Width = 120
			.Left = 5
			.Top = 228
			.Height = 24
			.Enabled = True
			.Visible = True
			.Caption = "Quit"
			.ForeColor = muSettings.InterfaceBorderColor
			.SetFont(New System.Drawing.Font("Microsoft Sans Serif", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0))
			.DrawBackImage = True
			.FontFormat = CType(5, DrawTextFormat)
			.ControlImageRect = New Rectangle(0, 0, 120, 32)
		End With
		Me.AddChild(CType(btnQuit, UIControl))

		'btnCancel initial props
		btnCancel = New UIButton(oUILib)
		With btnCancel
			.ControlName = "btnCancel"
			.Width = 120
			.Left = Me.Width - 125
			.Top = 228
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

		'txtDetails initial props
		txtDetails = New UITextBox(oUILib)
		With txtDetails
			.ControlName = "txtDetails"
			.Left = 5
			.Top = 25
			.Width = 246
			.Height = 203
			.Enabled = True
			.Visible = True
			.Caption = sText
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
		Me.AddChild(CType(txtDetails, UIControl))

		MyBase.moUILib.RemoveWindow(Me.ControlName)
		MyBase.moUILib.AddWindow(Me)

		If goSound Is Nothing = False Then goSound.StartSound("UserInterface\Alert.wav", False, SoundMgr.SoundUsage.eALWAYS_PLAY, Nothing, Nothing)
	End Sub

	Private Sub btnQuit_Click(ByVal sName As String) Handles btnQuit.Click
		frmMain.mbfrmConfirmHandled = True
		MyBase.moUILib.RemoveWindow(Me.ControlName)
		frmMain.Close()
	End Sub

	Public Shared Function CheckLogoffCapability() As String
		Dim sResult As String = ""
		If goCurrentPlayer Is Nothing = False Then
			If goCurrentPlayer.bInPirateSpawn = True Then
				sResult = "There are pirates currently on our home planet with the intent to raid our colony. They will continue to raid our colony even while you are away." & vbCrLf & vbCrLf & "Are you sure you wish to proceed?"
			ElseIf goCurrentPlayer.bInNegativeCashFlow = True Then
				sResult = "You presently are in negative cash flow meaning your civilization is steadily losing credits from the treasury." & vbCrLf & vbCrLf & "Are you sure you wish to proceed?"
			Else : sResult = ""
			End If
		Else : sResult = ""
		End If
		Return sResult
	End Function

	Private Sub btnCancel_Click(ByVal sName As String) Handles btnCancel.Click
		MyBase.moUILib.RemoveWindow(Me.ControlName)
	End Sub
End Class